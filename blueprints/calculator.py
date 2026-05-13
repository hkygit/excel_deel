# blueprints/calculator.py
import os
import logging
import gc
import openpyxl
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, after_this_request

from services import FileService, ComputeEngine
from core.data_access_layer import get_data_access_layer
from core.decorators import handle_exceptions
from core.response import error_response
from core.excel_utils import ExcelFile
from core.csv_utils import build_column_mapping

logger = logging.getLogger(__name__)
calc_bp = Blueprint('calc', __name__, url_prefix='/calc', template_folder='../templates/calculator')

@calc_bp.route('/')
def index():
    return render_template('calculator/index.html')

@calc_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    compute_engine = current_app.container.compute_engine()
    dal = get_data_access_layer()

    if 'file' not in request.files:
        return error_response('未上传文件', 400)
    file = request.files['file']
    if file.filename == '':
        return error_response('文件名为空', 400)

    filepath = file_service.save_upload(file)
    calc_type = request.form.get('calc_type')
    if calc_type not in ('column', 'aggregate'):
        file_service.cleanup([filepath])
        return error_response('无效的操作类型', 400)

    # ---------- 取表头行号（默认1） ----------
    header_row_str = request.form.get('header_row', '1').strip()
    try:
        header_row = int(header_row_str)
        if header_row < 1:
            raise ValueError
    except ValueError:
        return error_response('表头行号必须为正整数', 400)
    header_index = header_row - 1   # pandas 的 header 参数

    result_path = None
    if calc_type == 'column':
        expression_str = request.form.get('expression', '').strip()
        result_col_input = request.form.get('result_col_column')
        merge_by_cols_str = request.form.get('merge_by_cols', '').strip()

        if not expression_str:
            return error_response('表达式不能为空', 400)

        df = pd.read_excel(filepath, header=header_index, dtype=str)
        df = dal._auto_clean_dataframe(df)

        variables = compute_engine.extract_variables(expression_str)
        missing_cols = [v for v in variables if v not in df.columns]
        if missing_cols:
            return error_response(f'列不存在: {missing_cols}', 400)

        def calc_row(row):
            values = {col: row[col] for col in variables}
            return compute_engine.evaluate_expression(expression_str, values)
        df['Result'] = df.apply(calc_row, axis=1)

        with ExcelFile(filepath) as ef:
            ws = ef.active_sheet
            # 处理结果列
            try:
                result_col_idx = openpyxl.utils.column_index_from_string(result_col_input)
                is_new_col = False
                result_col_name = result_col_input
            except:
                is_new_col = True
                result_col_name = result_col_input.strip()
                max_col = ws.max_column
                result_col_idx = max_col + 1
                ws.insert_cols(result_col_idx)
                # 新列标题写入 header_row 行
                _safe_write_cell(ws, header_row, result_col_idx, result_col_name)

            # 若不是新列，但标题为空，则写入
            if not is_new_col:
                existing = _safe_get_cell_value(ws, header_row, result_col_idx)
                if existing is None or str(existing).strip() == '':
                    _safe_write_cell(ws, header_row, result_col_idx, result_col_name)

            start_row, end_row = header_row + 1, ws.max_row
            for i, val in enumerate(df['Result']):
                ws.cell(start_row + i, result_col_idx).value = val

            if merge_by_cols_str:
                merge_keys = [openpyxl.utils.column_index_from_string(c.strip()) for c in merge_by_cols_str.split(',')]
                if merge_keys:
                    from services.excel.excel_service import ExcelService
                    ExcelService.copy_merge_structure(ws, merge_keys[0], result_col_idx, start_row, end_row)

            base_name = os.path.splitext(os.path.basename(filepath))[0]
            output_filename = f"{base_name}_处理版.xlsx"
            result_path = os.path.join(os.path.dirname(filepath), output_filename)
            ef.save(result_path)
            ef._workbook.close()
            gc.collect()

    elif calc_type == 'aggregate':
        group_cols_str = request.form.get('group_cols', '').strip()
        agg_cols_str = request.form.get('agg_cols', '').strip()
        agg_func = request.form.get('agg_func', 'sum')
        result_col_input = request.form.get('result_col_aggregate', '').strip()

        if not group_cols_str or not agg_cols_str or not result_col_input:
            return error_response('分组列、聚合列和结果列不能为空', 400)

        group_cols_raw = [c.strip() for c in group_cols_str.split(',')]
        agg_cols_raw = [c.strip() for c in agg_cols_str.split(',')]

        # 构建列字母→列名映射
        try:
            df_header = pd.read_excel(filepath, header=header_index, nrows=0)
        except Exception as e:
            return error_response(f'读取表头失败: {str(e)}', 400)
        headers = list(df_header.columns)
        letter_to_name, name_to_letter = build_column_mapping(headers)
        logger.info(f"表头 {len(headers)} 列，前5: {headers[:5]}")

        def resolve_col(col):
            col = col.strip()
            if col in headers:
                return col
            upper = col.upper()
            if upper in letter_to_name:
                return letter_to_name[upper]
            return col

        group_cols = [resolve_col(c) for c in group_cols_raw]
        agg_cols = [resolve_col(c) for c in agg_cols_raw]
        logger.info(f"分组列: {group_cols_raw} -> {group_cols}")
        logger.info(f"聚合列: {agg_cols_raw} -> {agg_cols}")

        df = pd.read_excel(filepath, header=header_index, dtype=str)
        df = dal._auto_clean_dataframe(df)

        missing = [c for c in group_cols + agg_cols if c not in df.columns]
        if missing:
            return error_response(f'列不存在: {missing}', 400)

        result_df = compute_engine.simple_aggregate(df, group_cols, agg_cols, agg_func)

        # 回填到原始 Excel
        with ExcelFile(filepath) as ef:
            ws = ef.active_sheet
            # 建立 header_row 行的列名→列索引映射
            name_to_col_idx = {}
            for idx, cell in enumerate(ws[header_row], start=1):
                if cell.value is not None:
                    name_to_col_idx[str(cell.value).strip()] = idx

            def get_col_index(col):
                col = col.strip()
                if col in name_to_col_idx:
                    return name_to_col_idx[col]
                try:
                    return openpyxl.utils.column_index_from_string(col)
                except ValueError:
                    raise ValueError(f"列 '{col}' 在原始表中不存在")

            group_col_indices = [get_col_index(c) for c in group_cols]

            # 处理结果列
            try:
                result_col_idx = openpyxl.utils.column_index_from_string(result_col_input)
                is_new_col = False
                result_col_name = result_col_input
            except:
                is_new_col = True
                result_col_name = result_col_input.strip()
                result_col_idx = ws.max_column + 1
                ws.insert_cols(result_col_idx)
                # 新列标题写在 header_row 行
                _safe_write_cell(ws, header_row, result_col_idx, result_col_name)

            if not is_new_col:
                # 若标题为空，则在 header_row 行写入
                existing = _safe_get_cell_value(ws, header_row, result_col_idx)
                if existing is None or str(existing).strip() == '':
                    _safe_write_cell(ws, header_row, result_col_idx, result_col_name)

            # 构建聚合结果映射
            result_map = {}
            for _, row in result_df.iterrows():
                key = tuple(str(row[c]) for c in group_cols)
                result_map[key] = row[agg_cols[0]]

            start_row, end_row = header_row + 1, ws.max_row
            for row in range(start_row, end_row + 1):
                key = tuple(str(ws.cell(row, idx).value) for idx in group_col_indices)
                ws.cell(row, result_col_idx).value = result_map.get(key, '')

            base_name = os.path.splitext(os.path.basename(filepath))[0]
            output_filename = f"{base_name}_处理版.xlsx"
            result_path = os.path.join(os.path.dirname(filepath), output_filename)
            ef.save(result_path)
            ef._workbook.close()
            gc.collect()

    @after_this_request
    def cleanup_files(response):
        try:
            if filepath and os.path.exists(filepath):
                os.remove(filepath)
            if result_path and os.path.exists(result_path):
                os.remove(result_path)
        except Exception as e:
            logger.warning(f"清理临时文件失败: {e}")
        return response

    return send_file(result_path, as_attachment=True, download_name=os.path.basename(result_path))


def _safe_get_cell_value(ws, row, col):
    """获取单元格值，兼容合并单元格"""
    cell = ws.cell(row, col)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                return ws.cell(merged_range.min_row, merged_range.min_col).value
    return cell.value


def _safe_write_cell(ws, row, col, value):
    """安全写入单元格，若为合并单元格则先取消合并"""
    cell = ws.cell(row, col)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        # 取消包含该单元格的所有合并区域
        to_unmerge = []
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
                to_unmerge.append(str(merged_range))
        for rng in to_unmerge:
            ws.unmerge_cells(rng)
    ws.cell(row, col).value = value
