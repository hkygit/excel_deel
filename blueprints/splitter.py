# blueprints/splitter.py
import os
import tempfile
import zipfile
import logging
from pathlib import Path
from flask import Blueprint, render_template, request, send_file, current_app
from openpyxl import load_workbook, Workbook
from copy import copy
import pandas as pd

from utils.validators import sanitize_filename
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.filter_engine import parse_conditions_from_form, ConditionEngine

logger = logging.getLogger(__name__)
splitter_bp = Blueprint('splitter', __name__, url_prefix='/splitter', template_folder='../templates/splitter')

# ========== 工具函数 ==========

def copy_cell(source_cell, target_cell):
    try:
        if hasattr(source_cell, 'value') and source_cell.value is not None:
            target_cell.value = source_cell.value
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = source_cell.number_format
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.protection = copy(source_cell.protection)
    except AttributeError:
        pass

def copy_worksheet_without_data(source_ws, target_wb, keep_rows_set):
    target_ws = target_wb.create_sheet(title=source_ws.title)
    if keep_rows_set is None:
        keep_rows_set = set(range(1, source_ws.max_row + 1))
    old_to_new = {}
    new_row = 1
    for old_row in range(1, source_ws.max_row + 1):
        if old_row in keep_rows_set:
            old_to_new[old_row] = new_row
            new_row += 1
    for col in source_ws.column_dimensions:
        target_ws.column_dimensions[col].width = source_ws.column_dimensions[col].width
    for old_row, new_row_num in old_to_new.items():
        if old_row in source_ws.row_dimensions:
            target_ws.row_dimensions[new_row_num].height = source_ws.row_dimensions[old_row].height
    merged_cells_to_apply = []
    for merged_range in source_ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
        rows_in_range = set(range(min_row, max_row + 1))
        if not rows_in_range.intersection(keep_rows_set):
            continue
        new_min_row = old_to_new.get(min_row, min_row)
        new_max_row = old_to_new.get(max_row, max_row)
        if new_min_row and new_max_row:
            merged_cells_to_apply.append((new_min_row, min_col, new_max_row, max_col))
    max_col = source_ws.max_column
    is_top_left = {}
    for merged_range in source_ws.merged_cells.ranges:
        min_row, min_col = merged_range.min_row, merged_range.min_col
        is_top_left[(min_row, min_col)] = True
    for old_row, new_row_num in old_to_new.items():
        for col in range(1, max_col + 1):
            if (old_row, col) in is_top_left:
                pass
            else:
                skip = False
                for merged_range in source_ws.merged_cells.ranges:
                    if (merged_range.min_row <= old_row <= merged_range.max_row and
                        merged_range.min_col <= col <= merged_range.max_col):
                        if old_row != merged_range.min_row or col != merged_range.min_col:
                            skip = True
                            break
                if skip:
                    continue
            src_cell = source_ws.cell(row=old_row, column=col)
            dst_cell = target_ws.cell(row=new_row_num, column=col)
            copy_cell(src_cell, dst_cell)
    for min_row, min_col, max_row, max_col in merged_cells_to_apply:
        if min_row < max_row or min_col < max_col:
            target_ws.merge_cells(start_row=min_row, start_column=min_col,
                                  end_row=max_row, end_column=max_col)
    return target_ws

def repair_excel_with_pandas(input_path):
    """尝试用 pandas 修复无法打开的 Excel 文件，注意：此方法会丢失公式和样式"""
    import pandas as pd
    temp_xlsx = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_xlsx.close()
    try:
        with pd.ExcelFile(input_path) as xls:
            sheet_names = xls.sheet_names
            wb = Workbook()
            wb.remove(wb.active)
            for sheet_name in sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, keep_default_na=False)
                ws = wb.create_sheet(title=sheet_name)
                for col_idx, col_name in enumerate(df.columns, 1):
                    ws.cell(row=1, column=col_idx, value=col_name)
                for row_idx, row in enumerate(df.values, 2):
                    for col_idx, val in enumerate(row, 1):
                        ws.cell(row=row_idx, column=col_idx, value=val)
            wb.save(temp_xlsx.name)
    except Exception as e:
        try:
            os.unlink(temp_xlsx.name)
        except:
            pass
        raise e
    return temp_xlsx.name

def split_excel_to_temp(input_path, split_column_name, include_values, exclude_values, skip_sheet_indices, temp_dir):
    wb = load_workbook(input_path, data_only=True)   # 读取公式计算结果
    sheets = wb.worksheets
    first_sheet = sheets[0]
    headers = [cell.value for cell in first_sheet[1]]
    if split_column_name not in headers:
        wb.close()
        raise ValidationError(f"拆分列 '{split_column_name}' 不在第一个工作表的表头中")
    split_col_idx = headers.index(split_column_name) + 1

    all_values = set()
    for idx, sheet in enumerate(sheets):
        if (idx + 1) in skip_sheet_indices:
            continue
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=split_col_idx)
            if cell.value is not None:
                val = str(cell.value).strip()
                if val:
                    all_values.add(val)

    if include_values:
        target_values = [v for v in all_values if v in include_values]
    elif exclude_values:
        target_values = [v for v in all_values if v not in exclude_values]
    else:
        target_values = list(all_values)

    if not target_values:
        wb.close()
        raise ValidationError("没有找到任何需要拆分的值")

    generated_files = []
    for split_val in target_values:
        safe_val = sanitize_filename(str(split_val))
        new_wb = Workbook()
        default_sheet = new_wb.active
        new_wb.remove(default_sheet)
        for idx, src_ws in enumerate(sheets):
            if (idx + 1) in skip_sheet_indices:
                continue
            keep_rows = {1}
            for row in range(2, src_ws.max_row + 1):
                cell_val = src_ws.cell(row=row, column=split_col_idx).value
                if cell_val is not None and str(cell_val) == str(split_val):
                    keep_rows.add(row)
            copy_worksheet_without_data(src_ws, new_wb, keep_rows)
        out_file = Path(temp_dir) / f"{safe_val}.xlsx"
        new_wb.save(str(out_file))
        new_wb.close()
        generated_files.append(str(out_file))
    wb.close()
    return generated_files

def filter_workbook_inplace(input_path, form, skip_sheet_indices=None):
    """
    在 openpyxl 层面筛选行，保留样式、公式结果和合并单元格。
    使用 data_only=True 读取，确保公式计算结果可用。
    """
    if skip_sheet_indices is None:
        skip_sheet_indices = set()

    exact, fuzzy = parse_conditions_from_form(form)
    if not exact and not fuzzy:
        return input_path

    engine = ConditionEngine()

    wb = load_workbook(input_path, data_only=True)   # 关键：读取缓存的计算结果
    try:
        for idx, ws in enumerate(wb.worksheets, start=1):
            if idx in skip_sheet_indices:
                continue

            headers = [cell.value for cell in ws[1]]
            missing_fields = []
            for cond in exact + fuzzy:
                if cond.get('field', '') not in headers:
                    missing_fields.append(cond.get('field', ''))
            if missing_fields:
                logger.warning(f"工作表 '{ws.title}' 缺少字段 {missing_fields}，跳过筛选")
                continue

            col_map = {h: i + 1 for i, h in enumerate(headers)}
            rows_to_delete = []
            for row_idx in range(2, ws.max_row + 1):
                row_dict = {}
                for col_name, col_idx in col_map.items():
                    cell = ws.cell(row=row_idx, column=col_idx)
                    row_dict[col_name] = str(cell.value) if cell.value is not None else ''
                if not engine.apply_to_row(row_dict, exact, fuzzy):
                    rows_to_delete.append(row_idx)

            for row_idx in reversed(rows_to_delete):
                ws.delete_rows(row_idx)

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        wb.save(tmp.name)
        return tmp.name
    finally:
        wb.close()

# ---------- 路由 ----------
@splitter_bp.route('/', methods=['GET'])
def index():
    return render_template('splitter/index.html')

@splitter_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    if 'file' not in request.files:
        raise ValidationError('未上传文件')
    file = request.files['file']
    if file.filename == '':
        raise ValidationError('文件名为空')

    split_column = request.form.get('split_column', '').strip()
    if not split_column:
        raise ValidationError('请指定拆分列名')

    include_values_str = request.form.get('include_values', '').strip()
    exclude_values_str = request.form.get('exclude_values', '').strip()
    skip_sheets_str = request.form.get('skip_sheets', '').strip()

    include_values = [v.strip() for v in include_values_str.split(',') if v.strip()] if include_values_str else []
    exclude_values = [v.strip() for v in exclude_values_str.split(',') if v.strip()] if exclude_values_str else []
    skip_sheet_indices = {int(v.strip()) for v in skip_sheets_str.split(',') if v.strip().isdigit()} if skip_sheets_str else set()

    if include_values and exclude_values:
        raise ValidationError('正选和反选不能同时使用')

    save_path = file_service.save_upload(file, subdir='splitter', allowed_extensions={'xlsx','xls','xlsm'})
    if not save_path.lower().endswith('.xlsx'):
        new_path = save_path + '.xlsx'
        os.rename(save_path, new_path)
        save_path = new_path

    # 尝试用 data_only=True 打开，若失败则修复（修复会丢失公式，但至少能继续）
    try:
        load_workbook(save_path, data_only=True)
    except Exception:
        logger.warning("openpyxl 打开失败，尝试用 pandas 修复（公式将丢失）")
        try:
            repaired_path = repair_excel_with_pandas(save_path)
            os.remove(save_path)
            save_path = repaired_path
        except Exception as repair_err:
            raise ValidationError(f'文件无法被识别为有效的 Excel 文件: {str(repair_err)}')

    filter_path = filter_workbook_inplace(save_path, request.form, skip_sheet_indices)
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            generated_files = split_excel_to_temp(
                input_path=filter_path,
                split_column_name=split_column,
                include_values=include_values,
                exclude_values=exclude_values,
                skip_sheet_indices=skip_sheet_indices,
                temp_dir=temp_dir
            )
            zip_buffer = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
            zip_buffer.close()
            with zipfile.ZipFile(zip_buffer.name, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fpath in generated_files:
                    arcname = os.path.basename(fpath)
                    zf.write(fpath, arcname)

            from datetime import datetime
            original_name = os.path.splitext(file.filename)[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            download_name = f"{sanitize_filename(original_name)}_拆分_{timestamp}.zip"

            return send_file(
                zip_buffer.name,
                as_attachment=True,
                download_name=download_name,
                mimetype='application/zip'
            )
    finally:
        if filter_path != save_path:
            try:
                os.remove(filter_path)
            except:
                pass
