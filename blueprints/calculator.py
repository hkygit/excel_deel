# blueprints/calculator.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app
from openpyxl.utils.exceptions import InvalidFileException

from services import FileService, ComputeEngine
from core.decorators import handle_exceptions
from core.response import error_response
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start
from core.exceptions import ValidationError
from core.excel_helpers import normalize_excel_file
from services.excel.calculator_processor import ColumnCalculatorProcessor, AggregateBackfillProcessor

logger = logging.getLogger(__name__)
calc_bp = Blueprint('calc', __name__, url_prefix='/calc', template_folder='../templates/calculator')




def ensure_excel_compatible(filepath: str, file_service) -> (str, str):
    """
    确保文件可以被后续处理（openpyxl 和 pandas）正常读取。
    返回 (规范化的文件路径, 临时文件标记)

    注意：不在这里删除任何文件，文件清理由调用者负责
    """
    from core.excel_helpers import normalize_excel_file, safe_delete_file

    normalized_path = None
    try:
        normalized_path = normalize_excel_file(filepath)
        logger.info(f"文件规范化完成: {filepath} -> {normalized_path}")
        return normalized_path, normalized_path

    except Exception as e:
        logger.error(f"Excel 文件处理失败: {e}")
        # 清理可能创建的文件
        if normalized_path and os.path.exists(normalized_path) and normalized_path != filepath:
            safe_delete_file(normalized_path)
        raise ValidationError(f"Excel 文件处理失败: {str(e)}。请确保文件是有效的 Excel 文件且未损坏。")




@calc_bp.route('/')
def index():
    return render_template('calculator/index.html')


@calc_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    compute_engine = current_app.container.compute_engine()

    if 'file' not in request.files:
        return error_response('未上传文件', 400)
    file = request.files['file']
    if file.filename == '':
        return error_response('文件名为空', 400)

    original_filename = file.filename
    ext = os.path.splitext(original_filename)[1].lower()
    if ext not in ('.xlsx', '.xls', '.xlsm'):
        raise ValidationError(f"不支持的文件格式: {ext}，请上传 .xlsx 或 .xls 文件")

    # 保存原始上传文件
    filepath = file_service.save_upload(file)
    temp_processed = None

    # 确保文件兼容，总是转换为临时文件
    filepath, temp_processed = ensure_excel_compatible(filepath, file_service)

    calc_type = request.form.get('calc_type')
    if calc_type not in ('column', 'aggregate'):
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        return error_response('无效的操作类型', 400)

    # 表头行号：支持 header_row=0 表示无表头
    header_row_str = request.form.get('header_row', '').strip()
    if header_row_str:
        try:
            header_row = int(header_row_str)
            if header_row < 0:
                raise ValueError
        except ValueError:
            return error_response('表头行号必须为非负整数', 400)
    else:
        try:
            data_start = detect_header_start(filepath)
            header_row = data_start - 1
            if header_row < 0:
                header_row = 0
        except Exception:
            header_row = 0

    result_path = None
    try:
        if calc_type == 'column':
            expression = request.form.get('expression', '').strip()
            result_col = request.form.get('result_col_column')
            merge_by_cols = request.form.get('merge_by_cols', '').strip()

            if not expression:
                return error_response('表达式不能为空', 400)

            processor = ColumnCalculatorProcessor(
                filepath=filepath,
                header_row=header_row,
                expression=expression,
                result_col_input=result_col,
                merge_by_cols_str=merge_by_cols,
                compute_engine=compute_engine,
                already_normalized=True,
            )
            result_path = processor.process()

        else:  # aggregate
            group_cols_str = request.form.get('group_cols', '').strip()
            agg_cols_str = request.form.get('agg_cols', '').strip()
            agg_func = request.form.get('agg_func', 'sum')
            result_col = request.form.get('result_col_aggregate', '').strip()

            if not group_cols_str or not agg_cols_str or not result_col:
                return error_response('分组列、聚合列和结果列不能为空', 400)

            group_cols_raw = [c.strip() for c in group_cols_str.split(',')]
            agg_cols_raw = [c.strip() for c in agg_cols_str.split(',')]

            processor = AggregateBackfillProcessor(
                filepath=filepath,
                header_row=header_row,
                group_cols_raw=group_cols_raw,
                agg_cols_raw=agg_cols_raw,
                agg_func=agg_func,
                result_col_input=result_col,
                compute_engine=compute_engine,
                already_normalized=True,
            )
            result_path = processor.process()

    except InvalidFileException as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise ValidationError(f"文件格式错误，无法打开: {str(e)}。请确保文件是有效的 Excel 文件（.xlsx/.xls）且未被损坏。")
    except Exception as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise

    # 清理临时文件（不删除最终结果）
    try:
        from core.excel_helpers import safe_delete_file

        # 清理处理器创建的临时文件
        if temp_processed and os.path.exists(temp_processed) and temp_processed != result_path:
            if safe_delete_file(temp_processed):
                logger.info(f"已清理临时文件: {temp_processed}")

        # 清理原始上传文件（由文件服务处理）
        # 注意：这里不直接删除，由文件服务的清理机制处理

    except Exception as e:
        logger.warning(f"清理临时文件失败: {e}")

    download_name = generate_output_filename('calc', original_filename, suffix='_计算结果')
    return send_file(result_path, as_attachment=True, download_name=download_name)
