# blueprints/calculator.py
import os
import logging
import tempfile
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, after_this_request
from openpyxl.utils.exceptions import InvalidFileException

from services import FileService, ComputeEngine
from core.decorators import handle_exceptions
from core.response import error_response
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start
from core.exceptions import ValidationError
from services.excel.calculator_processor import ColumnCalculatorProcessor, AggregateBackfillProcessor

logger = logging.getLogger(__name__)
calc_bp = Blueprint('calc', __name__, url_prefix='/calc', template_folder='../templates/calculator')

def convert_xls_to_xlsx(filepath):
    """将 .xls 文件转换为 .xlsx，返回新文件路径（临时文件）"""
    # 读取原始 .xls 文件
    try:
        df = pd.read_excel(filepath, sheet_name=None, dtype=str)
        # 创建临时 .xlsx 文件
        fd, new_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            for sheet_name, data in df.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info(f"已将 .xls 文件转换为 .xlsx: {new_path}")
        return new_path
    except Exception as e:
        raise ValidationError(f"无法转换 .xls 文件: {str(e)}")

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
    converted = False
    # 如果是 .xls 文件，自动转换为 .xlsx
    if ext == '.xls':
        try:
            new_path = convert_xls_to_xlsx(filepath)
            # 删除原始 .xls 文件，使用新文件
            file_service.cleanup([filepath])
            filepath = new_path
            converted = True
        except Exception as e:
            file_service.cleanup([filepath])
            raise ValidationError(f"旧版 Excel 文件转换失败: {str(e)}")

    calc_type = request.form.get('calc_type')
    if calc_type not in ('column', 'aggregate'):
        file_service.cleanup([filepath])
        return error_response('无效的操作类型', 400)

    # 表头行号
    header_row_str = request.form.get('header_row', '').strip()
    if header_row_str:
        try:
            header_row = int(header_row_str)
            if header_row < 1:
                raise ValueError
        except ValueError:
            return error_response('表头行号必须为正整数', 400)
    else:
        try:
            data_start = detect_header_start(filepath)
            header_row = data_start - 1
            if header_row < 1:
                header_row = 1
        except Exception:
            header_row = 1

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
                compute_engine=compute_engine
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
                compute_engine=compute_engine
            )
            result_path = processor.process()

    except InvalidFileException as e:
        file_service.cleanup([filepath])
        raise ValidationError(f"文件格式错误，无法打开: {str(e)}。请确保文件是有效的 Excel 文件（.xlsx/.xls）且未被损坏。")
    except Exception as e:
        file_service.cleanup([filepath])
        raise e

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

    download_name = generate_output_filename('calc', original_filename, suffix='_计算结果')
    return send_file(result_path, as_attachment=True, download_name=download_name)
