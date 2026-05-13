# blueprints/calculator.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app, after_this_request

from services import FileService, ComputeEngine
from core.decorators import handle_exceptions
from core.response import error_response
from services.excel.calculator_processor import ColumnCalculatorProcessor, AggregateBackfillProcessor

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

    # 表头行号解析
    header_row_str = request.form.get('header_row', '1').strip()
    try:
        header_row = int(header_row_str)
        if header_row < 1:
            raise ValueError
    except ValueError:
        return error_response('表头行号必须为正整数', 400)

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

    return send_file(result_path, as_attachment=True, download_name=os.path.basename(result_path))
