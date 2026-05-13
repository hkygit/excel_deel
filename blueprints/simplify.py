# blueprints/simplify.py
import json, os, time, uuid, logging
from flask import Blueprint, render_template, request, send_file, current_app
from werkzeug.utils import secure_filename

from core.csv_utils import split_fields, read_csv_header, get_column_letter
from services.simplify import (
    submit_task, load_status, get_log,
    ColumnMapper, validate_csv_files
)
from core.response import make_response

logger = logging.getLogger(__name__)
simplify_bp = Blueprint('simplify', __name__, url_prefix='/simplify', template_folder='../templates/simplify')

@simplify_bp.route('/')
def index():
    return render_template('simplify/index.html')

@simplify_bp.route('/process', methods=['POST'])
def process():
    logger.info("开始处理请求")

    if 'header_file' not in request.files:
        return make_response(False, None, '未上传表头文件', 400)
    header_file = request.files['header_file']
    if header_file.filename == '':
        return make_response(False, None, '表头文件名为空', 400)

    data_files = request.files.getlist('data_files')
    data_files = [f for f in data_files if f.filename != '']

    def check_ext(file):
        ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
        return ext == 'csv'

    if not check_ext(header_file):
        return make_response(False, None, '表头文件必须是 CSV 格式', 400)
    for df in data_files:
        if not check_ext(df):
            return make_response(False, None, f'数据文件 {df.filename} 不是 CSV 格式', 400)

    params_str = request.form.get('params', '{}')
    try:
        params = json.loads(params_str)
    except json.JSONDecodeError:
        return make_response(False, None, '参数格式错误', 400)

    key_cols_raw = params.get('key_cols', '')
    keep_cols_raw = params.get('keep_cols', '')
    non_sum_raw = params.get('non_sum_fields', '')
    column_mode = params.get('column_mode', 'name')

    upload_folder = current_app.config['SIMPLIFY_UPLOAD_FOLDER']
    preview_rows = current_app.config.get('SIMPLIFY_PREVIEW_ROWS', 100)

    os.makedirs(upload_folder, exist_ok=True)

    file_service = current_app.container.file_service()
    saved_paths = []
    try:
        header_filename = secure_filename(header_file.filename)
        header_path = file_service.save_upload(header_file, subdir='simplify', allowed_extensions={'csv'})
        saved_paths.append(header_path)

        data_paths = []
        for df in data_files:
            data_path = file_service.save_upload(df, subdir='simplify', allowed_extensions={'csv'})
            saved_paths.append(data_path)
            data_paths.append(data_path)

        headers = read_csv_header(header_path)
        mapper = ColumnMapper(headers)

        if column_mode == 'letter':
            def convert_letters(s, field_name=None):
                if not s:
                    return []
                try:
                    return mapper.resolve_user_input(s, 'letter')
                except ValueError as e:
                    if field_name:
                        raise ValueError(f"字段 '{field_name}' 中的列字母无效: {str(e)}")
                    raise

            try:
                key_cols = convert_letters(key_cols_raw, '主键列')
                keep_cols = convert_letters(keep_cols_raw, '聚合列') if keep_cols_raw else []
                non_sum_fields = convert_letters(non_sum_raw, '非聚合列') if non_sum_raw else []
                filters = params.get('filters', [])
                for f in filters:
                    if 'field' in f:
                        field_val = f['field']
                        converted = convert_letters(field_val, '筛选条件')
                        if not converted:
                            raise ValueError(f"筛选条件中的列字母 '{field_val}' 不存在")
                        f['field'] = converted[0]
            except ValueError as e:
                file_service.cleanup(saved_paths)
                return make_response(False, None, str(e), 400)
        else:
            key_cols = split_fields(key_cols_raw)
            keep_cols = split_fields(keep_cols_raw)
            non_sum_fields = split_fields(non_sum_raw)
            filters = params.get('filters', [])

        params['key_cols'] = key_cols
        params['keep_cols'] = keep_cols
        params['non_sum_fields'] = non_sum_fields
        params['filters'] = filters

        logger.info("开始预检")
        valid, error_msg = validate_csv_files(header_path, data_paths, params, preview_rows)
        logger.info(f"预检完成, valid={valid}, error_msg={error_msg}")
        if not valid:
            file_service.cleanup(saved_paths)
            return make_response(False, None, error_msg, 400)

        original_filename = os.path.splitext(header_file.filename)[0]
        task_id = submit_task(header_path, data_paths, params, original_filename)
        return make_response(True, {'task_id': task_id})

    except Exception as e:
        file_service.cleanup(saved_paths)
        logger.exception("处理异常")
        return make_response(False, None, f'服务器内部错误: {str(e)}', 500)

@simplify_bp.route('/status/<task_id>')
def status(task_id):
    status_data = load_status(task_id)
    if not status_data:
        return make_response(False, None, '任务不存在', 404)
    log_content = get_log(task_id)
    status_data['log'] = log_content
    return make_response(True, status_data)

@simplify_bp.route('/download/<task_id>')
def download(task_id):
    status_data = load_status(task_id)
    if not status_data:
        return make_response(False, None, '任务不存在', 404)
    if status_data['status'] != 'completed':
        return make_response(False, None, '任务未完成', 400)
    result_path = status_data.get('result_path')
    if not result_path or not os.path.exists(result_path):
        return make_response(False, None, '结果文件不存在', 404)
    if result_path.endswith('.xlsx'):
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    else:
        mimetype = 'text/csv'
    return send_file(result_path, as_attachment=True, download_name=os.path.basename(result_path), mimetype=mimetype)
