# blueprints/matcher.py
import logging
from flask import Blueprint, render_template, request, send_file, current_app

from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from services.excel.matcher import ExcelMatcher

logger = logging.getLogger(__name__)
match_bp = Blueprint('match', __name__, url_prefix='/match', template_folder='../templates/matcher')

@match_bp.route('/')
def index():
    return render_template('matcher/index.html')

@match_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()

    target_file = request.files.get('target_file')
    source_file = request.files.get('source_file')
    if not target_file or not source_file:
        raise ValidationError('请上传两个文件')

    target_path = file_service.save_upload(target_file)
    source_path = file_service.save_upload(source_file)

    config = {
        'target_key_cols': request.form.get('target_key_cols', '').strip(),
        'source_match_cols': request.form.get('source_match_cols', '').strip(),
        'source_value_col': request.form.get('source_value_col', '').strip(),
        'result_col': request.form.get('result_col', '').strip(),
        'source_sheets': request.form.get('source_sheets', '').strip()
    }

    if not all([config['target_key_cols'], config['source_match_cols'], config['source_value_col'], config['result_col']]):
        raise ValidationError('请填写完整配置项')

    sheet_indices = None
    if config['source_sheets']:
        try:
            sheet_indices = [int(x.strip()) for x in config['source_sheets'].split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            raise ValidationError('Sheet索引格式错误，请输入逗号分隔的数字，例如 1,2,3')

    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    success, result = matcher.match()
    if not success:
        error_msg = result if isinstance(result, str) else '\n'.join(result)
        raise ValidationError(error_msg)

    return send_file(result, as_attachment=True, download_name=f"matched_{target_file.filename}")

@match_bp.route('/debug', methods=['POST'])
@handle_exceptions()
def debug():
    file_service = current_app.container.file_service()

    target_file = request.files.get('target_file')
    source_file = request.files.get('source_file')
    if not target_file or not source_file:
        return render_template('error.html', error='请上传两个文件'), 400

    target_path = file_service.save_upload(target_file)
    source_path = file_service.save_upload(source_file)

    config = {
        'target_key_cols': request.form.get('target_key_cols', '').strip(),
        'source_match_cols': request.form.get('source_match_cols', '').strip(),
        'source_value_col': request.form.get('source_value_col', '').strip(),
        'result_col': request.form.get('result_col', '').strip(),
        'source_sheets': request.form.get('source_sheets', '').strip()
    }

    if not all([config['target_key_cols'], config['source_match_cols'], config['source_value_col'], config['result_col']]):
        return render_template('error.html', error='请填写完整配置项'), 400

    sheet_indices = None
    if config['source_sheets']:
        try:
            sheet_indices = [int(x.strip()) for x in config['source_sheets'].split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            return render_template('error.html', error='Sheet索引格式错误，请输入逗号分隔的数字'), 400

    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    info = matcher.debug_info()
    if 'error' in info:
        return render_template('error.html', error=info['error']), 400

    return render_template('debug.html', config=config, info=info)
