# blueprints/matcher.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app

from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from services.excel.matcher import ExcelMatcher

logger = logging.getLogger(__name__)
match_bp = Blueprint('match', __name__, url_prefix='/match', template_folder='../templates/matcher')

ALLOWED_TARGET_EXT = {'.xlsx'}

def _check_target_extension(filename: str):
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_TARGET_EXT:
        raise ValidationError(f"目标文件格式不支持，请使用 .xlsx 格式（当前: {ext}）。如为 .xls 文件，请先用 Excel 另存为 .xlsx。")

def _fix_target_extension(filepath: str) -> str:
    """确保目标文件路径以 .xlsx 结尾，否则重命名文件"""
    if not filepath.lower().endswith('.xlsx'):
        new_path = filepath + '.xlsx'
        try:
            os.rename(filepath, new_path)
            logger.info(f"目标文件扩展名已修正: {filepath} -> {new_path}")
            return new_path
        except Exception as e:
            logger.error(f"重命名文件失败: {e}")
    return filepath

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

    _check_target_extension(target_file.filename)

    target_path = file_service.save_upload(target_file)
    target_path = _fix_target_extension(target_path)   # 修复扩展名
    source_path = file_service.save_upload(source_file)

    config = {
        'target_key_cols': request.form.get('target_key_cols', '').strip(),
        'source_match_cols': request.form.get('source_match_cols', '').strip(),
        'source_value_col': request.form.get('source_value_col', '').strip(),
        'result_col': request.form.get('result_col', '').strip(),
        'source_sheets': request.form.get('source_sheets', '').strip(),
        'allow_duplicate_keys': request.form.get('allow_duplicate_keys') == 'on'
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

    _check_target_extension(target_file.filename)

    target_path = file_service.save_upload(target_file)
    target_path = _fix_target_extension(target_path)   # 修复扩展名
    source_path = file_service.save_upload(source_file)

    config = {
        'target_key_cols': request.form.get('target_key_cols', '').strip(),
        'source_match_cols': request.form.get('source_match_cols', '').strip(),
        'source_value_col': request.form.get('source_value_col', '').strip(),
        'result_col': request.form.get('result_col', '').strip(),
        'source_sheets': request.form.get('source_sheets', '').strip(),
        'allow_duplicate_keys': request.form.get('allow_duplicate_keys') == 'on'
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
