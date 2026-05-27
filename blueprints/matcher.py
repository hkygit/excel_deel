# blueprints/matcher.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app

from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.file_naming import generate_output_filename
from core.csv_utils import split_fields
from services.excel.matcher import ExcelMatcher

logger = logging.getLogger(__name__)
match_bp = Blueprint('match', __name__, url_prefix='/match', template_folder='../templates/matcher')

ALLOWED_TARGET_EXT = {'.xlsx'}

def _check_target_extension(filename: str):
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_TARGET_EXT:
        raise ValidationError(f"目标文件格式不支持，请使用 .xlsx 格式（当前: {ext}）。如为 .xls 文件，请先用 Excel 另存为 .xlsx。")

def _fix_target_extension(filepath: str) -> str:
    if not filepath.lower().endswith('.xlsx'):
        new_path = filepath + '.xlsx'
        try:
            os.rename(filepath, new_path)
            return new_path
        except Exception:
            pass
    return filepath

def _parse_multi_cols(raw: str) -> list:
    """解析多列输入，支持英文逗号、中文逗号、空格分隔，返回去重后的列表"""
    if not raw:
        return []
    # 使用 split_fields 工具函数，它已经支持逗号、空格、换行分隔
    return split_fields(raw)

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
    target_path = _fix_target_extension(target_path)
    source_path = file_service.save_upload(source_file)

    # 解析多列输入
    target_key_cols_raw = request.form.get('target_key_cols', '').strip()
    source_match_cols_raw = request.form.get('source_match_cols', '').strip()
    source_value_cols_raw = request.form.get('source_value_col', '').strip()
    result_cols_raw = request.form.get('result_col', '').strip()
    source_sheets_raw = request.form.get('source_sheets', '').strip()
    allow_duplicate_keys = request.form.get('allow_duplicate_keys') == 'on'

    if not target_key_cols_raw or not source_match_cols_raw or not source_value_cols_raw or not result_cols_raw:
        raise ValidationError('请填写完整配置项')

    target_key_cols = _parse_multi_cols(target_key_cols_raw)
    source_match_cols = _parse_multi_cols(source_match_cols_raw)
    source_value_cols = _parse_multi_cols(source_value_cols_raw)
    result_cols = _parse_multi_cols(result_cols_raw)

    # 验证数量一致性
    if len(source_value_cols) != len(result_cols):
        raise ValidationError(f'源文件值列数量({len(source_value_cols)})与结果列数量({len(result_cols)})不一致')
    if not target_key_cols or not source_match_cols:
        raise ValidationError('主键列和匹配列不能为空')
    if len(target_key_cols) != len(source_match_cols):
        raise ValidationError(f'目标主键列数量({len(target_key_cols)})与源匹配列数量({len(source_match_cols)})不一致')

    sheet_indices = None
    if source_sheets_raw:
        try:
            sheet_indices = [int(x.strip()) for x in source_sheets_raw.split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            raise ValidationError('Sheet索引格式错误，请输入逗号分隔的数字，例如 1,2,3')

    config = {
        'target_key_cols': target_key_cols,
        'source_match_cols': source_match_cols,
        'source_value_cols': source_value_cols,
        'result_cols': result_cols,
        'allow_duplicate_keys': allow_duplicate_keys
    }

    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    success, result = matcher.match()
    if not success:
        error_msg = result if isinstance(result, str) else '\n'.join(result)
        raise ValidationError(error_msg)

    download_name = generate_output_filename('match', target_file.filename, suffix='_匹配结果')
    return send_file(result, as_attachment=True, download_name=download_name)

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
    target_path = _fix_target_extension(target_path)
    source_path = file_service.save_upload(source_file)

    target_key_cols_raw = request.form.get('target_key_cols', '').strip()
    source_match_cols_raw = request.form.get('source_match_cols', '').strip()
    source_value_cols_raw = request.form.get('source_value_col', '').strip()
    result_cols_raw = request.form.get('result_col', '').strip()
    source_sheets_raw = request.form.get('source_sheets', '').strip()
    allow_duplicate_keys = request.form.get('allow_duplicate_keys') == 'on'

    if not target_key_cols_raw or not source_match_cols_raw or not source_value_cols_raw or not result_cols_raw:
        return render_template('error.html', error='请填写完整配置项'), 400

    target_key_cols = _parse_multi_cols(target_key_cols_raw)
    source_match_cols = _parse_multi_cols(source_match_cols_raw)
    source_value_cols = _parse_multi_cols(source_value_cols_raw)
    result_cols = _parse_multi_cols(result_cols_raw)

    if len(source_value_cols) != len(result_cols):
        return render_template('error.html', error=f'源文件值列数量({len(source_value_cols)})与结果列数量({len(result_cols)})不一致'), 400

    sheet_indices = None
    if source_sheets_raw:
        try:
            sheet_indices = [int(x.strip()) for x in source_sheets_raw.split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            return render_template('error.html', error='Sheet索引格式错误，请输入逗号分隔的数字'), 400

    config = {
        'target_key_cols': target_key_cols,
        'source_match_cols': source_match_cols,
        'source_value_cols': source_value_cols,
        'result_cols': result_cols,
        'allow_duplicate_keys': allow_duplicate_keys
    }

    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    info = matcher.debug_info()
    if 'error' in info:
        return render_template('error.html', error=info['error']), 400

    return render_template('debug.html', config=config, info=info)
