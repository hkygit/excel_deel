# blueprints/comparator.py
import io
import logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, jsonify

from services import FileService
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import error_response, success_response
from core.file_naming import generate_output_filename
from schemas.comparator import CompareFullParams, CompareFieldMatchParams
from services.excel.comparator import ExcelComparator

logger = logging.getLogger(__name__)
compare_bp = Blueprint('compare', __name__, url_prefix='/compare', template_folder='../templates/comparator')

@compare_bp.route('/')
def index():
    return render_template('comparator/index.html')

@compare_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()

    file1 = request.files.get('file1')
    file2 = request.files.get('file2')
    if not file1 or not file2:
        raise ValidationError('请上传两个文件')

    compare_mode = request.form.get('compare_mode', 'full')
    sheet1 = request.form.get('sheet1', '').strip() or None
    sheet2 = request.form.get('sheet2', '').strip() or None

    path1 = file_service.save_upload(file1)
    path2 = file_service.save_upload(file2)

    try:
        if compare_mode == 'full':
            key_fields_str = request.form.get('key_fields', '').strip()
            compare_fields_str = request.form.get('compare_fields', '').strip()
            if not key_fields_str or not compare_fields_str:
                raise ValidationError('主键字段和对比字段不能为空')
            key_fields = [k.strip() for k in key_fields_str.split(',') if k.strip()]
            compare_fields = [c.strip() for c in compare_fields_str.split(',') if c.strip()]
            params = CompareFullParams(key_fields=key_fields, compare_fields=compare_fields, sheet1=sheet1, sheet2=sheet2)
            result = ExcelComparator.full_compare(path1, path2, params.key_fields, params.compare_fields,
                                                   params.sheet1, params.sheet2)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sheet_written = False
                if not result['only_in_1'].empty:
                    result['only_in_1'].to_excel(writer, sheet_name='仅存在于文件1', index=False)
                    sheet_written = True
                if not result['only_in_2'].empty:
                    result['only_in_2'].to_excel(writer, sheet_name='仅存在于文件2', index=False)
                    sheet_written = True
                if result['diff'] is not None and not result['diff'].empty:
                    result['diff'].to_excel(writer, sheet_name='内容不一致', index=False)
                    sheet_written = True
            if not sheet_written:
                return jsonify({'code': 200, 'message': '没有发现任何差异'}), 200
            output.seek(0)
            download_name = generate_output_filename('compare', file1.filename, suffix='_对比结果')
            return send_file(output, as_attachment=True, download_name=download_name,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            match_field = request.form.get('match_field', '').strip()
            if not match_field:
                raise ValidationError('请指定匹配字段')
            params = CompareFieldMatchParams(match_field=match_field, sheet1=sheet1, sheet2=sheet2)
            only_in_1, only_in_2 = ExcelComparator.field_match(path1, path2, params.match_field,
                                                                params.sheet1, params.sheet2)
            output_data = []
            if not only_in_1.empty:
                output_data.append(('文件1差异', only_in_1))
            if not only_in_2.empty:
                output_data.append(('文件2差异', only_in_2))
            if not output_data:
                return jsonify({'code': 200, 'message': '两个文件中该字段值完全匹配，无差异'}), 200
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in output_data:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            output.seek(0)
            download_name = generate_output_filename('compare', file1.filename, suffix='_字段差异')
            return send_file(output, as_attachment=True, download_name=download_name,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    finally:
        file_service.cleanup([path1, path2])
