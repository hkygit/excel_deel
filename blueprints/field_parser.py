# blueprints/field_parser.py
import io
import logging
import pandas as pd
from flask import Blueprint, render_template, request, jsonify, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.csv_utils import split_fields

logger = logging.getLogger(__name__)
field_parser_bp = Blueprint('field_parser', __name__, url_prefix='/field-parser', template_folder='../templates/field_parser')

@field_parser_bp.route('/')
def index():
    return render_template('field_parser/index.html')

@field_parser_bp.route('/analyze', methods=['GET', 'POST'])
@handle_exceptions()
def analyze():
    if request.method == 'GET':
        return jsonify({
            'message': '请使用 POST 方法提交 JSON 数据',
            'fields': {'standard_headers': '标准列名列表（文本）', 'compare_headers': '比较列名列表（文本）'}
        })

    if not request.is_json:
        raise ValidationError("仅支持 JSON 格式的请求")

    data = request.get_json()
    text_a = data.get('standard_headers', '')
    text_b = data.get('compare_headers', '')
    if not text_a or not text_b:
        raise ValidationError("请提供 standard_headers 和 compare_headers 字段")

    service = current_app.container.field_parser()
    # 使用 split_fields 解析列名
    cols_a = split_fields(text_a)
    cols_b = split_fields(text_b)

    if not cols_a or not cols_b:
        raise ValidationError("未能解析出有效的列名，请检查输入格式")

    result = service.match_headers(cols_a, cols_b)
    result['cols_a'] = cols_a
    result['cols_b'] = cols_b
    return jsonify(result)

@field_parser_bp.route('/export', methods=['POST'])
@handle_exceptions()
def export():
    """接收完整映射结果和原始列名，生成可编辑的 Excel 文件"""
    if not request.is_json:
        raise ValidationError("仅支持 JSON 格式的请求")

    data = request.get_json()
    matched = data.get('matched', [])
    only_a = data.get('only_a', [])
    only_b = data.get('only_b', [])
    cols_a = data.get('cols_a', [])
    cols_b = data.get('cols_b', [])

    rows = []
    for item in matched:
        rows.append({
            '标准字段': item['standard'],
            '比较字段': item['compare'],
            '相似度': item['score'],
            '备注': '自动匹配'
        })
    for col in only_a:
        rows.append({
            '标准字段': col,
            '比较字段': '',
            '相似度': '',
            '备注': '仅标准'
        })
    for col in only_b:
        rows.append({
            '标准字段': '',
            '比较字段': col,
            '相似度': '',
            '备注': '仅比较'
        })
    df_mapping = pd.DataFrame(rows, columns=['标准字段', '比较字段', '相似度', '备注'])

    max_len = max(len(cols_a), len(cols_b))
    cols_a_padded = list(cols_a) + [''] * (max_len - len(cols_a))
    cols_b_padded = list(cols_b) + [''] * (max_len - len(cols_b))
    df_original = pd.DataFrame({
        '标准文件列名': cols_a_padded,
        '比较文件列名': cols_b_padded
    })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_mapping.to_excel(writer, sheet_name='字段映射', index=False)
        df_original.to_excel(writer, sheet_name='原始数据', index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name='字段映射结果.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
