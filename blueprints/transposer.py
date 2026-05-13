import io
import logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.csv_utils import split_fields
from services.transposer.service import pivot_table

logger = logging.getLogger(__name__)
transposer_bp = Blueprint('transposer', __name__, url_prefix='/transposer', template_folder='../templates/transposer')

@transposer_bp.route('/')
def index():
    return render_template('transposer/index.html')

@transposer_bp.route('/convert', methods=['POST'])
@handle_exceptions()
def convert():
    if 'file' not in request.files:
        raise ValidationError('未上传文件')
    file = request.files['file']
    if file.filename == '':
        raise ValidationError('文件名为空')

    # 读取参数
    category_col_raw = request.form.get('category_col', '').strip()
    value_cols_raw = request.form.get('value_cols', '').strip()
    index_cols_raw = request.form.get('index_cols', '').strip()

    if not category_col_raw:
        raise ValidationError('请指定分类列')
    if not value_cols_raw:
        raise ValidationError('请指定值列（可多选）')

    category_col = category_col_raw
    value_cols = split_fields(value_cols_raw)
    index_cols = split_fields(index_cols_raw) if index_cols_raw else None

    # 读取 Excel
    try:
        df = pd.read_excel(file, dtype=str)
    except Exception as e:
        raise ValidationError(f'无法读取文件: {str(e)}')

    # 校验列存在
    all_needed = [category_col] + value_cols
    if index_cols:
        all_needed += index_cols
    missing = [c for c in all_needed if c not in df.columns]
    if missing:
        raise ValidationError(f'列不存在: {", ".join(missing)}')

    # 执行转置
    try:
        result_df = pivot_table(df, index_cols, category_col, value_cols)
    except Exception as e:
        raise ValidationError(f'转换失败: {str(e)}')

    # 输出 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='转置结果')
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name='转置结果.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )