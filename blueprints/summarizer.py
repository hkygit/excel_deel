# blueprints/summarizer.py
import os
import tempfile
import logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app

from services import FileService, ComputeEngine, DataFormatter, DataWriter
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import error_response
from core.csv_utils import split_fields

logger = logging.getLogger(__name__)
summary_bp = Blueprint('summary', __name__, url_prefix='/summary', template_folder='../templates/summarizer')

def _dedup(lst):
    """去重并保持顺序"""
    seen = set()
    result = []
    for item in lst:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result

@summary_bp.route('/')
def index():
    return render_template('summarizer/index.html')

@summary_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    compute_engine = current_app.container.compute_engine()
    data_formatter = current_app.container.data_formatter()
    data_writer = current_app.container.data_writer()

    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        raise ValidationError('没有选择文件')

    field_str = request.form.get('fields', '').strip()
    unique_str = request.form.get('unique_keys', '').strip()
    sum_str = request.form.get('sum_fields', '').strip()

    # 统一处理表头定位：优先使用 header_row，其次使用 skip_header_rows
    header_row_str = request.form.get('header_row', '').strip()
    skip_header_str = request.form.get('skip_header_rows', '0').strip()

    try:
        if header_row_str:
            header_row = int(header_row_str)
            if header_row < 1:
                raise ValueError
            skip_header_rows = header_row - 1
        else:
            skip_header_rows = int(skip_header_str)
            if skip_header_rows < 0:
                skip_header_rows = 0
    except ValueError:
        raise ValidationError('表头行号或跳过行数必须为有效正整数')

    # 解析并去重字段列表（解决前端输入重复列导致 groupby 结果列冲突的问题）
    field_list = _dedup(split_fields(field_str))
    if not field_list:
        raise ValidationError('字段列表不能为空')

    unique_keys = _dedup(split_fields(unique_str))
    sum_fields = _dedup(split_fields(sum_str))

    logger.info(f"解析并去重后的字段: {field_list}")
    logger.info(f"唯一键: {unique_keys}")
    logger.info(f"求和字段: {sum_fields}")

    saved_paths = []
    for file in files:
        if file and file.filename:
            path = file_service.save_upload(file)
            saved_paths.append(path)

    if not saved_paths:
        raise ValidationError('没有有效文件上传')

    all_dfs = []
    failed_files = []
    for path in saved_paths:
        try:
            if skip_header_rows > 0:
                all_sheets = pd.read_excel(path, sheet_name=None, dtype=str,
                                           skiprows=skip_header_rows, header=0)
            else:
                all_sheets = pd.read_excel(path, sheet_name=None, dtype=str)
            if not all_sheets:
                raise ValueError(f"文件 {path} 无有效 sheet")
            df_list = []
            for sheet_df in all_sheets.values():
                df_list.append(sheet_df)
            df = pd.concat(df_list, ignore_index=True, sort=False)
            existing_fields = [f for f in field_list if f in df.columns]
            missing = set(field_list) - set(existing_fields)
            if missing:
                logger.warning(f"文件 {os.path.basename(path)} 缺少字段: {missing}")
            df = df[existing_fields]
            for field in field_list:
                if field not in df.columns:
                    df[field] = pd.NA
            all_dfs.append(df)
        except Exception as e:
            logger.warning(f"读取文件 {path} 失败: {e}")
            failed_files.append(os.path.basename(path))
            continue

    if not all_dfs:
        error_msg = "没有成功读取任何有效文件"
        if failed_files:
            error_msg += f"，以下文件读取失败: {', '.join(failed_files)}"
        raise ValidationError(error_msg)

    combined = pd.concat(all_dfs, ignore_index=True, sort=False)

    for field in sum_fields:
        if field not in combined.columns:
            combined[field] = 0
        combined[field] = pd.to_numeric(combined[field], errors='coerce').fillna(0)

    if not unique_keys:
        result = combined.drop_duplicates(subset=field_list, keep='first')
    else:
        valid_keys = [k for k in unique_keys if k in combined.columns]
        if not valid_keys:
            logger.warning("指定的唯一键字段都不存在，按全部字段去重")
            result = combined.drop_duplicates(subset=field_list, keep='first')
        else:
            missing_keys = set(unique_keys) - set(valid_keys)
            if missing_keys:
                logger.warning(f"唯一键字段不存在，将被忽略: {missing_keys}")
            # 将唯一键列中的 NaN 填充为 ''，确保空值行参与聚合而不被丢弃
            combined[valid_keys] = combined[valid_keys].fillna('')
            result = compute_engine.aggregate_by_keys(combined, valid_keys, sum_fields)
            # 再次确保输出中空值列为 ''
            result[valid_keys] = result[valid_keys].fillna('')

    # 调试：打印结果列信息
    logger.info(f"聚合完成，结果列: {list(result.columns)}，行数: {len(result)}")

    result = data_formatter.round_numeric_columns(result)

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    data_writer.write_excel(result, tmp_path)
    return send_file(tmp_path, as_attachment=True, download_name='汇总结果.xlsx')
