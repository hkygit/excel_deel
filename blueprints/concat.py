# blueprints/concat.py
"""分组拼接蓝图：只做参数解析与调度，不含业务逻辑。"""
import os
import re
import logging
from flask import Blueprint, render_template, request, send_file, current_app

from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.file_naming import build_content_disposition

logger = logging.getLogger(__name__)
concat_bp = Blueprint(
    "concat", __name__,
    url_prefix="/concat",
    template_folder="../templates/concat",
)


@concat_bp.route("/")
def index():
    return render_template("concat/index.html")


@concat_bp.route("/process", methods=["POST"])
@handle_exceptions()
def process():
    file = request.files.get("file")
    if not file or not file.filename:
        raise ValidationError("请上传文件")

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".xlsx", ".xls", ".xlsm"):
        raise ValidationError(f"不支持的文件格式: {ext}")

    group_col = (request.form.get("group_col") or "").strip()
    if not group_col:
        raise ValidationError("请指定分组列")

    concat_cols_raw = (request.form.get("concat_cols") or "").strip()
    if not concat_cols_raw:
        raise ValidationError("请至少指定一个拼接列")
    concat_cols = [
        c.strip() for c in re.split(r"[,，\s]+", concat_cols_raw) if c.strip()
    ]
    if not concat_cols:
        raise ValidationError("请至少指定一个拼接列")

    # 分隔符：支持 \\t \\n 转义；允许空串
    separator = request.form.get("separator", ",")
    separator = (
        separator.replace("\\t", "\t").replace("\\n", "\n").replace("\\r", "\r")
    )

    sort_str = (request.form.get("sort_ascending", "true") or "").lower()
    sort_ascending = sort_str not in ("false", "0", "no", "off")

    dup_keep = (request.form.get("dup_keep", "first") or "").strip().lower()
    if dup_keep not in ("first", "last"):
        dup_keep = "first"

    header_row = None
    header_row_raw = (request.form.get("header_row") or "").strip()
    if header_row_raw:
        try:
            header_row = int(header_row_raw)
            if header_row < 1:
                raise ValueError
        except ValueError:
            raise ValidationError("表头行号必须为正整数")

    file_service = current_app.container.file_service()
    config = current_app.container.concat_config()
    service = current_app.container.concat_service()

    path = file_service.save_upload(file)
    try:
        size = os.path.getsize(path)
        if size > config.max_file_size_bytes:
            mb = config.max_file_size_bytes // (1024 * 1024)
            raise ValidationError(f"大文件处理暂未实现，请使用小于{mb}MB的文件")

        result_path = service.execute(
            file_path=path,
            original_filename=file.filename,
            group_col=group_col,
            concat_cols=concat_cols,
            separator=separator,
            sort_ascending=sort_ascending,
            dup_keep=dup_keep,
            header_row=header_row,
        )

        base = os.path.splitext(os.path.basename(file.filename))[0]
        base = re.sub(r'[\\/*?:"<>|]', "_", base) or "result"
        download_name = f"{base}_分组拼接.xlsx"

        resp = send_file(
            result_path,
            as_attachment=True,
            download_name=download_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )
        resp.headers["Content-Disposition"] = build_content_disposition(download_name)
        return resp
    finally:
        file_service.cleanup([path])
