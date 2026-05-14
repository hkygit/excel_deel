# blueprints/convert.py
import os, tempfile, logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, jsonify, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import make_response
from services.convert.service import ConvertService
from core.task.decorator import async_task

logger = logging.getLogger(__name__)
convert_bp = Blueprint('convert', __name__, url_prefix='/convert', template_folder='../templates/convert')

@convert_bp.route('/')
def index():
    return render_template('convert/index.html')

@convert_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file = request.files.get('file')
    if not file or file.filename == '':
        raise ValidationError("请上传文件")
    target_fmt = request.form.get('target_format', 'csv')
    encoding = request.form.get('encoding', 'utf-8-sig')
    separator = request.form.get('separator', ',')

    upload_dir = tempfile.mkdtemp(prefix='convert_')
    save_path = os.path.join(upload_dir, file.filename)
    file.save(save_path)

    app = current_app._get_current_object()

    @async_task(original_filename='convert')
    def run_conversion(task_id):
        with app.app_context():
            task_storage = app.container.task_storage()
            def progress_cb(**kw):
                task_storage.update(task_id, **kw)
            service = ConvertService(
                file_path=save_path,
                target_format=target_fmt,
                encoding=encoding,
                separator=separator,
                progress_callback=progress_cb
            )
            result_path = service.execute()
            return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}

    result = run_conversion()
    return make_response(True, {'task_id': result['task_id']})

@convert_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data:
        return make_response(False, None, '任务不存在', 404)
    return make_response(True, data)

@convert_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status')!='completed':
        return make_response(False, None, '不可下载', 400)
    path = data['result_path']
    if not path or not os.path.exists(path):
        return make_response(False, None, '文件不存在', 404)
    basename = os.path.basename(path)
    return send_file(path, as_attachment=True, download_name=basename)
