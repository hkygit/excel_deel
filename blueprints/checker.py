# blueprints/checker.py
import os, json, logging, tempfile
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import make_response
from services.checker import CheckerService
from core.task.decorator import async_task

logger = logging.getLogger(__name__)
checker_bp = Blueprint('checker', __name__, url_prefix='/checker', template_folder='../templates/checker')

@checker_bp.route('/precheck', methods=['POST'])
@handle_exceptions()
def precheck():
    files_a = request.files.getlist('files_a')
    files_b = request.files.getlist('files_b')
    if not files_a or not files_b:
        raise ValidationError("请上传两批文件")
    saved_a, saved_b = [], []
    for f in files_a:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"pre_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_a.append(p)
    for f in files_b:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"pre_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_b.append(p)
    svc = CheckerService(files_a=saved_a, files_b=saved_b, config=current_app.container.config.provided.checker())
    res = svc.precheck()
    for p in saved_a+saved_b: os.remove(p)
    return make_response(True, res)

@checker_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    files_a = request.files.getlist('files_a')
    files_b = request.files.getlist('files_b')
    if not files_a or not files_b:
        raise ValidationError("请上传两批文件")
    saved_a, saved_b = [], []
    for f in files_a:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"chk_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_a.append(p)
    for f in files_b:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"chk_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_b.append(p)

    key_map_a = [k.strip() for k in request.form['key_map_a'].split(',') if k.strip()]
    key_map_b_str = request.form.get('key_map_b','')
    key_map_b = [k.strip() for k in key_map_b_str.split(',') if k.strip()] if key_map_b_str else key_map_a
    compare_map = json.loads(request.form['compare_map'])
    filters_a = json.loads(request.form.get('filters_a','[]'))
    dup_action = request.form.get('duplicate_action','error')

    app = current_app._get_current_object()
    cfg = app.container.config.provided.checker()
    pre_svc = CheckerService(files_a=saved_a, files_b=saved_b, config=cfg)
    pre_data = pre_svc.precheck()

    @async_task(original_filename='checker')
    def run_check(task_id):
        with app.app_context():
            task_storage = app.container.task_storage()
            def progress_cb(**kw):
                task_storage.update(task_id, **kw)
            service = CheckerService(
                files_a=saved_a, files_b=saved_b,
                key_map_a=key_map_a, key_map_b=key_map_b,
                compare_map=compare_map, filters_a=filters_a,
                duplicate_action=dup_action, config=cfg,
                progress_callback=progress_cb,
                precheck_data=pre_data
            )
            result_path = service.execute()
            return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}

    result = run_check()
    return make_response(True, {'task_id': result['task_id']})

@checker_bp.route('/status/<task_id>')
def status(task_id):
    data = current_app.container.task_storage().load(task_id)
    if not data:
        return make_response(False, None, '任务不存在', 404)
    return make_response(True, data)

@checker_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status')!='completed':
        return make_response(False, None, '不可下载', 400)
    path = data['result_path']
    if not path or not os.path.exists(path):
        return make_response(False, None, '文件不存在', 404)
    return send_file(path, as_attachment=True, download_name='校验结果.xlsx')

@checker_bp.route('/')
def index():
    return render_template('checker/index.html')
