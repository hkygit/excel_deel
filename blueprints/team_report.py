# blueprints/team_report.py
import json, os, logging, tempfile, pandas as pd
from flask import Blueprint, render_template, request, jsonify, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from datetime import datetime

logger = logging.getLogger(__name__)

team_report_bp = Blueprint('team_report', __name__, url_prefix='/team-report',
                           template_folder='../templates/team_report')

def _save_uploaded(file, prefix='tr_'):
    if not file or file.filename == '':
        return None
    from werkzeug.utils import secure_filename
    upload_dir = tempfile.mkdtemp(prefix=prefix)
    safe_name = secure_filename(file.filename)
    path = os.path.join(upload_dir, safe_name)
    file.save(path)
    return path

@team_report_bp.route('/')
def index():
    return render_template('team_report/index.html')

@team_report_bp.route('/generate', methods=['POST'])
@handle_exceptions()
def generate():
    mapping_file = request.files.get('mapping_file')
    today_file = request.files.get('today_file')
    yesterday_file = request.files.get('yesterday_file')
    monthly_files = request.files.getlist('monthly_files')  # 新增

    if not today_file:
        raise ValidationError("请上传今日司机流水文件")

    mapping_path = _save_uploaded(mapping_file, prefix='mapping_') if mapping_file else None
    today_path = _save_uploaded(today_file, prefix='today_')
    yesterday_path = _save_uploaded(yesterday_file, prefix='yesterday_') if yesterday_file else None
    monthly_paths = []
    for mf in monthly_files:
        if mf and mf.filename:
            p = _save_uploaded(mf, prefix='monthly_')
            if p:
                monthly_paths.append(p)

    daily_order_target = request.form.get('daily_order_target') or None
    daily_apply_target = request.form.get('daily_apply_target') or None
    monthly_order_target = request.form.get('monthly_order_target') or None
    monthly_apply_target = request.form.get('monthly_apply_target') or None

    app = current_app._get_current_object()
    task_manager = app.container.task_manager()
    task_storage = app.container.task_storage()

    task_id = task_manager.create_task(
        params={'mapping_path': mapping_path, 'today_path': today_path,
                'yesterday_path': yesterday_path, 'monthly_paths': monthly_paths},
        original_filename='team_report'
    )

    def run_task():
        with app.app_context():
            service = app.container.team_report_service()
            try:
                task_storage.update(task_id, progress_percent=10, progress_message='加载映射...')
                if mapping_path and os.path.exists(mapping_path):
                    mapping = service.load_mapping(mapping_path)
                else:
                    raise ValidationError("未提供队伍明细文件且无默认映射")

                task_storage.update(task_id, progress_percent=30, progress_message='处理今日数据...')
                today_df = service._prepare_flow(today_path, mapping)
                yesterday_df = None
                if yesterday_path and os.path.exists(yesterday_path):
                    task_storage.update(task_id, progress_percent=50, progress_message='处理昨日数据...')
                    yesterday_df = service._prepare_flow(yesterday_path, mapping)

                # 月度数据处理
                monthly_df = None
                if monthly_paths:
                    task_storage.update(task_id, progress_percent=60, progress_message='处理月度数据...')
                    monthly_df = service.aggregate_monthly(monthly_paths, mapping)

                task_storage.update(task_id, progress_percent=70, progress_message='生成报表...')
                report_df = service.build_report_data(today_df, yesterday_df, monthly_df)

                task_storage.update(task_id, progress_percent=80, progress_message='写入格式化 Excel...')
                now = datetime.now()
                report_date = now.strftime('%Y/%m/%d')
                month = now.strftime('%m')
                output_dir = app.config.get('OUTPUT_FOLDER', 'outputs')
                os.makedirs(output_dir, exist_ok=True)
                result_path = os.path.join(output_dir, f'team_report_{task_id}.xlsx')

                targets = {
                    'daily_order': int(daily_order_target) if daily_order_target else service.config.daily_order_target,
                    'daily_apply': int(daily_apply_target) if daily_apply_target else service.config.daily_apply_target,
                    'monthly_order': int(monthly_order_target) if monthly_order_target else service.config.monthly_order_target,
                    'monthly_apply': int(monthly_apply_target) if monthly_apply_target else service.config.monthly_apply_target,
                }
                writer = app.container.team_report_writer()
                rows = service.report_to_rows(report_df,
                                              daily_order=targets['daily_order'],
                                              daily_apply=targets['daily_apply'],
                                              monthly_order=targets['monthly_order'],
                                              monthly_apply=targets['monthly_apply'])
                writer.write(rows, result_path, report_date, month,
                             targets['daily_order'], targets['daily_apply'],
                             targets['monthly_order'], targets['monthly_apply'])

                task_storage.update(task_id, status='completed', result_path=result_path,
                                    progress_percent=100, completed_at=pd.Timestamp.now().isoformat())
            except Exception as e:
                logger.exception("生成日报失败")
                task_storage.update(task_id, status='failed', error=str(e))
            finally:
                for p in [mapping_path, today_path, yesterday_path] + monthly_paths:
                    if p and os.path.exists(p):
                        try: os.remove(p)
                        except: pass

    task_manager.start_task(task_id, run_task, ())
    return jsonify({'task_id': task_id})

@team_report_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data:
        return jsonify({'error': '任务不存在'}), 404
    return jsonify(data)

@team_report_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status') != 'completed':
        return jsonify({'error': '报表不可用'}), 400
    path = data.get('result_path')
    if not path or not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404
    response = send_file(
        path,
        as_attachment=True,
        download_name='每日数据对比.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        conditional=True
    )
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    response.headers['X-Accel-Buffering'] = 'no'
    return response
