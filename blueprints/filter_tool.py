# blueprints/filter_tool.py
import io, logging, os, tempfile, zipfile
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from services import FileService
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.filter_engine import parse_conditions_from_form, ConditionEngine
from core.column_resolver import resolve_columns
from core.excel_writer import write_excel_response, write_excel_to_path

logger = logging.getLogger(__name__)
filter_bp = Blueprint('filter', __name__, url_prefix='/filter', template_folder='../templates/filter_tool')

engine = ConditionEngine()

def process_single_df(df, form):
    exact, fuzzy = parse_conditions_from_form(form)
    if exact or fuzzy:
        filtered = engine.apply_to_dataframe(df, exact, fuzzy)
    else:
        filtered = df.copy()
    selected = form.get('selected_columns', '')
    if selected:
        filtered = resolve_columns(filtered, selected)
    return filtered

def _read_file(filepath, form):
    exact, fuzzy = parse_conditions_from_form(form)
    required = set(c['field'] for c in exact)
    required.update(c['field'] for c in fuzzy)

    selected_raw = form.get('selected_columns', '')
    usecols = None
    if selected_raw:
        from core.csv_utils import split_fields
        candidates = split_fields(selected_raw)
        if filepath.lower().endswith(('.xlsx', '.xls')):
            first_df = pd.read_excel(filepath, nrows=0)
        else:
            first_df = pd.read_csv(filepath, nrows=0, encoding='utf-8')
        from core.csv_utils import build_column_mapping
        col_map = build_column_mapping(first_df.columns.tolist())[0]
        usecols = []
        for cand in candidates:
            if cand.isalpha() and len(cand)<=3:
                upper = cand.upper()
                if upper in col_map:
                    usecols.append(col_map[upper])
            elif cand in first_df.columns:
                usecols.append(cand)
        for rc in required:
            if rc not in usecols and rc in first_df.columns:
                usecols.append(rc)

    row_limit_str = form.get('uploadRowsLimit', '')
    nrows = int(row_limit_str) if row_limit_str.isdigit() and int(row_limit_str) > 0 else None

    if filepath.lower().endswith(('.xlsx', '.xls')):
        df = pd.read_excel(filepath, usecols=usecols, nrows=nrows)
    else:
        df = pd.read_csv(filepath, usecols=usecols, nrows=nrows, encoding='utf-8')
    if nrows:
        df = df.head(nrows)
    return df

@filter_bp.route('/')
def index():
    return render_template('filter_tool/index.html')

@filter_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    action = request.form.get('action', 'generate_excel')
    # 切片模式
    truncated = request.form.get('truncated_content', '').strip()
    if truncated:
        from io import StringIO
        df = pd.read_csv(StringIO(truncated), sep=None, engine='python')
        row_limit = request.form.get('uploadRowsLimit', '')
        if row_limit.isdigit() and int(row_limit) > 0:
            df = df.head(int(row_limit))
        filtered = process_single_df(df, request.form)
        if action == 'get_values':
            target = request.form.get('target_field','').strip()
            if not target: raise ValidationError('请填写字段名')
            vals = filtered[target].dropna().astype(str).unique()
            return '\n'.join(sorted(vals)), 200, {'Content-Type': 'text/plain; charset=utf-8'}
        return write_excel_response(filtered, '筛选结果.xlsx', '筛选数据')

    # 完整上传
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        raise ValidationError('请至少上传一个文件')

    file_service = current_app.container.file_service()
    saved_paths = []
    try:
        ALLOWED = {'xlsx', 'xls', 'csv'}
        for f in files:
            if f and f.filename:
                path = file_service.save_upload(f, allowed_extensions=ALLOWED)
                saved_paths.append(path)
    except Exception as e:
        file_service.cleanup(saved_paths)
        raise ValidationError(f"文件保存失败: {e}")

    total_size = sum(os.path.getsize(p) for p in saved_paths)
    async_requested = request.form.get('async') == 'true' or total_size > 100 * 1024 * 1024

    if not async_requested:
        try:
            if len(saved_paths) == 1:
                df = _read_file(saved_paths[0], request.form)
                filtered = process_single_df(df, request.form)
                if action == 'get_values':
                    target = request.form.get('target_field','').strip()
                    if not target: raise ValidationError('请填写字段名')
                    vals = filtered[target].dropna().astype(str).unique()
                    return '\n'.join(sorted(vals)), 200, {'Content-Type': 'text/plain; charset=utf-8'}
                return write_excel_response(filtered, f"filtered_{os.path.basename(saved_paths[0])}", '筛选数据')

            tmp_dir = tempfile.mkdtemp(prefix='filter_')
            results = []
            for path in saved_paths:
                try:
                    df = _read_file(path, request.form)
                    filtered = process_single_df(df, request.form)
                    out_name = os.path.splitext(os.path.basename(path))[0] + '_筛选.xlsx'
                    out_path = os.path.join(tmp_dir, out_name)
                    write_excel_to_path(filtered, out_path, '筛选数据')
                    results.append(out_path)
                except Exception as e:
                    logger.warning(f"文件 {path} 处理失败: {e}")
            if not results:
                raise ValidationError("所有文件处理失败")
            zip_path = tempfile.NamedTemporaryFile(suffix='.zip', delete=False).name
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for rp in results:
                    zf.write(rp, os.path.basename(rp))
            return send_file(zip_path, as_attachment=True, download_name='筛选结果.zip',
                             mimetype='application/zip')
        finally:
            file_service.cleanup(saved_paths)

    # 异步模式 (略，此处保持一致，但代码太长我们简短处理)
    app = current_app._get_current_object()
    task_manager = app.container.task_manager()
    task_storage = app.container.task_storage()
    form_data = dict(request.form)

    task_id = task_manager.create_task(params=form_data, original_filename='filter')
    task_storage.update(task_id, total_files=len(saved_paths))

    def run():
        with app.app_context():
            storage = app.container.task_storage()
            try:
                storage.update(task_id, progress_percent=0, progress_message='开始处理...')
                if len(saved_paths) == 1:
                    df = _read_file(saved_paths[0], form_data)
                    storage.update(task_id, progress_percent=30, progress_message='读取完毕，筛选...')
                    filtered = process_single_df(df, form_data)
                    if form_data.get('action') == 'get_values':
                        target = form_data.get('target_field','').strip()
                        vals = filtered[target].dropna().astype(str).unique()
                        text_path = tempfile.NamedTemporaryFile(suffix='.txt', delete=False).name
                        with open(text_path, 'w', encoding='utf-8') as f:
                            f.write('\n'.join(sorted(vals)))
                        storage.update(task_id, status='completed', result_path=text_path, progress_percent=100)
                        return
                    out_path = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
                    write_excel_to_path(filtered, out_path, '筛选数据')
                    storage.update(task_id, status='completed', result_path=out_path, progress_percent=100)
                    return

                tmp_dir = tempfile.mkdtemp(prefix='filter_')
                results = []
                total = len(saved_paths)
                for idx, path in enumerate(saved_paths):
                    try:
                        storage.update(task_id, progress_percent=int((idx)/total*90), progress_message=f'处理第 {idx+1}/{total} 个文件...')
                        df = _read_file(path, form_data)
                        filtered = process_single_df(df, form_data)
                        out_name = os.path.splitext(os.path.basename(path))[0] + '_筛选.xlsx'
                        out_path = os.path.join(tmp_dir, out_name)
                        write_excel_to_path(filtered, out_path, '筛选数据')
                        results.append(out_path)
                    except Exception as e:
                        logger.warning(f"文件 {path} 处理失败: {e}")
                        storage.append_log(task_id, f"文件 {path} 失败: {e}")
                if not results:
                    raise RuntimeError("所有文件处理失败")

                zip_path = tempfile.NamedTemporaryFile(suffix='.zip', delete=False).name
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for rp in results:
                        zf.write(rp, os.path.basename(rp))
                storage.update(task_id, status='completed', result_path=zip_path, progress_percent=100)
            except Exception as e:
                logger.exception("任务执行失败")
                storage.update(task_id, status='failed', error=str(e), progress_percent=0)
            finally:
                file_service.cleanup(saved_paths)

    task_manager.start_task(task_id, run, ())
    return jsonify({'task_id': task_id})

@filter_bp.route('/status/<task_id>')
def task_status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data: return jsonify({'error': '任务不存在'}), 404
    data['log'] = storage.get_log(task_id, max_lines=100)
    return jsonify(data)

@filter_bp.route('/download/<task_id>')
def task_download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status') != 'completed': return jsonify({'error': '文件不可用'}), 400
    path = data.get('result_path')
    if not path or not os.path.exists(path): return jsonify({'error': '文件不存在'}), 404
    basename = os.path.basename(path)
    mimetype = 'application/zip' if path.endswith('.zip') else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    if path.endswith('.txt'): mimetype = 'text/plain; charset=utf-8'
    return send_file(path, as_attachment=True, download_name=basename, mimetype=mimetype)
