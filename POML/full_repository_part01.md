# Project Object Markup Language (POML)

- **repository**: `project`
- **export_filter**: `full_repository`
- **total_files**: 167

## Files

- **app.py**
  - language: `python`
  - content:
    ```python
#!/usr/bin/env python3
# ---------------------------------------------------------------------------
# 环境守卫：强制切到项目 .venv
# 无论用哪个 python 启动（全局 Python、别的项目 venv、IDE 里选的），
# 都先切到 E:\project\.venv\Scripts\python.exe 再执行。
# ---------------------------------------------------------------------------
import os as _os
import sys as _sys
import subprocess as _subprocess
from pathlib import Path as _Path

_PROJECT_ROOT = _Path(__file__).resolve().parent
_VENV_PY = _PROJECT_ROOT / ".venv" / "Scripts" / "python.exe"

if _VENV_PY.exists():
    try:
        _current = _Path(_sys.executable).resolve()
        _target = _VENV_PY.resolve()
        if _current != _target:
            print(f"[环境切换] {_current} → {_target}")
            _sys.stdout.flush()
            _env = _os.environ.copy()
            _env.pop("PYTHONPATH", None)
            _env.pop("PYTHONHOME", None)
            _env["PATH"] = _os.pathsep.join(
                p for p in _env.get("PATH", "").split(_os.pathsep)
                if not (".venv" in p.lower() and "scripts" in p.lower()
                        and str(_PROJECT_ROOT).lower() not in p.lower())
            )
            _sys.exit(_subprocess.run(
                [str(_target), str(_Path(__file__).resolve())] + _sys.argv[1:],
                env=_env,
            ).returncode)
    except OSError:
        pass
# ---------------------------------------------------------------------------
import os, logging, threading, time
from flask import Flask, render_template, request, g
from werkzeug.exceptions import HTTPException, NotFound
from core.di import Container
from core.config import settings
from core.logging_setup import init_app as init_logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(name)s %(message)s')
logger = logging.getLogger(__name__)

from blueprints.calculator import calc_bp
from blueprints.comparator import compare_bp
from blueprints.matcher import match_bp
from blueprints.summarizer import summary_bp
from blueprints.splitter import splitter_bp
from blueprints.field_parser import field_parser_bp
from blueprints.checker import checker_bp
from blueprints.convert import convert_bp
from blueprints.team_report import team_report_bp
from blueprints.transposer import transposer_bp
from blueprints.subtotal import subtotal_bp

def clean_old_files(directory, age_hours):
    now = time.time()
    if not os.path.exists(directory): return
    for root, dirs, files in os.walk(directory):
        for f in files:
            path = os.path.join(root, f)
            try:
                if os.path.getmtime(path) < now - age_hours * 3600:
                    os.remove(path)
                    logger.debug(f"删除过期文件: {path}")
            except Exception as e:
                logger.warning(f"删除文件失败 {path}: {e}")

def scheduled_cleanup(app):
    while True:
        time.sleep(3600)
        with app.app_context():
            cfg = app.config
            clean_old_files(cfg['UPLOAD_FOLDER'], 24)
            logger.info("执行定期清理任务")

def create_app():
    app = Flask(__name__)
    app.debug = False
    raw_dict = settings.model_dump()
    flat_config = {}
    for key, value in raw_dict.items():
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                flat_key = f"{key}_{sub_key}".upper()
                flat_config[flat_key] = sub_value
        else:
            flat_config[key.upper()] = value
    app.config.update(flat_config)
    app.config['MAX_CONTENT_LENGTH'] = settings.max_content_length

    container = Container()
    app.container = container

    os.makedirs(app.config.get('UPLOAD_FOLDER', 'uploads'), exist_ok=True)

    app.register_blueprint(calc_bp, url_prefix='/calc')
    app.register_blueprint(compare_bp, url_prefix='/compare')
    app.register_blueprint(match_bp, url_prefix='/match')
    app.register_blueprint(summary_bp, url_prefix='/summary')
    app.register_blueprint(splitter_bp, url_prefix='/splitter')
    app.register_blueprint(field_parser_bp, url_prefix='/field-parser')
    app.register_blueprint(checker_bp, url_prefix='/checker')
    app.register_blueprint(convert_bp, url_prefix='/convert')
    app.register_blueprint(team_report_bp, url_prefix='/team-report')
    app.register_blueprint(transposer_bp, url_prefix='/transposer')
    app.register_blueprint(subtotal_bp, url_prefix='/subtotal')

    init_logging(app)

    @app.route('/')
    def index():
        return render_template('index.html')

    @app.route('/favicon.ico')
    def favicon():
        return '', 204

    @app.errorhandler(NotFound)
    def handle_not_found(e):
        ignored = ('/favicon.ico', '/robots.txt', '/apple-touch-icon', '/.env', '/wp-admin', '/.well-known/')
        if any(request.path.startswith(p) for p in ignored):
            return '', 404
        logger.warning(f"404 Not Found: {request.path}")
        return render_template('error.html', error='页面未找到'), 404

    @app.errorhandler(405)
    def handle_method_not_allowed(e):
        logger.warning(f"405 Method Not Allowed: {request.method} {request.path}")
        return render_template('error.html', error='请求方法不允许'), 405

    @app.errorhandler(HTTPException)
    def handle_http_exception(e):
        logger.warning(f"HTTP {e.code}: {request.path}")
        return render_template('error.html', error=e.description), e.code

    @app.errorhandler(Exception)
    def handle_general_exception(e):
        logger.exception("未处理的异常")
        return render_template('error.html', error='服务器内部错误'), 500

    if not app.debug:
        cleanup_thread = threading.Thread(target=scheduled_cleanup, args=(app,), daemon=True)
        cleanup_thread.start()
        logger.info("已启动后台清理线程")

    # 修复 #2: 服务启动时恢复孤儿任务（pending/running → failed）
    try:
        container.task_storage().recover_orphan_tasks()
        logger.info("孤儿任务恢复完成")
    except Exception as e:
        logger.warning(f"恢复孤儿任务失败: {e}")

    return app

app = create_app()

if __name__ == '__main__':
    from waitress import serve
    try:
        port = int(os.environ.get('PORT', 5000))
        logger.info(f"启动服务...端口 {port}")
        serve(app, host='127.0.0.1', port=port,
              max_request_body_size=2 * 1024 * 1024 * 1024,
              channel_timeout=3600, connection_limit=100,
              clear_untrusted_proxy_headers=True, _quiet=False)
    except Exception as e:
        logger.exception("服务启动失败")
    ```

- **blueprints/__init__.py**
  - language: `python`
  - content:
    ```python
"""Flask 蓝图包。

各蓝图定义在 blueprints/<name>.py 中，由 app.py 显式注册。

本 __init__ 刻意不导入任何蓝图，以避免：
- 循环导入（blueprint 之间可能相互引用工具模块）
- 注册顺序问题（依赖 app.container 初始化完成）
"""
    ```

- **blueprints/calculator.py**
  - language: `python`
  - content:
    ```python
# blueprints/calculator.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app
from openpyxl.utils.exceptions import InvalidFileException

from services import FileService, ComputeEngine
from core.decorators import handle_exceptions
from core.response import error_response
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start
from core.exceptions import ValidationError
from core.excel_helpers import normalize_excel_file
from services.excel.calculator_processor import ColumnCalculatorProcessor, AggregateBackfillProcessor

logger = logging.getLogger(__name__)
calc_bp = Blueprint('calc', __name__, url_prefix='/calc', template_folder='../templates/calculator')




def ensure_excel_compatible(filepath: str, file_service) -> (str, str):
    """
    确保文件可以被后续处理（openpyxl 和 pandas）正常读取。
    返回 (规范化的文件路径, 临时文件标记)

    注意：不在这里删除任何文件，文件清理由调用者负责
    """
    from core.excel_helpers import normalize_excel_file, safe_delete_file

    normalized_path = None
    try:
        normalized_path = normalize_excel_file(filepath)
        logger.info(f"文件规范化完成: {filepath} -> {normalized_path}")
        return normalized_path, normalized_path

    except Exception as e:
        logger.error(f"Excel 文件处理失败: {e}")
        # 清理可能创建的文件
        if normalized_path and os.path.exists(normalized_path) and normalized_path != filepath:
            safe_delete_file(normalized_path)
        raise ValidationError(f"Excel 文件处理失败: {str(e)}。请确保文件是有效的 Excel 文件且未损坏。")




@calc_bp.route('/')
def index():
    return render_template('calculator/index.html')


@calc_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    compute_engine = current_app.container.compute_engine()

    if 'file' not in request.files:
        return error_response('未上传文件', 400)
    file = request.files['file']
    if file.filename == '':
        return error_response('文件名为空', 400)

    original_filename = file.filename
    ext = os.path.splitext(original_filename)[1].lower()
    if ext not in ('.xlsx', '.xls', '.xlsm'):
        raise ValidationError(f"不支持的文件格式: {ext}，请上传 .xlsx 或 .xls 文件")

    # 保存原始上传文件
    filepath = file_service.save_upload(file)
    temp_processed = None

    # 确保文件兼容，总是转换为临时文件
    filepath, temp_processed = ensure_excel_compatible(filepath, file_service)

    calc_type = request.form.get('calc_type')
    if calc_type not in ('column', 'aggregate'):
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        return error_response('无效的操作类型', 400)

    # 表头行号：支持 header_row=0 表示无表头
    header_row_str = request.form.get('header_row', '').strip()
    if header_row_str:
        try:
            header_row = int(header_row_str)
            if header_row < 0:
                raise ValueError
        except ValueError:
            return error_response('表头行号必须为非负整数', 400)
    else:
        try:
            data_start = detect_header_start(filepath)
            header_row = data_start - 1
            if header_row < 0:
                header_row = 0
        except Exception:
            header_row = 0

    result_path = None
    try:
        if calc_type == 'column':
            expression = request.form.get('expression', '').strip()
            result_col = request.form.get('result_col_column')
            merge_by_cols = request.form.get('merge_by_cols', '').strip()

            if not expression:
                return error_response('表达式不能为空', 400)

            processor = ColumnCalculatorProcessor(
                filepath=filepath,
                header_row=header_row,
                expression=expression,
                result_col_input=result_col,
                merge_by_cols_str=merge_by_cols,
                compute_engine=compute_engine,
                already_normalized=True,
            )
            result_path = processor.process()

        else:  # aggregate
            group_cols_str = request.form.get('group_cols', '').strip()
            agg_cols_str = request.form.get('agg_cols', '').strip()
            agg_func = request.form.get('agg_func', 'sum')
            result_col = request.form.get('result_col_aggregate', '').strip()

            if not group_cols_str or not agg_cols_str or not result_col:
                return error_response('分组列、聚合列和结果列不能为空', 400)

            group_cols_raw = [c.strip() for c in group_cols_str.split(',')]
            agg_cols_raw = [c.strip() for c in agg_cols_str.split(',')]

            processor = AggregateBackfillProcessor(
                filepath=filepath,
                header_row=header_row,
                group_cols_raw=group_cols_raw,
                agg_cols_raw=agg_cols_raw,
                agg_func=agg_func,
                result_col_input=result_col,
                compute_engine=compute_engine,
                already_normalized=True,
            )
            result_path = processor.process()

    except InvalidFileException as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise ValidationError(f"文件格式错误，无法打开: {str(e)}。请确保文件是有效的 Excel 文件（.xlsx/.xls）且未被损坏。")
    except Exception as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise

    # 清理临时文件（不删除最终结果）
    try:
        from core.excel_helpers import safe_delete_file

        # 清理处理器创建的临时文件
        if temp_processed and os.path.exists(temp_processed) and temp_processed != result_path:
            if safe_delete_file(temp_processed):
                logger.info(f"已清理临时文件: {temp_processed}")

        # 清理原始上传文件（由文件服务处理）
        # 注意：这里不直接删除，由文件服务的清理机制处理

    except Exception as e:
        logger.warning(f"清理临时文件失败: {e}")

    download_name = generate_output_filename('calc', original_filename, suffix='_计算结果')
    return send_file(result_path, as_attachment=True, download_name=download_name)
    ```

- **blueprints/checker.py**
  - language: `python`
  - content:
    ```python
# blueprints/checker.py
import os, json, logging, tempfile
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import make_response
from services.checker import CheckerService
from core.task.decorator import async_task
from core.file_naming import generate_output_filename

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
    try:
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
        return make_response(True, res)
    finally:
        # 修复 #7: 确保临时文件始终清理
        for p in saved_a + saved_b:
            try:
                os.remove(p)
            except Exception:
                logger.warning(f"清理临时文件失败: {p}")

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

    # 修复 #8: request.form['key_map_a'] → .get() + 校验
    key_map_a_raw = request.form.get('key_map_a')
    if not key_map_a_raw:
        raise ValidationError("缺少主键列映射参数 key_map_a")
    key_map_a = [k.strip() for k in key_map_a_raw.split(',') if k.strip()]

    # 修复 #8: request.form['compare_map'] → .get() + 默认
    compare_map = json.loads(request.form.get('compare_map', '[]'))

    # 原始逻辑保留：key_map_b 空字符串 → key_map_a
    key_map_b_str = request.form.get('key_map_b', '')
    key_map_b = [k.strip() for k in key_map_b_str.split(',') if k.strip()] if key_map_b_str else key_map_a

    filters_a = json.loads(request.form.get('filters_a', '[]'))
    filters_b = json.loads(request.form.get('filters_b', '[]'))
    dup_action = request.form.get('duplicate_action', 'error')

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
            try:
                service = CheckerService(
                    files_a=saved_a, files_b=saved_b,
                    key_map_a=key_map_a, key_map_b=key_map_b,
                    compare_map=compare_map, filters_a=filters_a, filters_b=filters_b,
                    duplicate_action=dup_action, config=cfg,
                    progress_callback=progress_cb,
                    precheck_data=pre_data
                )
                result_path = service.execute()
                return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}
            finally:
                # 修复 #7: 闭包内清理临时文件
                for p in saved_a + saved_b:
                    try:
                        os.remove(p)
                    except Exception:
                        logger.warning(f"清理临时文件失败: {p}")

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
    original_filename = data.get('original_filename', '校验结果.xlsx')
    download_name = generate_output_filename('checker', original_filename, suffix='_校验结果')
    return send_file(path, as_attachment=True, download_name=download_name)

@checker_bp.route('/')
def index():
    return render_template('checker/index.html')
    ```

- **blueprints/comparator.py**
  - language: `python`
  - content:
    ```python
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
    ```

- **blueprints/convert.py**
  - language: `python`
  - content:
    ```python
# blueprints/convert.py
import os, tempfile, logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, jsonify, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import make_response
from services.convert.service import ConvertService
from core.task.decorator import async_task
from core.file_naming import generate_output_filename

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
            try:
                service = ConvertService(
                    file_path=save_path,
                    target_format=target_fmt,
                    encoding=encoding,
                    separator=separator,
                    progress_callback=progress_cb
                )
                result_path = service.execute()
                return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}
            finally:
                # 修复 #7: 闭包内清理临时文件和目录
                try:
                    if os.path.exists(save_path):
                        os.remove(save_path)
                except Exception:
                    logger.warning(f"清理临时文件失败: {save_path}")
                try:
                    import shutil
                    shutil.rmtree(upload_dir, ignore_errors=True)
                except Exception:
                    logger.warning(f"清理临时目录失败: {upload_dir}")

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
    original = data.get('original_filename', 'file')
    safe_name = generate_output_filename('convert', original, suffix='_转换结果')
    return send_file(path, as_attachment=True, download_name=safe_name)
    ```

- **blueprints/field_parser.py**
  - language: `python`
  - content:
    ```python
# blueprints/field_parser.py
import io, logging, pandas as pd
from flask import Blueprint, render_template, request, jsonify, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.csv_utils import split_fields
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
field_parser_bp = Blueprint('field_parser', __name__, url_prefix='/field-parser', template_folder='../templates/field_parser')

@field_parser_bp.route('/')
def index():
    return render_template('field_parser/index.html')

@field_parser_bp.route('/analyze', methods=['GET', 'POST'])
@handle_exceptions()
def analyze():
    if request.method == 'GET':
        return jsonify({'message': '请使用 POST 方法提交 JSON 数据', 'fields': {'standard_headers': '标准列名列表（文本）', 'compare_headers': '比较列名列表（文本）'}})

    if not request.is_json:
        raise ValidationError("仅支持 JSON 格式的请求")

    data = request.get_json()
    text_a = data.get('standard_headers', '')
    text_b = data.get('compare_headers', '')
    if not text_a or not text_b:
        raise ValidationError("请提供 standard_headers 和 compare_headers 字段")

    service = current_app.container.field_parser()
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
        rows.append({'标准字段': item['standard'], '比较字段': item['compare'], '相似度': item['score'], '备注': '自动匹配'})
    for col in only_a:
        rows.append({'标准字段': col, '比较字段': '', '相似度': '', '备注': '仅标准'})
    for col in only_b:
        rows.append({'标准字段': '', '比较字段': col, '相似度': '', '备注': '仅比较'})

    df_mapping = pd.DataFrame(rows, columns=['标准字段', '比较字段', '相似度', '备注'])

    max_len = max(len(cols_a), len(cols_b))
    cols_a_padded = list(cols_a) + [''] * (max_len - len(cols_a))
    cols_b_padded = list(cols_b) + [''] * (max_len - len(cols_b))
    df_original = pd.DataFrame({'标准文件列名': cols_a_padded, '比较文件列名': cols_b_padded})

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_mapping.to_excel(writer, sheet_name='字段映射', index=False)
        df_original.to_excel(writer, sheet_name='原始数据', index=False)
    output.seek(0)

    download_name = generate_output_filename('field_parser', 'mapping.xlsx', suffix='_字段映射')
    return send_file(output, as_attachment=True, download_name=download_name,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    ```

- **blueprints/matcher.py**
  - language: `python`
  - content:
    ```python
# blueprints/matcher.py
import os
import io
import logging
import tempfile
import pandas as pd
from openpyxl import Workbook
from flask import Blueprint, render_template, request, send_file, current_app

from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.file_naming import generate_output_filename
from core.csv_utils import split_fields
from core.excel_helpers import safe_delete_file
from services.excel.matcher import ExcelMatcher

logger = logging.getLogger(__name__)

match_bp = Blueprint('match', __name__, url_prefix='/match', template_folder='../templates/matcher')

ALLOWED_TARGET_EXT = {'.xlsx'}


def _check_target_extension(filename: str):
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_TARGET_EXT:
        raise ValidationError(
            f"目标文件格式不支持，请使用 .xlsx 格式（当前: {ext}）。如为 .xls 文件，请先用 Excel 另存为 .xlsx。"
        )


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
    if not raw:
        return []
    return split_fields(raw)


def _resolve_duplicate_action(form) -> str:
    """解析目标键重复时的处理方式。

    新参数 duplicate_action 优先：
        - 'error'  : 检测到重复键时报错（默认）
        - 'first'  : 只填充第一个匹配行，其余行保留原值
        - 'all'    : 填充所有匹配行
    兼容旧参数 allow_duplicate_keys='on' → 'first'（旧语义：只填第一个）
    都未传时：'error'
    """
    action = (form.get('duplicate_action') or '').strip().lower()
    if action in ('error', 'first', 'all'):
        return action
    if form.get('allow_duplicate_keys') == 'on':
        return 'first'
    return 'error'


def _write_dataframe_to_worksheet(ws, df):
    """
    将 DataFrame 写入 worksheet（替代 pandas.to_excel，避免临时文件冲突）。
    """
    # 写表头
    ws.append([str(c) for c in df.columns])
    # 写数据
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))


def _ensure_excel_compatible(filepath: str, sheet_indices=None, file_label="文件") -> str:
    """
    只提取指定的 Sheet（默认为第一个），生成新的临时 .xlsx 文件。
    使用内存缓冲区保存，避免 Windows 临时文件问题。
    """
    logger.info(f"[规范化] 开始处理 {file_label}: {filepath}, 目标Sheet索引: {sheet_indices}")

    all_sheet_names = []
    try:
        import openpyxl as _oxl
        wb = _oxl.load_workbook(filepath, read_only=True, data_only=True)
        all_sheet_names = wb.sheetnames
        wb.close()
        logger.info(f"[规范化] 成功获取所有Sheet名称: {all_sheet_names}")
    except Exception as e:
        logger.warning(f"[规范化] 无法用 openpyxl 读取工作表名称: {e}，将使用 pandas 直接读取指定 Sheet")

    if sheet_indices is None:
        target_indices = [0]
    else:
        target_indices = [idx - 1 for idx in sheet_indices]

    dfs = {}
    failed_sheets = []

    for idx in target_indices:
        logger.info(f"[规范化] 尝试读取 {file_label} 的Sheet索引 {idx}（第{idx+1}个）")
        try:
            df = pd.read_excel(filepath, sheet_name=idx, dtype=str, engine='openpyxl')
            dfs[idx] = df
            logger.info(f"[规范化] {file_label} Sheet索引 {idx} 读取成功，行数: {len(df)}")
        except Exception as e:
            failed_sheets.append(f"第{idx+1}个Sheet")
            logger.warning(f"[规范化] 读取 {file_label} Sheet索引 {idx} 失败: {e}")

    if not dfs:
        raise ValidationError(
            f"{file_label} 无法读取所需的 Sheet（{', '.join(failed_sheets)}）。"
            "请确认这些 Sheet 未损坏且为有效的 Excel 表格。"
        )

    # === 使用 BytesIO 在内存中保存，避免 Windows 临时文件问题 ===
    wb_new = Workbook()
    default_sheet = wb_new.active
    wb_new.remove(default_sheet)

    for idx, df in dfs.items():
        try:
            sheet_name = all_sheet_names[idx] if idx < len(all_sheet_names) else f"Sheet{idx+1}"
        except Exception:
            sheet_name = f"Sheet{idx+1}"
        sheet_name = str(sheet_name).strip()[:31] or f"Sheet{idx+1}"
        for ch in r'[]:*?/\\':
            sheet_name = sheet_name.replace(ch, '_')

        ws = wb_new.create_sheet(title=sheet_name)
        _write_dataframe_to_worksheet(ws, df)
        logger.info(f"[规范化] 已写入 {file_label} Sheet: {sheet_name}")

    # 保存到内存缓冲区
    buffer = io.BytesIO()
    wb_new.save(buffer)
    wb_new.close()
    buffer.seek(0)

    # 将内存中的 Excel 写入一个安全的磁盘文件（供后续 openpyxl 处理）
    tmp_dir = tempfile.mkdtemp(prefix='matcher_norm_')
    new_path = os.path.join(tmp_dir, 'normalized.xlsx')
    with open(new_path, 'wb') as f:
        f.write(buffer.getvalue())

    logger.info(f"[规范化] {file_label} 新临时文件生成: {new_path}")
    return new_path


@match_bp.route('/')
def index():
    return render_template('matcher/index.html')


@match_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    logger.info("[匹配填充] 收到 POST 请求")
    file_service = current_app.container.file_service()

    target_file = request.files.get('target_file')
    source_file = request.files.get('source_file')
    if not target_file or not source_file:
        raise ValidationError('请上传两个文件')

    _check_target_extension(target_file.filename)

    logger.info("[匹配填充] 保存上传文件...")
    target_raw_path = file_service.save_upload(target_file)
    source_raw_path = file_service.save_upload(source_file)

    target_raw_path = _fix_target_extension(target_raw_path)

    source_sheets_raw = request.form.get('source_sheets', '').strip()
    sheet_indices = None
    if source_sheets_raw:
        try:
            sheet_indices = [int(x.strip()) for x in source_sheets_raw.split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            raise ValidationError('Sheet索引格式错误，请输入逗号分隔的数字，例如 1,2,3')

    temp_files_to_clean = []

    try:
        logger.info("[匹配填充] 开始规范化目标文件（只读第1个Sheet）")
        target_path = _ensure_excel_compatible(target_raw_path, sheet_indices=[1], file_label="目标文件")
        temp_files_to_clean.append(target_path)
        temp_files_to_clean.append(target_raw_path)

        logger.info("[匹配填充] 开始规范化源文件")
        source_path = _ensure_excel_compatible(source_raw_path, sheet_indices=sheet_indices, file_label="源文件")
        temp_files_to_clean.append(source_path)
        temp_files_to_clean.append(source_raw_path)
    except ValidationError as e:
        for p in temp_files_to_clean:
            safe_delete_file(p)
        raise

    target_key_cols_raw = request.form.get('target_key_cols', '').strip()
    source_match_cols_raw = request.form.get('source_match_cols', '').strip()
    source_value_cols_raw = request.form.get('source_value_col', '').strip()
    result_cols_raw = request.form.get('result_col', '').strip()
    duplicate_action = _resolve_duplicate_action(request.form)

    if not target_key_cols_raw or not source_match_cols_raw or not source_value_cols_raw or not result_cols_raw:
        raise ValidationError('请填写完整配置项')

    target_key_cols = _parse_multi_cols(target_key_cols_raw)
    source_match_cols = _parse_multi_cols(source_match_cols_raw)
    source_value_cols = _parse_multi_cols(source_value_cols_raw)
    result_cols = _parse_multi_cols(result_cols_raw)

    if len(source_value_cols) != len(result_cols):
        raise ValidationError(f'源文件值列数量({len(source_value_cols)})与结果列数量({len(result_cols)})不一致')
    if not target_key_cols or not source_match_cols:
        raise ValidationError('主键列和匹配列不能为空')
    if len(target_key_cols) != len(source_match_cols):
        raise ValidationError(
            f'目标主键列数量({len(target_key_cols)})与源匹配列数量({len(source_match_cols)})不一致'
        )

    config = {
        'target_key_cols': target_key_cols,
        'source_match_cols': source_match_cols,
        'source_value_cols': source_value_cols,
        'result_cols': result_cols,
        'duplicate_action': duplicate_action,
        'generate_report': request.form.get('generate_report', 'on') != 'off',
    }

    logger.info("[匹配填充] 开始执行 ExcelMatcher.match()")
    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    success, result = matcher.match()
    if not success:
        error_msg = result if isinstance(result, str) else '\n'.join(result)
        raise ValidationError(error_msg)

    for p in temp_files_to_clean:
        if p and os.path.exists(p) and p != result:
            safe_delete_file(p)

    download_name = generate_output_filename('match', target_file.filename, suffix='_匹配结果')
    logger.info("[匹配填充] 处理完成，返回文件")
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

    target_raw_path = file_service.save_upload(target_file)
    target_raw_path = _fix_target_extension(target_raw_path)
    source_raw_path = file_service.save_upload(source_file)

    source_sheets_raw = request.form.get('source_sheets', '').strip()
    sheet_indices = None
    if source_sheets_raw:
        try:
            sheet_indices = [int(x.strip()) for x in source_sheets_raw.split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            return render_template('error.html', error='Sheet索引格式错误，请输入逗号分隔的数字'), 400

    temp_files_to_clean = []

    try:
        target_path = _ensure_excel_compatible(target_raw_path, sheet_indices=[1], file_label="目标文件")
        temp_files_to_clean.append(target_path)
        temp_files_to_clean.append(target_raw_path)

        source_path = _ensure_excel_compatible(source_raw_path, sheet_indices=sheet_indices, file_label="源文件")
        temp_files_to_clean.append(source_path)
        temp_files_to_clean.append(source_raw_path)
    except ValidationError as e:
        for p in temp_files_to_clean:
            safe_delete_file(p)
        raise

    target_key_cols_raw = request.form.get('target_key_cols', '').strip()
    source_match_cols_raw = request.form.get('source_match_cols', '').strip()
    source_value_cols_raw = request.form.get('source_value_col', '').strip()
    result_cols_raw = request.form.get('result_col', '').strip()
    duplicate_action = _resolve_duplicate_action(request.form)

    if not target_key_cols_raw or not source_match_cols_raw or not source_value_cols_raw or not result_cols_raw:
        return render_template('error.html', error='请填写完整配置项'), 400

    target_key_cols = _parse_multi_cols(target_key_cols_raw)
    source_match_cols = _parse_multi_cols(source_match_cols_raw)
    source_value_cols = _parse_multi_cols(source_value_cols_raw)
    result_cols = _parse_multi_cols(result_cols_raw)

    if len(source_value_cols) != len(result_cols):
        return render_template(
            'error.html',
            error=f'源文件值列数量({len(source_value_cols)})与结果列数量({len(result_cols)})不一致'
        ), 400

    config = {
        'target_key_cols': target_key_cols,
        'source_match_cols': source_match_cols,
        'source_value_cols': source_value_cols,
        'result_cols': result_cols,
        'duplicate_action': duplicate_action,
    }

    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    info = matcher.debug_info()
    if 'error' in info:
        return render_template('error.html', error=info['error']), 400

    for p in temp_files_to_clean:
        safe_delete_file(p)

    return render_template('debug.html', config=config, info=info)
    ```

- **blueprints/splitter.py**
  - language: `python`
  - content:
    ```python
# blueprints/splitter.py
import os
import tempfile
import zipfile
import logging
from pathlib import Path
from flask import Blueprint, render_template, request, send_file, current_app
from openpyxl import load_workbook, Workbook
from copy import copy
import pandas as pd

from core.validators import sanitize_filename
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.filter_engine import parse_conditions_from_form, ConditionEngine
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
splitter_bp = Blueprint('splitter', __name__, url_prefix='/splitter', template_folder='../templates/splitter')

def copy_cell(source_cell, target_cell):
    try:
        if hasattr(source_cell, 'value') and source_cell.value is not None:
            target_cell.value = source_cell.value
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = source_cell.number_format
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.protection = copy(source_cell.protection)
    except AttributeError:
        pass

def copy_worksheet_without_data(source_ws, target_wb, keep_rows_set):
    target_ws = target_wb.create_sheet(title=source_ws.title)
    if keep_rows_set is None:
        keep_rows_set = set(range(1, source_ws.max_row + 1))
    old_to_new = {}
    new_row = 1
    for old_row in range(1, source_ws.max_row + 1):
        if old_row in keep_rows_set:
            old_to_new[old_row] = new_row
            new_row += 1
    for col in source_ws.column_dimensions:
        target_ws.column_dimensions[col].width = source_ws.column_dimensions[col].width
    for old_row, new_row_num in old_to_new.items():
        if old_row in source_ws.row_dimensions:
            target_ws.row_dimensions[new_row_num].height = source_ws.row_dimensions[old_row].height
    merged_cells_to_apply = []
    for merged_range in source_ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
        rows_in_range = set(range(min_row, max_row + 1))
        if not rows_in_range.intersection(keep_rows_set):
            continue
        new_min_row = old_to_new.get(min_row, min_row)
        new_max_row = old_to_new.get(max_row, max_row)
        if new_min_row and new_max_row:
            merged_cells_to_apply.append((new_min_row, min_col, new_max_row, max_col))
    max_col = source_ws.max_column
    is_top_left = {}
    for merged_range in source_ws.merged_cells.ranges:
        min_row, min_col = merged_range.min_row, merged_range.min_col
        is_top_left[(min_row, min_col)] = True
    for old_row, new_row_num in old_to_new.items():
        for col in range(1, max_col + 1):
            if (old_row, col) in is_top_left:
                pass
            else:
                skip = False
                for merged_range in source_ws.merged_cells.ranges:
                    if (merged_range.min_row <= old_row <= merged_range.max_row and
                        merged_range.min_col <= col <= merged_range.max_col):
                        if old_row != merged_range.min_row or col != merged_range.min_col:
                            skip = True
                            break
                if skip:
                    continue
            src_cell = source_ws.cell(row=old_row, column=col)
            dst_cell = target_ws.cell(row=new_row_num, column=col)
            copy_cell(src_cell, dst_cell)
    for min_row, min_col, max_row, max_col in merged_cells_to_apply:
        if min_row < max_row or min_col < max_col:
            target_ws.merge_cells(start_row=min_row, start_column=min_col,
                                  end_row=max_row, end_column=max_col)
    return target_ws

def repair_excel_with_pandas(input_path):
    import pandas as pd
    temp_xlsx = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_xlsx.close()
    try:
        with pd.ExcelFile(input_path) as xls:
            sheet_names = xls.sheet_names
            wb = Workbook()
            wb.remove(wb.active)
            for sheet_name in sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, keep_default_na=False)
                ws = wb.create_sheet(title=sheet_name)
                for col_idx, col_name in enumerate(df.columns, 1):
                    ws.cell(row=1, column=col_idx, value=col_name)
                for row_idx, row in enumerate(df.values, 2):
                    for col_idx, val in enumerate(row, 1):
                        ws.cell(row=row_idx, column=col_idx, value=val)
            wb.save(temp_xlsx.name)
    except Exception as e:
        try:
            os.unlink(temp_xlsx.name)
        except:
            pass
        raise e
    return temp_xlsx.name

def split_excel_to_temp(input_path, split_column_name, include_values, exclude_values, skip_sheet_indices, temp_dir):
    wb = load_workbook(input_path, data_only=True)
    sheets = wb.worksheets
    first_sheet = sheets[0]
    headers = [cell.value for cell in first_sheet[1]]
    if split_column_name not in headers:
        wb.close()
        raise ValidationError(f"拆分列 '{split_column_name}' 不在第一个工作表的表头中")
    split_col_idx = headers.index(split_column_name) + 1

    all_values = set()
    for idx, sheet in enumerate(sheets):
        if (idx + 1) in skip_sheet_indices:
            continue
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=split_col_idx)
            if cell.value is not None:
                val = str(cell.value).strip()
                if val:
                    all_values.add(val)

    if include_values:
        target_values = [v for v in all_values if v in include_values]
    elif exclude_values:
        target_values = [v for v in all_values if v not in exclude_values]
    else:
        target_values = list(all_values)

    if not target_values:
        wb.close()
        raise ValidationError("没有找到任何需要拆分的值")

    generated_files = []
    for split_val in target_values:
        safe_val = sanitize_filename(str(split_val))
        new_wb = Workbook()
        default_sheet = new_wb.active
        new_wb.remove(default_sheet)
        for idx, src_ws in enumerate(sheets):
            if (idx + 1) in skip_sheet_indices:
                continue
            keep_rows = {1}
            for row in range(2, src_ws.max_row + 1):
                cell_val = src_ws.cell(row=row, column=split_col_idx).value
                if cell_val is not None and str(cell_val) == str(split_val):
                    keep_rows.add(row)
            copy_worksheet_without_data(src_ws, new_wb, keep_rows)
        out_file = Path(temp_dir) / f"{safe_val}.xlsx"
        new_wb.save(str(out_file))
        new_wb.close()
        generated_files.append(str(out_file))
    wb.close()
    return generated_files

def filter_workbook_inplace(input_path, form, skip_sheet_indices=None):
    if skip_sheet_indices is None:
        skip_sheet_indices = set()

    exact, fuzzy = parse_conditions_from_form(form)
    if not exact and not fuzzy:
        return input_path

    engine = ConditionEngine()

    wb = load_workbook(input_path, data_only=True)
    try:
        for idx, ws in enumerate(wb.worksheets, start=1):
            if idx in skip_sheet_indices:
                continue

            headers = [cell.value for cell in ws[1]]
            missing_fields = []
            for cond in exact + fuzzy:
                if cond.get('field', '') not in headers:
                    missing_fields.append(cond.get('field', ''))
            if missing_fields:
                logger.warning(f"工作表 '{ws.title}' 缺少字段 {missing_fields}，跳过筛选")
                continue

            col_map = {h: i + 1 for i, h in enumerate(headers)}
            rows_to_delete = []
            for row_idx in range(2, ws.max_row + 1):
                row_dict = {}
                for col_name, col_idx in col_map.items():
                    cell = ws.cell(row=row_idx, column=col_idx)
                    row_dict[col_name] = str(cell.value) if cell.value is not None else ''
                if not engine.apply_to_row(row_dict, exact, fuzzy):
                    rows_to_delete.append(row_idx)

            for row_idx in reversed(rows_to_delete):
                ws.delete_rows(row_idx)

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        wb.save(tmp.name)
        return tmp.name
    finally:
        wb.close()

# ---------- 路由 ----------
@splitter_bp.route('/', methods=['GET'])
def index():
    return render_template('splitter/index.html')

@splitter_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    if 'file' not in request.files:
        raise ValidationError('未上传文件')
    file = request.files['file']
    if file.filename == '':
        raise ValidationError('文件名为空')

    split_column = request.form.get('split_column', '').strip()
    if not split_column:
        raise ValidationError('请指定拆分列名')

    include_values_str = request.form.get('include_values', '').strip()
    exclude_values_str = request.form.get('exclude_values', '').strip()
    skip_sheets_str = request.form.get('skip_sheets', '').strip()

    include_values = [v.strip() for v in include_values_str.split(',') if v.strip()] if include_values_str else []
    exclude_values = [v.strip() for v in exclude_values_str.split(',') if v.strip()] if exclude_values_str else []
    skip_sheet_indices = {int(v.strip()) for v in skip_sheets_str.split(',') if v.strip().isdigit()} if skip_sheets_str else set()

    if include_values and exclude_values:
        raise ValidationError('正选和反选不能同时使用')

    save_path = file_service.save_upload(file, subdir='splitter', allowed_extensions={'xlsx','xls','xlsm'})
    if not save_path.lower().endswith('.xlsx'):
        new_path = save_path + '.xlsx'
        os.rename(save_path, new_path)
        save_path = new_path

    try:
        load_workbook(save_path, data_only=True)
    except Exception:
        logger.warning("openpyxl 打开失败，尝试用 pandas 修复（公式将丢失）")
        try:
            repaired_path = repair_excel_with_pandas(save_path)
            os.remove(save_path)
            save_path = repaired_path
        except Exception as repair_err:
            raise ValidationError(f'文件无法被识别为有效的 Excel 文件: {str(repair_err)}')

    filter_path = filter_workbook_inplace(save_path, request.form, skip_sheet_indices)
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            generated_files = split_excel_to_temp(
                input_path=filter_path,
                split_column_name=split_column,
                include_values=include_values,
                exclude_values=exclude_values,
                skip_sheet_indices=skip_sheet_indices,
                temp_dir=temp_dir
            )
            zip_buffer = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
            zip_buffer.close()
            with zipfile.ZipFile(zip_buffer.name, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fpath in generated_files:
                    arcname = os.path.basename(fpath)
                    zf.write(fpath, arcname)

            from datetime import datetime
            original_name = os.path.splitext(file.filename)[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            download_name = generate_output_filename('splitter', file.filename, suffix='_拆分结果')
            return send_file(
                zip_buffer.name,
                as_attachment=True,
                download_name=download_name,
                mimetype='application/zip'
            )
    finally:
        if filter_path != save_path:
            try:
                os.remove(filter_path)
            except:
                pass
    ```

- **blueprints/subtotal.py**
  - language: `python`
  - content:
    ```python
# blueprints/subtotal.py
import os
import logging
import openpyxl
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.task.decorator import async_task
from services.subtotal.engine import SubtotalEngine, _build_header_map, resolve_single_col
from core.csv_utils import split_fields
from core.file_naming import generate_output_filename, build_content_disposition

logger = logging.getLogger(__name__)
subtotal_bp = Blueprint('subtotal', __name__, url_prefix='/subtotal',
                        template_folder='../templates/subtotal')

@subtotal_bp.route('/')
def index():
    return render_template('subtotal/index.html')

def _check_columns(raw, category, header_map, max_col):
    if not raw:
        return
    items = split_fields(raw)
    for item in items:
        if '-' in item:
            parts = item.split('-')
            if len(parts) != 2:
                raise ValidationError(f"{category}列范围格式错误：{item}")
            start, end = parts[0].strip(), parts[1].strip()
            start_idx = resolve_single_col(start, header_map)
            end_idx = resolve_single_col(end, header_map)
            if start_idx > max_col:
                raise ValidationError(f"{category}列 '{start}' 超出工作表列范围（共{max_col}列）")
            if end_idx > max_col:
                raise ValidationError(f"{category}列 '{end}' 超出工作表列范围（共{max_col}列）")
        else:
            idx = resolve_single_col(item, header_map)
            if idx > max_col:
                raise ValidationError(f"{category}列 '{item}' 超出工作表列范围（共{max_col}列）")

@subtotal_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file = request.files.get('file')
    if not file or file.filename == '':
        raise ValidationError('请上传文件')

    original_filename = file.filename
    key_cols_raw = request.form.get('key_cols', '').strip()
    if not key_cols_raw:
        raise ValidationError('请指定主键列')

    dimension_cols_raw = request.form.get('dimension_cols', '').strip() or None
    calc_cols_raw = request.form.get('calc_cols', '').strip()
    if not calc_cols_raw:
        raise ValidationError('请指定计算列')

    agg_func_default = request.form.get('agg_func_default', 'sum').strip()
    agg_overrides_raw = request.form.get('agg_overrides', '').strip()
    sheets_mode = request.form.get('sheets_mode', '1').strip()

    file_service = current_app.container.file_service()
    path = file_service.save_upload(file, allowed_extensions={'xlsx', 'xls'})

    wb = openpyxl.load_workbook(path)
    first_ws = wb.active
    max_col = first_ws.max_column
    header_map = _build_header_map(first_ws)
    wb.close()

    _check_columns(key_cols_raw, '主键', header_map, max_col)
    _check_columns(calc_cols_raw, '计算', header_map, max_col)
    if dimension_cols_raw:
        _check_columns(dimension_cols_raw, '维度', header_map, max_col)

    agg_funcs = {'default': agg_func_default}
    if agg_overrides_raw:
        for part in agg_overrides_raw.split(','):
            if ':' in part:
                col, func = part.split(':', 1)
                col = col.strip().upper()
                func = func.strip().lower()
                if func not in ('sum', 'avg', 'min', 'max', 'count'):
                    raise ValidationError(f"不支持的聚合函数: {func}")
                try:
                    openpyxl.utils.column_index_from_string(col)
                except:
                    raise ValidationError(f"无效的列字母: {col}")
                agg_funcs[col] = func

    wb2 = openpyxl.load_workbook(path, read_only=True)
    total_rows = sum(sheet.max_row for sheet in wb2.worksheets)
    wb2.close()
    data_rows = total_rows - 1

    app = current_app._get_current_object()

    if data_rows > 5000:
        task_storage = app.container.task_storage()

        @async_task(original_filename='subtotal')
        def run_subtotal(task_id):
            with app.app_context():
                def progress_cb(percent, message):
                    task_storage.update(task_id, progress_percent=percent, progress_message=message)
                try:
                    progress_cb(0, "开始处理...")
                    engine = SubtotalEngine(
                        path, key_cols_raw, calc_cols_raw, agg_funcs,
                        dimension_cols=dimension_cols_raw,
                        sheets=sheets_mode,
                        original_filename=original_filename
                    )
                    out_path = engine.process()
                    progress_cb(100, "完成")
                    # 保存原始文件名到任务数据中，便于下载时生成统一名称
                    return {'result_path': out_path, 'original_filename': original_filename}
                except Exception as e:
                    import traceback
                    error_msg = f"{str(e)}\n{traceback.format_exc()}"
                    task_storage.append_log(task_id, error_msg)
                    raise
                finally:
                    file_service.cleanup([path])

        result = run_subtotal()
        return jsonify({'task_id': result['task_id']})
    else:
        try:
            engine = SubtotalEngine(
                path, key_cols_raw, calc_cols_raw, agg_funcs,
                dimension_cols=dimension_cols_raw,
                sheets=sheets_mode,
                original_filename=original_filename
            )
            out_path = engine.process()
            download_name = generate_output_filename('subtotal', original_filename, suffix='_小计')
            resp = send_file(out_path, as_attachment=True,
                             download_name=download_name,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp.headers['Content-Disposition'] = build_content_disposition(download_name)
            return resp
        finally:
            file_service.cleanup([path])

@subtotal_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data:
        return jsonify({'error': '任务不存在'}), 404
    log = storage.get_log(task_id, max_lines=100)
    data['log'] = log
    return jsonify(data)

@subtotal_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status') != 'completed':
        return jsonify({'error': '结果不可用'}), 400
    path = data.get('result_path')
    if not path or not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404

    # 使用统一命名函数生成下载名
    original_filename = data.get('original_filename', 'data.xlsx')
    download_name = generate_output_filename('subtotal', original_filename, suffix='_小计')
    resp = send_file(path, as_attachment=True, download_name=download_name)
    resp.headers['Content-Disposition'] = build_content_disposition(download_name)
    return resp
    ```

- **blueprints/summarizer.py**
  - language: `python`
  - content:
    ```python
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
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start

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

    header_row_str = request.form.get('header_row', '').strip()
    skip_header_str = request.form.get('skip_header_rows', '0').strip()

    saved_paths = []
    for file in files:
        if file and file.filename:
            path = file_service.save_upload(file)
            saved_paths.append(path)

    if not saved_paths:
        raise ValidationError('没有有效文件上传')

    # 确定表头行
    try:
        if header_row_str:
            header_row = int(header_row_str)
            if header_row < 1:
                raise ValueError
            skip_header_rows = header_row - 1
        else:
            # 自动检测第一个文件的表头
            data_start = detect_header_start(saved_paths[0])
            header_row = data_start - 1 if data_start > 1 else 1
            if header_row < 1:
                header_row = 1
            skip_header_rows = header_row - 1
    except ValueError:
        raise ValidationError('表头行号或跳过行数必须为有效正整数')

    # 解析字段
    field_list = _dedup(split_fields(field_str))
    if not field_list:
        raise ValidationError('字段列表不能为空')

    unique_keys = _dedup(split_fields(unique_str))
    sum_fields = _dedup(split_fields(sum_str))

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
            combined[valid_keys] = combined[valid_keys].fillna('')
            result = compute_engine.aggregate_by_keys(combined, valid_keys, sum_fields)
            result[valid_keys] = result[valid_keys].fillna('')

    result = data_formatter.round_numeric_columns(result)

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    data_writer.write_excel(result, tmp_path)

    download_name = generate_output_filename('summary', files[0].filename, suffix='_汇总结果')
    return send_file(tmp_path, as_attachment=True, download_name=download_name)
    ```

- **blueprints/team_report.py**
  - language: `python`
  - content:
    ```python
# blueprints/team_report.py
import json, os, logging, tempfile, pandas as pd
from flask import Blueprint, render_template, request, jsonify, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from datetime import datetime
from core.task.decorator import async_task
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
team_report_bp = Blueprint('team_report', __name__, url_prefix='/team-report', template_folder='../templates/team_report')

def _save_uploaded(file, prefix='tr_'):
    if not file or file.filename == '': return None
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
    monthly_files = request.files.getlist('monthly_files')

    if not today_file:
        raise ValidationError("请上传今日司机流水文件")

    mapping_path = _save_uploaded(mapping_file, prefix='mapping_') if mapping_file else None
    today_path = _save_uploaded(today_file, prefix='today_')
    yesterday_path = _save_uploaded(yesterday_file, prefix='yesterday_') if yesterday_file else None
    monthly_paths = [_save_uploaded(mf, prefix='monthly_') for mf in monthly_files if mf and mf.filename]

    daily_order_target = request.form.get('daily_order_target') or None
    daily_apply_target = request.form.get('daily_apply_target') or None
    monthly_order_target = request.form.get('monthly_order_target') or None
    monthly_apply_target = request.form.get('monthly_apply_target') or None

    app = current_app._get_current_object()

    @async_task(original_filename='team_report')
    def run_report(task_id):
        with app.app_context():
            service = app.container.team_report_service()
            task_storage = app.container.task_storage()
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

                return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}
            finally:
                for p in [mapping_path, today_path, yesterday_path] + monthly_paths:
                    if p and os.path.exists(p):
                        try: os.remove(p)
                        except: pass

    result = run_report()
    return jsonify({'task_id': result['task_id']})

@team_report_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data: return jsonify({'error': '任务不存在'}), 404
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
    original = data.get('original_filename', 'team_report.xlsx')
    download_name = generate_output_filename('team_report', original, suffix='_每日数据对比')
    response = send_file(path, as_attachment=True, download_name=download_name,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         conditional=True)
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    response.headers['X-Accel-Buffering'] = 'no'
    return response
    ```

- **blueprints/transposer.py**
  - language: `python`
  - content:
    ```python
import io, logging, pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.csv_utils import split_fields
from services.transposer.service import pivot_table
from core.file_naming import generate_output_filename

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

    try:
        df = pd.read_excel(file, dtype=str)
    except Exception as e:
        raise ValidationError(f'无法读取文件: {str(e)}')

    all_needed = [category_col] + value_cols + (index_cols if index_cols else [])
    missing = [c for c in all_needed if c not in df.columns]
    if missing:
        raise ValidationError(f'列不存在: {", ".join(missing)}')

    try:
        result_df = pivot_table(df, index_cols, category_col, value_cols)
    except Exception as e:
        raise ValidationError(f'转换失败: {str(e)}')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='转置结果')
    output.seek(0)

    download_name = generate_output_filename('transposer', file.filename, suffix='_转置结果')
    return send_file(output, as_attachment=True, download_name=download_name,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    ```

- **core/__init__.py**
  - language: `python`
  - content:
    ```python
# core/__init__.py
"""核心基础设施层"""
from .exceptions import *
from .config import settings
from .expression_engine import ExpressionEngine, get_expression_engine
from .data_access_layer import DataAccessLayer, get_data_access_layer
    ```

