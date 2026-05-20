#!/usr/bin/env python3
"""
一键部署“数据小计”功能模块
- 创建 blueprints/subtotal.py, services/subtotal/engine.py, 前端模板及 JS
- 修改 app.py, templates/base.html, templates/index.html
- 清理 Python 缓存
- 终止旧进程并重启服务
- 集成验证：模拟调用核心功能，检查输出
"""

import os
import sys
import shutil
import subprocess
import time
import signal

# ============================================================
# 新增文件内容定义
# ============================================================

# ---------- services/subtotal/__init__.py ----------
SUBTOTAL_INIT = '''# services/subtotal/__init__.py
'''

# ---------- services/subtotal/engine.py ----------
SUBTOTAL_ENGINE = '''# services/subtotal/engine.py
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import column_index_from_string
from core.csv_utils import split_fields

SUBTOTAL_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
SUBTOTAL_FONT = Font(bold=True)
THIN_BORDER = Border(
    left=Side('thin'), right=Side('thin'),
    top=Side('thin'), bottom=Side('thin')
)

class SubtotalEngine:
    """数据小计核心引擎"""

    def __init__(self, filepath, key_cols_str, calc_cols_str, agg_funcs,
                 dimension_cols=None, sheets='1'):
        self.filepath = filepath
        self.key_cols = self._parse_cols(key_cols_str)
        self.calc_cols = self._parse_cols(calc_cols_str)
        self.dimension_cols = self._parse_cols(dimension_cols) if dimension_cols else []
        self.agg_funcs = agg_funcs      # dict with 'default' and optional letter->func
        self.sheets = sheets            # '1', 'all', or '1,2,3'

    @staticmethod
    def _parse_cols(raw):
        if not raw:
            return []
        items = split_fields(raw)
        result = []
        for item in items:
            if '-' in item:
                start, end = item.split('-', 1)
                start_idx = column_index_from_string(start.strip())
                end_idx = column_index_from_string(end.strip())
                for i in range(start_idx, end_idx + 1):
                    result.append(i)
            else:
                result.append(column_index_from_string(item.strip()))
        return sorted(list(set(result)))

    def _get_agg_func(self, col_letter):
        if col_letter in self.agg_funcs:
            return self.agg_funcs[col_letter]
        return self.agg_funcs.get('default', 'sum')

    def process(self):
        wb = openpyxl.load_workbook(self.filepath)
        sheet_names = wb.sheetnames
        if self.sheets == 'all':
            target_sheets = sheet_names
        else:
            indices = [int(x.strip()) for x in self.sheets.split(',')
                       if x.strip().isdigit()]
            target_sheets = [sheet_names[i - 1] for i in indices
                             if 1 <= i <= len(sheet_names)]

        for sheet_name in target_sheets:
            ws = wb[sheet_name]
            self._process_sheet(ws)

        out_path = self.filepath.rsplit('.', 1)[0] + '_小计.xlsx'
        wb.save(out_path)
        wb.close()
        return out_path

    def _process_sheet(self, ws):
        max_col = ws.max_column
        data_start = 2
        max_row = ws.max_row

        # 1. 按主键列分组
        groups = {}
        for row_idx in range(data_start, max_row + 1):
            key = tuple(ws.cell(row_idx, col).value for col in self.key_cols)
            groups.setdefault(key, []).append(row_idx)

        # 2. 从下往上插入小计行
        sorted_groups = sorted(groups.items(),
                               key=lambda x: max(x[1]), reverse=True)

        for key, rows in sorted_groups:
            last_row = max(rows)          # 组内最后一行（原始行号）
            ws.insert_rows(last_row + 1)   # 在下方插入一行
            subtotal_row = last_row + 1

            # ---------- 整行样式 ----------
            for col_idx in range(1, max_col + 1):
                cell = ws.cell(subtotal_row, col_idx)
                cell.fill = SUBTOTAL_FILL
                cell.font = SUBTOTAL_FONT
                cell.border = THIN_BORDER

            # ---------- 主键列 ----------
            for i, col in enumerate(self.key_cols):
                ws.cell(subtotal_row, col).value = key[i]

            # ---------- 维度列 ----------
            for col in self.dimension_cols:
                ws.cell(subtotal_row, col).value = '合计'

            # ---------- 计算列 ----------
            for col in self.calc_cols:
                col_letter = openpyxl.utils.get_column_letter(col)
                func = self._get_agg_func(col_letter)
                values = []
                for r in rows:
                    val = ws.cell(r, col).value
                    try:
                        val = float(val)
                    except (TypeError, ValueError):
                        val = 0
                    values.append(val)

                if func == 'sum':
                    result = sum(values)
                elif func == 'avg':
                    result = sum(values) / len(values) if values else 0
                elif func == 'min':
                    result = min(values) if values else 0
                elif func == 'max':
                    result = max(values) if values else 0
                elif func == 'count':
                    result = len(values)
                else:
                    result = sum(values)
                ws.cell(subtotal_row, col).value = result
'''

# ---------- blueprints/subtotal.py ----------
SUBTOTAL_BP = '''# blueprints/subtotal.py
import os
import logging
import openpyxl
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.task.decorator import async_task
from services.subtotal.engine import SubtotalEngine
from core.csv_utils import split_fields

logger = logging.getLogger(__name__)
subtotal_bp = Blueprint('subtotal', __name__, url_prefix='/subtotal',
                        template_folder='../templates/subtotal')

@subtotal_bp.route('/')
def index():
    return render_template('subtotal/index.html')

@subtotal_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file = request.files.get('file')
    if not file or file.filename == '':
        raise ValidationError('请上传文件')

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

    # 保存上传文件
    file_service = current_app.container.file_service()
    path = file_service.save_upload(file, allowed_extensions={'xlsx', 'xls'})

    # ----- 校验列是否存在（利用 openpyxl 快速读取表头） -----
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]

    def letter_to_col(letter):
        return openpyxl.utils.column_index_from_string(letter.strip())

    # 维度列校验
    if dimension_cols_raw:
        dim_letters = split_fields(dimension_cols_raw)
        for l in dim_letters:
            idx = letter_to_col(l)
            if idx > len(headers) or headers[idx - 1] is None:
                wb.close()
                raise ValidationError(f"维度列 '{l}' 在文件中不存在")
    # 计算列校验
    calc_letters = split_fields(calc_cols_raw)
    for l in calc_letters:
        idx = letter_to_col(l)
        if idx > len(headers) or headers[idx - 1] is None:
            wb.close()
            raise ValidationError(f"计算列 '{l}' 在文件中不存在")
    # 主键列校验
    key_letters = split_fields(key_cols_raw)
    for l in key_letters:
        idx = letter_to_col(l)
        if idx > len(headers) or headers[idx - 1] is None:
            wb.close()
            raise ValidationError(f"主键列 '{l}' 在文件中不存在")
    wb.close()

    # ----- 解析聚合函数覆盖 -----
    agg_funcs = {'default': agg_func_default}
    if agg_overrides_raw:
        parts = agg_overrides_raw.split(',')
        for part in parts:
            if ':' in part:
                col, func = part.split(':', 1)
                col = col.strip().upper()
                func = func.strip().lower()
                if func not in ('sum', 'avg', 'min', 'max', 'count'):
                    raise ValidationError(f"不支持的聚合函数: {func}")
                try:
                    letter_to_col(col)
                except Exception:
                    raise ValidationError(f"无效的列字母: {col}")
                agg_funcs[col] = func

    # ----- 判断是否异步（行数 > 5000） -----
    wb2 = openpyxl.load_workbook(path, read_only=True)
    total_rows = 0
    for sheet in wb2.worksheets:
        total_rows += sheet.max_row
    wb2.close()
    data_rows = total_rows - (1 if total_rows > 0 else 0)

    app = current_app._get_current_object()

    if data_rows > 5000:
        task_storage = app.container.task_storage()

        @async_task(original_filename='subtotal')
        def run_subtotal(task_id):
            with app.app_context():
                def progress_cb(percent, message):
                    task_storage.update(task_id,
                                        progress_percent=percent,
                                        progress_message=message)
                try:
                    progress_cb(0, "开始处理...")
                    engine = SubtotalEngine(
                        path, key_cols_raw, calc_cols_raw, agg_funcs,
                        dimension_cols=dimension_cols_raw,
                        sheets=sheets_mode
                    )
                    out_path = engine.process()
                    progress_cb(100, "完成")
                    return {'result_path': out_path}
                except Exception as e:
                    import traceback
                    error_msg = f"{str(e)}\\n{traceback.format_exc()}"
                    task_storage.append_log(task_id, error_msg)
                    raise
                finally:
                    file_service.cleanup([path])

        result = run_subtotal()
        return jsonify({'task_id': result['task_id']})
    else:
        # 同步处理
        try:
            engine = SubtotalEngine(
                path, key_cols_raw, calc_cols_raw, agg_funcs,
                dimension_cols=dimension_cols_raw,
                sheets=sheets_mode
            )
            out_path = engine.process()
            return send_file(out_path, as_attachment=True,
                             download_name=os.path.basename(out_path),
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
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
    return send_file(path, as_attachment=True,
                     download_name='subtotal_result.xlsx')
'''

# ---------- templates/subtotal/index.html ----------
SUBTOTAL_HTML = '''{% extends "base.html" %}
{% block title %}数据小计{% endblock %}
{% block head %}
<style>
    .override-row { display: flex; gap: 10px; margin-bottom: 8px; align-items: center; }
    .sheets-radio { margin-right: 15px; }
</style>
{% endblock %}
{% block content %}
<h2>📊 数据小计</h2>
<p>根据主键列对数据进行分组，在每组最后一行下方插入小计行。</p>
<form id="subtotalForm" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
    </div>
    <div class="form-section">
        <h3>🔑 主键列（分组依据，支持多列和范围，如 A,B,C 或 A-C）</h3>
        <input type="text" name="key_cols" placeholder="A,B" required>
    </div>
    <div class="form-section">
        <h3>📏 维度列（在小计行显示“合计”，可选，支持多列）</h3>
        <input type="text" name="dimension_cols" placeholder="C,D">
        <div class="hint">如果指定的列不存在，将报错。</div>
    </div>
    <div class="form-section">
        <h3>🧮 计算列（支持多列和范围）</h3>
        <input type="text" name="calc_cols" placeholder="E,F,G-H" required>
    </div>
    <div class="form-section">
        <h3>⚙️ 聚合函数</h3>
        <div style="display: flex; align-items: center; gap: 10px;">
            <label>默认函数：</label>
            <select name="agg_func_default" id="aggFuncDefault">
                <option value="sum">求和</option>
                <option value="avg">平均值</option>
                <option value="min">最小值</option>
                <option value="max">最大值</option>
                <option value="count">计数</option>
            </select>
            <button type="button" class="btn btn-secondary" id="addOverrideBtn">+ 特殊列函数</button>
        </div>
        <div id="overrideContainer" style="margin-top:10px;"></div>
        <div class="hint">点击“+”为特定列设置不同的聚合函数，格式：列字母:函数，如 C:avg</div>
    </div>
    <div class="form-section">
        <h3>📄 处理Sheet</h3>
        <div>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="1" checked> 仅第一个Sheet</label>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="all"> 所有Sheet</label>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="custom"> 指定Sheet索引（逗号分隔）</label>
            <input type="text" name="sheets_custom" id="sheetsCustom" placeholder="1,2,3" style="width:150px; display:none;">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始处理</button>
</form>

<div id="statusArea" style="display:none; margin-top:20px;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <div id="downloadLink"></div>
</div>

<template id="overrideRowTemplate">
    <div class="override-row">
        <span>列：</span>
        <input type="text" class="override-col" placeholder="C" style="width:80px;">
        <span>函数：</span>
        <select class="override-func">
            <option value="sum">求和</option>
            <option value="avg">平均值</option>
            <option value="min">最小值</option>
            <option value="max">最大值</option>
            <option value="count">计数</option>
        </select>
        <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">删除</button>
    </div>
</template>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/subtotal.js') }}"></script>
{% endblock %}
'''

# ---------- static/js/subtotal.js ----------
SUBTOTAL_JS = '''// static/js/subtotal.js
(function() {
    const form = document.getElementById('subtotalForm');
    const statusArea = document.getElementById('statusArea');
    const addOverrideBtn = document.getElementById('addOverrideBtn');
    const overrideContainer = document.getElementById('overrideContainer');
    const sheetsModeRadios = document.getElementsByName('sheets_mode');
    const sheetsCustom = document.getElementById('sheetsCustom');

    addOverrideBtn.addEventListener('click', function() {
        const template = document.getElementById('overrideRowTemplate');
        const clone = template.content.cloneNode(true);
        overrideContainer.appendChild(clone);
    });

    sheetsModeRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            sheetsCustom.style.display = this.value === 'custom' ? 'inline-block' : 'none';
        });
    });

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        // 收集特殊列函数
        const overrides = [];
        document.querySelectorAll('.override-row').forEach(row => {
            const col = row.querySelector('.override-col').value.trim().toUpperCase();
            const func = row.querySelector('.override-func').value;
            if (col) {
                overrides.push(col + ':' + func);
            }
        });

        const formData = new FormData(form);
        if (overrides.length > 0) {
            formData.set('agg_overrides', overrides.join(','));
        }
        const mode = formData.get('sheets_mode');
        if (mode === 'custom') {
            const customVal = sheetsCustom.value.trim();
            formData.set('sheets_mode', customVal || '1');
        }

        const submitBtn = form.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        statusArea.style.display = 'block';
        document.getElementById('taskStatus').textContent = '提交中...';
        document.getElementById('progressBar').style.width = '0%';
        document.getElementById('downloadLink').innerHTML = '';

        try {
            const resp = await fetch('/subtotal/process', { method: 'POST', body: formData });
            if (!resp.ok) {
                const err = await resp.json();
                document.getElementById('taskStatus').textContent = '错误: ' + (err.error || err.message);
                submitBtn.disabled = false;
                return;
            }
            const contentType = resp.headers.get('Content-Type');
            if (contentType && contentType.includes('application/json')) {
                // 异步任务
                const result = await resp.json();
                const taskId = result.task_id;
                pollTask(taskId, submitBtn);
            } else {
                // 同步下载
                const blob = await resp.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'subtotal_result.xlsx';
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                document.getElementById('taskStatus').textContent = '处理完成，已下载。';
                submitBtn.disabled = false;
            }
        } catch (err) {
            document.getElementById('taskStatus').textContent = '网络错误: ' + err.message;
            submitBtn.disabled = false;
        }
    });

    function pollTask(taskId, submitBtn) {
        const interval = setInterval(async () => {
            try {
                const resp = await fetch('/subtotal/status/' + taskId);
                const status = await resp.json();
                if (status.error) {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '错误: ' + status.error;
                    submitBtn.disabled = false;
                    return;
                }
                if (status.progress_percent !== undefined) {
                    document.getElementById('progressBar').style.width = status.progress_percent + '%';
                }
                document.getElementById('taskStatus').textContent =
                    status.status === 'completed' ? '完成' : (status.status || '处理中...');
                if (status.status === 'completed') {
                    clearInterval(interval);
                    document.getElementById('downloadLink').innerHTML =
                        `<a href="/subtotal/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    submitBtn.disabled = false;
                } else if (status.status === 'failed') {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '失败';
                    submitBtn.disabled = false;
                }
            } catch (err) {
                console.error(err);
            }
        }, 2000);
    }
})();
'''

# ---------- 修改后的 app.py ----------
MODIFIED_APP = '''#!/usr/bin/env python3
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

    return app

if __name__ == '__main__':
    app = create_app()
    from waitress import serve
    try:
        logger.info("启动服务...")
        serve(app, host='0.0.0.0', port=5000,
              max_request_body_size=2*1024*1024*1024,
              channel_timeout=3600, connection_limit=100,
              clear_untrusted_proxy_headers=True, _quiet=False)
    except Exception as e:
        logger.exception("服务启动失败")
'''

# ---------- 修改后的 templates/base.html ----------
MODIFIED_BASE_HTML = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{% block title %}Excel 智能工具集{% endblock %}</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
    <script src="{{ url_for('static', filename='js/main.js') }}" defer></script>
    {% block head %}{% endblock %}
</head>
<body>
    <div class="container">
        <h1>📊 Excel 智能工具集</h1>
        <nav>
            <ul>
                <li><a href="{{ url_for('index') }}">🏠 首页</a></li>
                <li><a href="{{ url_for('calc.index') }}">🧮 列计算/聚合</a></li>
                <li><a href="{{ url_for('compare.index') }}">🔍 文件对比</a></li>
                <li><a href="{{ url_for('match.index') }}">🔗 匹配填充</a></li>
                <li><a href="{{ url_for('summary.index') }}">📋 汇总工具</a></li>
                <li><a href="{{ url_for('splitter.index') }}">✂️ 文件拆分</a></li>
                <li><a href="{{ url_for('field_parser.index') }}">🔗 字段映射</a></li>
                <li><a href="{{ url_for('checker.index') }}">✅ 文件校验</a></li>
                <li><a href="{{ url_for('convert.index') }}">🔄 格式转换</a></li>
                <li><a href="{{ url_for('subtotal.index') }}">📊 数据小计</a></li>
                <li><a href="{{ url_for('team_report.index') }}">📊 车队日报</a></li>
                <li><a href="{{ url_for('transposer.index') }}">🔄 行列转置</a></li>
            </ul>
        </nav>
        {% block content %}{% endblock %}
        <div id="message" class="message"></div>
    </div>
    {% block scripts %}{% endblock %}
</body>
</html>
'''

# ---------- 修改后的 templates/index.html ----------
MODIFIED_INDEX_HTML = '''{% extends "base.html" %}
{% block title %}首页 - Excel 工具集{% endblock %}
{% block content %}
<h2>👋 欢迎使用 Excel 智能工具集</h2>
<p style="font-size: 1.1rem; margin-bottom: 32px;">请从导航栏选择所需功能：</p>

<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 24px;">
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🧮 列计算/聚合回填</h3>
        <p>支持多列表达式计算，或按分组聚合回填结果列，保留原有格式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔍 文件对比</h3>
        <p>完整对比（主键匹配）或仅字段值匹配，输出差异报告。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔗 匹配填充</h3>
        <p>根据源文件匹配目标文件，填充值并保留格式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📋 汇总工具</h3>
        <p>合并多个文件，按唯一键汇总求和。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">✂️ 文件拆分</h3>
        <p>按列值拆分Excel为多个文件，保留合并单元格与样式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔗 字段映射</h3>
        <p>智能匹配不同文件间的列名。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">✅ 文件校验</h3>
        <p>根据主键对比两个文件群，输出差异报告。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔄 格式转换</h3>
        <p>Excel 与 CSV 互转，支持编码选择。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📊 数据小计</h3>
        <p>按指定主键列分组，在每组下方插入小计行，计算合计、平均值等。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📊 车队日报</h3>
        <p>生成车队业绩日报，含上线率、目标完成率。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔄 行列转置</h3>
        <p>长表转宽表，取消分类列。</p>
    </div>
</div>
{% endblock %}
'''

# ============================================================
# 工具函数
# ============================================================

def clean_cache(root_dir):
    removed_dirs = 0
    removed_files = 0
    for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
        for d in list(dirnames):
            if d == '__pycache__':
                full = os.path.join(dirpath, d)
                shutil.rmtree(full, ignore_errors=True)
                removed_dirs += 1
        for f in filenames:
            if f.endswith('.pyc'):
                full = os.path.join(dirpath, f)
                os.remove(full)
                removed_files += 1
    print(f"缓存清理完毕：删除 {removed_dirs} 个 __pycache__ 目录，{removed_files} 个 .pyc 文件。")

def kill_port(port):
    if os.name == 'nt':
        import subprocess
        result = subprocess.run(f'netstat -ano | findstr :{port}', shell=True,
                                capture_output=True, text=True)
        for line in result.stdout.splitlines():
            if f':{port}' in line and 'LISTENING' in line:
                pid = line.strip().split()[-1]
                subprocess.run(f'taskkill /PID {pid} /F', shell=True)
                print(f"已终止 Windows 进程 PID {pid} (端口 {port})")
    else:
        import signal
        import subprocess
        result = subprocess.run(f'lsof -ti :{port}', shell=True,
                                capture_output=True, text=True)
        pids = result.stdout.strip().split()
        for pid in pids:
            os.kill(int(pid), signal.SIGKILL)
            print(f"已终止进程 PID {pid} (端口 {port})")

def verify_deployment():
    print("正在进行集成验证...")
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from app import create_app
    app = create_app()
    app.config['TESTING'] = True
    with app.test_client() as client:
        import pandas as pd
        import io
        df = pd.DataFrame({
            'A': ['X', 'X', 'Y'],
            'B': [10, 20, 30],
            'C': [1, 2, 3]
        })
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B,C',
            'agg_func_default': 'sum',
            'dimension_cols': '',
            'sheets_mode': '1'
        }
        resp = client.post('/subtotal/process', data=data,
                           content_type='multipart/form-data')
        if resp.status_code != 200:
            print(f"验证失败：状态码 {resp.status_code}")
            return False
        result_df = pd.read_excel(io.BytesIO(resp.data), sheet_name=None)
        sheet = result_df['Sheet1']
        if len(sheet) != 5:
            print(f"验证失败：预期 5 行，实际 {len(sheet)} 行")
            return False
        print("验证成功：小计功能正常。")
        return True

def main():
    root = os.path.dirname(os.path.abspath(__file__))

    # 1. 创建新目录
    os.makedirs(os.path.join(root, 'services', ''), exist_ok=True)
    os.makedirs(os.path.join(root, 'templates', ''), exist_ok=True)

    # 2. 写入新文件
    new_files = {
        'services/subtotal/__init__.py': SUBTOTAL_INIT,
        'services/subtotal/engine.py': SUBTOTAL_ENGINE,
        'blueprints/subtotal.py': SUBTOTAL_BP,
        'templates/subtotal/index.html': SUBTOTAL_HTML,
        'static/js/subtotal.js': SUBTOTAL_JS,
    }
    for rel_path, content in new_files.items():
        full_path = os.path.join(root, rel_path)
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"已创建文件: {rel_path}")

    # 3. 修改现有文件
    modify_files = {
        'app.py': MODIFIED_APP,
        'templates/base.html': MODIFIED_BASE_HTML,
        'templates/index.html': MODIFIED_INDEX_HTML,
    }
    for rel_path, content in modify_files.items():
        full_path = os.path.join(root, rel_path)
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"已更新文件: {rel_path}")

    # 4. 清理缓存
    clean_cache(root)

    # 5. 验证功能
    if not verify_deployment():
        print("部署中止：集成验证未通过。请检查错误信息。")
        sys.exit(1)

    # 6. 重启服务
    print("正在重启服务...")
    kill_port(5000)
    time.sleep(2)
    subprocess.Popen([sys.executable, os.path.join(root, 'app.py')], cwd=root)
    print("新服务已启动 (http://127.0.0.1:5000)。")

    print("\n✅ 一键部署完成！")

if __name__ == '__main__':
    main()