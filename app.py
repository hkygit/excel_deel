#!/usr/bin/env python3
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

    return app

app = create_app()

if __name__ == '__main__':
    from waitress import serve
    try:
        port = int(os.environ.get('PORT', 5000))
        logger.info(f"启动服务...端口 {port}")
        serve(app, host='0.0.0.0', port=port,
              max_request_body_size=2*1024*1024*1024,
              channel_timeout=3600, connection_limit=100,
              clear_untrusted_proxy_headers=True, _quiet=False)
    except Exception as e:
        logger.exception("服务启动失败")
