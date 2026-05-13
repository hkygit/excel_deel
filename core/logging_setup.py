# core/logging_setup.py
"""结构化日志配置与请求ID管理"""
import logging
import uuid
from flask import request, has_request_context, g
from datetime import datetime, timezone

def setup_logging(app):
    """配置应用日志格式"""
    # 移除默认处理器
    app.logger.handlers.clear()
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s %(name)s [%(request_id)s] %(message)s'
    ))
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.INFO)

def get_request_id():
    """获取当前请求的唯一ID，不存在则生成"""
    if has_request_context():
        if not hasattr(g, 'request_id'):
            g.request_id = str(uuid.uuid4())[:8]
        return g.request_id
    return 'no-request'

class RequestIDFilter(logging.Filter):
    """将 request_id 注入到日志记录中"""
    def filter(self, record):
        record.request_id = get_request_id()
        return True

def init_app(app):
    """在 Flask app 上配置日志和请求钩子"""
    setup_logging(app)
    # 添加过滤器，使得所有 logger 都能获得 request_id
    for logger_name in logging.root.manager.loggerDict:
        logger = logging.getLogger(logger_name)
        logger.addFilter(RequestIDFilter())
    app.logger.addFilter(RequestIDFilter())

    @app.before_request
    def start_timer():
        g.start_time = datetime.now(timezone.utc)

    @app.after_request
    def log_request(response):
        elapsed = (datetime.now(timezone.utc) - g.start_time).total_seconds()
        app.logger.info(
            "%s %s %s %.3fs", request.method, request.path, response.status_code, elapsed
        )
        return response
