# core/decorators.py
import logging
from functools import wraps
from flask import jsonify, request
from core.exceptions import AppException
from core.response import make_response

logger = logging.getLogger(__name__)

def handle_exceptions():
    """统一异常处理装饰器：捕获 AppException 并返回统一 JSON 格式"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except AppException as e:
                logger.warning(f"业务异常: {e.message}")
                return make_response(False, None, e.message, e.code)
            except Exception as e:
                logger.exception("未处理异常")
                return make_response(False, None, "服务器内部错误，请稍后重试", 500)
        return wrapper
    return decorator
