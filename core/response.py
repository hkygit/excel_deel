# core/response.py
"""统一 JSON 响应构建工具"""
from flask import jsonify

def make_response(success=True, data=None, error=None, code=200):
    """构建统一的 JSON 响应
    :param success: 操作是否成功
    :param data: 成功时返回的数据
    :param error: 错误时的描述信息
    :param code: HTTP 状态码
    """
    body = {
        'success': success,
        'data': data,
        'error': error,
        'code': code
    }
    response = jsonify(body)
    response.status_code = code
    return response

def success_response(data=None, message='success'):
    """快捷成功响应，兼容旧接口"""
    return make_response(True, data, None, 200)

def error_response(message, code=400):
    """快捷错误响应，兼容旧接口"""
    return make_response(False, None, message, code)
