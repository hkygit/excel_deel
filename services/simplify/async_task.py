# services/simplify/async_task.py
import os
from typing import List, Dict
from flask import current_app
from .task import SimplifyTask

def init_config(app):
    """保留兼容性接口，不再需要持有全局状态"""
    pass

def submit_task(header_path: str, data_paths: List[str], params: Dict,
                original_filename: str = None) -> str:
    app = current_app._get_current_object()
    manager = app.container.task_manager()
    task_id = manager.create_task(params, original_filename)
    # 直接创建 SimplifyTask 并执行（通过 manager 提交）
    task = SimplifyTask(task_id, header_path, data_paths, params, original_filename, app)
    manager.start_task(task_id, task.execute, ())
    return task_id

def load_status(task_id: str):
    app = current_app._get_current_object()
    storage = app.container.task_storage()
    return storage.load(task_id)

def get_log(task_id: str, max_lines: int = 2000) -> str:
    app = current_app._get_current_object()
    storage = app.container.task_storage()
    return storage.get_log(task_id, max_lines)
