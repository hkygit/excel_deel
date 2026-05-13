# services/simplify/async_task.py
import os
from typing import List, Dict
from flask import current_app

def init_config(app):
    """保留兼容性接口，不再需要持有全局状态"""
    pass

def _run_task(task_id: str, header_path: str, data_paths: List[str],
              params: Dict, original_filename: str, app):
    """异步任务执行体，通过传入的 app 实例访问容器"""
    with app.app_context():
        storage = app.container.task_storage()
        manager = app.container.task_manager()
        try:
            storage.append_log(task_id, "开始处理...")
            service = app.container.simplify_service(header_path, data_paths, params, original_filename)
            result_path, result_format, input_rows, output_rows = service.execute()
            manager.complete_task(task_id, result_path=result_path, result_format=result_format,
                                  total_input_rows=input_rows, total_output_rows=output_rows)
            storage.append_log(task_id, f"处理完成，输出文件: {result_path}，格式: {result_format}，输入行数: {input_rows}，输出行数: {output_rows}")
        except Exception as e:
            import traceback
            error_msg = f"{str(e)}\n{traceback.format_exc()}"
            storage.append_log(task_id, error_msg)
            manager.fail_task(task_id, error=error_msg)
        finally:
            all_paths = [header_path] + data_paths
            for p in all_paths:
                try:
                    if os.path.exists(p):
                        os.remove(p)
                        storage.append_log(task_id, f"已删除临时文件: {p}")
                except Exception:
                    pass

def submit_task(header_path: str, data_paths: List[str], params: Dict,
                original_filename: str = None) -> str:
    app = current_app._get_current_object()
    manager = app.container.task_manager()
    task_id = manager.create_task(params, original_filename)
    manager.start_task(task_id, _run_task, (task_id, header_path, data_paths, params, original_filename, app))
    return task_id

def load_status(task_id: str):
    app = current_app._get_current_object()
    storage = app.container.task_storage()
    return storage.load(task_id)

def get_log(task_id: str, max_lines: int = 2000) -> str:
    app = current_app._get_current_object()
    storage = app.container.task_storage()
    return storage.get_log(task_id, max_lines)
