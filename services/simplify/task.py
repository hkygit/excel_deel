# services/simplify/task.py
import os
from typing import List, Dict
from core.task.base import Task
from flask import current_app

class SimplifyTask(Task):
    """异步聚合任务，封装原有 _run_task 逻辑"""
    def __init__(self, task_id: str, header_path: str, data_paths: List[str],
                 params: Dict, original_filename: str, app):
        self.task_id = task_id
        self.header_path = header_path
        self.data_paths = data_paths
        self.params = params
        self.original_filename = original_filename
        self.app = app

    def execute(self):
        with self.app.app_context():
            storage = self.app.container.task_storage()
            manager = self.app.container.task_manager()
            try:
                storage.append_log(self.task_id, "开始处理...")
                service = self.app.container.simplify_service(
                    self.header_path, self.data_paths, self.params, self.original_filename
                )
                result_path, result_format, input_rows, output_rows = service.execute()
                manager.complete_task(
                    self.task_id,
                    result_path=result_path,
                    result_format=result_format,
                    total_input_rows=input_rows,
                    total_output_rows=output_rows
                )
                storage.append_log(
                    self.task_id,
                    f"处理完成，输出文件: {result_path}，格式: {result_format}，输入行数: {input_rows}，输出行数: {output_rows}"
                )
            except Exception as e:
                import traceback
                error_msg = f"{str(e)}\n{traceback.format_exc()}"
                storage.append_log(self.task_id, error_msg)
                manager.fail_task(self.task_id, error=error_msg)
            finally:
                all_paths = [self.header_path] + self.data_paths
                for p in all_paths:
                    try:
                        if os.path.exists(p):
                            os.remove(p)
                            storage.append_log(self.task_id, f"已删除临时文件: {p}")
                    except Exception:
                        pass
