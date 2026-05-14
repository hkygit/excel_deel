# core/task/decorator.py
import functools
from datetime import datetime
from flask import current_app
from .state import TaskStatus

def async_task(original_filename=None):
    """
    装饰器：将函数包装为异步任务。
    被装饰函数必须接受 task_id 作为第一个参数，并返回一个字典（将作为任务结果保存）。
    示例：
        @async_task(original_filename="my_task")
        def my_task(task_id, param1, param2):
            # 耗时操作
            return {"result_path": "/path/to/result"}
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            app = current_app._get_current_object()
            manager = app.container.task_manager()
            storage = app.container.task_storage()
            task_id = manager.create_task(params={}, original_filename=original_filename)

            def run():
                with app.app_context():
                    try:
                        storage.update(task_id, status=TaskStatus.RUNNING, started_at=datetime.now().isoformat())
                        result = func(task_id, *args, **kwargs)
                        # 移除可能由任务返回的 completed_at，因为 complete_task 会自己设置
                        if result and 'completed_at' in result:
                            result.pop('completed_at')
                        manager.complete_task(task_id, **result)
                    except Exception as e:
                        import traceback
                        error_msg = f"{str(e)}\n{traceback.format_exc()}"
                        storage.append_log(task_id, error_msg)
                        manager.fail_task(task_id, error=error_msg)

            manager.start_task(task_id, run, ())
            return {"task_id": task_id}
        return wrapper
    return decorator
