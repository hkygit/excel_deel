# core/task/manager.py
import threading
import uuid
from datetime import datetime
from typing import Callable, Dict, Optional
from .state import TaskStatus
from .storage import TaskStorage
from .logger import log_event

class TaskManager:
    def __init__(self, storage: TaskStorage, max_concurrent: int = 5):
        self.storage = storage
        self._lock = threading.Lock()
        self._running_count = 0
        self._max_concurrent = max_concurrent

    def submit(self, task_id: str, target: Callable, args: tuple) -> bool:
        with self._lock:
            if self._running_count >= self._max_concurrent:
                return False
            self._running_count += 1
        def wrapper():
            try:
                target(*args)
            finally:
                with self._lock:
                    self._running_count -= 1
        thread = threading.Thread(target=wrapper, daemon=True)
        thread.start()
        return True

    def create_task(self, params: Dict, original_filename: str = None) -> str:
        task_id = str(uuid.uuid4())
        status = {
            'task_id': task_id,
            'status': TaskStatus.PENDING,
            'created_at': datetime.now().isoformat(),
            'started_at': None,
            'completed_at': None,
            'result_path': None,
            'result_format': None,
            'error': None,
            'total_input_rows': None,
            'total_output_rows': None,
            'params': params,
            'original_filename': original_filename
        }
        self.storage.save(task_id, status)
        log_event(task_id, 'task_created', params=params)
        return task_id

    def start_task(self, task_id: str, target: Callable, args: tuple):
        self.storage.update(task_id, status=TaskStatus.RUNNING, started_at=datetime.now().isoformat())
        log_event(task_id, 'task_started')
        if not self.submit(task_id, target, args):
            self.storage.update(task_id, status=TaskStatus.FAILED, error="系统繁忙，请稍后再试")
            return False
        return True

    def complete_task(self, task_id: str, **result):
        self.storage.update(task_id, status=TaskStatus.COMPLETED, completed_at=datetime.now().isoformat(), **result)
        log_event(task_id, 'task_completed', **result)

    def fail_task(self, task_id: str, error: str):
        self.storage.update(task_id, status=TaskStatus.FAILED, error=error, completed_at=datetime.now().isoformat())
        log_event(task_id, 'task_failed', error=error)

    def cancel_task(self, task_id: str):
        self.storage.update(task_id, status=TaskStatus.CANCELLED, completed_at=datetime.now().isoformat())
        log_event(task_id, 'task_cancelled')
