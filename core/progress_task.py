# core/progress_task.py
from abc import ABC, abstractmethod
from typing import Optional, Callable
from core.task.base import Task

class ProgressTask(Task, ABC):
    """支持进度回调的任务基类"""
    def __init__(self, progress_callback: Optional[Callable] = None):
        self.progress_callback = progress_callback
        self._total_rows = 0
        self._processed_rows = 0

    def set_total_rows(self, total: int):
        self._total_rows = total

    def set_progress(self, percent=None, message=None, force=True):
        if self.progress_callback:
            if percent is None and self._total_rows > 0:
                percent = min(int(self._processed_rows / self._total_rows * 100), 99)
            self.progress_callback(
                percent=percent,
                message=message,
                processed_rows=self._processed_rows,
                total_rows=self._total_rows
            )
