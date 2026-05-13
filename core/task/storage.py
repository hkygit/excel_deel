# core/task/storage.py
import json
import os
import threading
from datetime import datetime
from typing import Dict, Optional
import logging

logger = logging.getLogger(__name__)

class TaskStorage:
    def __init__(self, status_dir: str, output_dir: str):
        self.status_dir = status_dir
        self.output_dir = output_dir
        self._lock = threading.RLock()
        os.makedirs(status_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

    def _path(self, task_id: str) -> str:
        return os.path.join(self.status_dir, f"{task_id}.json")

    def _log_path(self, task_id: str) -> str:
        return os.path.join(self.status_dir, f"{task_id}.log")

    def save(self, task_id: str, data: Dict):
        with self._lock:
            with open(self._path(task_id), 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

    def load(self, task_id: str) -> Optional[Dict]:
        path = self._path(task_id)
        if not os.path.exists(path):
            return None
        with self._lock:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if not content:
                        logger.warning(f"任务文件为空: {path}")
                        return None
                    return json.loads(content)
            except json.JSONDecodeError:
                logger.error(f"任务文件损坏: {path}")
                return None
            except Exception as e:
                logger.error(f"读取任务文件失败: {path} - {e}")
                return None

    def update(self, task_id: str, **fields):
        with self._lock:
            data = self.load(task_id)
            if data is None:
                raise ValueError(f"Task {task_id} not found")
            data.update(fields)
            self.save(task_id, data)

    def delete(self, task_id: str):
        path = self._path(task_id)
        if os.path.exists(path):
            os.remove(path)

    def get_output_dir(self) -> str:
        return self.output_dir

    def append_log(self, task_id: str, message: str):
        log_path = self._log_path(task_id)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")

    def get_log(self, task_id: str, max_lines: int = 2000) -> str:
        log_path = self._log_path(task_id)
        if not os.path.exists(log_path):
            return ""
        with open(log_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        if len(lines) > max_lines:
            lines = lines[-max_lines:]
        return ''.join(lines)
