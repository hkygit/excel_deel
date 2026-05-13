# utils/task/logger.py
import json
import logging
from datetime import datetime
from typing import Any

_logger = logging.getLogger('app')

def log_event(task_id: str, event: str, **kwargs: Any):
    log_entry = {
        'timestamp': datetime.now().isoformat(),
        'task_id': task_id,
        'event': event,
        **kwargs
    }
    _logger.info(json.dumps(log_entry, ensure_ascii=False))
