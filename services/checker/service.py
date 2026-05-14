# services/checker/service.py
import os
import logging
from typing import List, Dict, Optional, Callable
import pandas as pd

from core.progress_task import ProgressTask
from core.exceptions import ValidationError
from services.data.row_counter import RowCounter
from .strategies import MemoryCheckerStrategy, SqliteCheckerStrategy

logger = logging.getLogger(__name__)


class CheckerService(ProgressTask):
    """文件校验服务（策略上下文）"""
    def __init__(self, files_a=None, files_b=None,
                 key_map_a=None, key_map_b=None,
                 compare_map=None, filters_a=None,
                 duplicate_action=None, config=None,
                 progress_callback=None,
                 precheck_data=None):
        super().__init__(progress_callback)
        self.files_a = files_a or []
        self.files_b = files_b or []
        self.key_map_a = key_map_a or []
        self.key_map_b = key_map_b or []
        self.compare_map = compare_map or []
        self.filters_a = filters_a or []
        self.config = config
        self.duplicate_action = duplicate_action or (config.duplicate_action if config else "error")
        self.precheck_data = precheck_data

    def precheck(self) -> dict:
        ta = sum(RowCounter.count(f) for f in self.files_a)
        tb = sum(RowCounter.count(f) for f in self.files_b)
        total = ta + tb
        total_size = sum(os.path.getsize(f) for f in self.files_a + self.files_b) / (1024 * 1024)
        return {
            "total_files_a": len(self.files_a),
            "total_files_b": len(self.files_b),
            "total_rows_a": ta, "total_rows_b": tb,
            "total_rows": total,
            "estimated_memory_mb": round(total * 200 / (1024*1024), 2),
            "estimated_time_seconds": round(total / 20000, 2),
            "total_size_mb": round(total_size, 2),
            "advice": "数据量大" if total > 500000 else "安全",
            "can_proceed": True
        }

    def execute(self) -> str:
        total_rows_a = self.precheck_data.get('total_rows_a', 0) if self.precheck_data else sum(RowCounter.count(f) for f in self.files_a)
        max_memory_rows = self.config.max_memory_rows if self.config else 200000
        use_sqlite = total_rows_a > max_memory_rows

        strategy_class = SqliteCheckerStrategy if use_sqlite else MemoryCheckerStrategy
        strategy = strategy_class(
            files_a=self.files_a,
            files_b=self.files_b,
            key_map_a=self.key_map_a,
            key_map_b=self.key_map_b,
            compare_map=self.compare_map,
            filters_a=self.filters_a,
            duplicate_action=self.duplicate_action,
            config=self.config,
            progress_callback=self.progress_callback,
            precheck_data=self.precheck_data
        )
        strategy._total_rows = self._total_rows
        strategy._processed_rows = self._processed_rows
        return strategy.execute()
