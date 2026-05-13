# services/checker/precheck.py
import os
import pandas as pd
from typing import List
from core.exceptions import ValidationError

class PrecheckEngine:
    """预检：快速扫描文件，返回统计信息"""

    @staticmethod
    def estimate(files: List[str], sample_rows: int = 1000) -> dict:
        total_rows = 0
        total_size_mb = 0
        for f in files:
            if not os.path.exists(f):
                continue
            total_size_mb += os.path.getsize(f) / (1024 * 1024)
            try:
                # 采样读取
                sample = pd.read_excel(f, nrows=sample_rows)
                # 粗略推算总行数（假设文件均匀）
                # 真实实现需更精细
            except Exception:
                pass
        return {
            "file_count": len(files),
            "total_size_mb": round(total_size_mb, 2),
            "estimated_rows": total_rows,
            "warnings": [],
        }
