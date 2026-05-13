# services/data/report_writer.py
import tempfile, os, logging
from typing import Dict, Optional, Callable
import pandas as pd

logger = logging.getLogger(__name__)

class ReportWriter:
    """通用 Excel 报告写入器，支持超时跳过单个 Sheet"""
    def __init__(self, timeout_seconds: Optional[int] = None, progress_callback: Optional[Callable] = None):
        self.timeout_seconds = timeout_seconds
        self.progress_callback = progress_callback

    def write(self, sheets: Dict[str, pd.DataFrame], start_time: float) -> str:
        """sheets: {sheet_name: DataFrame}"""
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            for name, df in sheets.items():
                if df.empty:
                    continue
                # 特殊处理标准群独有超时
                if name == "标准群独有" and self.timeout_seconds:
                    import time
                    if time.time() - start_time > self.timeout_seconds:
                        logger.warning(f"跳过 {name}，已超时")
                        continue
                df.to_excel(writer, sheet_name=name, index=False)
        return tmp.name
