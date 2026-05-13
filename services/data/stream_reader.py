import os, logging, csv
from typing import List, Optional, Callable, Generator, Dict, Any
from utils.condition_filter import ConditionFilter, Condition

logger = logging.getLogger(__name__)


class StreamReader:
    """流式读取 Excel/CSV，返回清洗后的字典行，支持通用进度回调"""

    @staticmethod
    def read_excel(
        filepath: str,
        columns: List[str],
        filters: Optional[List[Dict]] = None,
        progress_callback: Optional[Callable] = None,
        total_rows: int = 0
    ) -> Generator[Dict[str, str], None, None]:
        """
        读取 Excel（兼容旧回调）
        旧回调形式: progress_callback(processed_rows=..., total_rows=...)
        """
        from openpyxl import load_workbook
        wb = load_workbook(filepath, read_only=True)
        ws = wb.active
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        col_indices = {col: headers.index(col) for col in columns if col in headers}
        if not col_indices:
            wb.close()
            raise ValueError(f"文件 {filepath} 缺少必要列: {columns}")
        filter_conditions = [Condition(**f) for f in filters] if filters else []
        processed = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            d = {}
            for col in columns:
                if col in col_indices:
                    val = row[col_indices[col]]
                    d[col] = str(val).strip().lower() if val is not None else ''
                else:
                    d[col] = ''
            if filter_conditions:
                prepared = d.copy()
                for f in filters:
                    if f.get('op') in ('>','<','>=','<='):
                        try: prepared[f['field']] = float(prepared[f['field']])
                        except: pass
                if not ConditionFilter.apply_to_dict(prepared, filter_conditions):
                    processed += 1
                    if progress_callback:
                        progress_callback(processed_rows=processed, total_rows=total_rows)
                    continue
            processed += 1
            if progress_callback:
                progress_callback(processed_rows=processed, total_rows=total_rows)
            yield d
        wb.close()

    @classmethod
    def read_excel_with_progress(
        cls,
        filepath: str,
        columns: List[str],
        filters: Optional[List[Dict]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None,
        total_rows: int = 0
    ) -> Generator[Dict[str, str], None, None]:
        """
        读取 Excel，支持 ProgressTask 风格的进度回调
        progress_callback(percent, message)
        """
        processed = 0
        total = total_rows if total_rows else 1

        def old_callback(processed_rows, total_rows):
            nonlocal processed
            processed = processed_rows
            if progress_callback:
                pct = min(int(processed / total * 100), 99) if total else 50
                msg = f"已读取 {processed}/{total}"
                progress_callback(pct, msg)

        yield from cls.read_excel(filepath, columns, filters, old_callback, total)
        if progress_callback:
            progress_callback(100, "读取完成")

