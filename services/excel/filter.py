# services/excel/filter.py
import pandas as pd
import logging
from typing import List, Tuple, Optional
from core.filter_engine import ConditionEngine

logger = logging.getLogger(__name__)

class ExcelFilter:
    """Excel 数据筛选核心类（委托给 core.filter_engine）"""

    @staticmethod
    def apply_filters(df: pd.DataFrame, conditions: List[Tuple[str, List[str]]],
                      fuzzy_field: Optional[str] = None, fuzzy_value: Optional[str] = None) -> pd.DataFrame:
        engine = ConditionEngine()
        exact = []
        fuzzy = []
        for field, values in conditions:
            if values:
                exact.append({'field': field, 'op': '=', 'values': values})
        if fuzzy_field and fuzzy_value:
            fuzzy.append({'field': fuzzy_field, 'op': '包含', 'value': fuzzy_value})
        return engine.apply_to_dataframe(df, exact, fuzzy)

    @staticmethod
    def read_and_combine(file_paths: List[str]) -> pd.DataFrame:
        from core.excel_utils import ExcelFile
        dfs = []
        for path in file_paths:
            df = ExcelFile.read_dataframe(path)
            dfs.append(df)
        if not dfs:
            raise ValueError("没有成功读取任何文件")
        return pd.concat(dfs, ignore_index=True)

    @staticmethod
    def round_numeric_columns(df: pd.DataFrame, decimals: int = 2) -> pd.DataFrame:
        df = df.copy()
        for col in df.columns:
            numeric_series = pd.to_numeric(df[col], errors='coerce')
            if numeric_series.notna().any():
                df[col] = numeric_series.apply(lambda x: round(x, decimals) if pd.notnull(x) else x)
        return df
