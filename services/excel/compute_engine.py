# services/excel/compute_engine.py
import pandas as pd
from typing import List, Dict, Any, Optional, Union
from core.expression_engine import get_expression_engine
from services.excel.aggregator import Aggregator

class ComputeEngine:
    @staticmethod
    def evaluate_expression(expr: str, values: Dict[str, float]) -> float:
        engine = get_expression_engine()
        return engine.evaluate(expr, values)

    @staticmethod
    def extract_variables(expr: str) -> set:
        import re
        return set(re.findall(r'[A-Z]+', expr.upper()))

    @staticmethod
    def aggregate(df: pd.DataFrame, group_keys: List[str],
                  agg_spec: Dict[str, Union[str, List[str]]]) -> pd.DataFrame:
        return Aggregator.group_aggregate(df, group_keys, agg_spec)

    @staticmethod
    def simple_aggregate(df: pd.DataFrame, group_keys: List[str],
                         agg_cols: List[str], func: str = 'sum') -> pd.DataFrame:
        return Aggregator.simple_aggregate(df, group_keys, agg_cols, func)

    @staticmethod
    def aggregate_by_keys(df: pd.DataFrame, unique_keys: List[str],
                          sum_fields: List[str], other_fields: Optional[List[str]] = None) -> pd.DataFrame:
        return Aggregator.aggregate_by_keys(df, unique_keys, sum_fields, other_fields)
