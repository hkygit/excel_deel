# services/excel/aggregator.py
"""通用聚合器（纯静态方法）。"""
import pandas as pd
from typing import List, Dict, Union, Optional


class Aggregator:
    """通用聚合器（静态方法）。"""

    @staticmethod
    def group_aggregate(df: pd.DataFrame, group_keys: List[str],
                        agg_spec: Dict[str, Union[str, List[str]]]) -> pd.DataFrame:
        if not group_keys:
            return df.drop_duplicates()
        return df.groupby(group_keys, as_index=False).agg(agg_spec)

    @staticmethod
    def aggregate_by_keys(df: pd.DataFrame, unique_keys: List[str],
                          sum_fields: List[str],
                          other_fields: Optional[List[str]] = None) -> pd.DataFrame:
        """按唯一键聚合：sum_fields 求和，other_fields 取第一条记录。

        实现上组合 drop_duplicates 与 groupby sum，避免 pandas 在某些
        版本上对 agg(dict) 的内部错误。
        """
        if other_fields is None:
            other_fields = [
                c for c in df.columns
                if c not in unique_keys and c not in sum_fields
            ]
        if not unique_keys:
            return df.drop_duplicates()

        # 1. 求和部分
        if sum_fields:
            sum_part = df.groupby(unique_keys, as_index=False)[sum_fields].sum()
        else:
            sum_part = pd.DataFrame(columns=unique_keys)

        # 2. 取第一条记录部分
        if other_fields:
            first_part = df[unique_keys + other_fields].drop_duplicates(
                subset=unique_keys, keep='first'
            )
        else:
            first_part = pd.DataFrame(columns=unique_keys)

        # 3. 合并
        if not sum_part.empty and not first_part.empty:
            result = sum_part.merge(first_part, on=unique_keys, how='outer')
        elif not sum_part.empty:
            result = sum_part
        elif not first_part.empty:
            result = first_part
        else:
            result = df[unique_keys].drop_duplicates()

        return result

    @staticmethod
    def simple_aggregate(df: pd.DataFrame, group_keys: List[str],
                         agg_cols: List[str], func: str = 'sum') -> pd.DataFrame:
        agg_dict = {col: func for col in agg_cols}
        return Aggregator.group_aggregate(df, group_keys, agg_dict)
