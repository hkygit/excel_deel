# utils/common/aggregator.py
import pandas as pd
from typing import List, Dict, Any, Union, Optional

class Aggregator:
    """通用聚合器（静态方法）"""

    @staticmethod
    def group_aggregate(df: pd.DataFrame, group_keys: List[str],
                        agg_spec: Dict[str, Union[str, List[str]]]) -> pd.DataFrame:
        if not group_keys:
            return df.drop_duplicates()
        return df.groupby(group_keys, as_index=False).agg(agg_spec)

    @staticmethod
    def aggregate_by_keys(df: pd.DataFrame, unique_keys: List[str],
                          sum_fields: List[str], other_fields: Optional[List[str]] = None) -> pd.DataFrame:
        """
        按唯一键聚合：sum_fields 求和，other_fields 取第一条记录。
        不使用 groupby.agg 字典，改为组合 drop_duplicates 和 groupby sum，避免 pandas 内部错误。
        """
        if other_fields is None:
            other_fields = [c for c in df.columns if c not in unique_keys and c not in sum_fields]
        if not unique_keys:
            return df.drop_duplicates()

        # 1. 求和部分
        if sum_fields:
            sum_part = df.groupby(unique_keys, as_index=False)[sum_fields].sum()
        else:
            sum_part = pd.DataFrame(columns=unique_keys)  # 空框架，仅用于后续合并

        # 2. 取第一条记录部分（直接用 drop_duplicates 保留第一个出现）
        if other_fields:
            first_part = df[unique_keys + other_fields].drop_duplicates(subset=unique_keys, keep='first')
        else:
            first_part = pd.DataFrame(columns=unique_keys)

        # 3. 合并求和与首条记录
        if not sum_part.empty and not first_part.empty:
            result = sum_part.merge(first_part, on=unique_keys, how='outer')
        elif not sum_part.empty:
            result = sum_part
        elif not first_part.empty:
            result = first_part
        else:
            # 没有求和也没有其他字段时，直接去重返回唯一键
            result = df[unique_keys].drop_duplicates()

        return result

    @staticmethod
    def simple_aggregate(df: pd.DataFrame, group_keys: List[str],
                         agg_cols: List[str], func: str = 'sum') -> pd.DataFrame:
        agg_dict = {col: func for col in agg_cols}
        return Aggregator.group_aggregate(df, group_keys, agg_dict)


# 向后兼容包装
class AggregateCalculator:
    def __init__(self, excel_file):
        self.excel_file = excel_file

    def aggregate(self, group_cols, agg_cols, agg_func, result_col):
        raise NotImplementedError
