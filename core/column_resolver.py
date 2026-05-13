# core/column_resolver.py
"""公共列解析：将用户输入的列名/列字母转换为DataFrame子集"""
from typing import List, Dict
import pandas as pd
from core.csv_utils import split_fields
from core.exceptions import ValidationError

def resolve_columns(df: pd.DataFrame, col_str: str) -> pd.DataFrame:
    if not col_str.strip():
        return df
    candidates = split_fields(col_str)
    from core.csv_utils import build_column_mapping
    col_map = build_column_mapping(df.columns.tolist())[0]
    final = []
    for cand in candidates:
        if not cand:
            continue
        if cand.isalpha() and len(cand) <= 3:   # 列字母
            upper = cand.upper()
            if upper in col_map:
                final.append(col_map[upper])
                continue
        if cand in df.columns:                   # 列名
            final.append(cand)
    if not final:
        raise ValidationError("指定的列均不存在或无效")
    return df[final]
