# services/transposer/service.py
import pandas as pd
from typing import List, Optional


def pivot_table(df: pd.DataFrame,
                index_cols: Optional[List[str]],
                category_col: str,
                value_cols: List[str]) -> pd.DataFrame:
    """
    将分类列的值转置为新的列名，同时保留其他固定列。

    :param df: 原始数据
    :param index_cols: 索引列列表；若为空则自动生成组内序号（此时固定列不会被保留）
    :param category_col: 分类列名（如“医疗类别”）
    :param value_cols: 值列名列表（如“a数据”，“b数据”）
    :return: 转置后的 DataFrame，固定列在左侧，拼接后的新值列严格按“类别+值列名”排列（无空格）
    """
    df = df.copy()

    # 如果未提供索引列，自动生成组内序号，并在此情况下不保留固定列
    auto_index = False
    if not index_cols:
        df['_idx'] = df.groupby(category_col).cumcount()
        index_cols = ['_idx']
        auto_index = True

    # 区分固定列：所有列 - category_col - value_cols
    # 注意：value_cols 中可能包含已被选为索引的列，但用户应避免错误输入
    other_cols = [c for c in df.columns 
                  if c not in value_cols and c != category_col]

    # 如果使用自动序号，则不保留任何固定列（因为固定列与自动序号没有实际关联）
    if auto_index:
        other_cols = []
    else:
        # 确保索引列被包含在 other_cols 中（通常已包含）
        # 并去除重复
        other_cols = list(dict.fromkeys(other_cols))

    # 取出固定列的唯一值（按 index_cols 去重）
    if other_cols:
        fixed_df = df[other_cols].drop_duplicates(subset=index_cols)
        # 将固定列设置为索引以便后续合并
        fixed_df = fixed_df.set_index(index_cols)
    else:
        fixed_df = pd.DataFrame()

    # 执行 pivot：包括索引列、分类列、值列
    pivot_input = df[index_cols + [category_col] + value_cols]
    pivoted = pivot_input.pivot(index=index_cols, columns=category_col, values=value_cols)

    # 扁平化多层列名并去除空格，连接方式为“类别+值列名”，不包含任何分隔符
    new_columns = []
    for col in pivoted.columns:
        if isinstance(col, tuple):
            # 元组顺序可能是 (value, category) 或 (category, value)，需要统一
            # 通过检查 col 中哪个是 value_cols 中的列来确定顺序
            cat, val = None, None
            for part in col:
                if part in value_cols:
                    val = str(part).strip()
                else:
                    cat = str(part).strip()
            if cat and val:
                new_col = f"{cat}{val}"
                new_columns.append(new_col)
            else:
                # 退化情况
                new_columns.append(''.join(str(x).strip() for x in col))
        else:
            new_columns.append(str(col).strip())
    pivoted.columns = new_columns

    # 重置索引以便与固定列合并
    pivoted.reset_index(inplace=True)
    # 删除可能由自动索引添加的 '_idx' 列
    if '_idx' in pivoted.columns:
        pivoted.drop(columns=['_idx'], inplace=True)

    # 合并固定列
    if not fixed_df.empty:
        # 将 pivoted 的索引列也设为索引以执行合并
        pivoted_merge = pivoted.set_index(index_cols)
        result = fixed_df.join(pivoted_merge, how='right')
        result.reset_index(inplace=True)
    else:
        result = pivoted

    # 重新排序列：固定列（other_cols 中出现在 result 中的列）→ 新的值列
    fixed_cols_in_result = [c for c in other_cols if c in result.columns]
    new_value_cols = [c for c in result.columns if c not in fixed_cols_in_result]
    result = result[fixed_cols_in_result + new_value_cols]

    # 对新的值列按“类别+值列名”排序（例如：门诊a数据、门诊b数据、住院a数据、住院b数据、其他a数据、其他b数据）
    # 此处保持原顺序，仅确保无空白
    return result