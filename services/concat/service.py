# services/concat/service.py
"""分组拼接服务。

按分组列把多行数据的拼接列合并成单行，支持多列独立拼接、
自然排序、去重、空值过滤、超长保护。"""
import os
import re
import logging
from typing import List, Optional

import pandas as pd

from core.exceptions import ValidationError
from core.header_detection import detect_header_start

logger = logging.getLogger(__name__)


def _natural_sort_key(s):
    """自然排序键：连续数字段按数值、其余按小写字符串。"""
    result = []
    for part in re.split(r"(\d+)", str(s)):
        if part.isdigit():
            result.append((0, int(part)))
        else:
            result.append((1, part.lower()))
    return result


def _is_empty(v) -> bool:
    """空值判定：None / NaN / trim 后空串 均视为空；字符串 'nan'/'None' 不视为空。"""
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


class ConcatService:
    """分组拼接服务（无状态，注入 config）。"""

    def __init__(self, config):
        self.config = config

    # ------------------------------------------------------------------ #
    # 主流程
    # ------------------------------------------------------------------ #
    def execute(
        self,
        file_path: str,
        original_filename: str,
        group_col: str,
        concat_cols: List[str],
        separator: str = ",",
        sort_ascending: bool = True,
        dup_keep: str = "first",
        header_row: Optional[int] = None,
    ) -> str:
        # 1) 参数校验
        if not group_col or not group_col.strip():
            raise ValidationError("请指定分组列")
        if not concat_cols:
            raise ValidationError("请至少指定一个拼接列")
        if dup_keep not in ("first", "last"):
            raise ValidationError(f"无效的去重保留策略: {dup_keep}")

        # 2) 确定表头行（1-based → pandas 0-based）
        if header_row is not None:
            if header_row < 1:
                raise ValidationError("表头行号必须 >= 1")
            pandas_header = header_row - 1
        else:
            try:
                data_start = detect_header_start(file_path)  # 1-based
            except Exception:
                data_start = 2
            pandas_header = max(data_start - 2, 0)

        # 3) 读取（整表 str，na_filter=False，保留精度）
        try:
            df = pd.read_excel(
                file_path,
                sheet_name=0,
                dtype=str,
                na_filter=False,
                header=pandas_header,
            )
        except Exception as e:  # noqa: BLE001
            raise ValidationError(f"无法读取 Excel 文件: {e}")

        if df.empty:
            raise ValidationError("文件无数据")

        # 4) 解析列标识（先列字母、后列名）
        group_col_name = self._resolve_col(df, group_col)
        concat_col_names = [self._resolve_col(df, c) for c in concat_cols]

        # 5) 分组列空值检查（先全表检查，有则立即报错）
        if df[group_col_name].map(_is_empty).any():
            raise ValidationError(
                f"分组列 [{group_col}] 包含空值，已终止处理"
            )

        # 6) 预处理：trim 分组列 + 拼接列（其他列不动）
        df = df.copy()
        df[group_col_name] = df[group_col_name].astype(str).str.strip()
        for c in concat_col_names:
            df[c] = df[c].astype(str).str.strip()

        # 7) 分组处理
        other_cols = [
            c for c in df.columns
            if c != group_col_name and c not in concat_col_names
        ]
        output_order = list(df.columns)

        result_rows = []
        for group_val, group_df in df.groupby(
            group_col_name, sort=False, dropna=False
        ):
            row = {}
            first_row = group_df.iloc[0]
            for c in other_cols:
                row[c] = first_row[c]
            row[group_col_name] = group_val
            for c in concat_col_names:
                row[c] = self._concat_column(
                    group_df[c], separator, sort_ascending, dup_keep, c
                )
            result_rows.append(row)

        result_df = pd.DataFrame(result_rows, columns=output_order)

        # 8) 写出
        out_dir = os.path.dirname(file_path)
        base = os.path.splitext(os.path.basename(original_filename))[0]
        base = re.sub(r'[\\/*?:"<>|]', "_", base) or "result"
        out_path = os.path.join(out_dir, f"{base}_分组拼接.xlsx")
        result_df.to_excel(out_path, index=False)
        return out_path

    # ------------------------------------------------------------------ #
    # 辅助
    # ------------------------------------------------------------------ #
    @staticmethod
    def _resolve_col(df: pd.DataFrame, col_str: str) -> str:
        """先按列字母、后按列名解析。返回解析后的原列名。"""
        from openpyxl.utils import column_index_from_string

        col_str = col_str.strip()
        if not col_str:
            raise ValidationError("列标识不能为空")

        upper = col_str.upper()
        if 1 <= len(upper) <= 3 and upper.isalpha():
            try:
                idx = column_index_from_string(upper)
                if 1 <= idx <= len(df.columns):
                    return df.columns[idx - 1]
            except Exception:  # noqa: BLE001
                pass

        if col_str in df.columns:
            return col_str

        raise ValidationError(f"列 '{col_str}' 在文件中不存在")

    def _concat_column(
        self,
        series: pd.Series,
        separator: str,
        sort_ascending: bool,
        dup_keep: str,
        col_name: str,
    ) -> str:
        # 1) 过滤空值
        values = [str(v).strip() for v in series if not _is_empty(v)]
        if not values:
            return ""

        # 2) 自然排序
        values = sorted(values, key=_natural_sort_key, reverse=not sort_ascending)

        # 3) 去重
        if dup_keep == "first":
            seen = set()
            kept = []
            for v in values:
                if v not in seen:
                    seen.add(v)
                    kept.append(v)
        else:  # last
            seen = set()
            kept = []
            for v in reversed(values):
                if v not in seen:
                    seen.add(v)
                    kept.append(v)
            kept.reverse()

        # 4) 去重后再按同一规则排序一次（稳定升/降序）
        kept = sorted(kept, key=_natural_sort_key, reverse=not sort_ascending)

        # 5) 拼接
        result = separator.join(kept)

        # 6) 长度检查
        if len(result) > self.config.max_cell_chars:
            raise ValidationError(
                f"拼接结果超过Excel单单元格上限{self.config.max_cell_chars}字符，"
                f"当前长度 {len(result)}（列: {col_name}）"
            )
        return result
