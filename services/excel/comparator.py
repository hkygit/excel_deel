# services/excel/comparator.py
import pandas as pd
import logging
from typing import Dict, Any, Tuple, List, Optional

logger = logging.getLogger(__name__)

class ExcelComparator:
    """Excel 文件对比核心类（业务服务）"""

    @staticmethod
    def _read_sheet(filepath: str, sheet: Optional[str]) -> pd.DataFrame:
        if sheet is None or sheet == '':
            return pd.read_excel(filepath, sheet_name=0)
        try:
            idx = int(sheet)
            return pd.read_excel(filepath, sheet_name=idx)
        except ValueError:
            return pd.read_excel(filepath, sheet_name=sheet)

    @classmethod
    def full_compare(cls, file1: str, file2: str,
                     key_fields: List[str], compare_fields: List[str],
                     sheet1: Optional[str] = None, sheet2: Optional[str] = None) -> Dict[str, Any]:
        df1 = cls._read_sheet(file1, sheet1)
        df2 = cls._read_sheet(file2, sheet2)

        cols1, cols2 = set(df1.columns), set(df2.columns)
        missing_keys1 = [k for k in key_fields if k not in cols1]
        missing_keys2 = [k for k in key_fields if k not in cols2]
        if missing_keys1 or missing_keys2:
            raise ValueError(f"主键字段不存在: {missing_keys1 + missing_keys2}")

        missing_comp1 = [c for c in compare_fields if c not in cols1]
        missing_comp2 = [c for c in compare_fields if c not in cols2]
        if missing_comp1 or missing_comp2:
            raise ValueError(f"对比字段不存在: {missing_comp1 + missing_comp2}")

        if df1.duplicated(subset=key_fields).any():
            raise ValueError("文件1中主键字段存在重复行")
        if df2.duplicated(subset=key_fields).any():
            raise ValueError("文件2中主键字段存在重复行")

        df1.set_index(key_fields, inplace=True)
        df2.set_index(key_fields, inplace=True)

        df1_comp = df1[compare_fields] if set(compare_fields).issubset(df1.columns) else df1.reindex(columns=compare_fields)
        df2_comp = df2[compare_fields] if set(compare_fields).issubset(df2.columns) else df2.reindex(columns=compare_fields)

        only_in_1 = df1_comp.loc[~df1_comp.index.isin(df2_comp.index)]
        only_in_2 = df2_comp.loc[~df2_comp.index.isin(df1_comp.index)]

        common_keys = df1_comp.index.intersection(df2_comp.index)
        diff_rows = []
        for key in common_keys:
            row1, row2 = df1_comp.loc[key], df2_comp.loc[key]
            if not row1.equals(row2):
                diff = {'主键': key}
                for f in compare_fields:
                    diff[f'文件1_{f}'] = row1[f]
                    diff[f'文件2_{f}'] = row2[f]
                    diff[f'{f}_是否一致'] = row1[f] == row2[f]
                diff_rows.append(diff)

        df_diff = pd.DataFrame(diff_rows) if diff_rows else None

        if not only_in_1.empty:
            only_in_1.reset_index(inplace=True)
        if not only_in_2.empty:
            only_in_2.reset_index(inplace=True)

        return {'only_in_1': only_in_1, 'only_in_2': only_in_2, 'diff': df_diff}

    @classmethod
    def field_match(cls, file1: str, file2: str,
                    match_field: str,
                    sheet1: Optional[str] = None, sheet2: Optional[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
        df1 = cls._read_sheet(file1, sheet1)
        df2 = cls._read_sheet(file2, sheet2)
        if match_field not in df1.columns:
            raise ValueError(f"字段 '{match_field}' 在文件1中不存在")
        if match_field not in df2.columns:
            raise ValueError(f"字段 '{match_field}' 在文件2中不存在")
        set1 = set(df1[match_field].astype(str).unique())
        set2 = set(df2[match_field].astype(str).unique())
        only_in_1 = set1 - set2
        only_in_2 = set2 - set1
        filtered_df1 = df1[df1[match_field].astype(str).isin(only_in_1)] if only_in_1 else pd.DataFrame()
        filtered_df2 = df2[df2[match_field].astype(str).isin(only_in_2)] if only_in_2 else pd.DataFrame()
        return filtered_df1, filtered_df2
