#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最终修复：使用 openpyxl 的 data_only=True 模式重新保存文件，避免公式求值错误。
"""

import os
import tempfile
import logging
from openpyxl import load_workbook

logger = logging.getLogger(__name__)

# 只修改 core/excel_helpers.py 中的 normalize_excel_file 函数
NEW_EXCEL_HELPERS = '''# core/excel_helpers.py
"""Excel 操作公共工具函数"""
import re
import logging
from typing import List, Tuple, Optional, Dict, Any
import openpyxl
import os
import tempfile
import pandas as pd
import numpy as np

logger = logging.getLogger(__name__)

def safe_get_cell_value(ws, row: int, col: int) -> Any:
    """安全获取单元格值，支持合并单元格"""
    cell = ws.cell(row, col)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
                return ws.cell(merged_range.min_row, merged_range.min_col).value
    return cell.value


def safe_write_cell(ws, row: int, col: int, value: Any) -> None:
    """安全写入单元格值，若单元格属于合并区域则先拆分"""
    cell = ws.cell(row, col)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        to_unmerge = []
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
                to_unmerge.append(str(merged_range))
        for rng in to_unmerge:
            ws.unmerge_cells(rng)
    ws.cell(row, col).value = value


def clear_merges_in_column(ws, col_idx: int) -> None:
    """清除某一列上的所有合并单元格"""
    to_remove = []
    for merged in ws.merged_cells.ranges:
        if merged.min_col <= col_idx <= merged.max_col:
            to_remove.append(merged)
    for merged in to_remove:
        ws.unmerge_cells(str(merged))


def build_column_name_mapping(ws, header_row: int) -> Dict[str, int]:
    """
    构建列名到列索引的映射，正确处理合并单元格。
    :return: {列名: 列索引}
    """
    name_to_idx = {}
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(header_row, col_idx)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            for merged in ws.merged_cells.ranges:
                if (merged.min_row <= header_row <= merged.max_row and
                    merged.min_col <= col_idx <= merged.max_col):
                    cell = ws.cell(merged.min_row, merged.min_col)
                    break
        if cell.value is not None:
            name = str(cell.value).strip()
            if name:
                name_to_idx[name] = col_idx
    return name_to_idx


def normalize_excel_file(filepath: str) -> str:
    """
    将任何 Excel 文件转换为标准 .xlsx 临时文件，使用 openpyxl 的 data_only 模式，
    确保公式被计算为值，避免后续读取时出错。
    返回新的临时文件路径，调用者负责清理。
    """
    # 优先使用 openpyxl 读取（data_only=True）
    try:
        # 使用 data_only=True 加载工作簿，读取公式计算后的值
        wb = load_workbook(filepath, data_only=True, keep_vba=False)
        fd, temp_xlsx = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        wb.save(temp_xlsx)
        wb.close()
        logger.info(f"使用 openpyxl(data_only) 成功转换为标准格式: {temp_xlsx}")
        return temp_xlsx
    except Exception as e:
        logger.warning(f"openpyxl 转换失败: {e}，回退到 pandas 方式")
        # 回退：使用 pandas 以字符串类型读取（可能会遇到公式问题，但作为最后手段）
        try:
            sheets = pd.read_excel(filepath, sheet_name=None, dtype=str, engine='openpyxl')
            if not sheets:
                raise ValueError("文件无工作表")
            fd, temp_xlsx = tempfile.mkstemp(suffix='.xlsx')
            os.close(fd)
            with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
                for sheet_name, df in sheets.items():
                    if df.empty:
                        pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            logger.info(f"使用 pandas 成功转换为标准格式: {temp_xlsx}")
            return temp_xlsx
        except Exception as e2:
            raise ValueError(f"无法读取或转换 Excel 文件: {str(e2)}")


def safe_cleanup_dataframe(df: pd.DataFrame, auto_convert_numeric: bool = True) -> pd.DataFrame:
    """
    安全清洗 DataFrame：
    - 去除字符串首尾空格，将空字符串转为 NaN。
    - 可选：如果某列所有非空值都能转换为数值，则转换为数值。
    """
    df = df.copy()
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].fillna('').astype(str).str.strip()
        df[col] = df[col].replace(r'^\s*$', None, regex=True)

        if auto_convert_numeric:
            non_null = df[col].dropna()
            if len(non_null) > 0:
                try:
                    pd.to_numeric(non_null, errors='raise')
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    logger.debug(f"列 '{col}' 已自动转换为数值类型")
                except (ValueError, TypeError):
                    pass
    return df
'''

def main():
    # 目标文件路径
    target = "core/excel_helpers.py"
    with open(target, 'w', encoding='utf-8') as f:
        f.write(NEW_EXCEL_HELPERS.lstrip())
    print(f"已更新: {target}")
    print("修复完成。现在优先使用 openpyxl data_only 模式转换文件，避免公式求值错误。")

if __name__ == "__main__":
    main()