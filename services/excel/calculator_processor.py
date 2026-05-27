# services/excel/calculator_processor.py
"""Excel 列计算与聚合回填处理器"""
import os
import gc
import logging
import numpy as np
import openpyxl
import pandas as pd
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional, Tuple

from core.excel_utils import ExcelFile, copy_merge_structure
from core.data_access_layer import get_data_access_layer
from core.csv_utils import build_column_mapping
from core.excel_helpers import (
    safe_get_cell_value, safe_write_cell, clear_merges_in_column,
    build_column_name_mapping, normalize_excel_file, safe_cleanup_dataframe
)
from services.excel.compute_engine import ComputeEngine

logger = logging.getLogger(__name__)


class BaseExcelProcessor(ABC):
    """Excel 处理器基类，封装公共的写入逻辑"""

    def __init__(self, filepath: str, header_row: int, compute_engine: ComputeEngine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_file = None

    def _get_work_filepath(self) -> str:
        """获取可写的工作文件路径，统一通过 normalize_excel_file 转换为临时 .xlsx"""
        # 总是进行规范化，确保文件可读
        self.temp_file = normalize_excel_file(self.original_filepath)
        return self.temp_file

    def _cleanup_temp(self):
        if self.temp_file and os.path.exists(self.temp_file):
            try:
                os.unlink(self.temp_file)
            except Exception as e:
                logger.warning(f"清理临时文件失败: {e}")

    def _resolve_result_column(self, ws, result_col_input: str) -> Tuple[int, str]:
        """
        解析结果列：支持列字母或列名。
        返回 (列索引, 列名)
        """
        # 1. 尝试作为列字母解析
        try:
            col_idx = openpyxl.utils.column_index_from_string(result_col_input.upper())
            header_cell = ws.cell(self.header_row, col_idx)
            col_name = str(header_cell.value).strip() if header_cell.value else result_col_input
            return col_idx, col_name
        except Exception:
            pass

        # 2. 尝试作为列名查找
        name_to_idx = build_column_name_mapping(ws, self.header_row)
        if result_col_input in name_to_idx:
            col_idx = name_to_idx[result_col_input]
            return col_idx, result_col_input

        # 3. 新列：追加到末尾
        new_col_idx = ws.max_column + 1
        ws.insert_cols(new_col_idx)
        safe_write_cell(ws, self.header_row, new_col_idx, result_col_input)
        return new_col_idx, result_col_input

    def _write_result_column(self, ws, result_col_idx: int, values: List[Any],
                             start_row: int, end_row: int) -> None:
        """将计算结果写入指定列，并清理原有合并单元格"""
        clear_merges_in_column(ws, result_col_idx)
        for i, val in enumerate(values):
            row = start_row + i
            if isinstance(val, (np.integer, np.floating)):
                val = val.item()
            safe_write_cell(ws, row, result_col_idx, val)

    def _copy_merge_structure(self, ws, src_col_letter: str, dst_col_idx: int,
                              start_row: int, end_row: int) -> None:
        """复制源列的合并结构到目标列"""
        try:
            src_col_idx = openpyxl.utils.column_index_from_string(src_col_letter)
        except Exception:
            return
        copy_merge_structure(ws, src_col_idx, ws, dst_col_idx, start_row, end_row)

    def _get_output_path(self) -> str:
        """生成输出文件路径"""
        base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
        output_dir = os.path.dirname(self.original_filepath)
        return os.path.join(output_dir, f"{base_name}_处理版.xlsx")

    @abstractmethod
    def process(self) -> str:
        pass


class ColumnCalculatorProcessor(BaseExcelProcessor):
    """列计算处理器"""

    def __init__(self, filepath: str, header_row: int, expression: str,
                 result_col_input: str, merge_by_cols_str: str, compute_engine: ComputeEngine):
        super().__init__(filepath, header_row, compute_engine)
        self.expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols = [c.strip() for c in merge_by_cols_str.split(',') if c.strip()] if merge_by_cols_str else []

    def _remove_result_column_from_expr(self, result_col_letter: str) -> str:
        """从表达式中移除结果列字母（如果存在）"""
        if not result_col_letter:
            return self.expression
        tokens = []
        current = []
        for ch in self.expression.upper():
            if ch.isalpha():
                current.append(ch)
            else:
                if current:
                    tokens.append(''.join(current))
                    current = []
                tokens.append(ch)
        if current:
            tokens.append(''.join(current))

        filtered = []
        for token in tokens:
            if token == result_col_letter:
                continue
            filtered.append(token)
        new_expr = ''.join(filtered)
        new_expr = new_expr.replace('+-', '-').replace('-+', '-').replace('++', '+').replace('--', '+')
        new_expr = new_expr.lstrip('+-*/').rstrip('+-*/')
        if new_expr != self.expression:
            logger.warning(f"表达式中包含了结果列 {result_col_letter}，已自动移除: {self.expression} -> {new_expr}")
        return new_expr

    def process(self) -> str:
        filepath = self._get_work_filepath()
        try:
            # 强制以字符串类型读取，避免公式错误
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = safe_cleanup_dataframe(df, auto_convert_numeric=True)
            headers = list(df.columns)

            letter_to_name, name_to_letter = build_column_mapping(headers)

            with ExcelFile(filepath) as ef:
                ws = ef.active_sheet
                result_col_idx, _ = self._resolve_result_column(ws, self.result_col_input)

                result_col_letter = None
                try:
                    result_col_letter = openpyxl.utils.get_column_letter(result_col_idx)
                except Exception:
                    pass

                expr = self.expression
                if result_col_letter:
                    expr = self._remove_result_column_from_expr(result_col_letter)
                    if not expr.strip():
                        raise ValueError(f"表达式在移除结果列 {result_col_letter} 后为空")

                variables = self.compute_engine.extract_variables(expr)
                mapped_vars = {}
                missing = []
                for letter in variables:
                    if letter in letter_to_name:
                        mapped_vars[letter] = letter_to_name[letter]
                    else:
                        missing.append(letter)
                if missing:
                    raise ValueError(f"表达式中的列字母 {missing} 无法映射到文件中的任何列")

                def calc_row(row):
                    context = {letter: row[col_name] for letter, col_name in mapped_vars.items()}
                    return self.compute_engine.evaluate_expression(expr, context)
                df["Result"] = df.apply(calc_row, axis=1)

                start_row = self.header_row + 1
                end_row = ws.max_row
                values = [None if pd.isna(v) else v for v in df["Result"].tolist()]

                self._write_result_column(ws, result_col_idx, values, start_row, end_row)

                if variables:
                    ref_col_letter = sorted(variables)[0]
                    self._copy_merge_structure(ws, ref_col_letter, result_col_idx, start_row, end_row)

                if self.merge_by_cols:
                    merge_indices = []
                    for col in self.merge_by_cols:
                        try:
                            merge_indices.append(openpyxl.utils.column_index_from_string(col))
                        except Exception:
                            if col in name_to_letter:
                                merge_indices.append(openpyxl.utils.column_index_from_string(name_to_letter[col]))
                            else:
                                raise ValueError(f"合并依据列 {col} 无效")
                    group_first_row = {}
                    for row in range(start_row, end_row + 1):
                        key = tuple(safe_get_cell_value(ws, row, col) for col in merge_indices)
                        if key not in group_first_row:
                            group_first_row[key] = row
                    for row in range(start_row, end_row + 1):
                        key = tuple(safe_get_cell_value(ws, row, col) for col in merge_indices)
                        if row != group_first_row.get(key):
                            safe_write_cell(ws, row, result_col_idx, None)

                result_path = self._get_output_path()
                ef.save(result_path)
                ef._workbook.close()

            gc.collect()
            return result_path
        finally:
            self._cleanup_temp()


class AggregateBackfillProcessor(BaseExcelProcessor):
    """聚合回填处理器"""

    def __init__(self, filepath: str, header_row: int,
                 group_cols_raw: List[str], agg_cols_raw: List[str],
                 agg_func: str, result_col_input: str, compute_engine: ComputeEngine):
        super().__init__(filepath, header_row, compute_engine)
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input

    def process(self) -> str:
        filepath = self._get_work_filepath()
        try:
            # 读取表头（仅列名），先不读取数据
            df_header = pd.read_excel(filepath, header=self.header_row - 1, nrows=0)
            headers = list(df_header.columns)
            letter_to_name, name_to_letter = build_column_mapping(headers)

            def resolve_col(col: str) -> str:
                col = col.strip()
                if col in headers:
                    return col
                upper = col.upper()
                if upper in letter_to_name:
                    return letter_to_name[upper]
                raise ValueError(f"列 '{col}' 无法映射到实际列名")

            group_cols = [resolve_col(c) for c in self.group_cols_raw]
            agg_cols = [resolve_col(c) for c in self.agg_cols_raw]

            # 以字符串类型读取数据
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = safe_cleanup_dataframe(df, auto_convert_numeric=True)

            missing = [c for c in group_cols + agg_cols if c not in df.columns]
            if missing:
                raise ValueError(f"列不存在: {missing}")

            result_df = self.compute_engine.simple_aggregate(df, group_cols, agg_cols, self.agg_func)

            with ExcelFile(filepath) as ef:
                ws = ef.active_sheet
                name_to_idx = build_column_name_mapping(ws, self.header_row)

                def get_col_index(col: str) -> int:
                    col = col.strip()
                    if col in name_to_idx:
                        return name_to_idx[col]
                    try:
                        return openpyxl.utils.column_index_from_string(col)
                    except Exception:
                        raise ValueError(f"列 '{col}' 在原始表中不存在")

                group_col_indices = [get_col_index(c) for c in group_cols]

                result_col_idx, _ = self._resolve_result_column(ws, self.result_col_input)

                result_map = {}
                for _, row in result_df.iterrows():
                    key = tuple(str(row[c]) for c in group_cols)
                    val = row[agg_cols[0]]
                    if hasattr(val, 'item'):
                        val = val.item()
                    result_map[key] = val

                start_row = self.header_row + 1
                end_row = ws.max_row
                for row in range(start_row, end_row + 1):
                    key = tuple(str(safe_get_cell_value(ws, row, idx)) for idx in group_col_indices)
                    val = result_map.get(key, "")
                    safe_write_cell(ws, row, result_col_idx, val)

                result_path = self._get_output_path()
                ef.save(result_path)
                ef._workbook.close()

            gc.collect()
            return result_path
        finally:
            self._cleanup_temp()
