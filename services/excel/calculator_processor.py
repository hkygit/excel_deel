#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 列计算与聚合回填处理器 (拆分自 blueprints/calculator.py)
"""

import os
import gc
import openpyxl
import pandas as pd
from core.excel_utils import ExcelFile, copy_merge_structure
from core.data_access_layer import get_data_access_layer
from core.csv_utils import build_column_mapping
from core.excel_utils import ExcelFile
from services.excel.compute_engine import ComputeEngine


class ColumnCalculatorProcessor:
    """列计算处理器"""
    def __init__(self, filepath, header_row, expression, result_col_input, merge_by_cols_str, compute_engine):
        self.filepath = filepath
        self.header_row = header_row
        self.expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols_str = merge_by_cols_str
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()

    def process(self):
        df = pd.read_excel(self.filepath, header=self.header_row - 1, dtype=str)
        df = self.dal._auto_clean_dataframe(df)

        variables = self.compute_engine.extract_variables(self.expression)
        missing_cols = [v for v in variables if v not in df.columns]
        if missing_cols:
            raise ValueError(f"列不存在: {missing_cols}")

        def calc_row(row):
            values = {col: row[col] for col in variables}
            return self.compute_engine.evaluate_expression(self.expression, values)
        df["Result"] = df.apply(calc_row, axis=1)

        with ExcelFile(self.filepath) as ef:
            ws = ef.active_sheet
            # 处理结果列
            try:
                result_col_idx = openpyxl.utils.column_index_from_string(self.result_col_input)
                is_new_col = False
                result_col_name = self.result_col_input
            except:
                is_new_col = True
                result_col_name = self.result_col_input.strip()
                max_col = ws.max_column
                result_col_idx = max_col + 1
                ws.insert_cols(result_col_idx)
                self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

            if not is_new_col:
                existing = self._safe_get_cell_value(ws, self.header_row, result_col_idx)
                if existing is None or str(existing).strip() == "":
                    self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

            start_row, end_row = self.header_row + 1, ws.max_row
            for i, val in enumerate(df["Result"]):
                ws.cell(start_row + i, result_col_idx).value = val

            if self.merge_by_cols_str:
                merge_keys = [openpyxl.utils.column_index_from_string(c.strip()) for c in self.merge_by_cols_str.split(",")]
                if merge_keys:
                    copy_merge_structure(ws, merge_keys[0], result_col_idx, start_row, end_row)

            base_name = os.path.splitext(os.path.basename(self.filepath))[0]
            output_filename = f"{base_name}_处理版.xlsx"
            result_path = os.path.join(os.path.dirname(self.filepath), output_filename)
            ef.save(result_path)
            ef._workbook.close()
            gc.collect()
            return result_path

    @staticmethod
    def _safe_get_cell_value(ws, row, col):
        cell = ws.cell(row, col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                    return ws.cell(merged_range.min_row, merged_range.min_col).value
        return cell.value

    @staticmethod
    def _safe_write_cell(ws, row, col, value):
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


class AggregateBackfillProcessor:
    """聚合回填处理器"""
    def __init__(self, filepath, header_row, group_cols_raw, agg_cols_raw, agg_func, result_col_input, compute_engine):
        self.filepath = filepath
        self.header_row = header_row
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()

    def process(self):
        # 构建列字母→列名映射
        try:
            df_header = pd.read_excel(self.filepath, header=self.header_row - 1, nrows=0)
        except Exception as e:
            raise ValueError(f"读取表头失败: {str(e)}")
        headers = list(df_header.columns)
        letter_to_name, name_to_letter = build_column_mapping(headers)

        def resolve_col(col):
            col = col.strip()
            if col in headers:
                return col
            upper = col.upper()
            if upper in letter_to_name:
                return letter_to_name[upper]
            return col

        group_cols = [resolve_col(c) for c in self.group_cols_raw]
        agg_cols = [resolve_col(c) for c in self.agg_cols_raw]

        df = pd.read_excel(self.filepath, header=self.header_row - 1, dtype=str)
        df = self.dal._auto_clean_dataframe(df)

        missing = [c for c in group_cols + agg_cols if c not in df.columns]
        if missing:
            raise ValueError(f"列不存在: {missing}")

        result_df = self.compute_engine.simple_aggregate(df, group_cols, agg_cols, self.agg_func)

        # 回填到原始 Excel
        with ExcelFile(self.filepath) as ef:
            ws = ef.active_sheet
            # 建立 header_row 行的列名→列索引映射
            name_to_col_idx = {}
            for idx, cell in enumerate(ws[self.header_row], start=1):
                if cell.value is not None:
                    name_to_col_idx[str(cell.value).strip()] = idx

            def get_col_index(col):
                col = col.strip()
                if col in name_to_col_idx:
                    return name_to_col_idx[col]
                try:
                    return openpyxl.utils.column_index_from_string(col)
                except ValueError:
                    raise ValueError(f"列 '{col}' 在原始表中不存在")

            group_col_indices = [get_col_index(c) for c in group_cols]

            # 处理结果列
            try:
                result_col_idx = openpyxl.utils.column_index_from_string(self.result_col_input)
                is_new_col = False
                result_col_name = self.result_col_input
            except:
                is_new_col = True
                result_col_name = self.result_col_input.strip()
                result_col_idx = ws.max_column + 1
                ws.insert_cols(result_col_idx)
                self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

            if not is_new_col:
                existing = self._safe_get_cell_value(ws, self.header_row, result_col_idx)
                if existing is None or str(existing).strip() == "":
                    self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

            # 构建聚合结果映射
            result_map = {}
            for _, row in result_df.iterrows():
                key = tuple(str(row[c]) for c in group_cols)
                result_map[key] = row[agg_cols[0]]

            start_row, end_row = self.header_row + 1, ws.max_row
            for row in range(start_row, end_row + 1):
                key = tuple(str(ws.cell(row, idx).value) for idx in group_col_indices)
                ws.cell(row, result_col_idx).value = result_map.get(key, "")

            base_name = os.path.splitext(os.path.basename(self.filepath))[0]
            output_filename = f"{base_name}_处理版.xlsx"
            result_path = os.path.join(os.path.dirname(self.filepath), output_filename)
            ef.save(result_path)
            ef._workbook.close()
            gc.collect()
            return result_path

    @staticmethod
    def _safe_get_cell_value(ws, row, col):
        cell = ws.cell(row, col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                    return ws.cell(merged_range.min_row, merged_range.min_col).value
        return cell.value

    @staticmethod
    def _safe_write_cell(ws, row, col, value):
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