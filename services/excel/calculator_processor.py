#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 列计算与聚合回填处理器
支持 .xls 文件自动转换为 .xlsx
"""

import os
import gc
import logging
import tempfile
import openpyxl
import pandas as pd
import re
from core.excel_utils import ExcelFile, copy_merge_structure
from core.data_access_layer import get_data_access_layer
from core.csv_utils import build_column_mapping
from services.excel.compute_engine import ComputeEngine

logger = logging.getLogger(__name__)

def convert_xls_to_xlsx(filepath):
    """将 .xls 文件转换为临时 .xlsx 文件"""
    if not filepath.lower().endswith('.xls') or filepath.lower().endswith('.xlsx'):
        return filepath
    logger.info(f"正在将 .xls 文件转换为 .xlsx: {filepath}")
    try:
        # 使用 pandas 读取 .xls 并写入临时 .xlsx
        df = pd.read_excel(filepath, sheet_name=None, dtype=str)
        fd, temp_xlsx = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
            for sheet_name, sheet_df in df.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info(f"转换成功: {temp_xlsx}")
        return temp_xlsx
    except Exception as e:
        logger.error(f"转换 .xls 文件失败: {e}")
        raise ValueError(f"无法转换 .xls 文件，请确保文件未损坏且可读: {e}")

class ColumnCalculatorProcessor:
    """列计算处理器（自动过滤结果列，支持 .xls）"""
    def __init__(self, filepath, header_row, expression, result_col_input, merge_by_cols_str, compute_engine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.original_expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols_str = merge_by_cols_str
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_file = None

    def _remove_result_column_from_expr(self, expression, result_col_letter):
        """彻底从表达式中移除结果列字母"""
        if not result_col_letter or result_col_letter not in expression:
            return expression
        pattern = r'(?:[+\-*/]?)' + re.escape(result_col_letter) + r'(?:[+\-*/]?)'
        new_expr = re.sub(pattern, '', expression)
        new_expr = re.sub(r'[+\-*/]{2,}', '', new_expr)
        new_expr = re.sub(r'^[+\-*/]', '', new_expr)
        new_expr = re.sub(r'[+\-*/]$', '', new_expr)
        if new_expr != expression:
            logger.warning(f"表达式中包含了结果列 {result_col_letter}，已自动移除: {expression} -> {new_expr}")
        return new_expr

    def process(self):
        # 0. 如果是 .xls 文件，先转换为临时 .xlsx
        ext = os.path.splitext(self.original_filepath)[1].lower()
        if ext == '.xls':
            self.temp_file = convert_xls_to_xlsx(self.original_filepath)
            filepath = self.temp_file
        else:
            filepath = self.original_filepath

        try:
            # 1. 读取 DataFrame 并获取列名
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = self.dal._auto_clean_dataframe(df)
            headers = list(df.columns)

            # 2. 构建列字母到列名的映射
            letter_to_name, name_to_letter = build_column_mapping(headers)
            logger.debug(f"列字母映射: {letter_to_name}")

            # 3. 获取结果列字母
            result_col_letter = None
            try:
                col_idx = openpyxl.utils.column_index_from_string(self.result_col_input)
                result_col_letter = openpyxl.utils.get_column_letter(col_idx)
            except:
                pass

            # 4. 自动移除结果列字母
            expression = self.original_expression
            if result_col_letter:
                expression = self._remove_result_column_from_expr(expression, result_col_letter)
                if not expression.strip():
                    raise ValueError(f"表达式在移除结果列 {result_col_letter} 后为空")

            # 5. 提取并映射变量
            variables = self.compute_engine.extract_variables(expression)
            logger.info(f"表达式中的列字母: {variables}")
            mapped_vars = {}
            missing = []
            for letter in variables:
                if letter in letter_to_name:
                    mapped_vars[letter] = letter_to_name[letter]
                else:
                    missing.append(letter)
            if missing:
                raise ValueError(f"表达式中的列字母 {missing} 无法映射到文件中的任何列（可用列字母: {list(letter_to_name.keys())}）")

            # 6. 计算
            def calc_row(row):
                context = {letter: row[col_name] for letter, col_name in mapped_vars.items()}
                return self.compute_engine.evaluate_expression(expression, context)
            df["Result"] = df.apply(calc_row, axis=1)
            logger.info("表达式计算完成")

            # 7. 打开原始 Excel 写入（注意：如果是转换后的临时文件，原始文件无法写入，需要将结果写回原始文件？）
            # 策略：如果原文件是 .xls，则最终输出新的 .xlsx 文件，保留在原文件同目录
            if ext == '.xls':
                # 原文件是 .xls，不能直接用 openpyxl 修改，我们直接基于转换后的临时文件写入结果，保存到新文件
                output_base = os.path.splitext(os.path.basename(self.original_filepath))[0]
                output_dir = os.path.dirname(self.original_filepath)
                output_filename = f"{output_base}_处理版.xlsx"
                result_path = os.path.join(output_dir, output_filename)
                # 将计算结果写入临时文件并另存
                with ExcelFile(filepath) as ef:
                    ws = ef.active_sheet
                    # 处理结果列...
                    # 这里需要与下面逻辑类似，但为了避免重复，复用同一份代码块较复杂。
                    # 鉴于时间，我们简单说明：可以调用一个通用写入函数。
                    # 为了快速修复，我们假设原文件不是 .xls，但实际需要处理。这里提供完善方案。
                    # 由于代码量较大，我将在下面统一写入代码块中处理两种情况。
                pass

            # 统一写入逻辑（适配原始文件和临时文件）
            with ExcelFile(filepath) as ef:
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

                start_row = self.header_row + 1
                end_row = ws.max_row

                # 清除结果列合并
                to_remove = []
                for merged in ws.merged_cells.ranges:
                    if merged.min_col <= result_col_idx <= merged.max_col:
                        to_remove.append(merged)
                for merged in to_remove:
                    ws.unmerge_cells(str(merged))

                # 写入计算结果
                for i, val in enumerate(df["Result"]):
                    ws.cell(start_row + i, result_col_idx).value = val

                # 处理合并依据列
                if self.merge_by_cols_str:
                    merge_keys = [openpyxl.utils.column_index_from_string(c.strip()) for c in self.merge_by_cols_str.split(",") if c.strip()]
                    if merge_keys:
                        ref_col = merge_keys[0]
                        has_merge = any(m.min_col <= ref_col <= m.max_col for m in ws.merged_cells.ranges)
                        if has_merge:
                            try:
                                copy_merge_structure(ws, ref_col, ws, result_col_idx, start_row, end_row)
                                logger.info(f"已复制合并结构 from {self.merge_by_cols_str} to {self.result_col_input}")
                            except Exception as e:
                                logger.warning(f"复制合并结构失败: {e}")
                        else:
                            logger.info(f"源列 {self.merge_by_cols_str} 无合并单元格，跳过")

                # 确定输出路径
                if ext == '.xls':
                    # 对于 .xls 文件，输出到原目录下，文件名加“_处理版.xlsx”
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")
                else:
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")

                ef.save(result_path)
                ef._workbook.close()

            gc.collect()
            return result_path

        finally:
            # 清理临时文件（如果创建了）
            if self.temp_file and os.path.exists(self.temp_file):
                try:
                    os.unlink(self.temp_file)
                except:
                    pass

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
    """聚合回填处理器（同样支持 .xls 转换）"""
    def __init__(self, filepath, header_row, group_cols_raw, agg_cols_raw, agg_func, result_col_input, compute_engine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_file = None

    def process(self):
        ext = os.path.splitext(self.original_filepath)[1].lower()
        if ext == '.xls':
            self.temp_file = convert_xls_to_xlsx(self.original_filepath)
            filepath = self.temp_file
        else:
            filepath = self.original_filepath

        try:
            # 读取表头获取列名
            df_header = pd.read_excel(filepath, header=self.header_row - 1, nrows=0)
            headers = list(df_header.columns)
            letter_to_name, name_to_letter = build_column_mapping(headers)

            def resolve_col(col):
                col = col.strip()
                if col in headers:
                    return col
                upper = col.upper()
                if upper in letter_to_name:
                    return letter_to_name[upper]
                raise ValueError(f"列 '{col}' 无法映射到实际列名")

            group_cols = [resolve_col(c) for c in self.group_cols_raw]
            agg_cols = [resolve_col(c) for c in self.agg_cols_raw]

            # 读取数据并清洗
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = self.dal._auto_clean_dataframe(df)

            missing = [c for c in group_cols + agg_cols if c not in df.columns]
            if missing:
                raise ValueError(f"列不存在: {missing}")

            result_df = self.compute_engine.simple_aggregate(df, group_cols, agg_cols, self.agg_func)

            # 回填到原始 Excel
            with ExcelFile(filepath) as ef:
                ws = ef.active_sheet
                # 建立列名到列索引的映射
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

                start_row = self.header_row + 1
                end_row = ws.max_row
                for row in range(start_row, end_row + 1):
                    key = tuple(str(ws.cell(row, idx).value) for idx in group_col_indices)
                    ws.cell(row, result_col_idx).value = result_map.get(key, "")

                if ext == '.xls':
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")
                else:
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")

                ef.save(result_path)
                ef._workbook.close()

            gc.collect()
            return result_path
        finally:
            if self.temp_file and os.path.exists(self.temp_file):
                try:
                    os.unlink(self.temp_file)
                except:
                    pass

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
