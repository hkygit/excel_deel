# core/excel_utils.py
"""统一的 Excel 工具：列转换、文件包装、合并单元格复制"""
import re
import logging
from typing import Union, List, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from copy import copy

logger = logging.getLogger(__name__)

def col_letter_to_index(letter: str) -> int:
    return column_index_from_string(letter.upper())

def col_index_to_letter(index: int) -> str:
    return get_column_letter(index)

def validate_column_letters(letters: List[str]) -> bool:
    pattern = re.compile(r'^[A-Z]+$')
    return all(pattern.match(l.upper()) for l in letters)

class ExcelFile:
    def __init__(self, filepath: str, data_only: bool = True):
        self.filepath = filepath
        self.data_only = data_only
        self._workbook: Optional[openpyxl.Workbook] = None
        self._active_sheet = None

    def __enter__(self):
        self.load()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def load(self) -> 'ExcelFile':
        self._workbook = openpyxl.load_workbook(self.filepath, data_only=self.data_only)
        return self

    def close(self):
        if self._workbook:
            self._workbook.close()
            self._workbook = None

    def save(self, filepath: str = None) -> str:
        if not self._workbook:
            raise ValueError("未加载工作簿")
        target = filepath or self.filepath
        self._workbook.save(target)
        return target

    @property
    def active_sheet(self):
        if self._active_sheet is None:
            self._active_sheet = self._workbook.active
        return self._active_sheet

    def get_sheet(self, name_or_index: Union[str, int]):
        if isinstance(name_or_index, int):
            return self._workbook.worksheets[name_or_index]
        return self._workbook[name_or_index]

    def get_headers(self, sheet=None, row: int = 1) -> List[str]:
        ws = sheet or self.active_sheet
        headers = []
        for cell in ws[row]:
            if cell.value is None:
                break
            headers.append(str(cell.value).strip())
        return headers

    def get_cell_value(self, row: int, col: Union[int, str], sheet=None) -> Any:
        ws = sheet or self.active_sheet
        if isinstance(col, str):
            col = col_letter_to_index(col)
        cell = ws.cell(row, col)
        for merged in ws.merged_cells.ranges:
            if cell.coordinate in merged:
                return ws.cell(merged.min_row, merged.min_col).value
        return cell.value

    def copy_merge_structure(self, src_col: Union[int, str], dst_col: Union[int, str],
                             start_row: int, end_row: int, sheet=None):
        ws = sheet or self.active_sheet
        if isinstance(src_col, str):
            src_col = col_letter_to_index(src_col)
        if isinstance(dst_col, str):
            dst_col = col_letter_to_index(dst_col)

        src_ranges = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= src_col <= merged.max_col:
                if merged.min_row <= end_row and merged.max_row >= start_row:
                    src_ranges.append((merged.min_row, merged.max_row))

        to_remove = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= dst_col <= merged.max_col:
                to_remove.append(merged)
        for merged in to_remove:
            ws.unmerge_cells(str(merged))

        for min_r, max_r in src_ranges:
            if min_r < max_r:
                ws.merge_cells(start_row=min_r, start_column=dst_col,
                               end_row=max_r, end_column=dst_col)

    @staticmethod
    def read_dataframe(filepath: str, sheet: Union[str, int] = 0, **kwargs) -> pd.DataFrame:
        try:
            return pd.read_excel(filepath, sheet_name=sheet, **kwargs)
        except Exception as e:
            logger.error(f"读取Excel失败 {filepath}: {e}")
            raise ValueError(f"无法读取Excel文件: {e}")

    @staticmethod
    def write_dataframe(df: pd.DataFrame, filepath: str, sheet_name: str = 'Sheet1'):
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return filepath

def copy_merge_structure(src_ws, src_col, dst_ws, dst_col, start_row, end_row):
    """兼容旧接口的独立函数"""
    ef = ExcelFile('')  # 占位，仅用于调用实例方法
    ef._workbook = src_ws.parent
    ef.copy_merge_structure(src_col, dst_col, start_row, end_row, sheet=dst_ws)


# ========== 以下为兼容 ExcelService 而添加的静态方法 ==========

    @staticmethod
    def copy_cell_style(src_cell, dst_cell):
        """复制单元格样式（兼容 ExcelFile.copy_cell_style）"""
        from copy import copy
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.border = copy(src_cell.border)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.number_format = src_cell.number_format
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.protection = copy(src_cell.protection)
        dst_cell.value = src_cell.value

    @staticmethod
    def copy_merge_structure_static(ws, src_col, dst_col, start_row, end_row):
        """
        静态合并单元格结构复制（兼容 copy_merge_structure）
        :param ws: 工作表对象
        :param src_col: 源列（字母或数字索引）
        :param dst_col: 目标列（字母或数字索引）
        :param start_row: 起始行
        :param end_row: 结束行
        """
        # 复用 ExcelFile 的实例方法，创建一个临时对象
        ef = ExcelFile("")
        ef._workbook = ws.parent
        ef.copy_merge_structure(src_col, dst_col, start_row, end_row, sheet=ws)

    @staticmethod
    def get_cell_value_static(ws, row, col):
        """
        静态获取单元格值（兼容 ExcelFile.get_cell_value）
        :param ws: 工作表对象
        :param row: 行号
        :param col: 列字母或数字索引
        """
        ef = ExcelFile("")
        ef._workbook = ws.parent
        return ef.get_cell_value(row, col, sheet=ws)
