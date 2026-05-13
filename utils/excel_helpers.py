# utils/excel_helpers.py
"""向后兼容层：委托给 core.excel_utils"""
from core.excel_utils import (
    ExcelFile, col_letter_to_index, col_index_to_letter,
    validate_column_letters, copy_merge_structure
)
# 保留旧的导入路径，确保其他模块无需修改
__all__ = ['ExcelFile', 'col_letter_to_index', 'col_index_to_letter', 'validate_column_letters', 'copy_merge_structure']
