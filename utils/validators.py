# utils/validators.py
"""向后兼容：委托给 core/validators.py"""
from core.validators import (
    validate_column_letters,
    validate_sheet_index,
    validate_expression,
    sanitize_filename
)
