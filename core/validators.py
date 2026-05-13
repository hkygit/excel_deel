# core/validators.py
"""通用验证函数"""
import re
from typing import List, Union

def validate_column_letters(letters: List[str]) -> bool:
    pattern = re.compile(r'^[A-Z]+$')
    return all(pattern.match(l.upper()) for l in letters)

def validate_sheet_index(idx: Union[str, int]) -> bool:
    try:
        num = int(idx)
        return num >= 1
    except (ValueError, TypeError):
        return False

def validate_expression(expr: str) -> bool:
    if not expr or not expr.strip():
        return False
    allowed_chars = set('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789+-*/() ')
    if not all(c in allowed_chars for c in expr.upper()):
        return False
    if expr.strip().endswith(('+', '-', '*', '/', '(', ' ')):
        return False
    stack = []
    for ch in expr:
        if ch == '(':
            stack.append(ch)
        elif ch == ')':
            if not stack:
                return False
            stack.pop()
    return len(stack) == 0

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', '_', str(name))
