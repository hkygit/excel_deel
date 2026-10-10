#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修复脚本 v3：_remove_result_column_from_expr 改用 token 级匹配

问题：
    services/excel/calculator_processor.py 中的 _remove_result_column_from_expr
    使用 `expr.replace(result_col_letter, '')` 做子串替换，会把"含有"结果列字母
    的标识符一并破坏：
        'AA+B' 结果列 'A' → '+B'   (AA 被误拆)
        'AB+C' 结果列 'B' → 'A+C'  (AB 被误拆)

修复：
    改用 token 级匹配。将表达式分词为 [标识符 | 数字 | 运算符 | 括号]，
    只移除"恰好等于结果列字母的独立标识符 token"，并连带移除其邻近的
    二元运算符。这样：
        'A+B'   结果列 'A' → 'B'
        'AA+B'  结果列 'A' → 'AA+B'   (AA 完整保留)
        'AB+C'  结果列 'B' → 'AB+C'   (AB 完整保留)
        'A+B-C' 结果列 'B' → 'A-C'

本脚本只写文件，不运行测试。
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CALCULATOR_PROCESSOR_PATH = os.path.join(
    BASE_DIR, "services", "excel", "calculator_processor.py"
)


# ============================================================================
# 新的 _remove_result_column_from_expr 方法实现
# ============================================================================
# 使用原始字符串，避免 \\d \\+ \\- 等转义问题
NEW_METHOD = r'''    def _remove_result_column_from_expr(self, result_col_letter: str) -> str:
        """从表达式中移除与结果列字母"完全匹配"的标识符 token。

        使用 token 级匹配（而不是 str.replace 子串替换），避免破坏那些
        只是"包含"结果列字母的标识符。例如结果列为 'A' 时：
            'A+B'   -> 'B'        移除独立的 'A'
            'AA+B'  -> 'AA+B'     'AA' 是独立标识符，不受影响
            'AB+A'  -> 'AB'       仅移除独立的 'A'，不动 'AB'
            'A+B-C' -> 'A-C'      移除 'B' 及其邻近的二元运算符

        当某个匹配的标识符被移除时，会一并移除它邻近的二元运算符，避免
        残留 '(A+)*C'、'A+-C' 这类非法表达式。
        """
        if not result_col_letter or not self.expression:
            return self.expression

        target = result_col_letter.upper()
        expr_upper = self.expression.upper()

        # 分词：标识符（首字符字母）、数字（含小数）、运算符、括号
        import re
        token_re = re.compile(r'[A-Z][A-Z0-9]*|\d+(?:\.\d+)?|[()+\-*/]')
        tokens = token_re.findall(expr_upper)

        # 快速路径：没有与结果列完全相等的标识符 token，无需移除
        if not any(t == target and t.isalpha() for t in tokens):
            return self.expression

        binary_ops = {'+', '-', '*', '/'}

        # 标记需要移除的 token 下标
        to_remove = set()
        n = len(tokens)
        for i, tok in enumerate(tokens):
            if tok == target and tok.isalpha():
                to_remove.add(i)
                # 优先连同前面的二元运算符一起移除
                if i > 0 and tokens[i - 1] in binary_ops:
                    to_remove.add(i - 1)
                # 否则，连同后面的二元运算符一起移除
                elif i + 1 < n and tokens[i + 1] in binary_ops:
                    to_remove.add(i + 1)

        remaining = [tok for i, tok in enumerate(tokens) if i not in to_remove]

        # 清理首尾残留的运算符
        while remaining and remaining[0] in binary_ops:
            remaining.pop(0)
        while remaining and remaining[-1] in binary_ops:
            remaining.pop()

        new_expr = ''.join(remaining).strip()
        if not new_expr:
            new_expr = '0'

        if new_expr != expr_upper:
            logger.warning(
                f"表达式中包含了结果列 {target}，已自动移除: "
                f"{self.expression} -> {new_expr}"
            )

        return new_expr

'''


def patch_calculator_processor() -> None:
    """用新的 token 级实现替换旧的 _remove_result_column_from_expr。"""
    if not os.path.exists(CALCULATOR_PROCESSOR_PATH):
        raise FileNotFoundError(f"找不到文件: {CALCULATOR_PROCESSOR_PATH}")

    with open(CALCULATOR_PROCESSOR_PATH, "r", encoding="utf-8") as f:
        content = f.read()

    # 定位方法起点
    start_marker = (
        "    def _remove_result_column_from_expr(self, result_col_letter: str) -> str:"
    )
    start_idx = content.find(start_marker)
    if start_idx == -1:
        raise RuntimeError(
            "未找到 _remove_result_column_from_expr 方法定义，"
            "可能文件已被修改或方法签名不同。"
        )

    # 定位方法终点：下一个同缩进的方法定义（"    def "）
    end_search = content.find("\n    def ", start_idx + len(start_marker))
    if end_search == -1:
        raise RuntimeError(
            "未能在 _remove_result_column_from_expr 之后找到下一个方法定义，"
            "无法确定替换范围。"
        )
    # end_search 指向 '\n'，content[end_search:] 从 '\n    def ...' 开始
    end_idx = end_search

    new_content = content[:start_idx] + NEW_METHOD + content[end_idx + 1:]

    with open(CALCULATOR_PROCESSOR_PATH, "w", encoding="utf-8") as f:
        f.write(new_content)

    print(f"[OK] {CALCULATOR_PROCESSOR_PATH}: 已替换 _remove_result_column_from_expr")


# ============================================================================
# 测试文件
# ============================================================================
TEST_CONTENT = r'''# tests/test_remove_result_column_expr.py
"""
针对 ColumnCalculatorProcessor._remove_result_column_from_expr 的回归测试。

核心目标：
    验证"用 token 匹配替换子串 replace"这一修复，确保：
    - 恰好等于结果列的独立标识符被移除
    - 只是包含结果列字母的标识符（如 'AA', 'AB'）不被破坏
    - 移除时连带清理邻近的二元运算符，不产生语法上非法的表达式
"""
import pytest

from services.excel.calculator_processor import ColumnCalculatorProcessor
from services.excel.compute_engine import ComputeEngine


def _make_processor(expression: str) -> ColumnCalculatorProcessor:
    """构造一个仅用于测试 _remove_result_column_from_expr 的处理器实例。"""
    return ColumnCalculatorProcessor(
        filepath="dummy.xlsx",      # 不用于本测试
        header_row=1,
        expression=expression,
        result_col_input="Z",        # 不用于本测试
        merge_by_cols_str="",
        compute_engine=ComputeEngine(),
    )


class TestRemoveResultColumnFromExpr:
    # ---------- 快速路径：无匹配 ----------
    def test_no_match_returns_original(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("D") == "A+B"

    def test_empty_expression(self):
        p = _make_processor("")
        assert p._remove_result_column_from_expr("A") == ""

    def test_empty_result_col(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("") == "A+B"

    # ---------- 关键回归：不得破坏子串标识符 ----------
    def test_aa_identifier_not_corrupted_by_removing_a(self):
        """'AA+B' 移除 'A'，'AA' 必须完整保留。"""
        p = _make_processor("AA+B")
        assert p._remove_result_column_from_expr("A") == "AA+B"

    def test_ab_identifier_not_corrupted_by_removing_a(self):
        """'AB+C' 移除 'A'，'AB' 必须完整保留。"""
        p = _make_processor("AB+C")
        assert p._remove_result_column_from_expr("A") == "AB+C"

    def test_ab_identifier_not_corrupted_by_removing_b(self):
        """'AB+C' 移除 'B'，'AB' 必须完整保留。"""
        p = _make_processor("AB+C")
        assert p._remove_result_column_from_expr("B") == "AB+C"

    def test_multi_letter_identifiers_untouched(self):
        """'AB+AC' 里没有任何 token 恰好等于 'A'。"""
        p = _make_processor("AB+AC")
        assert p._remove_result_column_from_expr("A") == "AB+AC"

    # ---------- 标准移除：独立标识符 ----------
    def test_removes_standalone_at_start(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("A") == "B"

    def test_removes_standalone_at_end(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("B") == "A"

    def test_removes_standalone_in_middle(self):
        p = _make_processor("A+B-C")
        assert p._remove_result_column_from_expr("B") == "A-C"

    def test_removes_in_longer_chain(self):
        p = _make_processor("A+B+C+D")
        assert p._remove_result_column_from_expr("C") == "A+B+D"

    # ---------- 括号与一元运算符 ----------
    def test_removes_with_parens(self):
        p = _make_processor("(A+B)*C")
        assert p._remove_result_column_from_expr("B") == "(A)*C"

    def test_removes_inside_nested_parens(self):
        p = _make_processor("((A+B)-C)")
        assert p._remove_result_column_from_expr("B") == "((A)-C)"

    # ---------- 边界情况 ----------
    def test_whole_expression_is_result_column(self):
        """表达式只有一个 token，且正是结果列本身。"""
        p = _make_processor("A")
        assert p._remove_result_column_from_expr("A") == "0"

    def test_lowercase_input_is_normalized(self):
        """小写输入应被正确识别并移除。"""
        p = _make_processor("a+b")
        assert p._remove_result_column_from_expr("A") == "B"

    def test_number_literals_are_not_mistaken_for_identifiers(self):
        """数字 token 不应被当作标识符。'A*2+B' 移除 'A' → '2+B'。"""
        p = _make_processor("A*2+B")
        assert p._remove_result_column_from_expr("A") == "2+B"

    def test_decimal_number_literal_untouched(self):
        p = _make_processor("A*1.5+B")
        assert p._remove_result_column_from_expr("A") == "1.5+B"
'''


def write_test_file() -> None:
    """写入新的单元测试文件。"""
    path = os.path.join(BASE_DIR, "tests", "test_remove_result_column_expr.py")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(TEST_CONTENT)
    print(f"[OK] {path}: 已写入")


# ============================================================================
# main
# ============================================================================
def main() -> None:
    print("=" * 72)
    print("修复脚本 v3")
    print("  · calculator_processor.py::_remove_result_column_from_expr")
    print("    由子串 replace 改为 token 级精确匹配")
    print("  · 新增 tests/test_remove_result_column_expr.py")
    print("=" * 72)

    patch_calculator_processor()
    write_test_file()

    print()
    print("=" * 72)
    print("✅ 修复完成（未运行测试）")
    print("=" * 72)
    print()
    print("请手动执行验证：")
    print("    python -m pytest tests/test_remove_result_column_expr.py -v")
    print("    python -m pytest tests/test_integration_calculator.py -v")
    print("    python -m pytest tests/ -v")


if __name__ == "__main__":
    main()