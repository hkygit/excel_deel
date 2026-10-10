# tests/test_remove_result_column_expr.py
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
