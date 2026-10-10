#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P0-2 修复脚本：消除 calc 模块的重复 normalize

改动：
1. services/excel/calculator_processor.py
   - BaseExcelProcessor.__init__ 新增 already_normalized 参数
   - _get_work_filepath 在该参数为 True 时直接返回原路径，不调用 normalize
   - ColumnCalculatorProcessor / AggregateBackfillProcessor 透传该参数
2. blueprints/calculator.py
   - 构造两个 processor 时显式传入 already_normalized=True
3. 新增 tests/test_calc_no_double_normalize.py
   - 断言 already_normalized=True 时不会调用 normalize_excel_file
   - 断言默认行为（already_normalized=False）仍然调用 normalize
   - 断言原路径未被误加入 temp_files

本脚本采用精确字符串替换，不重写整个文件，避免大文件重写风险。
本脚本只写文件，不运行测试。
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PROCESSOR_PATH = os.path.join(
    BASE_DIR, "services", "excel", "calculator_processor.py"
)
CALCULATOR_BP_PATH = os.path.join(BASE_DIR, "blueprints", "calculator.py")
TEST_PATH = os.path.join(
    BASE_DIR, "tests", "test_calc_no_double_normalize.py"
)


# ============================================================================
# 精确替换定义
# ============================================================================
PROCESSOR_REPLACEMENTS = [
    # ---------- BaseExcelProcessor.__init__ ----------
    (
        """    def __init__(self, filepath: str, header_row: int, compute_engine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_files = []  # 跟踪所有临时文件
        self._work_filepath = None

        # 注册清理函数
        atexit.register(self._cleanup_all)
""",
        """    def __init__(self, filepath: str, header_row: int, compute_engine,
                 already_normalized: bool = False):
        self.original_filepath = filepath
        self.header_row = header_row
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_files = []  # 跟踪所有临时文件
        self._work_filepath = None
        # P0-2: 上游（blueprints/*.py）若已调用 normalize_excel_file，
        # 应显式传入 already_normalized=True，避免重复规范化。
        self.already_normalized = already_normalized

        # 注册清理函数
        atexit.register(self._cleanup_all)
""",
    ),
    # ---------- _get_work_filepath ----------
    (
        '''    def _get_work_filepath(self) -> str:
        """
        获取工作文件路径，确保文件可读
        """
        if self._work_filepath and os.path.exists(self._work_filepath):
            return self._work_filepath

        # 检查文件是否存在
        if not os.path.exists(self.original_filepath):
            raise FileNotFoundError(f"原始文件不存在: {self.original_filepath}")

        # 总是进行规范化，确保文件格式正确
        try:
            self._work_filepath = normalize_excel_file(self.original_filepath)
            if self._work_filepath and os.path.exists(self._work_filepath):
                # 记录临时文件以便后续清理
                if self._work_filepath != self.original_filepath:
                    self.temp_files.append(self._work_filepath)
                return self._work_filepath
            else:
                raise ValueError(f"规范化文件创建失败: {self._work_filepath}")
        except Exception as e:
            logger.error(f"获取工作文件路径失败: {e}")
            raise
''',
        '''    def _get_work_filepath(self) -> str:
        """
        获取工作文件路径，确保文件可读。

        P0-2: 上游若已通过 already_normalized=True 声明文件已规范化，
        则直接复用原路径，跳过二次 normalize（消除一次完整读写）。
        """
        if self._work_filepath and os.path.exists(self._work_filepath):
            return self._work_filepath

        # 检查文件是否存在
        if not os.path.exists(self.original_filepath):
            raise FileNotFoundError(f"原始文件不存在: {self.original_filepath}")

        # P0-2: 上游已规范化的路径直接复用，零复制
        if self.already_normalized:
            if not self.original_filepath.lower().endswith('.xlsx'):
                # 上游声明"已规范化"但扩展名不是 .xlsx —— 保守回退到 normalize
                logger.warning(
                    f"already_normalized=True 但文件不是 .xlsx "
                    f"({self.original_filepath})，回退到 normalize_excel_file"
                )
            else:
                self._work_filepath = self.original_filepath
                logger.debug(
                    f"已规范化，直接复用原路径: {self.original_filepath}"
                )
                return self._work_filepath

        # 兜底路径：独立调用（未声明 already_normalized）时做完整规范化
        try:
            self._work_filepath = normalize_excel_file(self.original_filepath)
            if self._work_filepath and os.path.exists(self._work_filepath):
                # 记录临时文件以便后续清理（仅当确实新建了临时文件）
                if self._work_filepath != self.original_filepath:
                    self.temp_files.append(self._work_filepath)
                return self._work_filepath
            else:
                raise ValueError(f"规范化文件创建失败: {self._work_filepath}")
        except Exception as e:
            logger.error(f"获取工作文件路径失败: {e}")
            raise
''',
    ),
    # ---------- ColumnCalculatorProcessor.__init__ ----------
    (
        """    def __init__(self, filepath: str, header_row: int, expression: str,
                 result_col_input: str, merge_by_cols_str: str, compute_engine):
        super().__init__(filepath, header_row, compute_engine)
        self.expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols = [c.strip() for c in merge_by_cols_str.split(',') if c.strip()] if merge_by_cols_str else []
""",
        """    def __init__(self, filepath: str, header_row: int, expression: str,
                 result_col_input: str, merge_by_cols_str: str, compute_engine,
                 already_normalized: bool = False):
        super().__init__(filepath, header_row, compute_engine,
                         already_normalized=already_normalized)
        self.expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols = [c.strip() for c in merge_by_cols_str.split(',') if c.strip()] if merge_by_cols_str else []
""",
    ),
    # ---------- AggregateBackfillProcessor.__init__ ----------
    (
        """    def __init__(self, filepath: str, header_row: int,
                 group_cols_raw: List[str], agg_cols_raw: List[str],
                 agg_func: str, result_col_input: str, compute_engine):
        super().__init__(filepath, header_row, compute_engine)
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input
""",
        """    def __init__(self, filepath: str, header_row: int,
                 group_cols_raw: List[str], agg_cols_raw: List[str],
                 agg_func: str, result_col_input: str, compute_engine,
                 already_normalized: bool = False):
        super().__init__(filepath, header_row, compute_engine,
                         already_normalized=already_normalized)
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input
""",
    ),
]


CALCULATOR_BP_REPLACEMENTS = [
    # ---------- ColumnCalculatorProcessor 调用点 ----------
    (
        """            processor = ColumnCalculatorProcessor(
                filepath=filepath,
                header_row=header_row,
                expression=expression,
                result_col_input=result_col,
                merge_by_cols_str=merge_by_cols,
                compute_engine=compute_engine
            )""",
        """            processor = ColumnCalculatorProcessor(
                filepath=filepath,
                header_row=header_row,
                expression=expression,
                result_col_input=result_col,
                merge_by_cols_str=merge_by_cols,
                compute_engine=compute_engine,
                already_normalized=True,
            )""",
    ),
    # ---------- AggregateBackfillProcessor 调用点 ----------
    (
        """            processor = AggregateBackfillProcessor(
                filepath=filepath,
                header_row=header_row,
                group_cols_raw=group_cols_raw,
                agg_cols_raw=agg_cols_raw,
                agg_func=agg_func,
                result_col_input=result_col,
                compute_engine=compute_engine
            )""",
        """            processor = AggregateBackfillProcessor(
                filepath=filepath,
                header_row=header_row,
                group_cols_raw=group_cols_raw,
                agg_cols_raw=agg_cols_raw,
                agg_func=agg_func,
                result_col_input=result_col,
                compute_engine=compute_engine,
                already_normalized=True,
            )""",
    ),
]


# ============================================================================
# 应用替换
# ============================================================================
def apply_replacements(path: str, replacements: list, label: str) -> None:
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到文件: {path}")

    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    for i, (old, new) in enumerate(replacements, 1):
        if old not in content:
            raise RuntimeError(
                f"{label}: 第 {i} 个替换锚点未找到。\n"
                f"可能是文件已被修改。锚点前 80 字符:\n"
                f"{old[:80]!r}"
            )
        # 每个锚点只应出现一次，避免误伤
        if content.count(old) > 1:
            raise RuntimeError(
                f"{label}: 第 {i} 个锚点在文件中出现 {content.count(old)} 次，"
                f"预期 1 次"
            )
        content = content.replace(old, new)

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[OK] {os.path.relpath(path, BASE_DIR)}: "
          f"已应用 {len(replacements)} 处替换")


# ============================================================================
# 新测试文件
# ============================================================================
TEST_CONTENT = r'''# tests/test_calc_no_double_normalize.py
"""P0-2 回归测试：calc 模块不应重复 normalize。

覆盖：
- already_normalized=True 时，processor 不调用 normalize_excel_file
- already_normalized=False（默认）时，processor 调用 normalize_excel_file
- already_normalized=True 时，原路径不会进入 temp_files（避免误删）
- already_normalized=True 且路径不是 .xlsx 时，回退到 normalize
- 端到端：通过蓝图调用时，normalize_excel_file 仅被调用一次
"""
import os

import pytest
from openpyxl import Workbook

import services.excel.calculator_processor as proc_module
from services.excel.calculator_processor import (
    ColumnCalculatorProcessor,
    AggregateBackfillProcessor,
)
from services.excel.compute_engine import ComputeEngine


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------
def _make_xlsx(path, n_rows: int = 3, n_cols: int = 3):
    wb = Workbook()
    ws = wb.active
    ws.append([f"c{c+1}" for c in range(n_cols)])
    for r in range(1, n_rows):
        ws.append([r * (c + 1) for c in range(n_cols)])
    wb.save(str(path))
    wb.close()
    return str(path)


class _NormalizeCounter:
    """包装 normalize_excel_file，统计调用次数。"""
    def __init__(self, monkeypatch):
        self.count = 0
        self.calls = []
        self._original = proc_module.normalize_excel_file

        def wrapper(*args, **kwargs):
            self.count += 1
            self.calls.append((args, kwargs))
            return self._original(*args, **kwargs)

        monkeypatch.setattr(
            proc_module, "normalize_excel_file", wrapper
        )


# ---------------------------------------------------------------------------
# BaseExcelProcessor._get_work_filepath 的行为
# ---------------------------------------------------------------------------
class TestGetWorkFilepathBehavior:

    def test_already_normalized_skips_normalize(self, tmp_path, monkeypatch):
        """already_normalized=True 时，normalize_excel_file 不应被调用。"""
        src = _make_xlsx(tmp_path / "input.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        result = p._get_work_filepath()

        assert result == src
        assert counter.count == 0, (
            f"already_normalized=True 时不应调用 normalize_excel_file，"
            f"实际调用 {counter.count} 次"
        )

    def test_default_calls_normalize(self, tmp_path, monkeypatch):
        """默认（already_normalized=False）时，应调用 normalize_excel_file。"""
        src = _make_xlsx(tmp_path / "input.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
        )
        result = p._get_work_filepath()

        assert counter.count == 1
        # P0-1 之后有效 .xlsx 会返回原路径
        assert result == src

    def test_already_normalized_path_not_in_temp_files(self, tmp_path):
        """already_normalized=True 时，原路径不应进入 temp_files（避免误删）。"""
        src = _make_xlsx(tmp_path / "input.xlsx")

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        result = p._get_work_filepath()

        assert result == src
        assert src not in p.temp_files, (
            "原路径不应进入 temp_files，否则 cleanup 会误删原文件"
        )

    def test_already_normalized_non_xlsx_falls_back(self, tmp_path, monkeypatch):
        """already_normalized=True 但扩展名不是 .xlsx → 回退到 normalize。"""
        # 故意造一个 .xls 文件
        src = tmp_path / "input.xls"
        wb = Workbook()
        wb.active.append(["A"])
        wb.active.append([1])
        wb.save(str(src))
        wb.close()

        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=str(src),
            header_row=1,
            expression="A+1",
            result_col_input="B",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        p._get_work_filepath()

        assert counter.count == 1, (
            "非 .xlsx 文件即使声明 already_normalized=True，也应回退到 normalize"
        )

    def test_already_normalized_result_is_cached(self, tmp_path, monkeypatch):
        """重复调用 _get_work_filepath 不应重复 normalize。"""
        src = _make_xlsx(tmp_path / "input.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        r1 = p._get_work_filepath()
        r2 = p._get_work_filepath()

        assert r1 == r2 == src
        assert counter.count == 0


# ---------------------------------------------------------------------------
# 两个 Processor 子类都支持 already_normalized
# ---------------------------------------------------------------------------
class TestProcessorSubclasses:

    def test_column_processor_accepts_flag(self, tmp_path, monkeypatch):
        src = _make_xlsx(tmp_path / "c.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src, header_row=1,
            expression="A+B", result_col_input="D",
            merge_by_cols_str="", compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        assert p.already_normalized is True
        p._get_work_filepath()
        assert counter.count == 0

    def test_aggregate_processor_accepts_flag(self, tmp_path, monkeypatch):
        src = _make_xlsx(tmp_path / "a.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = AggregateBackfillProcessor(
            filepath=src, header_row=1,
            group_cols_raw=["A"], agg_cols_raw=["B"],
            agg_func="sum", result_col_input="D",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        assert p.already_normalized is True
        p._get_work_filepath()
        assert counter.count == 0

    def test_default_flag_is_false(self, tmp_path):
        src = _make_xlsx(tmp_path / "d.xlsx")
        p = ColumnCalculatorProcessor(
            filepath=src, header_row=1,
            expression="A+B", result_col_input="D",
            merge_by_cols_str="", compute_engine=ComputeEngine(),
        )
        assert p.already_normalized is False


# ---------------------------------------------------------------------------
# 端到端：通过蓝图触发
# ---------------------------------------------------------------------------
class TestEndToEndViaBlueprint:
    def test_blueprint_triggers_exactly_one_normalize(
        self, client, sample_excel_file, monkeypatch
    ):
        """通过 /calc/process 调用，normalize_excel_file 应恰好被调用 1 次。

        蓝图层调用 1 次；processor 层因 already_normalized=True 应跳过。
        """
        import core.excel_helpers as helpers

        counter = {"n": 0}
        original = helpers.normalize_excel_file

        def counting(*args, **kwargs):
            counter["n"] += 1
            return original(*args, **kwargs)

        monkeypatch.setattr(helpers, "normalize_excel_file", counting)

        data = {
            "file": (sample_excel_file, "test.xlsx"),
            "calc_type": "column",
            "header_row": "1",
            "expression": "A+B",
            "result_col_column": "D",
            "merge_by_cols": "",
        }
        response = client.post(
            "/calc/process", data=data, content_type="multipart/form-data"
        )
        assert response.status_code == 200, (
            f"请求失败: {response.status_code}"
        )
        assert counter["n"] == 1, (
            f"normalize_excel_file 应被调用 1 次（蓝图），实际 {counter['n']} 次"
        )
'''


# ============================================================================
# 写文件
# ============================================================================
def write_file(path: str, content: str) -> None:
    os.makedirs(os.path.dirname(path) or BASE_DIR, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[OK] 已写入 {os.path.relpath(path, BASE_DIR)}  "
          f"({len(content):,} 字符)")


def main() -> None:
    print("=" * 72)
    print("P0-2：消除 calc 模块的重复 normalize")
    print("=" * 72)

    apply_replacements(PROCESSOR_PATH, PROCESSOR_REPLACEMENTS,
                       "calculator_processor.py")
    apply_replacements(CALCULATOR_BP_PATH, CALCULATOR_BP_REPLACEMENTS,
                       "blueprints/calculator.py")
    write_file(TEST_PATH, TEST_CONTENT)

    print()
    print("=" * 72)
    print("✅ 全部文件已写入（未运行任何测试）")
    print("=" * 72)
    print()
    print("核心变更：")
    print("  services/excel/calculator_processor.py")
    print("    · BaseExcelProcessor.__init__ 新增 already_normalized 参数")
    print("    · _get_work_filepath 在该参数为 True 时直接返回原路径")
    print("    · 非 .xlsx 路径 + already_normalized=True → 记警告并回退")
    print("    · ColumnCalculatorProcessor / AggregateBackfillProcessor")
    print("      透传该参数")
    print()
    print("  blueprints/calculator.py")
    print("    · 两处 processor 构造均显式传入 already_normalized=True")
    print()
    print("  tests/test_calc_no_double_normalize.py  (新增)")
    print("    · already_normalized=True → normalize 调用次数为 0")
    print("    · 默认（False）→ normalize 调用次数为 1")
    print("    · 原路径不进 temp_files（避免误删）")
    print("    · 非 .xlsx + already_normalized=True → 回退 normalize")
    print("    · 端到端：/calc/process 请求中 normalize 仅调用 1 次")
    print()
    print("预期收益：")
    print("  · /calc/process 请求中 normalize_excel_file 从 2 次降为 1 次")
    print("  · 免去 1 次 openpyxl 打开 + 1 次尺寸校验")
    print("  · 语义明确：谁 normalize 谁负责清理")
    print("  · 兼容性：独立调用 processor 的行为完全不变")
    print()
    print("下一步（由你手动执行）：")
    print("    python -m pytest tests/test_calc_no_double_normalize.py -v")
    print("    python -m pytest tests/test_excel_helpers.py -v")
    print("    python -m pytest tests/test_integration_calculator.py -v")
    print("    python -m pytest tests/ -v")


if __name__ == "__main__":
    main()