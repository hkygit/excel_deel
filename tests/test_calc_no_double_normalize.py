# tests/test_calc_no_double_normalize.py
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
