# tests/test_excel_helpers.py
"""
针对 core.excel_helpers.normalize_excel_file 的回归测试。

重点验证：
- 正常文件能够被完整复制（不丢数据）
- 超过 max_rows / max_cols 时抛出 ExcelNormalizationError（而非静默截断）
- ExcelNormalizationError 继承自 ValueError（向后兼容）
- 文件不存在时抛 FileNotFoundError

Windows 注意事项：
- openpyxl 在非 read_only 模式下可能在 Workbook.close() 之后仍短暂
  持有文件句柄。这里统一使用 read_only=True 打开，保证 close() 会
  真正释放底层 zip 归档。
- 清理输出文件时使用带重试的 _robust_unlink，避免偶发的文件锁让
  测试因清理失败而失败（文件在系统 temp 目录，最终会被 OS 回收）。
"""
import gc
import os
import time

import pytest
from openpyxl import Workbook, load_workbook

from core.excel_helpers import (
    normalize_excel_file,
    ExcelNormalizationError,
    MAX_EXCEL_ROWS,
    MAX_EXCEL_COLS,
)


def _make_simple_xlsx(path, n_rows: int = 2, n_cols: int = 2):
    """生成一个简单 Excel 文件，返回其路径。"""
    wb = Workbook()
    ws = wb.active
    # 表头
    ws.append([f"col{c + 1}" for c in range(n_cols)])
    # 数据行
    for r in range(1, n_rows):
        ws.append([r * (c + 1) for c in range(n_cols)])
    wb.save(str(path))
    wb.close()
    return str(path)


def _robust_unlink(path: str, attempts: int = 5, delay: float = 0.1) -> None:
    """
    尽力清理系统临时目录下由 normalize_excel_file 生成的文件。

    Windows 上 openpyxl / zipfile 可能在 close() 后仍短暂持有文件句柄，
    这里做几次带 sleep 的重试；最后仍失败则吞掉异常，交给 OS 清理 /tmp。
    """
    gc.collect()
    for _ in range(attempts):
        if not os.path.exists(path):
            return
        try:
            os.unlink(path)
            return
        except PermissionError:
            gc.collect()
            time.sleep(delay)
    # 最后一次尝试；仍然失败就忽略（系统临时目录会被 OS 回收）
    try:
        os.unlink(path)
    except PermissionError:
        pass


class TestNormalizeExcelFile:
    def test_normal_file_is_copied_intact(self, tmp_path):
        """正常文件应该被完整复制，不丢任何数据。"""
        src = _make_simple_xlsx(tmp_path / "simple.xlsx", n_rows=5, n_cols=3)

        out = normalize_excel_file(src)
        assert os.path.exists(out), "规范化后的文件应该存在"

        try:
            # read_only=True 保证 wb.close() 会真正释放底层文件句柄
            wb_out = load_workbook(out, read_only=True)
            try:
                ws_out = wb_out.active
                # read_only 模式下 max_row 可能不准，直接迭代读取
                rows = list(ws_out.iter_rows(values_only=True))
                # 5 行数据（含表头）
                assert len(rows) == 5, f"期望 5 行，实际 {len(rows)}"
                # 3 列
                assert len(rows[0]) == 3, f"期望 3 列，实际 {len(rows[0])}"
                # 抽查内容
                assert rows[0] == ('col1', 'col2', 'col3')
                # 第 4 行：r=3, [3*(c+1)] = (3, 6, 9)
                assert rows[3] == (3, 6, 9)
            finally:
                wb_out.close()
        finally:
            _robust_unlink(out)

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            normalize_excel_file("/nonexistent/path/to/file.xlsx")

    def test_exceed_rows_raises_instead_of_truncating(self, tmp_path):
        """
        关键回归：文件行数超过 max_rows 时必须抛异常，而不是静默截断。
        这里用 max_rows=5 来模拟上限，实际文件有 1+10=11 行。
        """
        src = _make_simple_xlsx(tmp_path / "many_rows.xlsx", n_rows=11, n_cols=2)

        with pytest.raises(ExcelNormalizationError, match="行数"):
            normalize_excel_file(src, max_rows=5, max_cols=100)

    def test_exceed_cols_raises_instead_of_truncating(self, tmp_path):
        """关键回归：文件列数超过 max_cols 时必须抛异常。"""
        src = _make_simple_xlsx(tmp_path / "wide.xlsx", n_rows=2, n_cols=5)

        with pytest.raises(ExcelNormalizationError, match="列数"):
            normalize_excel_file(src, max_rows=100, max_cols=3)

    def test_error_is_value_error_subclass(self):
        """向后兼容：ExcelNormalizationError 必须是 ValueError 的子类。"""
        assert issubclass(ExcelNormalizationError, ValueError)

    def test_default_limits_are_excel_hard_limits(self):
        """默认上限应是 Excel 硬限制，而不是旧版静默截断的 100000 / 100。"""
        assert MAX_EXCEL_ROWS == 1_048_576
        assert MAX_EXCEL_COLS == 16_384

    def test_below_limit_succeeds(self, tmp_path):
        """刚好等于上限时不应报错（边界条件）。"""
        # 3 行文件（含表头），max_rows=3 刚好容纳
        src = _make_simple_xlsx(tmp_path / "boundary.xlsx", n_rows=3, n_cols=2)

        out = normalize_excel_file(src, max_rows=3, max_cols=2)
        try:
            assert os.path.exists(out)
            wb = load_workbook(out, read_only=True)
            try:
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
                assert len(rows) == 3
                assert len(rows[0]) == 2
            finally:
                wb.close()
        finally:
            _robust_unlink(out)

    def test_no_temp_file_left_on_size_error(self, tmp_path):
        """
        触发尺寸上限时，不应在临时目录留下半成品 normalized_*.xlsx 文件。
        """
        import tempfile
        src = _make_simple_xlsx(tmp_path / "big.xlsx", n_rows=20, n_cols=2)

        before = set(os.listdir(tempfile.gettempdir()))
        with pytest.raises(ExcelNormalizationError):
            normalize_excel_file(src, max_rows=5, max_cols=100)
        after = set(os.listdir(tempfile.gettempdir()))

        leaked = [
            f for f in (after - before)
            if f.startswith("normalized_") and f.endswith(".xlsx")
        ]
        assert leaked == [], f"发现残留的临时文件: {leaked}"
