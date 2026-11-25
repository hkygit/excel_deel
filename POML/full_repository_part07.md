- **tests/test_core_task.py**
  - language: `python`
  - content:
    ```python
# tests/test_core_task.py
"""TaskStorage / TaskManager 测试。

覆盖：
- 基础 CRUD（save / load / update / delete）
- 未知字段的 extra 往返与合并
- 任务日志隔离
- WAL 模式已启用
- 崩溃恢复（recover_orphan_tasks）
- 多线程并发更新不丢数据
- 旧 JSON 文件自动迁移
- TaskManager 状态机
"""
import json
import os
import sqlite3
import threading

import pytest

from core.task.storage import TaskStorage
from core.task.manager import TaskManager
from core.task.state import TaskStatus
from core.task.base import Task


class DummyTask(Task):
    def __init__(self, should_fail=False, result_data=None):
        self.should_fail = should_fail
        self.result_data = result_data or {"output": "success"}

    def execute(self):
        if self.should_fail:
            raise ValueError("模拟任务失败")
        return self.result_data


# ============================================================================
# TaskStorage
# ============================================================================
class TestTaskStorage:
    @pytest.fixture
    def storage(self, tmp_path):
        status_dir = tmp_path / "status"
        output_dir = tmp_path / "output"
        return TaskStorage(str(status_dir), str(output_dir))

    # ---------- 基础 CRUD ----------
    def test_save_and_load(self, storage):
        task_id = "test-123"
        data = {"status": "pending", "value": 42}
        storage.save(task_id, data)
        loaded = storage.load(task_id)
        assert loaded["value"] == 42
        assert loaded["status"] == "pending"
        assert loaded["task_id"] == task_id

    def test_update(self, storage):
        task_id = "test-123"
        storage.save(task_id, {"status": "pending"})
        storage.update(task_id, status="completed", result="ok")
        loaded = storage.load(task_id)
        assert loaded["status"] == "completed"
        assert loaded["result"] == "ok"

    def test_delete(self, storage):
        task_id = "delete-me"
        storage.save(task_id, {})
        assert storage.load(task_id) is not None
        storage.delete(task_id)
        assert storage.load(task_id) is None

    def test_update_nonexistent_raises(self, storage):
        with pytest.raises(ValueError, match="not found"):
            storage.update("does-not-exist", status="completed")

    # ---------- extra 字段 ----------
    def test_extra_fields_roundtrip(self, storage):
        task_id = "extra-test"
        storage.save(
            task_id,
            {"status": "pending", "custom_a": 1, "custom_b": "x"},
        )
        loaded = storage.load(task_id)
        assert loaded["custom_a"] == 1
        assert loaded["custom_b"] == "x"

    def test_extra_field_merges_on_update(self, storage):
        task_id = "merge-test"
        storage.save(task_id, {"status": "pending", "custom_a": 1})
        storage.update(task_id, custom_b=2)
        loaded = storage.load(task_id)
        assert loaded["custom_a"] == 1
        assert loaded["custom_b"] == 2

    def test_params_dict_roundtrip(self, storage):
        """params 字典应通过 JSON 往返无损。"""
        task_id = "params-test"
        params = {"file": "x.csv", "options": {"sep": ",", "enc": "utf-8"}}
        storage.save(task_id, {"status": "pending", "params": params})
        loaded = storage.load(task_id)
        assert loaded["params"] == params

    # ---------- 日志 ----------
    def test_append_and_get_log(self, storage):
        task_id = "log-test"
        storage.append_log(task_id, "first line")
        storage.append_log(task_id, "second line")
        log = storage.get_log(task_id)
        assert "first line" in log
        assert "second line" in log
        # 顺序应保持
        assert log.index("first line") < log.index("second line")

    def test_logs_isolated_by_task(self, storage):
        storage.append_log("task-1", "hello 1")
        storage.append_log("task-2", "hello 2")
        assert "hello 1" in storage.get_log("task-1")
        assert "hello 2" not in storage.get_log("task-1")

    def test_get_log_limit(self, storage):
        for i in range(10):
            storage.append_log("many", f"line-{i}")
        log = storage.get_log("many", max_lines=3)
        # 应该只保留最后 3 行
        assert "line-7" in log
        assert "line-8" in log
        assert "line-9" in log
        assert "line-0" not in log

    # ---------- SQLite 特性 ----------
    def test_wal_mode_enabled(self, storage):
        conn = sqlite3.connect(storage.db_path)
        try:
            mode = conn.execute("PRAGMA journal_mode").fetchone()[0]
            assert mode.lower() == "wal"
        finally:
            conn.close()

    def test_db_file_created(self, storage):
        assert os.path.exists(storage.db_path)
        assert storage.db_path.endswith("tasks.db")

    # ---------- 崩溃恢复 ----------
    def test_recover_orphan_tasks(self, storage):
        storage.save("orphan-1", {"status": "running"})
        storage.save("orphan-2", {"status": "pending"})
        storage.save("done-1", {"status": "completed"})
        storage.save("failed-1", {"status": "failed"})

        count = storage.recover_orphan_tasks()
        assert count == 2

        assert storage.load("orphan-1")["status"] == "failed"
        assert storage.load("orphan-2")["status"] == "failed"
        assert storage.load("done-1")["status"] == "completed"
        assert storage.load("failed-1")["status"] == "failed"
        # 错误信息应该保留提示
        assert "服务重启" in storage.load("orphan-1")["error"]

    def test_recover_orphan_tasks_empty(self, storage):
        """没有 pending/running 时返回 0。"""
        storage.save("x", {"status": "completed"})
        assert storage.recover_orphan_tasks() == 0

    # ---------- 并发 ----------
    def test_concurrent_updates_do_not_lose_data(self, tmp_path):
        """多线程同时更新同一 task 的不同字段，字段不应丢失。"""
        storage = TaskStorage(
            str(tmp_path / "s"), str(tmp_path / "o")
        )
        storage.save("conc", {"status": "running"})

        def worker(i):
            storage.update("conc", **{f"field_{i}": i})

        threads = [
            threading.Thread(target=worker, args=(i,)) for i in range(20)
        ]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        loaded = storage.load("conc")
        for i in range(20):
            assert loaded.get(f"field_{i}") == i, f"field_{i} lost"

    # ---------- 旧 JSON 迁移 ----------
    def test_json_migration(self, tmp_path):
        """旧的 .json 任务文件应被迁移到 SQLite。"""
        status_dir = tmp_path / "status"
        status_dir.mkdir()
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        legacy = {
            "task_id": "legacy-1",
            "status": "completed",
            "created_at": "2026-01-01T00:00:00",
            "result_path": "/tmp/x.xlsx",
        }
        (status_dir / "legacy-1.json").write_text(
            json.dumps(legacy, ensure_ascii=False), encoding="utf-8"
        )
        (status_dir / "legacy-1.log").write_text(
            "[2026-01-01 00:00:01] hello\n"
            "[2026-01-01 00:00:02] world\n",
            encoding="utf-8",
        )

        # 初始化时自动迁移
        storage = TaskStorage(str(status_dir), str(output_dir))

        loaded = storage.load("legacy-1")
        assert loaded is not None
        assert loaded["status"] == "completed"
        assert loaded["result_path"] == "/tmp/x.xlsx"

        log = storage.get_log("legacy-1")
        assert "hello" in log
        assert "world" in log

        # 旧文件应被移动到 .migrated/
        migrated = status_dir / ".migrated"
        assert migrated.is_dir()
        assert (migrated / "legacy-1.json").exists()

    def test_json_migration_skips_corrupt(self, tmp_path):
        """损坏的 JSON 文件不应阻塞其他文件的迁移。"""
        status_dir = tmp_path / "status"
        status_dir.mkdir()
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        (status_dir / "corrupt.json").write_text(
            "{not valid json", encoding="utf-8"
        )
        (status_dir / "good.json").write_text(
            json.dumps({"task_id": "good", "status": "completed"}),
            encoding="utf-8",
        )

        storage = TaskStorage(str(status_dir), str(output_dir))
        assert storage.load("good") is not None
        # 损坏的文件保持原样（不能被识别为有效任务）
        assert storage.load("corrupt") is None


# ============================================================================
# TaskManager（状态机）
# ============================================================================
class TestTaskManager:
    @pytest.fixture
    def storage(self, tmp_path):
        status_dir = tmp_path / "status"
        output_dir = tmp_path / "output"
        return TaskStorage(str(status_dir), str(output_dir))

    @pytest.fixture
    def manager(self, storage):
        return TaskManager(storage)

    def test_create_task(self, manager):
        params = {"file": "test.csv"}
        task_id = manager.create_task(params, original_filename="test")
        assert task_id is not None
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.PENDING
        assert status["params"] == params

    def test_submit_and_complete(self, manager):
        task_id = manager.create_task({})
        manager.storage.update(task_id, status=TaskStatus.RUNNING)
        manager.complete_task(
            task_id, result_path="/out.xlsx", total_input_rows=100
        )
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.COMPLETED
        assert status["result_path"] == "/out.xlsx"

    def test_fail_task(self, manager):
        task_id = manager.create_task({})
        manager.fail_task(task_id, error="Something wrong")
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.FAILED
        assert status["error"] == "Something wrong"

    def test_cancel_task(self, manager):
        task_id = manager.create_task({})
        manager.cancel_task(task_id)
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.CANCELLED

    def test_manager_with_recover(self, manager, storage):
        """模拟服务重启后，遗留的 running 任务应被恢复。"""
        task_id = manager.create_task({})
        manager.storage.update(task_id, status=TaskStatus.RUNNING)
        # 模拟重启
        count = storage.recover_orphan_tasks()
        assert count == 1
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.FAILED
    ```

- **tests/test_data_access_layer.py**
  - language: `python`
  - content:
    ```python
import pytest
import pandas as pd
from core.data_access_layer import DataAccessLayer

class TestDataAccessLayer:
    def test_auto_clean_strings(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'A': [' 10 ', '20', ' 30 ']})
        cleaned = dal._auto_clean_dataframe(df)
        # 数值转换后应为数字，但测试允许字符串或数字（向后兼容）
        assert cleaned['A'].tolist() == [10.0, 20.0, 30.0] or cleaned['A'].tolist() == ['10', '20', '30']

    def test_dtype_hints(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'amount': ['100', '200', 'invalid']})
        cleaned = dal._auto_clean_dataframe(df, dtype_hints={'amount': float})
        # dtype_hints 会强制转换，无效值变为 NaN
        assert cleaned['amount'].isna().sum() == 1

    def test_empty_string_handling(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'A': ['', '   ', '10']})
        cleaned = dal._auto_clean_dataframe(df)
        # 空字符串变为 NaN，'10' 转为数字或保持字符串
        assert pd.isna(cleaned['A'].iloc[0])
        assert pd.isna(cleaned['A'].iloc[1])
        assert cleaned['A'].iloc[2] == 10.0 or cleaned['A'].iloc[2] == '10'
    ```

- **tests/test_di_container.py**
  - language: `python`
  - content:
    ```python
# tests/test_di_container.py
"""DI 容器（P1-2）回归测试。

覆盖：
- Container 可以实例化，且默认读取全局 settings
- task_storage 的 status_dir 与 output_dir 分离
- task_storage / task_manager 为 Singleton（同一实例）
- checker_service / team_report_service 为 Factory（不同实例）
- 规范化的子配置 provider 返回正确的配置对象
- 向后兼容：config.provided.xxx 依旧可用
- Container(config=...) 覆盖配置后所有依赖同步更新
- field_parser 的 similarity_threshold 来自 field_parser_config
"""
import os

import pytest
from dependency_injector import providers

from core.config import Settings, settings as global_settings
from core.di import Container, _task_status_dir
from core.task.storage import TaskStorage
from core.task.manager import TaskManager
from services.checker.service import CheckerService
from services.team_report.service import TeamReportService


def _make_settings(tmp_path):
    """构造一个纯净的 Settings（不读 .env），输出目录指向临时目录。"""
    return Settings(
        _env_file=None,
        output_folder=str(tmp_path / "outputs"),
        upload_folder=str(tmp_path / "uploads"),
    )


@pytest.fixture
def container(tmp_path):
    """使用临时目录的 Settings 覆盖 container 的 config provider。"""
    s = _make_settings(tmp_path)
    return Container(config=providers.Singleton(lambda: s))


# ============================================================================
# 基础实例化
# ============================================================================
class TestContainerInstantiation:
    def test_container_can_be_instantiated(self):
        c = Container()
        assert c is not None
        # 不覆盖时，config 返回全局 settings
        assert c.config() is global_settings

    def test_default_config_returns_global_settings(self):
        c = Container()
        assert c.config() is global_settings


# ============================================================================
# task_storage 目录分离（本次修复的核心）
# ============================================================================
class TestTaskStorageDirs:
    def test_status_dir_separated_from_output_dir(self, container, tmp_path):
        """关键回归：status_dir 必须与 output_dir 不同。"""
        storage = container.task_storage()
        output_dir = str(tmp_path / "outputs")
        expected_status_dir = os.path.join(output_dir, "task_status")

        assert storage.output_dir == output_dir
        assert storage.status_dir == expected_status_dir
        assert storage.status_dir != storage.output_dir

    def test_status_dir_created_on_init(self, container, tmp_path):
        """首次访问 task_storage 时，两个目录都应被创建。"""
        storage = container.task_storage()
        assert os.path.isdir(storage.status_dir)
        assert os.path.isdir(storage.output_dir)

    def test_task_status_dir_helper(self, tmp_path):
        """_task_status_dir 的纯函数行为。"""
        s = _make_settings(tmp_path)
        result = _task_status_dir(s)
        assert result == os.path.join(
            str(tmp_path / "outputs"), "task_status"
        )

    def test_db_file_lives_in_status_dir(self, container):
        storage = container.task_storage()
        # tasks.db 应位于 status_dir 下
        assert storage.db_path.startswith(storage.status_dir)
        assert os.path.basename(storage.db_path) == "tasks.db"


# ============================================================================
# Singleton / Factory 语义
# ============================================================================
class TestProviderScopes:
    def test_task_storage_is_singleton(self, container):
        assert container.task_storage() is container.task_storage()

    def test_task_manager_is_singleton(self, container):
        assert container.task_manager() is container.task_manager()

    def test_task_manager_shares_storage(self, container):
        storage = container.task_storage()
        manager = container.task_manager()
        assert manager.storage is storage

    def test_checker_service_is_factory(self, container):
        assert container.checker_service() is not container.checker_service()

    def test_team_report_service_is_factory(self, container):
        assert (
            container.team_report_service()
            is not container.team_report_service()
        )

    def test_field_parser_is_singleton(self, container):
        assert container.field_parser() is container.field_parser()


# ============================================================================
# 规范化子配置 provider
# ============================================================================
class TestSubConfigProviders:
    def test_checker_config(self, container):
        cfg = container.checker_config()
        assert cfg is not None
        assert hasattr(cfg, "similarity_threshold")
        assert hasattr(cfg, "duplicate_action")

    def test_team_report_config(self, container):
        cfg = container.team_report_config()
        assert cfg is not None
        assert cfg.daily_order_target == 1150
        assert cfg.daily_apply_target == 120

    def test_field_parser_config(self, container):
        cfg = container.field_parser_config()
        assert cfg is not None
        assert hasattr(cfg, "similarity_threshold")

    def test_splitter_config(self, container):
        assert container.splitter_config() is not None

    def test_convert_config(self, container):
        assert container.convert_config() is not None

    def test_sub_config_stable_between_calls(self, container):
        """Callable provider 每次返回 settings 上的同一子配置实例。"""
        assert container.checker_config() is container.checker_config()


# ============================================================================
# 向后兼容：config.provided.xxx
# ============================================================================
class TestBackwardCompat:
    """保留 container.config.provided.xxx 的旧调用方式。"""

    def test_config_provided_checker(self, container):
        cfg = container.config.provided.checker()
        assert cfg is not None
        assert hasattr(cfg, "similarity_threshold")

    def test_config_provided_team_report(self, container):
        cfg = container.config.provided.team_report()
        assert cfg is not None

    def test_provided_and_new_provider_same_threshold(self, container):
        """两种访问方式应反映同一份配置。"""
        old = container.config.provided.checker()
        new = container.checker_config()
        assert old.similarity_threshold == new.similarity_threshold
        assert old.duplicate_action == new.duplicate_action


# ============================================================================
# 依赖注入的 wiring
# ============================================================================
class TestWiring:
    def test_checker_service_uses_checker_config(self, tmp_path):
        """checker_service 的 config 应来自 checker_config。"""
        s = _make_settings(tmp_path)
        s.checker.similarity_threshold = 999
        c = Container(config=providers.Singleton(lambda: s))
        svc = c.checker_service()
        assert isinstance(svc, CheckerService)
        assert svc.config.similarity_threshold == 999

    def test_team_report_service_uses_team_report_config(self, tmp_path):
        """team_report_service 的 config 应来自 team_report_config。"""
        s = _make_settings(tmp_path)
        s.team_report.daily_order_target = 777
        c = Container(config=providers.Singleton(lambda: s))
        svc = c.team_report_service()
        assert isinstance(svc, TeamReportService)
        assert svc.config.daily_order_target == 777

    def test_field_parser_uses_field_parser_config(self, tmp_path):
        """field_parser 的 similarity_threshold 来自 field_parser_config。"""
        s = _make_settings(tmp_path)
        s.field_parser.similarity_threshold = 42
        c = Container(config=providers.Singleton(lambda: s))
        parser = c.field_parser()
        # FieldParserService.__init__ 会做 / 100
        assert parser.threshold == 0.42

    def test_override_config_propagates(self, tmp_path):
        """覆盖 config 后，所有子 provider 应反映新配置。"""
        s = _make_settings(tmp_path)
        s.checker.duplicate_action = "keep_first"
        c = Container(config=providers.Singleton(lambda: s))

        assert c.checker_config().duplicate_action == "keep_first"
        assert c.config.provided.checker().duplicate_action == "keep_first"
    ```

- **tests/test_excel_helpers.py**
  - language: `python`
  - content:
    ```python
# tests/test_excel_helpers.py
"""
针对 core.excel_helpers.normalize_excel_file 的回归测试。

P0 语义：
- 有效 .xlsx → 直接返回原路径（零复制）
- .xls / 损坏 .xlsx → 走转换路径，返回新的临时文件
- 尺寸超限 → 抛 ExcelNormalizationError
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
    ws.append([f"col{c + 1}" for c in range(n_cols)])
    for r in range(1, n_rows):
        ws.append([r * (c + 1) for c in range(n_cols)])
    wb.save(str(path))
    wb.close()
    return str(path)


def _robust_unlink(path: str, attempts: int = 5, delay: float = 0.1) -> None:
    """尽力清理系统临时目录下由 normalize_excel_file 生成的文件。

    Windows 上 openpyxl / zipfile 可能在 close() 后仍短暂持有文件句柄，
    因此做几次带 sleep 的重试；最后仍失败则吞掉异常。
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
    try:
        os.unlink(path)
    except PermissionError:
        pass


def _temp_normalized_files() -> set:
    """返回系统临时目录下所有 normalized_*.xlsx 文件名。"""
    import tempfile
    return {
        f for f in os.listdir(tempfile.gettempdir())
        if f.startswith("normalized_") and f.endswith(".xlsx")
    }


class TestNormalizeExcelFile:
    # ---------- P0 核心：有效 .xlsx 零复制 ----------
    def test_valid_xlsx_returns_original_path_and_is_intact(self, tmp_path):
        """P0 核心：有效 .xlsx 直接返回原路径，且内容不变。"""
        src = _make_simple_xlsx(tmp_path / "simple.xlsx", n_rows=5, n_cols=3)

        out = normalize_excel_file(src)

        # 关键断言：返回的应该就是原路径
        assert out == src, f"应返回原路径 {src}，实际 {out}"
        assert os.path.exists(out)

        # 内容应完整
        wb = load_workbook(out, read_only=True)
        try:
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            assert len(rows) == 5
            assert rows[0] == ('col1', 'col2', 'col3')
            assert rows[3] == (3, 6, 9)
        finally:
            wb.close()
        # 注意：out 就是 src，由 tmp_path fixture 自动清理，无需 unlink

    def test_valid_xlsx_does_not_create_temp_file(self, tmp_path):
        """P0 核心：有效 .xlsx 不创建任何 normalized_*.xlsx 临时文件。"""
        src = _make_simple_xlsx(tmp_path / "no_copy.xlsx", n_rows=3, n_cols=2)
        before = _temp_normalized_files()

        out = normalize_excel_file(src)

        after = _temp_normalized_files()
        new_files = after - before
        assert new_files == set(), f"不应创建临时文件: {new_files}"
        assert out == src

    def test_multiple_calls_no_temp_file_accumulation(self, tmp_path):
        """连续多次调用，每次都返回原路径，不累积临时文件。"""
        src = _make_simple_xlsx(tmp_path / "multi.xlsx", n_rows=3, n_cols=2)
        before = _temp_normalized_files()

        results = [normalize_excel_file(src) for _ in range(3)]

        after = _temp_normalized_files()
        new_files = after - before
        assert new_files == set(), f"多次调用不应产生临时文件: {new_files}"
        assert all(r == src for r in results)

    def test_below_limit_returns_original_path(self, tmp_path):
        """恰好等于上限的 .xlsx 仍应返回原路径。"""
        src = _make_simple_xlsx(tmp_path / "boundary.xlsx", n_rows=3, n_cols=2)

        out = normalize_excel_file(src, max_rows=3, max_cols=2)
        assert out == src

        wb = load_workbook(out, read_only=True)
        try:
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            assert len(rows) == 3
            assert len(rows[0]) == 2
        finally:
            wb.close()

    # ---------- 损坏文件仍走转换路径 ----------
    def test_corrupt_xlsx_does_not_return_original_path(self, tmp_path):
        """有 magic 但内部损坏的 .xlsx 不应直接返回原路径。"""
        bad = tmp_path / "corrupt.xlsx"
        # 有 ZIP magic，但内容是垃圾字节
        bad.write_bytes(b"PK\x03\x04" + b"\x00" * 200)

        out = normalize_excel_file(str(bad))

        # 关键断言：不能返回原路径（因为文件无效）
        assert out != str(bad), "损坏文件不应直接返回原路径"
        assert os.path.exists(out), "应返回一个可用的文件路径"
        _robust_unlink(out)

    # ---------- 基础异常 ----------
    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            normalize_excel_file("/nonexistent/path/to/file.xlsx")

    # ---------- 尺寸上限 ----------
    def test_exceed_rows_raises_instead_of_truncating(self, tmp_path):
        """行数超限时抛 ExcelNormalizationError。"""
        src = _make_simple_xlsx(tmp_path / "many_rows.xlsx", n_rows=11, n_cols=2)

        with pytest.raises(ExcelNormalizationError, match="行数"):
            normalize_excel_file(src, max_rows=5, max_cols=100)

    def test_exceed_cols_raises_instead_of_truncating(self, tmp_path):
        """列数超限时抛 ExcelNormalizationError。"""
        src = _make_simple_xlsx(tmp_path / "wide.xlsx", n_rows=2, n_cols=5)

        with pytest.raises(ExcelNormalizationError, match="列数"):
            normalize_excel_file(src, max_rows=100, max_cols=3)

    def test_no_temp_file_left_on_size_error(self, tmp_path):
        """尺寸超限时不应留下任何 normalized_*.xlsx 临时文件。"""
        src = _make_simple_xlsx(tmp_path / "big.xlsx", n_rows=20, n_cols=2)
        before = _temp_normalized_files()

        with pytest.raises(ExcelNormalizationError):
            normalize_excel_file(src, max_rows=5, max_cols=100)

        after = _temp_normalized_files()
        leaked = after - before
        assert leaked == set(), f"发现残留的临时文件: {leaked}"

    # ---------- 常量与异常类型 ----------
    def test_error_is_value_error_subclass(self):
        """向后兼容：ExcelNormalizationError 必须是 ValueError 的子类。"""
        assert issubclass(ExcelNormalizationError, ValueError)

    def test_default_limits_are_excel_hard_limits(self):
        """默认上限应是 Excel 硬限制，而不是旧版静默截断的 100000 / 100。"""
        assert MAX_EXCEL_ROWS == 1_048_576
        assert MAX_EXCEL_COLS == 16_384
    ```

- **tests/test_expression_engine.py**
  - language: `python`
  - content:
    ```python
import pytest
import numpy as np
from core.expression_engine import ExpressionEngine

class TestExpressionEngine:
    @pytest.fixture
    def engine(self):
        return ExpressionEngine()

    def test_basic_arithmetic(self, engine):
        assert engine.evaluate("10+20", {}) == 30
        assert engine.evaluate("10-5", {}) == 5
        assert engine.evaluate("3*4", {}) == 12
        assert engine.evaluate("10/2", {}) == 5

    def test_with_variables(self, engine):
        context = {"A": 10, "B": 5}
        assert engine.evaluate("A+B", context) == 15
        assert engine.evaluate("A*B", context) == 50

    def test_string_conversion(self, engine):
        context = {"A": "10", "B": " 5 "}
        assert engine.evaluate("A+B", context) == 15

    def test_missing_variable_nan(self, engine):
        engine.missing_value_strategy = 'nan'
        assert np.isnan(engine.evaluate("X+Y", {"X": 10}))

    def test_missing_variable_zero(self):
        engine = ExpressionEngine(missing_value_strategy='zero')
        assert engine.evaluate("X+Y", {"X": 10}) == 10.0

    def test_division_by_zero(self, engine):
        assert np.isnan(engine.evaluate("10/0", {}))

    def test_expression_caching(self, engine):
        expr = "A+B*C"
        context = {"A": 1, "B": 2, "C": 3}
        assert engine.evaluate(expr, context) == 7
        assert expr.upper() in engine._cache

    def test_invalid_expression(self, engine):
        with pytest.raises(ValueError):
            engine.evaluate("10+", {})
    ```

- **tests/test_field_parser.py**
  - language: `python`
  - content:
    ```python
import pytest
from services.field_parser.service import FieldParserService

class TestFieldParserService:
    @pytest.fixture
    def service(self):
        return FieldParserService(similarity_threshold=30)

    def test_parse_field_list_commas(self, service):
        result = service.parse_field_list("姓名, 年龄, 性别")
        assert result == ["姓名", "年龄", "性别"]

    def test_match_exact(self, service):
        result = service.match_headers(["ID", "Name"], ["ID", "Name"])
        assert len(result["matched"]) == 2
        assert result["only_a"] == []
        assert result["only_b"] == []

    def test_match_similar_chinese(self, service):
        result = service.match_headers(["姓名", "年龄"], ["名字", "岁数"])
        assert len(result["matched"]) == 1
        assert result["matched"][0]["standard"] == "姓名"
        assert result["matched"][0]["compare"] == "名字"
        assert len(result["only_a"]) == 1
        assert "年龄" in result["only_a"]
        assert len(result["only_b"]) == 1
        assert "岁数" in result["only_b"]

    def test_match_below_threshold(self, service):
        service.threshold = 0.9
        result = service.match_headers(["姓名"], ["名字"])
        assert len(result["matched"]) == 0
    ```

- **tests/test_file_naming.py**
  - language: `python`
  - content:
    ```python
# tests/test_file_naming.py
from core.file_naming import generate_output_filename

class TestFileNaming:
    def test_basic(self):
        name = generate_output_filename('subtotal', '数据.xlsx')
        assert name == 'subtotal_数据_小计.xlsx'

    def test_with_path(self):
        name = generate_output_filename('match', '/path/to/目标文件.xlsx')
        assert name == 'match_目标文件_小计.xlsx'

    def test_custom_suffix(self):
        name = generate_output_filename('filter', 'data.csv', suffix='_汇总')
        assert name == 'filter_data_汇总.xlsx'
    ```

- **tests/test_filter.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
import pytest
from core.filter_engine import parse_conditions_from_form, ConditionEngine
from core.column_resolver import resolve_columns
from core.exceptions import ValidationError

engine = ConditionEngine()

class TestParseFiltersFromForm:
    def test_empty(self):
        form = {}
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact == []
        assert fuzzy == []

    def test_single_equals(self):
        form = {
            'condition_field_0': 'name',
            'condition_op_0': '=',
            'condition_value_0': 'Alice'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(exact) == 1
        assert exact[0]['field'] == 'name'
        assert exact[0]['op'] == '='
        assert exact[0]['values'] == ['Alice']

    def test_multi_values_space(self):
        form = {
            'condition_field_0': 'dept',
            'condition_op_0': '=',
            'condition_value_0': 'HR  IT'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact[0]['values'] == ['HR', 'IT']

    def test_not_equal(self):
        form = {
            'condition_field_0': 'score',
            'condition_op_0': '!=',
            'condition_value_0': '100'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact[0]['op'] == '!='

    def test_contains(self):
        form = {
            'condition_field_0': 'name',
            'condition_op_0': '包含',
            'condition_value_0': 'Al'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(fuzzy) == 1
        assert fuzzy[0]['op'] == '包含'

    def test_compatibility_old_format(self):
        form = {
            'condition_field_0': 'dept',
            'condition_values_0': 'HR IT'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(exact) == 1
        assert exact[0]['values'] == ['HR', 'IT']
        assert exact[0]['op'] == '='

class TestApplyConditions:
    def test_equals(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie']})
        exact = [{'field': 'name', 'op': '=', 'values': ['Alice', 'Charlie']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['name']) == ['Alice', 'Charlie']

    def test_not_equals(self):
        df = pd.DataFrame({'score': [100, 90, 80]})
        exact = [{'field': 'score', 'op': '!=', 'values': ['90']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['score']) == [100, 80]

    def test_greater_than(self):
        df = pd.DataFrame({'age': [25, 30, 35]})
        exact = [{'field': 'age', 'op': '>', 'values': ['28']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['age']) == [30, 35]

    def test_contains(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Alfred']})
        fuzzy = [{'field': 'name', 'op': '包含', 'value': 'Al'}]
        result = engine.apply_to_dataframe(df, [], fuzzy)
        assert list(result['name']) == ['Alice', 'Alfred']

    def test_not_contains(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Alfred']})
        fuzzy = [{'field': 'name', 'op': '不包含', 'value': 'Al'}]
        result = engine.apply_to_dataframe(df, [], fuzzy)
        assert list(result['name']) == ['Bob']

    def test_combined(self):
        df = pd.DataFrame({
            'dept': ['HR', 'IT', 'HR', 'IT'],
            'salary': [5000, 6000, 5500, 7000],
            'name': ['Alice', 'Bob', 'Charlie', 'David']
        })
        exact = [{'field': 'dept', 'op': '=', 'values': ['HR']}]
        fuzzy = [{'field': 'name', 'op': '包含', 'value': 'li'}]
        result = engine.apply_to_dataframe(df, exact, fuzzy)
        assert len(result) == 2
        assert set(result['name']) == {'Alice', 'Charlie'}

class TestResolveColumns:
    def test_by_name(self):
        df = pd.DataFrame({'A': [1], 'B': [2], 'C': [3]})
        res = resolve_columns(df, 'A,C')
        assert list(res.columns) == ['A', 'C']

    def test_by_letter(self):
        df = pd.DataFrame({'X': [1], 'Y': [2], 'Z': [3]})
        res = resolve_columns(df, 'A,C')
        assert list(res.columns) == ['X', 'Z']

    def test_mixed(self):
        df = pd.DataFrame({'Name': ['a'], 'Age': [1], 'Score': [2]})
        res = resolve_columns(df, 'Name,B')
        assert list(res.columns) == ['Name', 'Age']

    def test_invalid_raises(self):
        df = pd.DataFrame({'X': [1]})
        with pytest.raises(ValidationError):
            resolve_columns(df, 'NonExist')
    ```

- **tests/test_header_detection.py**
  - language: `python`
  - content:
    ```python
# tests/test_header_detection.py
import pytest
import pandas as pd
from core.header_detection import SimpleExcelHeaderDetector, detect_header_start, HeaderDetectionConfig
import openpyxl

class TestHeaderDetection:
    @pytest.fixture
    def simple_file(self, tmp_path):
        df = pd.DataFrame({'姓名': ['Alice', 'Bob'], '年龄': [30, 25]})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        return str(path)

    def test_basic_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(simple_file)
        assert result.header_rows == [0]
        assert result.data_start_row == 1
        assert '姓名' in result.column_names
        assert result.confidence > 0.5

    def test_detect_header_start(self, simple_file):
        start = detect_header_start(simple_file)
        assert start == 2

    def test_skip_total_row(self, tmp_path):
        df = pd.DataFrame({'名称': ['X', '合计'], '金额': [10, 30]})
        path = tmp_path / 'skip.xlsx'
        df.to_excel(path, index=False)
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(path)
        assert result.header_rows == [0]

    def test_workbook_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        wb = openpyxl.load_workbook(simple_file)
        result1 = detector.detect_from_workbook(wb)
        result2 = detector.detect_from_workbook(wb)
        assert result1.header_rows == [0]
        assert result2.metrics.get('workbook_cache_hit') is True
        wb.close()
    ```

- **tests/test_integration_calculator.py**
  - language: `python`
  - content:
    ```python
import io
from openpyxl import load_workbook


class TestCalculatorIntegration:
    def test_column_calculation(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'column',
            'header_row': '1',
            'expression': 'A+B',
            'result_col_column': 'D',
            'merge_by_cols': ''
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        assert response.headers['Content-Disposition'].startswith('attachment')
        assert '.xlsx' in response.headers['Content-Disposition']
        wb = load_workbook(io.BytesIO(response.data))
        ws = wb.active
        assert ws['D2'].value == 30
        assert ws['D3'].value == 8

    def test_aggregate_backfill(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'aggregate',
            'header_row': '1',
            'group_cols': 'A',
            'agg_cols': 'B',
            'agg_func': 'sum',
            'result_col_aggregate': 'E'
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        assert '.xlsx' in response.headers['Content-Disposition']
        wb = load_workbook(io.BytesIO(response.data))
        ws = wb.active
        assert ws['E2'].value == 20
        assert ws['E3'].value == 3

    def test_missing_file_error(self, client):
        response = client.post('/calc/process', data={'calc_type': 'column'})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '未上传文件' in response.get_json()['error']

    def test_invalid_expression_error(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'column',
            'header_row': '1',
            'expression': '',
            'result_col_column': 'D'
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '表达式' in response.get_json()['error']
    ```

- **tests/test_integration_checker.py**
  - language: `python`
  - content:
    ```python

import io
import json
import pandas as pd
import pytest
from unittest.mock import patch

class TestCheckerIntegration:
    def test_index_page(self, client):
        response = client.get('/checker/')
        assert response.status_code == 200
        assert '文件校验' in response.get_data(as_text=True)

    @pytest.fixture
    def sample_files(self, tmp_path):
        df_a1 = pd.DataFrame({"ID": [1, 2], "Name": ["Alice", "Bob"], "Score": [100, 90]})
        df_a2 = pd.DataFrame({"ID": [3, 4], "Name": ["Charlie", "David"], "Score": [80, 70]})
        df_b = pd.DataFrame({"ID": [1, 2, 5], "Name": ["Alice", "Bob", "Eve"], "Score": [100, 95, 60]})
        return {'a1': df_a1, 'a2': df_a2, 'b': df_b}

    def test_full_check_async(self, client, sample_files):
        def df_to_bytes(df, filename):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w:
                df.to_excel(w, index=False)
            buf.seek(0)
            return (buf, filename)

        files_a = [df_to_bytes(sample_files['a1'], 'a1.xlsx'), df_to_bytes(sample_files['a2'], 'a2.xlsx')]
        files_b = [df_to_bytes(sample_files['b'], 'b.xlsx')]

        data = {
            'key_map_a': 'ID',
            'key_map_b': '',
            'compare_map': json.dumps([{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}]),
            'filters_a': '[]',
            'duplicate_action': 'error',
            'files_a': files_a,
            'files_b': files_b,
        }

        # Mock TaskManager.start_task 使其同步执行目标函数
        app = client.application
        original_tm = app.container.task_manager()
        def sync_start(task_id, target, args):
            target(*args)  # 直接调用，不启动线程
        with patch.object(original_tm, 'start_task', side_effect=sync_start):
            response = client.post('/checker/process', data=data, content_type='multipart/form-data')
            assert response.status_code == 200
            result = response.get_json()
            assert result['success'] == True  # 现在返回统一格式
            task_id = result['data']['task_id']

            # 任务应已同步完成，直接检查状态
            status_resp = client.get(f'/checker/status/{task_id}')
            status = status_resp.get_json()['data']
            assert status['status'] == 'completed'

            # 下载结果文件
            download_resp = client.get(f'/checker/download/{task_id}')
            assert download_resp.status_code == 200
            result_df = pd.read_excel(io.BytesIO(download_resp.data), sheet_name=None)
            assert '差异数据' in result_df
            assert len(result_df['差异数据']) == 1
            assert '标准群独有' in result_df
            assert len(result_df['标准群独有']) == 2
            assert '比较群独有' in result_df
            assert len(result_df['比较群独有']) == 1

    def test_missing_files_error(self, client):
        response = client.post('/checker/process', data={})
        assert response.status_code == 400
        json_data = response.get_json()
        assert not json_data['success']
        assert '请上传两批文件' in json_data['error']

    def test_invalid_task_status(self, client):
        response = client.get('/checker/status/nonexistent')
        assert response.status_code == 404
        assert not response.get_json()['success']
    ```

- **tests/test_integration_comparator.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_comparator.py
import io
import json
import pandas as pd

class TestComparatorIntegration:
    def test_full_compare(self, client, sample_excel_comparison_files):
        f1, f2 = sample_excel_comparison_files
        data = {
            'file1': (f1, 'file1.xlsx'),
            'file2': (f2, 'file2.xlsx'),
            'compare_mode': 'full',
            'key_fields': 'id',
            'compare_fields': 'name,score',
            'sheet1': '',
            'sheet2': ''
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content_disp = response.headers['Content-Disposition']
        # 新文件名格式：compare_file1_对比结果.xlsx
        assert 'compare_file1_' in content_disp and '.xlsx' in content_disp
        result = pd.ExcelFile(io.BytesIO(response.data))
        sheet_names = result.sheet_names
        assert '仅存在于文件2' in sheet_names

    def test_field_match_mode(self, client, sample_excel_comparison_files):
        f1, f2 = sample_excel_comparison_files
        data = {
            'file1': (f1, 'file1.xlsx'),
            'file2': (f2, 'file2.xlsx'),
            'compare_mode': 'field_match',
            'match_field': 'name',
            'sheet1': '',
            'sheet2': ''
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content_disp = response.headers.get('Content-Disposition', '')
        assert 'filename=' in content_disp

    def test_no_difference_response(self, client):
        import pandas as pd
        from io import BytesIO
        df = pd.DataFrame({'id': [1], 'name': ['A']})
        f1 = BytesIO()
        f2 = BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)
        data = {
            'file1': (f1, 'same1.xlsx'),
            'file2': (f2, 'same2.xlsx'),
            'compare_mode': 'field_match',
            'match_field': 'name',
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        resp_json = response.get_json()
        assert resp_json is not None
        assert resp_json['message'] == '两个文件中该字段值完全匹配，无差异'
    ```

- **tests/test_integration_convert.py**
  - language: `python`
  - content:
    ```python
import io
import time
import pandas as pd
import pytest
import zipfile

class TestConvertIntegration:
    def test_convert_excel_to_csv(self, client):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'file': (buf, 'test.xlsx'),
            'target_format': 'csv',
            'encoding': 'utf-8-sig',
            'separator': ','
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        max_attempts = 10
        status = None
        for _ in range(max_attempts):
            status_resp = client.get(f'/convert/status/{task_id}')
            status = status_resp.get_json()['data']
            if status['status'] == 'completed':
                break
            time.sleep(0.5)
        assert status is not None and status['status'] == 'completed'
        download_resp = client.get(f'/convert/download/{task_id}')
        assert download_resp.status_code == 200
        content = download_resp.data
        if content[:2] == b'PK':
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                with zf.open(zf.namelist()[0]) as f:
                    result_df = pd.read_csv(f)
        else:
            result_df = pd.read_csv(io.BytesIO(content))
        assert list(result_df.columns) == ['A', 'B']
        assert len(result_df) == 2

    def test_convert_csv_to_excel(self, client):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        buf = io.StringIO()
        df.to_csv(buf, index=False, encoding='utf-8-sig')
        buf.seek(0)
        data = {
            'file': (io.BytesIO(buf.getvalue().encode('utf-8')), 'test.csv'),
            'target_format': 'xlsx',
            'encoding': 'utf-8-sig',
            'separator': ','
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        for _ in range(10):
            status_resp = client.get(f'/convert/status/{task_id}')
            status = status_resp.get_json()['data']
            if status['status'] == 'completed':
                break
            time.sleep(0.5)
        assert status['status'] == 'completed'
        download_resp = client.get(f'/convert/download/{task_id}')
        assert download_resp.status_code == 200
        result_df = pd.read_excel(io.BytesIO(download_resp.data))
        assert list(result_df.columns) == ['A', 'B']

    def test_missing_file_error(self, client):
        response = client.post('/convert/process', data={'target_format': 'csv'})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '请上传文件' in response.get_json()['error']

    def test_invalid_target_format(self, client):
        df = pd.DataFrame({"A": [1]})
        buf = io.BytesIO()
        df.to_excel(buf, index=False, engine='openpyxl')
        buf.seek(0)
        data = {
            'file': (buf, 'test.xlsx'),
            'target_format': 'pdf',
            'encoding': 'utf-8-sig',
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        for _ in range(10):
            status = client.get(f'/convert/status/{task_id}').get_json()['data']
            if status['status'] == 'failed':
                break
            time.sleep(0.5)
        assert status['status'] == 'failed'
    ```

- **tests/test_integration_field_parser.py**
  - language: `python`
  - content:
    ```python
import json
import pytest

class TestFieldParserIntegration:
    def test_analyze_success(self, client):
        data = {
            "standard_headers": "姓名, 年龄, 性别",
            "compare_headers": "名字, 岁数, 性别"
        }
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 200
        result = response.get_json()
        assert len(result["matched"]) == 2

    def test_analyze_missing_field(self, client):
        data = {"standard_headers": "姓名"}
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 400
        result = response.get_json()
        assert "compare_headers" in result['error']

    def test_analyze_empty(self, client):
        data = {"standard_headers": "", "compare_headers": ""}
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 400
    ```

- **tests/test_integration_matcher.py**
  - language: `python`
  - content:
    ```python
import io
import pytest
import pandas as pd
from openpyxl import Workbook

class TestMatcherIntegration:
    @pytest.fixture
    def target_file(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value1'
        ws['C1'] = 'Value2'
        ws['A2'] = 1
        ws['B2'] = ''
        ws['C2'] = ''
        ws['A3'] = 2
        ws['B3'] = ''
        ws['C3'] = ''
        ws['A4'] = 3
        ws['B4'] = ''
        ws['C4'] = ''
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    @pytest.fixture
    def target_file_duplicate(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value1'
        ws['C1'] = 'Value2'
        ws['A2'] = 1
        ws['B2'] = ''
        ws['C2'] = ''
        ws['A3'] = 1
        ws['B3'] = ''
        ws['C3'] = ''
        ws['A4'] = 2
        ws['B4'] = ''
        ws['C4'] = ''
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    @pytest.fixture
    def source_file(self):
        df = pd.DataFrame({'ID': [1, 2, 3], 'Score1': [100, 200, 300], 'Score2': [101, 202, 303]})
        f = io.BytesIO()
        with pd.ExcelWriter(f, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        f.seek(0)
        return f

    def test_match_single_column(self, client, target_file, source_file):
        """测试单列匹配（向后兼容）"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert result_df.loc[0, 'Value1'] == 100
        assert result_df.loc[1, 'Value1'] == 200
        assert result_df.loc[2, 'Value1'] == 300

    def test_match_multiple_columns(self, client, target_file, source_file):
        """测试多列匹配：源值列 B,C -> 结果列 B,C"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert result_df.loc[0, 'Value1'] == 100
        assert result_df.loc[1, 'Value1'] == 200
        assert result_df.loc[2, 'Value1'] == 300
        assert result_df.loc[0, 'Value2'] == 101
        assert result_df.loc[1, 'Value2'] == 202
        assert result_df.loc[2, 'Value2'] == 303

    def test_match_multiple_columns_with_new_result_names(self, client, target_file, source_file):
        """测试多列匹配，结果列为新建列名"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': '新分数1,新分数2',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert '新分数1' in result_df.columns
        assert '新分数2' in result_df.columns
        assert result_df.loc[0, '新分数1'] == 100
        assert result_df.loc[0, '新分数2'] == 101

    def test_match_column_count_mismatch(self, client, target_file, source_file):
        """测试源值列与结果列数量不一致时返回错误"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B',   # 数量不匹配
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        resp_json = response.get_json()
        assert not resp_json['success']
        assert '数量' in resp_json['error']

    def test_match_with_duplicate_keys_disabled(self, client, target_file_duplicate, source_file):
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        resp_json = response.get_json()
        assert not resp_json['success']
        assert '重复' in resp_json['error']

    def test_match_with_duplicate_keys_enabled(self, client, target_file_duplicate, source_file):
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
            'allow_duplicate_keys': 'on',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        # 第一行填充，第二行（重复键）留空
        assert result_df.loc[0, 'Value1'] == 100
        assert result_df.loc[1, 'Value1'] == '' or pd.isna(result_df.loc[1, 'Value1'])
        assert result_df.loc[2, 'Value1'] == 200

    def test_debug_mode_multiple_columns(self, client, target_file, source_file):
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
        }
        response = client.post('/match/debug', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content = response.data.decode('utf-8')
        assert 'source_value_cols' in content or 'config' in content

    def test_match_with_generate_report_off(self, client, target_file, source_file):
        """测试关闭校验报告开关：结果 xlsx 不应包含'校验报告' sheet"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            'generate_report': 'off',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        xlsx = pd.ExcelFile(io.BytesIO(response.data))
        assert '校验报告' not in xlsx.sheet_names

    def test_match_with_generate_report_default_on(self, client, target_file, source_file):
        """测试默认行为：不传 generate_report 时，应生成校验报告（向后兼容）"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        xlsx = pd.ExcelFile(io.BytesIO(response.data))
        assert '校验报告' in xlsx.sheet_names

    @pytest.fixture
    def target_file_duplicate_unmatched(self):
        """ID=99 出现两次但源里没有 99；ID=1 匹配。
        用于验证 all 模式下未匹配重复行填 0（与 first 模式第 1 行一致）"""
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value1'
        ws['A2'] = 99
        ws['B2'] = 'orig1'
        ws['A3'] = 99
        ws['B3'] = 'orig2'
        ws['A4'] = 1
        ws['B4'] = ''
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    def test_duplicate_action_first_explicit(self, client, target_file_duplicate, source_file):
        """显式 duplicate_action='first'，行为应与 allow_duplicate_keys='on' 一致"""
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            'duplicate_action': 'first',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        # 第 1 行（ID=1）填 100
        assert result_df.loc[0, 'Value1'] == 100
        # 第 2 行（ID=1，重复）保留原值（空）
        assert result_df.loc[1, 'Value1'] == '' or pd.isna(result_df.loc[1, 'Value1'])
        # 第 3 行（ID=2）填 200
        assert result_df.loc[2, 'Value1'] == 200

    def test_duplicate_action_all(self, client, target_file_duplicate, source_file):
        """duplicate_action='all'，所有匹配行都填充相同值"""
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            'duplicate_action': 'all',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        # 两行 ID=1 都被填充 100
        assert result_df.loc[0, 'Value1'] == 100
        assert result_df.loc[1, 'Value1'] == 100
        # 第 3 行（ID=2）填 200
        assert result_df.loc[2, 'Value1'] == 200

    def test_duplicate_action_all_with_unmatched(self, client, target_file_duplicate_unmatched, source_file):
        """duplicate_action='all' 且重复 key 在源中无匹配：填 0（与 first 模式第 1 行一致）"""
        data = {
            'target_file': (target_file_duplicate_unmatched, 'target_dup_unmatched.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            'duplicate_action': 'all',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        # ID=99 两行均未匹配，都填 0（覆盖原值 orig1/orig2）
        assert result_df.loc[0, 'Value1'] == 0
        assert result_df.loc[1, 'Value1'] == 0
        # 第 3 行（ID=1）填 100
        assert result_df.loc[2, 'Value1'] == 100

    def test_duplicate_action_invalid_value(self, client, target_file_duplicate, source_file):
        """非法 duplicate_action 值降级为 'error'，检测到重复键时报错"""
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            'duplicate_action': 'xxx',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        resp_json = response.get_json()
        assert not resp_json['success']
        assert '重复' in resp_json['error']
    ```

- **tests/test_integration_splitter.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_splitter.py
import io, zipfile, pytest
from openpyxl import Workbook

class TestSplitterIntegration:
    @pytest.fixture
    def multi_sheet_excel(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = 'Data'
        ws1['A1'] = 'Category'
        ws1['B1'] = 'Value'
        ws1['A2'] = 'A'
        ws1['B2'] = 10
        ws1['A3'] = 'B'
        ws1['B3'] = 20
        ws2 = wb.create_sheet('Extra')
        ws2['A1'] = 'Category'
        ws2['A2'] = 'A'
        ws2['A3'] = 'B'
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    def test_split_basic(self, client, multi_sheet_excel):
        data = {
            'file': (multi_sheet_excel, 'test.xlsx'),
            'split_column': 'Category',
            'include_values': '',
            'exclude_values': '',
            'skip_sheets': '2'
        }
        response = client.post('/splitter/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        # 检查 Content-Disposition 是否包含统一前缀
        content_disp = response.headers.get('Content-Disposition', '')
        assert 'splitter_test' in content_disp or 'splitter' in content_disp
        zip_data = io.BytesIO(response.data)
        with zipfile.ZipFile(zip_data) as zf:
            names = zf.namelist()
            assert any('A.xlsx' in n for n in names)
            assert any('B.xlsx' in n for n in names)
    ```

- **tests/test_integration_subtotal.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_subtotal.py
import io
import pandas as pd

class TestSubtotalIntegration:
    def test_sync_endpoint(self, client):
        # 使用字母列名
        df = pd.DataFrame({
            'A': ['X', 'X', 'Y'],
            'B': [10, 20, 30],
            'C': [100, 200, 300]
        })
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B,C',
            'agg_func_default': 'sum',
            'sheets_mode': '1'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'subtotal_test' in resp.headers.get('Content-Disposition', '')
        result_df = pd.read_excel(io.BytesIO(resp.data), sheet_name=None)
        ws = result_df['Sheet1']
        assert len(ws) == 5
        assert ws.iloc[2]['B'] == 30
        assert ws.iloc[2]['C'] == 300

    def test_missing_dimension_col(self, client):
        df = pd.DataFrame({'A': [1]})
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'A',
            'dimension_cols': 'Z'   # 不存在的列字母
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 400
        assert '无效' in resp.get_json()['error'] or '超出' in resp.get_json()['error']

    def test_async_large_file(self, client):
        df = pd.DataFrame({'A': ['X', 'Y'] * 3000, 'B': range(6000)})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test_large.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B',
            'agg_func_default': 'sum'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'task_id' in resp.get_json()
    ```

- **tests/test_integration_summarizer.py**
  - language: `python`
  - content:
    ```python
import io
import pytest
import pandas as pd
import numpy as np

class TestSummarizerIntegration:
    @pytest.fixture
    def excel_files(self):
        df1 = pd.DataFrame({'机构': ['A', 'A'], '金额': [100, 200], '日期': ['2025-01-01', '2025-01-02']})
        df2 = pd.DataFrame({'机构': ['B'], '金额': [300], '日期': ['2025-01-03']})
        f1 = io.BytesIO()
        f2 = io.BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df1.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df2.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)
        return [f1, f2]

    def test_summarize_basic(self, client, excel_files):
        data = {
            'files': [(excel_files[0], 'file1.xlsx'), (excel_files[1], 'file2.xlsx')],
            'fields': '机构\n金额',
            'unique_keys': '机构',
            'sum_fields': '金额'
        }
        response = client.post('/summary/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert len(result_df) == 2
        assert result_df.loc[result_df['机构'] == 'A', '金额'].iloc[0] == 300
        assert result_df.loc[result_df['机构'] == 'B', '金额'].iloc[0] == 300

    def test_missing_fields_error(self, client):
        response = client.post('/summary/process', data={})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '没有选择文件' in response.get_json()['error']

    # --------------- 唯一键包含空值场景 ---------------
    def test_unique_keys_with_empty_values(self, client):
        # 构造两个文件，唯一键为 机构+日期，其中一行的日期为空
        df1 = pd.DataFrame({'机构': ['A', 'A', 'B'], '日期': ['2025-01', '', '2025-02'], '金额': [100, 200, 300]})
        df2 = pd.DataFrame({'机构': ['B'], '日期': ['2025-02'], '金额': [400]})

        f1 = io.BytesIO()
        f2 = io.BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df1.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df2.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)

        data = {
            'files': [(f1, 'file1.xlsx'), (f2, 'file2.xlsx')],
            'fields': '机构\n日期\n金额',
            'unique_keys': '机构,日期',
            'sum_fields': '金额'
        }
        response = client.post('/summary/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))

        # 预期结果：
        # A + ''  -> 200
        # A + '2025-01' -> 100
        # B + '2025-02' -> 700
        # 共3行
        assert len(result_df) == 3

        # 空值在 Excel 中读回可能是 NaN，统一处理后再比较
        # 将日期列的空值统一转换为空字符串
        result_df['日期'] = result_df['日期'].fillna('')

        row_a_empty = result_df[(result_df['机构'] == 'A') & (result_df['日期'] == '')]
        assert len(row_a_empty) == 1
        assert row_a_empty.iloc[0]['金额'] == 200

        row_b = result_df[(result_df['机构'] == 'B') & (result_df['日期'] == '2025-02')]
        assert len(row_b) == 1
        assert row_b.iloc[0]['金额'] == 700
    ```

- **tests/test_integration_team_report.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_team_report.py
import io, time, pandas as pd, pytest

class TestTeamReportIntegration:
    def test_generate_with_monthly(self, client):
        mapping_df = pd.DataFrame({"队员ID":["1","2"],"队伍名称":["骑士队","米兰队"],"队长姓名":["袁队","安队"],"司管":["方行","任翔"]})
        bm = io.BytesIO(); mapping_df.to_excel(bm,index=False); bm.seek(0)
        today_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[10,15],"报单数":[2,3],"转单数":[0,0],"新桔完单数":[5,6],"在线天数":[1,0]})
        bt = io.BytesIO(); today_df.to_excel(bt,index=False); bt.seek(0)
        yesterday_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[9,12],"报单数":[1,2],"转单数":[0,0],"新桔完单数":[4,5],"在线天数":[1,0]})
        by = io.BytesIO(); yesterday_df.to_excel(by,index=False); by.seek(0)
        monthly_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[100,150],"报单数":[20,30],"转单数":[5,10],"新桔完单数":[50,60]})
        bmon = io.BytesIO(); monthly_df.to_excel(bmon,index=False); bmon.seek(0)
        data = {
            "mapping_file": (bm, "m.xlsx"),
            "today_file": (bt, "t.xlsx"),
            "yesterday_file": (by, "y.xlsx"),
            "monthly_files": (bmon, "monthly.xlsx"),
            "daily_order_target": "1150",
            "daily_apply_target": "120",
            "monthly_order_target": "35650",
            "monthly_apply_target": "3720"
        }
        resp = client.post("/team-report/generate", data=data, content_type="multipart/form-data")
        assert resp.status_code == 200
        tid = resp.get_json()["task_id"]
        for _ in range(15):
            s = client.get(f"/team-report/status/{tid}").get_json()
            if s["status"] == "completed":
                break
            time.sleep(0.5)
        dl = client.get(f"/team-report/download/{tid}")
        assert dl.status_code == 200 and len(dl.data) > 0

    def test_generate_without_monthly(self, client):
        mapping_df = pd.DataFrame({"队员ID":["1"],"队伍名称":["骑士队"],"队长姓名":["袁队"],"司管":["方行"]})
        bm = io.BytesIO(); mapping_df.to_excel(bm,index=False); bm.seek(0)
        today_df = pd.DataFrame({"司机ID":["1"],"司机状态":["已上岗"],"总完单数":[10],"报单数":[2],"转单数":[0],"新桔完单数":[5],"在线天数":[1]})
        bt = io.BytesIO(); today_df.to_excel(bt,index=False); bt.seek(0)
        data = {"mapping_file":(bm,"m.xlsx"),"today_file":(bt,"t.xlsx")}
        resp = client.post("/team-report/generate", data=data, content_type="multipart/form-data")
        assert resp.status_code == 200
        tid = resp.get_json()["task_id"]
        for _ in range(10):
            s = client.get(f"/team-report/status/{tid}").get_json()
            if s["status"] == "completed":
                break
            time.sleep(0.3)
        dl = client.get(f"/team-report/download/{tid}")
        assert dl.status_code == 200
    ```

