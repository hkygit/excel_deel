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
