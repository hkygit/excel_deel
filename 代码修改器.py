#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P1-1 修复脚本：TaskStorage 迁移到 SQLite + WAL

新增能力：
- 原子性：所有写操作走 SQLite 事务（BEGIN IMMEDIATE）
- 多进程安全：WAL 模式（读不阻塞写、写不阻塞读）
- 崩溃恢复：recover_orphan_tasks() 把 pending/running 标记为 failed
- 高效查询：按 status / created_at 建立索引
- 旧数据迁移：__init__ 时自动导入 status_dir/*.json
- 未知字段：写入 extra 列（JSON），load 时 merge 回 dict

接口完全兼容：
- save / load / update / delete / append_log / get_log / get_output_dir
- 所有 blueprint、manager、decorator 零改动

本脚本只写文件，不运行测试。
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

STORAGE_PATH = os.path.join(BASE_DIR, "core", "task", "storage.py")
TEST_PATH = os.path.join(BASE_DIR, "tests", "test_core_task.py")


# ============================================================================
# 1) core/task/storage.py 完整内容
# ============================================================================
STORAGE_CONTENT = r'''# core/task/storage.py
"""基于 SQLite + WAL 的任务状态存储。

设计要点：
- 单文件数据库（status_dir/tasks.db），零外部依赖（sqlite3 是标准库）
- WAL 模式：读不阻塞写，写不阻塞读，多进程安全
- 每个操作独立连接，避免跨线程/跨进程连接共享问题
- 所有写操作走 BEGIN IMMEDIATE 事务，保证原子性
- 自动迁移旧的 JSON 文件到数据库（备份到 .migrated/ 子目录）
- 未知字段自动落 extra 列（JSON），load 时 merge 回 dict

公开接口与旧版 JSON 实现保持一致，因此 blueprint / manager / decorator
无需任何改动。
"""
import json
import logging
import os
import sqlite3
import threading
from datetime import datetime
from typing import Any, Dict, Optional

logger = logging.getLogger(__name__)


# tasks 表中除主键 task_id 之外的固定列
_TASK_COLUMNS = (
    'status', 'created_at', 'started_at', 'completed_at',
    'result_path', 'result_format', 'error',
    'total_input_rows', 'total_output_rows',
    'progress_percent', 'progress_message',
    'params',                # 存储为 JSON 字符串
    'original_filename',
    'heartbeat',
)

# 需要在写入时 json.dumps / 读取时 json.loads 的列
_JSON_COLUMNS = {'params'}


class TaskStorage:
    """SQLite 支持的任务状态存储。

    与旧版 JSON 实现保持完全一致的方法签名，同时提供：
    - 原子性更新（SQLite 事务）
    - 多进程并发安全（WAL 模式）
    - 崩溃恢复（recover_orphan_tasks）
    - 高效按状态 / 时间查询（索引）
    """

    DB_FILENAME = 'tasks.db'

    def __init__(
        self,
        status_dir: str,
        output_dir: str,
        db_path: Optional[str] = None,
        auto_migrate: bool = True,
    ):
        self.status_dir = status_dir
        self.output_dir = output_dir
        self.db_path = db_path or os.path.join(status_dir, self.DB_FILENAME)

        # 兼容旧调用方：保留一个进程内的锁。真正的并发安全由 SQLite 提供。
        self._lock = threading.RLock()

        os.makedirs(status_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        self._init_db()

        if auto_migrate:
            try:
                self._migrate_json_files()
            except Exception as e:  # noqa: BLE001
                logger.warning(f"迁移旧 JSON 任务文件失败（不影响新功能）: {e}")

    # ------------------------------------------------------------------ #
    # 连接与初始化
    # ------------------------------------------------------------------ #
    def _connect(self) -> sqlite3.Connection:
        """打开一个新的连接。

        - isolation_level=None：显式控制事务（BEGIN IMMEDIATE / COMMIT）
        - PRAGMA 是 per-connection 的，因此每次连接都要设置一次
        - check_same_thread 保持默认 True：每次操作独立连接，不跨线程复用
        """
        conn = sqlite3.connect(
            self.db_path,
            timeout=10.0,
            isolation_level=None,
        )
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA busy_timeout=5000")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    def _init_db(self) -> None:
        """创建表与索引（幂等）。"""
        conn = self._connect()
        try:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id           TEXT PRIMARY KEY,
                    status            TEXT NOT NULL,
                    created_at        TEXT NOT NULL,
                    started_at        TEXT,
                    completed_at      TEXT,
                    result_path       TEXT,
                    result_format     TEXT,
                    error             TEXT,
                    total_input_rows  INTEGER,
                    total_output_rows INTEGER,
                    progress_percent  INTEGER,
                    progress_message  TEXT,
                    params            TEXT,
                    original_filename TEXT,
                    heartbeat         TEXT,
                    extra             TEXT
                );

                CREATE INDEX IF NOT EXISTS idx_tasks_status
                    ON tasks(status);
                CREATE INDEX IF NOT EXISTS idx_tasks_created
                    ON tasks(created_at DESC);

                CREATE TABLE IF NOT EXISTS task_logs (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id    TEXT NOT NULL,
                    timestamp  TEXT NOT NULL,
                    message    TEXT
                );

                CREATE INDEX IF NOT EXISTS idx_logs_task
                    ON task_logs(task_id, id);
            """)
        finally:
            conn.close()
        logger.debug(f"TaskStorage 初始化完成: {self.db_path}")

    # ------------------------------------------------------------------ #
    # 序列化辅助
    # ------------------------------------------------------------------ #
    def _row_to_dict(self, row: sqlite3.Row) -> Dict[str, Any]:
        """把 sqlite3.Row 转为 dict，解析 JSON 列，merge extra。"""
        d: Dict[str, Any] = {}
        for key in row.keys():
            if key == 'extra':
                continue
            d[key] = row[key]

        # 解析 JSON 列
        for col in _JSON_COLUMNS:
            if col in d and isinstance(d[col], str):
                try:
                    d[col] = json.loads(d[col])
                except (json.JSONDecodeError, TypeError):
                    # 保持原字符串
                    pass

        # merge extra
        extra_raw = row['extra']
        if extra_raw:
            try:
                extra = json.loads(extra_raw)
                if isinstance(extra, dict):
                    # 不覆盖已知列
                    for k, v in extra.items():
                        d.setdefault(k, v)
            except (json.JSONDecodeError, TypeError):
                logger.warning(f"无法解析 extra 字段: {extra_raw!r}")
        return d

    def _split_fields(self, data: Dict[str, Any]):
        """把 dict 拆成 (已知列 dict, 未知列 dict)。"""
        known: Dict[str, Any] = {}
        extra: Dict[str, Any] = {}
        for k, v in data.items():
            if k == 'task_id':
                continue
            if k in _TASK_COLUMNS:
                if k in _JSON_COLUMNS and not isinstance(v, str):
                    v = json.dumps(v, ensure_ascii=False)
                known[k] = v
            else:
                extra[k] = v
        return known, extra

    # ------------------------------------------------------------------ #
    # 核心 CRUD
    # ------------------------------------------------------------------ #
    def save(self, task_id: str, data: Dict) -> None:
        """完整保存任务状态（upsert）。

        - 若 task_id 不存在：INSERT 所有字段
        - 若 task_id 已存在：仅更新 data 中提供的字段（合并语义）
          · 已知列逐个更新
          · 未知字段 merge 到 extra
        """
        payload = dict(data)
        payload.pop('task_id', None)   # task_id 来自参数，忽略 data 里可能的重复
        payload.setdefault('status', 'pending')
        payload.setdefault('created_at', datetime.now().isoformat())

        known, extra = self._split_fields(payload)

        with self._lock:
            conn = self._connect()
            try:
                conn.execute("BEGIN IMMEDIATE")

                cur = conn.execute(
                    "SELECT extra FROM tasks WHERE task_id = ?", (task_id,)
                )
                row = cur.fetchone()
                exists = row is not None

                if exists:
                    # ---------------- UPDATE ----------------
                    set_parts, sql_params = [], []
                    for k, v in known.items():
                        if k in _JSON_COLUMNS and not isinstance(v, str):
                            v = json.dumps(v, ensure_ascii=False)
                        set_parts.append(f"{k} = ?")
                        sql_params.append(v)

                    if extra:
                        existing_extra = {}
                        if row['extra']:
                            try:
                                existing_extra = json.loads(row['extra']) or {}
                            except (json.JSONDecodeError, TypeError):
                                existing_extra = {}
                        existing_extra.update(extra)
                        set_parts.append("extra = ?")
                        sql_params.append(
                            json.dumps(existing_extra, ensure_ascii=False)
                        )

                    if set_parts:
                        sql_params.append(task_id)
                        conn.execute(
                            f"UPDATE tasks SET {', '.join(set_parts)} "
                            f"WHERE task_id = ?",
                            sql_params,
                        )
                else:
                    # ---------------- INSERT ----------------
                    cols = ['task_id'] + list(known.keys()) + ['extra']
                    placeholders = ','.join(['?'] * len(cols))
                    values = (
                        [task_id]
                        + list(known.values())
                        + [json.dumps(extra, ensure_ascii=False) if extra else None]
                    )
                    conn.execute(
                        f"INSERT INTO tasks ({','.join(cols)}) "
                        f"VALUES ({placeholders})",
                        values,
                    )

                conn.execute("COMMIT")
            except Exception:
                try:
                    conn.execute("ROLLBACK")
                except sqlite3.OperationalError:
                    pass
                raise
            finally:
                conn.close()

    def load(self, task_id: str) -> Optional[Dict]:
        conn = self._connect()
        try:
            cur = conn.execute(
                "SELECT * FROM tasks WHERE task_id = ?", (task_id,)
            )
            row = cur.fetchone()
            if row is None:
                return None
            return self._row_to_dict(row)
        finally:
            conn.close()

    def update(self, task_id: str, **fields) -> None:
        """更新部分字段。

        若 task_id 不存在，抛 ValueError（与旧实现行为一致）。
        """
        if not fields:
            return

        with self._lock:
            conn = self._connect()
            try:
                conn.execute("BEGIN IMMEDIATE")

                cur = conn.execute(
                    "SELECT extra FROM tasks WHERE task_id = ?", (task_id,)
                )
                row = cur.fetchone()
                if row is None:
                    conn.execute("ROLLBACK")
                    raise ValueError(f"Task {task_id} not found")

                known, extra = self._split_fields(fields)

                set_parts, sql_params = [], []
                for k, v in known.items():
                    if k in _JSON_COLUMNS and not isinstance(v, str):
                        v = json.dumps(v, ensure_ascii=False)
                    set_parts.append(f"{k} = ?")
                    sql_params.append(v)

                if extra:
                    existing_extra = {}
                    if row['extra']:
                        try:
                            existing_extra = json.loads(row['extra']) or {}
                        except (json.JSONDecodeError, TypeError):
                            existing_extra = {}
                    existing_extra.update(extra)
                    set_parts.append("extra = ?")
                    sql_params.append(
                        json.dumps(existing_extra, ensure_ascii=False)
                    )

                if set_parts:
                    sql_params.append(task_id)
                    conn.execute(
                        f"UPDATE tasks SET {', '.join(set_parts)} "
                        f"WHERE task_id = ?",
                        sql_params,
                    )

                conn.execute("COMMIT")
            except ValueError:
                raise
            except Exception:
                try:
                    conn.execute("ROLLBACK")
                except sqlite3.OperationalError:
                    pass
                raise
            finally:
                conn.close()

    def delete(self, task_id: str) -> None:
        with self._lock:
            conn = self._connect()
            try:
                conn.execute("BEGIN IMMEDIATE")
                conn.execute(
                    "DELETE FROM tasks WHERE task_id = ?", (task_id,)
                )
                conn.execute(
                    "DELETE FROM task_logs WHERE task_id = ?", (task_id,)
                )
                conn.execute("COMMIT")
            except Exception:
                try:
                    conn.execute("ROLLBACK")
                except sqlite3.OperationalError:
                    pass
                raise
            finally:
                conn.close()

    def get_output_dir(self) -> str:
        return self.output_dir

    # ------------------------------------------------------------------ #
    # 日志
    # ------------------------------------------------------------------ #
    def append_log(self, task_id: str, message: str) -> None:
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(
                "INSERT INTO task_logs (task_id, timestamp, message) "
                "VALUES (?, ?, ?)",
                (task_id, ts, message),
            )
            conn.execute("COMMIT")
        except Exception:
            try:
                conn.execute("ROLLBACK")
            except sqlite3.OperationalError:
                pass
            raise
        finally:
            conn.close()

    def get_log(self, task_id: str, max_lines: int = 2000) -> str:
        conn = self._connect()
        try:
            cur = conn.execute(
                "SELECT timestamp, message FROM task_logs "
                "WHERE task_id = ? ORDER BY id DESC LIMIT ?",
                (task_id, max_lines),
            )
            rows = list(reversed(cur.fetchall()))
            return ''.join(
                f"[{r['timestamp']}] {r['message']}\n" for r in rows
            )
        finally:
            conn.close()

    # ------------------------------------------------------------------ #
    # 崩溃恢复
    # ------------------------------------------------------------------ #
    def recover_orphan_tasks(
        self, reason: str = '服务重启导致任务中断'
    ) -> int:
        """把所有 pending / running 的任务标记为 failed。

        服务启动时调用，避免 daemon 线程被杀后任务永远卡在 running。

        Returns:
            被恢复的任务数量。
        """
        now = datetime.now().isoformat()
        with self._lock:
            conn = self._connect()
            try:
                conn.execute("BEGIN IMMEDIATE")
                cur = conn.execute(
                    "SELECT COUNT(*) AS c FROM tasks "
                    "WHERE status IN ('pending', 'running')"
                )
                count = cur.fetchone()['c']
                if count > 0:
                    conn.execute(
                        "UPDATE tasks "
                        "SET status = 'failed', error = ?, completed_at = ? "
                        "WHERE status IN ('pending', 'running')",
                        (reason, now),
                    )
                conn.execute("COMMIT")
                if count:
                    logger.warning(
                        f"已恢复 {count} 个僵尸任务为 failed"
                    )
                return count
            except Exception:
                try:
                    conn.execute("ROLLBACK")
                except sqlite3.OperationalError:
                    pass
                raise
            finally:
                conn.close()

    # ------------------------------------------------------------------ #
    # 旧 JSON 迁移
    # ------------------------------------------------------------------ #
    def _migrate_json_files(self) -> None:
        """扫描 status_dir/*.json 导入到 DB，然后移动到 .migrated/。"""
        if not os.path.isdir(self.status_dir):
            return

        try:
            json_files = [
                f for f in os.listdir(self.status_dir)
                if f.endswith('.json') and not f.startswith('.')
            ]
        except OSError as e:
            logger.warning(f"无法列出 status_dir: {e}")
            return

        if not json_files:
            return

        migrated_dir = os.path.join(self.status_dir, '.migrated')
        os.makedirs(migrated_dir, exist_ok=True)

        imported = 0
        for fname in json_files:
            json_path = os.path.join(self.status_dir, fname)
            task_id = fname[:-5]  # 去掉 .json

            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                if not content:
                    continue
                data = json.loads(content)
                if not isinstance(data, dict):
                    logger.warning(f"跳过非 dict 的旧任务文件: {json_path}")
                    continue
            except (json.JSONDecodeError, OSError) as e:
                logger.warning(f"跳过损坏的 JSON 任务文件 {json_path}: {e}")
                continue

            # 已存在则跳过（但仍移动旧文件）
            if self.load(task_id) is None:
                try:
                    data.setdefault('task_id', task_id)
                    self.save(task_id, data)
                    imported += 1
                except Exception as e:  # noqa: BLE001
                    logger.warning(f"导入旧任务 {task_id} 失败: {e}")
                    continue

            # 迁移对应的 .log 文件（如果有）
            log_path = os.path.join(self.status_dir, f"{task_id}.log")
            if os.path.exists(log_path):
                try:
                    log_rows = []
                    with open(log_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.rstrip('\n')
                            if not line:
                                continue
                            if line.startswith('[') and ']' in line:
                                ts, msg = line[1:].split(']', 1)
                                msg = msg.lstrip()
                            else:
                                ts = datetime.now().strftime(
                                    '%Y-%m-%d %H:%M:%S'
                                )
                                msg = line
                            log_rows.append((task_id, ts, msg))

                    if log_rows:
                        conn = self._connect()
                        try:
                            conn.execute("BEGIN IMMEDIATE")
                            conn.executemany(
                                "INSERT INTO task_logs "
                                "(task_id, timestamp, message) "
                                "VALUES (?, ?, ?)",
                                log_rows,
                            )
                            conn.execute("COMMIT")
                        except Exception:
                            try:
                                conn.execute("ROLLBACK")
                            except sqlite3.OperationalError:
                                pass
                            raise
                        finally:
                            conn.close()

                    os.replace(
                        log_path, os.path.join(migrated_dir, f"{task_id}.log")
                    )
                except Exception as e:  # noqa: BLE001
                    logger.warning(f"迁移日志 {log_path} 失败: {e}")

            # 移动 JSON 文件
            try:
                os.replace(json_path, os.path.join(migrated_dir, fname))
            except OSError as e:
                logger.warning(
                    f"移动旧任务文件 {json_path} 失败: {e}"
                )

        if imported:
            logger.info(f"已从 JSON 迁移 {imported} 个任务到 SQLite")
'''


# ============================================================================
# 2) tests/test_core_task.py 完整内容
# ============================================================================
TEST_CONTENT = r'''# tests/test_core_task.py
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
'''


# ============================================================================
# 写文件
# ============================================================================
def write_file(path: str, content: str) -> None:
    os.makedirs(os.path.dirname(path) or BASE_DIR, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[OK] 已写入 {os.path.relpath(path, BASE_DIR)}  ({len(content):,} 字符)")


def main() -> None:
    print("=" * 72)
    print("P1-1：TaskStorage → SQLite + WAL")
    print("=" * 72)

    write_file(STORAGE_PATH, STORAGE_CONTENT)
    write_file(TEST_PATH, TEST_CONTENT)

    print()
    print("=" * 72)
    print("✅ 全部文件已写入（未运行任何测试）")
    print("=" * 72)
    print()
    print("核心变更：")
    print("  core/task/storage.py")
    print("    · SQLite 后端（tasks.db），零外部依赖")
    print("    · WAL 模式 + busy_timeout=5000（多进程安全）")
    print("    · 所有写操作走 BEGIN IMMEDIATE 事务（原子性）")
    print("    · 未知字段落 extra 列（JSON），load 时 merge")
    print("    · 日志迁移到 task_logs 表（按 task_id + id 索引）")
    print("    · 新增 recover_orphan_tasks()（崩溃恢复）")
    print("    · __init__ 自动迁移 status_dir/*.json → .migrated/")
    print("    · 公开方法签名与旧版完全一致（零调用方改动）")
    print()
    print("  tests/test_core_task.py  (重写)")
    print("    · 保留原有 8 个测试")
    print("    · 新增：extra 合并、日志隔离、日志 limit、WAL 校验")
    print("    · 新增：崩溃恢复、并发更新不丢数据、JSON 迁移（含损坏文件）")
    print("    · 新增：manager 与 recover 联动")
    print()
    print("下一步（由你手动执行）：")
    print("    python -m pytest tests/test_core_task.py -v")
    print("    python -m pytest tests/test_integration_checker.py tests/test_integration_convert.py tests/test_integration_subtotal.py -v")
    print("    python -m pytest tests/ -v")
    print()
    print("建议后续小步骤（本次未做）：")
    print("  · app.py::create_app 中调用 container.task_storage().recover_orphan_tasks()")
    print("    （可选：接入 P1-2 DI 容器重构时一起做）")


if __name__ == "__main__":
    main()