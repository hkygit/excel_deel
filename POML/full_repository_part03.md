- **core/sqlite_buffer.py**
  - language: `python`
  - content:
    ```python
# core/sqlite_buffer.py
import sqlite3, tempfile, os, logging
from typing import List, Optional

logger = logging.getLogger(__name__)

class SqliteBuffer:
    """管理临时 SQLite 数据库，支持建表、建索引、批量插入、去重"""
    def __init__(self, cache_size: int = -100000):
        fd, self.db_path = tempfile.mkstemp(suffix='.db')
        os.close(fd)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA synchronous = OFF")
        self.conn.execute("PRAGMA journal_mode = MEMORY")
        self.conn.execute(f"PRAGMA cache_size = {cache_size}")

    def create_table(self, table: str, columns: List[str], key_columns: List[str] = None):
        col_defs = ', '.join([f'`{c}` TEXT' for c in columns])
        self.conn.execute(f"CREATE TABLE {table} ({col_defs})")
        if key_columns:
            for col in key_columns:
                self.conn.execute(f'CREATE INDEX IF NOT EXISTS idx_{table}_{col} ON {table}("{col}")')

    def insert_batch(self, table: str, columns: List[str], rows: list):
        placeholders = ', '.join(['?' for _ in columns])
        self.conn.executemany(f"INSERT INTO {table} VALUES ({placeholders})", rows)

    def remove_duplicates(self, table: str, key_columns: List[str], keep: str = 'first'):
        if keep == 'first':
            # 使用 rowid 保留第一条（按插入顺序）
            cols = ','.join(key_columns)
            self.conn.execute(f"CREATE TABLE {table}_dedup AS SELECT * FROM {table} WHERE rowid IN (SELECT MIN(rowid) FROM {table} GROUP BY {cols})")
            self.conn.execute(f"DROP TABLE {table}")
            self.conn.execute(f"ALTER TABLE {table}_dedup RENAME TO {table}")

    def execute(self, sql, params=None):
        return self.conn.execute(sql, params or [])

    def fetch_df(self, sql):
        import pandas as pd
        return pd.read_sql(sql, self.conn)

    def close(self):
        self.conn.close()
        try:
            os.unlink(self.db_path)
        except Exception as e:
            logger.warning(f"清理临时数据库失败: {e}")
    ```

- **core/task/__init__.py**
  - language: `python`
  - content:
    ```python

    ```

- **core/task/base.py**
  - language: `python`
  - content:
    ```python
# core/task/base.py
from abc import ABC, abstractmethod
from typing import Any

class Task(ABC):
    @abstractmethod
    def execute(self) -> Any:
        pass
    ```

- **core/task/decorator.py**
  - language: `python`
  - content:
    ```python
# core/task/decorator.py
import functools
from datetime import datetime
from flask import current_app
from .state import TaskStatus

def async_task(original_filename=None):
    """
    装饰器：将函数包装为异步任务。
    被装饰函数必须接受 task_id 作为第一个参数，并返回一个字典（将作为任务结果保存）。
    示例：
        @async_task(original_filename="my_task")
        def my_task(task_id, param1, param2):
            # 耗时操作
            return {"result_path": "/path/to/result"}
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            app = current_app._get_current_object()
            manager = app.container.task_manager()
            storage = app.container.task_storage()
            task_id = manager.create_task(params={}, original_filename=original_filename)

            def run():
                with app.app_context():
                    try:
                        storage.update(task_id, status=TaskStatus.RUNNING, started_at=datetime.now().isoformat())
                        result = func(task_id, *args, **kwargs)
                        # 移除可能由任务返回的 completed_at，因为 complete_task 会自己设置
                        if result and 'completed_at' in result:
                            result.pop('completed_at')
                        manager.complete_task(task_id, **result)
                    except Exception as e:
                        import traceback
                        error_msg = f"{str(e)}\n{traceback.format_exc()}"
                        storage.append_log(task_id, error_msg)
                        manager.fail_task(task_id, error=error_msg)

            manager.start_task(task_id, run, ())
            return {"task_id": task_id}
        return wrapper
    return decorator
    ```

- **core/task/logger.py**
  - language: `python`
  - content:
    ```python
# utils/task/logger.py
import json
import logging
from datetime import datetime
from typing import Any

_logger = logging.getLogger('app')

def log_event(task_id: str, event: str, **kwargs: Any):
    log_entry = {
        'timestamp': datetime.now().isoformat(),
        'task_id': task_id,
        'event': event,
        **kwargs
    }
    _logger.info(json.dumps(log_entry, ensure_ascii=False))
    ```

- **core/task/manager.py**
  - language: `python`
  - content:
    ```python
# core/task/manager.py
import threading
import uuid
from datetime import datetime
from typing import Callable, Dict, Optional
from .state import TaskStatus
from .storage import TaskStorage
from .logger import log_event

class TaskManager:
    def __init__(self, storage: TaskStorage, max_concurrent: int = 5):
        self.storage = storage
        self._lock = threading.Lock()
        self._running_count = 0
        self._max_concurrent = max_concurrent

    def submit(self, task_id: str, target: Callable, args: tuple) -> bool:
        with self._lock:
            if self._running_count >= self._max_concurrent:
                return False
            self._running_count += 1
        def wrapper():
            try:
                target(*args)
            finally:
                with self._lock:
                    self._running_count -= 1
        thread = threading.Thread(target=wrapper, daemon=True)
        thread.start()
        return True

    def create_task(self, params: Dict, original_filename: str = None) -> str:
        task_id = str(uuid.uuid4())
        status = {
            'task_id': task_id,
            'status': TaskStatus.PENDING,
            'created_at': datetime.now().isoformat(),
            'started_at': None,
            'completed_at': None,
            'result_path': None,
            'result_format': None,
            'error': None,
            'total_input_rows': None,
            'total_output_rows': None,
            'params': params,
            'original_filename': original_filename
        }
        self.storage.save(task_id, status)
        log_event(task_id, 'task_created', params=params)
        return task_id

    def start_task(self, task_id: str, target: Callable, args: tuple):
        self.storage.update(task_id, status=TaskStatus.RUNNING, started_at=datetime.now().isoformat())
        log_event(task_id, 'task_started')
        if not self.submit(task_id, target, args):
            self.storage.update(task_id, status=TaskStatus.FAILED, error="系统繁忙，请稍后再试")
            return False
        return True

    def complete_task(self, task_id: str, **result):
        self.storage.update(task_id, status=TaskStatus.COMPLETED, completed_at=datetime.now().isoformat(), **result)
        log_event(task_id, 'task_completed', **result)

    def fail_task(self, task_id: str, error: str):
        self.storage.update(task_id, status=TaskStatus.FAILED, error=error, completed_at=datetime.now().isoformat())
        log_event(task_id, 'task_failed', error=error)

    def cancel_task(self, task_id: str):
        self.storage.update(task_id, status=TaskStatus.CANCELLED, completed_at=datetime.now().isoformat())
        log_event(task_id, 'task_cancelled')
    ```

- **core/task/state.py**
  - language: `python`
  - content:
    ```python
# utils/task/state.py
class TaskStatus:
    PENDING = 'pending'
    RUNNING = 'running'
    COMPLETED = 'completed'
    FAILED = 'failed'
    CANCELLED = 'cancelled'
    ```

- **core/task/storage.py**
  - language: `python`
  - content:
    ```python
# core/task/storage.py
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
    ```

- **core/validators.py**
  - language: `python`
  - content:
    ```python
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
    ```

- **install_requirements.py**
  - language: `python`
  - content:
    ```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健壮的依赖安装脚本
- 自动轮换多个镜像源（阿里云、豆瓣、华为、清华、官方）
- 跳过 pip 升级，避免 403 错误
- 支持超时和重试
"""

import subprocess
import sys
import time
import os

REQUIREMENTS_FILE = "requirements.txt"
MAX_RETRIES = 2          # 每个镜像重试次数
RETRY_DELAY = 5          # 重试间隔（秒）
TIMEOUT = 180            # 安装超时（秒）

# 镜像源列表（按优先级排序）
MIRRORS = [
    "https://mirrors.aliyun.com/pypi/simple/",
    "https://pypi.douban.com/simple/",
    "https://repo.huaweicloud.com/repository/pypi/simple/",
    "https://pypi.tuna.tsinghua.edu.cn/simple/",
    "https://pypi.org/simple/",          # 官方源，最后尝试
]

def run_pip_install(mirror):
    """尝试使用指定镜像源安装依赖"""
    host = mirror.split("//")[1].split("/")[0]
    cmd = [
        sys.executable, "-m", "pip", "install",
        "-r", REQUIREMENTS_FILE,
        "--timeout", str(TIMEOUT),
        "--retries", str(MAX_RETRIES),
        "-i", mirror,
        "--trusted-host", host,
        "--exists-action", "w",
    ]
    print(f"\n>>> 使用镜像源: {mirror}")
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            result = subprocess.run(cmd, timeout=TIMEOUT, check=False, capture_output=True, text=True)
            if result.returncode == 0:
                print("✅ 安装成功！")
                print(result.stdout)
                return True
            else:
                print(f"安装失败 (尝试 {attempt}/{MAX_RETRIES})，返回码 {result.returncode}")
                if result.stderr:
                    print("错误详情:", result.stderr[-500:])  # 只打印最后500字符
                if attempt < MAX_RETRIES:
                    print(f"等待 {RETRY_DELAY} 秒后重试...")
                    time.sleep(RETRY_DELAY)
        except subprocess.TimeoutExpired:
            print(f"安装超时（{TIMEOUT}秒），尝试 {attempt}/{MAX_RETRIES}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)
    return False

def main():
    print("=" * 60)
    print("Python 依赖安装工具（多镜像源自动轮换）")
    print(f"Python 版本: {sys.version}")
    print(f"依赖文件: {REQUIREMENTS_FILE}")
    if not os.path.exists(REQUIREMENTS_FILE):
        print(f"错误: 找不到 {REQUIREMENTS_FILE} 文件")
        sys.exit(1)
    print("=" * 60)

    # 不升级 pip，避免 403 错误
    print("跳过 pip 升级，使用当前版本。")

    for mirror in MIRRORS:
        success = run_pip_install(mirror)
        if success:
            print("\n🎉 所有依赖安装成功！")
            return
        print(f"镜像源 {mirror} 失败，切换到下一个...")
        time.sleep(RETRY_DELAY)

    print("\n❌ 所有镜像源均失败，请手动安装缺失的包。")
    print("建议手动执行：pip install -r requirements.txt")
    sys.exit(1)

if __name__ == "__main__":
    main()
    ```

- **requirements.txt**
  - language: ``
  - content:
    ```
flask>=2.3.3
pandas==2.3.3
numpy==2.4.3
openpyxl==3.1.2
xlrd==2.0.1
gunicorn==21.2.0
waitress>=2.0.0
Werkzeug>=2.0.0
psutil>=5.9.0
pydantic>=2.0.0
pydantic-settings>=2.0.0
dependency-injector>=4.41.0
pytest>=8.0.0
pytest-cov>=4.1.0
    ```

- **run_tests.py**
  - language: `python`
  - content:
    ```python
#!/usr/bin/env python3
"""
测试运行脚本

⭐ 跳过不常用模块：
   默认跳过 SKIP_TEST_FILES 中列出的测试文件（当前：车队日报）。
   命令行加 --all 时跑全量，加 --skip 追加跳过文件。

- 全部通过：仅打印成功信息，不生成日志
- 有失败或异常：落盘到 log/ 目录
- 区分"pytest 启动失败"与"真正 0 失败"

用法：
    python run_tests.py                 # 默认（跳过车队日报）
    python run_tests.py --all           # 全量
    python run_tests.py --skip tests/test_xxx.py,tests/test_yyy.py  # 追加跳过
    python run_tests.py --show-skip     # 显示当前跳过列表
"""

import os
import sys
import subprocess
import re
import traceback
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# 环境守卫：强制切换到项目 .venv
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
VENV_PY = PROJECT_ROOT / ".venv" / "Scripts" / "python.exe"
LOG_DIR = PROJECT_ROOT / "log"
TEST_TIMEOUT_SECONDS = 30 * 60


def _ensure_project_venv():
    if not VENV_PY.exists():
        print(f"⚠️ 项目 venv 不存在: {VENV_PY}")
        print("   请先运行 fix_env.py 或手动创建 .venv")
        os.environ.pop("PYTHONPATH", None)
        return

    try:
        current = Path(sys.executable).resolve()
        target = VENV_PY.resolve()
    except OSError:
        return

    if current == target:
        if os.environ.get("PYTHONPATH"):
            print(f"[环境清理] PYTHONPATH 已移除")
            os.environ.pop("PYTHONPATH", None)
        return

    print(f"[环境切换] 当前解释器: {current}")
    print(f"[环境切换] 项目 venv:  {target}")
    print(f"[环境切换] 重新启动 run_tests.py 到项目 venv ...")
    sys.stdout.flush()
    sys.stderr.flush()

    clean_env = os.environ.copy()
    clean_env.pop("PYTHONPATH", None)
    clean_env.pop("PYTHONHOME", None)
    path_parts = clean_env.get("PATH", "").split(os.pathsep)
    filtered = []
    for p in path_parts:
        pl = p.lower()
        if ".venv" in pl and "scripts" in pl and str(PROJECT_ROOT).lower() not in pl:
            continue
        filtered.append(p)
    clean_env["PATH"] = os.pathsep.join(filtered)

    try:
        result = subprocess.run(
            [str(target), str(Path(__file__).resolve())] + sys.argv[1:],
            env=clean_env,
        )
        sys.exit(result.returncode)
    except Exception as e:
        print(f"[环境切换] 失败: {e}")
        os.environ.pop("PYTHONPATH", None)


_ensure_project_venv()


# ---------------------------------------------------------------------------
# 跳过列表（默认跳过的测试文件）
# 要恢复某个文件：从列表里删掉即可
# 要新增跳过：往列表里加相对路径
# ---------------------------------------------------------------------------
SKIP_TEST_FILES = [
    # ---- 车队日报（不常用，跳过） ----
    "tests/test_team_report_long_running.py",
    "tests/test_integration_team_report.py",
    "tests/test_team_report_service.py",
]


# ---------------------------------------------------------------------------
# 参数解析
# ---------------------------------------------------------------------------
def parse_args():
    run_all = "--all" in sys.argv
    show_skip = "--show-skip" in sys.argv

    extra_skip = []
    for i, arg in enumerate(sys.argv):
        if arg == "--skip" and i + 1 < len(sys.argv):
            extra_skip = [
                s.strip() for s in sys.argv[i + 1].split(",") if s.strip()
            ]
        elif arg.startswith("--skip="):
            extra_skip = [
                s.strip() for s in arg[len("--skip="):].split(",") if s.strip()
            ]

    return run_all, show_skip, extra_skip


# ---------------------------------------------------------------------------
# 构造命令
# ---------------------------------------------------------------------------
def build_cmd(run_all, extra_skip):
    cmd = [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"]

    if run_all:
        print("🔍 全量模式（不跳过任何测试）")
        return cmd

    skip_files = list(SKIP_TEST_FILES) + [s for s in extra_skip if s not in SKIP_TEST_FILES]

    # 过滤不存在的文件（避免 pytest 因为 --ignore 不存在的路径报错）
    valid_skip = []
    for f in skip_files:
        p = PROJECT_ROOT / f
        if p.exists():
            valid_skip.append(f)
        else:
            print(f"⚠️ 跳过列表中文件不存在，忽略: {f}")

    if valid_skip:
        print(f"⚡ 跳过 {len(valid_skip)} 个文件（共约 {len(skip_files)} 个请求跳过）")
        for f in valid_skip:
            cmd.extend(["--ignore", f])
    else:
        print("⚡ 跳过列表为空，等价于全量")

    return cmd


# ---------------------------------------------------------------------------
# 运行
# ---------------------------------------------------------------------------
def run_tests(cmd):
    print(f"🚀 运行测试: {' '.join(cmd)}")
    print(f"   解释器: {sys.executable}")
    print(f"   版本:   {sys.version.split()[0]}")

    started = datetime.now()
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=TEST_TIMEOUT_SECONDS,
        )
        output = (result.stdout or "") + (result.stderr or "")
        elapsed = (datetime.now() - started).total_seconds()
        return result.returncode, output, None, elapsed
    except subprocess.TimeoutExpired as e:
        partial = ""
        for stream in (e.stdout, e.stderr):
            if not stream:
                continue
            partial += stream if isinstance(stream, str) else stream.decode("utf-8", "replace")
        elapsed = (datetime.now() - started).total_seconds()
        return -1, partial, f"pytest 运行超时（>{TEST_TIMEOUT_SECONDS}s）", elapsed
    except Exception as e:
        elapsed = (datetime.now() - started).total_seconds()
        return -1, "", f"无法启动 pytest: {e}\n{traceback.format_exc()}", elapsed


def extract_failed_tests(output: str) -> list:
    tests = []
    for pat in (r"^(.+?\.py::[^ ]+)\s+FAILED", r"^(.+?\.py::[^ ]+)\s+ERROR"):
        for m in re.finditer(pat, output or "", re.MULTILINE):
            name = m.group(1)
            if name not in tests:
                tests.append(name)
    return tests


def looks_like_pytest_bootstrap_failure(output: str) -> bool:
    if not output:
        return True
    if "test session starts" in output:
        return False
    markers = ("ModuleNotFoundError", "ImportError:",
               "Traceback (most recent call last)", "No module named")
    return any(m in output for m in markers)


def save_failure_log(output, failed_tests, returncode=None,
                     error_msg=None, elapsed=None, bootstrap_failed=False):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"test_failures_{ts}.log"

    lines = [f"测试运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"]
    lines.append(f"解释器: {sys.executable}")
    lines.append(f"版本:   {sys.version.split()[0]}")
    if elapsed is not None:
        lines.append(f"耗时: {elapsed:.2f}s")
    if returncode is not None:
        lines.append(f"pytest 退出码: {returncode}")
    if error_msg:
        lines.append(f"运行异常: {error_msg}")
    if bootstrap_failed:
        lines.append("⚠️ pytest 启动失败（未能正常 collect）")
    lines.append(f"失败/错误测试数量: {len(failed_tests)}")
    lines.append("-" * 60)

    if failed_tests:
        lines.append("失败/错误测试列表:")
        for t in failed_tests:
            lines.append(f"  - {t}")
        lines.append("-" * 60)

    log_file.write_text("\n".join(lines) + "\n\n" + (output or "(无输出)"), encoding="utf-8")

    print("\n" + "=" * 60)
    if bootstrap_failed:
        print("❌ pytest 启动失败（还没开始 collect）")
    else:
        print("⚠️ 测试未通过（或运行异常），完整输出如下：")
    print("=" * 60)
    print(output if output else "(无输出)")
    print("=" * 60)
    print(f"\n📁 失败日志已保存至: {log_file}")

    if bootstrap_failed:
        print("\n💡 诊断建议：")
        print(f"   1) 确认项目 venv 依赖齐全：")
        print(f"      {VENV_PY} -m pip list")
        print(f"   2) 若缺包，重装：")
        print(f"      {VENV_PY} -m pip install -r requirements.txt")
    elif failed_tests:
        print("\n失败/错误测试列表：")
        for t in failed_tests:
            print(f"   ❌ {t}")


def main():
    run_all, show_skip, extra_skip = parse_args()

    if show_skip:
        print("当前默认跳过的测试文件：")
        for f in SKIP_TEST_FILES:
            exists = "✓" if (PROJECT_ROOT / f).exists() else "✗ (不存在)"
            print(f"  {exists} {f}")
        return 0

    cmd = build_cmd(run_all, extra_skip)
    returncode, output, error_msg, elapsed = run_tests(cmd)

    if returncode == 0 and error_msg is None:
        print(f"\n✅ 所有测试通过！耗时 {elapsed:.2f}s。未生成日志文件。")
        return 0

    failed_tests = extract_failed_tests(output)
    bootstrap_failed = (not failed_tests) and looks_like_pytest_bootstrap_failure(output)

    save_failure_log(output, failed_tests, returncode=returncode,
                     error_msg=error_msg, elapsed=elapsed,
                     bootstrap_failed=bootstrap_failed)
    sys.exit(returncode if returncode != 0 else 1)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        tb = traceback.format_exc()
        try:
            save_failure_log(tb, [], returncode=-1,
                             error_msg="run_tests.py 自身异常",
                             bootstrap_failed=True)
        except Exception:
            print(tb, file=sys.stderr)
        sys.exit(1)
    ```

- **runtime.txt**
  - language: ``
  - content:
    ```
python-3.11

    ```

- **schemas/comparator.py**
  - language: `python`
  - content:
    ```python
# schemas/comparator.py
from pydantic import BaseModel, Field, field_validator
from typing import List, Optional

class CompareFullParams(BaseModel):
    key_fields: List[str] = Field(..., min_length=1, description="主键字段列表")
    compare_fields: List[str] = Field(..., min_length=1, description="对比字段列表")
    sheet1: Optional[str] = None
    sheet2: Optional[str] = None

    @field_validator('key_fields', 'compare_fields', mode='before')
    @classmethod
    def check_not_empty(cls, v):
        if not v:
            raise ValueError('字段列表不能为空')
        return [item.strip() for item in v if item and item.strip()]

class CompareFieldMatchParams(BaseModel):
    match_field: str = Field(..., description="匹配字段名")
    sheet1: Optional[str] = None
    sheet2: Optional[str] = None

    @field_validator('match_field')
    @classmethod
    def not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('匹配字段名不能为空')
        return v.strip()
    ```

- **schemas/filter.py**
  - language: `python`
  - content:
    ```python
# schemas/filter.py
from pydantic import BaseModel, Field, field_validator
from typing import List, Optional

class FilterCondition(BaseModel):
    field: str
    values: List[str]

    @field_validator('field')
    @classmethod
    def field_not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('字段名不能为空')
        return v.strip()

    @field_validator('values')
    @classmethod
    def values_not_empty(cls, v):
        if not v:
            raise ValueError('值列表不能为空')
        return [val.strip() for val in v if val.strip()]

class FilterParams(BaseModel):
    files: List[str] = Field(..., min_length=1, description="文件路径列表")
    conditions: List[FilterCondition] = Field(default_factory=list)
    fuzzy_field: Optional[str] = None
    fuzzy_value: Optional[str] = None
    action: str = Field(..., description="操作类型: get_values 或 generate_excel")
    target_field: Optional[str] = None

    @field_validator('action')
    @classmethod
    def validate_action(cls, v):
        if v not in ('get_values', 'generate_excel'):
            raise ValueError('action 必须是 get_values 或 generate_excel')
        return v

    @field_validator('target_field')
    @classmethod
    def validate_target_field(cls, v, info):
        if info.data.get('action') == 'get_values' and not v:
            raise ValueError('获取值时必须指定 target_field')
        return v
    ```

- **services/__init__.py**
  - language: `python`
  - content:
    ```python
"""业务领域服务层（精简导出）—— 仅保留最常用的服务，其余请按需显式导入"""
from services.data.reader import DataReader
from services.data.writer import DataWriter
from services.data.formatter import DataFormatter
from services.data.file_service import FileService
from services.excel.compute_engine import ComputeEngine
from core.filter_engine import FilterEngine
from services.excel.aggregator import Aggregator

__all__ = [
    'DataReader', 'DataWriter', 'DataFormatter', 'FileService',
    'ComputeEngine', 'FilterEngine', 'Aggregator',
]
    ```

- **services/checker/__init__.py**
  - language: `python`
  - content:
    ```python
"""文件校验服务"""
from .service import CheckerService
    ```

- **services/checker/constants.py**
  - language: `python`
  - content:
    ```python
# services/checker/constants.py
"""校验模块常量，避免硬编码"""

SHEET_NAMES = {
    'DIFF': '差异数据',
    'ONLY_STD': '标准群独有',
    'ONLY_CMP': '比较群独有'
}
    ```

- **services/checker/service.py**
  - language: `python`
  - content:
    ```python
# services/checker/service.py
import os
import logging
from typing import List, Dict, Optional, Callable
import pandas as pd

from core.progress_task import ProgressTask
from core.exceptions import ValidationError
from services.data.row_counter import RowCounter
from .strategies import MemoryCheckerStrategy, SqliteCheckerStrategy

logger = logging.getLogger(__name__)


class CheckerService(ProgressTask):
    """文件校验服务（策略上下文）"""
    def __init__(self, files_a=None, files_b=None,
                 key_map_a=None, key_map_b=None,
                 compare_map=None, filters_a=None, filters_b=None,
                 duplicate_action=None, config=None,
                 progress_callback=None,
                 precheck_data=None):
        super().__init__(progress_callback)
        self.files_a = files_a or []
        self.files_b = files_b or []
        self.key_map_a = key_map_a or []
        self.key_map_b = key_map_b or []
        self.compare_map = compare_map or []
        self.filters_a = filters_a or []
        self.filters_b = filters_b or []
        self.config = config
        self.duplicate_action = duplicate_action or (config.duplicate_action if config else "error")
        self.precheck_data = precheck_data

    def precheck(self) -> dict:
        ta = sum(RowCounter.count(f) for f in self.files_a)
        tb = sum(RowCounter.count(f) for f in self.files_b)
        total = ta + tb
        total_size = sum(os.path.getsize(f) for f in self.files_a + self.files_b) / (1024 * 1024)
        return {
            "total_files_a": len(self.files_a),
            "total_files_b": len(self.files_b),
            "total_rows_a": ta, "total_rows_b": tb,
            "total_rows": total,
            "estimated_memory_mb": round(total * 200 / (1024*1024), 2),
            "estimated_time_seconds": round(total / 20000, 2),
            "total_size_mb": round(total_size, 2),
            "advice": "数据量大" if total > 500000 else "安全",
            "can_proceed": True
        }

    def execute(self) -> str:
        total_rows_a = self.precheck_data.get('total_rows_a', 0) if self.precheck_data else sum(RowCounter.count(f) for f in self.files_a)
        max_memory_rows = self.config.max_memory_rows if self.config else 200000
        use_sqlite = total_rows_a > max_memory_rows

        strategy_class = SqliteCheckerStrategy if use_sqlite else MemoryCheckerStrategy
        strategy = strategy_class(
            files_a=self.files_a,
            files_b=self.files_b,
            key_map_a=self.key_map_a,
            key_map_b=self.key_map_b,
            compare_map=self.compare_map,
            filters_a=self.filters_a,
            filters_b=self.filters_b,
            duplicate_action=self.duplicate_action,
            config=self.config,
            progress_callback=self.progress_callback,
            precheck_data=self.precheck_data
        )
        strategy._total_rows = self._total_rows
        strategy._processed_rows = self._processed_rows
        return strategy.execute()
    ```

- **services/checker/strategies.py**
  - language: `python`
  - content:
    ```python
# services/checker/strategies.py
import os
import time
import logging
from typing import List, Dict, Optional, Callable, Set
import pandas as pd
from openpyxl import Workbook

from core.progress_task import ProgressTask
from core.exceptions import ValidationError
from services.data.stream_reader import StreamReader
from services.data.row_counter import RowCounter
from core.sqlite_buffer import SqliteBuffer
from services.data.report_writer import ReportWriter

logger = logging.getLogger(__name__)
BATCH_SIZE = 10000


class BaseCheckerStrategy(ProgressTask):
    """校验策略基类"""
    def __init__(self, files_a, files_b, key_map_a, key_map_b,
                 compare_map, filters_a, filters_b, duplicate_action, config,
                 progress_callback=None, precheck_data=None):
        super().__init__(progress_callback)
        self.files_a = files_a or []
        self.files_b = files_b or []
        self.key_map_a = key_map_a or []
        self.key_map_b = key_map_b or []
        self.compare_map = compare_map or []
        self.filters_a = filters_a or []
        self.filters_b = filters_b or []
        self.config = config
        self.duplicate_action = duplicate_action or (config.duplicate_action if config else "error")
        self.precheck_data = precheck_data

    def _make_key(self, row, key_cols):
        values = []
        for k in key_cols:
            try:
                value = row[k]
            except Exception:
                value = row.get(k) if isinstance(row, dict) else None
            values.append(value)
        return tuple(values)


class MemoryCheckerStrategy(BaseCheckerStrategy):
    """内存流式对比策略"""
    def execute(self) -> str:
        start_time = time.time()
        if self.precheck_data:
            total_rows_a = self.precheck_data.get('total_rows_a', 0)
            total_rows_b = self.precheck_data.get('total_rows_b', 0)
        else:
            total_rows_a = sum(RowCounter.count(f) for f in self.files_a)
            total_rows_b = sum(RowCounter.count(f) for f in self.files_b)
        self._total_rows = total_rows_a + total_rows_b

        self.set_progress(0, "读取标准群...")

        needed_a = set(self.key_map_a)
        for m in self.compare_map:
            needed_a.add(m["standard"])
        for f in self.filters_a:
            if "field" in f:
                needed_a.add(f["field"])
        needed_a.discard("")

        std_rows = []
        processed_a = 0
        for file in self.files_a:
            for row in StreamReader.read_excel(file, list(needed_a), self.filters_a):
                std_rows.append(row)
                processed_a += 1
                if processed_a % 5000 == 0:
                    pct = min(int(processed_a / total_rows_a * 100), 30) if total_rows_a else 0
                    self.set_progress(pct, f"读取标准群 {processed_a}/{total_rows_a}")
        std_df = pd.DataFrame(std_rows)
        if std_df.empty:
            raise ValidationError("标准群无有效数据")

        if self.duplicate_action == "error":
            if std_df.duplicated(subset=self.key_map_a).any():
                raise ValidationError("标准群主键重复")
        else:
            std_df = std_df.drop_duplicates(subset=self.key_map_a, keep='first')

        std_df = std_df.set_index(self.key_map_a)

        cmp_cols_a = [m["standard"] for m in self.compare_map]
        cmp_cols_b = [m["compare"] for m in self.compare_map]

        wb = Workbook(write_only=True)
        ws_diff = wb.create_sheet("差异数据")
        ws_std_only = wb.create_sheet("标准群独有")
        ws_cmp_only = wb.create_sheet("比较群独有")

        diff_header = (list(self.key_map_a) +
                       [f"{col}_标准" for col in cmp_cols_a] +
                       [f"{col}_比较" for col in cmp_cols_b])
        ws_diff.append(diff_header)
        ws_std_only.append(list(self.key_map_a) + cmp_cols_a)
        ws_cmp_only.append(list(self.key_map_b) + cmp_cols_b)

        matched_std_keys: Set[tuple] = set()

        self.set_progress(30, "开始流式对比...")

        needed_b = set(self.key_map_b)
        for m in self.compare_map:
            needed_b.add(m["compare"])
        for f in self.filters_b:
            if "field" in f:
                needed_b.add(f["field"])
        needed_b.discard("")

        batch_buffer = []
        processed_b = 0
        cmp_only_rows = []

        for file in self.files_b:
            for row in StreamReader.read_excel(file, list(needed_b), self.filters_b):
                batch_buffer.append(row)
                processed_b += 1
                if len(batch_buffer) >= BATCH_SIZE:
                    self._process_batch(batch_buffer, std_df, cmp_cols_a, cmp_cols_b,
                                        ws_diff, matched_std_keys, cmp_only_rows)
                    batch_buffer = []
                    pct = 30 + int((processed_b / total_rows_b) * 65) if total_rows_b else 50
                    self.set_progress(pct, f"对比中 {processed_b}/{total_rows_b}")

        if batch_buffer:
            self._process_batch(batch_buffer, std_df, cmp_cols_a, cmp_cols_b,
                                ws_diff, matched_std_keys, cmp_only_rows)

        for row in cmp_only_rows:
            ws_cmp_only.append(list(row))

        self.set_progress(96, "计算标准群独有...")
        all_std_keys = set(self._make_key(row, self.key_map_a) for row in std_df.reset_index().to_dict('records'))
        only_std_keys = all_std_keys - matched_std_keys
        if only_std_keys:
            std_only_df = std_df.reset_index()
            std_only_df['_key'] = std_only_df.apply(lambda r: self._make_key(r, self.key_map_a), axis=1)
            for _, row in std_only_df[std_only_df['_key'].isin(only_std_keys)].iterrows():
                values = [row[k] for k in self.key_map_a] + [row[col] for col in cmp_cols_a]
                ws_std_only.append(values)

        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        out_path = tmp.name
        wb.save(out_path)
        wb.close()

        self.set_progress(100, "完成")
        return out_path

    def _process_batch(self, batch_rows, std_df, cmp_cols_a, cmp_cols_b,
                       ws_diff, matched_std_keys, cmp_only_buf):
        cmp_df = pd.DataFrame(batch_rows)
        cmp_df['_key'] = cmp_df.apply(lambda r: self._make_key(r, self.key_map_b), axis=1)

        merged = pd.merge(
            std_df.reset_index(), cmp_df,
            left_on=self.key_map_a, right_on=self.key_map_b,
            how='inner', suffixes=('_标准', '_比较')
        )

        for _, match_row in merged.iterrows():
            matched_std_keys.add(self._make_key(match_row, self.key_map_a))

        if merged.empty:
            for _, row in cmp_df.iterrows():
                cmp_only_buf.append(row.drop(labels=['_key']).to_dict())
            return

        diff_mask = pd.Series(False, index=merged.index)
        for ca, cb in zip(cmp_cols_a, cmp_cols_b):
            left = merged[f'{ca}_标准']
            right = merged[f'{cb}_比较']
            eq = left.eq(right)
            both_null = left.isna() & right.isna()
            diff_mask |= ~(eq.fillna(False) | both_null)
        diff_rows = merged[diff_mask]

        if not diff_rows.empty:
            for _, row in diff_rows.iterrows():
                values = [row[k] for k in self.key_map_a]
                for ca in cmp_cols_a:
                    values.append(row[f'{ca}_标准'])
                for cb in cmp_cols_b:
                    values.append(row[f'{cb}_比较'])
                ws_diff.append(values)

        merged_keys = set(self._make_key(row, self.key_map_b) for _, row in merged.iterrows())
        for _, row in cmp_df.iterrows():
            if row['_key'] not in merged_keys:
                cmp_only_buf.append(row.drop(labels=['_key']).to_dict())


class SqliteCheckerStrategy(BaseCheckerStrategy):
    """SQLite 兜底对比策略"""
    def execute(self) -> str:
        start_time = time.time()
        if self.precheck_data:
            total_rows_a = self.precheck_data.get('total_rows_a', 0)
            total_rows_b = self.precheck_data.get('total_rows_b', 0)
        else:
            total_rows_a = sum(RowCounter.count(f) for f in self.files_a)
            total_rows_b = sum(RowCounter.count(f) for f in self.files_b)
        self._total_rows = total_rows_a + total_rows_b

        self.set_progress(0, "正在准备（SQLite模式）...")

        needed_a = set(self.key_map_a)
        for m in self.compare_map:
            needed_a.add(m["standard"])
        for f in self.filters_a:
            if "field" in f:
                needed_a.add(f["field"])
        needed_a.discard("")
        needed_b = set(self.key_map_b)
        for m in self.compare_map:
            needed_b.add(m["compare"])

        db = SqliteBuffer()
        try:
            db.create_table("std", list(needed_a), self.key_map_a)
            db.create_table("cmp", list(needed_b), self.key_map_b)

            self.set_progress(0, "导入标准群...")
            self._import_files(db, "std", self.files_a, list(needed_a), self.filters_a)
            self.set_progress(50, "导入比较群...")
            self._import_files(db, "cmp", self.files_b, list(needed_b), self.filters_b)

            if self.duplicate_action == "error":
                for tbl, keys, label in [("std", self.key_map_a, "标准"), ("cmp", self.key_map_b, "比较")]:
                    quoted_keys = ",".join([f'"{k}"' for k in keys])
                    sql = f"SELECT {quoted_keys} FROM {tbl} GROUP BY {quoted_keys} HAVING COUNT(*)>1"
                    if db.execute(sql).fetchone():
                        raise ValidationError(f"{label}群主键重复")
            else:
                db.remove_duplicates("std", self.key_map_a)
                db.remove_duplicates("cmp", self.key_map_b)

            cmp_a = [m["standard"] for m in self.compare_map]
            cmp_b = [m["compare"] for m in self.compare_map]
            quoted_std_keys = [f'std."{ka}"' for ka in self.key_map_a]
            quoted_cmp_keys = [f'cmp."{kb}"' for kb in self.key_map_b]
            join_on = " AND ".join([f"{a} = {b}" for a, b in zip(quoted_std_keys, quoted_cmp_keys)])
            eq_parts = [
                f'((std."{ca}" = cmp."{cb}") OR (std."{ca}" IS NULL AND cmp."{cb}" IS NULL))'
                for ca, cb in zip(cmp_a, cmp_b)
            ]
            diff_sql = f"SELECT std.*, cmp.* FROM std JOIN cmp ON {join_on} WHERE NOT ({' AND '.join(eq_parts)})"
            only_a_sql = f"SELECT std.* FROM std LEFT JOIN cmp ON {join_on} WHERE cmp.\"{self.key_map_b[0]}\" IS NULL"
            only_b_sql = f"SELECT cmp.* FROM cmp LEFT JOIN std ON {join_on} WHERE std.\"{self.key_map_a[0]}\" IS NULL"

            self.set_progress(80, "执行对比查询...")
            diff_df = db.fetch_df(diff_sql)
            only_a_df = db.fetch_df(only_a_sql)
            only_b_df = db.fetch_df(only_b_sql)

            self.set_progress(90, "写入结果...")
            sheets = {}
            if not diff_df.empty:
                sheets["差异数据"] = diff_df
            if not only_a_df.empty:
                sheets["标准群独有"] = only_a_df
            if not only_b_df.empty:
                sheets["比较群独有"] = only_b_df

            writer = ReportWriter(
                timeout_seconds=self.config.sheet2_timeout_seconds if self.config else None,
                progress_callback=lambda **kw: self.set_progress(**kw)
            )
            result_path = writer.write(sheets, start_time)
            self.set_progress(100, "完成")
            return result_path
        finally:
            db.close()

    def _import_files(self, db, table, files, columns, filters=None):
        batch = []
        batch_size = 5000
        for file in files:
            for row in StreamReader.read_excel(file, columns, filters):
                batch.append([row[c] for c in columns])
                self._processed_rows += 1
                if len(batch) >= batch_size:
                    db.insert_batch(table, columns, batch)
                    batch.clear()
            if batch:
                db.insert_batch(table, columns, batch)
                batch.clear()
    ```

- **services/common/excel_styles.py**
  - language: `python`
  - content:
    ```python
# services/common/excel_styles.py
"""公共 Excel 样式定义，避免硬编码"""
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# 字体
TITLE_FONT = Font(name='微软雅黑', bold=True, size=14)
HEADER_FONT_BOLD = Font(name='微软雅黑', bold=True, size=10)
HEADER_FONT_RED = Font(name='微软雅黑', bold=True, size=10, color="FF0000")
NORMAL_FONT = Font(name='微软雅黑', size=10)
BOLD_FONT = Font(name='微软雅黑', bold=True, size=10)
NEGATIVE_FONT = Font(name='微软雅黑', size=10, color="FF0000")

# 填充
GRAY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
GREEN_FILL = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
RED_FILL = PatternFill(start_color="F4B7BE", end_color="F4B7BE", fill_type="solid")
LIGHT_GRAY_FILL = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

# 边框
THIN_BORDER = Border(
    left=Side('thin'), right=Side('thin'),
    top=Side('thin'), bottom=Side('thin')
)

# 对齐
CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
    ```

- **services/convert/__init__.py**
  - language: `python`
  - content:
    ```python
"""文件格式转换服务"""
from .service import ConvertService
    ```

- **services/convert/service.py**
  - language: `python`
  - content:
    ```python
# services/convert/service.py
import os, tempfile, zipfile, csv, logging
import pandas as pd
from typing import Optional, Callable
from core.exceptions import ValidationError
from core.task.base import Task

logger = logging.getLogger(__name__)

class ConvertService(Task):
    def __init__(self, file_path: str, target_format: str = 'csv',
                 encoding: str = 'utf-8-sig', separator: str = ',',
                 progress_callback: Optional[Callable] = None):
        self.file_path = file_path
        self.target_format = target_format
        self.encoding = encoding
        self.separator = separator
        self.progress_callback = progress_callback

    def precheck(self) -> dict:
        ext = os.path.splitext(self.file_path)[1].lower()
        if ext in ('.xlsx', '.xls', '.xlsm'):
            from openpyxl import load_workbook
            wb = load_workbook(self.file_path, read_only=True)
            rows = sum(ws.max_row - 1 for ws in wb.worksheets)
            wb.close()
            return {"type": "excel", "sheets": len(wb.sheetnames), "total_rows": rows}
        else:
            size = os.path.getsize(self.file_path)
            return {"type": "csv", "total_rows": max(size // 200, 1)}

    def execute(self) -> str:
        ext = os.path.splitext(self.file_path)[1].lower()
        if self.target_format == 'csv' and ext in ('.xlsx', '.xls', '.xlsm'):
            return self._excel_to_csv()
        elif self.target_format == 'xlsx' and ext in ('.csv',):
            return self._csv_to_excel()
        else:
            raise ValidationError(f"不支持的转换: {ext} -> {self.target_format}")

    def _excel_to_csv(self) -> str:
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path, read_only=True)
        sheet_names = wb.sheetnames
        if len(sheet_names) == 1:
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            total_rows = ws.max_row - 1
            with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w', encoding=self.encoding, newline='') as f:
                writer = csv.writer(f, delimiter=self.separator)
                writer.writerow(headers)
                buffer = []
                buffer_size = 200000   # 20 万行内存缓冲
                processed = 0
                self._update_progress(0, f"开始转换 (共约 {total_rows:,} 行)...")
                for row in ws.iter_rows(min_row=2, values_only=True):
                    buffer.append(row)
                    processed += 1
                    if len(buffer) >= buffer_size:
                        writer.writerows(buffer)
                        buffer.clear()
                        # 只在大块完成时更新进度
                        pct = int(processed / total_rows * 100) if total_rows else 50
                        self._update_progress(pct, f"已转换 {processed:,} / {total_rows:,} 行")
                # 写入剩余行
                if buffer:
                    writer.writerows(buffer)
                out_path = f.name
            wb.close()
            self._update_progress(100, "转换完成")
            return out_path
        else:
            # 多 sheet 打包不变
            zip_path = tempfile.NamedTemporaryFile(suffix='.zip', delete=False).name
            tmp_files = []
            total = len(sheet_names)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for i, name in enumerate(sheet_names):
                    self._update_progress(i/total*100, f"转换工作表 {name} ({i+1}/{total})")
                    df = pd.read_excel(self.file_path, sheet_name=name, engine='openpyxl')
                    csv_name = f"{os.path.splitext(os.path.basename(self.file_path))[0]}_{name}.csv"
                    tmp_csv = tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w', encoding=self.encoding)
                    df.to_csv(tmp_csv.name, index=False, encoding=self.encoding, sep=self.separator)
                    tmp_files.append(tmp_csv.name)
                    zf.write(tmp_csv.name, arcname=csv_name)
            for f in tmp_files:
                try: os.unlink(f)
                except: pass
            wb.close()
            self._update_progress(100, "转换完成")
            return zip_path

    def _csv_to_excel(self) -> str:
        total_size = os.path.getsize(self.file_path)
        estimated_rows = max(total_size // 100, 1)
        self._update_progress(0, "开始读取CSV...")

        chunks = []
        total_rows = 0
        for chunk in pd.read_csv(self.file_path, encoding=self.encoding, sep=self.separator, chunksize=100000):
            chunks.append(chunk)
            total_rows += len(chunk)
            pct = min(int(total_rows / estimated_rows * 100), 99)
            self._update_progress(pct, f"已读取 {total_rows:,} 行")

        df = pd.concat(chunks)
        self._update_progress(95, "正在写入Excel...")
        out_path = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
        df.to_excel(out_path, index=False, engine='openpyxl')
        self._update_progress(100, "转换完成")
        return out_path

    def _update_progress(self, percent, message):
        if self.progress_callback:
            self.progress_callback(percent=percent, message=message)
    ```

- **services/data/__init__.py**
  - language: `python`
  - content:
    ```python
"""数据读写领域服务"""
    ```

