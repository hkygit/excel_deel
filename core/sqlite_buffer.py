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
