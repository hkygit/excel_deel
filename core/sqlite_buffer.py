# core/sqlite_buffer.py
import sqlite3, tempfile, os, logging
from typing import List, Optional

logger = logging.getLogger(__name__)


class SqliteBuffer:
    """管理临时 SQLite 数据库，支持建表、建索引、批量插入、去重"""

    @staticmethod
    def _escape_identifier(name: str) -> str:
        """转义 SQLite 标识符（表名/列名/索引名）。

        SQLite 标识符用双引号包裹，内部的双引号需加倍（""）。
        拒绝包含空字节和分号的危险输入。
        """
        if not name or '\x00' in name or ';' in name:
            raise ValueError(f"非法的标识符名称: {name!r}")
        # SQLite 双引号转义：将 " 替换为 ""
        escaped = name.replace('"', '""')
        return f'"{escaped}"'

    @staticmethod
    def _safe_index_name(table: str, col: str, max_len: int = 50) -> str:
        """生成安全的索引名：清洗特殊字符并截断，再用双引号包裹。"""
        # 只保留字母、数字、下划线、中文
        safe = ''.join(c if c.isalnum() or c in ('_', '\u4e00-\u9fff') else '_' for c in f"{table}_{col}")
        safe = safe[:max_len]
        return SqliteBuffer._escape_identifier(f"idx_{safe}")

    def __init__(self, cache_size: int = -100000):
        fd, self.db_path = tempfile.mkstemp(suffix='.db')
        os.close(fd)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA synchronous = OFF")
        self.conn.execute("PRAGMA journal_mode = MEMORY")
        self.conn.execute(f"PRAGMA cache_size = {cache_size}")

    def create_table(self, table: str, columns: List[str], key_columns: List[str] = None):
        escaped_table = self._escape_identifier(table)
        col_defs = ', '.join([f'{self._escape_identifier(c)} TEXT' for c in columns])
        self.conn.execute(f"CREATE TABLE {escaped_table} ({col_defs})")
        if key_columns:
            for col in key_columns:
                escaped_col = self._escape_identifier(col)
                index_name = self._safe_index_name(table, col)
                self.conn.execute(
                    f'CREATE INDEX IF NOT EXISTS {index_name} ON {escaped_table}({escaped_col})'
                )

    def insert_batch(self, table: str, columns: List[str], rows: list):
        placeholders = ', '.join(['?' for _ in columns])
        escaped_table = self._escape_identifier(table)
        escaped_cols = ', '.join([self._escape_identifier(c) for c in columns])
        self.conn.executemany(
            f"INSERT INTO {escaped_table} ({escaped_cols}) VALUES ({placeholders})",
            rows
        )

    def remove_duplicates(self, table: str, key_columns: List[str], keep: str = 'first'):
        if keep == 'first':
            # 使用 rowid 保留第一条（按插入顺序）
            escaped_table = self._escape_identifier(table)
            escaped_cols = ', '.join([self._escape_identifier(c) for c in key_columns])
            self.conn.execute(
                f"CREATE TABLE {escaped_table}_dedup "
                f"AS SELECT * FROM {escaped_table} "
                f"WHERE rowid IN (SELECT MIN(rowid) FROM {escaped_table} GROUP BY {escaped_cols})"
            )
            self.conn.execute(f"DROP TABLE {escaped_table}")
            self.conn.execute(f"ALTER TABLE {escaped_table}_dedup RENAME TO {escaped_table}")

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