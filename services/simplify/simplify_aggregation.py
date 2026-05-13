# utils/simplify/simplify_aggregation.py  (优化版)
import sqlite3
import tempfile
import os
import pandas as pd
from typing import List, Dict, Optional
from abc import ABC, abstractmethod


class AggregationStrategy(ABC):
    @abstractmethod
    def add_row(self, row: Dict[str, str], key_cols: List[str],
                sum_fields: List[str], non_sum_fields: List[str],
                override_markers: Optional[Dict[str, List[str]]] = None):
        pass

    @abstractmethod
    def finalize(self, key_cols: List[str], sum_fields: List[str],
                 non_sum_fields: List[str]) -> pd.DataFrame:
        pass


class MemoryAggregation(AggregationStrategy):
    def __init__(self):
        self.data = {}  # key_tuple -> [sum_dict, first_dict, count]

    def add_row(self, row: Dict[str, str], key_cols: List[str],
                sum_fields: List[str], non_sum_fields: List[str],
                override_markers: Optional[Dict[str, List[str]]] = None):
        if override_markers is None:
            override_markers = {}
        # 预处理覆盖标记集合
        markers_per_col = {}
        for col, markers in override_markers.items():
            markers_per_col[col] = {str(m).strip() for m in markers} | {''}

        key = tuple(row.get(k, '') for k in key_cols)
        if key not in self.data:
            sum_dict = {col: 0.0 for col in sum_fields}
            first_dict = {col: '' for col in non_sum_fields}
            self.data[key] = [sum_dict, first_dict, 1]
        else:
            sum_dict, first_dict, cnt = self.data[key]
            self.data[key][2] = cnt + 1

        # 累加 sum_fields —— 非数值不归零，直接跳过
        for col in sum_fields:
            val = row.get(col, '')
            if val == '':
                continue
            try:
                num = float(val)
                sum_dict[col] += num
            except (ValueError, TypeError):
                # 非数值记录日志，忽略该值
                pass

        # 更新 non_sum_fields （逻辑不变）
        for col in non_sum_fields:
            current_val = row.get(col, '')
            if current_val == '':
                continue
            stored_val = first_dict.get(col, '')
            allowed_markers = markers_per_col.get(col, {''})
            if stored_val in allowed_markers:
                first_dict[col] = current_val

    def finalize(self, key_cols: List[str], sum_fields: List[str],
                 non_sum_fields: List[str]) -> pd.DataFrame:
        rows = []
        for key_tuple, (sum_dict, first_dict, cnt) in self.data.items():
            row = {}
            for i, col in enumerate(key_cols):
                row[col] = key_tuple[i]
            row.update(sum_dict)
            row.update(first_dict)
            row['原始行数'] = cnt
            rows.append(row)
        return pd.DataFrame(rows)

    def size(self) -> int:
        return len(self.data)


class SqliteAggregation(AggregationStrategy):
    def __init__(self, max_memory_keys: int):
        self.max_memory_keys = max_memory_keys
        self.conn = None
        self.cursor = None
        self.db_path = None
        self.batch_buffer = []
        self.batch_size = 5000
        self.row_count = 0
        self.override_markers = None

    def _init_db(self, key_cols: List[str], sum_fields: List[str], non_sum_fields: List[str]):
        self.db_path = tempfile.NamedTemporaryFile(suffix='.db', delete=False).name
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA synchronous = OFF")
        self.conn.execute("PRAGMA journal_mode = MEMORY")
        self.cursor = self.conn.cursor()
        cols_def = []
        for col in key_cols:
            cols_def.append(f'`{col}` TEXT')
        for col in sum_fields:
            cols_def.append(f'`{col}` REAL')
        for col in non_sum_fields:
            cols_def.append(f'`{col}` TEXT DEFAULT NULL')
        cols_def.append('`__count__` INTEGER')
        create_sql = f"CREATE TABLE agg ({', '.join(cols_def)})"
        self.cursor.execute(create_sql)
        self.conn.commit()

    def add_row(self, row: Dict[str, str], key_cols: List[str],
                sum_fields: List[str], non_sum_fields: List[str],
                override_markers: Optional[Dict[str, List[str]]] = None):
        if self.conn is None:
            self._init_db(key_cols, sum_fields, non_sum_fields)
            self.override_markers = override_markers or {}
            self._markers_per_col = {}
            for col, markers in self.override_markers.items():
                self._markers_per_col[col] = {str(m).strip() for m in markers} | {''}

        cols = key_cols + sum_fields + non_sum_fields + ['__count__']
        values = []
        for col in key_cols:
            values.append(row.get(col, ''))
        for col in sum_fields:
            val = row.get(col, '')
            if val == '':
                values.append(None)  # 空值视为 NULL，聚合时会被忽略
            else:
                try:
                    num = float(val)
                    values.append(num)
                except (ValueError, TypeError):
                    values.append(None)   # 非数值同样置 NULL
        for col in non_sum_fields:
            val = row.get(col, '')
            markers = self._markers_per_col.get(col, {''})
            if val in markers:
                values.append(None)
            else:
                values.append(val)
        values.append(1)
        self.batch_buffer.append(values)
        if len(self.batch_buffer) >= self.batch_size:
            self._flush_buffer()

    def _flush_buffer(self):
        if not self.batch_buffer:
            return
        placeholders = ','.join(['?'] * len(self.batch_buffer[0]))
        self.cursor.executemany(f"INSERT INTO agg VALUES ({placeholders})", self.batch_buffer)
        self.row_count += len(self.batch_buffer)
        self.batch_buffer = []
        self.conn.commit()

    def finalize(self, key_cols: List[str], sum_fields: List[str],
                 non_sum_fields: List[str]) -> pd.DataFrame:
        if self.conn is None:
            return pd.DataFrame()
        self._flush_buffer()
        self.conn.commit()
        select_parts = []
        seen = set()
        for col in key_cols:
            if col not in seen:
                seen.add(col)
                select_parts.append(f'`{col}`')
        for col in sum_fields:
            if col not in seen:
                seen.add(col)
                select_parts.append(f'SUM(`{col}`) AS `{col}`')
        for col in non_sum_fields:
            if col not in seen:
                seen.add(col)
                select_parts.append(f'MIN(`{col}`) AS `{col}`')
        select_parts.append('SUM(`__count__`) AS `原始行数`')
        group_by = ', '.join([f'`{col}`' for col in key_cols])
        sql = f"SELECT {', '.join(select_parts)} FROM agg GROUP BY {group_by}"
        df = pd.read_sql(sql, self.conn)
        # 将 NULL 转为空字符串（仅非数值列）
        for col in non_sum_fields:
            if col in df.columns:
                df[col] = df[col].fillna('')
        self.conn.close()
        os.unlink(self.db_path)
        return df


class AggregationStrategyFactory:
    @staticmethod
    def create(estimated_keys: int, threshold: int = 500000) -> AggregationStrategy:
        if estimated_keys < threshold:
            return MemoryAggregation()
        else:
            return SqliteAggregation(threshold)