- **services/data/file_service.py**
  - language: `python`
  - content:
    ```python
# utils/services/file_service.py
import os
import uuid
import logging
import tempfile
from typing import List, Optional
from werkzeug.utils import secure_filename
from flask import current_app

logger = logging.getLogger(__name__)

class FileService:
    """统一文件管理服务"""

    @staticmethod
    def save_upload(file, subdir: str = '', allowed_extensions: Optional[List[str]] = None) -> str:
        if not file or not file.filename:
            raise ValueError("无效文件")
        if allowed_extensions is None:
            allowed_extensions = current_app.config.get('ALLOWED_EXTENSIONS', {'xlsx', 'xls', 'csv'})

        # 分离文件名和扩展名，保留原始扩展名
        original_filename = file.filename
        name, ext = os.path.splitext(original_filename)
        # 对文件名部分进行安全处理，扩展名保留原样（不经过 secure_filename，避免点号被替换）
        safe_name = secure_filename(name) if name else 'file'
        if not safe_name:
            safe_name = 'file'
        # 确保扩展名以小写形式保存
        ext_lower = ext.lower()
        # 如果扩展名不在允许列表中，但文件实际是有效的 Excel，我们仍保留（后续会转换）
        # 这里不做严格检查，交给后续处理
        filename = f"{safe_name}{ext_lower}"

        unique_id = uuid.uuid4().hex
        base_dir = current_app.config.get('UPLOAD_FOLDER', 'uploads')
        save_dir = os.path.join(base_dir, subdir) if subdir else base_dir
        os.makedirs(save_dir, exist_ok=True)
        save_path = os.path.join(save_dir, f"{unique_id}_{filename}")
        file.save(save_path)
        logger.debug(f"文件已保存: {save_path}")
        return save_path

    @staticmethod
    def cleanup(paths: List[str]):
        for path in paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    logger.debug(f"已删除文件: {path}")
            except Exception as e:
                logger.warning(f"删除文件失败 {path}: {e}")

    @staticmethod
    def get_temp_path(suffix: str = '') -> str:
        fd, path = tempfile.mkstemp(suffix=suffix)
        os.close(fd)
        return path

    @staticmethod
    def safe_filename(name: str) -> str:
        return secure_filename(name)
    ```

- **services/data/formatter.py**
  - language: `python`
  - content:
    ```python
# services/data/formatter.py
import pandas as pd
import numpy as np
import logging

logger = logging.getLogger(__name__)

class DataFormatter:
    """统一数据格式化服务"""

    @staticmethod
    def round_numeric_columns(df: pd.DataFrame, decimals: int = 2) -> pd.DataFrame:
        df = df.copy()
        for col in df.columns:
            numeric_series = pd.to_numeric(df[col], errors='coerce')
            if numeric_series.notna().any():
                df[col] = numeric_series.apply(lambda x: round(x, decimals) if pd.notnull(x) else x)
        return df

    @staticmethod
    def safe_format_dataframe(df: pd.DataFrame) -> pd.DataFrame:
        """
        安全格式化 DataFrame：
        - 数值列：保持原样，不做转换
        - 字符串列：将 NaN/None 转为空字符串，并去除首尾空格
        - 日期时间列：转为字符串
        """
        df_copy = df.copy()
        for col in df_copy.columns:
            if pd.api.types.is_numeric_dtype(df_copy[col]):
                # 数值列：保留原始数据类型和缺失值表示
                continue
            elif pd.api.types.is_datetime64_any_dtype(df_copy[col]):
                df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d %H:%M:%S')
            else:
                # 非数值、非日期：转为字符串并清理
                df_copy[col] = df_copy[col].fillna('').astype(str).str.strip()
                df_copy[col] = df_copy[col].replace('nan', '')
        return df_copy

    @staticmethod
    def fillna_empty(df: pd.DataFrame) -> pd.DataFrame:
        return df.fillna('')
    ```

- **services/data/reader.py**
  - language: `python`
  - content:
    ```python
# utils/services/reader.py
import csv
import pandas as pd
from typing import Union, Iterator, Dict, List, Optional
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class DataReader:
    """统一数据读取服务，支持 CSV 和 Excel"""

    @staticmethod
    def read_excel(filepath: str, sheet: Union[str, int] = 0, **kwargs) -> pd.DataFrame:
        """读取 Excel 文件为 DataFrame"""
        try:
            return pd.read_excel(filepath, sheet_name=sheet, **kwargs)
        except Exception as e:
            logger.error(f"读取Excel失败 {filepath}: {e}")
            raise ValueError(f"无法读取Excel文件: {e}")

    @staticmethod
    def read_csv(filepath: str, encoding: str = 'utf-8', **kwargs) -> pd.DataFrame:
        """读取 CSV 文件为 DataFrame"""
        try:
            return pd.read_csv(filepath, encoding=encoding, **kwargs)
        except Exception as e:
            logger.error(f"读取CSV失败 {filepath}: {e}")
            raise ValueError(f"无法读取CSV文件: {e}")

    @staticmethod
    def read_auto(filepath: str, **kwargs) -> pd.DataFrame:
        """自动检测文件类型并读取"""
        ext = Path(filepath).suffix.lower()
        if ext in ('.xlsx', '.xls'):
            return DataReader.read_excel(filepath, **kwargs)
        elif ext == '.csv':
            return DataReader.read_csv(filepath, **kwargs)
        else:
            raise ValueError(f"不支持的文件类型: {ext}")
    ```

- **services/data/report_writer.py**
  - language: `python`
  - content:
    ```python
# services/data/report_writer.py
import tempfile, os, logging
from typing import Dict, Optional, Callable
import pandas as pd

logger = logging.getLogger(__name__)

class ReportWriter:
    """通用 Excel 报告写入器，支持超时跳过单个 Sheet"""
    def __init__(self, timeout_seconds: Optional[int] = None, progress_callback: Optional[Callable] = None):
        self.timeout_seconds = timeout_seconds
        self.progress_callback = progress_callback

    def write(self, sheets: Dict[str, pd.DataFrame], start_time: float) -> str:
        """sheets: {sheet_name: DataFrame}"""
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            for name, df in sheets.items():
                if df.empty:
                    continue
                # 特殊处理标准群独有超时
                if name == "标准群独有" and self.timeout_seconds:
                    import time
                    if time.time() - start_time > self.timeout_seconds:
                        logger.warning(f"跳过 {name}，已超时")
                        continue
                df.to_excel(writer, sheet_name=name, index=False)
        return tmp.name
    ```

- **services/data/row_counter.py**
  - language: `python`
  - content:
    ```python
import os
import logging
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


class RowCounter:
    """通用文件行数统计器（不含标题行）"""
    @staticmethod
    def count_excel(filepath: str) -> int:
        try:
            wb = load_workbook(filepath, read_only=True)
            rows = sum(ws.max_row - 1 for ws in wb.worksheets)
            wb.close()
            return max(rows, 0)
        except Exception as e:
            logger.warning(f"统计 Excel 行数失败 {filepath}: {e}")
            return 0

    @staticmethod
    def count_csv(filepath: str, encoding='utf-8') -> int:
        try:
            with open(filepath, 'r', encoding=encoding) as f:
                return sum(1 for _ in f) - 1
        except Exception as e:
            logger.warning(f"统计 CSV 行数失败 {filepath}: {e}")
            return 0

    @classmethod
    def count(cls, filepath: str) -> int:
        ext = os.path.splitext(filepath)[1].lower()
        if ext in ('.xlsx', '.xls', '.xlsm'):
            return cls.count_excel(filepath)
        return cls.count_csv(filepath)

    ```

- **services/data/stream_reader.py**
  - language: `python`
  - content:
    ```python
import os, logging, csv
from typing import List, Optional, Callable, Generator, Dict, Any
from core.filter_engine import ConditionFilter, Condition

logger = logging.getLogger(__name__)


class StreamReader:
    """流式读取 Excel/CSV，返回清洗后的字典行，支持通用进度回调"""

    @staticmethod
    def read_excel(
        filepath: str,
        columns: List[str],
        filters: Optional[List[Dict]] = None,
        progress_callback: Optional[Callable] = None,
        total_rows: int = 0
    ) -> Generator[Dict[str, str], None, None]:
        """
        读取 Excel（兼容旧回调）
        旧回调形式: progress_callback(processed_rows=..., total_rows=...)
        """
        from openpyxl import load_workbook
        wb = load_workbook(filepath, read_only=True)
        ws = wb.active
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        col_indices = {col: headers.index(col) for col in columns if col in headers}
        if not col_indices:
            wb.close()
            raise ValueError(f"文件 {filepath} 缺少必要列: {columns}")
        filter_conditions = [Condition(**f) for f in filters] if filters else []
        processed = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            d = {}
            for col in columns:
                if col in col_indices:
                    val = row[col_indices[col]]
                    if isinstance(val, str):
                        val = val.strip()
                    d[col] = val
                else:
                    d[col] = None
            if filter_conditions:
                prepared = d.copy()
                for f in filters:
                    if f.get('op') in ('>','<','>=','<='):
                        try:
                            prepared[f['field']] = float(prepared[f['field']])
                        except Exception:
                            pass
                if not ConditionFilter.apply_to_dict(prepared, filter_conditions):
                    processed += 1
                    if progress_callback:
                        progress_callback(processed_rows=processed, total_rows=total_rows)
                    continue
            processed += 1
            if progress_callback:
                progress_callback(processed_rows=processed, total_rows=total_rows)
            yield d
        wb.close()

    @classmethod
    def read_excel_with_progress(
        cls,
        filepath: str,
        columns: List[str],
        filters: Optional[List[Dict]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None,
        total_rows: int = 0
    ) -> Generator[Dict[str, str], None, None]:
        """
        读取 Excel，支持 ProgressTask 风格的进度回调
        progress_callback(percent, message)
        """
        processed = 0
        total = total_rows if total_rows else 1

        def old_callback(processed_rows, total_rows):
            nonlocal processed
            processed = processed_rows
            if progress_callback:
                pct = min(int(processed / total * 100), 99) if total else 50
                msg = f"已读取 {processed}/{total}"
                progress_callback(pct, msg)

        yield from cls.read_excel(filepath, columns, filters, old_callback, total)
        if progress_callback:
            progress_callback(100, "读取完成")

    ```

- **services/data/writer.py**
  - language: `python`
  - content:
    ```python
# utils/services/writer.py
import pandas as pd
from pathlib import Path
import logging
from typing import Optional

logger = logging.getLogger(__name__)

class DataWriter:
    """统一数据写入服务"""

    @staticmethod
    def write_excel(df: pd.DataFrame, filepath: str, sheet_name: str = 'Sheet1',
                    index: bool = False, engine: str = 'openpyxl') -> None:
        """
        将 DataFrame 写入 Excel 文件
        :param df: 要写入的数据
        :param filepath: 目标文件路径
        :param sheet_name: 工作表名称
        :param index: 是否写入行索引
        :param engine: Excel 引擎，默认 openpyxl
        """
        try:
            df.to_excel(filepath, sheet_name=sheet_name, index=index, engine=engine)
            logger.debug(f"已写入 Excel 文件: {filepath}")
        except Exception as e:
            logger.error(f"写入 Excel 失败 {filepath}: {e}")
            raise ValueError(f"无法写入 Excel 文件: {e}")

    @staticmethod
    def write_csv(df: pd.DataFrame, filepath: str, encoding: str = 'utf-8-sig',
                  index: bool = False, **kwargs) -> None:
        """
        将 DataFrame 写入 CSV 文件
        :param df: 要写入的数据
        :param filepath: 目标文件路径
        :param encoding: 编码格式
        :param index: 是否写入行索引
        :param kwargs: 其他 pandas to_csv 参数
        """
        try:
            df.to_csv(filepath, encoding=encoding, index=index, **kwargs)
            logger.debug(f"已写入 CSV 文件: {filepath}")
        except Exception as e:
            logger.error(f"写入 CSV 失败 {filepath}: {e}")
            raise ValueError(f"无法写入 CSV 文件: {e}")

    @staticmethod
    def write_auto(df: pd.DataFrame, filepath: str, **kwargs) -> None:
        """
        根据文件扩展名自动选择写入格式
        :param df: 要写入的数据
        :param filepath: 目标文件路径
        :param kwargs: 传递给具体写入方法的参数
        """
        ext = Path(filepath).suffix.lower()
        if ext in ('.xlsx', '.xls'):
            DataWriter.write_excel(df, filepath, **kwargs)
        elif ext == '.csv':
            DataWriter.write_csv(df, filepath, **kwargs)
        else:
            raise ValueError(f"不支持的文件类型: {ext}")

    @staticmethod
    def write_excel_with_formatting(df: pd.DataFrame, filepath: str,
                                    sheet_name: str = 'Sheet1',
                                    float_format: Optional[str] = None,
                                    **kwargs) -> None:
        """
        写入 Excel 并支持格式化选项
        :param float_format: 浮点数格式化字符串，如 '%.2f'
        """
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False,
                            float_format=float_format, **kwargs)
            logger.debug(f"已写入带格式的 Excel 文件: {filepath}")
        except Exception as e:
            logger.error(f"写入带格式的 Excel 失败 {filepath}: {e}")
            raise
    ```

- **services/excel/__init__.py**
  - language: `python`
  - content:
    ```python
"""Excel处理领域服务"""
    ```

- **services/excel/aggregator.py**
  - language: `python`
  - content:
    ```python
# services/excel/aggregator.py
"""通用聚合器（纯静态方法）。"""
import pandas as pd
from typing import List, Dict, Union, Optional


class Aggregator:
    """通用聚合器（静态方法）。"""

    @staticmethod
    def group_aggregate(df: pd.DataFrame, group_keys: List[str],
                        agg_spec: Dict[str, Union[str, List[str]]]) -> pd.DataFrame:
        if not group_keys:
            return df.drop_duplicates()
        return df.groupby(group_keys, as_index=False).agg(agg_spec)

    @staticmethod
    def aggregate_by_keys(df: pd.DataFrame, unique_keys: List[str],
                          sum_fields: List[str],
                          other_fields: Optional[List[str]] = None) -> pd.DataFrame:
        """按唯一键聚合：sum_fields 求和，other_fields 取第一条记录。

        实现上组合 drop_duplicates 与 groupby sum，避免 pandas 在某些
        版本上对 agg(dict) 的内部错误。
        """
        if other_fields is None:
            other_fields = [
                c for c in df.columns
                if c not in unique_keys and c not in sum_fields
            ]
        if not unique_keys:
            return df.drop_duplicates()

        # 1. 求和部分
        if sum_fields:
            sum_part = df.groupby(unique_keys, as_index=False)[sum_fields].sum()
        else:
            sum_part = pd.DataFrame(columns=unique_keys)

        # 2. 取第一条记录部分
        if other_fields:
            first_part = df[unique_keys + other_fields].drop_duplicates(
                subset=unique_keys, keep='first'
            )
        else:
            first_part = pd.DataFrame(columns=unique_keys)

        # 3. 合并
        if not sum_part.empty and not first_part.empty:
            result = sum_part.merge(first_part, on=unique_keys, how='outer')
        elif not sum_part.empty:
            result = sum_part
        elif not first_part.empty:
            result = first_part
        else:
            result = df[unique_keys].drop_duplicates()

        return result

    @staticmethod
    def simple_aggregate(df: pd.DataFrame, group_keys: List[str],
                         agg_cols: List[str], func: str = 'sum') -> pd.DataFrame:
        agg_dict = {col: func for col in agg_cols}
        return Aggregator.group_aggregate(df, group_keys, agg_dict)
    ```

- **services/excel/calculator.py**
  - language: `python`
  - content:
    ```python
# services/excel/calculator.py
import numpy as np
from typing import List, Tuple
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment
from core.excel_utils import ExcelFile
from core.expression_engine import get_expression_engine, extract_column_letters

class ColumnCalculator:
    def __init__(self, excel_file: ExcelFile):
        self.excel = excel_file
        self.ws = excel_file.active_sheet

    def get_data_range(self, header_row: int = 1) -> Tuple[int, int]:
        return header_row + 1, self.ws.max_row

    def calculate_column(self, expression: str, result_col: str,
                         merge_by_cols: List[str] = None) -> None:
        engine = get_expression_engine()
        col_letters = extract_column_letters(expression)
        if not col_letters:
            raise ValueError("表达式中未包含列字母")

        try:
            result_col_idx = column_index_from_string(result_col)
            is_new_col = False
        except:
            is_new_col = True
            result_col_name = result_col.strip()

        if not is_new_col:
            result_col_idx = column_index_from_string(result_col)
        else:
            max_col = self.ws.max_column
            result_col_idx = max_col + 1
            self.ws.insert_cols(result_col_idx)
            self.ws.cell(1, result_col_idx).value = result_col_name

        start_row, end_row = self.get_data_range()

        col_data = {}
        for col_letter in col_letters:
            col_idx = column_index_from_string(col_letter)
            vals = []
            for row in range(start_row, end_row + 1):
                val = self.excel.get_cell_value(row, col_idx, sheet=self.ws)
                try:
                    val = float(val) if val is not None else 0
                except:
                    val = 0
                vals.append(val)
            col_data[col_letter] = vals

        calc_vals = []
        first_letter = list(col_data.keys())[0]
        for i in range(len(col_data[first_letter])):
            row_vals = {letter: col_data[letter][i] for letter in col_letters}
            result = engine.evaluate(expression, row_vals)
            calc_vals.append(result)

        self._clear_merges_in_column(result_col_idx)
        for i, val in enumerate(calc_vals):
            row = start_row + i
            self.ws.cell(row, result_col_idx).value = None if np.isnan(val) else val

        ref_col = sorted(col_letters)[0]
        ref_idx = column_index_from_string(ref_col)
        self.excel.copy_merge_structure(ref_idx, result_col_idx, start_row, end_row, sheet=self.ws)

        if merge_by_cols:
            self._merge_by_keys(merge_by_cols, result_col_idx, start_row, end_row)

        for row in range(start_row, end_row + 1):
            self.ws.cell(row, result_col_idx).alignment = Alignment(horizontal='right')

    def _clear_merges_in_column(self, col_idx: int):
        to_remove = []
        for merged in self.ws.merged_cells.ranges:
            if merged.min_col <= col_idx <= merged.max_col:
                to_remove.append(merged)
        for merged in to_remove:
            self.ws.unmerge_cells(str(merged))

    def _merge_by_keys(self, merge_cols, result_col_idx, start_row, end_row):
        merge_indices = [column_index_from_string(c.strip()) for c in merge_cols]
        group_first_row = {}
        for row in range(start_row, end_row + 1):
            key = tuple(self.excel.get_cell_value(row, col, sheet=self.ws) for col in merge_indices)
            if key not in group_first_row:
                group_first_row[key] = row

        for row in range(start_row, end_row + 1):
            key = tuple(self.excel.get_cell_value(row, col, sheet=self.ws) for col in merge_indices)
            first_row = group_first_row.get(key)
            if row != first_row:
                self.ws.cell(row, result_col_idx).value = None

        for key, first_row in group_first_row.items():
            last_row = first_row
            for row in range(first_row, end_row + 1):
                if tuple(self.excel.get_cell_value(row, col, sheet=self.ws) for col in merge_indices) == key:
                    last_row = row
                else:
                    break
            if first_row < last_row:
                self.ws.merge_cells(start_row=first_row, start_column=result_col_idx,
                                    end_row=last_row, end_column=result_col_idx)
    ```

- **services/excel/calculator_processor.py**
  - language: `python`
  - content:
    ```python
"""
Excel 列计算与聚合回填处理器
修复版本：解决文件锁定和读取问题
"""
import os
import gc
import logging
import numpy as np
import openpyxl
import pandas as pd
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional, Tuple
import tempfile
import atexit

from core.excel_utils import ExcelFile, copy_merge_structure
from core.data_access_layer import get_data_access_layer
from core.csv_utils import build_column_mapping
from core.excel_helpers import (
    safe_get_cell_value, safe_write_cell, clear_merges_in_column,
    build_column_name_mapping, normalize_excel_file, safe_cleanup_dataframe,
    safe_delete_file
)

logger = logging.getLogger(__name__)


class BaseExcelProcessor(ABC):
    """Excel 处理器基类，封装公共的写入逻辑"""

    def __init__(self, filepath: str, header_row: int, compute_engine,
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

    def _cleanup_all(self):
        """清理所有临时文件"""
        for temp_file in self.temp_files[:]:  # 使用副本遍历
            if temp_file and os.path.exists(temp_file):
                try:
                    safe_delete_file(temp_file)
                    if temp_file in self.temp_files:
                        self.temp_files.remove(temp_file)
                except Exception as e:
                    logger.debug(f"清理临时文件失败(忽略): {temp_file}, {e}")

    def _get_work_filepath(self) -> str:
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

    def _cleanup_temp(self):
        """清理临时文件"""
        for temp_file in self.temp_files[:]:
            if temp_file and os.path.exists(temp_file) and temp_file != self.original_filepath:
                try:
                    safe_delete_file(temp_file)
                    if temp_file in self.temp_files:
                        self.temp_files.remove(temp_file)
                except Exception as e:
                    logger.debug(f"清理临时文件失败(忽略): {temp_file}, {e}")

    def _resolve_result_column(self, ws, result_col_input: str) -> Tuple[int, str]:
        """
        解析结果列：支持列字母或列名。
        返回 (列索引, 列名)
        """
        if not result_col_input or not result_col_input.strip():
            raise ValueError("结果列不能为空")

        result_col_input = result_col_input.strip()

        # 1. 尝试作为列字母解析
        try:
            col_idx = openpyxl.utils.column_index_from_string(result_col_input.upper())
            if col_idx > ws.max_column:
                # 如果列索引超过现有列，插入新列
                ws.insert_cols(col_idx)
            header_cell = ws.cell(self.header_row, col_idx)
            col_name = str(header_cell.value).strip() if header_cell.value else result_col_input
            return col_idx, col_name
        except Exception:
            pass

        # 2. 尝试作为列名查找
        name_to_idx = build_column_name_mapping(ws, self.header_row)
        if result_col_input in name_to_idx:
            col_idx = name_to_idx[result_col_input]
            return col_idx, result_col_input

        # 3. 新列：追加到末尾
        new_col_idx = ws.max_column + 1
        ws.insert_cols(new_col_idx)
        safe_write_cell(ws, self.header_row, new_col_idx, result_col_input)
        return new_col_idx, result_col_input

    def _write_result_column(self, ws, result_col_idx: int, values: List[Any],
                             start_row: int, end_row: int) -> None:
        """将计算结果写入指定列，并清理原有合并单元格"""
        clear_merges_in_column(ws, result_col_idx)

        # 确保values长度与行数匹配
        expected_rows = end_row - start_row + 1
        if len(values) != expected_rows:
            logger.warning(f"值数量({len(values)})与行数({expected_rows})不匹配，进行截断或填充")
            if len(values) < expected_rows:
                values = values + [None] * (expected_rows - len(values))
            else:
                values = values[:expected_rows]

        for i, val in enumerate(values):
            row = start_row + i
            if isinstance(val, (np.integer, np.floating)):
                val = val.item()
            safe_write_cell(ws, row, result_col_idx, val)

    def _copy_merge_structure(self, ws, src_col_letter: str, dst_col_idx: int,
                              start_row: int, end_row: int) -> None:
        """复制源列的合并结构到目标列"""
        try:
            src_col_idx = openpyxl.utils.column_index_from_string(src_col_letter)
        except Exception:
            return
        copy_merge_structure(ws, src_col_idx, ws, dst_col_idx, start_row, end_row)

    def _get_output_path(self) -> str:
        """生成输出文件路径"""
        base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
        output_dir = os.path.dirname(self.original_filepath)
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"{base_name}_处理结果_{timestamp}.xlsx")

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        return output_path

    @abstractmethod
    def process(self) -> str:
        pass


class ColumnCalculatorProcessor(BaseExcelProcessor):
    """列计算处理器"""

    def __init__(self, filepath: str, header_row: int, expression: str,
                 result_col_input: str, merge_by_cols_str: str, compute_engine,
                 already_normalized: bool = False):
        super().__init__(filepath, header_row, compute_engine,
                         already_normalized=already_normalized)
        self.expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols = [c.strip() for c in merge_by_cols_str.split(',') if c.strip()] if merge_by_cols_str else []

    def _remove_result_column_from_expr(self, result_col_letter: str) -> str:
        """从表达式中移除与结果列字母"完全匹配"的标识符 token。

        使用 token 级匹配（而不是 str.replace 子串替换），避免破坏那些
        只是"包含"结果列字母的标识符。例如结果列为 'A' 时：
            'A+B'   -> 'B'        移除独立的 'A'
            'AA+B'  -> 'AA+B'     'AA' 是独立标识符，不受影响
            'AB+A'  -> 'AB'       仅移除独立的 'A'，不动 'AB'
            'A+B-C' -> 'A-C'      移除 'B' 及其邻近的二元运算符

        当某个匹配的标识符被移除时，会一并移除它邻近的二元运算符，避免
        残留 '(A+)*C'、'A+-C' 这类非法表达式。
        """
        if not result_col_letter or not self.expression:
            return self.expression

        target = result_col_letter.upper()
        expr_upper = self.expression.upper()

        # 分词：标识符（首字符字母）、数字（含小数）、运算符、括号
        import re
        token_re = re.compile(r'[A-Z][A-Z0-9]*|\d+(?:\.\d+)?|[()+\-*/]')
        tokens = token_re.findall(expr_upper)

        # 快速路径：没有与结果列完全相等的标识符 token，无需移除
        if not any(t == target and t.isalpha() for t in tokens):
            return self.expression

        binary_ops = {'+', '-', '*', '/'}

        # 标记需要移除的 token 下标
        to_remove = set()
        n = len(tokens)
        for i, tok in enumerate(tokens):
            if tok == target and tok.isalpha():
                to_remove.add(i)
                # 优先连同前面的二元运算符一起移除
                if i > 0 and tokens[i - 1] in binary_ops:
                    to_remove.add(i - 1)
                # 否则，连同后面的二元运算符一起移除
                elif i + 1 < n and tokens[i + 1] in binary_ops:
                    to_remove.add(i + 1)

        remaining = [tok for i, tok in enumerate(tokens) if i not in to_remove]

        # 清理首尾残留的运算符
        while remaining and remaining[0] in binary_ops:
            remaining.pop(0)
        while remaining and remaining[-1] in binary_ops:
            remaining.pop()

        new_expr = ''.join(remaining).strip()
        if not new_expr:
            new_expr = '0'

        if new_expr != expr_upper:
            logger.warning(
                f"表达式中包含了结果列 {target}，已自动移除: "
                f"{self.expression} -> {new_expr}"
            )

        return new_expr

    def process(self) -> str:
        # 使用缓存的或转换后的工作文件路径
        filepath = self._get_work_filepath()

        if not os.path.exists(filepath):
            raise FileNotFoundError(f"工作文件不存在: {filepath}")

        try:
            # 以字符串类型读取，避免公式错误
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str, na_filter=False)
            df = safe_cleanup_dataframe(df, auto_convert_numeric=True)

            if df.empty:
                raise ValueError("Excel文件为空或没有数据")

            headers = list(df.columns)
            letter_to_name, name_to_letter = build_column_mapping(headers)

            with ExcelFile(filepath) as ef:
                ws = ef.active_sheet
                result_col_idx, _ = self._resolve_result_column(ws, self.result_col_input)

                result_col_letter = None
                try:
                    result_col_letter = openpyxl.utils.get_column_letter(result_col_idx)
                except Exception:
                    pass

                expr = self.expression
                if result_col_letter:
                    expr = self._remove_result_column_from_expr(result_col_letter)
                    if not expr or not expr.strip():
                        raise ValueError(f"表达式在移除结果列 {result_col_letter} 后为空")

                variables = self.compute_engine.extract_variables(expr)
                mapped_vars = {}
                missing = []
                for letter in variables:
                    if letter in letter_to_name:
                        mapped_vars[letter] = letter_to_name[letter]
                    else:
                        missing.append(letter)
                if missing:
                    raise ValueError(f"表达式中的列字母 {missing} 无法映射到文件中的任何列")

                def calc_row(row):
                    context = {letter: row[col_name] for letter, col_name in mapped_vars.items()}
                    return self.compute_engine.evaluate_expression(expr, context)

                df["Result"] = df.apply(calc_row, axis=1)

                start_row = self.header_row + 1
                end_row = ws.max_row
                values = [None if pd.isna(v) else v for v in df["Result"].tolist()]

                self._write_result_column(ws, result_col_idx, values, start_row, end_row)

                if variables:
                    ref_col_letter = sorted(variables)[0]
                    self._copy_merge_structure(ws, ref_col_letter, result_col_idx, start_row, end_row)

                if self.merge_by_cols:
                    merge_indices = []
                    for col in self.merge_by_cols:
                        try:
                            merge_indices.append(openpyxl.utils.column_index_from_string(col))
                        except Exception:
                            if col in name_to_letter:
                                merge_indices.append(openpyxl.utils.column_index_from_string(name_to_letter[col]))
                            else:
                                raise ValueError(f"合并依据列 {col} 无效")

                    group_first_row = {}
                    for row in range(start_row, end_row + 1):
                        key = tuple(safe_get_cell_value(ws, row, col) for col in merge_indices)
                        if key not in group_first_row:
                            group_first_row[key] = row

                    for row in range(start_row, end_row + 1):
                        key = tuple(safe_get_cell_value(ws, row, col) for col in merge_indices)
                        if row != group_first_row.get(key):
                            safe_write_cell(ws, row, result_col_idx, None)

                result_path = self._get_output_path()
                ef.save(result_path)
                # ef._workbook.close()  # ExcelFile上下文管理器会自动关闭

            gc.collect()
            return result_path
        finally:
            self._cleanup_temp()


class AggregateBackfillProcessor(BaseExcelProcessor):
    """聚合回填处理器"""

    def __init__(self, filepath: str, header_row: int,
                 group_cols_raw: List[str], agg_cols_raw: List[str],
                 agg_func: str, result_col_input: str, compute_engine,
                 already_normalized: bool = False):
        super().__init__(filepath, header_row, compute_engine,
                         already_normalized=already_normalized)
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input

    def process(self) -> str:
        # 使用缓存的或转换后的工作文件路径
        filepath = self._get_work_filepath()

        if not os.path.exists(filepath):
            raise FileNotFoundError(f"工作文件不存在: {filepath}")

        try:
            # 读取表头（仅列名），先不读取数据
            try:
                df_header = pd.read_excel(filepath, header=self.header_row - 1, nrows=0)
            except Exception as e:
                logger.error(f"读取表头失败: {e}")
                # 尝试用不同方式读取
                df_header = pd.read_excel(filepath, header=None, nrows=1)
                # 如果指定了标题行，但读取失败，尝试推断标题
                if self.header_row > 1:
                    df_header = pd.read_excel(filepath, header=self.header_row - 2, nrows=0)

            headers = list(df_header.columns)
            letter_to_name, name_to_letter = build_column_mapping(headers)

            def resolve_col(col: str) -> str:
                col = col.strip()
                if not col:
                    raise ValueError(f"列名为空")
                if col in headers:
                    return col
                upper = col.upper()
                if upper in letter_to_name:
                    return letter_to_name[upper]
                raise ValueError(f"列 '{col}' 无法映射到实际列名")

            group_cols = [resolve_col(c) for c in self.group_cols_raw if c.strip()]
            agg_cols = [resolve_col(c) for c in self.agg_cols_raw if c.strip()]

            if not group_cols:
                raise ValueError("分组列不能为空")
            if not agg_cols:
                raise ValueError("聚合列不能为空")

            # 以字符串类型读取数据
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str, na_filter=False)
            df = safe_cleanup_dataframe(df, auto_convert_numeric=True)

            if df.empty:
                raise ValueError("Excel文件为空或没有数据")

            missing = [c for c in group_cols + agg_cols if c not in df.columns]
            if missing:
                raise ValueError(f"列不存在: {missing}")

            result_df = self.compute_engine.simple_aggregate(df, group_cols, agg_cols, self.agg_func)

            with ExcelFile(filepath) as ef:
                ws = ef.active_sheet
                name_to_idx = build_column_name_mapping(ws, self.header_row)

                def get_col_index(col: str) -> int:
                    col = col.strip()
                    if col in name_to_idx:
                        return name_to_idx[col]
                    try:
                        return openpyxl.utils.column_index_from_string(col)
                    except Exception:
                        raise ValueError(f"列 '{col}' 在原始表中不存在")

                group_col_indices = [get_col_index(c) for c in group_cols]

                result_col_idx, _ = self._resolve_result_column(ws, self.result_col_input)

                result_map = {}
                for _, row in result_df.iterrows():
                    key = tuple(str(row[c]) for c in group_cols)
                    val = row[agg_cols[0]]
                    if hasattr(val, 'item'):
                        val = val.item()
                    result_map[key] = val

                start_row = self.header_row + 1
                end_row = ws.max_row
                for row in range(start_row, end_row + 1):
                    key = tuple(str(safe_get_cell_value(ws, row, idx)) for idx in group_col_indices)
                    val = result_map.get(key, "")
                    safe_write_cell(ws, row, result_col_idx, val)

                result_path = self._get_output_path()
                ef.save(result_path)
                # ef._workbook.close()  # ExcelFile上下文管理器会自动关闭

            gc.collect()
            return result_path
        finally:
            self._cleanup_temp()
    ```

- **services/excel/comparator.py**
  - language: `python`
  - content:
    ```python
# services/excel/comparator.py
import pandas as pd
import logging
from typing import Dict, Any, Tuple, List, Optional

logger = logging.getLogger(__name__)

class ExcelComparator:
    """Excel 文件对比核心类（业务服务）"""

    @staticmethod
    def _read_sheet(filepath: str, sheet: Optional[str]) -> pd.DataFrame:
        if sheet is None or sheet == '':
            return pd.read_excel(filepath, sheet_name=0)
        try:
            idx = int(sheet)
            return pd.read_excel(filepath, sheet_name=idx)
        except ValueError:
            return pd.read_excel(filepath, sheet_name=sheet)

    @classmethod
    def full_compare(cls, file1: str, file2: str,
                     key_fields: List[str], compare_fields: List[str],
                     sheet1: Optional[str] = None, sheet2: Optional[str] = None) -> Dict[str, Any]:
        df1 = cls._read_sheet(file1, sheet1)
        df2 = cls._read_sheet(file2, sheet2)

        cols1, cols2 = set(df1.columns), set(df2.columns)
        missing_keys1 = [k for k in key_fields if k not in cols1]
        missing_keys2 = [k for k in key_fields if k not in cols2]
        if missing_keys1 or missing_keys2:
            raise ValueError(f"主键字段不存在: {missing_keys1 + missing_keys2}")

        missing_comp1 = [c for c in compare_fields if c not in cols1]
        missing_comp2 = [c for c in compare_fields if c not in cols2]
        if missing_comp1 or missing_comp2:
            raise ValueError(f"对比字段不存在: {missing_comp1 + missing_comp2}")

        if df1.duplicated(subset=key_fields).any():
            raise ValueError("文件1中主键字段存在重复行")
        if df2.duplicated(subset=key_fields).any():
            raise ValueError("文件2中主键字段存在重复行")

        df1.set_index(key_fields, inplace=True)
        df2.set_index(key_fields, inplace=True)

        df1_comp = df1[compare_fields] if set(compare_fields).issubset(df1.columns) else df1.reindex(columns=compare_fields)
        df2_comp = df2[compare_fields] if set(compare_fields).issubset(df2.columns) else df2.reindex(columns=compare_fields)

        only_in_1 = df1_comp.loc[~df1_comp.index.isin(df2_comp.index)]
        only_in_2 = df2_comp.loc[~df2_comp.index.isin(df1_comp.index)]

        common_keys = df1_comp.index.intersection(df2_comp.index)
        diff_rows = []
        for key in common_keys:
            row1, row2 = df1_comp.loc[key], df2_comp.loc[key]
            if not row1.equals(row2):
                diff = {'主键': key}
                for f in compare_fields:
                    diff[f'文件1_{f}'] = row1[f]
                    diff[f'文件2_{f}'] = row2[f]
                    diff[f'{f}_是否一致'] = row1[f] == row2[f]
                diff_rows.append(diff)

        df_diff = pd.DataFrame(diff_rows) if diff_rows else None

        if not only_in_1.empty:
            only_in_1.reset_index(inplace=True)
        if not only_in_2.empty:
            only_in_2.reset_index(inplace=True)

        return {'only_in_1': only_in_1, 'only_in_2': only_in_2, 'diff': df_diff}

    @classmethod
    def field_match(cls, file1: str, file2: str,
                    match_field: str,
                    sheet1: Optional[str] = None, sheet2: Optional[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
        df1 = cls._read_sheet(file1, sheet1)
        df2 = cls._read_sheet(file2, sheet2)
        if match_field not in df1.columns:
            raise ValueError(f"字段 '{match_field}' 在文件1中不存在")
        if match_field not in df2.columns:
            raise ValueError(f"字段 '{match_field}' 在文件2中不存在")
        set1 = set(df1[match_field].astype(str).unique())
        set2 = set(df2[match_field].astype(str).unique())
        only_in_1 = set1 - set2
        only_in_2 = set2 - set1
        filtered_df1 = df1[df1[match_field].astype(str).isin(only_in_1)] if only_in_1 else pd.DataFrame()
        filtered_df2 = df2[df2[match_field].astype(str).isin(only_in_2)] if only_in_2 else pd.DataFrame()
        return filtered_df1, filtered_df2
    ```

- **services/excel/compute_engine.py**
  - language: `python`
  - content:
    ```python
# services/excel/compute_engine.py
import pandas as pd
from typing import List, Dict, Any, Optional, Union
from core.expression_engine import get_expression_engine
from services.excel.aggregator import Aggregator

class ComputeEngine:
    @staticmethod
    def evaluate_expression(expr: str, values: Dict[str, float]) -> float:
        engine = get_expression_engine()
        return engine.evaluate(expr, values)

    @staticmethod
    def extract_variables(expr: str) -> set:
        import re
        return set(re.findall(r'[A-Z]+', expr.upper()))

    @staticmethod
    def aggregate(df: pd.DataFrame, group_keys: List[str],
                  agg_spec: Dict[str, Union[str, List[str]]]) -> pd.DataFrame:
        return Aggregator.group_aggregate(df, group_keys, agg_spec)

    @staticmethod
    def simple_aggregate(df: pd.DataFrame, group_keys: List[str],
                         agg_cols: List[str], func: str = 'sum') -> pd.DataFrame:
        return Aggregator.simple_aggregate(df, group_keys, agg_cols, func)

    @staticmethod
    def aggregate_by_keys(df: pd.DataFrame, unique_keys: List[str],
                          sum_fields: List[str], other_fields: Optional[List[str]] = None) -> pd.DataFrame:
        return Aggregator.aggregate_by_keys(df, unique_keys, sum_fields, other_fields)
    ```

- **services/excel/filter.py**
  - language: `python`
  - content:
    ```python
# services/excel/filter.py
import pandas as pd
import logging
from typing import List, Tuple, Optional
from core.filter_engine import ConditionEngine

logger = logging.getLogger(__name__)

class ExcelFilter:
    """Excel 数据筛选核心类（委托给 core.filter_engine）"""

    @staticmethod
    def apply_filters(df: pd.DataFrame, conditions: List[Tuple[str, List[str]]],
                      fuzzy_field: Optional[str] = None, fuzzy_value: Optional[str] = None) -> pd.DataFrame:
        engine = ConditionEngine()
        exact = []
        fuzzy = []
        for field, values in conditions:
            if values:
                exact.append({'field': field, 'op': '=', 'values': values})
        if fuzzy_field and fuzzy_value:
            fuzzy.append({'field': fuzzy_field, 'op': '包含', 'value': fuzzy_value})
        return engine.apply_to_dataframe(df, exact, fuzzy)

    @staticmethod
    def read_and_combine(file_paths: List[str]) -> pd.DataFrame:
        from core.excel_utils import ExcelFile
        dfs = []
        for path in file_paths:
            df = ExcelFile.read_dataframe(path)
            dfs.append(df)
        if not dfs:
            raise ValueError("没有成功读取任何文件")
        return pd.concat(dfs, ignore_index=True)

    @staticmethod
    def round_numeric_columns(df: pd.DataFrame, decimals: int = 2) -> pd.DataFrame:
        df = df.copy()
        for col in df.columns:
            numeric_series = pd.to_numeric(df[col], errors='coerce')
            if numeric_series.notna().any():
                df[col] = numeric_series.apply(lambda x: round(x, decimals) if pd.notnull(x) else x)
        return df
    ```

- **services/excel/matcher.py**
  - language: `python`
  - content:
    ```python
import os
import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Alignment
from core.excel_utils import ExcelFile, col_letter_to_index
from core.excel_utils import copy_merge_structure
from services.data.reader import DataReader

class ExcelMatcher:
    """Excel 数据匹配填充核心类（支持多列值填充）"""

    def __init__(self, target_path, source_path, config, sheet_indices=None, data_reader=None):
        self.target_path = target_path
        self.source_path = source_path
        self.config = config
        self.sheet_indices = sheet_indices
        self.errors = []
        self.target_wb = None
        self.target_ws = None
        self.source_df = None
        self.data_reader = data_reader or DataReader()

        # 解析 duplicate_action（新参数），并与旧参数 allow_duplicate_keys 兼容
        # 优先级：duplicate_action 合法值 > allow_duplicate_keys=True → 'first' > 'error'
        _action = config.get('duplicate_action', '')
        if isinstance(_action, str):
            _action = _action.strip().lower()
        if _action not in ('error', 'first', 'all'):
            if config.get('allow_duplicate_keys'):
                _action = 'first'
            else:
                _action = 'error'
        self.duplicate_action = _action
        # 向后兼容：保留旧属性名，供外部潜在读取（内部逻辑统一用 duplicate_action）
        self.allow_duplicate_keys = (_action != 'error')

        # 优化 5：校验报告开关，默认开启（向后兼容）
        self.generate_report = config.get('generate_report', True)

        # 解析配置
        self.target_key_cols = config['target_key_cols']          # 列字母列表，如 ['A', 'B']
        self.source_match_cols = config['source_match_cols']      # 列字母列表
        self.source_value_cols = config['source_value_cols']      # 列字母列表，如 ['B', 'C']
        self.result_cols = config['result_cols']                  # 结果列列表，可以是列字母或新列名

        # 验证长度一致性
        if len(self.source_value_cols) != len(self.result_cols):
            self.errors.append(f"源值列数量({len(self.source_value_cols)})与结果列数量({len(self.result_cols)})不一致")
        if len(self.target_key_cols) != len(self.source_match_cols):
            self.errors.append(f"目标主键列数量({len(self.target_key_cols)})与源匹配列数量({len(self.source_match_cols)})不一致")

    def _validate_column_letters(self):
        """验证列字母有效性，并转换索引"""
        try:
            # 转换目标主键列索引
            self.target_key_indices = [column_index_from_string(col) for col in self.target_key_cols]
            # 转换源匹配列索引
            self.source_match_indices = [col_letter_to_index(col) for col in self.source_match_cols]
            # 转换源值列索引
            self.source_value_indices = [col_letter_to_index(col) for col in self.source_value_cols]

            # 处理结果列：区分是列字母还是新列名
            self.result_is_new = []
            self.result_col_indices = []
            self.result_col_names = []
            for rc in self.result_cols:
                try:
                    idx = column_index_from_string(rc)
                    self.result_is_new.append(False)
                    self.result_col_indices.append(idx)
                    self.result_col_names.append(None)
                except:
                    self.result_is_new.append(True)
                    self.result_col_indices.append(None)
                    self.result_col_names.append(rc.strip())
        except Exception as e:
            self.errors.append(f"列字母解析错误: {str(e)}")
            return False
        return True

    def _load_source(self):
        """加载源文件（支持多sheet）"""
        try:
            if self.sheet_indices:
                all_sheets = pd.read_excel(self.source_path, sheet_name=None, dtype=str)
                sheet_names = list(all_sheets.keys())
                selected_sheets = []
                for idx in self.sheet_indices:
                    if idx < 1 or idx > len(sheet_names):
                        self.errors.append(f"Sheet索引 {idx} 超出范围（共 {len(sheet_names)} 个sheet）")
                        return False
                    selected_sheets.append(sheet_names[idx-1])
                df_list = []
                for sheet_name in selected_sheets:
                    df = all_sheets[sheet_name]
                    df_list.append(df)
                self.source_df = pd.concat(df_list, ignore_index=True, sort=False)
            else:
                self.source_df = self.data_reader.read_excel(self.source_path, dtype=str)

            if self.source_df.empty:
                self.errors.append("源文件无数据")
                return False

            self.source_col_names = list(self.source_df.columns)
            max_cols = len(self.source_col_names)

            # 验证列范围
            for col_letter, idx in zip(self.source_match_cols, self.source_match_indices):
                if idx > max_cols:
                    self.errors.append(f"源文件匹配列 {col_letter} 超出列范围（共 {max_cols} 列）")
                    return False
            for col_letter, idx in zip(self.source_value_cols, self.source_value_indices):
                if idx > max_cols:
                    self.errors.append(f"源文件值列 {col_letter} 超出列范围（共 {max_cols} 列）")
                    return False
            return True
        except Exception as e:
            self.errors.append(f"读取源文件失败: {str(e)}")
            return False

    def _load_target(self):
        """加载目标文件"""
        try:
            self.target_wb = ExcelFile(self.target_path)
            self.target_wb.load()
            self.target_ws = self.target_wb.active_sheet
            max_col = self.target_ws.max_column
            for col_letter, idx in zip(self.target_key_cols, self.target_key_indices):
                if idx > max_col:
                    self.errors.append(f"目标文件键列 {col_letter} 超出最大列数 {max_col}")
                    return False
            return True
        except Exception as e:
            msg = str(e)
            if 'does not support' in msg or 'file format' in msg:
                self.errors.append("目标文件格式不支持，请使用 .xlsx 格式文件。如为 .xls 文件，请先用 Excel 另存为 .xlsx。")
            else:
                self.errors.append(f"读取目标文件失败: {msg}")
            return False

    def _extract_target_keys(self):
        """提取目标文件中所有行的组合主键

        性能优化（优化 3）：
        先一次性检测 key 列是否与任何合并区间相交。
        - 不相交：走 ws.cell(row, col).value 快速路径，跳过 O(M) 的 merged_cells 遍历
        - 相交：走 get_cell_value 慢路径（原逻辑），保证合并单元格取值正确

        两种路径在"key 列不在合并区间"时结果完全一致：
        get_cell_value 的内部循环此时必不命中，返回值就是 cell.value。
        """
        start_row = 2
        end_row = self.target_ws.max_row
        keys = []

        # 检测 key 列是否与任何合并区间相交（一次性开销，避免 R 次重复遍历）
        key_cols_set = set(self.target_key_indices)
        key_in_merge = False
        for merged in self.target_ws.merged_cells.ranges:
            if any(merged.min_col <= c <= merged.max_col for c in key_cols_set):
                key_in_merge = True
                break

        ws = self.target_ws
        if key_in_merge:
            # 慢速路径：key 列可能被合并，需走 get_cell_value
            for row in range(start_row, end_row + 1):
                key = []
                for col in self.target_key_indices:
                    val = self.target_wb.get_cell_value(row, col, sheet=ws)
                    key.append(str(val).strip() if val else '')
                keys.append(tuple(key))
        else:
            # 快速路径：直接读 cell.value
            for row in range(start_row, end_row + 1):
                key = []
                for col in self.target_key_indices:
                    val = ws.cell(row, col).value
                    key.append(str(val).strip() if val else '')
                keys.append(tuple(key))
        return keys, start_row, end_row

    def _build_lookup_dict(self):
        """构建源文件的查找字典：key -> 元组(值列1, 值列2, ...)"""
        # 获取列名
        match_col_names = [self.source_col_names[idx - 1] for idx in self.source_match_indices]
        value_col_names = [self.source_col_names[idx - 1] for idx in self.source_value_indices]

        lookup = {}
        for _, row in self.source_df.iterrows():
            key = tuple(str(row[col]).strip() if pd.notna(row[col]) else '' for col in match_col_names)
            values = tuple(
                row[col] if pd.notna(row[col]) else 0
                for col in value_col_names
            )
            # 如果允许重复键，保留第一次出现（与原行为一致，不覆盖）
            if key not in lookup:
                lookup[key] = values
        return lookup

    def _check_duplicate_target_keys(self, target_keys):
        """检查目标键是否重复，返回重复的键列表"""
        seen = set()
        duplicates = set()
        for key in target_keys:
            if key in seen:
                duplicates.add(key)
            else:
                seen.add(key)
        return duplicates

    def _write_result_columns(self, keys, lookup, start_row):
        """写入多个结果列（使用 self.result_col_indices 并原地更新新建列的索引）"""
        # 先处理新建列（插入列），更新 self.result_col_indices
        for i, (is_new, idx, col_name) in enumerate(zip(self.result_is_new, self.result_col_indices, self.result_col_names)):
            if is_new:
                max_col = self.target_ws.max_column
                new_idx = max_col + 1
                self.result_col_indices[i] = new_idx
                self.target_ws.insert_cols(new_idx)
                self.target_ws.cell(1, new_idx).value = col_name

        # 清除每个结果列的合并单元格
        for idx in self.result_col_indices:
            to_remove = []
            for merged in self.target_ws.merged_cells.ranges:
                if merged.min_col <= idx <= merged.max_col:
                    to_remove.append(merged)
            for merged in to_remove:
                self.target_ws.unmerge_cells(str(merged))

        # 写入数据
        # duplicate_action='first'：跳过重复键，仅填充第一个匹配行（其余行保留原值）
        # duplicate_action='all'  ：不跳过，所有匹配行都填充相同值
        # duplicate_action='error'：match() 已提前报错，正常不会走到重复键分支
        seen_keys = set() if self.duplicate_action == 'first' else None
        for row_offset, key in enumerate(keys):
            row_num = start_row + row_offset
            if seen_keys is not None:
                if key in seen_keys:
                    continue
                seen_keys.add(key)

            values_tuple = lookup.get(key, tuple(0 for _ in self.result_col_indices))
            for i, idx in enumerate(self.result_col_indices):
                val = values_tuple[i] if i < len(values_tuple) else 0
                # 尝试转数值
                try:
                    val = float(val)
                except (TypeError, ValueError):
                    pass
                self.target_ws.cell(row_num, idx).value = val

        # 复制合并单元格结构（基于第一个主键列）
        if self.target_key_indices:
            ref_col = self.target_key_indices[0]
            for idx in self.result_col_indices:
                copy_merge_structure(self.target_ws, ref_col, self.target_ws, idx, start_row, self.target_ws.max_row)

        # 设置右对齐（优化 4：Alignment 对象提到循环外复用，openpyxl 按值存储样式索引，等价）
        right_align = Alignment(horizontal='right')
        for idx in self.result_col_indices:
            for row in range(start_row, self.target_ws.max_row + 1):
                cell = self.target_ws.cell(row, idx)
                cell.alignment = right_align

    def _generate_report(self, used_target_keys, lookup, start_row, target_keys_full=None):
        """生成校验报告 sheet，展示所有源值列的值

        性能优化：
        - 优化 2：target_keys_full 可由调用方传入，避免重复调用 _extract_target_keys
        - 优化 1：set(target_keys_used) 提到循环外，避免 O(S*R) 的重复哈希
        - 优化 4：Alignment 对象提到循环外复用
        """
        if target_keys_full is None:
            target_keys_full, _, _ = self._extract_target_keys()
        target_keys_used = used_target_keys
        source_keys_set = set(lookup.keys())
        # 优化 1：提前构建 target_keys_used 的 set，避免在"仅源文件"循环里每次重建
        target_keys_used_set = set(target_keys_used)

        # 准备列标题：行号、目标键、源匹配键、每个源值列的值、匹配状态、备注
        header_row = 1
        target_col_names = []
        for col_1based in self.target_key_indices:
            col_name = self.target_ws.cell(header_row, col_1based).value
            if not col_name:
                col_name = f"列{get_column_letter(col_1based)}"
            target_col_names.append(col_name)
        target_key_label = "|".join(target_col_names)
        target_file_name = os.path.basename(self.target_path)

        # 源值列列名（使用列字母作为标识）
        source_value_labels = [f"源文件值({col})" for col in self.source_value_cols]

        report_sheet = self.target_wb._workbook.create_sheet("校验报告")
        headers = ["行号", f"目标文件键列值\n({target_file_name}【{target_key_label}】)", "源文件匹配键"] + source_value_labels + ["匹配状态", "备注"]
        for col, header in enumerate(headers, 1):
            report_sheet.cell(1, col).value = header

        row_idx = 2
        # 遍历目标文件每一行（包括未匹配的）
        for i, (full_key, used_key) in enumerate(zip(target_keys_full, target_keys_used)):
            row_num = start_row + i
            if used_key in source_keys_set:
                val_tuple = lookup.get(used_key, tuple())
                # 匹配成功
                row_data = [row_num, "|".join(str(v) for v in full_key), "|".join(str(v) for v in used_key)]
                # 添加每个值列的值
                for col_idx in range(len(self.source_value_cols)):
                    val = val_tuple[col_idx] if col_idx < len(val_tuple) else ''
                    row_data.append(val)
                row_data.append(1)
                row_data.append('')
                for col_idx, val in enumerate(row_data, 1):
                    report_sheet.cell(row_idx, col_idx).value = val
                row_idx += 1
            else:
                # 仅目标存在，未匹配
                row_data = [row_num, "|".join(str(v) for v in full_key), "|".join(str(v) for v in used_key)]
                # 值列留空
                for _ in self.source_value_cols:
                    row_data.append('')
                row_data.append(0)
                row_data.append('')
                for col_idx, val in enumerate(row_data, 1):
                    report_sheet.cell(row_idx, col_idx).value = val
                row_idx += 1

        # 仅源文件存在的记录
        # 优化 1：使用提前构建好的 target_keys_used_set，避免 O(S*R) 的重复哈希
        for key, val_tuple in lookup.items():
            if key not in target_keys_used_set:
                row_data = ['-', '', "|".join(str(v) for v in key)]
                for col_idx in range(len(self.source_value_cols)):
                    val = val_tuple[col_idx] if col_idx < len(val_tuple) else ''
                    row_data.append(val)
                row_data.append(0)
                row_data.append('仅源文件')
                for col_idx, val in enumerate(row_data, 1):
                    report_sheet.cell(row_idx, col_idx).value = val
                row_idx += 1

        # 设置对齐（优化 4：Alignment 对象提到循环外复用）
        left_align = Alignment(horizontal='left')
        right_align = Alignment(horizontal='right')
        for col in range(1, len(headers) + 1):
            align = left_align if col == 1 else right_align
            for row in range(2, row_idx):
                cell = report_sheet.cell(row, col)
                cell.alignment = align

    def match(self):
        if not self._validate_column_letters():
            return False, self.errors
        if not self._load_source():
            return False, self.errors
        if not self._load_target():
            return False, self.errors

        target_keys, start_row, end_row = self._extract_target_keys()

        if self.duplicate_action == 'error':
            duplicates = self._check_duplicate_target_keys(target_keys)
            if duplicates:
                dup_examples = [f"'{k}'" for k in list(duplicates)[:5]]
                self.errors.append(f"目标文件键列存在重复值: {', '.join(dup_examples)}。如需继续填充，请将'目标键重复时的处理方式'设为'只填充第一个'或'填充所有匹配行'。")
                return False, self.errors

        lookup = self._build_lookup_dict()
        match_len = len(self.target_key_cols)
        used_target_keys = [key[:match_len] for key in target_keys]

        # 写入结果列
        self._write_result_columns(used_target_keys, lookup, start_row)

        # 生成报告（多列版本）
        # 优化 2：把已算好的 target_keys 传进去，避免 _generate_report 重新提取
        # 优化 5：generate_report=False 时跳过报告生成（大文件场景显著提速）
        if self.generate_report:
            self._generate_report(used_target_keys, lookup, start_row,
                                  target_keys_full=target_keys)

        out_path = os.path.join(os.path.dirname(self.target_path), f"matched_{os.path.basename(self.target_path)}")
        self.target_wb.save(out_path)
        return True, out_path

    def debug_info(self):
        if not self._validate_column_letters():
            return {'error': self.errors}
        if not self._load_source():
            return {'error': self.errors}
        if not self._load_target():
            return {'error': self.errors}

        target_keys, start_row, end_row = self._extract_target_keys()
        lookup = self._build_lookup_dict()
        match_len = len(self.target_key_cols)
        used_target_keys = [key[:match_len] for key in target_keys]

        target_set = set(used_target_keys)
        source_set = set(lookup.keys())
        common = target_set & source_set
        only_target = target_set - source_set
        only_source = source_set - target_set

        matched_pairs = []
        unmatched_pairs = []
        for i, (full_key, used_key) in enumerate(zip(target_keys, used_target_keys)):
            row_num = start_row + i
            matched_val = lookup.get(used_key)
            if matched_val is not None:
                matched_pairs.append((row_num, full_key, used_key, matched_val))
            else:
                unmatched_pairs.append((row_num, full_key, used_key))
            if len(matched_pairs) >= 20 and len(unmatched_pairs) >= 20:
                break

        return {
            'config': self.config,
            'target_rows': len(target_keys),
            'target_keys_sample': [(start_row + i, key) for i, key in enumerate(target_keys[:20])],
            'source_keys_sample': list(source_set)[:20],
            'source_values_sample': [(k, lookup[k]) for k in list(source_set)[:20]],
            'common_count': len(common),
            'only_target_count': len(only_target),
            'only_source_count': len(only_source),
            'matched_count': len(common),
            'unmatched_keys': list(only_target)[:20],
            'start_row': start_row,
            'end_row': end_row,
            'matched_pairs': matched_pairs,
            'unmatched_pairs': unmatched_pairs
        }
    ```

- **services/field_parser/__init__.py**
  - language: `python`
  - content:
    ```python
"""字段解析服务"""
from .service import FieldParserService
    ```

- **services/field_parser/service.py**
  - language: `python`
  - content:
    ```python
# services/field_parser/service.py
import re
import logging
from typing import List, Dict
from difflib import SequenceMatcher

logger = logging.getLogger(__name__)

COMMON_REPLACEMENTS = {
    'id': '标识',
    'tel': '电话',
    'phone': '电话',
    'mobile': '手机',
    'name': '姓名',
    'addr': '地址',
    'no': '编号',
    'num': '数量',
}

class FieldParserService:
    """字段解析与分层匹配服务"""

    def __init__(self, similarity_threshold: int = 30):
        self.threshold = similarity_threshold / 100.0

    def parse_field_list(self, text: str) -> List[str]:
        parts = re.split(r'[,，\s\n\r\t]+', text.strip())
        return [p.strip() for p in parts if p.strip()]

    def _normalize(self, s: str) -> str:
        s = s.strip().lower()
        s = re.sub(r'[\s\-_]+', '', s)
        for abbr, full in COMMON_REPLACEMENTS.items():
            if s == abbr:
                return full
        return s

    def match_headers(self, headers_a: List[str], headers_b: List[str]) -> Dict:
        norm_a = [self._normalize(h) for h in headers_a]
        norm_b = [self._normalize(h) for h in headers_b]

        # ---------- 第一步：完全匹配（标准化后相等） ----------
        exact_matched = []
        remaining_a = []
        remaining_b_idx = set(range(len(norm_b)))  # 还未匹配的 B 列索引

        for i, ha in enumerate(headers_a):
            na = norm_a[i]
            matched = False
            for j in list(remaining_b_idx):
                if na == norm_b[j]:
                    exact_matched.append({
                        "standard": ha,
                        "compare": headers_b[j],
                        "score": 1.0
                    })
                    remaining_b_idx.remove(j)
                    matched = True
                    break
            if not matched:
                remaining_a.append(i)

        # ---------- 第二步：相似度匹配（按分数从高到低） ----------
        # 收集所有可能配对的相似度
        candidates = []
        for i in remaining_a:
            for j in remaining_b_idx:
                score = self._similarity_score(norm_a[i], norm_b[j])
                if score >= self.threshold:
                    candidates.append((score, i, j))

        # 按分数降序
        candidates.sort(key=lambda x: x[0], reverse=True)

        used_a = set()
        used_b = set()
        fuzzy_matched = []

        for score, i, j in candidates:
            if i in used_a or j in used_b:
                continue
            fuzzy_matched.append({
                "standard": headers_a[i],
                "compare": headers_b[j],
                "score": round(score, 2)
            })
            used_a.add(i)
            used_b.add(j)

        # ---------- 汇总结果 ----------
        all_matched = exact_matched + fuzzy_matched
        matched_a_idx = set(range(len(headers_a))) - {i for i in remaining_a} | used_a
        matched_b_idx = set(range(len(headers_b))) - remaining_b_idx | used_b

        only_a = [headers_a[i] for i in range(len(headers_a)) if i not in matched_a_idx]
        only_b = [headers_b[j] for j in range(len(headers_b)) if j not in matched_b_idx]

        return {"matched": all_matched, "only_a": only_a, "only_b": only_b}

    def _similarity_score(self, a: str, b: str) -> float:
        if a == b:
            return 1.0
        if not a or not b:
            return 0.0

        # SequenceMatcher
        seq_sim = SequenceMatcher(None, a, b).ratio()

        # 归一化编辑距离
        len_a, len_b = len(a), len(b)
        dp = [[0]*(len_b+1) for _ in range(len_a+1)]
        for i in range(len_a+1):
            dp[i][0] = i
        for j in range(len_b+1):
            dp[0][j] = j
        for i in range(1, len_a+1):
            for j in range(1, len_b+1):
                cost = 0 if a[i-1] == b[j-1] else 1
                dp[i][j] = min(dp[i-1][j]+1, dp[i][j-1]+1, dp[i-1][j-1]+cost)
        edit_sim = 1.0 - (dp[len_a][len_b] / max(len_a, len_b))

        # Jaccard
        set_a, set_b = set(a), set(b)
        common = len(set_a & set_b)
        total = len(set_a | set_b)
        jaccard = common / total if total > 0 else 0.0

        return 0.4 * seq_sim + 0.3 * edit_sim + 0.3 * jaccard
    ```

- **services/subtotal/__init__.py**
  - language: `python`
  - content:
    ```python
# services/subtotal/__init__.py
    ```

