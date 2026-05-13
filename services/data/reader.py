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

    @staticmethod
    def read_csv_streaming(header_path: str, data_paths: List[str],
                           needed_columns: set, null_markers: Optional[Dict] = None,
                           clean_columns: Optional[List[str]] = None) -> Iterator[Dict[str, str]]:
        """
        流式读取 CSV（用于大文件），逐行返回字典
        此方法复用 simplify 模块的 StreamingCsvReader 逻辑
        """
        from services.simplify.reader import StreamingCsvReader
        # 先读取表头
        with open(header_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
        reader = StreamingCsvReader(header_path, data_paths, headers, needed_columns,
                                    null_markers=null_markers, clean_columns=clean_columns)
        for row in reader:
            yield row