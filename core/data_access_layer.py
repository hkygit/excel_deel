# core/data_access_layer.py
"""
统一数据访问层 (核心基础设施)
- 封装文件读写，自动类型推断与清理
- 可插拔的读取器/写入器
"""

import logging
from typing import Optional, Dict, Any
import pandas as pd
import numpy as np

from services.data.reader import DataReader
from services.data.writer import DataWriter

logger = logging.getLogger(__name__)


class DataAccessLayer:
    """统一数据访问层"""

    def __init__(self, reader: Optional[DataReader] = None, writer: Optional[DataWriter] = None):
        self.reader = reader or DataReader()
        self.writer = writer or DataWriter()
        logger.debug("DataAccessLayer 初始化完成")

    def read_excel(self, filepath: str, sheet: Any = 0,
                   dtype_hints: Optional[Dict[str, type]] = None,
                   auto_clean: bool = True) -> pd.DataFrame:
        logger.info(f"读取 Excel 文件: {filepath}, sheet={sheet}")
        df = self.reader.read_excel(filepath, sheet=sheet)

        if auto_clean:
            df = self._auto_clean_dataframe(df, dtype_hints)

        logger.debug(f"读取完成，形状: {df.shape}")
        return df

    def read_csv(self, filepath: str, encoding: str = 'utf-8',
                 dtype_hints: Optional[Dict[str, type]] = None,
                 auto_clean: bool = True) -> pd.DataFrame:
        logger.info(f"读取 CSV 文件: {filepath}")
        df = self.reader.read_csv(filepath, encoding=encoding)

        if auto_clean:
            df = self._auto_clean_dataframe(df, dtype_hints)

        logger.debug(f"读取完成，形状: {df.shape}")
        return df

    def read_auto(self, filepath: str, **kwargs) -> pd.DataFrame:
        logger.info(f"自动读取文件: {filepath}")
        df = self.reader.read_auto(filepath, **kwargs)
        df = self._auto_clean_dataframe(df)
        logger.debug(f"读取完成，形状: {df.shape}")
        return df

    def _auto_clean_dataframe(self, df: pd.DataFrame,
                              dtype_hints: Optional[Dict[str, type]] = None) -> pd.DataFrame:
        df = df.copy()

        if dtype_hints:
            for col, dtype in dtype_hints.items():
                if col in df.columns:
                    try:
                        df[col] = df[col].astype(dtype)
                    except Exception as e:
                        logger.warning(f"列 '{col}' 类型转换失败: {e}")

        for col in df.select_dtypes(include=['object']).columns:
            # 去除首尾空格
            df[col] = df[col].astype(str).str.strip()
            # 将空字符串替换为 NaN
            df[col] = df[col].replace(r'^\s*$', np.nan, regex=True)

            # 尝试转为数值
            numeric_series = pd.to_numeric(df[col], errors='coerce')
            # 如果存在至少一个有效数值，则将该列转换为数值类型
            if numeric_series.notna().any():
                df[col] = numeric_series
                logger.debug(f"列 '{col}' 已自动转换为数值类型")

        return df

    def write_excel(self, df: pd.DataFrame, filepath: str, **kwargs):
        logger.info(f"写入 Excel 文件: {filepath}")
        self.writer.write_excel(df, filepath, **kwargs)

    def write_csv(self, df: pd.DataFrame, filepath: str, **kwargs):
        logger.info(f"写入 CSV 文件: {filepath}")
        self.writer.write_csv(df, filepath, **kwargs)


_default_dal = DataAccessLayer()


def get_data_access_layer() -> DataAccessLayer:
    return _default_dal
