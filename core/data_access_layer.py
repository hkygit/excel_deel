# core/data_access_layer.py
"""统一数据访问层 (核心基础设施)"""
import logging
from typing import Optional, Dict, Any
import pandas as pd
import numpy as np
from core.excel_helpers import safe_cleanup_dataframe

logger = logging.getLogger(__name__)


class DataAccessLayer:
    """统一数据访问层"""

    def __init__(self, reader=None, writer=None):
        from services.data.reader import DataReader
        from services.data.writer import DataWriter
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
        """
        自动清洗 DataFrame：
        - 首先应用 dtype_hints（如果提供），使用 pd.to_numeric 安全转换数值类型。
        - 然后进行字符串清理和自动数值转换（保持向后兼容）。
        """
        df = df.copy()

        # 1. 应用 dtype_hints（优先，以满足测试期望）
        if dtype_hints:
            for col, dtype in dtype_hints.items():
                if col not in df.columns:
                    continue
                if dtype in (int, float, np.int64, np.float64):
                    # 使用 pd.to_numeric 将无效值转为 NaN
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                else:
                    try:
                        df[col] = df[col].astype(dtype)
                    except Exception as e:
                        logger.warning(f"列 '{col}' 类型转换失败: {e}")

        # 2. 通用清洗（字符串清理 + 自动数值转换）
        df = safe_cleanup_dataframe(df, auto_convert_numeric=True)

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
