# utils/simplify/reader.py
import csv
import os
from typing import List, Set, Dict, Iterator, Optional, Callable

# 默认的无效值标记（适用于所有非聚合列）
DEFAULT_NULL_MARKERS = {'none', 'null', ''}


def create_clean_func(null_markers: Optional[Dict[str, List[str]]] = None,
                      default_columns: Optional[List[str]] = None) -> Callable[[Dict[str, str]], Dict[str, str]]:
    """
    创建清洗函数
    :param null_markers: 列级空值标记映射，如 {"医院等级": ["0","无"]}
    :param default_columns: 需要应用默认空值标记的列名列表（通常是非聚合列）
    :return: 清洗函数
    """
    if null_markers is None:
        null_markers = {}

    # 预处理：将列级标记转为小写集合
    markers_per_col = {}
    for col, markers in null_markers.items():
        markers_per_col[col] = {str(m).strip().lower() for m in markers if m}

    def clean_func(row: Dict[str, str]) -> Dict[str, str]:
        # 应用列级自定义标记
        for col, markers in markers_per_col.items():
            if col in row:
                val = row[col]
                if val and val.strip().lower() in markers:
                    row[col] = ''
        # 对 default_columns 应用默认标记（如果该列没有被自定义标记覆盖）
        if default_columns:
            for col in default_columns:
                if col in row and col not in markers_per_col:
                    val = row[col]
                    if val and val.strip().lower() in DEFAULT_NULL_MARKERS:
                        row[col] = ''
        return row

    return clean_func


class StreamingCsvReader:
    def __init__(self, header_path: str, data_paths: List[str], headers: List[str],
                 needed_columns: Set[str],
                 null_markers: Optional[Dict[str, List[str]]] = None,
                 clean_columns: Optional[List[str]] = None):
        """
        流式 CSV 读取器
        :param header_path: 表头文件路径
        :param data_paths: 数据文件路径列表
        :param headers: 所有列名（用于对齐数据文件）
        :param needed_columns: 需要读取的列名集合
        :param null_markers: 可选，列级空值标记映射
        :param clean_columns: 可选，需要应用默认空值标记的列名列表（通常是非聚合列）
        """
        self.header_path = header_path
        self.data_paths = data_paths
        self.headers = headers
        self.needed_columns = needed_columns
        self.clean_func = create_clean_func(null_markers, clean_columns)

    def __iter__(self) -> Iterator[Dict[str, str]]:
        # 处理表头文件
        with open(self.header_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                filtered = self._filter_row(row)
                filtered = self.clean_func(filtered)
                yield filtered

        # 处理数据文件
        for data_path in self.data_paths:
            with open(data_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row_vals in reader:
                    if len(row_vals) != len(self.headers):
                        raise ValueError(f"数据文件 {os.path.basename(data_path)} 列数不一致")
                    row = {self.headers[i]: row_vals[i] for i in range(len(self.headers))}
                    filtered = self._filter_row(row)
                    filtered = self.clean_func(filtered)
                    yield filtered

    def _filter_row(self, row: Dict[str, str]) -> Dict[str, str]:
        return {k: v for k, v in row.items() if k in self.needed_columns}