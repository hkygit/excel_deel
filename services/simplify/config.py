# utils/simplify/config.py
from dataclasses import dataclass
from typing import List, Dict, Optional


@dataclass
class ProcessConfig:
    """处理配置对象（不可变）"""
    key_cols: List[str]
    keep_cols: List[str]
    non_sum_fields: List[str]
    filters: List[Dict]
    output_format: str
    max_keys_in_memory: int = 500000
    chunk_size: int = 100000
    non_sum_override_markers: Optional[Dict[str, List[str]]] = None

    def __post_init__(self):
        self.key_cols = list(dict.fromkeys(self.key_cols))
        self.keep_cols = list(dict.fromkeys(self.keep_cols))
        self.non_sum_fields = list(dict.fromkeys(self.non_sum_fields))
        if self.non_sum_override_markers is None:
            self.non_sum_override_markers = {}


class ConfigParser:
    @staticmethod
    def parse(params: Dict, max_keys_in_memory: int = 500000) -> ProcessConfig:
        return ProcessConfig(
            key_cols=params.get('key_cols', []),
            keep_cols=params.get('keep_cols', []),
            non_sum_fields=params.get('non_sum_fields', []),
            filters=params.get('filters', []),
            output_format=params.get('output_format', 'auto'),
            max_keys_in_memory=max_keys_in_memory,
            non_sum_override_markers=params.get('non_sum_override_markers', {})
        )