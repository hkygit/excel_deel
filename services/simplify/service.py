# utils/simplify/service.py  (优化版)
import os
import psutil
from typing import List, Dict, Tuple
from werkzeug.utils import secure_filename
from core.task.base import Task
from .config import ConfigParser
from .validation import ValidationEngine, ColumnMapper, MixedTypeChecker, ColumnTypeDetector
from .reader import StreamingCsvReader
from .simplify_aggregation import MemoryAggregation, AggregationStrategyFactory
from .output import OutputFormatter
from .helpers import RowFilter, build_column_mapping


# 内存聚合最大键数硬限制（避免极端情况）
HARD_MAX_KEYS_IN_MEMORY = 2_000_000


class SimplifyService(Task):
    def __init__(self, header_path: str, data_paths: List[str], params: Dict,
                 original_filename: str = None):
        self.header_path = header_path
        self.data_paths = data_paths
        self.params = params
        self.original_filename = original_filename

    def execute(self) -> Tuple[str, str, int, int]:
        config = ConfigParser.parse(self.params)

        validator = ValidationEngine(self.header_path, self.data_paths)
        result = validator.validate(config)
        if not result.is_valid:
            raise ValueError(result.error_msg)
        headers = validator.headers
        sample_rows = validator.sample_rows
        sample_dicts = []
        if sample_rows:
            for row_vals in sample_rows:
                row_dict = {headers[i]: row_vals[i] for i in range(len(headers))}
                sample_dicts.append(row_dict)

        required_cols = set(config.key_cols) | set(config.keep_cols) | set(config.non_sum_fields)
        for f in config.filters:
            required_cols.add(f.get('field', ''))
        needed_columns = required_cols

        detector = ColumnTypeDetector(headers, sample_dicts)
        date_columns = detector.detect_date_columns()
        numeric_cols, string_cols, mixed_cols = detector.detect_numeric_and_string(
            exclude_cols=set(config.key_cols) | set(config.non_sum_fields)
        )

        if config.keep_cols:
            sum_fields = config.keep_cols[:]
            non_sum_final = config.non_sum_fields[:]
        else:
            sum_fields = numeric_cols
            non_sum_final = config.non_sum_fields[:] + mixed_cols + string_cols
            non_sum_final = list(dict.fromkeys(non_sum_final))

        letter_to_name, name_to_letter = build_column_mapping(headers)
        filter_obj = RowFilter(config.filters, name_to_letter, headers, date_columns)

        # 动态阈值 + 硬上限
        mem_available = psutil.virtual_memory().available
        dynamic_threshold = int(mem_available / 200)
        dynamic_threshold = max(10000, min(dynamic_threshold, config.max_keys_in_memory, HARD_MAX_KEYS_IN_MEMORY))

        override_markers = config.non_sum_override_markers if hasattr(config, 'non_sum_override_markers') else {}

        reader = StreamingCsvReader(self.header_path, self.data_paths, headers, needed_columns)
        aggregator = MemoryAggregation()
        input_row_count = 0
        key_cols = config.key_cols
        for row in reader:
            input_row_count += 1
            if not filter_obj.apply(row):
                continue
            is_memory_agg = hasattr(aggregator, 'data') and hasattr(aggregator, 'size')
            if is_memory_agg and aggregator.size() > dynamic_threshold:
                new_agg = AggregationStrategyFactory.create(dynamic_threshold + 1)
                for key, (sum_dict, first_dict, cnt) in aggregator.data.items():
                    row_dict = dict(zip(key_cols, key))
                    row_dict.update(sum_dict)
                    row_dict.update(first_dict)
                    row_dict['__count__'] = cnt
                    new_agg.add_row(row_dict, key_cols, sum_fields, non_sum_final, override_markers)
                aggregator = new_agg
            aggregator.add_row(row, key_cols, sum_fields, non_sum_final, override_markers)

        result_df = aggregator.finalize(key_cols, sum_fields, non_sum_final)
        output_row_count = len(result_df)

        if self.original_filename:
            safe_name = secure_filename(self.original_filename).replace('.', '_')
            if self.data_paths:
                base_name = f"{safe_name}等{len(self.data_paths)+1}个文件（精简版）"
            else:
                base_name = f"{safe_name}（精简版）"
        else:
            base_name = "result"

        output_dir = os.path.dirname(self.header_path)
        output_path, actual_format = OutputFormatter.write(result_df, output_dir, base_name, config.output_format)
        return output_path, actual_format, input_row_count, output_row_count