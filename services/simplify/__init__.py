# utils/simplify/__init__.py
from typing import List, Dict, Tuple, Optional
from .validation import read_csv_header, ColumnMapper, CsvFileValidator, MixedTypeChecker, ValidationEngine, ValidationResult
from .config import ConfigParser, ProcessConfig
from .reader import StreamingCsvReader
from .simplify_aggregation import AggregationStrategy, MemoryAggregation, SqliteAggregation, AggregationStrategyFactory
from .output import OutputFormatter
from .service import SimplifyService
from .async_task import submit_task, load_status, get_log, init_config

def validate_csv_files(header_path: str, data_paths: List[str], params: Dict, preview_rows: int = 100) -> Tuple[bool, Optional[str]]:
    engine = ValidationEngine(header_path, data_paths, preview_rows)
    config = ConfigParser.parse(params)
    result = engine.validate(config)
    if result.is_valid:
        return True, None
    else:
        return False, result.error_msg

__all__ = [
    'ConfigParser', 'ProcessConfig',
    'ValidationEngine', 'ValidationResult', 'ColumnMapper', 'CsvFileValidator', 'MixedTypeChecker',
    'StreamingCsvReader',
    'AggregationStrategy', 'MemoryAggregation', 'SqliteAggregation', 'AggregationStrategyFactory',
    'OutputFormatter',
    'SimplifyService',
    'submit_task', 'load_status', 'get_log', 'init_config',
    'read_csv_header', 'validate_csv_files'
]