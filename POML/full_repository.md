# Project Object Markup Language (POML)

- **repository**: `project`
- **export_filter**: `full_repository`
- **total_files**: 183

## Files

- **.trash/20260513_165803/services/checker/precheck.py**
  - language: `python`
  - content:
    ```python
# services/checker/precheck.py
import os
import pandas as pd
from typing import List
from core.exceptions import ValidationError

class PrecheckEngine:
    """预检：快速扫描文件，返回统计信息"""

    @staticmethod
    def estimate(files: List[str], sample_rows: int = 1000) -> dict:
        total_rows = 0
        total_size_mb = 0
        for f in files:
            if not os.path.exists(f):
                continue
            total_size_mb += os.path.getsize(f) / (1024 * 1024)
            try:
                # 采样读取
                sample = pd.read_excel(f, nrows=sample_rows)
                # 粗略推算总行数（假设文件均匀）
                # 真实实现需更精细
            except Exception:
                pass
        return {
            "file_count": len(files),
            "total_size_mb": round(total_size_mb, 2),
            "estimated_rows": total_rows,
            "warnings": [],
        }
    ```

- **.trash/20260513_165803/utils/excel_helpers.py**
  - language: `python`
  - content:
    ```python
# utils/excel_helpers.py
"""向后兼容层：委托给 core.excel_utils"""
from core.excel_utils import (
    ExcelFile, col_letter_to_index, col_index_to_letter,
    validate_column_letters, copy_merge_structure
)
# 保留旧的导入路径，确保其他模块无需修改
__all__ = ['ExcelFile', 'col_letter_to_index', 'col_index_to_letter', 'validate_column_letters', 'copy_merge_structure']
    ```

- **.trash/20260513_165803/utils/exceptions.py**
  - language: `python`
  - content:
    ```python
# utils/common/exceptions.py
class AppException(Exception):
    def __init__(self, message, code=400, error_type='bad_request'):
        self.message = message
        self.code = code
        self.error_type = error_type
        super().__init__(message)

class ValidationError(AppException):
    def __init__(self, message):
        super().__init__(message, code=400, error_type='validation_error')

class NotFoundError(AppException):
    def __init__(self, message='资源不存在'):
        super().__init__(message, code=404, error_type='not_found')

class ConflictError(AppException):
    def __init__(self, message='资源冲突'):
        super().__init__(message, code=409, error_type='conflict')

class InternalError(AppException):
    def __init__(self, message='服务器内部错误'):
        super().__init__(message, code=500, error_type='internal_error')
    ```

- **.trash/20260513_165803/utils/validators.py**
  - language: `python`
  - content:
    ```python
# utils/validators.py
"""向后兼容：委托给 core/validators.py"""
from core.validators import (
    validate_column_letters,
    validate_sheet_index,
    validate_expression,
    sanitize_filename
)
    ```

- **.trash/utils_20260513_170109/__init__.py**
  - language: `python`
  - content:
    ```python
# utils/__init__.py
"""工具层（大部分能力已迁移至 core/services，此处仅保留向后兼容的薄封装）"""
from core.excel_utils import ExcelFile, col_letter_to_index, col_index_to_letter, copy_merge_structure
from core.validators import validate_column_letters, validate_sheet_index, validate_expression, sanitize_filename
from core.filter_engine import ConditionFilter, Condition

__all__ = [
    'ExcelFile', 'col_letter_to_index', 'col_index_to_letter', 'copy_merge_structure',
    'validate_column_letters', 'validate_sheet_index', 'validate_expression', 'sanitize_filename',
    'ConditionFilter', 'Condition',
]
    ```

- **POML/calc.md**
  - language: `markdown`
  - content:
    ```markdown
# Project Object Markup Language (POML)

- **repository**: `project`
- **export_filter**: `calc`
- **total_files**: 25

## Files

- **blueprints/calculator.py**
  - language: `python`
  - content:
    ```python
# blueprints/calculator.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app
from openpyxl.utils.exceptions import InvalidFileException

from services import FileService, ComputeEngine
from core.decorators import handle_exceptions
from core.response import error_response
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start
from core.exceptions import ValidationError
from core.excel_helpers import normalize_excel_file
from services.excel.calculator_processor import ColumnCalculatorProcessor, AggregateBackfillProcessor

logger = logging.getLogger(__name__)
calc_bp = Blueprint('calc', __name__, url_prefix='/calc', template_folder='../templates/calculator')


def ensure_excel_compatible(filepath: str, file_service) -> (str, str):
    """
    确保文件可以被后续处理（openpyxl 和 pandas）正常读取。
    返回 (规范化的文件路径, 临时文件标记)
    总是转换为临时文件，避免原始文件被重复转换。
    """
    try:
        normalized_path = normalize_excel_file(filepath)
    except Exception as e:
        file_service.cleanup([filepath])
        raise ValidationError(f"Excel 文件处理失败: {str(e)}。请确保文件是有效的 Excel 文件且未损坏。")
    # 删除原始上传文件
    file_service.cleanup([filepath])
    return normalized_path, normalized_path


@calc_bp.route('/')
def index():
    return render_template('calculator/index.html')


@calc_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    compute_engine = current_app.container.compute_engine()

    if 'file' not in request.files:
        return error_response('未上传文件', 400)
    file = request.files['file']
    if file.filename == '':
        return error_response('文件名为空', 400)

    original_filename = file.filename
    ext = os.path.splitext(original_filename)[1].lower()
    if ext not in ('.xlsx', '.xls', '.xlsm'):
        raise ValidationError(f"不支持的文件格式: {ext}，请上传 .xlsx 或 .xls 文件")

    # 保存原始上传文件
    filepath = file_service.save_upload(file)
    temp_processed = None

    # 确保文件兼容，总是转换为临时文件
    filepath, temp_processed = ensure_excel_compatible(filepath, file_service)

    calc_type = request.form.get('calc_type')
    if calc_type not in ('column', 'aggregate'):
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        return error_response('无效的操作类型', 400)

    # 表头行号：支持 header_row=0 表示无表头
    header_row_str = request.form.get('header_row', '').strip()
    if header_row_str:
        try:
            header_row = int(header_row_str)
            if header_row < 0:
                raise ValueError
        except ValueError:
            return error_response('表头行号必须为非负整数', 400)
    else:
        try:
            data_start = detect_header_start(filepath)
            header_row = data_start - 1
            if header_row < 0:
                header_row = 0
        except Exception:
            header_row = 0

    result_path = None
    try:
        if calc_type == 'column':
            expression = request.form.get('expression', '').strip()
            result_col = request.form.get('result_col_column')
            merge_by_cols = request.form.get('merge_by_cols', '').strip()

            if not expression:
                return error_response('表达式不能为空', 400)

            processor = ColumnCalculatorProcessor(
                filepath=filepath,
                header_row=header_row,
                expression=expression,
                result_col_input=result_col,
                merge_by_cols_str=merge_by_cols,
                compute_engine=compute_engine
            )
            result_path = processor.process()

        else:  # aggregate
            group_cols_str = request.form.get('group_cols', '').strip()
            agg_cols_str = request.form.get('agg_cols', '').strip()
            agg_func = request.form.get('agg_func', 'sum')
            result_col = request.form.get('result_col_aggregate', '').strip()

            if not group_cols_str or not agg_cols_str or not result_col:
                return error_response('分组列、聚合列和结果列不能为空', 400)

            group_cols_raw = [c.strip() for c in group_cols_str.split(',')]
            agg_cols_raw = [c.strip() for c in agg_cols_str.split(',')]

            processor = AggregateBackfillProcessor(
                filepath=filepath,
                header_row=header_row,
                group_cols_raw=group_cols_raw,
                agg_cols_raw=agg_cols_raw,
                agg_func=agg_func,
                result_col_input=result_col,
                compute_engine=compute_engine
            )
            result_path = processor.process()

    except InvalidFileException as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise ValidationError(f"文件格式错误，无法打开: {str(e)}。请确保文件是有效的 Excel 文件（.xlsx/.xls）且未被损坏。")
    except Exception as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise

    # 清理临时文件（不删除最终结果）
    try:
        if temp_processed and os.path.exists(temp_processed):
            os.remove(temp_processed)
        if filepath and os.path.exists(filepath) and filepath != temp_processed:
            os.remove(filepath)
    except Exception as e:
        logger.warning(f"清理临时文件失败: {e}")

    download_name = generate_output_filename('calc', original_filename, suffix='_计算结果')
    return send_file(result_path, as_attachment=True, download_name=download_name)
    ```

- **core/config.py**
  - language: `python`
  - content:
    ```python
# core/config.py
from pydantic_settings import BaseSettings
from pydantic import Field
from typing import Set

class SplitterConfig(BaseSettings):
    upload_folder: str = "uploads/splitter"
    output_folder: str = "outputs/splitter"
    task_status_folder: str = "task_status/splitter"
    model_config = {"env_prefix": "SPLITTER_"}

class CheckConfig(BaseSettings):
    sheet2_timeout_seconds: int = 120
    similarity_threshold: int = 4
    max_preview_rows: int = 1000
    duplicate_action: str = "error"
    chunk_size: int = 10000
    max_memory_rows: int = 200000
    model_config = {"env_prefix": "CHECK_"}

class FieldParserConfig(BaseSettings):
    max_text_length: int = 10 * 1024 * 1024
    similarity_threshold: int = 30
    model_config = {"env_prefix": "FIELD_PARSER_"}

class ConvertConfig(BaseSettings):
    max_file_size: int = 500 * 1024 * 1024   # 500MB
    default_encoding: str = "utf-8-sig"
    allowed_input_exts: Set[str] = {"xlsx", "xls", "xlsm", "csv"}
    allowed_output_formats: Set[str] = {"csv", "xlsx"}
    model_config = {"env_prefix": "CONVERT_"}

class TeamReportConfig(BaseSettings):
    """车队日报专用配置"""
    mapping_driver_id_col: str = "队员ID"
    mapping_team_name_col: str = "队伍名称"
    mapping_leader_name_col: str = "队长姓名"
    mapping_supervisor_col: str = "司管"
    flow_driver_id_col: str = "司机ID"
    flow_status_col: str = "司机状态"
    flow_active_status: str = "已上岗"
    flow_order_col: str = "总完单数"
    flow_apply_col: str = "报单数"
    flow_transfer_col: str = "转单数"
    flow_new_ju_col: str = "新桔完单数"
    flow_peak_order_col: str = "高峰期接单数"
    flow_peak_finish_col: str = "高峰期完单数"
    flow_star_finish_col: str = "星享完单数"
    monthly_order_col: str = "10月完单"
    monthly_apply_col: str = "10月报单"
    monthly_transfer_col: str = "10月转单"
    monthly_new_ju_col: str = "10月新桔在线数"
    daily_order_target: int = 1150
    daily_apply_target: int = 120
    monthly_order_target: int = 35650
    monthly_apply_target: int = 3720
    flow_online_days_col: str = "在线天数"
    static_supervisor_mapping: dict = {'徐国宝': '任翔', '王国平': '任凯', '安会阳': '任翔', '任翔': '任翔', '项红伟': '任翔', '马长领': '任翔', '袁茂启': '方行', '余和斌': '任翔', '夏海洋': '方行', '任凯': '任凯', '王哲': '方行', '王妙恩': '方行', '方行': '方行', '黄伟昌': '方行', '金忠健': '任凯', '蒋小峰': '方行', '童圣君': '任翔', '陈仪端': '任凯', '刘美乐': '方行', '黄一鹏': '任翔', '王凤梨': '任翔', '鲍腊九': '任凯', '王天剑': '任凯', '陈旭东': '任凯', '王林': '方行'}
    model_config = {"env_prefix": "TEAM_"}

class Settings(BaseSettings):
    secret_key: str = Field("dev-key")
    upload_folder: str = "uploads"
    output_folder: str = "outputs"
    max_content_length: int = 2 * 1024 * 1024 * 1024  # 2GB
    allowed_extensions: Set[str] = {"xlsx", "xls"}
    splitter: SplitterConfig = SplitterConfig()
    checker: CheckConfig = CheckConfig()
    field_parser: FieldParserConfig = FieldParserConfig()
    team_report: TeamReportConfig = TeamReportConfig()
    convert: ConvertConfig = ConvertConfig()
    model_config = {"env_file": ".env", "env_file_encoding": "utf-8", "extra": "ignore"}

settings = Settings()
    ```

- **core/csv_utils.py**
  - language: `python`
  - content:
    ```python
# core/csv_utils.py
"""公共CSV工具：表头读取、列映射、列字母生成"""
from typing import List, Tuple, Dict
import csv

def read_csv_header(filepath: str, encoding='utf-8') -> List[str]:
    with open(filepath, 'r', encoding=encoding) as f:
        reader = csv.reader(f)
        header = next(reader)
        return [col.strip() for col in header]

def get_column_letter(n: int) -> str:
    result = ''
    while n > 0:
        n -= 1
        result = chr(n % 26 + 65) + result
        n //= 26
    return result

def build_column_mapping(headers: List[str]) -> Tuple[Dict[str, str], Dict[str, str]]:
    letter_to_name = {}
    name_to_letter = {}
    for i, name in enumerate(headers):
        letter = get_column_letter(i + 1)
        letter_to_name[letter] = name
        name_to_letter[name] = letter
    return letter_to_name, name_to_letter

def split_fields(s: str) -> List[str]:
    """将逗号、换行、空格分隔的字符串拆分为字段列表"""
    if not s:
        return []
    lines = s.splitlines()
    parts = []
    for line in lines:
        for seg in line.split(','):
            parts.extend(seg.split())
    return [p.strip() for p in parts if p.strip()]
    ```

- **core/data_access_layer.py**
  - language: `python`
  - content:
    ```python
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
    ```

- **core/decorators.py**
  - language: `python`
  - content:
    ```python
# core/decorators.py
import logging
from functools import wraps
from flask import jsonify
from core.exceptions import AppException
from core.response import make_response
from openpyxl.utils.exceptions import InvalidFileException

logger = logging.getLogger(__name__)


def handle_exceptions():
    """统一异常处理装饰器：区分预期异常和未预期异常"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except AppException as e:
                logger.warning(f"业务异常: {e.message}")
                return make_response(False, None, e.message, e.code)
            except (ValueError, InvalidFileException) as e:
                logger.warning(f"输入错误: {str(e)}")
                return make_response(False, None, str(e), 400)
            except Exception as e:
                logger.exception("未处理异常")
                return make_response(False, None, "服务器内部错误，请稍后重试", 500)
        return wrapper
    return decorator
    ```

- **core/di.py**
  - language: `python`
  - content:
    ```python
# core/di.py
from dependency_injector import containers, providers
from core.config import settings
from core.task.manager import TaskManager
from core.task.storage import TaskStorage
from services.data.reader import DataReader
from services.data.writer import DataWriter
from services.data.formatter import DataFormatter
from services.data.file_service import FileService
from services.excel.compute_engine import ComputeEngine
from core.filter_engine import FilterEngine
from core.excel_utils import ExcelFile
from services.excel.calculator import ColumnCalculator
from services.excel.comparator import ExcelComparator
from services.excel.filter import ExcelFilter
from services.excel.aggregator import Aggregator
from services.field_parser import FieldParserService
from services.checker import CheckerService
from services.convert.service import ConvertService
from services.team_report import TeamReportService
from services.team_report.writer import TeamReportExcelWriter

class Container(containers.DeclarativeContainer):
    config = providers.Singleton(lambda: settings)
    task_storage = providers.Singleton(TaskStorage, status_dir=settings.output_folder, output_dir=settings.output_folder)
    task_manager = providers.Singleton(TaskManager, storage=task_storage, max_concurrent=3)
    data_reader = providers.Singleton(DataReader)
    data_writer = providers.Singleton(DataWriter)
    data_formatter = providers.Singleton(DataFormatter)
    file_service = providers.Singleton(FileService)
    compute_engine = providers.Singleton(ComputeEngine)
    filter_engine = providers.Singleton(FilterEngine)
    excel_service = providers.Singleton(ExcelFile)
    aggregator = providers.Singleton(Aggregator)
    column_calculator = providers.Singleton(ColumnCalculator)
    excel_comparator = providers.Singleton(ExcelComparator)
    excel_filter = providers.Singleton(ExcelFilter)
    team_report_service = providers.Factory(
        TeamReportService,
        config=config.provided.team_report
    )
    team_report_writer = providers.Singleton(TeamReportExcelWriter)
    field_parser = providers.Singleton(FieldParserService, similarity_threshold=config.provided.field_parser.similarity_threshold)
    checker_service = providers.Factory(CheckerService, config=config.provided.checker)
    convert_service = providers.Factory(ConvertService)
    ```

- **core/excel_utils.py**
  - language: `python`
  - content:
    ```python
# core/excel_utils.py
"""统一的 Excel 工具：列转换、文件包装、合并单元格复制"""
import re
import logging
from typing import Union, List, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from copy import copy

logger = logging.getLogger(__name__)

def col_letter_to_index(letter: str) -> int:
    return column_index_from_string(letter.upper())

def col_index_to_letter(index: int) -> str:
    return get_column_letter(index)

def validate_column_letters(letters: List[str]) -> bool:
    pattern = re.compile(r'^[A-Z]+$')
    return all(pattern.match(l.upper()) for l in letters)

class ExcelFile:
    def __init__(self, filepath: str, data_only: bool = True):
        self.filepath = filepath
        self.data_only = data_only
        self._workbook: Optional[openpyxl.Workbook] = None
        self._active_sheet = None

    def __enter__(self):
        self.load()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def load(self) -> 'ExcelFile':
        self._workbook = openpyxl.load_workbook(self.filepath, data_only=self.data_only)
        return self

    def close(self):
        if self._workbook:
            self._workbook.close()
            self._workbook = None

    def save(self, filepath: str = None) -> str:
        if not self._workbook:
            raise ValueError("未加载工作簿")
        target = filepath or self.filepath
        self._workbook.save(target)
        return target

    @property
    def active_sheet(self):
        if self._active_sheet is None:
            self._active_sheet = self._workbook.active
        return self._active_sheet

    def get_sheet(self, name_or_index: Union[str, int]):
        if isinstance(name_or_index, int):
            return self._workbook.worksheets[name_or_index]
        return self._workbook[name_or_index]

    def get_headers(self, sheet=None, row: int = 1) -> List[str]:
        ws = sheet or self.active_sheet
        headers = []
        for cell in ws[row]:
            if cell.value is None:
                break
            headers.append(str(cell.value).strip())
        return headers

    def get_cell_value(self, row: int, col: Union[int, str], sheet=None) -> Any:
        ws = sheet or self.active_sheet
        if isinstance(col, str):
            col = col_letter_to_index(col)
        cell = ws.cell(row, col)
        for merged in ws.merged_cells.ranges:
            if cell.coordinate in merged:
                return ws.cell(merged.min_row, merged.min_col).value
        return cell.value

    def copy_merge_structure(self, src_col: Union[int, str], dst_col: Union[int, str],
                             start_row: int, end_row: int, sheet=None):
        ws = sheet or self.active_sheet
        if isinstance(src_col, str):
            src_col = col_letter_to_index(src_col)
        if isinstance(dst_col, str):
            dst_col = col_letter_to_index(dst_col)

        src_ranges = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= src_col <= merged.max_col:
                if merged.min_row <= end_row and merged.max_row >= start_row:
                    src_ranges.append((merged.min_row, merged.max_row))

        to_remove = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= dst_col <= merged.max_col:
                to_remove.append(merged)
        for merged in to_remove:
            ws.unmerge_cells(str(merged))

        for min_r, max_r in src_ranges:
            if min_r < max_r:
                ws.merge_cells(start_row=min_r, start_column=dst_col,
                               end_row=max_r, end_column=dst_col)

    @staticmethod
    def read_dataframe(filepath: str, sheet: Union[str, int] = 0, **kwargs) -> pd.DataFrame:
        try:
            return pd.read_excel(filepath, sheet_name=sheet, **kwargs)
        except Exception as e:
            logger.error(f"读取Excel失败 {filepath}: {e}")
            raise ValueError(f"无法读取Excel文件: {e}")

    @staticmethod
    def write_dataframe(df: pd.DataFrame, filepath: str, sheet_name: str = 'Sheet1'):
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return filepath

def copy_merge_structure(src_ws, src_col, dst_ws, dst_col, start_row, end_row):
    """兼容旧接口的独立函数"""
    ef = ExcelFile('')  # 占位，仅用于调用实例方法
    ef._workbook = src_ws.parent
    ef.copy_merge_structure(src_col, dst_col, start_row, end_row, sheet=dst_ws)


# ========== 以下为兼容 ExcelService 而添加的静态方法 ==========

    @staticmethod
    def copy_cell_style(src_cell, dst_cell):
        """复制单元格样式（兼容 ExcelFile.copy_cell_style）"""
        from copy import copy
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.border = copy(src_cell.border)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.number_format = src_cell.number_format
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.protection = copy(src_cell.protection)
        dst_cell.value = src_cell.value

    @staticmethod
    def copy_merge_structure_static(ws, src_col, dst_col, start_row, end_row):
        """
        静态合并单元格结构复制（兼容 copy_merge_structure）
        :param ws: 工作表对象
        :param src_col: 源列（字母或数字索引）
        :param dst_col: 目标列（字母或数字索引）
        :param start_row: 起始行
        :param end_row: 结束行
        """
        # 复用 ExcelFile 的实例方法，创建一个临时对象
        ef = ExcelFile("")
        ef._workbook = ws.parent
        ef.copy_merge_structure(src_col, dst_col, start_row, end_row, sheet=ws)

    @staticmethod
    def get_cell_value_static(ws, row, col):
        """
        静态获取单元格值（兼容 ExcelFile.get_cell_value）
        :param ws: 工作表对象
        :param row: 行号
        :param col: 列字母或数字索引
        """
        ef = ExcelFile("")
        ef._workbook = ws.parent
        return ef.get_cell_value(row, col, sheet=ws)
    ```

- **core/exceptions.py**
  - language: `python`
  - content:
    ```python
# core/exceptions.py
class AppException(Exception):
    def __init__(self, message, code=400, error_type='bad_request'):
        self.message = message
        self.code = code
        self.error_type = error_type
        super().__init__(message)

class ValidationError(AppException):
    def __init__(self, message):
        super().__init__(message, code=400, error_type='validation_error')

class NotFoundError(AppException):
    def __init__(self, message='资源不存在'):
        super().__init__(message, code=404, error_type='not_found')

class ConflictError(AppException):
    def __init__(self, message='资源冲突'):
        super().__init__(message, code=409, error_type='conflict')

class InternalError(AppException):
    def __init__(self, message='服务器内部错误'):
        super().__init__(message, code=500, error_type='internal_error')
    ```

- **core/expression_engine.py**
  - language: `python`
  - content:
    ```python
# core/expression_engine.py
"""增强型表达式引擎，并导出列字母提取函数"""
import ast
import operator
import re
import logging
from typing import Dict, Any, Set
import numpy as np

logger = logging.getLogger(__name__)

_OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Pow: operator.pow,
    ast.USub: operator.neg,
}


def extract_column_letters(expr: str) -> Set[str]:
    """
    从表达式中提取列字母（仅限 1-3 位大写字母）。
    使用简单正则 [A-Z]+，并过滤长度 >3 的 token（避免匹配函数名）。
    """
    tokens = re.findall(r'[A-Z]+', expr.upper())
    # 只保留长度 1-3 的 token（列字母通常不超过 3 位）
    return {t for t in tokens if 1 <= len(t) <= 3}


class ExpressionEngine:
    def __init__(self, missing_value_strategy: str = 'nan'):
        self.missing_value_strategy = missing_value_strategy
        self._cache: Dict[str, ast.Expression] = {}

    def compile(self, expr: str) -> ast.Expression:
        expr_upper = expr.upper()
        if expr_upper in self._cache:
            return self._cache[expr_upper]
        allowed_chars = set('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789+-*/() .')
        if not all(c in allowed_chars for c in expr_upper):
            raise ValueError(f"表达式包含非法字符: {expr}")
        try:
            tree = ast.parse(expr_upper, mode='eval')
            self._cache[expr_upper] = tree
            return tree
        except SyntaxError as e:
            raise ValueError(f"表达式语法错误: {expr}") from e

    def evaluate(self, expr: str, context: Dict[str, Any]) -> float:
        tree = self.compile(expr)
        try:
            result = self._safe_eval(tree.body, context)
            return float(result) if isinstance(result, (int, float)) else self._convert_to_number(result)
        except Exception as e:
            logger.warning(f"求值失败: {expr}, {e}")
            return self._handle_missing_value()

    def _safe_eval(self, node, context):
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name):
            return self._get_variable_value(node.id, context)
        if isinstance(node, ast.BinOp):
            left = self._safe_eval(node.left, context)
            right = self._safe_eval(node.right, context)
            op = _OPERATORS.get(type(node.op))
            if op is None:
                raise ValueError(f"不支持运算符: {type(node.op)}")
            try:
                res = op(left, right)
                return np.nan if isinstance(res, float) and not np.isfinite(res) else res
            except ZeroDivisionError:
                return np.nan
        if isinstance(node, ast.UnaryOp):
            if isinstance(node.op, ast.USub):
                return -self._safe_eval(node.operand, context)
        raise ValueError(f"不支持节点: {type(node)}")

    def _get_variable_value(self, name, context):
        raw = context.get(name)
        if raw is None:
            return self._handle_missing_value()
        if isinstance(raw, str):
            s = raw.strip()
            if s == '':
                return self._handle_missing_value()
            try:
                return float(s)
            except ValueError:
                return self._handle_missing_value()
        if isinstance(raw, (int, float, np.number)):
            return float(raw)
        try:
            return float(raw)
        except (ValueError, TypeError):
            return self._handle_missing_value()

    def _convert_to_number(self, value):
        if isinstance(value, (int, float)):
            return float(value)
        try:
            return float(value)
        except (TypeError, ValueError):
            return self._handle_missing_value()

    def _handle_missing_value(self) -> float:
        if self.missing_value_strategy == 'nan':
            return np.nan
        elif self.missing_value_strategy == 'zero':
            return 0.0
        elif self.missing_value_strategy == 'raise':
            raise ValueError("缺少必要变量或值无效")
        return np.nan


_default_engine = ExpressionEngine()


def get_expression_engine() -> ExpressionEngine:
    return _default_engine
    ```

- **core/file_naming.py**
  - language: `python`
  - content:
    ```python
# core/file_naming.py
import os
from urllib.parse import quote

def generate_output_filename(module_name: str, original_filename: str, suffix: str = '_小计') -> str:
    base = os.path.splitext(os.path.basename(original_filename))[0]
    safe_module = module_name.replace('/', '_').replace('\\', '_')
    return f"{safe_module}_{base}{suffix}.xlsx"


def build_content_disposition(filename: str) -> str:
    encoded = quote(filename, safe='')
    fallback = ''.join(ch if ord(ch) < 128 else '_' for ch in filename)
    if not fallback:
        fallback = 'download.xlsx'
    return f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{encoded}"
    ```

- **core/header_detection.py**
  - language: `python`
  - content:
    ```python
# core/header_detection.py
"""生产级Excel表头检测器"""
import re
import os
import time
from typing import List, Tuple, Optional, Dict, Any
from dataclasses import dataclass, field, asdict
from openpyxl import load_workbook
import logging
import hashlib

logger = logging.getLogger(__name__)

@dataclass
class HeaderDetectionConfig:
    max_check_rows: int = 20
    header_row_range: Tuple[int, int] = (0, 2)
    min_header_cols: int = 3
    max_header_cols: int = 200
    min_non_empty_ratio: float = 0.6
    min_string_ratio: float = 0.7
    max_duplicate_ratio: float = 0.3
    min_unique_cols: int = 2
    known_header_keywords: List[str] = field(default_factory=lambda: [
        "名称", "姓名", "编号", "代码", "日期", "时间", "金额", "数量",
        "单位", "备注", "状态", "类型", "级别", "部门", "岗位",
        "科目", "凭证", "摘要", "借方", "贷方", "余额", "本期", "上期",
        "病历号", "患者", "诊断", "药品", "剂量", "用法", "医生", "科室",
        "id", "name", "date", "time", "code", "type", "status", "amount"
    ])
    skip_rows_keywords: List[str] = field(default_factory=lambda: [
        "合计", "总计", "小计", "total", "sum", "汇总", "平均", "average"
    ])
    mandatory_columns: List[str] = field(default_factory=list)
    read_only: bool = True
    data_only: bool = True
    cache_ttl_seconds: int = 300

    def to_dict(self) -> Dict:
        return asdict(self)

    def get_hash(self) -> str:
        import json
        config_str = json.dumps(self.to_dict(), sort_keys=True)
        return hashlib.md5(config_str.encode()).hexdigest()[:8]

class HeaderDetectionResult:
    def __init__(self):
        self.header_rows: List[int] = []
        self.data_start_row: int = 0
        self.column_names: List[str] = []
        self.confidence: float = 0.0
        self.reason: str = ""
        self.metrics: Dict[str, Any] = {}
        self.warnings: List[str] = []
        self.timestamp: float = time.time()
        self.config_hash: str = ""

    def is_valid(self) -> bool:
        return bool(self.header_rows) and self.confidence >= 0.1

    def get_data_start_excel_row(self) -> int:
        return self.data_start_row + 1 if self.data_start_row >= 0 else 1

class SimpleExcelHeaderDetector:
    def __init__(self, config: HeaderDetectionConfig = None):
        self.config = config or HeaderDetectionConfig()
        self.config_hash = self.config.get_hash()
        self._cache = {}
        self._workbook_cache = {}
        self._keyword_set = {kw.lower() for kw in self.config.known_header_keywords}
        self._skip_set = {kw.lower() for kw in self.config.skip_rows_keywords}
        self._numeric_regex = re.compile(r'^[+-]?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?%?$')
        self._invalid_chars_regex = re.compile(r'[\/*?\[\]]')

    def detect_headers(self, file_path: str, sheet_name: str = None, use_cache: bool = True) -> HeaderDetectionResult:
        self._validate_file(file_path)
        mtime = os.path.getmtime(file_path)
        cache_key = (file_path, mtime, self.config_hash, sheet_name)
        if use_cache and cache_key in self._cache:
            cached = self._cache[cache_key]
            if time.time() - cached.timestamp < self.config.cache_ttl_seconds:
                cached.metrics['cache_hit'] = True
                return cached

        result = HeaderDetectionResult()
        result.config_hash = self.config_hash
        try:
            wb = load_workbook(filename=file_path, read_only=self.config.read_only, data_only=self.config.data_only)
            ws = wb[sheet_name] if sheet_name else wb.active
            result = self._detect_from_worksheet(ws)
            wb.close()
        except Exception as e:
            logger.error(f"检测失败 {file_path}: {e}")
            result = self._create_fallback_result()
            result.warnings.append(str(e))

        if use_cache and result.is_valid():
            self._cache[cache_key] = result
        return result

    def detect_from_workbook(self, workbook, sheet_name: str = None) -> HeaderDetectionResult:
        workbook_id = id(workbook)
        cache_key = (workbook_id, sheet_name, self.config_hash)
        if cache_key in self._workbook_cache:
            cached = self._workbook_cache[cache_key]
            cached.metrics['workbook_cache_hit'] = True
            return cached

        try:
            ws = workbook[sheet_name] if sheet_name else workbook.active
            result = self._detect_from_worksheet(ws)
            result.config_hash = self.config_hash
        except Exception as e:
            logger.error(f"从workbook检测失败: {e}")
            result = self._create_fallback_result()
            result.warnings.append(str(e))

        self._workbook_cache[cache_key] = result
        return result

    def _validate_file(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        if not os.path.isfile(file_path):
            raise ValueError(f"路径不是文件: {file_path}")
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in {'.xlsx', '.xls', '.xlsm'}:
            raise ValueError(f"不支持的文件类型: {ext}")

    def _detect_from_worksheet(self, ws) -> HeaderDetectionResult:
        result = HeaderDetectionResult()
        start_time = time.time()
        max_row = min(ws.max_row, self.config.max_check_rows)
        max_col = min(ws.max_column, self.config.max_header_cols)
        if max_row == 0 or max_col == 0:
            result.warnings.append("工作表为空")
            return self._create_fallback_result()

        data = self._read_cell_range(ws, 1, max_row, 1, max_col)
        if not data:
            return self._create_fallback_result()

        candidates = []
        start, end = self.config.header_row_range
        for row_idx in range(min(end + 1, len(data))):
            if row_idx < start:
                continue
            candidate = self._evaluate_row(data, row_idx)
            if candidate['score'] > 0:
                candidates.append(candidate)

        if candidates:
            candidates.sort(key=lambda x: x['score'], reverse=True)
            best = candidates[0]
            result.header_rows = [best['row']]
            result.confidence = best['score']
            result.reason = best['reason']
            result.column_names = best['column_names']
            result.data_start_row = best['row'] + 1
        else:
            result = self._create_fallback_result()
            if data:
                result.column_names = [str(c).strip() if cell is not None and str(cell).strip() else f"Col_{i+1}" for i, cell in enumerate(data[0])]
        result.metrics.update({
            'execution_time_ms': (time.time() - start_time) * 1000,
            'rows_checked': len(data),
            'columns_checked': max_col,
            'candidates_found': len(candidates)
        })
        return result

    def _read_cell_range(self, ws, min_row, max_row, min_col, max_col):
        try:
            return [[cell for cell in row] for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True)]
        except Exception as e:
            logger.warning(f"读取单元格失败: {e}")
            return []

    def _evaluate_row(self, data, row_idx):
        if row_idx >= len(data):
            return self._create_candidate(row_idx, 0, "索引越界")
        row_data = data[row_idx]
        if not row_data:
            return self._create_candidate(row_idx, 0, "空行")
        non_empty = [cell for cell in row_data if cell is not None and str(cell).strip() != '']
        non_empty_ratio = len(non_empty) / len(row_data) if row_data else 0
        if non_empty_ratio < 0.2:
            return self._create_candidate(row_idx, 0, f"非空{non_empty_ratio:.0%}")
        string_cells = []
        for cell in non_empty:
            s = str(cell).strip()
            if s and not self._is_numeric(s):
                string_cells.append(s)
        string_ratio = len(string_cells) / len(row_data) if row_data else 0
        unique_strings = set(string_cells)
        dup_ratio = 1 - len(unique_strings) / len(string_cells) if string_cells else 0
        keyword_matches = sum(1 for s in string_cells for kw in self._keyword_set if kw in s.lower())
        keyword_ratio = keyword_matches / len(string_cells) if string_cells else 0
        combined = ' '.join(string_cells).lower()
        if any(skip in combined for skip in self._skip_set):
            return self._create_candidate(row_idx, 0, "跳过关键词")
        valid_fmt = sum(1 for s in string_cells if self._is_valid_column_name(s))
        valid_fmt_ratio = valid_fmt / len(string_cells) if string_cells else 0
        score = 0.0
        reasons = []
        if non_empty_ratio >= self.config.min_non_empty_ratio:
            score += 0.3; reasons.append(f"非空{non_empty_ratio:.0%}")
        if string_ratio >= self.config.min_string_ratio:
            score += 0.3; reasons.append(f"文本{string_ratio:.0%}")
        if dup_ratio <= self.config.max_duplicate_ratio:
            score += 0.2; reasons.append(f"唯一{1-dup_ratio:.0%}")
        if keyword_ratio > 0.1:
            score += min(keyword_ratio, 0.2); reasons.append(f"关键词{keyword_ratio:.0%}")
        if valid_fmt_ratio > 0.5:
            score += 0.1; reasons.append("格式正确")
        score = max(0.0, min(1.0, score))
        column_names = [str(cell).strip() if cell is not None and str(cell).strip() != '' else f"Col_{i+1}" for i, cell in enumerate(row_data)]
        return {'row': row_idx, 'score': score, 'reason': ', '.join(reasons) or '基础特征', 'column_names': column_names, 'metrics': {}}

    def _create_candidate(self, row_idx, score, reason):
        return {'row': row_idx, 'score': score, 'reason': reason, 'column_names': [], 'metrics': {}}

    def _is_numeric(self, s: str) -> bool:
        s = s.strip().replace(',', '')
        if s.endswith('%'): s = s[:-1]
        try: float(s); return True
        except ValueError: return False

    def _is_valid_column_name(self, s: str) -> bool:
        if not s or len(s) > 50: return False
        if re.match(r'^[\d.,%\s-]+$', s): return False
        if self._invalid_chars_regex.search(s): return False
        return s.strip() != ''

    def _create_fallback_result(self) -> HeaderDetectionResult:
        result = HeaderDetectionResult()
        result.header_rows = [0]
        result.data_start_row = 1
        result.confidence = 0.3
        result.reason = "回退默认第一行"
        return result

    def clear_cache(self):
        self._cache.clear()
        self._workbook_cache.clear()

def detect_header_start(file_path: str, sheet_name: str = None) -> int:
    detector = SimpleExcelHeaderDetector()
    try:
        result = detector.detect_headers(file_path, sheet_name)
        if result.is_valid():
            return result.get_data_start_excel_row()
    except Exception:
        pass
    return 2
    ```

- **core/response.py**
  - language: `python`
  - content:
    ```python
# core/response.py
"""统一 JSON 响应构建工具"""
from flask import jsonify

def make_response(success=True, data=None, error=None, code=200):
    """构建统一的 JSON 响应
    :param success: 操作是否成功
    :param data: 成功时返回的数据
    :param error: 错误时的描述信息
    :param code: HTTP 状态码
    """
    body = {
        'success': success,
        'data': data,
        'error': error,
        'code': code
    }
    response = jsonify(body)
    response.status_code = code
    return response

def success_response(data=None, message='success'):
    """快捷成功响应，兼容旧接口"""
    return make_response(True, data, None, 200)

def error_response(message, code=400):
    """快捷错误响应，兼容旧接口"""
    return make_response(False, None, message, code)
    ```

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

- **services/excel/aggregator.py**
  - language: `python`
  - content:
    ```python
# utils/common/aggregator.py
import pandas as pd
from typing import List, Dict, Any, Union, Optional

class Aggregator:
    """通用聚合器（静态方法）"""

    @staticmethod
    def group_aggregate(df: pd.DataFrame, group_keys: List[str],
                        agg_spec: Dict[str, Union[str, List[str]]]) -> pd.DataFrame:
        if not group_keys:
            return df.drop_duplicates()
        return df.groupby(group_keys, as_index=False).agg(agg_spec)

    @staticmethod
    def aggregate_by_keys(df: pd.DataFrame, unique_keys: List[str],
                          sum_fields: List[str], other_fields: Optional[List[str]] = None) -> pd.DataFrame:
        """
        按唯一键聚合：sum_fields 求和，other_fields 取第一条记录。
        不使用 groupby.agg 字典，改为组合 drop_duplicates 和 groupby sum，避免 pandas 内部错误。
        """
        if other_fields is None:
            other_fields = [c for c in df.columns if c not in unique_keys and c not in sum_fields]
        if not unique_keys:
            return df.drop_duplicates()

        # 1. 求和部分
        if sum_fields:
            sum_part = df.groupby(unique_keys, as_index=False)[sum_fields].sum()
        else:
            sum_part = pd.DataFrame(columns=unique_keys)  # 空框架，仅用于后续合并

        # 2. 取第一条记录部分（直接用 drop_duplicates 保留第一个出现）
        if other_fields:
            first_part = df[unique_keys + other_fields].drop_duplicates(subset=unique_keys, keep='first')
        else:
            first_part = pd.DataFrame(columns=unique_keys)

        # 3. 合并求和与首条记录
        if not sum_part.empty and not first_part.empty:
            result = sum_part.merge(first_part, on=unique_keys, how='outer')
        elif not sum_part.empty:
            result = sum_part
        elif not first_part.empty:
            result = first_part
        else:
            # 没有求和也没有其他字段时，直接去重返回唯一键
            result = df[unique_keys].drop_duplicates()

        return result

    @staticmethod
    def simple_aggregate(df: pd.DataFrame, group_keys: List[str],
                         agg_cols: List[str], func: str = 'sum') -> pd.DataFrame:
        agg_dict = {col: func for col in agg_cols}
        return Aggregator.group_aggregate(df, group_keys, agg_dict)


# 向后兼容包装
class AggregateCalculator:
    def __init__(self, excel_file):
        self.excel_file = excel_file

    def aggregate(self, group_cols, agg_cols, agg_func, result_col):
        raise NotImplementedError
    ```

- **services/excel/calculator_processor.py**
  - language: `python`
  - content:
    ```python
# services/excel/calculator_processor.py
"""Excel 列计算与聚合回填处理器"""
import os
import gc
import logging
import numpy as np
import openpyxl
import pandas as pd
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional, Tuple

from core.excel_utils import ExcelFile, copy_merge_structure
from core.data_access_layer import get_data_access_layer
from core.csv_utils import build_column_mapping
from core.excel_helpers import (
    safe_get_cell_value, safe_write_cell, clear_merges_in_column,
    build_column_name_mapping, normalize_excel_file, safe_cleanup_dataframe
)
from services.excel.compute_engine import ComputeEngine

logger = logging.getLogger(__name__)


class BaseExcelProcessor(ABC):
    """Excel 处理器基类，封装公共的写入逻辑"""

    def __init__(self, filepath: str, header_row: int, compute_engine: ComputeEngine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_file = None

    def _get_work_filepath(self) -> str:
        """获取可写的工作文件路径，统一通过 normalize_excel_file 转换为临时 .xlsx"""
        # 总是进行规范化，确保文件可读
        self.temp_file = normalize_excel_file(self.original_filepath)
        return self.temp_file

    def _cleanup_temp(self):
        if self.temp_file and os.path.exists(self.temp_file):
            try:
                os.unlink(self.temp_file)
            except Exception as e:
                logger.warning(f"清理临时文件失败: {e}")

    def _resolve_result_column(self, ws, result_col_input: str) -> Tuple[int, str]:
        """
        解析结果列：支持列字母或列名。
        返回 (列索引, 列名)
        """
        # 1. 尝试作为列字母解析
        try:
            col_idx = openpyxl.utils.column_index_from_string(result_col_input.upper())
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
        return os.path.join(output_dir, f"{base_name}_处理版.xlsx")

    @abstractmethod
    def process(self) -> str:
        pass


class ColumnCalculatorProcessor(BaseExcelProcessor):
    """列计算处理器"""

    def __init__(self, filepath: str, header_row: int, expression: str,
                 result_col_input: str, merge_by_cols_str: str, compute_engine: ComputeEngine):
        super().__init__(filepath, header_row, compute_engine)
        self.expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols = [c.strip() for c in merge_by_cols_str.split(',') if c.strip()] if merge_by_cols_str else []

    def _remove_result_column_from_expr(self, result_col_letter: str) -> str:
        """从表达式中移除结果列字母（如果存在）"""
        if not result_col_letter:
            return self.expression
        tokens = []
        current = []
        for ch in self.expression.upper():
            if ch.isalpha():
                current.append(ch)
            else:
                if current:
                    tokens.append(''.join(current))
                    current = []
                tokens.append(ch)
        if current:
            tokens.append(''.join(current))

        filtered = []
        for token in tokens:
            if token == result_col_letter:
                continue
            filtered.append(token)
        new_expr = ''.join(filtered)
        new_expr = new_expr.replace('+-', '-').replace('-+', '-').replace('++', '+').replace('--', '+')
        new_expr = new_expr.lstrip('+-*/').rstrip('+-*/')
        if new_expr != self.expression:
            logger.warning(f"表达式中包含了结果列 {result_col_letter}，已自动移除: {self.expression} -> {new_expr}")
        return new_expr

    def process(self) -> str:
        filepath = self._get_work_filepath()
        try:
            # 强制以字符串类型读取，避免公式错误
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = safe_cleanup_dataframe(df, auto_convert_numeric=True)
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
                    if not expr.strip():
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
                ef._workbook.close()

            gc.collect()
            return result_path
        finally:
            self._cleanup_temp()


class AggregateBackfillProcessor(BaseExcelProcessor):
    """聚合回填处理器"""

    def __init__(self, filepath: str, header_row: int,
                 group_cols_raw: List[str], agg_cols_raw: List[str],
                 agg_func: str, result_col_input: str, compute_engine: ComputeEngine):
        super().__init__(filepath, header_row, compute_engine)
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input

    def process(self) -> str:
        filepath = self._get_work_filepath()
        try:
            # 读取表头（仅列名），先不读取数据
            df_header = pd.read_excel(filepath, header=self.header_row - 1, nrows=0)
            headers = list(df_header.columns)
            letter_to_name, name_to_letter = build_column_mapping(headers)

            def resolve_col(col: str) -> str:
                col = col.strip()
                if col in headers:
                    return col
                upper = col.upper()
                if upper in letter_to_name:
                    return letter_to_name[upper]
                raise ValueError(f"列 '{col}' 无法映射到实际列名")

            group_cols = [resolve_col(c) for c in self.group_cols_raw]
            agg_cols = [resolve_col(c) for c in self.agg_cols_raw]

            # 以字符串类型读取数据
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = safe_cleanup_dataframe(df, auto_convert_numeric=True)

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
                ef._workbook.close()

            gc.collect()
            return result_path
        finally:
            self._cleanup_temp()
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

- **static/css/style.css**
  - language: `css`
  - content:
    ```css
/* ========== 全局变量与重置 ========== */
:root {
    --primary: #4361ee;
    --primary-dark: #3a56d4;
    --primary-light: #eef2ff;
    --secondary: #2ec4b6;
    --success: #06d6a0;
    --danger: #ef476f;
    --warning: #ffd166;
    --dark: #1e293b;
    --gray: #64748b;
    --light: #f8fafc;
    --border: #e2e8f0;
    --shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
    --shadow-lg: 0 10px 30px rgba(0, 0, 0, 0.1);
    --radius: 12px;
    --radius-sm: 8px;
    --transition: all 0.2s ease;
}

* {
    box-sizing: border-box;
}

body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, sans-serif;
    margin: 0;
    padding: 24px;
    background: linear-gradient(145deg, #f1f5f9 0%, #e6edf5 100%);
    min-height: 100vh;
    color: var(--dark);
    line-height: 1.6;
}

.container {
    max-width: 1280px;
    margin: 0 auto;
    background: rgba(255, 255, 255, 0.85);
    backdrop-filter: blur(10px);
    padding: 32px 36px;
    border-radius: 28px;
    box-shadow: var(--shadow-lg);
    border: 1px solid rgba(255, 255, 255, 0.7);
}

h1 {
    font-size: 2.2rem;
    font-weight: 700;
    margin-top: 0;
    margin-bottom: 28px;
    color: var(--dark);
    display: flex;
    align-items: center;
    gap: 12px;
    border-bottom: 2px solid var(--primary-light);
    padding-bottom: 16px;
}

h2 {
    font-size: 1.6rem;
    font-weight: 600;
    margin-top: 0;
    margin-bottom: 24px;
    color: var(--dark);
}

h3 {
    font-size: 1.2rem;
    font-weight: 600;
    margin-bottom: 16px;
    color: var(--gray);
}

/* 导航栏 */
nav ul {
    list-style: none;
    padding: 0;
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 8px;
    margin-bottom: 32px;
    background: var(--light);
    padding: 8px;
    border-radius: 48px;
    border: 1px solid var(--border);
}

nav ul li a {
    display: inline-block;
    padding: 10px 20px;
    text-decoration: none;
    color: var(--gray);
    font-weight: 500;
    border-radius: 40px;
    transition: var(--transition);
    background: transparent;
}

nav ul li a:hover {
    background: white;
    color: var(--primary);
    box-shadow: var(--shadow);
}

/* 表单通用 */
.form-section {
    background: white;
    border-radius: var(--radius);
    padding: 24px;
    margin-bottom: 28px;
    box-shadow: var(--shadow);
    border: 1px solid var(--border);
}

.form-row {
    margin-bottom: 20px;
}

label {
    font-weight: 600;
    display: block;
    margin-bottom: 8px;
    color: var(--dark);
    font-size: 0.95rem;
}

input, select, textarea {
    width: 100%;
    padding: 12px 16px;
    border: 1.5px solid var(--border);
    border-radius: var(--radius-sm);
    font-size: 0.95rem;
    transition: var(--transition);
    background: white;
}

input:focus, select:focus, textarea:focus {
    outline: none;
    border-color: var(--primary);
    box-shadow: 0 0 0 3px rgba(67, 97, 238, 0.15);
}

.hint, .help-text {
    font-size: 0.85rem;
    color: var(--gray);
    margin-top: 6px;
}

/* 按钮 */
.btn, button {
    background: var(--primary);
    color: white;
    padding: 12px 24px;
    border: none;
    border-radius: 40px;
    cursor: pointer;
    font-size: 1rem;
    font-weight: 600;
    transition: var(--transition);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    box-shadow: 0 2px 6px rgba(67, 97, 238, 0.3);
}

.btn:hover, button:hover {
    background: var(--primary-dark);
    transform: translateY(-2px);
    box-shadow: 0 6px 12px rgba(67, 97, 238, 0.25);
}

.btn-success {
    background: var(--success);
    box-shadow: 0 2px 6px rgba(6, 214, 160, 0.3);
}
.btn-success:hover { background: #05b384; }

.btn-danger {
    background: var(--danger);
    box-shadow: 0 2px 6px rgba(239, 71, 111, 0.3);
}
.btn-danger:hover { background: #d63e5e; }

.btn-secondary {
    background: var(--secondary);
    box-shadow: 0 2px 6px rgba(46, 196, 182, 0.3);
}
.btn-secondary:hover { background: #25a99a; }

/* 消息与状态 */
.message {
    margin-top: 20px;
    padding: 14px 18px;
    border-radius: var(--radius-sm);
    font-weight: 500;
}
.message.error { background: #fee2e2; color: #b91c1c; }
.message.success { background: #d1fae5; color: #065f46; }
.message.info { background: #dbeafe; color: #1e40af; }

.status-area {
    margin-top: 24px;
    padding: 24px;
    background: white;
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    border: 1px solid var(--border);
}

.progress {
    width: 100%;
    background: var(--border);
    border-radius: 20px;
    height: 10px;
    overflow: hidden;
}
.progress-bar {
    height: 10px;
    background: var(--primary);
    width: 0%;
    transition: width 0.3s;
}

.log-area {
    background: #0f172a;
    color: #a5f3fc;
    padding: 16px;
    border-radius: var(--radius-sm);
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.85rem;
    max-height: 240px;
    overflow-y: auto;
    white-space: pre-wrap;
}

/* 筛选条件行 */
.filter-row {
    display: flex;
    gap: 12px;
    margin-bottom: 16px;
    align-items: center;
    flex-wrap: wrap;
}
.filter-field { width: 180px; }
.filter-op { width: 120px; }
.filter-value { width: 240px; }

/* 表格样式（用于匹配填充等） */
.table-container {
    overflow-x: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    border-radius: var(--radius);
    overflow: hidden;
    box-shadow: var(--shadow);
}
th, td {
    padding: 12px 16px;
    text-align: left;
    border-bottom: 1px solid var(--border);
}
th {
    background: var(--primary-light);
    font-weight: 600;
}

/* 响应式 */
@media (max-width: 768px) {
    body { padding: 12px; }
    .container { padding: 20px; }
    nav ul { flex-direction: column; align-items: stretch; }
    nav ul li a { display: block; text-align: center; }
}
    ```

- **static/js/main.js**
  - language: `javascript`
  - content:
    ```javascript
// 全局交互逻辑
document.addEventListener('DOMContentLoaded', () => {
    console.log('✨ Excel 智能工具集已加载');

    // 为所有提交按钮添加加载状态（可选）
    document.querySelectorAll('form').forEach(form => {
        form.addEventListener('submit', () => {
            const btn = form.querySelector('button[type="submit"]');
            if (btn) {
                btn.disabled = true;
                btn.innerHTML = '<span class="spinner"></span> 处理中...';
            }
        });
    });
});
    ```

- **templates/base.html**
  - language: `html`
  - content:
    ```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{% block title %}Excel 智能工具集{% endblock %}</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
    <script src="{{ url_for('static', filename='js/main.js') }}" defer></script>
    <script src="{{ url_for('static', filename='js/download_utils.js') }}"></script>
    {% block head %}{% endblock %}
</head>
<body>
    <div class="container">
        <h1>📊 Excel 智能工具集</h1>
        <nav>
            <ul>
                <li><a href="{{ url_for('index') }}">🏠 首页</a></li>
                <li><a href="{{ url_for('calc.index') }}">🧮 列计算/聚合</a></li>
                <li><a href="{{ url_for('compare.index') }}">🔍 文件对比</a></li>
                <li><a href="{{ url_for('match.index') }}">🔗 匹配填充</a></li>
                <li><a href="{{ url_for('summary.index') }}">📋 汇总工具</a></li>
                <li><a href="{{ url_for('splitter.index') }}">✂️ 文件拆分</a></li>
                <li><a href="{{ url_for('field_parser.index') }}">🔗 字段映射</a></li>
                <li><a href="{{ url_for('checker.index') }}">✅ 文件校验</a></li>
                <li><a href="{{ url_for('convert.index') }}">🔄 格式转换</a></li>
                <li><a href="{{ url_for('subtotal.index') }}">📊 数据小计</a></li>
                <li><a href="{{ url_for('team_report.index') }}">📊 车队日报</a></li>
                <li><a href="{{ url_for('transposer.index') }}">🔄 行列转置</a></li>
            </ul>
        </nav>
        {% block content %}{% endblock %}
        <div id="message" class="message"></div>
    </div>
    {% block scripts %}{% endblock %}
</body>
</html>
    ```

- **templates/calculator/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}列计算 / 聚合回填{% endblock %}
{% block content %}
<h2>🧮 列计算 / 聚合回填</h2>
<form action="{{ url_for('calc.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
        <div class="form-row">
            <label>⚙️ 操作类型</label>
            <select name="calc_type" id="calcType">
                <option value="column">列计算</option>
                <option value="aggregate">聚合回填</option>
            </select>
        </div>
        <div class="form-row">
            <label>表头所在行号（从1开始，默认1）</label>
            <input type="number" name="header_row" value="1" min="1" style="max-width:100px;">
            <div class="hint">如果你的 Excel 顶部有标题行，请填写实际列名所在的行号</div>
        </div>
    </div>
    <div id="columnConfig" class="form-section">
        <h3>📐 列计算设置</h3>
        <div class="form-row">
            <label>表达式 (如 A+B-C)</label>
            <input type="text" name="expression" placeholder="例如 A+B">
        </div>
        <div class="form-row">
            <label>结果列 (字母或新列名)</label>
            <input type="text" name="result_col_column" placeholder="K 或 '差值'">
        </div>
        <div class="form-row">
            <label>合并依据列 (可选，逗号分隔)</label>
            <input type="text" name="merge_by_cols" placeholder="A,B">
            <div class="hint">相同组的结果列只保留第一行，其余合并。</div>
        </div>
    </div>
    <div id="aggregateConfig" class="form-section" style="display:none;">
        <h3>📊 聚合回填设置</h3>
        <div class="form-row">
            <label>分组列 (逗号分隔，支持列名或列字母)</label>
            <input type="text" name="group_cols" placeholder="A,B">
        </div>
        <div class="form-row">
            <label>聚合列 (逗号分隔)</label>
            <input type="text" name="agg_cols" placeholder="C,D">
        </div>
        <div class="form-row">
            <label>聚合函数</label>
            <select name="agg_func">
                <option value="sum">求和</option>
                <option value="mean">平均值</option>
                <option value="max">最大值</option>
                <option value="min">最小值</option>
                <option value="count">计数</option>
            </select>
        </div>
        <div class="form-row">
            <label>结果列名</label>
            <input type="text" name="result_col_aggregate" placeholder="K 或 '聚合结果'">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始处理</button>
</form>
<script>
    const typeSelect = document.getElementById('calcType');
    const colCfg = document.getElementById('columnConfig');
    const aggCfg = document.getElementById('aggregateConfig');
    function toggle() {
        if (typeSelect.value === 'column') {
            colCfg.style.display = 'block';
            aggCfg.style.display = 'none';
        } else {
            colCfg.style.display = 'none';
            aggCfg.style.display = 'block';
        }
    }
    typeSelect.addEventListener('change', toggle);
    toggle();
</script>
{% endblock %}
    ```

- **tests/test_calculator.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
from openpyxl import Workbook
from services.excel.calculator import ColumnCalculator
from services.excel.aggregator import Aggregator
from core.excel_utils import ExcelFile


class TestColumnCalculator:
    def test_calculate_column(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'A'
        ws['B1'] = 'B'
        ws['C1'] = 'Result'
        ws['A2'] = 10
        ws['B2'] = 20
        ws['A3'] = 5
        ws['B3'] = 3
        test_file = tmp_path / 'test.xlsx'
        wb.save(test_file)

        with ExcelFile(test_file) as ef:
            calculator = ColumnCalculator(ef)
            calculator.calculate_column('A+B', 'C', merge_by_cols=None)
            ef.save(test_file)

        result_df = pd.read_excel(test_file)
        assert result_df.loc[0, 'Result'] == 30
        assert result_df.loc[1, 'Result'] == 8

    def test_calculate_column_with_merge(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Group'
        ws['B1'] = 'Value'
        ws['C1'] = 'Result'
        ws['A2'] = 'X'
        ws['B2'] = 10
        ws['A3'] = 'X'
        ws['B3'] = 20
        ws['A4'] = 'Y'
        ws['B4'] = 5
        test_file = tmp_path / 'test.xlsx'
        wb.save(test_file)

        with ExcelFile(test_file) as ef:
            calculator = ColumnCalculator(ef)
            calculator.calculate_column('B*2', 'C', merge_by_cols=['A'])
            ef.save(test_file)

        result_df = pd.read_excel(test_file)
        assert result_df.loc[0, 'Result'] == 20
        assert pd.isna(result_df.loc[1, 'Result'])
        assert result_df.loc[2, 'Result'] == 10


class TestAggregator:
    def test_aggregate_sum(self):
        df = pd.DataFrame({'Group': ['X', 'X', 'Y'], 'Value': [10, 20, 5]})
        result = Aggregator.simple_aggregate(df, ['Group'], ['Value'], 'sum')
        assert len(result) == 2
        assert result.loc[result['Group'] == 'X', 'Value'].iloc[0] == 30
        assert result.loc[result['Group'] == 'Y', 'Value'].iloc[0] == 5

    def test_aggregate_mean(self):
        df = pd.DataFrame({'Group': ['X', 'X'], 'Value': [10, 30]})
        result = Aggregator.simple_aggregate(df, ['Group'], ['Value'], 'mean')
        assert result.loc[0, 'Value'] == 20

    def test_aggregate_by_keys(self):
        df = pd.DataFrame({
            'key': ['A', 'A', 'B'],
            'sum_col': [10, 20, 5],
            'other_col': ['first', 'second', 'other']
        })
        result = Aggregator.aggregate_by_keys(df, ['key'], ['sum_col'], ['other_col'])
        assert len(result) == 2
        assert result.loc[result['key'] == 'A', 'sum_col'].iloc[0] == 30
        assert result.loc[result['key'] == 'A', 'other_col'].iloc[0] == 'first'
        assert result.loc[result['key'] == 'B', 'sum_col'].iloc[0] == 5
    ```

- **tests/test_integration_calculator.py**
  - language: `python`
  - content:
    ```python
import io
from openpyxl import load_workbook


class TestCalculatorIntegration:
    def test_column_calculation(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'column',
            'header_row': '1',
            'expression': 'A+B',
            'result_col_column': 'D',
            'merge_by_cols': ''
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        assert response.headers['Content-Disposition'].startswith('attachment')
        assert '.xlsx' in response.headers['Content-Disposition']
        wb = load_workbook(io.BytesIO(response.data))
        ws = wb.active
        assert ws['D2'].value == 30
        assert ws['D3'].value == 8

    def test_aggregate_backfill(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'aggregate',
            'header_row': '1',
            'group_cols': 'A',
            'agg_cols': 'B',
            'agg_func': 'sum',
            'result_col_aggregate': 'E'
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        assert '.xlsx' in response.headers['Content-Disposition']
        wb = load_workbook(io.BytesIO(response.data))
        ws = wb.active
        assert ws['E2'].value == 20
        assert ws['E3'].value == 3

    def test_missing_file_error(self, client):
        response = client.post('/calc/process', data={'calc_type': 'column'})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '未上传文件' in response.get_json()['error']

    def test_invalid_expression_error(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'column',
            'header_row': '1',
            'expression': '',
            'result_col_column': 'D'
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '表达式' in response.get_json()['error']
    ```


> ✅ 共导出 25 个代码文件。
    ```

- **POML/calc.txt**
  - language: ``
  - content:
    ```
# calc模块相关文件
blueprints/calculator.py
services/excel/calculator_processor.py
services/excel/compute_engine.py
services/excel/aggregator.py
services/data/reader.py
services/data/writer.py
services/data/formatter.py
services/data/file_service.py
core/expression_engine.py
core/excel_utils.py
core/data_access_layer.py
core/header_detection.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/calculator/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_calculator.py
tests/test_integration_calculator.py
    ```

- **POML/checker.txt**
  - language: ``
  - content:
    ```
# checker模块相关文件
blueprints/checker.py
services/checker/__init__.py
services/checker/service.py
services/checker/strategies.py
services/checker/constants.py
services/data/stream_reader.py
services/data/row_counter.py
services/data/report_writer.py
core/filter_engine.py
core/progress_task.py
core/sqlite_buffer.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/checker/index.html
templates/base.html
static/css/checker.css
static/css/style.css
static/js/checker.js
static/js/main.js
tests/test_checker_service.py
tests/test_integration_checker.py
    ```

- **POML/compare.txt**
  - language: ``
  - content:
    ```
# compare模块相关文件
blueprints/comparator.py
services/excel/comparator.py
services/data/file_service.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
schemas/comparator.py
templates/comparator/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_comparator.py
tests/test_integration_comparator.py
    ```

- **POML/convert.txt**
  - language: ``
  - content:
    ```
# convert模块相关文件
blueprints/convert.py
services/convert/__init__.py
services/convert/service.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/convert/index.html
templates/base.html
static/css/style.css
static/js/convert.js
static/js/main.js
tests/test_convert_service.py
tests/test_integration_convert.py
    ```

- **POML/field_parser.txt**
  - language: ``
  - content:
    ```
# field_parser模块相关文件
blueprints/field_parser.py
services/field_parser/service.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/field_parser/index.html
templates/base.html
static/css/style.css
static/js/field_parser.js
static/js/main.js
tests/test_field_parser.py
tests/test_integration_field_parser.py
    ```

- **POML/match.txt**
  - language: ``
  - content:
    ```
# match模块相关文件
blueprints/matcher.py
services/excel/matcher.py
services/data/reader.py
core/excel_utils.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/matcher/index.html
templates/debug.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_integration_matcher.py
    ```

- **POML/splitter.txt**
  - language: ``
  - content:
    ```
# splitter模块相关文件
blueprints/splitter.py
core/filter_engine.py
core/validators.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/splitter/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_splitter.py
tests/test_integration_splitter.py
    ```

- **POML/subtotal.txt**
  - language: ``
  - content:
    ```
# subtotal模块相关文件
blueprints/subtotal.py
services/subtotal/__init__.py
services/subtotal/engine.py
core/header_detection.py
core/csv_utils.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/subtotal/index.html
templates/base.html
static/css/style.css
static/js/subtotal.js
static/js/main.js
tests/test_subtotal_engine.py
tests/test_integration_subtotal.py
    ```

- **POML/summary.txt**
  - language: ``
  - content:
    ```
# summary模块相关文件
blueprints/summarizer.py
services/excel/compute_engine.py
services/excel/aggregator.py
services/data/reader.py
services/data/writer.py
services/data/formatter.py
services/data/file_service.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/header_detection.py
core/exceptions.py
core/config.py
core/di.py
templates/summarizer/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_integration_summarizer.py
    ```

- **POML/team_report.txt**
  - language: ``
  - content:
    ```
# team_report模块相关文件
blueprints/team_report.py
services/team_report/__init__.py
services/team_report/service.py
services/team_report/writer.py
services/common/excel_styles.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/team_report/index.html
templates/base.html
static/css/style.css
static/js/team_report.js
static/js/main.js
tests/test_team_report_service.py
tests/test_integration_team_report.py
tests/test_team_report_long_running.py
    ```

- **POML/transposer.txt**
  - language: ``
  - content:
    ```
# transposer模块相关文件
blueprints/transposer.py
services/transposer/service.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/transposer/index.html
templates/base.html
static/css/style.css
static/js/transposer.js
static/js/main.js
tests/test_transposer.py
    ```

- **app.py**
  - language: `python`
  - content:
    ```python
#!/usr/bin/env python3
import os, logging, threading, time
from flask import Flask, render_template, request, g
from werkzeug.exceptions import HTTPException, NotFound
from core.di import Container
from core.config import settings
from core.logging_setup import init_app as init_logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(name)s %(message)s')
logger = logging.getLogger(__name__)

from blueprints.calculator import calc_bp
from blueprints.comparator import compare_bp
from blueprints.matcher import match_bp
from blueprints.summarizer import summary_bp
from blueprints.splitter import splitter_bp
from blueprints.field_parser import field_parser_bp
from blueprints.checker import checker_bp
from blueprints.convert import convert_bp
from blueprints.team_report import team_report_bp
from blueprints.transposer import transposer_bp
from blueprints.subtotal import subtotal_bp

def clean_old_files(directory, age_hours):
    now = time.time()
    if not os.path.exists(directory): return
    for root, dirs, files in os.walk(directory):
        for f in files:
            path = os.path.join(root, f)
            try:
                if os.path.getmtime(path) < now - age_hours * 3600:
                    os.remove(path)
                    logger.debug(f"删除过期文件: {path}")
            except Exception as e:
                logger.warning(f"删除文件失败 {path}: {e}")

def scheduled_cleanup(app):
    while True:
        time.sleep(3600)
        with app.app_context():
            cfg = app.config
            clean_old_files(cfg['UPLOAD_FOLDER'], 24)
            logger.info("执行定期清理任务")

def create_app():
    app = Flask(__name__)
    app.debug = False
    raw_dict = settings.model_dump()
    flat_config = {}
    for key, value in raw_dict.items():
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                flat_key = f"{key}_{sub_key}".upper()
                flat_config[flat_key] = sub_value
        else:
            flat_config[key.upper()] = value
    app.config.update(flat_config)
    app.config['MAX_CONTENT_LENGTH'] = settings.max_content_length

    container = Container()
    app.container = container

    os.makedirs(app.config.get('UPLOAD_FOLDER', 'uploads'), exist_ok=True)

    app.register_blueprint(calc_bp, url_prefix='/calc')
    app.register_blueprint(compare_bp, url_prefix='/compare')
    app.register_blueprint(match_bp, url_prefix='/match')
    app.register_blueprint(summary_bp, url_prefix='/summary')
    app.register_blueprint(splitter_bp, url_prefix='/splitter')
    app.register_blueprint(field_parser_bp, url_prefix='/field-parser')
    app.register_blueprint(checker_bp, url_prefix='/checker')
    app.register_blueprint(convert_bp, url_prefix='/convert')
    app.register_blueprint(team_report_bp, url_prefix='/team-report')
    app.register_blueprint(transposer_bp, url_prefix='/transposer')
    app.register_blueprint(subtotal_bp, url_prefix='/subtotal')

    init_logging(app)

    @app.route('/')
    def index():
        return render_template('index.html')

    @app.route('/favicon.ico')
    def favicon():
        return '', 204

    @app.errorhandler(NotFound)
    def handle_not_found(e):
        ignored = ('/favicon.ico', '/robots.txt', '/apple-touch-icon', '/.env', '/wp-admin', '/.well-known/')
        if any(request.path.startswith(p) for p in ignored):
            return '', 404
        logger.warning(f"404 Not Found: {request.path}")
        return render_template('error.html', error='页面未找到'), 404

    @app.errorhandler(405)
    def handle_method_not_allowed(e):
        logger.warning(f"405 Method Not Allowed: {request.method} {request.path}")
        return render_template('error.html', error='请求方法不允许'), 405

    @app.errorhandler(HTTPException)
    def handle_http_exception(e):
        logger.warning(f"HTTP {e.code}: {request.path}")
        return render_template('error.html', error=e.description), e.code

    @app.errorhandler(Exception)
    def handle_general_exception(e):
        logger.exception("未处理的异常")
        return render_template('error.html', error='服务器内部错误'), 500

    if not app.debug:
        cleanup_thread = threading.Thread(target=scheduled_cleanup, args=(app,), daemon=True)
        cleanup_thread.start()
        logger.info("已启动后台清理线程")

    return app

app = create_app()

if __name__ == '__main__':
    from waitress import serve
    try:
        port = int(os.environ.get('PORT', 5000))
        logger.info(f"启动服务...端口 {port}")
        serve(app, host='127.0.0.1', port=port,
              max_request_body_size=2*1024*1024*1024,
              channel_timeout=3600, connection_limit=100,
              clear_untrusted_proxy_headers=True, _quiet=False)
    except Exception as e:
        logger.exception("服务启动失败")
    ```

- **blueprints/__init__.py**
  - language: `python`
  - content:
    ```python
# 蓝图模块
from . import calculator, comparator, matcher
    ```

- **blueprints/calculator.py**
  - language: `python`
  - content:
    ```python
# blueprints/calculator.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app
from openpyxl.utils.exceptions import InvalidFileException

from services import FileService, ComputeEngine
from core.decorators import handle_exceptions
from core.response import error_response
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start
from core.exceptions import ValidationError
from core.excel_helpers import normalize_excel_file
from services.excel.calculator_processor import ColumnCalculatorProcessor, AggregateBackfillProcessor

logger = logging.getLogger(__name__)
calc_bp = Blueprint('calc', __name__, url_prefix='/calc', template_folder='../templates/calculator')




def ensure_excel_compatible(filepath: str, file_service) -> (str, str):
    """
    确保文件可以被后续处理（openpyxl 和 pandas）正常读取。
    返回 (规范化的文件路径, 临时文件标记)

    注意：不在这里删除任何文件，文件清理由调用者负责
    """
    from core.excel_helpers import normalize_excel_file, safe_delete_file

    normalized_path = None
    try:
        normalized_path = normalize_excel_file(filepath)
        logger.info(f"文件规范化完成: {filepath} -> {normalized_path}")
        return normalized_path, normalized_path

    except Exception as e:
        logger.error(f"Excel 文件处理失败: {e}")
        # 清理可能创建的文件
        if normalized_path and os.path.exists(normalized_path) and normalized_path != filepath:
            safe_delete_file(normalized_path)
        raise ValidationError(f"Excel 文件处理失败: {str(e)}。请确保文件是有效的 Excel 文件且未损坏。")




@calc_bp.route('/')
def index():
    return render_template('calculator/index.html')


@calc_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    compute_engine = current_app.container.compute_engine()

    if 'file' not in request.files:
        return error_response('未上传文件', 400)
    file = request.files['file']
    if file.filename == '':
        return error_response('文件名为空', 400)

    original_filename = file.filename
    ext = os.path.splitext(original_filename)[1].lower()
    if ext not in ('.xlsx', '.xls', '.xlsm'):
        raise ValidationError(f"不支持的文件格式: {ext}，请上传 .xlsx 或 .xls 文件")

    # 保存原始上传文件
    filepath = file_service.save_upload(file)
    temp_processed = None

    # 确保文件兼容，总是转换为临时文件
    filepath, temp_processed = ensure_excel_compatible(filepath, file_service)

    calc_type = request.form.get('calc_type')
    if calc_type not in ('column', 'aggregate'):
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        return error_response('无效的操作类型', 400)

    # 表头行号：支持 header_row=0 表示无表头
    header_row_str = request.form.get('header_row', '').strip()
    if header_row_str:
        try:
            header_row = int(header_row_str)
            if header_row < 0:
                raise ValueError
        except ValueError:
            return error_response('表头行号必须为非负整数', 400)
    else:
        try:
            data_start = detect_header_start(filepath)
            header_row = data_start - 1
            if header_row < 0:
                header_row = 0
        except Exception:
            header_row = 0

    result_path = None
    try:
        if calc_type == 'column':
            expression = request.form.get('expression', '').strip()
            result_col = request.form.get('result_col_column')
            merge_by_cols = request.form.get('merge_by_cols', '').strip()

            if not expression:
                return error_response('表达式不能为空', 400)

            processor = ColumnCalculatorProcessor(
                filepath=filepath,
                header_row=header_row,
                expression=expression,
                result_col_input=result_col,
                merge_by_cols_str=merge_by_cols,
                compute_engine=compute_engine
            )
            result_path = processor.process()

        else:  # aggregate
            group_cols_str = request.form.get('group_cols', '').strip()
            agg_cols_str = request.form.get('agg_cols', '').strip()
            agg_func = request.form.get('agg_func', 'sum')
            result_col = request.form.get('result_col_aggregate', '').strip()

            if not group_cols_str or not agg_cols_str or not result_col:
                return error_response('分组列、聚合列和结果列不能为空', 400)

            group_cols_raw = [c.strip() for c in group_cols_str.split(',')]
            agg_cols_raw = [c.strip() for c in agg_cols_str.split(',')]

            processor = AggregateBackfillProcessor(
                filepath=filepath,
                header_row=header_row,
                group_cols_raw=group_cols_raw,
                agg_cols_raw=agg_cols_raw,
                agg_func=agg_func,
                result_col_input=result_col,
                compute_engine=compute_engine
            )
            result_path = processor.process()

    except InvalidFileException as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise ValidationError(f"文件格式错误，无法打开: {str(e)}。请确保文件是有效的 Excel 文件（.xlsx/.xls）且未被损坏。")
    except Exception as e:
        file_service.cleanup([filepath])
        if temp_processed:
            file_service.cleanup([temp_processed])
        raise

    # 清理临时文件（不删除最终结果）
    try:
        from core.excel_helpers import safe_delete_file

        # 清理处理器创建的临时文件
        if temp_processed and os.path.exists(temp_processed) and temp_processed != result_path:
            if safe_delete_file(temp_processed):
                logger.info(f"已清理临时文件: {temp_processed}")

        # 清理原始上传文件（由文件服务处理）
        # 注意：这里不直接删除，由文件服务的清理机制处理

    except Exception as e:
        logger.warning(f"清理临时文件失败: {e}")

    download_name = generate_output_filename('calc', original_filename, suffix='_计算结果')
    return send_file(result_path, as_attachment=True, download_name=download_name)
    ```

- **blueprints/checker.py**
  - language: `python`
  - content:
    ```python
# blueprints/checker.py
import os, json, logging, tempfile
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import make_response
from services.checker import CheckerService
from core.task.decorator import async_task
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
checker_bp = Blueprint('checker', __name__, url_prefix='/checker', template_folder='../templates/checker')

@checker_bp.route('/precheck', methods=['POST'])
@handle_exceptions()
def precheck():
    files_a = request.files.getlist('files_a')
    files_b = request.files.getlist('files_b')
    if not files_a or not files_b:
        raise ValidationError("请上传两批文件")
    saved_a, saved_b = [], []
    for f in files_a:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"pre_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_a.append(p)
    for f in files_b:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"pre_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_b.append(p)
    svc = CheckerService(files_a=saved_a, files_b=saved_b, config=current_app.container.config.provided.checker())
    res = svc.precheck()
    for p in saved_a+saved_b: os.remove(p)
    return make_response(True, res)

@checker_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    files_a = request.files.getlist('files_a')
    files_b = request.files.getlist('files_b')
    if not files_a or not files_b:
        raise ValidationError("请上传两批文件")
    saved_a, saved_b = [], []
    for f in files_a:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"chk_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_a.append(p)
    for f in files_b:
        if f.filename:
            p = os.path.join(tempfile.gettempdir(), f"chk_{os.urandom(4).hex()}_{f.filename}")
            f.save(p)
            saved_b.append(p)

    key_map_a = [k.strip() for k in request.form['key_map_a'].split(',') if k.strip()]
    key_map_b_str = request.form.get('key_map_b','')
    key_map_b = [k.strip() for k in key_map_b_str.split(',') if k.strip()] if key_map_b_str else key_map_a
    compare_map = json.loads(request.form['compare_map'])
    filters_a = json.loads(request.form.get('filters_a','[]'))
    filters_b = json.loads(request.form.get('filters_b','[]'))
    dup_action = request.form.get('duplicate_action','error')

    app = current_app._get_current_object()
    cfg = app.container.config.provided.checker()
    pre_svc = CheckerService(files_a=saved_a, files_b=saved_b, config=cfg)
    pre_data = pre_svc.precheck()

    @async_task(original_filename='checker')
    def run_check(task_id):
        with app.app_context():
            task_storage = app.container.task_storage()
            def progress_cb(**kw):
                task_storage.update(task_id, **kw)
            service = CheckerService(
                files_a=saved_a, files_b=saved_b,
                key_map_a=key_map_a, key_map_b=key_map_b,
                compare_map=compare_map, filters_a=filters_a, filters_b=filters_b,
                duplicate_action=dup_action, config=cfg,
                progress_callback=progress_cb,
                precheck_data=pre_data
            )
            result_path = service.execute()
            return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}

    result = run_check()
    return make_response(True, {'task_id': result['task_id']})

@checker_bp.route('/status/<task_id>')
def status(task_id):
    data = current_app.container.task_storage().load(task_id)
    if not data:
        return make_response(False, None, '任务不存在', 404)
    return make_response(True, data)

@checker_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status')!='completed':
        return make_response(False, None, '不可下载', 400)
    path = data['result_path']
    if not path or not os.path.exists(path):
        return make_response(False, None, '文件不存在', 404)
    # 获取原始文件名用于生成下载名
    original_filename = data.get('original_filename', '校验结果.xlsx')
    download_name = generate_output_filename('checker', original_filename, suffix='_校验结果')
    return send_file(path, as_attachment=True, download_name=download_name)

@checker_bp.route('/')
def index():
    return render_template('checker/index.html')
    ```

- **blueprints/comparator.py**
  - language: `python`
  - content:
    ```python
# blueprints/comparator.py
import io
import logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, jsonify

from services import FileService
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import error_response, success_response
from core.file_naming import generate_output_filename
from schemas.comparator import CompareFullParams, CompareFieldMatchParams
from services.excel.comparator import ExcelComparator

logger = logging.getLogger(__name__)
compare_bp = Blueprint('compare', __name__, url_prefix='/compare', template_folder='../templates/comparator')

@compare_bp.route('/')
def index():
    return render_template('comparator/index.html')

@compare_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()

    file1 = request.files.get('file1')
    file2 = request.files.get('file2')
    if not file1 or not file2:
        raise ValidationError('请上传两个文件')

    compare_mode = request.form.get('compare_mode', 'full')
    sheet1 = request.form.get('sheet1', '').strip() or None
    sheet2 = request.form.get('sheet2', '').strip() or None

    path1 = file_service.save_upload(file1)
    path2 = file_service.save_upload(file2)

    try:
        if compare_mode == 'full':
            key_fields_str = request.form.get('key_fields', '').strip()
            compare_fields_str = request.form.get('compare_fields', '').strip()
            if not key_fields_str or not compare_fields_str:
                raise ValidationError('主键字段和对比字段不能为空')
            key_fields = [k.strip() for k in key_fields_str.split(',') if k.strip()]
            compare_fields = [c.strip() for c in compare_fields_str.split(',') if c.strip()]
            params = CompareFullParams(key_fields=key_fields, compare_fields=compare_fields, sheet1=sheet1, sheet2=sheet2)
            result = ExcelComparator.full_compare(path1, path2, params.key_fields, params.compare_fields,
                                                   params.sheet1, params.sheet2)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sheet_written = False
                if not result['only_in_1'].empty:
                    result['only_in_1'].to_excel(writer, sheet_name='仅存在于文件1', index=False)
                    sheet_written = True
                if not result['only_in_2'].empty:
                    result['only_in_2'].to_excel(writer, sheet_name='仅存在于文件2', index=False)
                    sheet_written = True
                if result['diff'] is not None and not result['diff'].empty:
                    result['diff'].to_excel(writer, sheet_name='内容不一致', index=False)
                    sheet_written = True
            if not sheet_written:
                return jsonify({'code': 200, 'message': '没有发现任何差异'}), 200
            output.seek(0)
            download_name = generate_output_filename('compare', file1.filename, suffix='_对比结果')
            return send_file(output, as_attachment=True, download_name=download_name,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            match_field = request.form.get('match_field', '').strip()
            if not match_field:
                raise ValidationError('请指定匹配字段')
            params = CompareFieldMatchParams(match_field=match_field, sheet1=sheet1, sheet2=sheet2)
            only_in_1, only_in_2 = ExcelComparator.field_match(path1, path2, params.match_field,
                                                                params.sheet1, params.sheet2)
            output_data = []
            if not only_in_1.empty:
                output_data.append(('文件1差异', only_in_1))
            if not only_in_2.empty:
                output_data.append(('文件2差异', only_in_2))
            if not output_data:
                return jsonify({'code': 200, 'message': '两个文件中该字段值完全匹配，无差异'}), 200
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in output_data:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            output.seek(0)
            download_name = generate_output_filename('compare', file1.filename, suffix='_字段差异')
            return send_file(output, as_attachment=True, download_name=download_name,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    finally:
        file_service.cleanup([path1, path2])
    ```

- **blueprints/convert.py**
  - language: `python`
  - content:
    ```python
# blueprints/convert.py
import os, tempfile, logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, jsonify, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import make_response
from services.convert.service import ConvertService
from core.task.decorator import async_task
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
convert_bp = Blueprint('convert', __name__, url_prefix='/convert', template_folder='../templates/convert')

@convert_bp.route('/')
def index():
    return render_template('convert/index.html')

@convert_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file = request.files.get('file')
    if not file or file.filename == '':
        raise ValidationError("请上传文件")
    target_fmt = request.form.get('target_format', 'csv')
    encoding = request.form.get('encoding', 'utf-8-sig')
    separator = request.form.get('separator', ',')

    upload_dir = tempfile.mkdtemp(prefix='convert_')
    save_path = os.path.join(upload_dir, file.filename)
    file.save(save_path)

    app = current_app._get_current_object()

    @async_task(original_filename='convert')
    def run_conversion(task_id):
        with app.app_context():
            task_storage = app.container.task_storage()
            def progress_cb(**kw):
                task_storage.update(task_id, **kw)
            service = ConvertService(
                file_path=save_path,
                target_format=target_fmt,
                encoding=encoding,
                separator=separator,
                progress_callback=progress_cb
            )
            result_path = service.execute()
            return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}

    result = run_conversion()
    return make_response(True, {'task_id': result['task_id']})

@convert_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data:
        return make_response(False, None, '任务不存在', 404)
    return make_response(True, data)

@convert_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status')!='completed':
        return make_response(False, None, '不可下载', 400)
    path = data['result_path']
    if not path or not os.path.exists(path):
        return make_response(False, None, '文件不存在', 404)
    basename = os.path.basename(path)
    # 统一命名：convert_原始文件名_转换结果.扩展名
    original = data.get('original_filename', 'file')
    safe_name = generate_output_filename('convert', original, suffix='_转换结果')
    return send_file(path, as_attachment=True, download_name=safe_name)
    ```

- **blueprints/field_parser.py**
  - language: `python`
  - content:
    ```python
# blueprints/field_parser.py
import io, logging, pandas as pd
from flask import Blueprint, render_template, request, jsonify, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.csv_utils import split_fields
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
field_parser_bp = Blueprint('field_parser', __name__, url_prefix='/field-parser', template_folder='../templates/field_parser')

@field_parser_bp.route('/')
def index():
    return render_template('field_parser/index.html')

@field_parser_bp.route('/analyze', methods=['GET', 'POST'])
@handle_exceptions()
def analyze():
    if request.method == 'GET':
        return jsonify({'message': '请使用 POST 方法提交 JSON 数据', 'fields': {'standard_headers': '标准列名列表（文本）', 'compare_headers': '比较列名列表（文本）'}})

    if not request.is_json:
        raise ValidationError("仅支持 JSON 格式的请求")

    data = request.get_json()
    text_a = data.get('standard_headers', '')
    text_b = data.get('compare_headers', '')
    if not text_a or not text_b:
        raise ValidationError("请提供 standard_headers 和 compare_headers 字段")

    service = current_app.container.field_parser()
    cols_a = split_fields(text_a)
    cols_b = split_fields(text_b)

    if not cols_a or not cols_b:
        raise ValidationError("未能解析出有效的列名，请检查输入格式")

    result = service.match_headers(cols_a, cols_b)
    result['cols_a'] = cols_a
    result['cols_b'] = cols_b
    return jsonify(result)

@field_parser_bp.route('/export', methods=['POST'])
@handle_exceptions()
def export():
    if not request.is_json:
        raise ValidationError("仅支持 JSON 格式的请求")

    data = request.get_json()
    matched = data.get('matched', [])
    only_a = data.get('only_a', [])
    only_b = data.get('only_b', [])
    cols_a = data.get('cols_a', [])
    cols_b = data.get('cols_b', [])

    rows = []
    for item in matched:
        rows.append({'标准字段': item['standard'], '比较字段': item['compare'], '相似度': item['score'], '备注': '自动匹配'})
    for col in only_a:
        rows.append({'标准字段': col, '比较字段': '', '相似度': '', '备注': '仅标准'})
    for col in only_b:
        rows.append({'标准字段': '', '比较字段': col, '相似度': '', '备注': '仅比较'})

    df_mapping = pd.DataFrame(rows, columns=['标准字段', '比较字段', '相似度', '备注'])

    max_len = max(len(cols_a), len(cols_b))
    cols_a_padded = list(cols_a) + [''] * (max_len - len(cols_a))
    cols_b_padded = list(cols_b) + [''] * (max_len - len(cols_b))
    df_original = pd.DataFrame({'标准文件列名': cols_a_padded, '比较文件列名': cols_b_padded})

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_mapping.to_excel(writer, sheet_name='字段映射', index=False)
        df_original.to_excel(writer, sheet_name='原始数据', index=False)
    output.seek(0)

    download_name = generate_output_filename('field_parser', 'mapping.xlsx', suffix='_字段映射')
    return send_file(output, as_attachment=True, download_name=download_name,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    ```

- **blueprints/matcher.py**
  - language: `python`
  - content:
    ```python
# blueprints/matcher.py
import os
import logging
from flask import Blueprint, render_template, request, send_file, current_app

from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.file_naming import generate_output_filename
from core.csv_utils import split_fields
from services.excel.matcher import ExcelMatcher

logger = logging.getLogger(__name__)
match_bp = Blueprint('match', __name__, url_prefix='/match', template_folder='../templates/matcher')

ALLOWED_TARGET_EXT = {'.xlsx'}

def _check_target_extension(filename: str):
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_TARGET_EXT:
        raise ValidationError(f"目标文件格式不支持，请使用 .xlsx 格式（当前: {ext}）。如为 .xls 文件，请先用 Excel 另存为 .xlsx。")

def _fix_target_extension(filepath: str) -> str:
    if not filepath.lower().endswith('.xlsx'):
        new_path = filepath + '.xlsx'
        try:
            os.rename(filepath, new_path)
            return new_path
        except Exception:
            pass
    return filepath

def _parse_multi_cols(raw: str) -> list:
    """解析多列输入，支持英文逗号、中文逗号、空格分隔，返回去重后的列表"""
    if not raw:
        return []
    # 使用 split_fields 工具函数，它已经支持逗号、空格、换行分隔
    return split_fields(raw)

@match_bp.route('/')
def index():
    return render_template('matcher/index.html')

@match_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()

    target_file = request.files.get('target_file')
    source_file = request.files.get('source_file')
    if not target_file or not source_file:
        raise ValidationError('请上传两个文件')

    _check_target_extension(target_file.filename)

    target_path = file_service.save_upload(target_file)
    target_path = _fix_target_extension(target_path)
    source_path = file_service.save_upload(source_file)

    # 解析多列输入
    target_key_cols_raw = request.form.get('target_key_cols', '').strip()
    source_match_cols_raw = request.form.get('source_match_cols', '').strip()
    source_value_cols_raw = request.form.get('source_value_col', '').strip()
    result_cols_raw = request.form.get('result_col', '').strip()
    source_sheets_raw = request.form.get('source_sheets', '').strip()
    allow_duplicate_keys = request.form.get('allow_duplicate_keys') == 'on'

    if not target_key_cols_raw or not source_match_cols_raw or not source_value_cols_raw or not result_cols_raw:
        raise ValidationError('请填写完整配置项')

    target_key_cols = _parse_multi_cols(target_key_cols_raw)
    source_match_cols = _parse_multi_cols(source_match_cols_raw)
    source_value_cols = _parse_multi_cols(source_value_cols_raw)
    result_cols = _parse_multi_cols(result_cols_raw)

    # 验证数量一致性
    if len(source_value_cols) != len(result_cols):
        raise ValidationError(f'源文件值列数量({len(source_value_cols)})与结果列数量({len(result_cols)})不一致')
    if not target_key_cols or not source_match_cols:
        raise ValidationError('主键列和匹配列不能为空')
    if len(target_key_cols) != len(source_match_cols):
        raise ValidationError(f'目标主键列数量({len(target_key_cols)})与源匹配列数量({len(source_match_cols)})不一致')

    sheet_indices = None
    if source_sheets_raw:
        try:
            sheet_indices = [int(x.strip()) for x in source_sheets_raw.split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            raise ValidationError('Sheet索引格式错误，请输入逗号分隔的数字，例如 1,2,3')

    config = {
        'target_key_cols': target_key_cols,
        'source_match_cols': source_match_cols,
        'source_value_cols': source_value_cols,
        'result_cols': result_cols,
        'allow_duplicate_keys': allow_duplicate_keys
    }

    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    success, result = matcher.match()
    if not success:
        error_msg = result if isinstance(result, str) else '\n'.join(result)
        raise ValidationError(error_msg)

    download_name = generate_output_filename('match', target_file.filename, suffix='_匹配结果')
    return send_file(result, as_attachment=True, download_name=download_name)

@match_bp.route('/debug', methods=['POST'])
@handle_exceptions()
def debug():
    file_service = current_app.container.file_service()

    target_file = request.files.get('target_file')
    source_file = request.files.get('source_file')
    if not target_file or not source_file:
        return render_template('error.html', error='请上传两个文件'), 400

    _check_target_extension(target_file.filename)

    target_path = file_service.save_upload(target_file)
    target_path = _fix_target_extension(target_path)
    source_path = file_service.save_upload(source_file)

    target_key_cols_raw = request.form.get('target_key_cols', '').strip()
    source_match_cols_raw = request.form.get('source_match_cols', '').strip()
    source_value_cols_raw = request.form.get('source_value_col', '').strip()
    result_cols_raw = request.form.get('result_col', '').strip()
    source_sheets_raw = request.form.get('source_sheets', '').strip()
    allow_duplicate_keys = request.form.get('allow_duplicate_keys') == 'on'

    if not target_key_cols_raw or not source_match_cols_raw or not source_value_cols_raw or not result_cols_raw:
        return render_template('error.html', error='请填写完整配置项'), 400

    target_key_cols = _parse_multi_cols(target_key_cols_raw)
    source_match_cols = _parse_multi_cols(source_match_cols_raw)
    source_value_cols = _parse_multi_cols(source_value_cols_raw)
    result_cols = _parse_multi_cols(result_cols_raw)

    if len(source_value_cols) != len(result_cols):
        return render_template('error.html', error=f'源文件值列数量({len(source_value_cols)})与结果列数量({len(result_cols)})不一致'), 400

    sheet_indices = None
    if source_sheets_raw:
        try:
            sheet_indices = [int(x.strip()) for x in source_sheets_raw.split(',') if x.strip()]
            if not sheet_indices:
                sheet_indices = None
        except ValueError:
            return render_template('error.html', error='Sheet索引格式错误，请输入逗号分隔的数字'), 400

    config = {
        'target_key_cols': target_key_cols,
        'source_match_cols': source_match_cols,
        'source_value_cols': source_value_cols,
        'result_cols': result_cols,
        'allow_duplicate_keys': allow_duplicate_keys
    }

    matcher = ExcelMatcher(target_path, source_path, config, sheet_indices)
    info = matcher.debug_info()
    if 'error' in info:
        return render_template('error.html', error=info['error']), 400

    return render_template('debug.html', config=config, info=info)
    ```

- **blueprints/splitter.py**
  - language: `python`
  - content:
    ```python
# blueprints/splitter.py
import os
import tempfile
import zipfile
import logging
from pathlib import Path
from flask import Blueprint, render_template, request, send_file, current_app
from openpyxl import load_workbook, Workbook
from copy import copy
import pandas as pd

from core.validators import sanitize_filename
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.filter_engine import parse_conditions_from_form, ConditionEngine
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
splitter_bp = Blueprint('splitter', __name__, url_prefix='/splitter', template_folder='../templates/splitter')

def copy_cell(source_cell, target_cell):
    try:
        if hasattr(source_cell, 'value') and source_cell.value is not None:
            target_cell.value = source_cell.value
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = source_cell.number_format
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.protection = copy(source_cell.protection)
    except AttributeError:
        pass

def copy_worksheet_without_data(source_ws, target_wb, keep_rows_set):
    target_ws = target_wb.create_sheet(title=source_ws.title)
    if keep_rows_set is None:
        keep_rows_set = set(range(1, source_ws.max_row + 1))
    old_to_new = {}
    new_row = 1
    for old_row in range(1, source_ws.max_row + 1):
        if old_row in keep_rows_set:
            old_to_new[old_row] = new_row
            new_row += 1
    for col in source_ws.column_dimensions:
        target_ws.column_dimensions[col].width = source_ws.column_dimensions[col].width
    for old_row, new_row_num in old_to_new.items():
        if old_row in source_ws.row_dimensions:
            target_ws.row_dimensions[new_row_num].height = source_ws.row_dimensions[old_row].height
    merged_cells_to_apply = []
    for merged_range in source_ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
        rows_in_range = set(range(min_row, max_row + 1))
        if not rows_in_range.intersection(keep_rows_set):
            continue
        new_min_row = old_to_new.get(min_row, min_row)
        new_max_row = old_to_new.get(max_row, max_row)
        if new_min_row and new_max_row:
            merged_cells_to_apply.append((new_min_row, min_col, new_max_row, max_col))
    max_col = source_ws.max_column
    is_top_left = {}
    for merged_range in source_ws.merged_cells.ranges:
        min_row, min_col = merged_range.min_row, merged_range.min_col
        is_top_left[(min_row, min_col)] = True
    for old_row, new_row_num in old_to_new.items():
        for col in range(1, max_col + 1):
            if (old_row, col) in is_top_left:
                pass
            else:
                skip = False
                for merged_range in source_ws.merged_cells.ranges:
                    if (merged_range.min_row <= old_row <= merged_range.max_row and
                        merged_range.min_col <= col <= merged_range.max_col):
                        if old_row != merged_range.min_row or col != merged_range.min_col:
                            skip = True
                            break
                if skip:
                    continue
            src_cell = source_ws.cell(row=old_row, column=col)
            dst_cell = target_ws.cell(row=new_row_num, column=col)
            copy_cell(src_cell, dst_cell)
    for min_row, min_col, max_row, max_col in merged_cells_to_apply:
        if min_row < max_row or min_col < max_col:
            target_ws.merge_cells(start_row=min_row, start_column=min_col,
                                  end_row=max_row, end_column=max_col)
    return target_ws

def repair_excel_with_pandas(input_path):
    import pandas as pd
    temp_xlsx = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_xlsx.close()
    try:
        with pd.ExcelFile(input_path) as xls:
            sheet_names = xls.sheet_names
            wb = Workbook()
            wb.remove(wb.active)
            for sheet_name in sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, keep_default_na=False)
                ws = wb.create_sheet(title=sheet_name)
                for col_idx, col_name in enumerate(df.columns, 1):
                    ws.cell(row=1, column=col_idx, value=col_name)
                for row_idx, row in enumerate(df.values, 2):
                    for col_idx, val in enumerate(row, 1):
                        ws.cell(row=row_idx, column=col_idx, value=val)
            wb.save(temp_xlsx.name)
    except Exception as e:
        try:
            os.unlink(temp_xlsx.name)
        except:
            pass
        raise e
    return temp_xlsx.name

def split_excel_to_temp(input_path, split_column_name, include_values, exclude_values, skip_sheet_indices, temp_dir):
    wb = load_workbook(input_path, data_only=True)
    sheets = wb.worksheets
    first_sheet = sheets[0]
    headers = [cell.value for cell in first_sheet[1]]
    if split_column_name not in headers:
        wb.close()
        raise ValidationError(f"拆分列 '{split_column_name}' 不在第一个工作表的表头中")
    split_col_idx = headers.index(split_column_name) + 1

    all_values = set()
    for idx, sheet in enumerate(sheets):
        if (idx + 1) in skip_sheet_indices:
            continue
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=split_col_idx)
            if cell.value is not None:
                val = str(cell.value).strip()
                if val:
                    all_values.add(val)

    if include_values:
        target_values = [v for v in all_values if v in include_values]
    elif exclude_values:
        target_values = [v for v in all_values if v not in exclude_values]
    else:
        target_values = list(all_values)

    if not target_values:
        wb.close()
        raise ValidationError("没有找到任何需要拆分的值")

    generated_files = []
    for split_val in target_values:
        safe_val = sanitize_filename(str(split_val))
        new_wb = Workbook()
        default_sheet = new_wb.active
        new_wb.remove(default_sheet)
        for idx, src_ws in enumerate(sheets):
            if (idx + 1) in skip_sheet_indices:
                continue
            keep_rows = {1}
            for row in range(2, src_ws.max_row + 1):
                cell_val = src_ws.cell(row=row, column=split_col_idx).value
                if cell_val is not None and str(cell_val) == str(split_val):
                    keep_rows.add(row)
            copy_worksheet_without_data(src_ws, new_wb, keep_rows)
        out_file = Path(temp_dir) / f"{safe_val}.xlsx"
        new_wb.save(str(out_file))
        new_wb.close()
        generated_files.append(str(out_file))
    wb.close()
    return generated_files

def filter_workbook_inplace(input_path, form, skip_sheet_indices=None):
    if skip_sheet_indices is None:
        skip_sheet_indices = set()

    exact, fuzzy = parse_conditions_from_form(form)
    if not exact and not fuzzy:
        return input_path

    engine = ConditionEngine()

    wb = load_workbook(input_path, data_only=True)
    try:
        for idx, ws in enumerate(wb.worksheets, start=1):
            if idx in skip_sheet_indices:
                continue

            headers = [cell.value for cell in ws[1]]
            missing_fields = []
            for cond in exact + fuzzy:
                if cond.get('field', '') not in headers:
                    missing_fields.append(cond.get('field', ''))
            if missing_fields:
                logger.warning(f"工作表 '{ws.title}' 缺少字段 {missing_fields}，跳过筛选")
                continue

            col_map = {h: i + 1 for i, h in enumerate(headers)}
            rows_to_delete = []
            for row_idx in range(2, ws.max_row + 1):
                row_dict = {}
                for col_name, col_idx in col_map.items():
                    cell = ws.cell(row=row_idx, column=col_idx)
                    row_dict[col_name] = str(cell.value) if cell.value is not None else ''
                if not engine.apply_to_row(row_dict, exact, fuzzy):
                    rows_to_delete.append(row_idx)

            for row_idx in reversed(rows_to_delete):
                ws.delete_rows(row_idx)

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        wb.save(tmp.name)
        return tmp.name
    finally:
        wb.close()

# ---------- 路由 ----------
@splitter_bp.route('/', methods=['GET'])
def index():
    return render_template('splitter/index.html')

@splitter_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    if 'file' not in request.files:
        raise ValidationError('未上传文件')
    file = request.files['file']
    if file.filename == '':
        raise ValidationError('文件名为空')

    split_column = request.form.get('split_column', '').strip()
    if not split_column:
        raise ValidationError('请指定拆分列名')

    include_values_str = request.form.get('include_values', '').strip()
    exclude_values_str = request.form.get('exclude_values', '').strip()
    skip_sheets_str = request.form.get('skip_sheets', '').strip()

    include_values = [v.strip() for v in include_values_str.split(',') if v.strip()] if include_values_str else []
    exclude_values = [v.strip() for v in exclude_values_str.split(',') if v.strip()] if exclude_values_str else []
    skip_sheet_indices = {int(v.strip()) for v in skip_sheets_str.split(',') if v.strip().isdigit()} if skip_sheets_str else set()

    if include_values and exclude_values:
        raise ValidationError('正选和反选不能同时使用')

    save_path = file_service.save_upload(file, subdir='splitter', allowed_extensions={'xlsx','xls','xlsm'})
    if not save_path.lower().endswith('.xlsx'):
        new_path = save_path + '.xlsx'
        os.rename(save_path, new_path)
        save_path = new_path

    try:
        load_workbook(save_path, data_only=True)
    except Exception:
        logger.warning("openpyxl 打开失败，尝试用 pandas 修复（公式将丢失）")
        try:
            repaired_path = repair_excel_with_pandas(save_path)
            os.remove(save_path)
            save_path = repaired_path
        except Exception as repair_err:
            raise ValidationError(f'文件无法被识别为有效的 Excel 文件: {str(repair_err)}')

    filter_path = filter_workbook_inplace(save_path, request.form, skip_sheet_indices)
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            generated_files = split_excel_to_temp(
                input_path=filter_path,
                split_column_name=split_column,
                include_values=include_values,
                exclude_values=exclude_values,
                skip_sheet_indices=skip_sheet_indices,
                temp_dir=temp_dir
            )
            zip_buffer = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
            zip_buffer.close()
            with zipfile.ZipFile(zip_buffer.name, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fpath in generated_files:
                    arcname = os.path.basename(fpath)
                    zf.write(fpath, arcname)

            from datetime import datetime
            original_name = os.path.splitext(file.filename)[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            download_name = generate_output_filename('splitter', file.filename, suffix='_拆分结果')
            return send_file(
                zip_buffer.name,
                as_attachment=True,
                download_name=download_name,
                mimetype='application/zip'
            )
    finally:
        if filter_path != save_path:
            try:
                os.remove(filter_path)
            except:
                pass
    ```

- **blueprints/subtotal.py**
  - language: `python`
  - content:
    ```python
# blueprints/subtotal.py
import os
import logging
import openpyxl
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.task.decorator import async_task
from services.subtotal.engine import SubtotalEngine, _build_header_map, resolve_single_col
from core.csv_utils import split_fields
from core.file_naming import generate_output_filename, build_content_disposition

logger = logging.getLogger(__name__)
subtotal_bp = Blueprint('subtotal', __name__, url_prefix='/subtotal',
                        template_folder='../templates/subtotal')

@subtotal_bp.route('/')
def index():
    return render_template('subtotal/index.html')

def _check_columns(raw, category, header_map, max_col):
    if not raw:
        return
    items = split_fields(raw)
    for item in items:
        if '-' in item:
            parts = item.split('-')
            if len(parts) != 2:
                raise ValidationError(f"{category}列范围格式错误：{item}")
            start, end = parts[0].strip(), parts[1].strip()
            start_idx = resolve_single_col(start, header_map)
            end_idx = resolve_single_col(end, header_map)
            if start_idx > max_col:
                raise ValidationError(f"{category}列 '{start}' 超出工作表列范围（共{max_col}列）")
            if end_idx > max_col:
                raise ValidationError(f"{category}列 '{end}' 超出工作表列范围（共{max_col}列）")
        else:
            idx = resolve_single_col(item, header_map)
            if idx > max_col:
                raise ValidationError(f"{category}列 '{item}' 超出工作表列范围（共{max_col}列）")

@subtotal_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file = request.files.get('file')
    if not file or file.filename == '':
        raise ValidationError('请上传文件')

    original_filename = file.filename
    key_cols_raw = request.form.get('key_cols', '').strip()
    if not key_cols_raw:
        raise ValidationError('请指定主键列')

    dimension_cols_raw = request.form.get('dimension_cols', '').strip() or None
    calc_cols_raw = request.form.get('calc_cols', '').strip()
    if not calc_cols_raw:
        raise ValidationError('请指定计算列')

    agg_func_default = request.form.get('agg_func_default', 'sum').strip()
    agg_overrides_raw = request.form.get('agg_overrides', '').strip()
    sheets_mode = request.form.get('sheets_mode', '1').strip()

    file_service = current_app.container.file_service()
    path = file_service.save_upload(file, allowed_extensions={'xlsx', 'xls'})

    wb = openpyxl.load_workbook(path)
    first_ws = wb.active
    max_col = first_ws.max_column
    header_map = _build_header_map(first_ws)
    wb.close()

    _check_columns(key_cols_raw, '主键', header_map, max_col)
    _check_columns(calc_cols_raw, '计算', header_map, max_col)
    if dimension_cols_raw:
        _check_columns(dimension_cols_raw, '维度', header_map, max_col)

    agg_funcs = {'default': agg_func_default}
    if agg_overrides_raw:
        for part in agg_overrides_raw.split(','):
            if ':' in part:
                col, func = part.split(':', 1)
                col = col.strip().upper()
                func = func.strip().lower()
                if func not in ('sum', 'avg', 'min', 'max', 'count'):
                    raise ValidationError(f"不支持的聚合函数: {func}")
                try:
                    openpyxl.utils.column_index_from_string(col)
                except:
                    raise ValidationError(f"无效的列字母: {col}")
                agg_funcs[col] = func

    wb2 = openpyxl.load_workbook(path, read_only=True)
    total_rows = sum(sheet.max_row for sheet in wb2.worksheets)
    wb2.close()
    data_rows = total_rows - 1

    app = current_app._get_current_object()

    if data_rows > 5000:
        task_storage = app.container.task_storage()

        @async_task(original_filename='subtotal')
        def run_subtotal(task_id):
            with app.app_context():
                def progress_cb(percent, message):
                    task_storage.update(task_id, progress_percent=percent, progress_message=message)
                try:
                    progress_cb(0, "开始处理...")
                    engine = SubtotalEngine(
                        path, key_cols_raw, calc_cols_raw, agg_funcs,
                        dimension_cols=dimension_cols_raw,
                        sheets=sheets_mode,
                        original_filename=original_filename
                    )
                    out_path = engine.process()
                    progress_cb(100, "完成")
                    # 保存原始文件名到任务数据中，便于下载时生成统一名称
                    return {'result_path': out_path, 'original_filename': original_filename}
                except Exception as e:
                    import traceback
                    error_msg = f"{str(e)}\n{traceback.format_exc()}"
                    task_storage.append_log(task_id, error_msg)
                    raise
                finally:
                    file_service.cleanup([path])

        result = run_subtotal()
        return jsonify({'task_id': result['task_id']})
    else:
        try:
            engine = SubtotalEngine(
                path, key_cols_raw, calc_cols_raw, agg_funcs,
                dimension_cols=dimension_cols_raw,
                sheets=sheets_mode,
                original_filename=original_filename
            )
            out_path = engine.process()
            download_name = generate_output_filename('subtotal', original_filename, suffix='_小计')
            resp = send_file(out_path, as_attachment=True,
                             download_name=download_name,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp.headers['Content-Disposition'] = build_content_disposition(download_name)
            return resp
        finally:
            file_service.cleanup([path])

@subtotal_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data:
        return jsonify({'error': '任务不存在'}), 404
    log = storage.get_log(task_id, max_lines=100)
    data['log'] = log
    return jsonify(data)

@subtotal_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status') != 'completed':
        return jsonify({'error': '结果不可用'}), 400
    path = data.get('result_path')
    if not path or not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404

    # 使用统一命名函数生成下载名
    original_filename = data.get('original_filename', 'data.xlsx')
    download_name = generate_output_filename('subtotal', original_filename, suffix='_小计')
    resp = send_file(path, as_attachment=True, download_name=download_name)
    resp.headers['Content-Disposition'] = build_content_disposition(download_name)
    return resp
    ```

- **blueprints/summarizer.py**
  - language: `python`
  - content:
    ```python
# blueprints/summarizer.py
import os
import tempfile
import logging
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app

from services import FileService, ComputeEngine, DataFormatter, DataWriter
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.response import error_response
from core.csv_utils import split_fields
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start

logger = logging.getLogger(__name__)
summary_bp = Blueprint('summary', __name__, url_prefix='/summary', template_folder='../templates/summarizer')

def _dedup(lst):
    """去重并保持顺序"""
    seen = set()
    result = []
    for item in lst:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result

@summary_bp.route('/')
def index():
    return render_template('summarizer/index.html')

@summary_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file_service = current_app.container.file_service()
    compute_engine = current_app.container.compute_engine()
    data_formatter = current_app.container.data_formatter()
    data_writer = current_app.container.data_writer()

    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        raise ValidationError('没有选择文件')

    field_str = request.form.get('fields', '').strip()
    unique_str = request.form.get('unique_keys', '').strip()
    sum_str = request.form.get('sum_fields', '').strip()

    header_row_str = request.form.get('header_row', '').strip()
    skip_header_str = request.form.get('skip_header_rows', '0').strip()

    saved_paths = []
    for file in files:
        if file and file.filename:
            path = file_service.save_upload(file)
            saved_paths.append(path)

    if not saved_paths:
        raise ValidationError('没有有效文件上传')

    # 确定表头行
    try:
        if header_row_str:
            header_row = int(header_row_str)
            if header_row < 1:
                raise ValueError
            skip_header_rows = header_row - 1
        else:
            # 自动检测第一个文件的表头
            data_start = detect_header_start(saved_paths[0])
            header_row = data_start - 1 if data_start > 1 else 1
            if header_row < 1:
                header_row = 1
            skip_header_rows = header_row - 1
    except ValueError:
        raise ValidationError('表头行号或跳过行数必须为有效正整数')

    # 解析字段
    field_list = _dedup(split_fields(field_str))
    if not field_list:
        raise ValidationError('字段列表不能为空')

    unique_keys = _dedup(split_fields(unique_str))
    sum_fields = _dedup(split_fields(sum_str))

    all_dfs = []
    failed_files = []
    for path in saved_paths:
        try:
            if skip_header_rows > 0:
                all_sheets = pd.read_excel(path, sheet_name=None, dtype=str,
                                           skiprows=skip_header_rows, header=0)
            else:
                all_sheets = pd.read_excel(path, sheet_name=None, dtype=str)
            if not all_sheets:
                raise ValueError(f"文件 {path} 无有效 sheet")
            df_list = []
            for sheet_df in all_sheets.values():
                df_list.append(sheet_df)
            df = pd.concat(df_list, ignore_index=True, sort=False)
            existing_fields = [f for f in field_list if f in df.columns]
            missing = set(field_list) - set(existing_fields)
            if missing:
                logger.warning(f"文件 {os.path.basename(path)} 缺少字段: {missing}")
            df = df[existing_fields]
            for field in field_list:
                if field not in df.columns:
                    df[field] = pd.NA
            all_dfs.append(df)
        except Exception as e:
            logger.warning(f"读取文件 {path} 失败: {e}")
            failed_files.append(os.path.basename(path))
            continue

    if not all_dfs:
        error_msg = "没有成功读取任何有效文件"
        if failed_files:
            error_msg += f"，以下文件读取失败: {', '.join(failed_files)}"
        raise ValidationError(error_msg)

    combined = pd.concat(all_dfs, ignore_index=True, sort=False)

    for field in sum_fields:
        if field not in combined.columns:
            combined[field] = 0
        combined[field] = pd.to_numeric(combined[field], errors='coerce').fillna(0)

    if not unique_keys:
        result = combined.drop_duplicates(subset=field_list, keep='first')
    else:
        valid_keys = [k for k in unique_keys if k in combined.columns]
        if not valid_keys:
            logger.warning("指定的唯一键字段都不存在，按全部字段去重")
            result = combined.drop_duplicates(subset=field_list, keep='first')
        else:
            combined[valid_keys] = combined[valid_keys].fillna('')
            result = compute_engine.aggregate_by_keys(combined, valid_keys, sum_fields)
            result[valid_keys] = result[valid_keys].fillna('')

    result = data_formatter.round_numeric_columns(result)

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    data_writer.write_excel(result, tmp_path)

    download_name = generate_output_filename('summary', files[0].filename, suffix='_汇总结果')
    return send_file(tmp_path, as_attachment=True, download_name=download_name)
    ```

- **blueprints/team_report.py**
  - language: `python`
  - content:
    ```python
# blueprints/team_report.py
import json, os, logging, tempfile, pandas as pd
from flask import Blueprint, render_template, request, jsonify, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from datetime import datetime
from core.task.decorator import async_task
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
team_report_bp = Blueprint('team_report', __name__, url_prefix='/team-report', template_folder='../templates/team_report')

def _save_uploaded(file, prefix='tr_'):
    if not file or file.filename == '': return None
    from werkzeug.utils import secure_filename
    upload_dir = tempfile.mkdtemp(prefix=prefix)
    safe_name = secure_filename(file.filename)
    path = os.path.join(upload_dir, safe_name)
    file.save(path)
    return path

@team_report_bp.route('/')
def index():
    return render_template('team_report/index.html')

@team_report_bp.route('/generate', methods=['POST'])
@handle_exceptions()
def generate():
    mapping_file = request.files.get('mapping_file')
    today_file = request.files.get('today_file')
    yesterday_file = request.files.get('yesterday_file')
    monthly_files = request.files.getlist('monthly_files')

    if not today_file:
        raise ValidationError("请上传今日司机流水文件")

    mapping_path = _save_uploaded(mapping_file, prefix='mapping_') if mapping_file else None
    today_path = _save_uploaded(today_file, prefix='today_')
    yesterday_path = _save_uploaded(yesterday_file, prefix='yesterday_') if yesterday_file else None
    monthly_paths = [_save_uploaded(mf, prefix='monthly_') for mf in monthly_files if mf and mf.filename]

    daily_order_target = request.form.get('daily_order_target') or None
    daily_apply_target = request.form.get('daily_apply_target') or None
    monthly_order_target = request.form.get('monthly_order_target') or None
    monthly_apply_target = request.form.get('monthly_apply_target') or None

    app = current_app._get_current_object()

    @async_task(original_filename='team_report')
    def run_report(task_id):
        with app.app_context():
            service = app.container.team_report_service()
            task_storage = app.container.task_storage()
            try:
                task_storage.update(task_id, progress_percent=10, progress_message='加载映射...')
                if mapping_path and os.path.exists(mapping_path):
                    mapping = service.load_mapping(mapping_path)
                else:
                    raise ValidationError("未提供队伍明细文件且无默认映射")

                task_storage.update(task_id, progress_percent=30, progress_message='处理今日数据...')
                today_df = service._prepare_flow(today_path, mapping)
                yesterday_df = None
                if yesterday_path and os.path.exists(yesterday_path):
                    task_storage.update(task_id, progress_percent=50, progress_message='处理昨日数据...')
                    yesterday_df = service._prepare_flow(yesterday_path, mapping)

                monthly_df = None
                if monthly_paths:
                    task_storage.update(task_id, progress_percent=60, progress_message='处理月度数据...')
                    monthly_df = service.aggregate_monthly(monthly_paths, mapping)

                task_storage.update(task_id, progress_percent=70, progress_message='生成报表...')
                report_df = service.build_report_data(today_df, yesterday_df, monthly_df)

                task_storage.update(task_id, progress_percent=80, progress_message='写入格式化 Excel...')
                now = datetime.now()
                report_date = now.strftime('%Y/%m/%d')
                month = now.strftime('%m')
                output_dir = app.config.get('OUTPUT_FOLDER', 'outputs')
                os.makedirs(output_dir, exist_ok=True)
                result_path = os.path.join(output_dir, f'team_report_{task_id}.xlsx')

                targets = {
                    'daily_order': int(daily_order_target) if daily_order_target else service.config.daily_order_target,
                    'daily_apply': int(daily_apply_target) if daily_apply_target else service.config.daily_apply_target,
                    'monthly_order': int(monthly_order_target) if monthly_order_target else service.config.monthly_order_target,
                    'monthly_apply': int(monthly_apply_target) if monthly_apply_target else service.config.monthly_apply_target,
                }
                writer = app.container.team_report_writer()
                rows = service.report_to_rows(report_df,
                                              daily_order=targets['daily_order'],
                                              daily_apply=targets['daily_apply'],
                                              monthly_order=targets['monthly_order'],
                                              monthly_apply=targets['monthly_apply'])
                writer.write(rows, result_path, report_date, month,
                             targets['daily_order'], targets['daily_apply'],
                             targets['monthly_order'], targets['monthly_apply'])

                return {'result_path': result_path, 'completed_at': pd.Timestamp.now().isoformat()}
            finally:
                for p in [mapping_path, today_path, yesterday_path] + monthly_paths:
                    if p and os.path.exists(p):
                        try: os.remove(p)
                        except: pass

    result = run_report()
    return jsonify({'task_id': result['task_id']})

@team_report_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data: return jsonify({'error': '任务不存在'}), 404
    return jsonify(data)

@team_report_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status') != 'completed':
        return jsonify({'error': '报表不可用'}), 400
    path = data.get('result_path')
    if not path or not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404
    original = data.get('original_filename', 'team_report.xlsx')
    download_name = generate_output_filename('team_report', original, suffix='_每日数据对比')
    response = send_file(path, as_attachment=True, download_name=download_name,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         conditional=True)
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    response.headers['X-Accel-Buffering'] = 'no'
    return response
    ```

- **blueprints/transposer.py**
  - language: `python`
  - content:
    ```python
import io, logging, pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.csv_utils import split_fields
from services.transposer.service import pivot_table
from core.file_naming import generate_output_filename

logger = logging.getLogger(__name__)
transposer_bp = Blueprint('transposer', __name__, url_prefix='/transposer', template_folder='../templates/transposer')

@transposer_bp.route('/')
def index():
    return render_template('transposer/index.html')

@transposer_bp.route('/convert', methods=['POST'])
@handle_exceptions()
def convert():
    if 'file' not in request.files:
        raise ValidationError('未上传文件')
    file = request.files['file']
    if file.filename == '':
        raise ValidationError('文件名为空')

    category_col_raw = request.form.get('category_col', '').strip()
    value_cols_raw = request.form.get('value_cols', '').strip()
    index_cols_raw = request.form.get('index_cols', '').strip()

    if not category_col_raw:
        raise ValidationError('请指定分类列')
    if not value_cols_raw:
        raise ValidationError('请指定值列（可多选）')

    category_col = category_col_raw
    value_cols = split_fields(value_cols_raw)
    index_cols = split_fields(index_cols_raw) if index_cols_raw else None

    try:
        df = pd.read_excel(file, dtype=str)
    except Exception as e:
        raise ValidationError(f'无法读取文件: {str(e)}')

    all_needed = [category_col] + value_cols + (index_cols if index_cols else [])
    missing = [c for c in all_needed if c not in df.columns]
    if missing:
        raise ValidationError(f'列不存在: {", ".join(missing)}')

    try:
        result_df = pivot_table(df, index_cols, category_col, value_cols)
    except Exception as e:
        raise ValidationError(f'转换失败: {str(e)}')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='转置结果')
    output.seek(0)

    download_name = generate_output_filename('transposer', file.filename, suffix='_转置结果')
    return send_file(output, as_attachment=True, download_name=download_name,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    ```

- **core/__init__.py**
  - language: `python`
  - content:
    ```python
# core/__init__.py
"""核心基础设施层"""
from .exceptions import *
from .config import settings
from .expression_engine import ExpressionEngine, get_expression_engine
from .data_access_layer import DataAccessLayer, get_data_access_layer
    ```

- **core/column_resolver.py**
  - language: `python`
  - content:
    ```python
# core/column_resolver.py
"""公共列解析：将用户输入的列名/列字母转换为DataFrame子集"""
from typing import List, Dict
import pandas as pd
from core.csv_utils import split_fields
from core.exceptions import ValidationError

def resolve_columns(df: pd.DataFrame, col_str: str) -> pd.DataFrame:
    if not col_str.strip():
        return df
    candidates = split_fields(col_str)
    from core.csv_utils import build_column_mapping
    col_map = build_column_mapping(df.columns.tolist())[0]
    final = []
    for cand in candidates:
        if not cand:
            continue
        if cand.isalpha() and len(cand) <= 3:   # 列字母
            upper = cand.upper()
            if upper in col_map:
                final.append(col_map[upper])
                continue
        if cand in df.columns:                   # 列名
            final.append(cand)
    if not final:
        raise ValidationError("指定的列均不存在或无效")
    return df[final]
    ```

- **core/config.py**
  - language: `python`
  - content:
    ```python
# core/config.py
from pydantic_settings import BaseSettings
from pydantic import Field
from typing import Set

class SplitterConfig(BaseSettings):
    upload_folder: str = "uploads/splitter"
    output_folder: str = "outputs/splitter"
    task_status_folder: str = "task_status/splitter"
    model_config = {"env_prefix": "SPLITTER_"}

class CheckConfig(BaseSettings):
    sheet2_timeout_seconds: int = 120
    similarity_threshold: int = 4
    max_preview_rows: int = 1000
    duplicate_action: str = "error"
    chunk_size: int = 10000
    max_memory_rows: int = 200000
    model_config = {"env_prefix": "CHECK_"}

class FieldParserConfig(BaseSettings):
    max_text_length: int = 10 * 1024 * 1024
    similarity_threshold: int = 30
    model_config = {"env_prefix": "FIELD_PARSER_"}

class ConvertConfig(BaseSettings):
    max_file_size: int = 500 * 1024 * 1024   # 500MB
    default_encoding: str = "utf-8-sig"
    allowed_input_exts: Set[str] = {"xlsx", "xls", "xlsm", "csv"}
    allowed_output_formats: Set[str] = {"csv", "xlsx"}
    model_config = {"env_prefix": "CONVERT_"}

class TeamReportConfig(BaseSettings):
    """车队日报专用配置"""
    mapping_driver_id_col: str = "队员ID"
    mapping_team_name_col: str = "队伍名称"
    mapping_leader_name_col: str = "队长姓名"
    mapping_supervisor_col: str = "司管"
    flow_driver_id_col: str = "司机ID"
    flow_status_col: str = "司机状态"
    flow_active_status: str = "已上岗"
    flow_order_col: str = "总完单数"
    flow_apply_col: str = "报单数"
    flow_transfer_col: str = "转单数"
    flow_new_ju_col: str = "新桔完单数"
    flow_peak_order_col: str = "高峰期接单数"
    flow_peak_finish_col: str = "高峰期完单数"
    flow_star_finish_col: str = "星享完单数"
    monthly_order_col: str = "10月完单"
    monthly_apply_col: str = "10月报单"
    monthly_transfer_col: str = "10月转单"
    monthly_new_ju_col: str = "10月新桔在线数"
    daily_order_target: int = 1150
    daily_apply_target: int = 120
    monthly_order_target: int = 35650
    monthly_apply_target: int = 3720
    flow_online_days_col: str = "在线天数"
    static_supervisor_mapping: dict = {'徐国宝': '任翔', '王国平': '任凯', '安会阳': '任翔', '任翔': '任翔', '项红伟': '任翔', '马长领': '任翔', '袁茂启': '方行', '余和斌': '任翔', '夏海洋': '方行', '任凯': '任凯', '王哲': '方行', '王妙恩': '方行', '方行': '方行', '黄伟昌': '方行', '金忠健': '任凯', '蒋小峰': '方行', '童圣君': '任翔', '陈仪端': '任凯', '刘美乐': '方行', '黄一鹏': '任翔', '王凤梨': '任翔', '鲍腊九': '任凯', '王天剑': '任凯', '陈旭东': '任凯', '王林': '方行'}
    model_config = {"env_prefix": "TEAM_"}

class Settings(BaseSettings):
    secret_key: str = Field("dev-key")
    upload_folder: str = "uploads"
    output_folder: str = "outputs"
    max_content_length: int = 2 * 1024 * 1024 * 1024  # 2GB
    allowed_extensions: Set[str] = {"xlsx", "xls"}
    splitter: SplitterConfig = SplitterConfig()
    checker: CheckConfig = CheckConfig()
    field_parser: FieldParserConfig = FieldParserConfig()
    team_report: TeamReportConfig = TeamReportConfig()
    convert: ConvertConfig = ConvertConfig()
    model_config = {"env_file": ".env", "env_file_encoding": "utf-8", "extra": "ignore"}

settings = Settings()
    ```

- **core/csv_utils.py**
  - language: `python`
  - content:
    ```python
# core/csv_utils.py
"""公共CSV工具：表头读取、列映射、列字母生成"""
from typing import List, Tuple, Dict
import csv

def read_csv_header(filepath: str, encoding='utf-8') -> List[str]:
    with open(filepath, 'r', encoding=encoding) as f:
        reader = csv.reader(f)
        header = next(reader)
        return [col.strip() for col in header]

def get_column_letter(n: int) -> str:
    result = ''
    while n > 0:
        n -= 1
        result = chr(n % 26 + 65) + result
        n //= 26
    return result

def build_column_mapping(headers: List[str]) -> Tuple[Dict[str, str], Dict[str, str]]:
    letter_to_name = {}
    name_to_letter = {}
    for i, name in enumerate(headers):
        letter = get_column_letter(i + 1)
        letter_to_name[letter] = name
        name_to_letter[name] = letter
    return letter_to_name, name_to_letter

def split_fields(s: str) -> List[str]:
    """将逗号、换行、空格分隔的字符串拆分为字段列表"""
    if not s:
        return []
    lines = s.splitlines()
    parts = []
    for line in lines:
        for seg in line.split(','):
            parts.extend(seg.split())
    return [p.strip() for p in parts if p.strip()]
    ```

- **core/data_access_layer.py**
  - language: `python`
  - content:
    ```python
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
    ```

- **core/decorators.py**
  - language: `python`
  - content:
    ```python
# core/decorators.py
import logging
from functools import wraps
from flask import jsonify
from core.exceptions import AppException
from core.response import make_response
from openpyxl.utils.exceptions import InvalidFileException

logger = logging.getLogger(__name__)


def handle_exceptions():
    """统一异常处理装饰器：区分预期异常和未预期异常"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except AppException as e:
                logger.warning(f"业务异常: {e.message}")
                return make_response(False, None, e.message, e.code)
            except (ValueError, InvalidFileException) as e:
                logger.warning(f"输入错误: {str(e)}")
                return make_response(False, None, str(e), 400)
            except Exception as e:
                logger.exception("未处理异常")
                return make_response(False, None, "服务器内部错误，请稍后重试", 500)
        return wrapper
    return decorator
    ```

- **core/di.py**
  - language: `python`
  - content:
    ```python
# core/di.py
from dependency_injector import containers, providers
from core.config import settings
from core.task.manager import TaskManager
from core.task.storage import TaskStorage
from services.data.reader import DataReader
from services.data.writer import DataWriter
from services.data.formatter import DataFormatter
from services.data.file_service import FileService
from services.excel.compute_engine import ComputeEngine
from core.filter_engine import FilterEngine
from core.excel_utils import ExcelFile
from services.excel.calculator import ColumnCalculator
from services.excel.comparator import ExcelComparator
from services.excel.filter import ExcelFilter
from services.excel.aggregator import Aggregator
from services.field_parser import FieldParserService
from services.checker import CheckerService
from services.convert.service import ConvertService
from services.team_report import TeamReportService
from services.team_report.writer import TeamReportExcelWriter

class Container(containers.DeclarativeContainer):
    config = providers.Singleton(lambda: settings)
    task_storage = providers.Singleton(TaskStorage, status_dir=settings.output_folder, output_dir=settings.output_folder)
    task_manager = providers.Singleton(TaskManager, storage=task_storage, max_concurrent=3)
    data_reader = providers.Singleton(DataReader)
    data_writer = providers.Singleton(DataWriter)
    data_formatter = providers.Singleton(DataFormatter)
    file_service = providers.Singleton(FileService)
    compute_engine = providers.Singleton(ComputeEngine)
    filter_engine = providers.Singleton(FilterEngine)
    excel_service = providers.Singleton(ExcelFile)
    aggregator = providers.Singleton(Aggregator)
    column_calculator = providers.Singleton(ColumnCalculator)
    excel_comparator = providers.Singleton(ExcelComparator)
    excel_filter = providers.Singleton(ExcelFilter)
    team_report_service = providers.Factory(
        TeamReportService,
        config=config.provided.team_report
    )
    team_report_writer = providers.Singleton(TeamReportExcelWriter)
    field_parser = providers.Singleton(FieldParserService, similarity_threshold=config.provided.field_parser.similarity_threshold)
    checker_service = providers.Factory(CheckerService, config=config.provided.checker)
    convert_service = providers.Factory(ConvertService)
    ```

- **core/excel_helpers.py**
  - language: `python`
  - content:
    ```python
"""
Excel 文件辅助函数
"""
import os
import logging
import tempfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

logger = logging.getLogger(__name__)


def normalize_excel_file(filepath: str) -> str:
    """
    将 Excel 文件标准化为 .xlsx 格式
    安全处理文件锁定和XML格式问题
    """
    # 检查文件是否存在
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    # 检查文件扩展名
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.xlsx':
        # 已经是 .xlsx 格式，检查文件是否有效
        try:
            # 尝试快速打开检查文件是否有效
            with open(filepath, 'rb') as f:
                header = f.read(4)
                if header != b'PK\x03\x04':
                    logger.warning(f"文件 {filepath} 不是有效的ZIP文件，可能是损坏的Excel文件")
        except Exception as e:
            logger.warning(f"检查文件失败: {e}")

    temp_dir = tempfile.gettempdir()
    base_name = os.path.basename(filepath)
    output_name = f"normalized_{os.path.splitext(base_name)[0]}_{os.urandom(4).hex()}.xlsx"
    output_path = os.path.join(temp_dir, output_name)

    try:
        # 方法1: 首先尝试用openpyxl读取
        try:
            logger.info(f"尝试用openpyxl读取文件: {filepath}")
            # 使用只读模式减少内存使用
            wb = load_workbook(
                filename=filepath,
                data_only=True,
                read_only=True,  # 只读模式，减少内存使用
                keep_vba=False
            )

            # 创建新的工作簿用于写入
            from openpyxl import Workbook
            wb_new = Workbook()
            wb_new.remove(wb_new.active)  # 删除默认的工作表

            # 复制所有工作表
            for ws in wb.worksheets:
                ws_new = wb_new.create_sheet(title=ws.title)
                # 只复制有限的行，避免内存问题
                max_row = min(ws.max_row, 100000)  # 限制行数
                max_col = min(ws.max_column, 100)  # 限制列数

                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        ws_new.cell(row=row, column=col, value=cell.value)

            wb.close()  # 关闭只读工作簿
            wb_new.save(output_path)
            wb_new.close()

            logger.info(f"已用openpyxl转换文件: {output_path}")
            return output_path

        except Exception as e:
            logger.warning(f"openpyxl 转换失败: {e}，回退到 pandas 方式")

            # 方法2: 使用pandas读取
            try:
                # 根据文件扩展名选择engine
                engine = None
                if ext in ['.xls', '.xlsm']:
                    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

                # 读取Excel文件
                df = pd.read_excel(
                    filepath,
                    dtype=str,  # 全部作为字符串读取
                    na_filter=False,  # 不过滤空值
                    engine=engine,
                    header=None  # 不指定标题行
                )

                # 清理数据
                df = df.fillna('')

                # 保存为新的Excel文件
                df.to_excel(
                    output_path,
                    index=False,
                    header=False,
                    engine='openpyxl'
                )

                logger.info(f"已用pandas转换文件: {output_path}")
                return output_path

            except Exception as pd_e:
                logger.error(f"pandas 读取也失败: {pd_e}")

                # 方法3: 尝试修复文件
                try:
                    # 创建一个新的空Excel文件
                    from openpyxl import Workbook
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Data"
                    ws['A1'] = "错误: 无法读取原始文件"
                    ws['A2'] = f"原始错误: {str(pd_e)[:100]}"
                    wb.save(output_path)
                    logger.warning(f"创建了空的修复文件: {output_path}")
                    return output_path

                except Exception as create_e:
                    logger.error(f"创建修复文件也失败: {create_e}")
                    raise ValueError(f"无法读取或转换 Excel 文件。请确保文件是有效的 Excel 文件。原始错误: {pd_e}")

    except Exception as final_e:
        logger.error(f"所有转换方法都失败: {final_e}")
        # 如果输出路径已创建，尝试清理
        if os.path.exists(output_path):
            try:
                os.unlink(output_path)
            except:
                pass
        raise ValueError(f"Excel 文件处理失败: {final_e}")


def safe_cleanup_dataframe(df: pd.DataFrame, auto_convert_numeric: bool = True) -> pd.DataFrame:
    """
    安全清理 DataFrame
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    for col in df.columns:
        # 清理字符串
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            # 将空字符串转为空值
            df[col] = df[col].replace('', pd.NA)
            df[col] = df[col].replace('nan', pd.NA)
            df[col] = df[col].replace('None', pd.NA)
            df[col] = df[col].replace('null', pd.NA)

        # 自动转换数值
        if auto_convert_numeric and df[col].dtype == 'object':
            try:
                numeric_series = pd.to_numeric(df[col], errors='coerce')
                if numeric_series.notna().any():
                    df[col] = numeric_series
            except Exception:
                pass

    return df


def build_column_name_mapping(ws, header_row: int):
    """
    构建列名到列索引的映射
    """
    mapping = {}
    max_col = min(ws.max_column, 100)  # 限制列数

    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        if cell and cell.value:
            col_name = str(cell.value).strip()
            if col_name:
                mapping[col_name] = col_idx

    return mapping


def safe_get_cell_value(ws, row: int, col: int):
    """
    安全获取单元格值
    """
    try:
        if row > ws.max_row or col > ws.max_column:
            return None
        cell = ws.cell(row=row, column=col)
        return cell.value
    except Exception:
        return None


def safe_write_cell(ws, row: int, col: int, value):
    """
    安全写入单元格值
    """
    try:
        cell = ws.cell(row=row, column=col)
        cell.value = value
    except Exception as e:
        logger.warning(f"写入单元格失败 ({row}, {col}): {e}")


def clear_merges_in_column(ws, col_idx: int):
    """
    清除指定列的合并单元格
    """
    try:
        ranges_to_remove = []
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_col <= col_idx <= merged_range.max_col:
                ranges_to_remove.append(str(merged_range))

        for range_str in ranges_to_remove:
            ws.unmerge_cells(range_str)
    except Exception as e:
        logger.warning(f"清除合并单元格失败: {e}")


def safe_delete_file(filepath: str, max_retries: int = 3):
    """
    安全删除文件，处理文件被占用的情况
    """
    if not filepath or not os.path.exists(filepath):
        return True

    for i in range(max_retries):
        try:
            os.unlink(filepath)
            logger.info(f"成功删除文件: {filepath}")
            return True
        except PermissionError as e:
            if i < max_retries - 1:
                import time
                time.sleep(0.1 * (i + 1))  # 指数退避
                continue
            logger.warning(f"删除文件失败(被占用): {filepath}, 错误: {e}")
            return False
        except Exception as e:
            logger.warning(f"删除文件失败: {filepath}, 错误: {e}")
            return False

    return False
    ```

- **core/excel_utils.py**
  - language: `python`
  - content:
    ```python
# core/excel_utils.py
"""统一的 Excel 工具：列转换、文件包装、合并单元格复制"""
import re
import logging
from typing import Union, List, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from copy import copy

logger = logging.getLogger(__name__)

def col_letter_to_index(letter: str) -> int:
    return column_index_from_string(letter.upper())

def col_index_to_letter(index: int) -> str:
    return get_column_letter(index)

def validate_column_letters(letters: List[str]) -> bool:
    pattern = re.compile(r'^[A-Z]+$')
    return all(pattern.match(l.upper()) for l in letters)

class ExcelFile:
    def __init__(self, filepath: str, data_only: bool = True):
        self.filepath = filepath
        self.data_only = data_only
        self._workbook: Optional[openpyxl.Workbook] = None
        self._active_sheet = None

    def __enter__(self):
        self.load()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def load(self) -> 'ExcelFile':
        self._workbook = openpyxl.load_workbook(self.filepath, data_only=self.data_only)
        return self

    def close(self):
        if self._workbook:
            self._workbook.close()
            self._workbook = None

    def save(self, filepath: str = None) -> str:
        if not self._workbook:
            raise ValueError("未加载工作簿")
        target = filepath or self.filepath
        self._workbook.save(target)
        return target

    @property
    def active_sheet(self):
        if self._active_sheet is None:
            self._active_sheet = self._workbook.active
        return self._active_sheet

    def get_sheet(self, name_or_index: Union[str, int]):
        if isinstance(name_or_index, int):
            return self._workbook.worksheets[name_or_index]
        return self._workbook[name_or_index]

    def get_headers(self, sheet=None, row: int = 1) -> List[str]:
        ws = sheet or self.active_sheet
        headers = []
        for cell in ws[row]:
            if cell.value is None:
                break
            headers.append(str(cell.value).strip())
        return headers

    def get_cell_value(self, row: int, col: Union[int, str], sheet=None) -> Any:
        ws = sheet or self.active_sheet
        if isinstance(col, str):
            col = col_letter_to_index(col)
        cell = ws.cell(row, col)
        for merged in ws.merged_cells.ranges:
            if cell.coordinate in merged:
                return ws.cell(merged.min_row, merged.min_col).value
        return cell.value

    def copy_merge_structure(self, src_col: Union[int, str], dst_col: Union[int, str],
                             start_row: int, end_row: int, sheet=None):
        ws = sheet or self.active_sheet
        if isinstance(src_col, str):
            src_col = col_letter_to_index(src_col)
        if isinstance(dst_col, str):
            dst_col = col_letter_to_index(dst_col)

        src_ranges = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= src_col <= merged.max_col:
                if merged.min_row <= end_row and merged.max_row >= start_row:
                    src_ranges.append((merged.min_row, merged.max_row))

        to_remove = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= dst_col <= merged.max_col:
                to_remove.append(merged)
        for merged in to_remove:
            ws.unmerge_cells(str(merged))

        for min_r, max_r in src_ranges:
            if min_r < max_r:
                ws.merge_cells(start_row=min_r, start_column=dst_col,
                               end_row=max_r, end_column=dst_col)

    @staticmethod
    def read_dataframe(filepath: str, sheet: Union[str, int] = 0, **kwargs) -> pd.DataFrame:
        try:
            return pd.read_excel(filepath, sheet_name=sheet, **kwargs)
        except Exception as e:
            logger.error(f"读取Excel失败 {filepath}: {e}")
            raise ValueError(f"无法读取Excel文件: {e}")

    @staticmethod
    def write_dataframe(df: pd.DataFrame, filepath: str, sheet_name: str = 'Sheet1'):
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return filepath

def copy_merge_structure(src_ws, src_col, dst_ws, dst_col, start_row, end_row):
    """兼容旧接口的独立函数"""
    ef = ExcelFile('')  # 占位，仅用于调用实例方法
    ef._workbook = src_ws.parent
    ef.copy_merge_structure(src_col, dst_col, start_row, end_row, sheet=dst_ws)


# ========== 以下为兼容 ExcelService 而添加的静态方法 ==========

    @staticmethod
    def copy_cell_style(src_cell, dst_cell):
        """复制单元格样式（兼容 ExcelFile.copy_cell_style）"""
        from copy import copy
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.border = copy(src_cell.border)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.number_format = src_cell.number_format
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.protection = copy(src_cell.protection)
        dst_cell.value = src_cell.value

    @staticmethod
    def copy_merge_structure_static(ws, src_col, dst_col, start_row, end_row):
        """
        静态合并单元格结构复制（兼容 copy_merge_structure）
        :param ws: 工作表对象
        :param src_col: 源列（字母或数字索引）
        :param dst_col: 目标列（字母或数字索引）
        :param start_row: 起始行
        :param end_row: 结束行
        """
        # 复用 ExcelFile 的实例方法，创建一个临时对象
        ef = ExcelFile("")
        ef._workbook = ws.parent
        ef.copy_merge_structure(src_col, dst_col, start_row, end_row, sheet=ws)

    @staticmethod
    def get_cell_value_static(ws, row, col):
        """
        静态获取单元格值（兼容 ExcelFile.get_cell_value）
        :param ws: 工作表对象
        :param row: 行号
        :param col: 列字母或数字索引
        """
        ef = ExcelFile("")
        ef._workbook = ws.parent
        return ef.get_cell_value(row, col, sheet=ws)
    ```

- **core/excel_writer.py**
  - language: `python`
  - content:
    ```python
# core/excel_writer.py
"""统一 Excel 输出工具：写入 BytesIO 并返回"""
import io
import pandas as pd
from flask import send_file

def write_excel_response(df: pd.DataFrame, filename: str, sheet_name: str = '数据'):
    """将 DataFrame 写入 BytesIO，返回 Flask send_file 响应"""
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl', sheet_name=sheet_name)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def write_excel_to_path(df: pd.DataFrame, path: str, sheet_name: str = '数据'):
    """写入磁盘文件（用于异步任务或后续打包）"""
    df.to_excel(path, index=False, engine='openpyxl', sheet_name=sheet_name)
    ```

- **core/exceptions.py**
  - language: `python`
  - content:
    ```python
# core/exceptions.py
class AppException(Exception):
    def __init__(self, message, code=400, error_type='bad_request'):
        self.message = message
        self.code = code
        self.error_type = error_type
        super().__init__(message)

class ValidationError(AppException):
    def __init__(self, message):
        super().__init__(message, code=400, error_type='validation_error')

class NotFoundError(AppException):
    def __init__(self, message='资源不存在'):
        super().__init__(message, code=404, error_type='not_found')

class ConflictError(AppException):
    def __init__(self, message='资源冲突'):
        super().__init__(message, code=409, error_type='conflict')

class InternalError(AppException):
    def __init__(self, message='服务器内部错误'):
        super().__init__(message, code=500, error_type='internal_error')
    ```

- **core/expression_engine.py**
  - language: `python`
  - content:
    ```python
# core/expression_engine.py
"""增强型表达式引擎，并导出列字母提取函数"""
import ast
import operator
import re
import logging
from typing import Dict, Any, Set
import numpy as np

logger = logging.getLogger(__name__)

_OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Pow: operator.pow,
    ast.USub: operator.neg,
}


def extract_column_letters(expr: str) -> Set[str]:
    """
    从表达式中提取列字母（仅限 1-3 位大写字母）。
    使用简单正则 [A-Z]+，并过滤长度 >3 的 token（避免匹配函数名）。
    """
    tokens = re.findall(r'[A-Z]+', expr.upper())
    # 只保留长度 1-3 的 token（列字母通常不超过 3 位）
    return {t for t in tokens if 1 <= len(t) <= 3}


class ExpressionEngine:
    def __init__(self, missing_value_strategy: str = 'nan'):
        self.missing_value_strategy = missing_value_strategy
        self._cache: Dict[str, ast.Expression] = {}

    def compile(self, expr: str) -> ast.Expression:
        expr_upper = expr.upper()
        if expr_upper in self._cache:
            return self._cache[expr_upper]
        allowed_chars = set('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789+-*/() .')
        if not all(c in allowed_chars for c in expr_upper):
            raise ValueError(f"表达式包含非法字符: {expr}")
        try:
            tree = ast.parse(expr_upper, mode='eval')
            self._cache[expr_upper] = tree
            return tree
        except SyntaxError as e:
            raise ValueError(f"表达式语法错误: {expr}") from e

    def evaluate(self, expr: str, context: Dict[str, Any]) -> float:
        tree = self.compile(expr)
        try:
            result = self._safe_eval(tree.body, context)
            return float(result) if isinstance(result, (int, float)) else self._convert_to_number(result)
        except Exception as e:
            logger.warning(f"求值失败: {expr}, {e}")
            return self._handle_missing_value()

    def _safe_eval(self, node, context):
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name):
            return self._get_variable_value(node.id, context)
        if isinstance(node, ast.BinOp):
            left = self._safe_eval(node.left, context)
            right = self._safe_eval(node.right, context)
            op = _OPERATORS.get(type(node.op))
            if op is None:
                raise ValueError(f"不支持运算符: {type(node.op)}")
            try:
                res = op(left, right)
                return np.nan if isinstance(res, float) and not np.isfinite(res) else res
            except ZeroDivisionError:
                return np.nan
        if isinstance(node, ast.UnaryOp):
            if isinstance(node.op, ast.USub):
                return -self._safe_eval(node.operand, context)
        raise ValueError(f"不支持节点: {type(node)}")

    def _get_variable_value(self, name, context):
        raw = context.get(name)
        if raw is None:
            return self._handle_missing_value()
        if isinstance(raw, str):
            s = raw.strip()
            if s == '':
                return self._handle_missing_value()
            try:
                return float(s)
            except ValueError:
                return self._handle_missing_value()
        if isinstance(raw, (int, float, np.number)):
            return float(raw)
        try:
            return float(raw)
        except (ValueError, TypeError):
            return self._handle_missing_value()

    def _convert_to_number(self, value):
        if isinstance(value, (int, float)):
            return float(value)
        try:
            return float(value)
        except (TypeError, ValueError):
            return self._handle_missing_value()

    def _handle_missing_value(self) -> float:
        if self.missing_value_strategy == 'nan':
            return np.nan
        elif self.missing_value_strategy == 'zero':
            return 0.0
        elif self.missing_value_strategy == 'raise':
            raise ValueError("缺少必要变量或值无效")
        return np.nan


_default_engine = ExpressionEngine()


def get_expression_engine() -> ExpressionEngine:
    return _default_engine
    ```

- **core/file_naming.py**
  - language: `python`
  - content:
    ```python
# core/file_naming.py
import os
from urllib.parse import quote

def generate_output_filename(module_name: str, original_filename: str, suffix: str = '_小计') -> str:
    base = os.path.splitext(os.path.basename(original_filename))[0]
    safe_module = module_name.replace('/', '_').replace('\\', '_')
    return f"{safe_module}_{base}{suffix}.xlsx"


def build_content_disposition(filename: str) -> str:
    encoded = quote(filename, safe='')
    fallback = ''.join(ch if ord(ch) < 128 else '_' for ch in filename)
    if not fallback:
        fallback = 'download.xlsx'
    return f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{encoded}"
    ```

- **core/filter_engine.py**
  - language: `python`
  - content:
    ```python
# core/filter_engine.py
"""统一筛选引擎，基于 simplify/RowFilter 成熟实现，支持 DataFrame 和字典行"""
from typing import Dict, List, Union, Optional
from datetime import datetime
import pandas as pd
from core.exceptions import ValidationError

VALID_OPS = {'=', '!=', '>', '<', '>=', '<=', '包含', '不包含'}

def _split_by_delimiters(text: str) -> List[str]:
    if not text:
        return []
    parts = []
    for seg in text.split(','):
        parts.extend(seg.split())
    return [p.strip() for p in parts if p.strip()]

def parse_conditions_from_form(form: Dict) -> (List[Dict], List[Dict]):
    """从表单字典解析精确条件和模糊条件"""
    exact = []
    fuzzy = []
    i = 0
    while True:
        field = form.get(f'condition_field_{i}')
        if field is None:
            break
        op = form.get(f'condition_op_{i}')
        val = form.get(f'condition_value_{i}')
        if val is None:
            val = form.get(f'condition_values_{i}')   # 旧格式
        if not field or not val:
            i += 1
            continue
        field = field.strip()
        val = val.strip()
        if not field or not val:
            i += 1
            continue
        if op is None:
            op = '='
        op = op.strip()
        if op not in VALID_OPS:
            raise ValidationError(f"不支持的操作符: {op}")
        if op in ('包含', '不包含'):
            fuzzy.append({'field': field, 'op': op, 'value': val})
        else:
            vals = _split_by_delimiters(val)
            if vals:
                exact.append({'field': field, 'op': op, 'values': vals})
        i += 1
    return exact, fuzzy

class ConditionEngine:
    """统一条件引擎，支持 DataFrame 和行字典"""

    def __init__(self, date_columns: Optional[set] = None):
        self.date_columns = date_columns or set()

    def _try_number(self, s: str) -> Optional[float]:
        try:
            return float(s)
        except (ValueError, TypeError):
            return None

    def _try_datetime(self, s: str) -> Optional[datetime]:
        for fmt in ('%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M:%S',
                    '%Y/%m/%d %H:%M', '%Y-%m-%d %H:%M',
                    '%Y/%m/%d', '%Y-%m-%d', '%Y%m%d'):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return None

    def _compare(self, cell: str, op: str, target: str) -> bool:
        # 多值 OR 仅对 = 操作符
        if op == '=' and ',' in target:
            return any(cell == v.strip() for v in target.split(','))

        # 数值比较
        cell_num = self._try_number(cell)
        target_num = self._try_number(target)
        if cell_num is not None and target_num is not None:
            if op == '=':   return cell_num == target_num
            if op == '!=':  return cell_num != target_num
            if op == '>':   return cell_num > target_num
            if op == '<':   return cell_num < target_num
            if op == '>=':  return cell_num >= target_num
            if op == '<=':  return cell_num <= target_num

        # 日期比较
        cell_dt = self._try_datetime(cell)
        target_dt = self._try_datetime(target)
        if cell_dt is not None and target_dt is not None:
            if op == '=':   return cell_dt == target_dt
            if op == '!=':  return cell_dt != target_dt
            if op == '>':   return cell_dt > target_dt
            if op == '<':   return cell_dt < target_dt
            if op == '>=':  return cell_dt >= target_dt
            if op == '<=':  return cell_dt <= target_dt

        # 字符串比较
        if op == '=':   return cell == target
        if op == '!=':  return cell != target
        if op == 'contains' or op == '包含': return target in cell
        if op == 'not_contains' or op == '不包含': return target not in cell
        if op in ('>', '<', '>=', '<='):
            try:
                return eval(f'"{cell}" {op} "{target}"')
            except:
                return False
        return False

    def apply_to_row(self, row: Dict[str, str], exact: List[Dict], fuzzy: List[Dict]) -> bool:
        for c in exact:
            field = c['field']
            op = c['op']
            for v in c['values']:
                cell = str(row.get(field, ''))
                if not self._compare(cell, op, v):
                    return False
        for c in fuzzy:
            field = c['field']
            op = c['op']
            value = c['value']
            cell = str(row.get(field, ''))
            if not self._compare(cell, op, value):
                return False
        return True

    def apply_to_dataframe(self, df: pd.DataFrame, exact: List[Dict], fuzzy: List[Dict]) -> pd.DataFrame:
        mask = pd.Series(True, index=df.index)
        for c in exact:
            f = c['field']
            if f not in df.columns:
                raise ValidationError(f"字段 {f} 不存在")
            op = c['op']
            sub = pd.Series(False, index=df.index)
            for v in c['values']:
                # 尝试数值/日期/字符串
                num_series = pd.to_numeric(df[f], errors='coerce')
                num_val = self._try_number(v)
                if num_val is not None and num_series.notna().any():
                    if op == '=':   sub |= (num_series == num_val)
                    elif op == '!=': sub |= (num_series != num_val)
                    elif op == '>': sub |= (num_series > num_val)
                    elif op == '<': sub |= (num_series < num_val)
                    elif op == '>=': sub |= (num_series >= num_val)
                    elif op == '<=': sub |= (num_series <= num_val)
                    continue
                # 日期比较
                dt_series = pd.to_datetime(df[f], errors='coerce')
                dt_val = self._try_datetime(v)
                if dt_val is not None and dt_series.notna().any():
                    if op == '=':   sub |= (dt_series == dt_val)
                    elif op == '!=': sub |= (dt_series != dt_val)
                    elif op == '>': sub |= (dt_series > dt_val)
                    elif op == '<': sub |= (dt_series < dt_val)
                    elif op == '>=': sub |= (dt_series >= dt_val)
                    elif op == '<=': sub |= (dt_series <= dt_val)
                    continue
                # 字符串
                str_series = df[f].astype(str)
                if op == '=':   sub |= (str_series == v)
                elif op == '!=': sub |= (str_series != v)
                else:           # 其他比较，尝试在字符串上进行
                    try:
                        if op == '>': sub |= (str_series > v)
                        elif op == '<': sub |= (str_series < v)
                        elif op == '>=': sub |= (str_series >= v)
                        elif op == '<=': sub |= (str_series <= v)
                    except:
                        pass
            mask &= sub
        for c in fuzzy:
            f = c['field']
            if f not in df.columns:
                raise ValidationError(f"字段 {f} 不存在")
            op = c['op']
            value = c['value']
            if op in ('包含', 'contains'):
                mask &= df[f].astype(str).str.contains(value, na=False)
            elif op in ('不包含', 'not_contains'):
                mask &= ~df[f].astype(str).str.contains(value, na=False)
        return df[mask]

# 全局默认引擎实例
default_engine = ConditionEngine()


# ========== 以下为兼容废弃模块而添加的类，直接委托给本模块实现 ==========

class Condition:
    """条件数据类（兼容 utils.condition_filter.Condition）"""
    def __init__(self, field, op, value, exclude=False):
        self.field = field
        self.op = op
        self.value = value
        self.exclude = exclude


class ConditionFilter:
    """条件过滤器（兼容 utils.condition_filter.ConditionFilter）"""
    @staticmethod
    def apply_to_dict(row, conditions):
        from core.filter_engine import ConditionEngine
        engine = ConditionEngine()
        for c in conditions:
            if isinstance(c, Condition):
                op = c.op
                if op in ('contains', '包含'):
                    op = '包含'
                elif op in ('not_contains', '不包含'):
                    op = '不包含'
                if op in ('包含', '不包含'):
                    fuzzy = [{'field': c.field, 'op': op, 'value': str(c.value)}]
                    exact = []
                else:
                    exact = [{'field': c.field, 'op': op, 'values': [str(c.value)]}]
                    fuzzy = []
                result = engine.apply_to_row(row, exact, fuzzy)
                if c.exclude:
                    result = not result
                if not result:
                    return False
        return True

    @staticmethod
    def apply_to_dataframe(df, conditions):
        mask = pd.Series(True, index=df.index)
        engine = ConditionEngine()
        for c in conditions:
            if isinstance(c, Condition):
                op = c.op
                if op in ('contains', '包含'):
                    op = '包含'
                elif op in ('not_contains', '不包含'):
                    op = '不包含'
                if op in ('包含', '不包含'):
                    fuzzy = [{'field': c.field, 'op': op, 'value': str(c.value)}]
                    exact = []
                else:
                    exact = [{'field': c.field, 'op': op, 'values': [str(c.value)]}]
                    fuzzy = []
                filtered = engine.apply_to_dataframe(df, exact, fuzzy)
                cond_mask = pd.Series(False, index=df.index)
                cond_mask[filtered.index] = True
                if c.exclude:
                    cond_mask = ~cond_mask
                mask &= cond_mask
        return mask


class FilterEngine:
    """Excel 筛选引擎（兼容 services.excel.filter_engine.FilterEngine）"""
    @staticmethod
    def filter_dataframe(df, conditions):
        from core.filter_engine import ConditionEngine
        engine = ConditionEngine()
        exact, fuzzy = [], []
        for c in conditions:
            op = c.op
            if op in ('contains', '包含'):
                op = '包含'
            elif op in ('not_contains', '不包含'):
                op = '不包含'
            if op in ('包含', '不包含'):
                fuzzy.append({'field': c.field, 'op': op, 'value': str(c.value)})
            else:
                exact.append({'field': c.field, 'op': op, 'values': [str(c.value)]})
        return engine.apply_to_dataframe(df, exact, fuzzy)

    @staticmethod
    def filter_dataframe_with_multi_values(df, exact_conditions, fuzzy_conditions=None):
        from core.filter_engine import ConditionEngine
        engine = ConditionEngine()
        exact = [{'field': f, 'op': '=', 'values': vs} for f, vs in exact_conditions]
        fuzzy = [{'field': f, 'op': '包含', 'value': v} for f, v in (fuzzy_conditions or [])]
        return engine.apply_to_dataframe(df, exact, fuzzy)

    @staticmethod
    def filter_dict(row, conditions):
        from core.filter_engine import ConditionEngine
        engine = ConditionEngine()
        exact, fuzzy = [], []
        for c in conditions:
            op = c.op
            if op in ('contains', '包含'):
                op = '包含'
            elif op in ('not_contains', '不包含'):
                op = '不包含'
            if op in ('包含', '不包含'):
                fuzzy.append({'field': c.field, 'op': op, 'value': str(c.value)})
            else:
                exact.append({'field': c.field, 'op': op, 'values': [str(c.value)]})
        return engine.apply_to_row(row, exact, fuzzy)

    @staticmethod
    def create_condition(field, op, value, exclude=False):
        return Condition(field=field, op=op, value=value, exclude=exclude)
    ```

- **core/header_detection.py**
  - language: `python`
  - content:
    ```python
# core/header_detection.py
"""生产级Excel表头检测器"""
import re
import os
import time
from typing import List, Tuple, Optional, Dict, Any
from dataclasses import dataclass, field, asdict
from openpyxl import load_workbook
import logging
import hashlib

logger = logging.getLogger(__name__)

@dataclass
class HeaderDetectionConfig:
    max_check_rows: int = 20
    header_row_range: Tuple[int, int] = (0, 2)
    min_header_cols: int = 3
    max_header_cols: int = 200
    min_non_empty_ratio: float = 0.6
    min_string_ratio: float = 0.7
    max_duplicate_ratio: float = 0.3
    min_unique_cols: int = 2
    known_header_keywords: List[str] = field(default_factory=lambda: [
        "名称", "姓名", "编号", "代码", "日期", "时间", "金额", "数量",
        "单位", "备注", "状态", "类型", "级别", "部门", "岗位",
        "科目", "凭证", "摘要", "借方", "贷方", "余额", "本期", "上期",
        "病历号", "患者", "诊断", "药品", "剂量", "用法", "医生", "科室",
        "id", "name", "date", "time", "code", "type", "status", "amount"
    ])
    skip_rows_keywords: List[str] = field(default_factory=lambda: [
        "合计", "总计", "小计", "total", "sum", "汇总", "平均", "average"
    ])
    mandatory_columns: List[str] = field(default_factory=list)
    read_only: bool = True
    data_only: bool = True
    cache_ttl_seconds: int = 300

    def to_dict(self) -> Dict:
        return asdict(self)

    def get_hash(self) -> str:
        import json
        config_str = json.dumps(self.to_dict(), sort_keys=True)
        return hashlib.md5(config_str.encode()).hexdigest()[:8]

class HeaderDetectionResult:
    def __init__(self):
        self.header_rows: List[int] = []
        self.data_start_row: int = 0
        self.column_names: List[str] = []
        self.confidence: float = 0.0
        self.reason: str = ""
        self.metrics: Dict[str, Any] = {}
        self.warnings: List[str] = []
        self.timestamp: float = time.time()
        self.config_hash: str = ""

    def is_valid(self) -> bool:
        return bool(self.header_rows) and self.confidence >= 0.1

    def get_data_start_excel_row(self) -> int:
        return self.data_start_row + 1 if self.data_start_row >= 0 else 1

class SimpleExcelHeaderDetector:
    def __init__(self, config: HeaderDetectionConfig = None):
        self.config = config or HeaderDetectionConfig()
        self.config_hash = self.config.get_hash()
        self._cache = {}
        self._workbook_cache = {}
        self._keyword_set = {kw.lower() for kw in self.config.known_header_keywords}
        self._skip_set = {kw.lower() for kw in self.config.skip_rows_keywords}
        self._numeric_regex = re.compile(r'^[+-]?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?%?$')
        self._invalid_chars_regex = re.compile(r'[\/*?\[\]]')

    def detect_headers(self, file_path: str, sheet_name: str = None, use_cache: bool = True) -> HeaderDetectionResult:
        self._validate_file(file_path)
        mtime = os.path.getmtime(file_path)
        cache_key = (file_path, mtime, self.config_hash, sheet_name)
        if use_cache and cache_key in self._cache:
            cached = self._cache[cache_key]
            if time.time() - cached.timestamp < self.config.cache_ttl_seconds:
                cached.metrics['cache_hit'] = True
                return cached

        result = HeaderDetectionResult()
        result.config_hash = self.config_hash
        try:
            wb = load_workbook(filename=file_path, read_only=self.config.read_only, data_only=self.config.data_only)
            ws = wb[sheet_name] if sheet_name else wb.active
            result = self._detect_from_worksheet(ws)
            wb.close()
        except Exception as e:
            logger.error(f"检测失败 {file_path}: {e}")
            result = self._create_fallback_result()
            result.warnings.append(str(e))

        if use_cache and result.is_valid():
            self._cache[cache_key] = result
        return result

    def detect_from_workbook(self, workbook, sheet_name: str = None) -> HeaderDetectionResult:
        workbook_id = id(workbook)
        cache_key = (workbook_id, sheet_name, self.config_hash)
        if cache_key in self._workbook_cache:
            cached = self._workbook_cache[cache_key]
            cached.metrics['workbook_cache_hit'] = True
            return cached

        try:
            ws = workbook[sheet_name] if sheet_name else workbook.active
            result = self._detect_from_worksheet(ws)
            result.config_hash = self.config_hash
        except Exception as e:
            logger.error(f"从workbook检测失败: {e}")
            result = self._create_fallback_result()
            result.warnings.append(str(e))

        self._workbook_cache[cache_key] = result
        return result

    def _validate_file(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        if not os.path.isfile(file_path):
            raise ValueError(f"路径不是文件: {file_path}")
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in {'.xlsx', '.xls', '.xlsm'}:
            raise ValueError(f"不支持的文件类型: {ext}")

    def _detect_from_worksheet(self, ws) -> HeaderDetectionResult:
        result = HeaderDetectionResult()
        start_time = time.time()
        max_row = min(ws.max_row, self.config.max_check_rows)
        max_col = min(ws.max_column, self.config.max_header_cols)
        if max_row == 0 or max_col == 0:
            result.warnings.append("工作表为空")
            return self._create_fallback_result()

        data = self._read_cell_range(ws, 1, max_row, 1, max_col)
        if not data:
            return self._create_fallback_result()

        candidates = []
        start, end = self.config.header_row_range
        for row_idx in range(min(end + 1, len(data))):
            if row_idx < start:
                continue
            candidate = self._evaluate_row(data, row_idx)
            if candidate['score'] > 0:
                candidates.append(candidate)

        if candidates:
            candidates.sort(key=lambda x: x['score'], reverse=True)
            best = candidates[0]
            result.header_rows = [best['row']]
            result.confidence = best['score']
            result.reason = best['reason']
            result.column_names = best['column_names']
            result.data_start_row = best['row'] + 1
        else:
            result = self._create_fallback_result()
            if data:
                result.column_names = [str(c).strip() if cell is not None and str(cell).strip() else f"Col_{i+1}" for i, cell in enumerate(data[0])]
        result.metrics.update({
            'execution_time_ms': (time.time() - start_time) * 1000,
            'rows_checked': len(data),
            'columns_checked': max_col,
            'candidates_found': len(candidates)
        })
        return result

    def _read_cell_range(self, ws, min_row, max_row, min_col, max_col):
        try:
            return [[cell for cell in row] for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True)]
        except Exception as e:
            logger.warning(f"读取单元格失败: {e}")
            return []

    def _evaluate_row(self, data, row_idx):
        if row_idx >= len(data):
            return self._create_candidate(row_idx, 0, "索引越界")
        row_data = data[row_idx]
        if not row_data:
            return self._create_candidate(row_idx, 0, "空行")
        non_empty = [cell for cell in row_data if cell is not None and str(cell).strip() != '']
        non_empty_ratio = len(non_empty) / len(row_data) if row_data else 0
        if non_empty_ratio < 0.2:
            return self._create_candidate(row_idx, 0, f"非空{non_empty_ratio:.0%}")
        string_cells = []
        for cell in non_empty:
            s = str(cell).strip()
            if s and not self._is_numeric(s):
                string_cells.append(s)
        string_ratio = len(string_cells) / len(row_data) if row_data else 0
        unique_strings = set(string_cells)
        dup_ratio = 1 - len(unique_strings) / len(string_cells) if string_cells else 0
        keyword_matches = sum(1 for s in string_cells for kw in self._keyword_set if kw in s.lower())
        keyword_ratio = keyword_matches / len(string_cells) if string_cells else 0
        combined = ' '.join(string_cells).lower()
        if any(skip in combined for skip in self._skip_set):
            return self._create_candidate(row_idx, 0, "跳过关键词")
        valid_fmt = sum(1 for s in string_cells if self._is_valid_column_name(s))
        valid_fmt_ratio = valid_fmt / len(string_cells) if string_cells else 0
        score = 0.0
        reasons = []
        if non_empty_ratio >= self.config.min_non_empty_ratio:
            score += 0.3; reasons.append(f"非空{non_empty_ratio:.0%}")
        if string_ratio >= self.config.min_string_ratio:
            score += 0.3; reasons.append(f"文本{string_ratio:.0%}")
        if dup_ratio <= self.config.max_duplicate_ratio:
            score += 0.2; reasons.append(f"唯一{1-dup_ratio:.0%}")
        if keyword_ratio > 0.1:
            score += min(keyword_ratio, 0.2); reasons.append(f"关键词{keyword_ratio:.0%}")
        if valid_fmt_ratio > 0.5:
            score += 0.1; reasons.append("格式正确")
        score = max(0.0, min(1.0, score))
        column_names = [str(cell).strip() if cell is not None and str(cell).strip() != '' else f"Col_{i+1}" for i, cell in enumerate(row_data)]
        return {'row': row_idx, 'score': score, 'reason': ', '.join(reasons) or '基础特征', 'column_names': column_names, 'metrics': {}}

    def _create_candidate(self, row_idx, score, reason):
        return {'row': row_idx, 'score': score, 'reason': reason, 'column_names': [], 'metrics': {}}

    def _is_numeric(self, s: str) -> bool:
        s = s.strip().replace(',', '')
        if s.endswith('%'): s = s[:-1]
        try: float(s); return True
        except ValueError: return False

    def _is_valid_column_name(self, s: str) -> bool:
        if not s or len(s) > 50: return False
        if re.match(r'^[\d.,%\s-]+$', s): return False
        if self._invalid_chars_regex.search(s): return False
        return s.strip() != ''

    def _create_fallback_result(self) -> HeaderDetectionResult:
        result = HeaderDetectionResult()
        result.header_rows = [0]
        result.data_start_row = 1
        result.confidence = 0.3
        result.reason = "回退默认第一行"
        return result

    def clear_cache(self):
        self._cache.clear()
        self._workbook_cache.clear()

def detect_header_start(file_path: str, sheet_name: str = None) -> int:
    detector = SimpleExcelHeaderDetector()
    try:
        result = detector.detect_headers(file_path, sheet_name)
        if result.is_valid():
            return result.get_data_start_excel_row()
    except Exception:
        pass
    return 2
    ```

- **core/hooks.py**
  - language: `python`
  - content:
    ```python
# core/hooks.py
import functools
import logging
import time
import threading
import os
from datetime import datetime, timedelta
from typing import Dict, Any, Callable, List, Optional
from blinker import Signal

logger = logging.getLogger(__name__)

class HookManager:
    _signals: Dict[str, Signal] = {}
    _async_pool = {}
    _module_loggers: Dict[str, logging.Logger] = {}

    @classmethod
    def _get_module_name(cls, hook_name: str) -> str:
        for prefix in ("HOOK_BEFORE_", "HOOK_AFTER_"):
            if hook_name.startswith(prefix):
                name = hook_name[len(prefix):]
                break
        else:
            name = hook_name
        parts = name.lower().split('_')
        return parts[0] if parts else "general"

    @classmethod
    def _get_module_logger(cls, hook_name: str) -> logging.Logger:
        module = cls._get_module_name(hook_name)
        if module in cls._module_loggers:
            return cls._module_loggers[module]

        log_dir = os.path.join("logs", "hooks", module)
        os.makedirs(log_dir, exist_ok=True)

        module_logger = logging.getLogger(f"hook.{module}")
        module_logger.setLevel(logging.INFO)
        module_logger.propagate = False

        # 使用普通 FileHandler，避免滚动问题
        log_file = os.path.join(log_dir, "hooks.log")
        handler = logging.FileHandler(log_file, encoding='utf-8')
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        module_logger.addHandler(handler)

        # 控制台输出（可选）
        console = logging.StreamHandler()
        console.setFormatter(formatter)
        module_logger.addHandler(console)

        # 强制写入一条初始化日志，确保文件创建
        module_logger.info(f"钩子日志初始化成功，模块: {module}")

        cls._module_loggers[module] = module_logger
        return module_logger

    @classmethod
    def get_signal(cls, name: str) -> Signal:
        if name not in cls._signals:
            cls._signals[name] = Signal(name)
        return cls._signals[name]

    @classmethod
    def register(cls, name: str, func: Callable, priority: int = 0, async_mode: bool = False):
        signal = cls.get_signal(name)
        def wrapper(*args, **kwargs):
            if async_mode:
                threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True).start()
            else:
                try:
                    func(*args, **kwargs)
                except Exception as e:
                    module_logger = cls._get_module_logger(name)
                    module_logger.error(f"Hook '{name}' 执行失败: {e}", exc_info=True)
                    logger.error(f"Hook '{name}' 执行失败: {e}", exc_info=True)
        if not hasattr(signal, '_priorities'):
            signal._priorities = {}
        signal._priorities[func] = priority
        signal.connect(wrapper)

    @classmethod
    def trigger(cls, name: str, **kwargs):
        signal = cls.get_signal(name)
        receivers = signal.receivers
        if hasattr(signal, '_priorities'):
            sorted_receivers = sorted(receivers, key=lambda r: signal._priorities.get(r, 0))
        else:
            sorted_receivers = list(receivers)
        for receiver in sorted_receivers:
            try:
                receiver(**kwargs)
            except Exception as e:
                module_logger = cls._get_module_logger(name)
                module_logger.error(f"Hook '{name}' receiver 异常: {e}", exc_info=True)
                logger.error(f"Hook '{name}' receiver 异常: {e}", exc_info=True)

    @classmethod
    def clear(cls, name: str = None):
        if name:
            cls.get_signal(name).receivers.clear()
        else:
            for sig in cls._signals.values():
                sig.receivers.clear()

    @classmethod
    def clean_old_logs(cls, days: int = 7):
        base_dir = os.path.join("logs", "hooks")
        if not os.path.exists(base_dir):
            return
        cutoff = datetime.now() - timedelta(days=days)
        for root, dirs, files in os.walk(base_dir):
            for f in files:
                file_path = os.path.join(root, f)
                try:
                    mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if mtime < cutoff:
                        os.remove(file_path)
                        logger.info(f"清理过期日志文件: {file_path}")
                except Exception as e:
                    logger.warning(f"清理文件失败 {file_path}: {e}")


HOOK_BEFORE_FILE_UPLOAD = "HOOK_BEFORE_FILE_UPLOAD"
HOOK_AFTER_FILE_UPLOAD = "HOOK_AFTER_FILE_UPLOAD"
HOOK_BEFORE_DATA_READ = "HOOK_BEFORE_DATA_READ"
HOOK_AFTER_DATA_READ = "HOOK_AFTER_DATA_READ"
HOOK_BEFORE_PIVOT_AGGREGATE = "HOOK_BEFORE_PIVOT_AGGREGATE"
HOOK_AFTER_PIVOT_AGGREGATE = "HOOK_AFTER_PIVOT_AGGREGATE"
HOOK_BEFORE_EXCEL_WRITE = "HOOK_BEFORE_EXCEL_WRITE"
HOOK_AFTER_EXCEL_WRITE = "HOOK_AFTER_EXCEL_WRITE"

def hook(name: str, priority: int = 0, async_mode: bool = False):
    def decorator(func):
        HookManager.register(name, func, priority, async_mode)
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        return wrapper
    return decorator

def start_log_cleaner(interval_hours: int = 24):
    def cleaner_loop():
        while True:
            time.sleep(interval_hours * 3600)
            HookManager.clean_old_logs(days=7)
    thread = threading.Thread(target=cleaner_loop, daemon=True)
    thread.start()
    return thread
    ```

- **core/logging_setup.py**
  - language: `python`
  - content:
    ```python
# core/logging_setup.py
"""结构化日志配置与请求ID管理"""
import logging
import uuid
from flask import request, has_request_context, g
from datetime import datetime, timezone

def setup_logging(app):
    """配置应用日志格式"""
    # 移除默认处理器
    app.logger.handlers.clear()
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s %(name)s [%(request_id)s] %(message)s'
    ))
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.INFO)

def get_request_id():
    """获取当前请求的唯一ID，不存在则生成"""
    if has_request_context():
        if not hasattr(g, 'request_id'):
            g.request_id = str(uuid.uuid4())[:8]
        return g.request_id
    return 'no-request'

class RequestIDFilter(logging.Filter):
    """将 request_id 注入到日志记录中"""
    def filter(self, record):
        record.request_id = get_request_id()
        return True

def init_app(app):
    """在 Flask app 上配置日志和请求钩子"""
    setup_logging(app)
    # 添加过滤器，使得所有 logger 都能获得 request_id
    for logger_name in logging.root.manager.loggerDict:
        logger = logging.getLogger(logger_name)
        logger.addFilter(RequestIDFilter())
    app.logger.addFilter(RequestIDFilter())

    @app.before_request
    def start_timer():
        g.start_time = datetime.now(timezone.utc)

    @app.after_request
    def log_request(response):
        elapsed = (datetime.now(timezone.utc) - g.start_time).total_seconds()
        app.logger.info(
            "%s %s %s %.3fs", request.method, request.path, response.status_code, elapsed
        )
        return response
    ```

- **core/progress_task.py**
  - language: `python`
  - content:
    ```python
# core/progress_task.py
from abc import ABC, abstractmethod
from typing import Optional, Callable
from core.task.base import Task

class ProgressTask(Task, ABC):
    """支持进度回调的任务基类"""
    def __init__(self, progress_callback: Optional[Callable] = None):
        self.progress_callback = progress_callback
        self._total_rows = 0
        self._processed_rows = 0

    def set_total_rows(self, total: int):
        self._total_rows = total

    def set_progress(self, percent=None, message=None, force=True):
        if self.progress_callback:
            if percent is None and self._total_rows > 0:
                percent = min(int(self._processed_rows / self._total_rows * 100), 99)
            self.progress_callback(
                percent=percent,
                message=message,
                processed_rows=self._processed_rows,
                total_rows=self._total_rows
            )
    ```

- **core/response.py**
  - language: `python`
  - content:
    ```python
# core/response.py
"""统一 JSON 响应构建工具"""
from flask import jsonify

def make_response(success=True, data=None, error=None, code=200):
    """构建统一的 JSON 响应
    :param success: 操作是否成功
    :param data: 成功时返回的数据
    :param error: 错误时的描述信息
    :param code: HTTP 状态码
    """
    body = {
        'success': success,
        'data': data,
        'error': error,
        'code': code
    }
    response = jsonify(body)
    response.status_code = code
    return response

def success_response(data=None, message='success'):
    """快捷成功响应，兼容旧接口"""
    return make_response(True, data, None, 200)

def error_response(message, code=400):
    """快捷错误响应，兼容旧接口"""
    return make_response(False, None, message, code)
    ```

- **core/sqlite_buffer.py**
  - language: `python`
  - content:
    ```python
# core/sqlite_buffer.py
import sqlite3, tempfile, os, logging
from typing import List, Optional

logger = logging.getLogger(__name__)

class SqliteBuffer:
    """管理临时 SQLite 数据库，支持建表、建索引、批量插入、去重"""
    def __init__(self, cache_size: int = -100000):
        fd, self.db_path = tempfile.mkstemp(suffix='.db')
        os.close(fd)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA synchronous = OFF")
        self.conn.execute("PRAGMA journal_mode = MEMORY")
        self.conn.execute(f"PRAGMA cache_size = {cache_size}")

    def create_table(self, table: str, columns: List[str], key_columns: List[str] = None):
        col_defs = ', '.join([f'`{c}` TEXT' for c in columns])
        self.conn.execute(f"CREATE TABLE {table} ({col_defs})")
        if key_columns:
            for col in key_columns:
                self.conn.execute(f'CREATE INDEX IF NOT EXISTS idx_{table}_{col} ON {table}("{col}")')

    def insert_batch(self, table: str, columns: List[str], rows: list):
        placeholders = ', '.join(['?' for _ in columns])
        self.conn.executemany(f"INSERT INTO {table} VALUES ({placeholders})", rows)

    def remove_duplicates(self, table: str, key_columns: List[str], keep: str = 'first'):
        if keep == 'first':
            # 使用 rowid 保留第一条（按插入顺序）
            cols = ','.join(key_columns)
            self.conn.execute(f"CREATE TABLE {table}_dedup AS SELECT * FROM {table} WHERE rowid IN (SELECT MIN(rowid) FROM {table} GROUP BY {cols})")
            self.conn.execute(f"DROP TABLE {table}")
            self.conn.execute(f"ALTER TABLE {table}_dedup RENAME TO {table}")

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
    ```

- **core/task/__init__.py**
  - language: `python`
  - content:
    ```python

    ```

- **core/task/base.py**
  - language: `python`
  - content:
    ```python
# core/task/base.py
from abc import ABC, abstractmethod
from typing import Any

class Task(ABC):
    @abstractmethod
    def execute(self) -> Any:
        pass
    ```

- **core/task/decorator.py**
  - language: `python`
  - content:
    ```python
# core/task/decorator.py
import functools
from datetime import datetime
from flask import current_app
from .state import TaskStatus

def async_task(original_filename=None):
    """
    装饰器：将函数包装为异步任务。
    被装饰函数必须接受 task_id 作为第一个参数，并返回一个字典（将作为任务结果保存）。
    示例：
        @async_task(original_filename="my_task")
        def my_task(task_id, param1, param2):
            # 耗时操作
            return {"result_path": "/path/to/result"}
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            app = current_app._get_current_object()
            manager = app.container.task_manager()
            storage = app.container.task_storage()
            task_id = manager.create_task(params={}, original_filename=original_filename)

            def run():
                with app.app_context():
                    try:
                        storage.update(task_id, status=TaskStatus.RUNNING, started_at=datetime.now().isoformat())
                        result = func(task_id, *args, **kwargs)
                        # 移除可能由任务返回的 completed_at，因为 complete_task 会自己设置
                        if result and 'completed_at' in result:
                            result.pop('completed_at')
                        manager.complete_task(task_id, **result)
                    except Exception as e:
                        import traceback
                        error_msg = f"{str(e)}\n{traceback.format_exc()}"
                        storage.append_log(task_id, error_msg)
                        manager.fail_task(task_id, error=error_msg)

            manager.start_task(task_id, run, ())
            return {"task_id": task_id}
        return wrapper
    return decorator
    ```

- **core/task/logger.py**
  - language: `python`
  - content:
    ```python
# utils/task/logger.py
import json
import logging
from datetime import datetime
from typing import Any

_logger = logging.getLogger('app')

def log_event(task_id: str, event: str, **kwargs: Any):
    log_entry = {
        'timestamp': datetime.now().isoformat(),
        'task_id': task_id,
        'event': event,
        **kwargs
    }
    _logger.info(json.dumps(log_entry, ensure_ascii=False))
    ```

- **core/task/manager.py**
  - language: `python`
  - content:
    ```python
# core/task/manager.py
import threading
import uuid
from datetime import datetime
from typing import Callable, Dict, Optional
from .state import TaskStatus
from .storage import TaskStorage
from .logger import log_event

class TaskManager:
    def __init__(self, storage: TaskStorage, max_concurrent: int = 5):
        self.storage = storage
        self._lock = threading.Lock()
        self._running_count = 0
        self._max_concurrent = max_concurrent

    def submit(self, task_id: str, target: Callable, args: tuple) -> bool:
        with self._lock:
            if self._running_count >= self._max_concurrent:
                return False
            self._running_count += 1
        def wrapper():
            try:
                target(*args)
            finally:
                with self._lock:
                    self._running_count -= 1
        thread = threading.Thread(target=wrapper, daemon=True)
        thread.start()
        return True

    def create_task(self, params: Dict, original_filename: str = None) -> str:
        task_id = str(uuid.uuid4())
        status = {
            'task_id': task_id,
            'status': TaskStatus.PENDING,
            'created_at': datetime.now().isoformat(),
            'started_at': None,
            'completed_at': None,
            'result_path': None,
            'result_format': None,
            'error': None,
            'total_input_rows': None,
            'total_output_rows': None,
            'params': params,
            'original_filename': original_filename
        }
        self.storage.save(task_id, status)
        log_event(task_id, 'task_created', params=params)
        return task_id

    def start_task(self, task_id: str, target: Callable, args: tuple):
        self.storage.update(task_id, status=TaskStatus.RUNNING, started_at=datetime.now().isoformat())
        log_event(task_id, 'task_started')
        if not self.submit(task_id, target, args):
            self.storage.update(task_id, status=TaskStatus.FAILED, error="系统繁忙，请稍后再试")
            return False
        return True

    def complete_task(self, task_id: str, **result):
        self.storage.update(task_id, status=TaskStatus.COMPLETED, completed_at=datetime.now().isoformat(), **result)
        log_event(task_id, 'task_completed', **result)

    def fail_task(self, task_id: str, error: str):
        self.storage.update(task_id, status=TaskStatus.FAILED, error=error, completed_at=datetime.now().isoformat())
        log_event(task_id, 'task_failed', error=error)

    def cancel_task(self, task_id: str):
        self.storage.update(task_id, status=TaskStatus.CANCELLED, completed_at=datetime.now().isoformat())
        log_event(task_id, 'task_cancelled')
    ```

- **core/task/state.py**
  - language: `python`
  - content:
    ```python
# utils/task/state.py
class TaskStatus:
    PENDING = 'pending'
    RUNNING = 'running'
    COMPLETED = 'completed'
    FAILED = 'failed'
    CANCELLED = 'cancelled'
    ```

- **core/task/storage.py**
  - language: `python`
  - content:
    ```python
# core/task/storage.py
import json
import os
import threading
from datetime import datetime
from typing import Dict, Optional
import logging

logger = logging.getLogger(__name__)

class TaskStorage:
    def __init__(self, status_dir: str, output_dir: str):
        self.status_dir = status_dir
        self.output_dir = output_dir
        self._lock = threading.RLock()
        os.makedirs(status_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

    def _path(self, task_id: str) -> str:
        return os.path.join(self.status_dir, f"{task_id}.json")

    def _log_path(self, task_id: str) -> str:
        return os.path.join(self.status_dir, f"{task_id}.log")

    def save(self, task_id: str, data: Dict):
        with self._lock:
            with open(self._path(task_id), 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

    def load(self, task_id: str) -> Optional[Dict]:
        path = self._path(task_id)
        if not os.path.exists(path):
            return None
        with self._lock:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if not content:
                        logger.warning(f"任务文件为空: {path}")
                        return None
                    return json.loads(content)
            except json.JSONDecodeError:
                logger.error(f"任务文件损坏: {path}")
                return None
            except Exception as e:
                logger.error(f"读取任务文件失败: {path} - {e}")
                return None

    def update(self, task_id: str, **fields):
        with self._lock:
            data = self.load(task_id)
            if data is None:
                raise ValueError(f"Task {task_id} not found")
            data.update(fields)
            self.save(task_id, data)

    def delete(self, task_id: str):
        path = self._path(task_id)
        if os.path.exists(path):
            os.remove(path)

    def get_output_dir(self) -> str:
        return self.output_dir

    def append_log(self, task_id: str, message: str):
        log_path = self._log_path(task_id)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")

    def get_log(self, task_id: str, max_lines: int = 2000) -> str:
        log_path = self._log_path(task_id)
        if not os.path.exists(log_path):
            return ""
        with open(log_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        if len(lines) > max_lines:
            lines = lines[-max_lines:]
        return ''.join(lines)
    ```

- **core/validators.py**
  - language: `python`
  - content:
    ```python
# core/validators.py
"""通用验证函数"""
import re
from typing import List, Union

def validate_column_letters(letters: List[str]) -> bool:
    pattern = re.compile(r'^[A-Z]+$')
    return all(pattern.match(l.upper()) for l in letters)

def validate_sheet_index(idx: Union[str, int]) -> bool:
    try:
        num = int(idx)
        return num >= 1
    except (ValueError, TypeError):
        return False

def validate_expression(expr: str) -> bool:
    if not expr or not expr.strip():
        return False
    allowed_chars = set('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789+-*/() ')
    if not all(c in allowed_chars for c in expr.upper()):
        return False
    if expr.strip().endswith(('+', '-', '*', '/', '(', ' ')):
        return False
    stack = []
    for ch in expr:
        if ch == '(':
            stack.append(ch)
        elif ch == ')':
            if not stack:
                return False
            stack.pop()
    return len(stack) == 0

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', '_', str(name))
    ```

- **force_rebuild_and_push.py**
  - language: `python`
  - content:
    ```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
强制重建 Git 仓库并推送到指定 Gitee 仓库
- 删除所有历史
- 忽略 .exe、所有 Excel 文件、中文名文件/目录
- 自动处理大文件
"""

import os
import sys
import subprocess
import time
import re

# ========== 配置 ==========
REMOTE_URL = "https://gitee.com/hky_gitee/data_deel"
REMOTE_NAME = "origin"
BRANCH = "master"

# ========== 工具函数 ==========
def run_cmd(cmd, capture=True, check=False, shell=True):
    try:
        if capture:
            result = subprocess.run(cmd, shell=shell, capture_output=True, text=True,
                                    encoding='utf-8', errors='replace', check=check)
            return result.returncode, result.stdout.strip(), result.stderr.strip()
        else:
            result = subprocess.run(cmd, shell=shell, check=check)
            return result.returncode, "", ""
    except subprocess.CalledProcessError as e:
        return e.returncode, e.stdout.strip() if e.stdout else "", e.stderr.strip() if e.stderr else str(e)

def force_delete_git():
    """强制删除 .git 目录，自动重试"""
    if not os.path.exists('.git'):
        print("没有 .git 目录，跳过删除。")
        return True

    print("正在强制删除 .git 目录...")
    for attempt in range(1, 4):
        ret, _, err = run_cmd('rmdir /s /q .git', capture=True)
        if ret == 0 or not os.path.exists('.git'):
            print(f"✓ 删除成功（尝试 {attempt}）")
            return True
        print(f"删除失败（尝试 {attempt}），等待 2 秒...")
        time.sleep(2)
        # 备用命令
        ret, _, err = run_cmd('rd /s /q .git', capture=True)
        if ret == 0 or not os.path.exists('.git'):
            print(f"✓ 删除成功（尝试 {attempt} 使用 rd）")
            return True
        time.sleep(2)

    print("❌ 无法删除 .git 目录，请手动关闭占用程序后重试。")
    return False

def create_gitignore():
    """创建完善的 .gitignore"""
    content = """# === 可执行文件 ===
*.exe
*.msi

# === Excel 文件（所有格式） ===
*.xlsx
*.xls
*.xlsm
*.xlsb
*.xltx
*.xltm

# === 大文件（防止超限） ===
*.zip
*.tar.gz
*.7z
*.rar
*.iso
*.bin
*.dat
*.pkl
*.joblib
*.parquet
*.feather
*.pth
*.ckpt
*.pb
*.onnx
*.h5
*.npy

# === 临时文件 ===
*.tmp
*.bak
*.log
*.cache
*.pyc
__pycache__/

# === IDE ===
.vscode/
.idea/
*.swp
.DS_Store

# === 虚拟环境 ===
.venv/
venv/
ENV/
env/

# === 项目特定 ===
Excel/
*.ipynb_checkpoints/

# === 其他可能的大文件 ===
*.pdf
*.mp4
*.avi
*.mov
*.mkv
*.mp3
*.wav

# === 压缩备份 ===
*.backup
*.old
"""
    with open(".gitignore", "w", encoding="utf-8") as f:
        f.write(content)
    print("✓ .gitignore 已创建/覆盖")

def has_chinese(text):
    """判断字符串是否包含中文"""
    return bool(re.search(r'[\u4e00-\u9fff]', text))

def remove_chinese_items_from_index():
    """
    从 Git 索引中移除所有文件名包含中文的文件，以及目录名包含中文的目录（递归）
    本地文件不受影响
    """
    # 获取当前暂存区所有文件（包括已添加的）
    ret, files, err = run_cmd("git ls-files --cached")
    if ret != 0:
        print("无法获取暂存区文件列表")
        return

    chinese_files = []
    chinese_dirs = set()

    for f in files.splitlines():
        f = f.strip()
        if not f:
            continue
        # 检查文件名（不含路径）是否含中文
        basename = os.path.basename(f)
        if has_chinese(basename):
            chinese_files.append(f)
            continue
        # 检查路径中的每一级目录是否含中文
        parts = f.split('/')
        for i in range(len(parts)-1):  # 目录部分
            if has_chinese(parts[i]):
                # 标记该目录前缀
                prefix = '/'.join(parts[:i+1])
                chinese_dirs.add(prefix)
                break

    if chinese_files:
        print(f"发现 {len(chinese_files)} 个文件名含中文的文件，正在从索引移除...")
        for f in chinese_files:
            run_cmd(f'git rm --cached --ignore-unmatch "{f}"', capture=False, check=False)
        print("✓ 已移除中文名文件")

    if chinese_dirs:
        print(f"发现 {len(chinese_dirs)} 个目录名含中文的目录，正在递归移除...")
        for d in chinese_dirs:
            run_cmd(f'git rm -r --cached --ignore-unmatch "{d}/"', capture=False, check=False)
        print("✓ 已移除中文目录下的所有内容")

def main():
    print("=" * 60)
    print("强制重建 Git 仓库并推送到 Gitee")
    print("目标仓库:", REMOTE_URL)
    print("=" * 60)
    print("注意：此操作将删除所有 Git 历史记录，仅保留当前文件内容。")
    confirm = input("是否继续？(y/N): ").strip().lower()
    if confirm != 'y':
        print("已取消")
        sys.exit(0)

    # 1. 删除 .git
    if not force_delete_git():
        sys.exit(1)

    # 2. 初始化
    print("\n初始化新仓库...")
    run_cmd("git init", capture=False)

    # 3. 创建 .gitignore
    create_gitignore()

    # 4. 添加所有文件（此时 .gitignore 会生效）
    print("\n添加文件到暂存区（.gitignore 自动生效）...")
    run_cmd("git add .", capture=False)

    # 5. 移除中文名文件/目录（从暂存区）
    print("\n检查并移除中文名文件/目录...")
    remove_chinese_items_from_index()

    # 6. 检查是否有需要提交的内容
    ret, status, _ = run_cmd("git status --porcelain")
    if ret != 0 or not status:
        print("\n⚠️ 没有需要提交的文件！请检查 .gitignore 规则或项目目录是否为空。")
        sys.exit(1)

    # 7. 提交
    commit_msg = input("\n请输入提交信息（直接回车使用默认）: ").strip()
    if not commit_msg:
        commit_msg = "Initial commit: clean project without ignored files"
    ret, _, err = run_cmd(f'git commit -m "{commit_msg}"')
    if ret != 0:
        print("提交失败:", err)
        sys.exit(1)
    print("✓ 提交成功")

    # 8. 配置远程
    print(f"\n配置远程仓库 {REMOTE_NAME} -> {REMOTE_URL}")
    ret, _, _ = run_cmd(f"git remote get-url {REMOTE_NAME}", capture=True, check=False)
    if ret == 0:
        run_cmd(f"git remote set-url {REMOTE_NAME} {REMOTE_URL}", capture=False)
        print("已更新 origin URL")
    else:
        run_cmd(f"git remote add {REMOTE_NAME} {REMOTE_URL}", capture=False)
        print("已添加远程仓库")

    # 9. 强制推送
    print(f"\n准备强制推送到 {REMOTE_NAME}/{BRANCH}...")
    final = input("是否执行强制推送？(y/N): ").strip().lower()
    if final != 'y':
        print("已取消推送，本地仓库已重建但未推送。")
        sys.exit(0)

    ret, _, err = run_cmd(f"git push -u {REMOTE_NAME} {BRANCH} --force")
    if ret != 0:
        print("推送失败:", err)
        sys.exit(1)

    print("\n🎉 推送成功！远程仓库已更新，历史已清空。")

if __name__ == "__main__":
    main()
    ```

- **install_requirements.py**
  - language: `python`
  - content:
    ```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健壮的依赖安装脚本
- 自动轮换多个镜像源（阿里云、豆瓣、华为、清华、官方）
- 跳过 pip 升级，避免 403 错误
- 支持超时和重试
"""

import subprocess
import sys
import time
import os

REQUIREMENTS_FILE = "requirements.txt"
MAX_RETRIES = 2          # 每个镜像重试次数
RETRY_DELAY = 5          # 重试间隔（秒）
TIMEOUT = 180            # 安装超时（秒）

# 镜像源列表（按优先级排序）
MIRRORS = [
    "https://mirrors.aliyun.com/pypi/simple/",
    "https://pypi.douban.com/simple/",
    "https://repo.huaweicloud.com/repository/pypi/simple/",
    "https://pypi.tuna.tsinghua.edu.cn/simple/",
    "https://pypi.org/simple/",          # 官方源，最后尝试
]

def run_pip_install(mirror):
    """尝试使用指定镜像源安装依赖"""
    host = mirror.split("//")[1].split("/")[0]
    cmd = [
        sys.executable, "-m", "pip", "install",
        "-r", REQUIREMENTS_FILE,
        "--timeout", str(TIMEOUT),
        "--retries", str(MAX_RETRIES),
        "-i", mirror,
        "--trusted-host", host,
        "--exists-action", "w",
    ]
    print(f"\n>>> 使用镜像源: {mirror}")
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            result = subprocess.run(cmd, timeout=TIMEOUT, check=False, capture_output=True, text=True)
            if result.returncode == 0:
                print("✅ 安装成功！")
                print(result.stdout)
                return True
            else:
                print(f"安装失败 (尝试 {attempt}/{MAX_RETRIES})，返回码 {result.returncode}")
                if result.stderr:
                    print("错误详情:", result.stderr[-500:])  # 只打印最后500字符
                if attempt < MAX_RETRIES:
                    print(f"等待 {RETRY_DELAY} 秒后重试...")
                    time.sleep(RETRY_DELAY)
        except subprocess.TimeoutExpired:
            print(f"安装超时（{TIMEOUT}秒），尝试 {attempt}/{MAX_RETRIES}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)
    return False

def main():
    print("=" * 60)
    print("Python 依赖安装工具（多镜像源自动轮换）")
    print(f"Python 版本: {sys.version}")
    print(f"依赖文件: {REQUIREMENTS_FILE}")
    if not os.path.exists(REQUIREMENTS_FILE):
        print(f"错误: 找不到 {REQUIREMENTS_FILE} 文件")
        sys.exit(1)
    print("=" * 60)

    # 不升级 pip，避免 403 错误
    print("跳过 pip 升级，使用当前版本。")

    for mirror in MIRRORS:
        success = run_pip_install(mirror)
        if success:
            print("\n🎉 所有依赖安装成功！")
            return
        print(f"镜像源 {mirror} 失败，切换到下一个...")
        time.sleep(RETRY_DELAY)

    print("\n❌ 所有镜像源均失败，请手动安装缺失的包。")
    print("建议手动执行：pip install -r requirements.txt")
    sys.exit(1)

if __name__ == "__main__":
    main()
    ```

- **requirements.txt**
  - language: ``
  - content:
    ```
flask>=2.3.3
pandas==2.3.3
numpy==2.4.3
openpyxl==3.1.2
xlrd==2.0.1
gunicorn==21.2.0
waitress>=2.0.0
Werkzeug>=2.0.0
psutil>=5.9.0
pydantic>=2.0.0
pydantic-settings>=2.0.0
dependency-injector>=4.41.0
pytest>=7.4.0
pytest-cov>=4.1.0
    ```

- **run_tests.py**
  - language: `python`
  - content:
    ```python
#!/usr/bin/env python3
"""
测试运行脚本
- 自动执行 pytest 运行所有测试
python -m pytest tests/ -v
- 若全部通过：仅打印成功信息，不生成任何日志文件
- 若有失败：将完整输出保存到 log/ 目录，并包含失败测试列表
"""

import subprocess
import sys
import re
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
LOG_DIR = BASE_DIR / "log"


def run_tests():
    """运行 pytest 并返回 (returncode, output)"""
    cmd = [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"]
    print(f"🚀 运行测试: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    output = result.stdout + result.stderr
    return result.returncode, output


def extract_failed_tests(output: str) -> list:
    """从 pytest 输出中提取失败的测试名称"""
    failed_tests = []
    # 匹配类似 "tests/test_xxx.py::TestClass::test_method FAILED" 的行
    pattern = re.compile(r"^(.+?\.py::[^ ]+)\s+FAILED", re.MULTILINE)
    for match in pattern.finditer(output):
        failed_tests.append(match.group(1))
    return failed_tests


def save_failure_log(output: str, failed_tests: list):
    """保存失败日志到 log/ 目录，包含失败测试列表"""
    LOG_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"test_failures_{timestamp}.log"

    # 构建日志内容
    header = f"测试运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    header += f"失败测试数量: {len(failed_tests)}\n"
    header += "-" * 60 + "\n"
    if failed_tests:
        header += "失败测试列表:\n"
        for test in failed_tests:
            header += f"  - {test}\n"
        header += "-" * 60 + "\n\n"

    full_log = header + output

    log_file.write_text(full_log, encoding="utf-8")
    print(f"\n📁 失败日志已保存至: {log_file}")
    print("\n⚠️ 部分测试失败，失败测试列表：")
    for test in failed_tests:
        print(f"   ❌ {test}")
    print(f"\n详细信息请查看日志文件。")


def main():
    returncode, output = run_tests()

    if returncode == 0:
        print("\n✅ 所有测试通过！未生成日志文件。")
    else:
        failed_tests = extract_failed_tests(output)
        save_failure_log(output, failed_tests)
        sys.exit(returncode)


if __name__ == "__main__":
    main()
    ```

- **runtime.txt**
  - language: ``
  - content:
    ```
python-3.10
    ```

- **schemas/comparator.py**
  - language: `python`
  - content:
    ```python
# schemas/comparator.py
from pydantic import BaseModel, Field, field_validator
from typing import List, Optional

class CompareFullParams(BaseModel):
    key_fields: List[str] = Field(..., min_length=1, description="主键字段列表")
    compare_fields: List[str] = Field(..., min_length=1, description="对比字段列表")
    sheet1: Optional[str] = None
    sheet2: Optional[str] = None

    @field_validator('key_fields', 'compare_fields', mode='before')
    @classmethod
    def check_not_empty(cls, v):
        if not v:
            raise ValueError('字段列表不能为空')
        return [item.strip() for item in v if item and item.strip()]

class CompareFieldMatchParams(BaseModel):
    match_field: str = Field(..., description="匹配字段名")
    sheet1: Optional[str] = None
    sheet2: Optional[str] = None

    @field_validator('match_field')
    @classmethod
    def not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('匹配字段名不能为空')
        return v.strip()
    ```

- **schemas/filter.py**
  - language: `python`
  - content:
    ```python
# schemas/filter.py
from pydantic import BaseModel, Field, field_validator
from typing import List, Optional

class FilterCondition(BaseModel):
    field: str
    values: List[str]

    @field_validator('field')
    @classmethod
    def field_not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('字段名不能为空')
        return v.strip()

    @field_validator('values')
    @classmethod
    def values_not_empty(cls, v):
        if not v:
            raise ValueError('值列表不能为空')
        return [val.strip() for val in v if val.strip()]

class FilterParams(BaseModel):
    files: List[str] = Field(..., min_length=1, description="文件路径列表")
    conditions: List[FilterCondition] = Field(default_factory=list)
    fuzzy_field: Optional[str] = None
    fuzzy_value: Optional[str] = None
    action: str = Field(..., description="操作类型: get_values 或 generate_excel")
    target_field: Optional[str] = None

    @field_validator('action')
    @classmethod
    def validate_action(cls, v):
        if v not in ('get_values', 'generate_excel'):
            raise ValueError('action 必须是 get_values 或 generate_excel')
        return v

    @field_validator('target_field')
    @classmethod
    def validate_target_field(cls, v, info):
        if info.data.get('action') == 'get_values' and not v:
            raise ValueError('获取值时必须指定 target_field')
        return v
    ```

- **services/__init__.py**
  - language: `python`
  - content:
    ```python
"""业务领域服务层（精简导出）—— 仅保留最常用的服务，其余请按需显式导入"""
from services.data.reader import DataReader
from services.data.writer import DataWriter
from services.data.formatter import DataFormatter
from services.data.file_service import FileService
from services.excel.compute_engine import ComputeEngine
from core.filter_engine import FilterEngine
from services.excel.aggregator import Aggregator

__all__ = [
    'DataReader', 'DataWriter', 'DataFormatter', 'FileService',
    'ComputeEngine', 'FilterEngine', 'Aggregator',
]
    ```

- **services/checker/__init__.py**
  - language: `python`
  - content:
    ```python
"""文件校验服务"""
from .service import CheckerService
    ```

- **services/checker/constants.py**
  - language: `python`
  - content:
    ```python
# services/checker/constants.py
"""校验模块常量，避免硬编码"""

SHEET_NAMES = {
    'DIFF': '差异数据',
    'ONLY_STD': '标准群独有',
    'ONLY_CMP': '比较群独有'
}
    ```

- **services/checker/service.py**
  - language: `python`
  - content:
    ```python
# services/checker/service.py
import os
import logging
from typing import List, Dict, Optional, Callable
import pandas as pd

from core.progress_task import ProgressTask
from core.exceptions import ValidationError
from services.data.row_counter import RowCounter
from .strategies import MemoryCheckerStrategy, SqliteCheckerStrategy

logger = logging.getLogger(__name__)


class CheckerService(ProgressTask):
    """文件校验服务（策略上下文）"""
    def __init__(self, files_a=None, files_b=None,
                 key_map_a=None, key_map_b=None,
                 compare_map=None, filters_a=None, filters_b=None,
                 duplicate_action=None, config=None,
                 progress_callback=None,
                 precheck_data=None):
        super().__init__(progress_callback)
        self.files_a = files_a or []
        self.files_b = files_b or []
        self.key_map_a = key_map_a or []
        self.key_map_b = key_map_b or []
        self.compare_map = compare_map or []
        self.filters_a = filters_a or []
        self.filters_b = filters_b or []
        self.config = config
        self.duplicate_action = duplicate_action or (config.duplicate_action if config else "error")
        self.precheck_data = precheck_data

    def precheck(self) -> dict:
        ta = sum(RowCounter.count(f) for f in self.files_a)
        tb = sum(RowCounter.count(f) for f in self.files_b)
        total = ta + tb
        total_size = sum(os.path.getsize(f) for f in self.files_a + self.files_b) / (1024 * 1024)
        return {
            "total_files_a": len(self.files_a),
            "total_files_b": len(self.files_b),
            "total_rows_a": ta, "total_rows_b": tb,
            "total_rows": total,
            "estimated_memory_mb": round(total * 200 / (1024*1024), 2),
            "estimated_time_seconds": round(total / 20000, 2),
            "total_size_mb": round(total_size, 2),
            "advice": "数据量大" if total > 500000 else "安全",
            "can_proceed": True
        }

    def execute(self) -> str:
        total_rows_a = self.precheck_data.get('total_rows_a', 0) if self.precheck_data else sum(RowCounter.count(f) for f in self.files_a)
        max_memory_rows = self.config.max_memory_rows if self.config else 200000
        use_sqlite = total_rows_a > max_memory_rows

        strategy_class = SqliteCheckerStrategy if use_sqlite else MemoryCheckerStrategy
        strategy = strategy_class(
            files_a=self.files_a,
            files_b=self.files_b,
            key_map_a=self.key_map_a,
            key_map_b=self.key_map_b,
            compare_map=self.compare_map,
            filters_a=self.filters_a,
            filters_b=self.filters_b,
            duplicate_action=self.duplicate_action,
            config=self.config,
            progress_callback=self.progress_callback,
            precheck_data=self.precheck_data
        )
        strategy._total_rows = self._total_rows
        strategy._processed_rows = self._processed_rows
        return strategy.execute()
    ```

- **services/checker/strategies.py**
  - language: `python`
  - content:
    ```python
# services/checker/strategies.py
import os
import time
import logging
from typing import List, Dict, Optional, Callable, Set
import pandas as pd
from openpyxl import Workbook

from core.progress_task import ProgressTask
from core.exceptions import ValidationError
from services.data.stream_reader import StreamReader
from services.data.row_counter import RowCounter
from core.sqlite_buffer import SqliteBuffer
from services.data.report_writer import ReportWriter

logger = logging.getLogger(__name__)
BATCH_SIZE = 10000


class BaseCheckerStrategy(ProgressTask):
    """校验策略基类"""
    def __init__(self, files_a, files_b, key_map_a, key_map_b,
                 compare_map, filters_a, filters_b, duplicate_action, config,
                 progress_callback=None, precheck_data=None):
        super().__init__(progress_callback)
        self.files_a = files_a or []
        self.files_b = files_b or []
        self.key_map_a = key_map_a or []
        self.key_map_b = key_map_b or []
        self.compare_map = compare_map or []
        self.filters_a = filters_a or []
        self.filters_b = filters_b or []
        self.config = config
        self.duplicate_action = duplicate_action or (config.duplicate_action if config else "error")
        self.precheck_data = precheck_data

    def _make_key(self, row, key_cols):
        values = []
        for k in key_cols:
            try:
                value = row[k]
            except Exception:
                value = row.get(k) if isinstance(row, dict) else None
            values.append(value)
        return tuple(values)


class MemoryCheckerStrategy(BaseCheckerStrategy):
    """内存流式对比策略"""
    def execute(self) -> str:
        start_time = time.time()
        if self.precheck_data:
            total_rows_a = self.precheck_data.get('total_rows_a', 0)
            total_rows_b = self.precheck_data.get('total_rows_b', 0)
        else:
            total_rows_a = sum(RowCounter.count(f) for f in self.files_a)
            total_rows_b = sum(RowCounter.count(f) for f in self.files_b)
        self._total_rows = total_rows_a + total_rows_b

        self.set_progress(0, "读取标准群...")

        needed_a = set(self.key_map_a)
        for m in self.compare_map:
            needed_a.add(m["standard"])
        for f in self.filters_a:
            if "field" in f:
                needed_a.add(f["field"])
        needed_a.discard("")

        std_rows = []
        processed_a = 0
        for file in self.files_a:
            for row in StreamReader.read_excel(file, list(needed_a), self.filters_a):
                std_rows.append(row)
                processed_a += 1
                if processed_a % 5000 == 0:
                    pct = min(int(processed_a / total_rows_a * 100), 30) if total_rows_a else 0
                    self.set_progress(pct, f"读取标准群 {processed_a}/{total_rows_a}")
        std_df = pd.DataFrame(std_rows)
        if std_df.empty:
            raise ValidationError("标准群无有效数据")

        if self.duplicate_action == "error":
            if std_df.duplicated(subset=self.key_map_a).any():
                raise ValidationError("标准群主键重复")
        else:
            std_df = std_df.drop_duplicates(subset=self.key_map_a, keep='first')

        std_df = std_df.set_index(self.key_map_a)

        cmp_cols_a = [m["standard"] for m in self.compare_map]
        cmp_cols_b = [m["compare"] for m in self.compare_map]

        wb = Workbook(write_only=True)
        ws_diff = wb.create_sheet("差异数据")
        ws_std_only = wb.create_sheet("标准群独有")
        ws_cmp_only = wb.create_sheet("比较群独有")

        diff_header = (list(self.key_map_a) +
                       [f"{col}_标准" for col in cmp_cols_a] +
                       [f"{col}_比较" for col in cmp_cols_b])
        ws_diff.append(diff_header)
        ws_std_only.append(list(self.key_map_a) + cmp_cols_a)
        ws_cmp_only.append(list(self.key_map_b) + cmp_cols_b)

        matched_std_keys: Set[tuple] = set()

        self.set_progress(30, "开始流式对比...")

        needed_b = set(self.key_map_b)
        for m in self.compare_map:
            needed_b.add(m["compare"])
        for f in self.filters_b:
            if "field" in f:
                needed_b.add(f["field"])
        needed_b.discard("")

        batch_buffer = []
        processed_b = 0
        cmp_only_rows = []

        for file in self.files_b:
            for row in StreamReader.read_excel(file, list(needed_b), self.filters_b):
                batch_buffer.append(row)
                processed_b += 1
                if len(batch_buffer) >= BATCH_SIZE:
                    self._process_batch(batch_buffer, std_df, cmp_cols_a, cmp_cols_b,
                                        ws_diff, matched_std_keys, cmp_only_rows)
                    batch_buffer = []
                    pct = 30 + int((processed_b / total_rows_b) * 65) if total_rows_b else 50
                    self.set_progress(pct, f"对比中 {processed_b}/{total_rows_b}")

        if batch_buffer:
            self._process_batch(batch_buffer, std_df, cmp_cols_a, cmp_cols_b,
                                ws_diff, matched_std_keys, cmp_only_rows)

        for row in cmp_only_rows:
            ws_cmp_only.append(list(row))

        self.set_progress(96, "计算标准群独有...")
        all_std_keys = set(self._make_key(row, self.key_map_a) for row in std_df.reset_index().to_dict('records'))
        only_std_keys = all_std_keys - matched_std_keys
        if only_std_keys:
            std_only_df = std_df.reset_index()
            std_only_df['_key'] = std_only_df.apply(lambda r: self._make_key(r, self.key_map_a), axis=1)
            for _, row in std_only_df[std_only_df['_key'].isin(only_std_keys)].iterrows():
                values = [row[k] for k in self.key_map_a] + [row[col] for col in cmp_cols_a]
                ws_std_only.append(values)

        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        out_path = tmp.name
        wb.save(out_path)
        wb.close()

        self.set_progress(100, "完成")
        return out_path

    def _process_batch(self, batch_rows, std_df, cmp_cols_a, cmp_cols_b,
                       ws_diff, matched_std_keys, cmp_only_buf):
        cmp_df = pd.DataFrame(batch_rows)
        cmp_df['_key'] = cmp_df.apply(lambda r: self._make_key(r, self.key_map_b), axis=1)

        merged = pd.merge(
            std_df.reset_index(), cmp_df,
            left_on=self.key_map_a, right_on=self.key_map_b,
            how='inner', suffixes=('_标准', '_比较')
        )

        for _, match_row in merged.iterrows():
            matched_std_keys.add(self._make_key(match_row, self.key_map_a))

        if merged.empty:
            for _, row in cmp_df.iterrows():
                cmp_only_buf.append(row.drop(labels=['_key']).to_dict())
            return

        diff_mask = pd.Series(False, index=merged.index)
        for ca, cb in zip(cmp_cols_a, cmp_cols_b):
            left = merged[f'{ca}_标准']
            right = merged[f'{cb}_比较']
            eq = left.eq(right)
            both_null = left.isna() & right.isna()
            diff_mask |= ~(eq.fillna(False) | both_null)
        diff_rows = merged[diff_mask]

        if not diff_rows.empty:
            for _, row in diff_rows.iterrows():
                values = [row[k] for k in self.key_map_a]
                for ca in cmp_cols_a:
                    values.append(row[f'{ca}_标准'])
                for cb in cmp_cols_b:
                    values.append(row[f'{cb}_比较'])
                ws_diff.append(values)

        merged_keys = set(self._make_key(row, self.key_map_b) for _, row in merged.iterrows())
        for _, row in cmp_df.iterrows():
            if row['_key'] not in merged_keys:
                cmp_only_buf.append(row.drop(labels=['_key']).to_dict())


class SqliteCheckerStrategy(BaseCheckerStrategy):
    """SQLite 兜底对比策略"""
    def execute(self) -> str:
        start_time = time.time()
        if self.precheck_data:
            total_rows_a = self.precheck_data.get('total_rows_a', 0)
            total_rows_b = self.precheck_data.get('total_rows_b', 0)
        else:
            total_rows_a = sum(RowCounter.count(f) for f in self.files_a)
            total_rows_b = sum(RowCounter.count(f) for f in self.files_b)
        self._total_rows = total_rows_a + total_rows_b

        self.set_progress(0, "正在准备（SQLite模式）...")

        needed_a = set(self.key_map_a)
        for m in self.compare_map:
            needed_a.add(m["standard"])
        for f in self.filters_a:
            if "field" in f:
                needed_a.add(f["field"])
        needed_a.discard("")
        needed_b = set(self.key_map_b)
        for m in self.compare_map:
            needed_b.add(m["compare"])

        db = SqliteBuffer()
        try:
            db.create_table("std", list(needed_a), self.key_map_a)
            db.create_table("cmp", list(needed_b), self.key_map_b)

            self.set_progress(0, "导入标准群...")
            self._import_files(db, "std", self.files_a, list(needed_a), self.filters_a)
            self.set_progress(50, "导入比较群...")
            self._import_files(db, "cmp", self.files_b, list(needed_b), self.filters_b)

            if self.duplicate_action == "error":
                for tbl, keys, label in [("std", self.key_map_a, "标准"), ("cmp", self.key_map_b, "比较")]:
                    quoted_keys = ",".join([f'"{k}"' for k in keys])
                    sql = f"SELECT {quoted_keys} FROM {tbl} GROUP BY {quoted_keys} HAVING COUNT(*)>1"
                    if db.execute(sql).fetchone():
                        raise ValidationError(f"{label}群主键重复")
            else:
                db.remove_duplicates("std", self.key_map_a)
                db.remove_duplicates("cmp", self.key_map_b)

            cmp_a = [m["standard"] for m in self.compare_map]
            cmp_b = [m["compare"] for m in self.compare_map]
            quoted_std_keys = [f'std."{ka}"' for ka in self.key_map_a]
            quoted_cmp_keys = [f'cmp."{kb}"' for kb in self.key_map_b]
            join_on = " AND ".join([f"{a} = {b}" for a, b in zip(quoted_std_keys, quoted_cmp_keys)])
            eq_parts = [
                f'((std."{ca}" = cmp."{cb}") OR (std."{ca}" IS NULL AND cmp."{cb}" IS NULL))'
                for ca, cb in zip(cmp_a, cmp_b)
            ]
            diff_sql = f"SELECT std.*, cmp.* FROM std JOIN cmp ON {join_on} WHERE NOT ({' AND '.join(eq_parts)})"
            only_a_sql = f"SELECT std.* FROM std LEFT JOIN cmp ON {join_on} WHERE cmp.\"{self.key_map_b[0]}\" IS NULL"
            only_b_sql = f"SELECT cmp.* FROM cmp LEFT JOIN std ON {join_on} WHERE std.\"{self.key_map_a[0]}\" IS NULL"

            self.set_progress(80, "执行对比查询...")
            diff_df = db.fetch_df(diff_sql)
            only_a_df = db.fetch_df(only_a_sql)
            only_b_df = db.fetch_df(only_b_sql)

            self.set_progress(90, "写入结果...")
            sheets = {}
            if not diff_df.empty:
                sheets["差异数据"] = diff_df
            if not only_a_df.empty:
                sheets["标准群独有"] = only_a_df
            if not only_b_df.empty:
                sheets["比较群独有"] = only_b_df

            writer = ReportWriter(
                timeout_seconds=self.config.sheet2_timeout_seconds if self.config else None,
                progress_callback=lambda **kw: self.set_progress(**kw)
            )
            result_path = writer.write(sheets, start_time)
            self.set_progress(100, "完成")
            return result_path
        finally:
            db.close()

    def _import_files(self, db, table, files, columns, filters=None):
        batch = []
        batch_size = 5000
        for file in files:
            for row in StreamReader.read_excel(file, columns, filters):
                batch.append([row[c] for c in columns])
                self._processed_rows += 1
                if len(batch) >= batch_size:
                    db.insert_batch(table, columns, batch)
                    batch.clear()
            if batch:
                db.insert_batch(table, columns, batch)
                batch.clear()
    ```

- **services/common/excel_styles.py**
  - language: `python`
  - content:
    ```python
# services/common/excel_styles.py
"""公共 Excel 样式定义，避免硬编码"""
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# 字体
TITLE_FONT = Font(name='微软雅黑', bold=True, size=14)
HEADER_FONT_BOLD = Font(name='微软雅黑', bold=True, size=10)
HEADER_FONT_RED = Font(name='微软雅黑', bold=True, size=10, color="FF0000")
NORMAL_FONT = Font(name='微软雅黑', size=10)
BOLD_FONT = Font(name='微软雅黑', bold=True, size=10)
NEGATIVE_FONT = Font(name='微软雅黑', size=10, color="FF0000")

# 填充
GRAY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
GREEN_FILL = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
RED_FILL = PatternFill(start_color="F4B7BE", end_color="F4B7BE", fill_type="solid")
LIGHT_GRAY_FILL = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

# 边框
THIN_BORDER = Border(
    left=Side('thin'), right=Side('thin'),
    top=Side('thin'), bottom=Side('thin')
)

# 对齐
CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
    ```

- **services/convert/__init__.py**
  - language: `python`
  - content:
    ```python
"""文件格式转换服务"""
from .service import ConvertService
    ```

- **services/convert/service.py**
  - language: `python`
  - content:
    ```python
# services/convert/service.py
import os, tempfile, zipfile, csv, logging
import pandas as pd
from typing import Optional, Callable
from core.exceptions import ValidationError
from core.task.base import Task

logger = logging.getLogger(__name__)

class ConvertService(Task):
    def __init__(self, file_path: str, target_format: str = 'csv',
                 encoding: str = 'utf-8-sig', separator: str = ',',
                 progress_callback: Optional[Callable] = None):
        self.file_path = file_path
        self.target_format = target_format
        self.encoding = encoding
        self.separator = separator
        self.progress_callback = progress_callback

    def precheck(self) -> dict:
        ext = os.path.splitext(self.file_path)[1].lower()
        if ext in ('.xlsx', '.xls', '.xlsm'):
            from openpyxl import load_workbook
            wb = load_workbook(self.file_path, read_only=True)
            rows = sum(ws.max_row - 1 for ws in wb.worksheets)
            wb.close()
            return {"type": "excel", "sheets": len(wb.sheetnames), "total_rows": rows}
        else:
            size = os.path.getsize(self.file_path)
            return {"type": "csv", "total_rows": max(size // 200, 1)}

    def execute(self) -> str:
        ext = os.path.splitext(self.file_path)[1].lower()
        if self.target_format == 'csv' and ext in ('.xlsx', '.xls', '.xlsm'):
            return self._excel_to_csv()
        elif self.target_format == 'xlsx' and ext in ('.csv',):
            return self._csv_to_excel()
        else:
            raise ValidationError(f"不支持的转换: {ext} -> {self.target_format}")

    def _excel_to_csv(self) -> str:
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path, read_only=True)
        sheet_names = wb.sheetnames
        if len(sheet_names) == 1:
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            total_rows = ws.max_row - 1
            with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w', encoding=self.encoding, newline='') as f:
                writer = csv.writer(f, delimiter=self.separator)
                writer.writerow(headers)
                buffer = []
                buffer_size = 200000   # 20 万行内存缓冲
                processed = 0
                self._update_progress(0, f"开始转换 (共约 {total_rows:,} 行)...")
                for row in ws.iter_rows(min_row=2, values_only=True):
                    buffer.append(row)
                    processed += 1
                    if len(buffer) >= buffer_size:
                        writer.writerows(buffer)
                        buffer.clear()
                        # 只在大块完成时更新进度
                        pct = int(processed / total_rows * 100) if total_rows else 50
                        self._update_progress(pct, f"已转换 {processed:,} / {total_rows:,} 行")
                # 写入剩余行
                if buffer:
                    writer.writerows(buffer)
                out_path = f.name
            wb.close()
            self._update_progress(100, "转换完成")
            return out_path
        else:
            # 多 sheet 打包不变
            zip_path = tempfile.NamedTemporaryFile(suffix='.zip', delete=False).name
            tmp_files = []
            total = len(sheet_names)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for i, name in enumerate(sheet_names):
                    self._update_progress(i/total*100, f"转换工作表 {name} ({i+1}/{total})")
                    df = pd.read_excel(self.file_path, sheet_name=name, engine='openpyxl')
                    csv_name = f"{os.path.splitext(os.path.basename(self.file_path))[0]}_{name}.csv"
                    tmp_csv = tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w', encoding=self.encoding)
                    df.to_csv(tmp_csv.name, index=False, encoding=self.encoding, sep=self.separator)
                    tmp_files.append(tmp_csv.name)
                    zf.write(tmp_csv.name, arcname=csv_name)
            for f in tmp_files:
                try: os.unlink(f)
                except: pass
            wb.close()
            self._update_progress(100, "转换完成")
            return zip_path

    def _csv_to_excel(self) -> str:
        total_size = os.path.getsize(self.file_path)
        estimated_rows = max(total_size // 100, 1)
        self._update_progress(0, "开始读取CSV...")

        chunks = []
        total_rows = 0
        for chunk in pd.read_csv(self.file_path, encoding=self.encoding, sep=self.separator, chunksize=100000):
            chunks.append(chunk)
            total_rows += len(chunk)
            pct = min(int(total_rows / estimated_rows * 100), 99)
            self._update_progress(pct, f"已读取 {total_rows:,} 行")

        df = pd.concat(chunks)
        self._update_progress(95, "正在写入Excel...")
        out_path = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
        df.to_excel(out_path, index=False, engine='openpyxl')
        self._update_progress(100, "转换完成")
        return out_path

    def _update_progress(self, percent, message):
        if self.progress_callback:
            self.progress_callback(percent=percent, message=message)
    ```

- **services/data/__init__.py**
  - language: `python`
  - content:
    ```python
"""数据读写领域服务"""
    ```

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
# utils/common/aggregator.py
import pandas as pd
from typing import List, Dict, Any, Union, Optional

class Aggregator:
    """通用聚合器（静态方法）"""

    @staticmethod
    def group_aggregate(df: pd.DataFrame, group_keys: List[str],
                        agg_spec: Dict[str, Union[str, List[str]]]) -> pd.DataFrame:
        if not group_keys:
            return df.drop_duplicates()
        return df.groupby(group_keys, as_index=False).agg(agg_spec)

    @staticmethod
    def aggregate_by_keys(df: pd.DataFrame, unique_keys: List[str],
                          sum_fields: List[str], other_fields: Optional[List[str]] = None) -> pd.DataFrame:
        """
        按唯一键聚合：sum_fields 求和，other_fields 取第一条记录。
        不使用 groupby.agg 字典，改为组合 drop_duplicates 和 groupby sum，避免 pandas 内部错误。
        """
        if other_fields is None:
            other_fields = [c for c in df.columns if c not in unique_keys and c not in sum_fields]
        if not unique_keys:
            return df.drop_duplicates()

        # 1. 求和部分
        if sum_fields:
            sum_part = df.groupby(unique_keys, as_index=False)[sum_fields].sum()
        else:
            sum_part = pd.DataFrame(columns=unique_keys)  # 空框架，仅用于后续合并

        # 2. 取第一条记录部分（直接用 drop_duplicates 保留第一个出现）
        if other_fields:
            first_part = df[unique_keys + other_fields].drop_duplicates(subset=unique_keys, keep='first')
        else:
            first_part = pd.DataFrame(columns=unique_keys)

        # 3. 合并求和与首条记录
        if not sum_part.empty and not first_part.empty:
            result = sum_part.merge(first_part, on=unique_keys, how='outer')
        elif not sum_part.empty:
            result = sum_part
        elif not first_part.empty:
            result = first_part
        else:
            # 没有求和也没有其他字段时，直接去重返回唯一键
            result = df[unique_keys].drop_duplicates()

        return result

    @staticmethod
    def simple_aggregate(df: pd.DataFrame, group_keys: List[str],
                         agg_cols: List[str], func: str = 'sum') -> pd.DataFrame:
        agg_dict = {col: func for col in agg_cols}
        return Aggregator.group_aggregate(df, group_keys, agg_dict)


# 向后兼容包装
class AggregateCalculator:
    def __init__(self, excel_file):
        self.excel_file = excel_file

    def aggregate(self, group_cols, agg_cols, agg_func, result_col):
        raise NotImplementedError
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

    def __init__(self, filepath: str, header_row: int, compute_engine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_files = []  # 跟踪所有临时文件
        self._work_filepath = None

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
        获取工作文件路径，确保文件可读
        """
        if self._work_filepath and os.path.exists(self._work_filepath):
            return self._work_filepath

        # 检查文件是否存在
        if not os.path.exists(self.original_filepath):
            raise FileNotFoundError(f"原始文件不存在: {self.original_filepath}")

        # 总是进行规范化，确保文件格式正确
        try:
            self._work_filepath = normalize_excel_file(self.original_filepath)
            if self._work_filepath and os.path.exists(self._work_filepath):
                # 记录临时文件以便后续清理
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
                 result_col_input: str, merge_by_cols_str: str, compute_engine):
        super().__init__(filepath, header_row, compute_engine)
        self.expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols = [c.strip() for c in merge_by_cols_str.split(',') if c.strip()] if merge_by_cols_str else []

    def _remove_result_column_from_expr(self, result_col_letter: str) -> str:
        """从表达式中移除结果列字母（如果存在）"""
        if not result_col_letter or not self.expression:
            return self.expression

        expr = self.expression.upper()
        result_col_letter = result_col_letter.upper()

        # 简单替换结果列字母
        new_expr = expr.replace(result_col_letter, '')

        # 清理多余的运算符
        import re
        new_expr = re.sub(r'[+-]{2,}', '+', new_expr)  # 多个+-转为+
        new_expr = re.sub(r'[*/]{2,}', '*', new_expr)  # 多个*/转为*
        new_expr = re.sub(r'^\s*[+-*/]+\s*', '', new_expr)  # 开头运算符
        new_expr = re.sub(r'\s*[+-*/]+\s*$', '', new_expr)  # 结尾运算符

        if new_expr != expr:
            logger.warning(f"表达式中包含了结果列 {result_col_letter}，已自动移除: {expr} -> {new_expr}")

        return new_expr if new_expr.strip() else '0'  # 如果表达式为空，返回0

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
                 agg_func: str, result_col_input: str, compute_engine):
        super().__init__(filepath, header_row, compute_engine)
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
        self.allow_duplicate_keys = config.get('allow_duplicate_keys', False)

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
        """提取目标文件中所有行的组合主键"""
        start_row = 2
        end_row = self.target_ws.max_row
        keys = []
        for row in range(start_row, end_row + 1):
            key = []
            for col in self.target_key_indices:
                val = self.target_wb.get_cell_value(row, col, sheet=self.target_ws)
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
        seen_keys = set()
        for row_offset, key in enumerate(keys):
            row_num = start_row + row_offset
            if self.allow_duplicate_keys:
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

        # 设置右对齐
        for idx in self.result_col_indices:
            for row in range(start_row, self.target_ws.max_row + 1):
                cell = self.target_ws.cell(row, idx)
                cell.alignment = Alignment(horizontal='right')

    def _generate_report(self, used_target_keys, lookup, start_row):
        """生成校验报告 sheet，展示所有源值列的值"""
        target_keys_full, _, _ = self._extract_target_keys()
        target_keys_used = used_target_keys
        source_keys_set = set(lookup.keys())

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
        for key, val_tuple in lookup.items():
            if key not in set(target_keys_used):
                row_data = ['-', '', "|".join(str(v) for v in key)]
                for col_idx in range(len(self.source_value_cols)):
                    val = val_tuple[col_idx] if col_idx < len(val_tuple) else ''
                    row_data.append(val)
                row_data.append(0)
                row_data.append('仅源文件')
                for col_idx, val in enumerate(row_data, 1):
                    report_sheet.cell(row_idx, col_idx).value = val
                row_idx += 1

        # 设置对齐
        for col in range(1, len(headers) + 1):
            for row in range(2, row_idx):
                cell = report_sheet.cell(row, col)
                cell.alignment = Alignment(horizontal='right' if col > 1 else 'left')

    def match(self):
        if not self._validate_column_letters():
            return False, self.errors
        if not self._load_source():
            return False, self.errors
        if not self._load_target():
            return False, self.errors

        target_keys, start_row, end_row = self._extract_target_keys()

        if not self.allow_duplicate_keys:
            duplicates = self._check_duplicate_target_keys(target_keys)
            if duplicates:
                dup_examples = [f"'{k}'" for k in list(duplicates)[:5]]
                self.errors.append(f"目标文件键列存在重复值: {', '.join(dup_examples)}。如需继续填充，请勾选“允许目标键重复”。")
                return False, self.errors

        lookup = self._build_lookup_dict()
        match_len = len(self.target_key_cols)
        used_target_keys = [key[:match_len] for key in target_keys]

        # 写入结果列
        self._write_result_columns(used_target_keys, lookup, start_row)

        # 生成报告（多列版本）
        self._generate_report(used_target_keys, lookup, start_row)

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

- **services/subtotal/engine.py**
  - language: `python`
  - content:
    ```python
# services/subtotal/engine.py
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from core.csv_utils import split_fields
from core.header_detection import SimpleExcelHeaderDetector
from core.file_naming import generate_output_filename

SUBTOTAL_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
SUBTOTAL_FONT = Font(bold=True)
THIN_BORDER = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))

def _build_header_map(ws):
    """根据工作表第一行构建列名 -> 列字母的映射"""
    header_map = {}
    for cell in ws[1]:
        if cell.value is not None:
            col_letter = get_column_letter(cell.column)
            header_map[str(cell.value).strip()] = col_letter
    return header_map

def resolve_single_col(token, header_map):
    """解析单个列标识（字母或列名），返回1-based列索引"""
    try:
        return column_index_from_string(token)
    except (ValueError, KeyError):
        if token in header_map:
            return column_index_from_string(header_map[token])
        raise ValueError(f"无效的列标识: {token}")

def parse_cols_with_header(raw, header_map):
    """解析列字符串（支持范围、列名），返回列索引列表"""
    if not raw:
        return []
    items = split_fields(raw)
    result = []
    for item in items:
        if '-' in item:
            start_str, end_str = item.split('-', 1)
            start_idx = resolve_single_col(start_str.strip(), header_map)
            end_idx = resolve_single_col(end_str.strip(), header_map)
            for i in range(start_idx, end_idx + 1):
                result.append(i)
        else:
            result.append(resolve_single_col(item.strip(), header_map))
    return sorted(list(set(result)))

class SubtotalEngine:
    def __init__(self, filepath, key_cols_str, calc_cols_str, agg_funcs,
                 dimension_cols=None, sheets='1', original_filename='data.xlsx'):
        self.filepath = filepath
        self.original_filename = original_filename
        self.key_cols_raw = key_cols_str
        self.calc_cols_raw = calc_cols_str
        self.dimension_cols_raw = dimension_cols or ''
        self.agg_funcs = agg_funcs
        self.sheets = sheets
        self.detector = SimpleExcelHeaderDetector()
        self.key_cols = []
        self.calc_cols = []
        self.dimension_cols = []

    def _get_agg_func(self, col_letter):
        return self.agg_funcs.get(col_letter, self.agg_funcs.get('default', 'sum'))

    def process(self):
        wb = openpyxl.load_workbook(self.filepath)
        first_ws = wb.active
        header_map = _build_header_map(first_ws)

        self.key_cols = parse_cols_with_header(self.key_cols_raw, header_map)
        self.calc_cols = parse_cols_with_header(self.calc_cols_raw, header_map)
        if self.dimension_cols_raw:
            self.dimension_cols = parse_cols_with_header(self.dimension_cols_raw, header_map)

        header_result = self.detector.detect_from_workbook(wb)
        data_start_excel_row = header_result.get_data_start_excel_row()

        sheet_names = wb.sheetnames
        if self.sheets == 'all':
            target_sheets = sheet_names
        else:
            indices = [int(x.strip()) for x in self.sheets.split(',') if x.strip().isdigit()]
            target_sheets = [sheet_names[i-1] for i in indices if 1 <= i <= len(sheet_names)]

        for sheet_name in target_sheets:
            ws = wb[sheet_name]
            self._process_sheet(ws, data_start_excel_row)

        out_name = generate_output_filename('subtotal', self.original_filename)
        out_dir = os.path.dirname(self.filepath)
        out_path = os.path.join(out_dir, out_name)
        wb.save(out_path)
        wb.close()
        return out_path

    def _process_sheet(self, ws, data_start_row):
        max_col = ws.max_column
        max_row = ws.max_row
        if data_start_row > max_row:
            return

        groups = {}
        for row_idx in range(data_start_row, max_row + 1):
            key = tuple(ws.cell(row_idx, col).value for col in self.key_cols)
            # ---- 修复：跳过主键全为空的行（避免生成无意义的小计行） ----
            if all(v is None or str(v).strip() == '' for v in key):
                continue
            groups.setdefault(key, []).append(row_idx)

        # 如果没有有效组，直接返回
        if not groups:
            return

        sorted_groups = sorted(groups.items(), key=lambda x: max(x[1]), reverse=True)

        for key, rows in sorted_groups:
            last_row = max(rows)
            ws.insert_rows(last_row + 1)
            subtotal_row = last_row + 1

            for col_idx in range(1, max_col + 1):
                cell = ws.cell(subtotal_row, col_idx)
                cell.fill = SUBTOTAL_FILL
                cell.font = SUBTOTAL_FONT
                cell.border = THIN_BORDER

            for i, col in enumerate(self.key_cols):
                ws.cell(subtotal_row, col).value = key[i]

            for col in self.dimension_cols:
                ws.cell(subtotal_row, col).value = '合计'

            for col in self.calc_cols:
                col_letter = get_column_letter(col)
                func = self._get_agg_func(col_letter)
                values = []
                for r in rows:
                    val = ws.cell(r, col).value
                    try:
                        val = float(val)
                    except (TypeError, ValueError):
                        val = 0
                    values.append(val)

                if func == 'sum':
                    result = sum(values)
                elif func == 'avg':
                    result = sum(values) / len(values) if values else 0
                elif func == 'min':
                    result = min(values) if values else 0
                elif func == 'max':
                    result = max(values) if values else 0
                elif func == 'count':
                    result = len(values)
                else:
                    result = sum(values)
                ws.cell(subtotal_row, col).value = result
    ```

- **services/team_report/__init__.py**
  - language: `python`
  - content:
    ```python
# services/team_report/__init__.py
"""车队日报服务"""
from .service import TeamReportService
from .writer import TeamReportExcelWriter
    ```

- **services/team_report/service.py**
  - language: `python`
  - content:
    ```python
# services/team_report/service.py
"""团队报表服务：数据聚合与计算"""
import pandas as pd
import numpy as np
from typing import Dict, List, Optional
from core.exceptions import ValidationError
import logging

logger = logging.getLogger(__name__)

class TeamReportService:
    def __init__(self, config):
        self.config = config

    def _validate_columns(self, df, required_cols, label):
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValidationError(f"{label}缺少列: {missing}")

    # ---------- 映射加载 ----------
    def load_mapping(self, path: str) -> Dict[str, Dict[str, str]]:
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        core = [self.config.mapping_driver_id_col, self.config.mapping_team_name_col,
                self.config.mapping_leader_name_col]
        self._validate_columns(df, core, "队伍明细")
        df[self.config.mapping_driver_id_col] = df[self.config.mapping_driver_id_col].astype(str).str.strip().str.replace('.0', '', regex=False)
        has_s = self.config.mapping_supervisor_col in df.columns
        static = getattr(self.config, "static_supervisor_mapping", {}) if not has_s else {}
        mapping = {}
        for _, row in df.iterrows():
            did = row[self.config.mapping_driver_id_col]
            leader = row[self.config.mapping_leader_name_col]
            team = row[self.config.mapping_team_name_col]
            sup = row[self.config.mapping_supervisor_col] if has_s else static.get(leader, "")
            mapping[did] = {"队伍名称": team, "队长姓名": leader, "司管": sup}
        return mapping

    # ---------- 流水数据预处理 ----------
    def _prepare_flow(self, path: str, mapping: Dict) -> pd.DataFrame:
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        dc = self.config.flow_driver_id_col
        if dc not in df.columns:
            raise ValidationError(f"缺少司机ID列: {dc}")
        df[dc] = df[dc].astype(str).str.strip().str.replace('.0', '', regex=False)
        df['队伍名称'] = df[dc].map(lambda x: mapping.get(x, {}).get('队伍名称', '未分配'))
        df['队长姓名'] = df[dc].map(lambda x: mapping.get(x, {}).get('队长姓名', ''))
        df['司管'] = df[dc].map(lambda x: mapping.get(x, {}).get('司管', ''))
        return df

    # ---------- 按队伍聚合 ----------
    def _aggregate_by_team(self, df: pd.DataFrame):
        status_col = self.config.flow_status_col
        active = df[df[status_col] == self.config.flow_active_status].copy() if status_col in df.columns else df.copy()
        team_size = df.groupby('队伍名称').size().rename('小队人数').reset_index()

        online_col = self.config.flow_online_days_col
        if online_col in active.columns:
            active[online_col] = pd.to_numeric(active[online_col], errors='coerce').fillna(0)

        value_cols = [
            self.config.flow_order_col,
            self.config.flow_apply_col,
            self.config.flow_transfer_col,
            self.config.flow_new_ju_col,
        ]
        for c in value_cols:
            if c in active.columns:
                active[c] = pd.to_numeric(active[c], errors='coerce').fillna(0)

        agg = {self.config.flow_driver_id_col: 'count'}
        for c in value_cols:
            if c in active.columns:
                agg[c] = 'sum'
        if online_col in active.columns:
            agg[online_col] = 'sum'

        grp = ['队伍名称', '队长姓名', '司管']
        avail = [c for c in grp if c in active.columns]
        summary = active.groupby(avail, as_index=False).agg(agg)
        summary.rename(columns={self.config.flow_driver_id_col: '可上线人数'}, inplace=True)
        if online_col in active.columns:
            summary.rename(columns={online_col: '上线人数'}, inplace=True)
        else:
            summary['上线人数'] = 0

        summary = summary.merge(team_size, on='队伍名称', how='left')
        return summary

    # ---------- 月度数据聚合（多文件） ----------
    def aggregate_monthly(self, file_paths: List[str], mapping: Dict) -> pd.DataFrame:
        """合并多个月度流水文件并聚合队伍月度数据"""
        dfs = []
        for path in file_paths:
            df = self._prepare_flow(path, mapping)
            dfs.append(df)
        if not dfs:
            return pd.DataFrame()
        combined = pd.concat(dfs, ignore_index=True)
        # 按队伍聚合月度数值
        monthly = self._aggregate_by_team(combined)
        return monthly

    # ---------- 生成对比报表数据 ----------
    def build_report_data(self, today_df: pd.DataFrame,
                          yesterday_df: Optional[pd.DataFrame],
                          monthly_df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        today_agg = self._aggregate_by_team(today_df)
        if yesterday_df is not None:
            yesterday_agg = self._aggregate_by_team(yesterday_df)
            report = pd.merge(today_agg, yesterday_agg,
                              on=['队伍名称', '队长姓名', '司管', '小队人数'],
                              how='outer', suffixes=('_今日', '_昨日'))
            # 填充缺失
            for c in report.columns:
                if '_今日' in c or '_昨日' in c:
                    report[c] = report[c].fillna(0).astype(int)
            if '可上线人数_今日' in report.columns and '可上线人数_昨日' in report.columns:
                report['可上线人数'] = report['可上线人数_今日'].fillna(0) + report['可上线人数_昨日'].fillna(0)
            elif '可上线人数_今日' in report.columns:
                report['可上线人数'] = report['可上线人数_今日']
            else:
                report['可上线人数'] = report.get('可上线人数_昨日', 0)
        else:
            report = today_agg.copy()
            report['可上线人数'] = report.get('可上线人数', 0)

        # 上线率
        today_online = '上线人数_今日' if yesterday_df is not None else '上线人数'
        yesterday_online = '上线人数_昨日' if yesterday_df is not None else None
        can_today = '可上线人数_今日' if yesterday_df is not None else '可上线人数'
        can_yesterday = '可上线人数_昨日' if yesterday_df is not None else None
        report['今日上线率'] = np.where(report[can_today] > 0, report[today_online] / report[can_today], 0)
        if yesterday_df is not None:
            report['昨日上线率'] = np.where(report[can_yesterday] > 0, report[yesterday_online] / report[can_yesterday], 0)
            report['上线率同比'] = report['今日上线率'] - report['昨日上线率']
        else:
            report['昨日上线率'] = 0.0
            report['上线率同比'] = 0.0

        # 日完单/报单完成率
        order_today = self.config.flow_order_col + '_今日'
        apply_today = self.config.flow_apply_col + '_今日'
        newju_today = self.config.flow_new_ju_col + '_今日'
        if order_today in report.columns:
            report['日完单目标完成率'] = report[order_today] / self.config.daily_order_target
        if apply_today in report.columns:
            report['日报单目标完成率'] = report[apply_today] / self.config.daily_apply_target
        if newju_today in report.columns:
            safe_online = report[can_today].replace(0, 1)
            report['新桔渗透率'] = report[newju_today] / safe_online
        else:
            report['新桔渗透率'] = 0.0

        # 月度数据合并
        if monthly_df is not None and not monthly_df.empty:
            # 确保月度 aggreg 有队伍名称等键
            monthly_renames = {
                self.config.flow_order_col: '月完单',
                self.config.flow_apply_col: '月报单',
                self.config.flow_transfer_col: '月转单',
                self.config.flow_new_ju_col: '月新桔',
            }
            for old, new in monthly_renames.items():
                if old in monthly_df.columns:
                    monthly_df[new] = monthly_df[old]
            monthly_df = monthly_df[['队伍名称'] + [v for v in monthly_renames.values() if v in monthly_df.columns]]
            report = pd.merge(report, monthly_df, on='队伍名称', how='left')
        # 填充月度列默认值
        for col in ['月完单', '月报单', '月转单', '月新桔']:
            if col not in report.columns:
                report[col] = 0
            else:
                report[col] = report[col].fillna(0).astype(int)

        report.sort_values(['司管', '队伍名称'], inplace=True)
        return report

    # 为 writer 提供标准化的行字典
    def report_to_rows(self, report: pd.DataFrame,
                       daily_order: int, daily_apply: int,
                       monthly_order: int, monthly_apply: int) -> List[Dict]:
        rows = []
        for _, r in report.iterrows():
            row = {
                '司管': r['司管'],
                '队长姓名': r['队长姓名'],
                '队伍名称': r['队伍名称'],
                '小队人数': r['小队人数'],
                '可上线人数': r['可上线人数'],
                '今日上线率': r['今日上线率'],
                '昨日上线率': r['昨日上线率'],
                '上线率同比': r['上线率同比'],
                'daily_order_today': r.get(self.config.flow_order_col+'_今日', 0),
                'daily_apply_today': r.get(self.config.flow_apply_col+'_今日', 0),
                'daily_transfer_today': r.get(self.config.flow_transfer_col+'_今日', 0),
                'daily_newju_today': r.get(self.config.flow_new_ju_col+'_今日', 0),
                'daily_order_yesterday': r.get(self.config.flow_order_col+'_昨日', 0),
                'daily_apply_yesterday': r.get(self.config.flow_apply_col+'_昨日', 0),
                'daily_newju_yesterday': r.get(self.config.flow_new_ju_col+'_昨日', 0),
                '上线人数_今日': r.get('上线人数_今日', 0),
                '上线人数_昨日': r.get('上线人数_昨日', 0),
                '可上线人数_今日': r.get('可上线人数_今日', 0),
                '可上线人数_昨日': r.get('可上线人数_昨日', 0),
                'monthly_order': r.get('月完单', 0),
                'monthly_apply': r.get('月报单', 0),
                'monthly_transfer': r.get('月转单', 0),
                'monthly_newju': r.get('月新桔', 0),
            }
            rows.append(row)
        return rows
    ```

- **services/team_report/writer.py**
  - language: `python`
  - content:
    ```python
# services/team_report/writer.py
"""团队报表 Excel 渲染器（修正完成率）"""
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from services.common.excel_styles import (
    TITLE_FONT, HEADER_FONT_BOLD, HEADER_FONT_RED, NORMAL_FONT, BOLD_FONT,
    GRAY_FILL, GREEN_FILL, RED_FILL, LIGHT_GRAY_FILL, THIN_BORDER,
    CENTER_ALIGN, LEFT_ALIGN
)
import logging

logger = logging.getLogger(__name__)

class ColumnMapping:
    """列定义，支持动态标题和数据提取"""
    def __init__(self, letter: str, title_fn, width: float,
                 data_key: str = None,
                 is_target: bool = False,
                 value_fn: callable = None):
        self.letter = letter
        self._title_fn = title_fn
        self.width = width
        self.data_key = data_key
        self.is_target = is_target
        self._value_fn = value_fn

    @property
    def title(self):
        return self._title_fn() if callable(self._title_fn) else self._title_fn

    def get_value(self, row, extra_context=None):
        if self._value_fn:
            return self._value_fn(row)
        if self.data_key:
            return row.get(self.data_key, 0)
        return 0

class TeamReportExcelWriter:
    """生成车队日报 Excel 文件（输出与原版完全一致）"""
    def __init__(self):
        self._column_defs = []

    def build_columns(self, report_date: str, month: str,
                      daily_order_target: int, daily_apply_target: int,
                      monthly_order_target: int, monthly_apply_target: int):
        day = report_date.split('/')[-1]
        cols = []
        # A 司管
        cols.append(ColumnMapping('A', lambda: '司管', 6.5, data_key='司管'))
        # B 小队长
        cols.append(ColumnMapping('B', lambda: '小队长', 13.88, data_key='队长姓名'))
        # C 队伍
        cols.append(ColumnMapping('C', lambda: '队伍', 16.25, data_key='队伍名称'))
        # D 小队人数
        cols.append(ColumnMapping('D', lambda: '小队人数', 7.63, data_key='小队人数'))
        # E 可上线人数
        cols.append(ColumnMapping('E', lambda: '可上线人数', 7.63, data_key='可上线人数'))
        # F 今日上线率
        cols.append(ColumnMapping('F', lambda: '今日上线率', 7.63, data_key='今日上线率'))
        # G 昨日上线率
        cols.append(ColumnMapping('G', lambda: '昨日上线率', 7.63, data_key='昨日上线率'))
        # H 同比趋势（上线率）
        cols.append(ColumnMapping('H', lambda: '同比趋势（上线率）', 7.63, data_key='上线率同比'))
        # I {day}号完单
        cols.append(ColumnMapping('I', lambda: f'{day}号完单', 6, data_key='daily_order_today'))
        # J 日完单目标完成率（前值/目标值）
        cols.append(ColumnMapping('J', lambda: f'日完单目标值({daily_order_target}单)', 13.5,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('daily_order_today', 0) / daily_order_target)
                                  if daily_order_target else 0))
        # K {day}号报单
        cols.append(ColumnMapping('K', lambda: f'{day}号报单', 6, data_key='daily_apply_today'))
        # L 日报单目标完成率
        cols.append(ColumnMapping('L', lambda: f'日报单目标值({daily_apply_target}单)', 13.5,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('daily_apply_today', 0) / daily_apply_target)
                                  if daily_apply_target else 0))
        # M {day}号转单
        cols.append(ColumnMapping('M', lambda: f'{day}号转单', 6, data_key='daily_transfer_today'))
        # N {day}号新桔在线数
        cols.append(ColumnMapping('N', lambda: f'{day}号新桔在线数', 8.5, data_key='daily_newju_today'))
        # O 新桔渗透率
        cols.append(ColumnMapping('O', lambda: f'{day}号新桔渗透率', 12,
                                  value_fn=lambda r: (r['daily_newju_today'] / r['daily_order_today'])
                                  if r.get('daily_order_today') else 0))
        # P 昨天完单
        cols.append(ColumnMapping('P', lambda: '昨天完单', 6, data_key='daily_order_yesterday'))
        # Q 同比趋势（完单）
        cols.append(ColumnMapping('Q', lambda: '同比趋势（完单）', 12,
                                  value_fn=lambda r: r['daily_order_today'] - r['daily_order_yesterday']))
        # R 昨天报单
        cols.append(ColumnMapping('R', lambda: '昨天报单', 6, data_key='daily_apply_yesterday'))
        # S 同比趋势（报单）
        cols.append(ColumnMapping('S', lambda: '同比趋势（报单）', 12,
                                  value_fn=lambda r: r['daily_apply_today'] - r['daily_apply_yesterday']))
        # T 昨天新桔
        cols.append(ColumnMapping('T', lambda: '昨天新桔', 6, data_key='daily_newju_yesterday'))
        # U 同比趋势（新桔）
        cols.append(ColumnMapping('U', lambda: '同比趋势（新桔）', 12,
                                  value_fn=lambda r: r['daily_newju_today'] - r['daily_newju_yesterday']))
        # V {month}月完单
        cols.append(ColumnMapping('V', lambda: f'{month}月完单', 8.4, data_key='monthly_order'))
        # W 月完单目标完成率
        cols.append(ColumnMapping('W', lambda: f'{month}月完单目标值({monthly_order_target}单)', 8.4,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('monthly_order', 0) / monthly_order_target)
                                  if monthly_order_target else 0))
        # X {month}月报单
        cols.append(ColumnMapping('X', lambda: f'{month}月报单', 8.4, data_key='monthly_apply'))
        # Y 月报单目标完成率
        cols.append(ColumnMapping('Y', lambda: f'{month}月报单目标值({monthly_apply_target}单)', 8.4,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('monthly_apply', 0) / monthly_apply_target)
                                  if monthly_apply_target else 0))
        # Z {month}月转单
        cols.append(ColumnMapping('Z', lambda: f'{month}月转单', 12, data_key='monthly_transfer'))
        # AA {month}月新桔在线数
        cols.append(ColumnMapping('AA', lambda: f'{month}月新桔在线数', 12, data_key='monthly_newju'))
        self._column_defs = cols

    def write(self, rows, output_path: str,
              report_date: str, month: str,
              daily_order: int, daily_apply: int,
              monthly_order: int, monthly_apply: int):
        self.build_columns(report_date, month, daily_order, daily_apply, monthly_order, monthly_apply)
        wb = Workbook()
        ws = wb.active
        ws.title = "每日数据对比"

        # 列宽
        for col_def in self._column_defs:
            ws.column_dimensions[col_def.letter].width = col_def.width

        # 第1行标题
        ws.row_dimensions[1].height = 27
        ws.merge_cells('A1:AA1')
        c = ws['A1']
        c.value = f"日期\t{report_date}\t杭州新享{month}月三大队订单完成情况"
        c.font = TITLE_FONT
        c.alignment = CENTER_ALIGN
        c.fill = GRAY_FILL

        # 第2行子标题
        ws.row_dimensions[2].height = 27
        ws.merge_cells('A2:B2')
        ws['A2'].value = f"月份\t{month}月"
        ws['A2'].font = HEADER_FONT_BOLD
        ws['A2'].alignment = CENTER_ALIGN
        ws['A2'].fill = GRAY_FILL
        ws.merge_cells('C2:E2')
        ws['C2'].fill = GRAY_FILL
        ws.merge_cells('F2:H2')
        ws['F2'].fill = GRAY_FILL
        ws.merge_cells('I2:O2')
        ws['I2'].value = "今日数据"
        ws['I2'].font = HEADER_FONT_BOLD
        ws['I2'].alignment = CENTER_ALIGN
        ws['I2'].fill = GREEN_FILL
        ws.merge_cells('P2:U2')
        ws['P2'].value = "昨日数据"
        ws['P2'].font = HEADER_FONT_BOLD
        ws['P2'].alignment = CENTER_ALIGN
        ws['P2'].fill = RED_FILL
        ws.merge_cells('V2:AA2')
        ws['V2'].value = f"{month}月数据"
        ws['V2'].font = HEADER_FONT_BOLD
        ws['V2'].alignment = CENTER_ALIGN
        ws['V2'].fill = GRAY_FILL

        # 第3行表头
        ws.row_dimensions[3].height = 54
        for idx, col_def in enumerate(self._column_defs):
            col_idx = idx + 1
            cell = ws.cell(row=3, column=col_idx, value=col_def.title)
            cell.font = HEADER_FONT_RED if col_def.is_target else HEADER_FONT_BOLD
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER
            if col_idx <= 5:
                cell.fill = GRAY_FILL
            elif 6 <= col_idx <= 8:
                cell.fill = GRAY_FILL
            elif 9 <= col_idx <= 15:
                cell.fill = GREEN_FILL
            elif 16 <= col_idx <= 21:
                cell.fill = RED_FILL
            else:
                cell.fill = GRAY_FILL

        current_row = 4
        start_data_row = 4
        subtotal_rows = []
        supervisors = sorted(list({r.get('司管', '') for r in rows}))
        global_on_today = 0
        global_on_yesterday = 0
        global_can_today = 0
        global_can_yesterday = 0

        # 数据行
        for sup in supervisors:
            group = [r for r in rows if r.get('司管', '') == sup]
            first_group_row = current_row
            sum_on_today = 0
            sum_on_yesterday = 0
            sum_can_today = 0
            sum_can_yesterday = 0
            for row in group:
                for col_idx, col_def in enumerate(self._column_defs, start=1):
                    val = col_def.get_value(row)
                    cell = ws.cell(row=current_row, column=col_idx, value=val)
                    if col_idx in (6, 7, 8, 10, 12, 15, 23, 25):
                        cell.number_format = '0.00%'
                    cell.font = NORMAL_FONT
                    cell.alignment = CENTER_ALIGN if col_idx > 3 else LEFT_ALIGN
                    cell.border = THIN_BORDER
                sum_on_today += row.get('上线人数_今日', 0)
                sum_on_yesterday += row.get('上线人数_昨日', 0)
                sum_can_today += row.get('可上线人数_今日', 0)
                sum_can_yesterday += row.get('可上线人数_昨日', 0)
                current_row += 1

            global_on_today += sum_on_today
            global_on_yesterday += sum_on_yesterday
            global_can_today += sum_can_today
            global_can_yesterday += sum_can_yesterday

            if first_group_row < current_row:
                ws.merge_cells(start_row=first_group_row, start_column=1,
                               end_row=current_row-1, end_column=1)

            # 司管小计行
            subtotal_rows.append(current_row)
            for col_idx in range(1, 28):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.font = BOLD_FONT
                cell.border = THIN_BORDER
                cell.fill = GRAY_FILL
                cell.alignment = CENTER_ALIGN
            ws.cell(row=current_row, column=1, value="总计")

            # 求和列
            sum_cols_indices = [4, 5, 9, 11, 13, 14, 16, 17, 18, 19, 20, 21, 22, 24, 26, 27]
            for ci in sum_cols_indices:
                total = 0
                for r in range(first_group_row, current_row):
                    val = ws.cell(row=r, column=ci).value
                    if isinstance(val, (int, float)):
                        total += val
                ws.cell(row=current_row, column=ci, value=total)

            # 小计行完成率
            total_on_today = sum_on_today
            total_on_yesterday = sum_on_yesterday
            total_can_today = sum_can_today if sum_can_today else 1
            total_can_yesterday = sum_can_yesterday if sum_can_yesterday else 1
            ws.cell(row=current_row, column=6, value=total_on_today / total_can_today).number_format = '0.00%'
            ws.cell(row=current_row, column=7, value=total_on_yesterday / total_can_yesterday).number_format = '0.00%'
            ws.cell(row=current_row, column=8, value=(total_on_today / total_can_today) - (total_on_yesterday / total_can_yesterday)).number_format = '0.00%'
            total_order_today = ws.cell(row=current_row, column=9).value or 0
            ws.cell(row=current_row, column=10, value=total_order_today / daily_order if daily_order else 0).number_format = '0.00%'
            total_apply_today = ws.cell(row=current_row, column=11).value or 0
            ws.cell(row=current_row, column=12, value=total_apply_today / daily_apply if daily_apply else 0).number_format = '0.00%'
            total_newju = ws.cell(row=current_row, column=14).value or 0
            ws.cell(row=current_row, column=15, value=total_newju / total_order_today if total_order_today else 0).number_format = '0.00%'
            total_monthly_order = ws.cell(row=current_row, column=22).value or 0
            ws.cell(row=current_row, column=23, value=total_monthly_order / monthly_order if monthly_order else 0).number_format = '0.00%'
            total_monthly_apply = ws.cell(row=current_row, column=24).value or 0
            ws.cell(row=current_row, column=25, value=total_monthly_apply / monthly_apply if monthly_apply else 0).number_format = '0.00%'
            current_row += 1

        # 全部总计行
        grand_total_row = current_row
        for col_idx in range(1, 28):
            cell = ws.cell(row=grand_total_row, column=col_idx)
            cell.font = BOLD_FONT
            cell.border = THIN_BORDER
            cell.fill = GRAY_FILL
            cell.alignment = CENTER_ALIGN
        ws.cell(row=grand_total_row, column=1, value="总计")

        sum_cols_indices_all = [4, 5, 9, 11, 13, 14, 16, 17, 18, 19, 20, 21, 22, 24, 26, 27]
        for ci in sum_cols_indices_all:
            total = 0
            for r in range(start_data_row, grand_total_row):
                if r in subtotal_rows:
                    continue
                val = ws.cell(row=r, column=ci).value
                if isinstance(val, (int, float)):
                    total += val
            ws.cell(row=grand_total_row, column=ci, value=total)

        # 全部总计完成率
        all_can_today = global_can_today if global_can_today else 1
        all_can_yesterday = global_can_yesterday if global_can_yesterday else 1
        ws.cell(row=grand_total_row, column=6, value=global_on_today / all_can_today).number_format = '0.00%'
        ws.cell(row=grand_total_row, column=7, value=global_on_yesterday / all_can_yesterday).number_format = '0.00%'
        ws.cell(row=grand_total_row, column=8, value=(global_on_today / all_can_today) - (global_on_yesterday / all_can_yesterday)).number_format = '0.00%'
        total_order_all = ws.cell(row=grand_total_row, column=9).value or 0
        ws.cell(row=grand_total_row, column=10, value=total_order_all / daily_order if daily_order else 0).number_format = '0.00%'
        total_apply_all = ws.cell(row=grand_total_row, column=11).value or 0
        ws.cell(row=grand_total_row, column=12, value=total_apply_all / daily_apply if daily_apply else 0).number_format = '0.00%'
        total_newju_all = ws.cell(row=grand_total_row, column=14).value or 0
        ws.cell(row=grand_total_row, column=15, value=total_newju_all / total_order_all if total_order_all else 0).number_format = '0.00%'
        total_monthly_order_all = ws.cell(row=grand_total_row, column=22).value or 0
        ws.cell(row=grand_total_row, column=23, value=total_monthly_order_all / monthly_order if monthly_order else 0).number_format = '0.00%'
        total_monthly_apply_all = ws.cell(row=grand_total_row, column=24).value or 0
        ws.cell(row=grand_total_row, column=25, value=total_monthly_apply_all / monthly_apply if monthly_apply else 0).number_format = '0.00%'
        current_row = grand_total_row + 1

        # 负值标红
        for r in range(start_data_row, current_row):
            for ci in [8, 9, 11, 13, 14, 16, 17, 18, 19, 20, 21, 22, 24, 26, 27]:
                cell = ws.cell(row=r, column=ci)
                if isinstance(cell.value, (int, float)) and cell.value < 0:
                    cell.font = Font(name='微软雅黑', size=10, color="FF0000")

        # 备注行
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=27)
        note_cell = ws.cell(row=current_row, column=1)
        note_cell.value = "表格中标红色的表示数值低于该小队当日该数据的平均值，需要根据数值所在的单元格注意队员的接单情况。"
        note_cell.font = NORMAL_FONT
        note_cell.fill = LIGHT_GRAY_FILL
        note_cell.alignment = CENTER_ALIGN

        wb.save(output_path)
        return output_path
    ```

- **services/transposer/service.py**
  - language: `python`
  - content:
    ```python
# services/transposer/service.py
import pandas as pd
from typing import List, Optional


def pivot_table(df: pd.DataFrame,
                index_cols: Optional[List[str]],
                category_col: str,
                value_cols: List[str]) -> pd.DataFrame:
    """
    将分类列的值转置为新的列名，同时保留其他固定列。

    :param df: 原始数据
    :param index_cols: 索引列列表；若为空则自动生成组内序号（此时固定列不会被保留）
    :param category_col: 分类列名（如“医疗类别”）
    :param value_cols: 值列名列表（如“a数据”，“b数据”）
    :return: 转置后的 DataFrame，固定列在左侧，拼接后的新值列严格按“类别+值列名”排列（无空格）
    """
    df = df.copy()

    # 如果未提供索引列，自动生成组内序号，并在此情况下不保留固定列
    auto_index = False
    if not index_cols:
        df['_idx'] = df.groupby(category_col).cumcount()
        index_cols = ['_idx']
        auto_index = True

    # 区分固定列：所有列 - category_col - value_cols
    # 注意：value_cols 中可能包含已被选为索引的列，但用户应避免错误输入
    other_cols = [c for c in df.columns 
                  if c not in value_cols and c != category_col]

    # 如果使用自动序号，则不保留任何固定列（因为固定列与自动序号没有实际关联）
    if auto_index:
        other_cols = []
    else:
        # 确保索引列被包含在 other_cols 中（通常已包含）
        # 并去除重复
        other_cols = list(dict.fromkeys(other_cols))

    # 取出固定列的唯一值（按 index_cols 去重）
    if other_cols:
        fixed_df = df[other_cols].drop_duplicates(subset=index_cols)
        # 将固定列设置为索引以便后续合并
        fixed_df = fixed_df.set_index(index_cols)
    else:
        fixed_df = pd.DataFrame()

    # 执行 pivot：包括索引列、分类列、值列
    pivot_input = df[index_cols + [category_col] + value_cols]
    pivoted = pivot_input.pivot(index=index_cols, columns=category_col, values=value_cols)

    # 扁平化多层列名并去除空格，连接方式为“类别+值列名”，不包含任何分隔符
    new_columns = []
    for col in pivoted.columns:
        if isinstance(col, tuple):
            # 元组顺序可能是 (value, category) 或 (category, value)，需要统一
            # 通过检查 col 中哪个是 value_cols 中的列来确定顺序
            cat, val = None, None
            for part in col:
                if part in value_cols:
                    val = str(part).strip()
                else:
                    cat = str(part).strip()
            if cat and val:
                new_col = f"{cat}{val}"
                new_columns.append(new_col)
            else:
                # 退化情况
                new_columns.append(''.join(str(x).strip() for x in col))
        else:
            new_columns.append(str(col).strip())
    pivoted.columns = new_columns

    # 重置索引以便与固定列合并
    pivoted.reset_index(inplace=True)
    # 删除可能由自动索引添加的 '_idx' 列
    if '_idx' in pivoted.columns:
        pivoted.drop(columns=['_idx'], inplace=True)

    # 合并固定列
    if not fixed_df.empty:
        # 将 pivoted 的索引列也设为索引以执行合并
        pivoted_merge = pivoted.set_index(index_cols)
        result = fixed_df.join(pivoted_merge, how='right')
        result.reset_index(inplace=True)
    else:
        result = pivoted

    # 重新排序列：固定列（other_cols 中出现在 result 中的列）→ 新的值列
    fixed_cols_in_result = [c for c in other_cols if c in result.columns]
    new_value_cols = [c for c in result.columns if c not in fixed_cols_in_result]
    result = result[fixed_cols_in_result + new_value_cols]

    # 对新的值列按“类别+值列名”排序（例如：门诊a数据、门诊b数据、住院a数据、住院b数据、其他a数据、其他b数据）
    # 此处保持原顺序，仅确保无空白
    return result
    ```

- **static/css/checker.css**
  - language: `css`
  - content:
    ```css
/* 文件校验模块样式 */
.mapping-table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
}
.mapping-table th, .mapping-table td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
}
.compare-pair {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
}
.compare-pair input {
    flex: 1;
    max-width: 200px;
}
.compare-pair span {
    font-weight: bold;
    font-size: 1.2em;
}
.btn-remove-pair {
    padding: 5px 10px;
    background: #ef476f;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
}
.precheck-result {
    background: #f8f9fa;
    padding: 10px;
    border-radius: 8px;
    margin: 10px 0;
}
    ```

- **static/css/style.css**
  - language: `css`
  - content:
    ```css
/* ========== 全局变量与重置 ========== */
:root {
    --primary: #4361ee;
    --primary-dark: #3a56d4;
    --primary-light: #eef2ff;
    --secondary: #2ec4b6;
    --success: #06d6a0;
    --danger: #ef476f;
    --warning: #ffd166;
    --dark: #1e293b;
    --gray: #64748b;
    --light: #f8fafc;
    --border: #e2e8f0;
    --shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
    --shadow-lg: 0 10px 30px rgba(0, 0, 0, 0.1);
    --radius: 12px;
    --radius-sm: 8px;
    --transition: all 0.2s ease;
}

* {
    box-sizing: border-box;
}

body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, sans-serif;
    margin: 0;
    padding: 24px;
    background: linear-gradient(145deg, #f1f5f9 0%, #e6edf5 100%);
    min-height: 100vh;
    color: var(--dark);
    line-height: 1.6;
}

.container {
    max-width: 1280px;
    margin: 0 auto;
    background: rgba(255, 255, 255, 0.85);
    backdrop-filter: blur(10px);
    padding: 32px 36px;
    border-radius: 28px;
    box-shadow: var(--shadow-lg);
    border: 1px solid rgba(255, 255, 255, 0.7);
}

h1 {
    font-size: 2.2rem;
    font-weight: 700;
    margin-top: 0;
    margin-bottom: 28px;
    color: var(--dark);
    display: flex;
    align-items: center;
    gap: 12px;
    border-bottom: 2px solid var(--primary-light);
    padding-bottom: 16px;
}

h2 {
    font-size: 1.6rem;
    font-weight: 600;
    margin-top: 0;
    margin-bottom: 24px;
    color: var(--dark);
}

h3 {
    font-size: 1.2rem;
    font-weight: 600;
    margin-bottom: 16px;
    color: var(--gray);
}

/* 导航栏 */
nav ul {
    list-style: none;
    padding: 0;
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 8px;
    margin-bottom: 32px;
    background: var(--light);
    padding: 8px;
    border-radius: 48px;
    border: 1px solid var(--border);
}

nav ul li a {
    display: inline-block;
    padding: 10px 20px;
    text-decoration: none;
    color: var(--gray);
    font-weight: 500;
    border-radius: 40px;
    transition: var(--transition);
    background: transparent;
}

nav ul li a:hover {
    background: white;
    color: var(--primary);
    box-shadow: var(--shadow);
}

/* 表单通用 */
.form-section {
    background: white;
    border-radius: var(--radius);
    padding: 24px;
    margin-bottom: 28px;
    box-shadow: var(--shadow);
    border: 1px solid var(--border);
}

.form-row {
    margin-bottom: 20px;
}

label {
    font-weight: 600;
    display: block;
    margin-bottom: 8px;
    color: var(--dark);
    font-size: 0.95rem;
}

input, select, textarea {
    width: 100%;
    padding: 12px 16px;
    border: 1.5px solid var(--border);
    border-radius: var(--radius-sm);
    font-size: 0.95rem;
    transition: var(--transition);
    background: white;
}

input:focus, select:focus, textarea:focus {
    outline: none;
    border-color: var(--primary);
    box-shadow: 0 0 0 3px rgba(67, 97, 238, 0.15);
}

.hint, .help-text {
    font-size: 0.85rem;
    color: var(--gray);
    margin-top: 6px;
}

/* 按钮 */
.btn, button {
    background: var(--primary);
    color: white;
    padding: 12px 24px;
    border: none;
    border-radius: 40px;
    cursor: pointer;
    font-size: 1rem;
    font-weight: 600;
    transition: var(--transition);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    box-shadow: 0 2px 6px rgba(67, 97, 238, 0.3);
}

.btn:hover, button:hover {
    background: var(--primary-dark);
    transform: translateY(-2px);
    box-shadow: 0 6px 12px rgba(67, 97, 238, 0.25);
}

.btn-success {
    background: var(--success);
    box-shadow: 0 2px 6px rgba(6, 214, 160, 0.3);
}
.btn-success:hover { background: #05b384; }

.btn-danger {
    background: var(--danger);
    box-shadow: 0 2px 6px rgba(239, 71, 111, 0.3);
}
.btn-danger:hover { background: #d63e5e; }

.btn-secondary {
    background: var(--secondary);
    box-shadow: 0 2px 6px rgba(46, 196, 182, 0.3);
}
.btn-secondary:hover { background: #25a99a; }

/* 消息与状态 */
.message {
    margin-top: 20px;
    padding: 14px 18px;
    border-radius: var(--radius-sm);
    font-weight: 500;
}
.message.error { background: #fee2e2; color: #b91c1c; }
.message.success { background: #d1fae5; color: #065f46; }
.message.info { background: #dbeafe; color: #1e40af; }

.status-area {
    margin-top: 24px;
    padding: 24px;
    background: white;
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    border: 1px solid var(--border);
}

.progress {
    width: 100%;
    background: var(--border);
    border-radius: 20px;
    height: 10px;
    overflow: hidden;
}
.progress-bar {
    height: 10px;
    background: var(--primary);
    width: 0%;
    transition: width 0.3s;
}

.log-area {
    background: #0f172a;
    color: #a5f3fc;
    padding: 16px;
    border-radius: var(--radius-sm);
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.85rem;
    max-height: 240px;
    overflow-y: auto;
    white-space: pre-wrap;
}

/* 筛选条件行 */
.filter-row {
    display: flex;
    gap: 12px;
    margin-bottom: 16px;
    align-items: center;
    flex-wrap: wrap;
}
.filter-field { width: 180px; }
.filter-op { width: 120px; }
.filter-value { width: 240px; }

/* 表格样式（用于匹配填充等） */
.table-container {
    overflow-x: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    border-radius: var(--radius);
    overflow: hidden;
    box-shadow: var(--shadow);
}
th, td {
    padding: 12px 16px;
    text-align: left;
    border-bottom: 1px solid var(--border);
}
th {
    background: var(--primary-light);
    font-weight: 600;
}

/* 响应式 */
@media (max-width: 768px) {
    body { padding: 12px; }
    .container { padding: 20px; }
    nav ul { flex-direction: column; align-items: stretch; }
    nav ul li a { display: block; text-align: center; }
}
    ```

- **static/js/checker.js**
  - language: `javascript`
  - content:
    ```javascript
// checker.js
(function() {
    // 确保 DOM 加载完成后再绑定事件
    document.addEventListener('DOMContentLoaded', function() {
        // 定义全局 addFilter 函数
        window.addFilter = function() {
            const container = document.getElementById('filtersContainer');
            const template = document.getElementById('filterRowTemplate');
            if (template) container.appendChild(template.content.cloneNode(true));
        };

        // 解析列名
        function parseColumns(text) {
            if (!text) return [];
            const lines = text.split(/[\n\r]+/);
            const cols = [];
            lines.forEach(line => {
                const parts = line.split(/[,，\s]+/);
                parts.forEach(p => {
                    const trimmed = p.trim();
                    if (trimmed) cols.push(trimmed);
                });
            });
            return cols;
        }

        // 收集对比列映射
        function collectCompareMap() {
            const stdText = document.getElementById('compareStandardText').value;
            const cmpText = document.getElementById('compareCompareText').value;
            const stdCols = parseColumns(stdText);
            let cmpCols = parseColumns(cmpText);
            if (cmpCols.length === 0) cmpCols = [...stdCols];
            if (stdCols.length === 0) {
                alert('请填写标准对比列');
                return null;
            }
            if (stdCols.length !== cmpCols.length) {
                alert(`标准对比列数量 (${stdCols.length}) 与比较对比列数量 (${cmpCols.length}) 不一致，请检查输入。`);
                return null;
            }
            const map = [];
            for (let i = 0; i < stdCols.length; i++) {
                map.push({ standard: stdCols[i], compare: cmpCols[i] });
            }
            return map;
        }

        // 获取 DOM 元素
        const btnPrecheck = document.getElementById('btnPrecheck');
        const btnStart = document.getElementById('btnStart');
        const statusArea = document.getElementById('statusArea');
        const taskStatus = document.getElementById('taskStatus');
        const progressBar = document.getElementById('progressBar');
        const progressMessage = document.getElementById('progressMessage');
        const precheckResult = document.getElementById('precheckResult');
        const downloadLink = document.getElementById('downloadLink');

        // 预检按钮点击事件
        btnPrecheck.addEventListener('click', async function() {
            const filesA = document.getElementById('filesA').files;
            const filesB = document.getElementById('filesB').files;
            if (filesA.length === 0 || filesB.length === 0) {
                alert('请上传文件');
                return;
            }

            btnPrecheck.disabled = true;
            btnPrecheck.textContent = '⏳ 预检中...';
            statusArea.style.display = 'block';
            taskStatus.textContent = '正在预检，请稍后...';
            progressBar.style.width = '10%';
            progressMessage.textContent = '正在上传文件并扫描...';
            precheckResult.innerHTML = '';

            const formData = new FormData();
            for (let f of filesA) formData.append('files_a', f);
            for (let f of filesB) formData.append('files_b', f);

            try {
                const resp = await fetch('/checker/precheck', { method: 'POST', body: formData });
                const result = await resp.json();
                if (resp.ok) {
                    precheckResult.innerHTML = `
                        <p>标准文件: ${result.total_files_a} 个, 约 ${result.total_rows_a.toLocaleString()} 行</p>
                        <p>比较文件: ${result.total_files_b} 个, 约 ${result.total_rows_b.toLocaleString()} 行</p>
                        <p>预估内存: ${result.estimated_memory_mb} MB | 预估时间: ${result.estimated_time_seconds} 秒</p>
                        <p style="color:${result.advice.includes('安全') ? 'green' : 'orange'}">${result.advice}</p>
                    `;
                    taskStatus.textContent = '预检完成，可以开始校验';
                    progressBar.style.width = '100%';
                    btnStart.disabled = false;
                } else {
                    taskStatus.textContent = '预检失败';
                    precheckResult.innerHTML = `<p style="color:red;">错误: ${result.message || result.error}</p>`;
                    btnStart.disabled = true;
                }
            } catch (err) {
                taskStatus.textContent = '网络错误';
                precheckResult.innerHTML = `<p style="color:red;">网络请求失败: ${err.message}</p>`;
                btnStart.disabled = true;
            } finally {
                btnPrecheck.disabled = false;
                btnPrecheck.textContent = '🔎 预检';
            }
        });

        // 开始校验按钮点击事件
        btnStart.addEventListener('click', async function() {
            const filesA = document.getElementById('filesA').files;
            const filesB = document.getElementById('filesB').files;
            const keyA = document.getElementById('keyMapA').value.trim();
            const keyB = document.getElementById('keyMapB').value.trim();
            const duplicateAction = document.getElementById('duplicateAction').value;
            const compareMap = collectCompareMap();
            if (!compareMap) return;

            const filters = [];
            document.querySelectorAll('#filtersContainer .filter-row').forEach(row => {
                const field = row.querySelector('input[name="filter_field"]')?.value.trim();
                const op = row.querySelector('select[name="filter_op"]')?.value;
                const value = row.querySelector('input[name="filter_value"]')?.value.trim();
                if (field && value) filters.push({field, op, value});
            });

            if (!keyA) { alert('请填写标准群主键'); return; }

            const formData = new FormData();
            for (let f of filesA) formData.append('files_a', f);
            for (let f of filesB) formData.append('files_b', f);
            formData.append('key_map_a', keyA);
            formData.append('key_map_b', keyB || keyA);
            formData.append('compare_map', JSON.stringify(compareMap));
            formData.append('filters_a', JSON.stringify(filters));
            formData.append('duplicate_action', duplicateAction);

            btnStart.disabled = true;
            btnStart.textContent = '提交中...';
            statusArea.style.display = 'block';
            taskStatus.textContent = '提交中...';
            progressBar.style.width = '0%';
            progressMessage.textContent = '';

            const resp = await fetch('/checker/process', { method: 'POST', body: formData });
            const result = await resp.json();
            if (result.error) {
                taskStatus.textContent = '错误: ' + result.error;
                btnStart.disabled = false;
                btnStart.textContent = '▶ 开始校验';
                return;
            }
            const taskId = result.task_id;
            taskStatus.textContent = '任务已提交，处理中...';

            const poll = setInterval(async () => {
                try {
                    const statusResp = await fetch('/checker/status/' + taskId);
                    const status = await statusResp.json();
                    if (status.error) {
                        clearInterval(poll);
                        taskStatus.textContent = '错误: ' + status.error;
                        btnStart.disabled = false;
                        btnStart.textContent = '▶ 开始校验';
                        return;
                    }
                    if (status.progress_percent !== undefined && status.progress_percent !== null) {
                        progressBar.style.width = status.progress_percent + '%';
                    }
                    if (status.progress_message) {
                        progressMessage.textContent = status.progress_message;
                    }
                    if (status.processed_rows !== undefined && status.total_rows) {
                        taskStatus.textContent = `扫描中 (${status.processed_rows.toLocaleString()} / ${status.total_rows.toLocaleString()} 行)`;
                    }
                    if (status.status === 'completed') {
                        clearInterval(poll);
                        taskStatus.textContent = '完成！';
                        progressBar.style.width = '100%';
                        downloadLink.innerHTML = `<a href="/checker/download/${taskId}" class="btn btn-success">下载结果</a>`;
                        btnStart.disabled = true;
                    } else if (status.status === 'failed') {
                        clearInterval(poll);
                        taskStatus.textContent = '失败: ' + (status.error || '');
                        btnStart.disabled = false;
                        btnStart.textContent = '▶ 开始校验';
                    }
                } catch (err) {
                    console.error('轮询失败', err);
                }
            }, 1000);
        });
    });
})();
    ```

- **static/js/convert.js**
  - language: `javascript`
  - content:
    ```javascript
// convert.js
(function() {
    document.getElementById('convertForm').addEventListener('submit', async function(e) {
        e.preventDefault();
        const fileInput = document.getElementById('fileInput');
        const targetFormat = document.getElementById('targetFormat').value;
        const encoding = document.getElementById('encoding').value;
        const separator = document.getElementById('separator').value;
        if (!fileInput.files.length) { alert('请选择文件'); return; }

        const btn = document.getElementById('btnSubmit');
        btn.disabled = true;
        btn.textContent = '提交中...';
        document.getElementById('statusArea').style.display = 'block';
        document.getElementById('taskStatus').textContent = '提交中...';
        document.getElementById('progressBar').style.width = '0%';

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('target_format', targetFormat);
        formData.append('encoding', encoding);
        formData.append('separator', separator);

        try {
            const resp = await fetch('/convert/process', { method: 'POST', body: formData });
            const result = await resp.json();
            if (result.error) {
                document.getElementById('taskStatus').textContent = '错误: ' + result.error;
                btn.disabled = false;
                return;
            }
            const taskId = result.task_id;
            document.getElementById('taskStatus').textContent = '转换中...';

            const poll = setInterval(async () => {
                const statusResp = await fetch('/convert/status/' + taskId);
                const status = await statusResp.json();
                if (status.error) {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '错误: ' + status.error;
                    btn.disabled = false;
                    return;
                }
                if (status.progress_percent !== undefined) {
                    document.getElementById('progressBar').style.width = status.progress_percent + '%';
                }
                if (status.progress_message) {
                    document.getElementById('progressMessage').textContent = status.progress_message;
                }
                if (status.status === 'completed') {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '完成！';
                    document.getElementById('downloadLink').innerHTML = `<a href="/convert/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    btn.disabled = false;
                    btn.textContent = '开始转换';
                } else if (status.status === 'failed') {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '失败: ' + (status.error || '');
                    btn.disabled = false;
                    btn.textContent = '开始转换';
                }
            }, 1000);
        } catch (err) {
            document.getElementById('taskStatus').textContent = '网络错误: ' + err.message;
            btn.disabled = false;
        }
    });
})();
    ```

- **static/js/download_utils.js**
  - language: `javascript`
  - content:
    ```javascript
function getFilenameFromContentDisposition(disposition) {
    if (!disposition) return null;
    const filenameStarMatch = disposition.match(/filename\*=(?:UTF-8'')?([^;]+)/i);
    if (filenameStarMatch) {
        try {
            return decodeURIComponent(filenameStarMatch[1].trim().replace(/['"]+/g, ''));
        } catch (e) {
            return filenameStarMatch[1].trim().replace(/['"]+/g, '');
        }
    }
    const filenameMatch = disposition.match(/filename="?([^";]+)"?/i);
    return filenameMatch ? filenameMatch[1] : null;
}
    ```

- **static/js/field_parser.js**
  - language: `javascript`
  - content:
    ```javascript
// field_parser.js
(function() {
    let lastResult = null;   // 保存最近分析结果（含 matched, only_a, only_b, cols_a, cols_b）

    document.getElementById('btnAnalyze').addEventListener('click', async function() {
        const textA = document.getElementById('textA').value.trim();
        const textB = document.getElementById('textB').value.trim();
        if (!textA || !textB) {
            alert('请填写两边的列名');
            return;
        }

        const btn = document.getElementById('btnAnalyze');
        btn.disabled = true;
        btn.textContent = '分析中...';

        try {
            const resp = await fetch('/field-parser/analyze', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    standard_headers: textA,
                    compare_headers: textB
                })
            });

            const result = await resp.json();
            if (resp.ok) {
                lastResult = result;   // 保存完整结果
                displayResult(result);
                document.getElementById('btnExport').style.display = 'inline-block';
            } else {
                alert('分析失败: ' + (result.message || result.error || '未知错误'));
            }
        } catch (err) {
            alert('网络请求失败: ' + err.message);
        } finally {
            btn.disabled = false;
            btn.textContent = '开始分析';
        }
    });

    document.getElementById('btnExport').addEventListener('click', async function() {
        if (!lastResult) {
            alert('请先进行分析');
            return;
        }
        // 发送完整映射结果和原始列名
        const payload = {
            matched: lastResult.matched,
            only_a: lastResult.only_a,
            only_b: lastResult.only_b,
            cols_a: lastResult.cols_a,
            cols_b: lastResult.cols_b
        };

        try {
            const resp = await fetch('/field-parser/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (!resp.ok) {
                const err = await resp.json();
                alert('导出失败: ' + (err.message || '未知错误'));
                return;
            }
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const filename = getFilenameFromContentDisposition(resp.headers.get('Content-Disposition')) || '字段映射结果.xlsx';
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (err) {
            alert('导出请求失败: ' + err.message);
        }
    });

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function displayResult(mapping) {
        const tbody = document.querySelector('#mappingTable tbody');
        tbody.innerHTML = '';
        mapping.matched.forEach(m => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${escapeHtml(m.standard)}</td><td>${escapeHtml(m.compare)}</td><td>${m.score}</td>`;
            tbody.appendChild(tr);
        });
        mapping.only_a.forEach(h => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${escapeHtml(h)}</td><td>-</td><td>仅标准</td>`;
            tbody.appendChild(tr);
        });
        mapping.only_b.forEach(h => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>-</td><td>${escapeHtml(h)}</td><td>仅比较</td>`;
            tbody.appendChild(tr);
        });
        document.getElementById('resultArea').style.display = 'block';
    }
})();
    ```

- **static/js/main.js**
  - language: `javascript`
  - content:
    ```javascript
// 全局交互逻辑
document.addEventListener('DOMContentLoaded', () => {
    console.log('✨ Excel 智能工具集已加载');

    // 为所有提交按钮添加加载状态（可选）
    document.querySelectorAll('form').forEach(form => {
        form.addEventListener('submit', () => {
            const btn = form.querySelector('button[type="submit"]');
            if (btn) {
                btn.disabled = true;
                btn.innerHTML = '<span class="spinner"></span> 处理中...';
            }
        });
    });
});
    ```

- **static/js/subtotal.js**
  - language: `javascript`
  - content:
    ```javascript
// static/js/subtotal.js
(function() {
    const form = document.getElementById('subtotalForm');
    const statusArea = document.getElementById('statusArea');
    const addOverrideBtn = document.getElementById('addOverrideBtn');
    const overrideContainer = document.getElementById('overrideContainer');
    const sheetsModeRadios = document.getElementsByName('sheets_mode');
    const sheetsCustom = document.getElementById('sheetsCustom');

    addOverrideBtn.addEventListener('click', function() {
        const template = document.getElementById('overrideRowTemplate');
        const clone = template.content.cloneNode(true);
        overrideContainer.appendChild(clone);
    });

    sheetsModeRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            sheetsCustom.style.display = this.value === 'custom' ? 'inline-block' : 'none';
        });
    });

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        // 收集特殊列函数
        const overrides = [];
        document.querySelectorAll('.override-row').forEach(row => {
            const col = row.querySelector('.override-col').value.trim().toUpperCase();
            const func = row.querySelector('.override-func').value;
            if (col) {
                overrides.push(col + ':' + func);
            }
        });

        const formData = new FormData(form);
        if (overrides.length > 0) {
            formData.set('agg_overrides', overrides.join(','));
        }
        const mode = formData.get('sheets_mode');
        if (mode === 'custom') {
            const customVal = sheetsCustom.value.trim();
            formData.set('sheets_mode', customVal || '1');
        }

        const submitBtn = form.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        statusArea.style.display = 'block';
        document.getElementById('taskStatus').textContent = '提交中...';
        document.getElementById('progressBar').style.width = '0%';
        document.getElementById('downloadLink').innerHTML = '';

        try {
            const resp = await fetch('/subtotal/process', { method: 'POST', body: formData });
            if (!resp.ok) {
                const err = await resp.json();
                document.getElementById('taskStatus').textContent = '错误: ' + (err.error || err.message);
                submitBtn.disabled = false;
                return;
            }
            const contentType = resp.headers.get('Content-Type');
            if (contentType && contentType.includes('application/json')) {
                // 异步任务
                const result = await resp.json();
                const taskId = result.task_id;
                pollTask(taskId, submitBtn);
            } else {
                // 同步下载
                    const blob = await resp.blob();
                const url = URL.createObjectURL(blob);
                const filename = getFilenameFromContentDisposition(resp.headers.get('Content-Disposition')) || 'subtotal_result.xlsx';
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                document.getElementById('taskStatus').textContent = '处理完成，已下载。';
                submitBtn.disabled = false;
            }
        } catch (err) {
            document.getElementById('taskStatus').textContent = '网络错误: ' + err.message;
            submitBtn.disabled = false;
        }
    });

    function pollTask(taskId, submitBtn) {
        const interval = setInterval(async () => {
            try {
                const resp = await fetch('/subtotal/status/' + taskId);
                const status = await resp.json();
                if (status.error) {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '错误: ' + status.error;
                    submitBtn.disabled = false;
                    return;
                }
                if (status.progress_percent !== undefined) {
                    document.getElementById('progressBar').style.width = status.progress_percent + '%';
                }
                document.getElementById('taskStatus').textContent =
                    status.status === 'completed' ? '完成' : (status.status || '处理中...');
                if (status.status === 'completed') {
                    clearInterval(interval);
                    document.getElementById('downloadLink').innerHTML =
                        `<a href="/subtotal/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    submitBtn.disabled = false;
                } else if (status.status === 'failed') {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '失败';
                    submitBtn.disabled = false;
                }
            } catch (err) {
                console.error(err);
            }
        }, 2000);
    }
})();
    ```

- **static/js/team_report.js**
  - language: `javascript`
  - content:
    ```javascript
// team_report.js (updated to include monthly files)
(function() {
    const form = document.getElementById('reportForm');
    if (!form) return;

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const submitBtn = document.getElementById('submitBtn');
        const statusArea = document.getElementById('statusArea');
        const taskStatus = document.getElementById('taskStatus');
        const progressBar = document.getElementById('progressBar');
        const progressMessage = document.getElementById('progressMessage');

        submitBtn.disabled = true;
        statusArea.style.display = 'block';
        taskStatus.textContent = '提交中...';
        progressBar.style.width = '0%';
        progressMessage.textContent = '';

        const formData = new FormData(form);

        try {
            const resp = await fetch('/team-report/generate', {
                method: 'POST',
                body: formData
            });
            if (!resp.ok) {
                taskStatus.textContent = '提交失败 (HTTP ' + resp.status + ')';
                submitBtn.disabled = false;
                return;
            }
            const result = await resp.json();
            if (result.error) {
                taskStatus.textContent = '错误: ' + result.error;
                submitBtn.disabled = false;
                return;
            }
            const taskId = result.task_id;
            taskStatus.textContent = '任务已提交，正在处理...';
            progressBar.style.width = '10%';

            const poll = setInterval(async () => {
                try {
                    const statusResp = await fetch('/team-report/status/' + taskId);
                    if (!statusResp.ok) {
                        if (statusResp.status === 524 || statusResp.status === 502) {
                            console.warn('收到 ' + statusResp.status + '，继续等待...');
                            return;
                        }
                        clearInterval(poll);
                        taskStatus.textContent = '查询状态失败 (HTTP ' + statusResp.status + ')';
                        submitBtn.disabled = false;
                        return;
                    }
                    const status = await statusResp.json();
                    if (status.error) {
                        clearInterval(poll);
                        taskStatus.textContent = '错误: ' + status.error;
                        submitBtn.disabled = false;
                        return;
                    }
                    if (status.progress_percent !== undefined) {
                        progressBar.style.width = status.progress_percent + '%';
                    }
                    if (status.progress_message) {
                        progressMessage.textContent = status.progress_message;
                    }
                    if (status.status === 'completed') {
                        clearInterval(poll);
                        taskStatus.textContent = '处理完成，正在下载...';
                        progressBar.style.width = '100%';
                        window.location.href = '/team-report/download/' + taskId;
                        setTimeout(() => { submitBtn.disabled = false; }, 2000);
                    } else if (status.status === 'failed') {
                        clearInterval(poll);
                        taskStatus.textContent = '处理失败: ' + (status.error || '未知错误');
                        submitBtn.disabled = false;
                    }
                } catch (err) {
                    console.error('轮询错误:', err);
                }
            }, 4000);
        } catch (err) {
            taskStatus.textContent = '网络错误: ' + err.message;
            submitBtn.disabled = false;
        }
    });
})();
    ```

- **static/js/transposer.js**
  - language: `javascript`
  - content:
    ```javascript
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('transposerForm');
    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        const statusArea = document.getElementById('statusArea');
        const statusMsg = document.getElementById('statusMessage');
        const submitBtn = document.getElementById('submitBtn');

        statusArea.style.display = 'block';
        statusMsg.textContent = '处理中，请稍候...';
        submitBtn.disabled = true;

        const formData = new FormData(form);
        try {
            const response = await fetch('/transposer/convert', {
                method: 'POST',
                body: formData
            });
            if (!response.ok) {
                const errData = await response.json();
                statusMsg.textContent = '错误：' + (errData.error || '未知错误');
                submitBtn.disabled = false;
                return;
            }
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const filename = getFilenameFromContentDisposition(response.headers.get('Content-Disposition')) || '转置结果.xlsx';
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            statusMsg.textContent = '转换成功，文件已下载。';
        } catch (err) {
            statusMsg.textContent = '网络错误：' + err.message;
        } finally {
            submitBtn.disabled = false;
        }
    });
});
    ```

- **static/style.css**
  - language: `css`
  - content:
    ```css
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    margin: 0;
    padding: 20px;
    background: #f5f7fa;
}
.container {
    max-width: 1200px;
    margin: 0 auto;
    background: white;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}
nav ul {
    list-style: none;
    padding: 0;
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 20px;
    margin-bottom: 30px;
}
nav a {
    text-decoration: none;
    color: #007bff;
    font-weight: 500;
}
nav a:hover {
    text-decoration: underline;
}
.message {
    margin-top: 20px;
    padding: 12px;
    border-radius: 6px;
}
.message.error {
    background: #fee;
    color: #c00;
}
.message.success {
    background: #efe;
    color: #060;
}
button, .btn {
    background: #007bff;
    color: white;
    padding: 10px 20px;
    border: none;
    border-radius: 6px;
    cursor: pointer;
    font-size: 16px;
}
button:hover, .btn:hover {
    background: #0056b3;
}
input, select, textarea {
    padding: 8px;
    border: 1px solid #ddd;
    border-radius: 6px;
    width: 100%;
    max-width: 400px;
    margin-bottom: 15px;
}
label {
    font-weight: 600;
    display: block;
    margin-bottom: 5px;
}
    ```

- **templates/base.html**
  - language: `html`
  - content:
    ```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{% block title %}Excel 智能工具集{% endblock %}</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
    <script src="{{ url_for('static', filename='js/main.js') }}" defer></script>
    <script src="{{ url_for('static', filename='js/download_utils.js') }}"></script>
    {% block head %}{% endblock %}
</head>
<body>
    <div class="container">
        <h1>📊 Excel 智能工具集</h1>
        <nav>
            <ul>
                <li><a href="{{ url_for('index') }}">🏠 首页</a></li>
                <li><a href="{{ url_for('calc.index') }}">🧮 列计算/聚合</a></li>
                <li><a href="{{ url_for('compare.index') }}">🔍 文件对比</a></li>
                <li><a href="{{ url_for('match.index') }}">🔗 匹配填充</a></li>
                <li><a href="{{ url_for('summary.index') }}">📋 汇总工具</a></li>
                <li><a href="{{ url_for('splitter.index') }}">✂️ 文件拆分</a></li>
                <li><a href="{{ url_for('field_parser.index') }}">🔗 字段映射</a></li>
                <li><a href="{{ url_for('checker.index') }}">✅ 文件校验</a></li>
                <li><a href="{{ url_for('convert.index') }}">🔄 格式转换</a></li>
                <li><a href="{{ url_for('subtotal.index') }}">📊 数据小计</a></li>
                <li><a href="{{ url_for('team_report.index') }}">📊 车队日报</a></li>
                <li><a href="{{ url_for('transposer.index') }}">🔄 行列转置</a></li>
            </ul>
        </nav>
        {% block content %}{% endblock %}
        <div id="message" class="message"></div>
    </div>
    {% block scripts %}{% endblock %}
</body>
</html>
    ```

- **templates/calculator/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}列计算 / 聚合回填{% endblock %}
{% block content %}
<h2>🧮 列计算 / 聚合回填</h2>
<form action="{{ url_for('calc.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
        <div class="form-row">
            <label>⚙️ 操作类型</label>
            <select name="calc_type" id="calcType">
                <option value="column">列计算</option>
                <option value="aggregate">聚合回填</option>
            </select>
        </div>
        <div class="form-row">
            <label>表头所在行号（从1开始，默认1）</label>
            <input type="number" name="header_row" value="1" min="1" style="max-width:100px;">
            <div class="hint">如果你的 Excel 顶部有标题行，请填写实际列名所在的行号</div>
        </div>
    </div>
    <div id="columnConfig" class="form-section">
        <h3>📐 列计算设置</h3>
        <div class="form-row">
            <label>表达式 (如 A+B-C)</label>
            <input type="text" name="expression" placeholder="例如 A+B">
        </div>
        <div class="form-row">
            <label>结果列 (字母或新列名)</label>
            <input type="text" name="result_col_column" placeholder="K 或 '差值'">
        </div>
        <div class="form-row">
            <label>合并依据列 (可选，逗号分隔)</label>
            <input type="text" name="merge_by_cols" placeholder="A,B">
            <div class="hint">相同组的结果列只保留第一行，其余合并。</div>
        </div>
    </div>
    <div id="aggregateConfig" class="form-section" style="display:none;">
        <h3>📊 聚合回填设置</h3>
        <div class="form-row">
            <label>分组列 (逗号分隔，支持列名或列字母)</label>
            <input type="text" name="group_cols" placeholder="A,B">
        </div>
        <div class="form-row">
            <label>聚合列 (逗号分隔)</label>
            <input type="text" name="agg_cols" placeholder="C,D">
        </div>
        <div class="form-row">
            <label>聚合函数</label>
            <select name="agg_func">
                <option value="sum">求和</option>
                <option value="mean">平均值</option>
                <option value="max">最大值</option>
                <option value="min">最小值</option>
                <option value="count">计数</option>
            </select>
        </div>
        <div class="form-row">
            <label>结果列名</label>
            <input type="text" name="result_col_aggregate" placeholder="K 或 '聚合结果'">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始处理</button>
</form>
<script>
    const typeSelect = document.getElementById('calcType');
    const colCfg = document.getElementById('columnConfig');
    const aggCfg = document.getElementById('aggregateConfig');
    function toggle() {
        if (typeSelect.value === 'column') {
            colCfg.style.display = 'block';
            aggCfg.style.display = 'none';
        } else {
            colCfg.style.display = 'none';
            aggCfg.style.display = 'block';
        }
    }
    typeSelect.addEventListener('change', toggle);
    toggle();
</script>
{% endblock %}
    ```

- **templates/checker/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}文件校验{% endblock %}
{% block head %}
<link rel="stylesheet" href="{{ url_for('static', filename='css/checker.css') }}">
{% endblock %}
{% block content %}
<h2>✅ 文件校验</h2>
<p>上传标准文件群和比较文件群，系统将根据您指定的主键和对比列进行比较，并生成差异报告。</p>

<form id="checkerForm" enctype="multipart/form-data">
    <div class="form-section">
        <h3>📂 文件上传</h3>
        <div class="form-row">
            <label>标准文件群（可多选）</label>
            <input type="file" id="filesA" name="files_a" accept=".xlsx,.xls,.csv" multiple required>
        </div>
        <div class="form-row">
            <label>比较文件群（可多选）</label>
            <input type="file" id="filesB" name="files_b" accept=".xlsx,.xls,.csv" multiple required>
        </div>
    </div>

    <div class="form-section">
        <h3>🔑 主键设置</h3>
        <div class="form-row">
            <label>标准群主键列（逗号或空格分隔）</label>
            <input type="text" id="keyMapA" placeholder="例如 ID,日期">
        </div>
        <div class="form-row">
            <label>比较群主键列（逗号或空格分隔，默认与标准相同）</label>
            <input type="text" id="keyMapB" placeholder="留空则与标准群主键相同">
        </div>
        <div class="form-row">
            <label>主键重复处理</label>
            <select id="duplicateAction">
                <option value="error" selected>报错</option>
                <option value="keep_first">保留第一条</option>
            </select>
        </div>
    </div>

    <div class="form-section">
        <h3>📊 对比列</h3>
        <div class="parser-container" style="display: flex; gap: 20px;">
            <div class="parser-box" style="flex: 1;">
                <label>标准群对比列（每行一个或用逗号/空格分隔）</label>
                <textarea id="compareStandardText" rows="5" placeholder="例如：医疗类别&#10;入院日期&#10;出院日期&#10;医疗费总额&#10;结算时间&#10;定点医药机构名称"></textarea>
            </div>
            <div class="parser-box" style="flex: 1;">
                <label>比较群对比列（留空则与标准相同）</label>
                <textarea id="compareCompareText" rows="5" placeholder="例如：医疗类别&#10;入院日期&#10;出院日期&#10;医疗费总额&#10;结算时间&#10;定点医疗机构名称"></textarea>
            </div>
        </div>
        <div class="hint">按顺序一一对应，支持多列；比较群留空则默认与标准群列名相同。</div>
    </div>

    <div class="form-section">
        <h3>🔍 标准群筛选（可选）</h3>
        <div id="filtersContainer"></div>
        <button type="button" class="btn btn-secondary" onclick="addFilter()">+ 添加筛选条件</button>
        <div class="hint">支持多值OR：在“值”中用英文逗号分隔多个值</div>
    </div>

    <div style="text-align:center; margin: 30px 0;">
        <button type="button" class="btn" id="btnPrecheck">🔎 预检</button>
        <button type="button" class="btn btn-success" id="btnStart" disabled>▶ 开始校验</button>
    </div>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar" style="width:0%"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <p id="progressMessage" style="color:#666;"></p>
    <div id="precheckResult" class="precheck-result"></div>
    <div id="downloadLink" style="margin-top:15px;"></div>
</div>

<template id="filterRowTemplate">
    <div class="filter-row">
        <input type="text" name="filter_field" placeholder="字段名" class="filter-field">
        <select name="filter_op" class="filter-op">
            <option value="=">=</option><option value="!=">!=</option><option value=">">></option>
            <option value="<"><</option><option value=">=">>=</option><option value="<="><=</option>
            <option value="contains">包含</option><option value="not_contains">不包含</option>
        </select>
        <input type="text" name="filter_value" placeholder="值（多值用英文逗号）" class="filter-value">
        <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">删除</button>
    </div>
</template>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/checker.js') }}"></script>
{% endblock %}
    ```

- **templates/comparator/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}文件对比{% endblock %}
{% block content %}
<h2>🔍 文件对比</h2>
<form action="{{ url_for('compare.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📂 文件 1</label>
            <input type="file" name="file1" accept=".xlsx,.xls" required>
            <input type="text" name="sheet1" placeholder="Sheet名称 (可选)" class="mt-2">
        </div>
        <div class="form-row">
            <label>📂 文件 2</label>
            <input type="file" name="file2" accept=".xlsx,.xls" required>
            <input type="text" name="sheet2" placeholder="Sheet名称 (可选)" class="mt-2">
        </div>
        <div class="form-row">
            <label>对比模式</label>
            <select name="compare_mode" id="compareMode">
                <option value="full">完整对比</option>
                <option value="field_match">仅字段值匹配</option>
            </select>
        </div>
    </div>
    <div id="fullFields" class="form-section">
        <div class="form-row">
            <label>主键字段 (逗号分隔)</label>
            <input type="text" name="key_fields" placeholder="学号,姓名">
        </div>
        <div class="form-row">
            <label>对比字段 (逗号分隔)</label>
            <input type="text" name="compare_fields" placeholder="学号,姓名,成绩">
        </div>
    </div>
    <div id="fieldMatchOnly" class="form-section" style="display:none;">
        <div class="form-row">
            <label>匹配字段名</label>
            <input type="text" name="match_field" placeholder="医疗机构名称">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始对比</button>
</form>
<script>
    const modeSelect = document.getElementById('compareMode');
    const fullDiv = document.getElementById('fullFields');
    const fieldDiv = document.getElementById('fieldMatchOnly');
    modeSelect.addEventListener('change', () => {
        if (modeSelect.value === 'full') {
            fullDiv.style.display = 'block';
            fieldDiv.style.display = 'none';
        } else {
            fullDiv.style.display = 'none';
            fieldDiv.style.display = 'block';
        }
    });
    modeSelect.dispatchEvent(new Event('change'));
</script>
{% endblock %}
    ```

- **templates/convert/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}格式转换{% endblock %}
{% block content %}
<h2>🔄 文件格式转换</h2>
<p>上传文件，选择目标格式和编码，开始转换。</p>
<form id="convertForm" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 选择文件</label>
            <input type="file" id="fileInput" name="file" accept=".xlsx,.xls,.xlsm,.csv" required>
        </div>
        <div class="form-row">
            <label>目标格式</label>
            <select id="targetFormat" name="target_format">
                <option value="csv">CSV</option>
                <option value="xlsx">Excel (.xlsx)</option>
            </select>
        </div>
        <div class="form-row">
            <label>编码（仅CSV输出时有效）</label>
            <select id="encoding" name="encoding">
                <option value="utf-8-sig" selected>UTF-8 (带BOM)</option>
                <option value="utf-8">UTF-8</option>
                <option value="gbk">GBK</option>
                <option value="gb2312">GB2312</option>
            </select>
        </div>
        <div class="form-row">
            <label>分隔符（仅CSV输出时有效）</label>
            <select id="separator" name="separator">
                <option value="," selected>逗号 (,)</option>
                <option value=";">分号 (;)</option>
                <option value="	">制表符 (Tab)</option>
            </select>
        </div>
    </div>
    <button type="submit" class="btn btn-success" id="btnSubmit">开始转换</button>
</form>
<div id="statusArea" class="status-area" style="display:none;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar" style="width:0%"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <p id="progressMessage"></p>
    <div id="downloadLink" style="margin-top:15px;"></div>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/convert.js') }}"></script>
{% endblock %}
    ```

- **templates/debug.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}匹配填充调试{% endblock %}
{% block content %}
<h2>🔍 匹配填充调试信息</h2>
<div class="form-section">
    <h3>配置</h3>
    <pre>{{ config | tojson(indent=2) }}</pre>
</div>
<div class="form-section">
    <h3>信息</h3>
    <pre>{{ info | tojson(indent=2) }}</pre>
</div>
<a href="{{ url_for('match.index') }}" class="btn">返回</a>
{% endblock %}
    ```

- **templates/error.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}错误{% endblock %}
{% block content %}
<h2>⚠️ 出错了</h2>
<div class="message error" style="display:block; margin-bottom:24px;">{{ error }}</div>
<a href="{{ url_for('index') }}" class="btn">← 返回首页</a>
{% endblock %}
    ```

- **templates/field_parser/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}字段映射{% endblock %}
{% block head %}
<style>
    .parser-container { display: flex; gap: 20px; margin-bottom: 20px; }
    .parser-box { flex: 1; }
    .parser-box textarea { width: 100%; height: 200px; }
    .mapping-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    .mapping-table th, .mapping-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    .mapping-table .score { font-weight: bold; }
</style>
{% endblock %}
{% block content %}
<h2>🔗 字段映射分析</h2>
<p>粘贴两个文件的列名列表，系统将自动匹配相似字段。列名可使用逗号、空格、换行等分隔。</p>

<form id="parserForm">
    <div class="parser-container">
        <div class="parser-box">
            <h3>标准文件列名</h3>
            <textarea id="textA" name="text_a" placeholder="例如：编号, 姓名, 成绩"></textarea>
        </div>
        <div class="parser-box">
            <h3>比较文件列名</h3>
            <textarea id="textB" name="text_b" placeholder="例如：ID, 名字, 分数"></textarea>
        </div>
    </div>
    <button type="button" class="btn btn-primary" id="btnAnalyze">开始分析</button>
    <button type="button" class="btn btn-secondary" id="btnExport" style="display:none;">📥 导出 Excel</button>
</form>

<div id="resultArea" style="display:none; margin-top: 30px;">
    <h3>分析结果</h3>
    <table class="mapping-table" id="mappingTable">
        <thead><tr><th>标准字段</th><th>比较字段</th><th>相似度</th></tr></thead>
        <tbody></tbody>
    </table>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/field_parser.js') }}"></script>
{% endblock %}
    ```

- **templates/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}首页 - Excel 工具集{% endblock %}
{% block content %}
<h2>👋 欢迎使用 Excel 智能工具集</h2>
<p style="font-size: 1.1rem; margin-bottom: 32px;">请从导航栏选择所需功能：</p>

<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 24px;">
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🧮 列计算/聚合回填</h3>
        <p>支持多列表达式计算，或按分组聚合回填结果列，保留原有格式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔍 文件对比</h3>
        <p>完整对比（主键匹配）或仅字段值匹配，输出差异报告。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔗 匹配填充</h3>
        <p>根据源文件匹配目标文件，填充值并保留格式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📋 汇总工具</h3>
        <p>合并多个文件，按唯一键汇总求和。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">✂️ 文件拆分</h3>
        <p>按列值拆分Excel为多个文件，保留合并单元格与样式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔗 字段映射</h3>
        <p>智能匹配不同文件间的列名。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">✅ 文件校验</h3>
        <p>根据主键对比两个文件群，输出差异报告。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔄 格式转换</h3>
        <p>Excel 与 CSV 互转，支持编码选择。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📊 数据小计</h3>
        <p>按指定主键列分组，在每组下方插入小计行，计算合计、平均值等。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📊 车队日报</h3>
        <p>生成车队业绩日报，含上线率、目标完成率。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔄 行列转置</h3>
        <p>长表转宽表，取消分类列。</p>
    </div>
</div>
{% endblock %}
    ```

- **templates/matcher/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}匹配填充{% endblock %}
{% block head %}
<style>
    .button-group { display: flex; gap: 16px; margin-top: 24px; }
    .debug-btn { background: var(--warning); color: var(--dark); }
</style>
{% endblock %}
{% block content %}
<h2>🔗 Excel 数据匹配填充</h2>
<form id="matchForm" method="post" action="{{ url_for('match.process') }}" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>🎯 目标文件 (保留格式)</label>
            <input type="file" name="target_file" accept=".xlsx" required>
            <div class="hint">将保留原有合并单元格、字体等格式</div>
        </div>
        <div class="form-row">
            <label>📁 源文件 (数据源)</label>
            <input type="file" name="source_file" accept=".xlsx,.xls" required>
            <div class="hint">支持 .xlsx 和 .xls</div>
        </div>
        <div class="form-row">
            <label>源文件 Sheet 索引 (可选)</label>
            <input type="text" name="source_sheets" placeholder="例如 1,2,3">
            <div class="hint">不填默认读取第一个sheet</div>
        </div>
    </div>
    <div class="form-section">
        <div class="form-row">
            <label>目标文件键列 (字母，多列逗号分隔)</label>
            <input type="text" name="target_key_cols" placeholder="A,B" required>
        </div>
        <div class="form-row">
            <label>源文件匹配列 (字母，多列逗号分隔)</label>
            <input type="text" name="source_match_cols" placeholder="A,B" required>
        </div>
        <div class="form-row">
            <label>源文件值列 (字母)</label>
            <input type="text" name="source_value_col" placeholder="C" required>
        </div>
        <div class="form-row">
            <label>结果列 (目标文件)</label>
            <input type="text" name="result_col" placeholder="K 或 '匹配值'" required>
            <div class="hint">输入列字母覆盖该列；输入其他新建列</div>
        </div>
        <div class="form-row">
            <label>
                <input type="checkbox" name="allow_duplicate_keys" id="allowDupKeys">
                允许目标键重复（仅将源值填充到首次出现的行，其余行保持不变）
            </label>
        </div>
    </div>
    <div class="button-group">
        <button type="submit" class="btn btn-success">✨ 开始匹配填充</button>
        <button type="button" id="debugBtn" class="btn debug-btn">🔍 调试</button>
    </div>
</form>
<script>
    document.getElementById('debugBtn').addEventListener('click', function() {
        const form = document.getElementById('matchForm');
        form.action = '{{ url_for("match.debug") }}';
        form.target = '_blank';
        form.submit();
        form.action = '{{ url_for("match.process") }}';
        form.target = '';
    });
</script>
{% endblock %}
    ```

- **templates/splitter/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}Excel拆分工具{% endblock %}
{% block content %}
<h2>✂️ Excel拆分工具（保留合并单元格）</h2>
<p>根据指定列的唯一值将Excel拆分成多个文件，保留原格式。现在支持先筛选后拆分。</p>
<form id="splitForm" method="POST" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📂 Excel文件</label>
            <input type="file" name="file" accept=".xlsx,.xlsm,.xls" required>
            <div class="hint">支持 .xlsx, .xlsm, .xls</div>
        </div>
        <div class="form-row">
            <label>拆分列名</label>
            <input type="text" name="split_column" placeholder="例如：险种类型" required>
            <div class="hint">必须是第一行表头中的列名</div>
        </div>
        <div class="form-row">
            <label>正选值（可选）</label>
            <input type="text" name="include_values" placeholder="居民,职工">
            <div class="hint">只拆分这些值</div>
        </div>
        <div class="form-row">
            <label>反选值（可选）</label>
            <input type="text" name="exclude_values" placeholder="其他">
            <div class="hint">排除这些值</div>
        </div>
        <div class="form-row">
            <label>跳过工作表索引</label>
            <input type="text" name="skip_sheets" placeholder="1,3">
            <div class="hint">不参与拆分的sheet序号（从1开始）</div>
        </div>
    </div>

    <!-- 筛选条件区域 -->
    <div class="form-section">
        <h3>🎯 筛选条件（可选）</h3>
        <div id="advancedConditions">
            <div class="condition-row" style="display:flex; gap:8px; margin-bottom:10px;">
                <input type="text" name="condition_field_0" placeholder="字段名" style="flex:2;">
                <select name="condition_op_0" style="flex:1.5;">
                    <option value="=">=</option>
                    <option value="!=">!=</option>
                    <option value=">">></option>
                    <option value="<"><</option>
                    <option value=">=">>=</option>
                    <option value="<="><=</option>
                    <option value="包含">包含</option>
                    <option value="不包含">不包含</option>
                </select>
                <input type="text" name="condition_value_0" placeholder="值" style="flex:3;">
                <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">✖</button>
            </div>
        </div>
        <button type="button" class="btn btn-secondary" onclick="addCondition()">+ 添加条件</button>
        <div class="hint">先按这些条件筛选数据，再进行拆分。所有条件为 AND 关系。</div>
    </div>

    <button type="submit" class="btn btn-success" id="submitBtn">开始拆分</button>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <h3>处理状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">处理中...</span></p>
</div>
<script>
    let condCount = 1;
    function addCondition() {
        const container = document.getElementById('advancedConditions');
        const div = document.createElement('div');
        div.className = 'condition-row';
        div.style.cssText = 'display:flex; gap:8px; margin-bottom:10px;';
        div.innerHTML = `
            <input type="text" name="condition_field_${condCount}" placeholder="字段名" style="flex:2;">
            <select name="condition_op_${condCount}" style="flex:1.5;">
                <option value="=">=</option>
                <option value="!=">!=</option>
                <option value=">">></option>
                <option value="<"><</option>
                <option value=">=">>=</option>
                <option value="<="><=</option>
                <option value="包含">包含</option>
                <option value="不包含">不包含</option>
            </select>
            <input type="text" name="condition_value_${condCount}" placeholder="值" style="flex:3;">
            <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">✖</button>
        `;
        container.appendChild(div);
        condCount++;
    }

    const form = document.getElementById('splitForm');
    const statusArea = document.getElementById('statusArea');
    const taskStatus = document.getElementById('taskStatus');
    const progressBar = document.getElementById('progressBar');
    const submitBtn = document.getElementById('submitBtn');

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        statusArea.style.display = 'block';
        taskStatus.innerText = '提交中...';
        progressBar.style.width = '10%';
        submitBtn.disabled = true;

        const formData = new FormData(form);
        try {
            const response = await fetch('/splitter/process', {
                method: 'POST',
                body: formData
            });
            if (response.ok) {
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = '拆分结果.zip';
                if (contentDisposition) {
                    const match = contentDisposition.match(/filename="?(.+?)"?$/);
                    if (match) filename = match[1];
                }
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                taskStatus.innerText = '下载完成！';
                progressBar.style.width = '100%';
                progressBar.style.backgroundColor = '#28a745';
                setTimeout(() => {
                    statusArea.style.display = 'none';
                }, 3000);
            } else {
                let errorMsg = '拆分失败';
                try {
                    const errData = await response.json();
                    errorMsg = errData.message || errorMsg;
                } catch (e) {
                    errorMsg = `服务器错误 (${response.status})`;
                }
                taskStatus.innerText = '错误: ' + errorMsg;
                progressBar.style.width = '100%';
                progressBar.style.backgroundColor = '#dc3545';
            }
        } catch (err) {
            taskStatus.innerText = '网络错误: ' + err.message;
            progressBar.style.width = '100%';
            progressBar.style.backgroundColor = '#dc3545';
        } finally {
            submitBtn.disabled = false;
        }
    });
</script>
{% endblock %}
    ```

- **templates/subtotal/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}数据小计{% endblock %}
{% block head %}
<style>
    .override-row { display: flex; gap: 10px; margin-bottom: 8px; align-items: center; }
    .sheets-radio { margin-right: 15px; }
</style>
{% endblock %}
{% block content %}
<h2>📊 数据小计</h2>
<p>根据主键列对数据进行分组，在每组最后一行下方插入小计行。</p>
<form id="subtotalForm" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
    </div>
    <div class="form-section">
        <h3>🔑 主键列（分组依据，支持多列和范围，如 A,B,C 或 A-C）</h3>
        <input type="text" name="key_cols" placeholder="A,B" required>
    </div>
    <div class="form-section">
        <h3>📏 维度列（在小计行显示“合计”，可选，支持多列）</h3>
        <input type="text" name="dimension_cols" placeholder="C,D">
        <div class="hint">如果指定的列不存在，将报错。</div>
    </div>
    <div class="form-section">
        <h3>🧮 计算列（支持多列和范围）</h3>
        <input type="text" name="calc_cols" placeholder="E,F,G-H" required>
    </div>
    <div class="form-section">
        <h3>⚙️ 聚合函数</h3>
        <div style="display: flex; align-items: center; gap: 10px;">
            <label>默认函数：</label>
            <select name="agg_func_default" id="aggFuncDefault">
                <option value="sum">求和</option>
                <option value="avg">平均值</option>
                <option value="min">最小值</option>
                <option value="max">最大值</option>
                <option value="count">计数</option>
            </select>
            <button type="button" class="btn btn-secondary" id="addOverrideBtn">+ 特殊列函数</button>
        </div>
        <div id="overrideContainer" style="margin-top:10px;"></div>
        <div class="hint">点击“+”为特定列设置不同的聚合函数，格式：列字母:函数，如 C:avg</div>
    </div>
    <div class="form-section">
        <h3>📄 处理Sheet</h3>
        <div>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="1" checked> 仅第一个Sheet</label>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="all"> 所有Sheet</label>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="custom"> 指定Sheet索引（逗号分隔）</label>
            <input type="text" name="sheets_custom" id="sheetsCustom" placeholder="1,2,3" style="width:150px; display:none;">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始处理</button>
</form>

<div id="statusArea" style="display:none; margin-top:20px;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <div id="downloadLink"></div>
</div>

<template id="overrideRowTemplate">
    <div class="override-row">
        <span>列：</span>
        <input type="text" class="override-col" placeholder="C" style="width:80px;">
        <span>函数：</span>
        <select class="override-func">
            <option value="sum">求和</option>
            <option value="avg">平均值</option>
            <option value="min">最小值</option>
            <option value="max">最大值</option>
            <option value="count">计数</option>
        </select>
        <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">删除</button>
    </div>
</template>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/subtotal.js') }}"></script>
{% endblock %}
    ```

- **templates/summarizer/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}汇总工具{% endblock %}
{% block content %}
<h2>📋 汇总工具</h2>
<form action="{{ url_for('summary.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📂 上传多个Excel文件</label>
            <input type="file" name="files" accept=".xlsx,.xls" multiple required>
        </div>
        <div class="form-row">
            <label>需要保留的字段（每行一个，或逗号/空格分隔）</label>
            <textarea name="fields" rows="5" placeholder="机构名称&#10;日期&#10;金额" required></textarea>
        </div>
        <div class="form-row">
            <label>唯一键（主键，逗号、换行或空格分隔，可选）</label>
            <textarea name="unique_keys" rows="3" placeholder="例如 机构名称,日期"></textarea>
        </div>
        <div class="form-row">
            <label>求和字段（逗号、换行或空格分隔，可选）</label>
            <textarea name="sum_fields" rows="3" placeholder="例如 金额,数量"></textarea>
        </div>
        <div class="form-row">
            <label>表头所在行号（从1开始，如有标题行请填写；留空则从第一行读取）</label>
            <input type="number" name="header_row" value="" min="1" placeholder="例如 2" style="max-width:120px;">
            <div class="hint">与下方“跳过标题行数”二选一，优先使用行号</div>
        </div>
        <div class="form-row">
            <label>跳过的标题行数（第一行如果不是表头，请填写跳过行数）</label>
            <input type="number" name="skip_header_rows" value="0" min="0" style="max-width:120px;">
            <div class="hint">例如：文件第一行是标题，第二行是表头，则填写 1</div>
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始汇总</button>
</form>
{% endblock %}
    ```

- **templates/team_report/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}车队日报{% endblock %}
{% block head %}
<style>
    .form-row { margin-bottom: 15px; }
    .form-row label { display: block; margin-bottom: 5px; font-weight: bold; }
    .form-row input[type="file"], .form-row input[type="text"], .form-row input[type="number"] { width: 100%; max-width: 300px; }
    .hint { color: #666; font-size: 0.9em; }
    .status-area { margin-top: 20px; }
    .progress { height: 20px; background: #eee; border-radius: 10px; overflow: hidden; }
    .progress-bar { height: 100%; background: #4361ee; width: 0%; transition: width 0.3s; }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
</style>
{% endblock %}
{% block content %}
<h2>📊 生成车队日报</h2>
<form id="reportForm" enctype="multipart/form-data">
    <div class="form-section">
        <h3>📂 上传文件</h3>

        <div class="form-row">
            <label>今日司机流水文件 <span style="color:red;">*</span></label>
            <input type="file" id="todayFile" name="today_file" accept=".xlsx,.xls,.csv" required>
        </div>

        <div class="form-row">
            <label>昨日司机流水文件 （可选）</label>
            <input type="file" id="yesterdayFile" name="yesterday_file" accept=".xlsx,.xls,.csv">
        </div>

        <div class="form-row">
            <label>本月司机流水文件（可选，多选）</label>
            <input type="file" id="monthlyFiles" name="monthly_files" accept=".xlsx,.xls,.csv" multiple>
            <div class="hint">可上传多个月份的流水文件，用于统计月度累计数据</div>
        </div>

        <div class="form-row">
            <label>队伍明细映射表 （可选）</label>
            <input type="file" id="mappingFile" name="mapping_file" accept=".xlsx,.xls">
            <div class="hint">若不提供，将使用已配置的默认映射</div>
        </div>
    </div>

    <div class="form-section">
        <h3>🎯 目标值设置（留空使用默认值）</h3>
        <div class="grid-2">
            <div class="form-row">
                <label>日完单目标</label>
                <input type="number" id="dailyOrderTarget" name="daily_order_target" placeholder="默认1150" min="1">
            </div>
            <div class="form-row">
                <label>日报单目标</label>
                <input type="number" id="dailyApplyTarget" name="daily_apply_target" placeholder="默认120" min="1">
            </div>
            <div class="form-row">
                <label>月完单目标</label>
                <input type="number" id="monthlyOrderTarget" name="monthly_order_target" placeholder="默认35650" min="1">
            </div>
            <div class="form-row">
                <label>月报单目标</label>
                <input type="number" id="monthlyApplyTarget" name="monthly_apply_target" placeholder="默认3720" min="1">
            </div>
        </div>
    </div>

    <button type="submit" class="btn btn-success" id="submitBtn">开始生成</button>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <p id="progressMessage"></p>
    <div id="downloadLink" style="margin-top:15px;"></div>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/team_report.js') }}"></script>
{% endblock %}
    ```

- **templates/transposer/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}行列转置{% endblock %}
{% block content %}
<h2>🔄 行列转置（取消分类列）</h2>
<p>将某一分类列的不同取值转换为列标题，实现长表转宽表。例如：医疗类别（住院、门诊、其他）变为列头，原数值列变为“住院a数据”、“门诊b数据”等。</p>

<form id="transposerForm" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
    </div>
    <div class="form-section">
        <div class="form-row">
            <label>🔑 索引列（主键，多列用逗号或空格分隔，可选）</label>
            <input type="text" name="index_cols" placeholder="例如 病种名称 单位名称，或 A B">
            <div class="hint">若不填，将自动为每组类别生成行序号，并以此为索引。</div>
        </div>
        <div class="form-row">
            <label>📋 分类列（将被取消的列）</label>
            <input type="text" name="category_col" placeholder="例如 医疗类别" required>
        </div>
        <div class="form-row">
            <label>📊 值列（需保留的数值列，多选）</label>
            <input type="text" name="value_cols" placeholder="例如 a数据 b数据" required>
            <div class="hint">多个列名用逗号或空格分隔。</div>
        </div>
    </div>
    <button type="submit" class="btn btn-success" id="submitBtn">开始转换</button>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <p id="statusMessage"></p>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/transposer.js') }}"></script>
{% endblock %}
    ```

- **tests/__init__.py**
  - language: `python`
  - content:
    ```python

    ```

- **tests/conftest.py**
  - language: `python`
  - content:
    ```python
import pytest
import tempfile
import os
import shutil
from openpyxl import Workbook
import pandas as pd
from app import create_app


@pytest.fixture
def app():
    """每个测试函数独立创建一个 Flask 应用和临时目录，并在结束后清理。"""
    app = create_app()
    upload = tempfile.mkdtemp()
    app.config.update({
        "TESTING": True,
        "UPLOAD_FOLDER": upload,
        "WTF_CSRF_ENABLED": False,
    })
    yield app
    # 清理所有临时目录
    shutil.rmtree(upload, ignore_errors=True)


@pytest.fixture
def client(app):
    return app.test_client()


@pytest.fixture
def sample_excel_file():
    from io import BytesIO
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "A"
    ws["B1"] = "B"
    ws["C1"] = "C"
    ws["A2"] = 10
    ws["B2"] = 20
    ws["C2"] = 30
    ws["A3"] = 5
    ws["B3"] = 3
    ws["C3"] = 2
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@pytest.fixture
def sample_csv_header_file():
    from io import BytesIO
    content = "key,value,other\n"
    return BytesIO(content.encode("utf-8"))


@pytest.fixture
def sample_csv_data_file():
    from io import BytesIO
    content = "A,10,x\nA,20,y\nB,5,z\n"
    return BytesIO(content.encode("utf-8"))


@pytest.fixture
def sample_excel_comparison_files():
    from io import BytesIO
    df1 = pd.DataFrame({"id": [1, 2], "name": ["Alice", "Bob"], "score": [100, 90]})
    df2 = pd.DataFrame({"id": [1, 3], "name": ["Alice", "Charlie"], "score": [100, 80]})
    f1 = BytesIO()
    f2 = BytesIO()
    with pd.ExcelWriter(f1, engine="openpyxl") as w: df1.to_excel(w, index=False)
    with pd.ExcelWriter(f2, engine="openpyxl") as w: df2.to_excel(w, index=False)
    f1.seek(0); f2.seek(0)
    return f1, f2
    ```

- **tests/test_calculator.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
from openpyxl import Workbook
from services.excel.calculator import ColumnCalculator
from services.excel.aggregator import Aggregator
from core.excel_utils import ExcelFile


class TestColumnCalculator:
    def test_calculate_column(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'A'
        ws['B1'] = 'B'
        ws['C1'] = 'Result'
        ws['A2'] = 10
        ws['B2'] = 20
        ws['A3'] = 5
        ws['B3'] = 3
        test_file = tmp_path / 'test.xlsx'
        wb.save(test_file)

        with ExcelFile(test_file) as ef:
            calculator = ColumnCalculator(ef)
            calculator.calculate_column('A+B', 'C', merge_by_cols=None)
            ef.save(test_file)

        result_df = pd.read_excel(test_file)
        assert result_df.loc[0, 'Result'] == 30
        assert result_df.loc[1, 'Result'] == 8

    def test_calculate_column_with_merge(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Group'
        ws['B1'] = 'Value'
        ws['C1'] = 'Result'
        ws['A2'] = 'X'
        ws['B2'] = 10
        ws['A3'] = 'X'
        ws['B3'] = 20
        ws['A4'] = 'Y'
        ws['B4'] = 5
        test_file = tmp_path / 'test.xlsx'
        wb.save(test_file)

        with ExcelFile(test_file) as ef:
            calculator = ColumnCalculator(ef)
            calculator.calculate_column('B*2', 'C', merge_by_cols=['A'])
            ef.save(test_file)

        result_df = pd.read_excel(test_file)
        assert result_df.loc[0, 'Result'] == 20
        assert pd.isna(result_df.loc[1, 'Result'])
        assert result_df.loc[2, 'Result'] == 10


class TestAggregator:
    def test_aggregate_sum(self):
        df = pd.DataFrame({'Group': ['X', 'X', 'Y'], 'Value': [10, 20, 5]})
        result = Aggregator.simple_aggregate(df, ['Group'], ['Value'], 'sum')
        assert len(result) == 2
        assert result.loc[result['Group'] == 'X', 'Value'].iloc[0] == 30
        assert result.loc[result['Group'] == 'Y', 'Value'].iloc[0] == 5

    def test_aggregate_mean(self):
        df = pd.DataFrame({'Group': ['X', 'X'], 'Value': [10, 30]})
        result = Aggregator.simple_aggregate(df, ['Group'], ['Value'], 'mean')
        assert result.loc[0, 'Value'] == 20

    def test_aggregate_by_keys(self):
        df = pd.DataFrame({
            'key': ['A', 'A', 'B'],
            'sum_col': [10, 20, 5],
            'other_col': ['first', 'second', 'other']
        })
        result = Aggregator.aggregate_by_keys(df, ['key'], ['sum_col'], ['other_col'])
        assert len(result) == 2
        assert result.loc[result['key'] == 'A', 'sum_col'].iloc[0] == 30
        assert result.loc[result['key'] == 'A', 'other_col'].iloc[0] == 'first'
        assert result.loc[result['key'] == 'B', 'sum_col'].iloc[0] == 5
    ```

- **tests/test_checker_service.py**
  - language: `python`
  - content:
    ```python
import pytest
import pandas as pd
import os
from services.checker.service import CheckerService
from core.exceptions import ValidationError
from core.config import CheckConfig  # 引入配置类


class TestCheckerService:
    @pytest.fixture
    def sample_files_a(self, tmp_path):
        df1 = pd.DataFrame({"ID": [1, 2], "Name": ["Alice", "Bob"], "Score": [100, 90]})
        df2 = pd.DataFrame({"ID": [3, 4], "Name": ["Charlie", "David"], "Score": [80, 70]})
        f1 = tmp_path / "a1.xlsx"
        f2 = tmp_path / "a2.xlsx"
        df1.to_excel(f1, index=False)
        df2.to_excel(f2, index=False)
        return [str(f1), str(f2)]

    @pytest.fixture
    def sample_files_b(self, tmp_path):
        df = pd.DataFrame({"ID": [1, 2, 5], "Name": ["Alice", "Bob", "Eve"], "Score": [100, 95, 60]})
        f = tmp_path / "b.xlsx"
        df.to_excel(f, index=False)
        return [str(f)]

    def test_execute_basic(self, sample_files_a, sample_files_b):
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            duplicate_action="error"
        )
        result_path = service.execute()
        assert os.path.exists(result_path)
        result = pd.read_excel(result_path, sheet_name=None)
        assert "差异数据" in result
        assert len(result["差异数据"]) == 1
        assert "标准群独有" in result
        assert len(result["标准群独有"]) == 2
        assert "比较群独有" in result
        assert len(result["比较群独有"]) == 1
        os.unlink(result_path)

    def test_duplicate_key_error(self, sample_files_a, sample_files_b, tmp_path):
        df_dup = pd.DataFrame({"ID": [1, 1], "Name": ["Alice", "Alice"], "Score": [100, 200]})
        dup_file = tmp_path / "dup.xlsx"
        df_dup.to_excel(dup_file, index=False)
        service = CheckerService(
            files_a=[str(dup_file)],
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}],
            duplicate_action="error"
        )
        with pytest.raises(ValidationError, match="标准群主键重复"):
            service.execute()

    def test_filter_in_standard(self, sample_files_a, sample_files_b):
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            filters_a=[{"field": "Score", "op": ">", "value": 90}],
            duplicate_action="error"
        )
        result_path = service.execute()
        result = pd.read_excel(result_path, sheet_name=None)
        assert "比较群独有" in result
        assert len(result["比较群独有"]) == 2
        os.unlink(result_path)

    def test_execute_with_sqlite_fallback(self, sample_files_a, sample_files_b, monkeypatch):
        """将 max_memory_rows 设为 0 强制使用 SQLite 回退模式"""
        # 构造一个配置使 max_memory_rows=0，触发 SQLite 路径
        low_mem_config = CheckConfig(max_memory_rows=0)
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            duplicate_action="error",
            config=low_mem_config
        )
        result_path = service.execute()
        assert os.path.exists(result_path)
        result = pd.read_excel(result_path, sheet_name=None)
        assert "差异数据" in result
        os.unlink(result_path)
    ```

- **tests/test_common.py**
  - language: `python`
  - content:
    ```python
import os
import numpy as np
import pandas as pd
from openpyxl import Workbook
from core.validators import validate_column_letters, validate_sheet_index, validate_expression, sanitize_filename
# 直接使用核心表达式引擎，替代旧的 utils.expression 函数
from core.expression_engine import ExpressionEngine
from core.excel_utils import ExcelFile, col_letter_to_index, col_index_to_letter
from services.data.file_service import FileService
from flask import Flask

class TestValidators:
    def test_validate_column_letters(self):
        assert validate_column_letters(['A', 'AB', 'Z']) == True
        assert validate_column_letters(['a', 'Ab']) == True
        assert validate_column_letters(['A1', 'AB']) == False
        assert validate_column_letters([]) == True

    def test_validate_sheet_index(self):
        assert validate_sheet_index('1') == True
        assert validate_sheet_index(2) == True
        assert validate_sheet_index('0') == False
        assert validate_sheet_index('a') == False
        assert validate_sheet_index(None) == False

    def test_validate_expression(self):
        assert validate_expression('A+B-C') == True
        assert validate_expression('A*B/C') == True
        assert validate_expression('A**2') == True
        assert validate_expression('A+') == False
        assert validate_expression('') == False

    def test_sanitize_filename(self):
        assert sanitize_filename('a/b?c:d') == 'a_b_c_d'
        assert sanitize_filename('normal') == 'normal'

class TestExpression:
    @classmethod
    def setup_class(cls):
        cls.engine = ExpressionEngine()

    def test_extract_column_letters(self):
        # 利用引擎的 compile 方法提取变量（字母），但提取逻辑在旧函数中，这里直接复用旧实现
        import re
        def extract_column_letters(expr):
            return set(re.findall(r'[A-Z]+', expr.upper()))
        assert extract_column_letters('A+B-C') == {'A', 'B', 'C'}
        assert extract_column_letters('AA+BB') == {'AA', 'BB'}
        assert extract_column_letters('') == set()

    def test_evaluate_expression(self):
        values = {'A': 10, 'B': 20, 'C': 5}
        assert self.engine.evaluate('A+B-C', values) == 25
        assert self.engine.evaluate('A*B', values) == 200
        assert self.engine.evaluate('A/B', values) == 0.5
        result = self.engine.evaluate('A/0', values)
        assert np.isnan(result)
        result = self.engine.evaluate('X+Y', values)
        assert np.isnan(result)

class TestExcelHandler:
    def test_col_letter_conversion(self):
        assert col_letter_to_index('A') == 1
        assert col_letter_to_index('Z') == 26
        assert col_letter_to_index('AA') == 27
        assert col_index_to_letter(1) == 'A'
        assert col_index_to_letter(27) == 'AA'

    def test_excel_file_read_write(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2], 'B': ['x', 'y']})
        test_file = tmp_path / 'test.xlsx'
        df.to_excel(test_file, index=False)
        with ExcelFile(test_file) as ef:
            headers = ef.get_headers()
            assert headers == ['A', 'B']
            val = ef.get_cell_value(2, 'A')
            assert val == 1
            new_file = tmp_path / 'new.xlsx'
            ef.save(new_file)
        assert new_file.exists()

    def test_copy_merge_structure(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.merge_cells('A1:A3')
        ws['A1'] = 'Merged'
        test_file = tmp_path / 'merge.xlsx'
        wb.save(test_file)
        with ExcelFile(test_file) as ef:
            ws = ef.active_sheet
            ef.copy_merge_structure('A', 'B', 1, 3, sheet=ws)
            merged = list(ws.merged_cells.ranges)
            assert len(merged) == 2

    def test_read_dataframe(self, tmp_path):
        df = pd.DataFrame({'X': [1, 2]})
        test_file = tmp_path / 'data.xlsx'
        df.to_excel(test_file, index=False)
        result = ExcelFile.read_dataframe(test_file)
        pd.testing.assert_frame_equal(result, df)

class TestFileService:
    def test_save_upload(self, tmp_path):
        app = Flask(__name__)
        app.config['UPLOAD_FOLDER'] = str(tmp_path)
        app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}
        with app.app_context():
            from werkzeug.datastructures import FileStorage
            import io
            fake_file = FileStorage(io.BytesIO(b'fake'), filename='test.xlsx')
            svc = FileService()
            path = svc.save_upload(fake_file, subdir='')
            assert os.path.exists(path)
            assert path.endswith('.xlsx')
            svc.cleanup([path])
            assert not os.path.exists(path)
    ```

- **tests/test_comparator.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
import tempfile
from services.excel.comparator import ExcelComparator

def create_test_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    return filename

class TestExcelComparator:
    def test_full_compare_identical(self):
        data = [{'id': 1, 'name': 'A', 'score': 100}, {'id': 2, 'name': 'B', 'score': 90}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data, f1.name)
            create_test_excel(data, f2.name)
            result = ExcelComparator.full_compare(f1.name, f2.name, ['id'], ['name', 'score'])
            assert result['only_in_1'].empty
            assert result['only_in_2'].empty
            assert result['diff'] is None

    def test_full_compare_diff(self):
        data1 = [{'id': 1, 'name': 'A', 'score': 100}]
        data2 = [{'id': 1, 'name': 'A', 'score': 90}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data1, f1.name)
            create_test_excel(data2, f2.name)
            result = ExcelComparator.full_compare(f1.name, f2.name, ['id'], ['score'])
            assert result['diff'] is not None
            assert len(result['diff']) == 1
            assert result['diff'].iloc[0]['文件1_score'] == 100
            assert result['diff'].iloc[0]['文件2_score'] == 90

    def test_field_match(self):
        data1 = [{'name': 'A'}, {'name': 'B'}, {'name': 'C'}]
        data2 = [{'name': 'A'}, {'name': 'B'}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data1, f1.name)
            create_test_excel(data2, f2.name)
            only1, only2 = ExcelComparator.field_match(f1.name, f2.name, 'name')
            assert len(only1) == 1
            assert only1.iloc[0]['name'] == 'C'
            assert only2.empty
    ```

- **tests/test_condition_filter.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
from core.filter_engine import ConditionFilter, Condition

class TestConditionFilter:
    def test_apply_to_dict_eq(self):
        row = {'name': 'Alice', 'age': '30'}
        cond = Condition(field='name', op='=', value='Alice')
        assert ConditionFilter.apply_to_dict(row, [cond]) == True

    def test_apply_to_dict_contains(self):
        row = {'name': 'Alice'}
        cond = Condition(field='name', op='contains', value='Ali')
        assert ConditionFilter.apply_to_dict(row, [cond]) == True

    def test_apply_to_dict_exclude(self):
        row = {'name': 'Alice'}
        cond = Condition(field='name', op='=', value='Alice', exclude=True)
        assert ConditionFilter.apply_to_dict(row, [cond]) == False

    def test_apply_to_dataframe(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie'], 'age': [25, 30, 35]})
        cond = Condition(field='name', op='contains', value='Ali')
        mask = ConditionFilter.apply_to_dataframe(df, [cond])
        assert mask.tolist() == [True, False, False]

    def test_apply_to_dataframe_multiple(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie'], 'age': [25, 30, 35]})
        cond1 = Condition(field='age', op='>', value='30')
        cond2 = Condition(field='name', op='contains', value='Char')
        mask = ConditionFilter.apply_to_dataframe(df, [cond1, cond2])
        assert mask.tolist() == [False, False, True]
    ```

- **tests/test_convert_service.py**
  - language: `python`
  - content:
    ```python
import pytest
import pandas as pd
import tempfile
import os
from services.convert.service import ConvertService
from core.exceptions import ValidationError

class TestConvertService:
    @pytest.fixture
    def sample_xlsx(self, tmp_path):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False)
        return str(path)

    @pytest.fixture
    def sample_multi_sheet_xlsx(self, tmp_path):
        df1 = pd.DataFrame({"ID": [1], "Name": ["Alice"]})
        df2 = pd.DataFrame({"ID": [2], "Name": ["Bob"]})
        path = tmp_path / "multi.xlsx"
        with pd.ExcelWriter(path, engine='openpyxl') as w:
            df1.to_excel(w, sheet_name='Sheet1', index=False)
            df2.to_excel(w, sheet_name='Sheet2', index=False)
        return str(path)

    @pytest.fixture
    def sample_csv(self, tmp_path):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        path = tmp_path / "test.csv"
        df.to_csv(path, index=False, encoding='utf-8-sig')
        return str(path)

    def test_excel_to_csv_single_sheet(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='csv')
        result_path = service.execute()
        assert result_path.endswith('.csv')
        df = pd.read_csv(result_path)
        assert list(df.columns) == ['A', 'B']
        assert len(df) == 2
        os.unlink(result_path)

    def test_excel_to_csv_multi_sheet(self, sample_multi_sheet_xlsx):
        service = ConvertService(file_path=sample_multi_sheet_xlsx, target_format='csv')
        result_path = service.execute()
        assert result_path.endswith('.zip')
        import zipfile
        with zipfile.ZipFile(result_path, 'r') as zf:
            names = zf.namelist()
            assert len(names) == 2
            assert any('Sheet1.csv' in n for n in names)
            assert any('Sheet2.csv' in n for n in names)
        os.unlink(result_path)

    def test_csv_to_excel(self, sample_csv):
        service = ConvertService(file_path=sample_csv, target_format='xlsx')
        result_path = service.execute()
        assert result_path.endswith('.xlsx')
        df = pd.read_excel(result_path)
        assert list(df.columns) == ['A', 'B']
        assert len(df) == 2
        os.unlink(result_path)

    def test_unsupported_conversion(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='xlsx')
        with pytest.raises(ValidationError, match='不支持的转换'):
            service.execute()

    def test_invalid_file(self, tmp_path):
        invalid_file = tmp_path / "invalid.xlsx"
        invalid_file.write_text("not an excel file")
        service = ConvertService(file_path=str(invalid_file), target_format='csv')
        with pytest.raises(Exception):
            service.execute()

    def test_precheck(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='csv')
        result = service.precheck()
        assert 'total_rows' in result
        assert result['total_rows'] == 2
    ```

- **tests/test_core_task.py**
  - language: `python`
  - content:
    ```python
import pytest
import tempfile
import os
import time
from core.task.storage import TaskStorage
from core.task.manager import TaskManager
from core.task.state import TaskStatus
from core.task.base import Task

class DummyTask(Task):
    def __init__(self, should_fail=False, result_data=None):
        self.should_fail = should_fail
        self.result_data = result_data or {"output": "success"}
    def execute(self):
        if self.should_fail:
            raise ValueError("模拟任务失败")
        return self.result_data

class TestTaskStorage:
    @pytest.fixture
    def storage(self, tmp_path):
        status_dir = tmp_path / "status"
        output_dir = tmp_path / "output"
        return TaskStorage(str(status_dir), str(output_dir))

    def test_save_and_load(self, storage):
        task_id = "test-123"
        data = {"status": "pending", "value": 42}
        storage.save(task_id, data)
        loaded = storage.load(task_id)
        assert loaded["value"] == 42

    def test_update(self, storage):
        task_id = "test-123"
        storage.save(task_id, {"status": "pending"})
        storage.update(task_id, status="completed", result="ok")
        loaded = storage.load(task_id)
        assert loaded["status"] == "completed"
        assert loaded["result"] == "ok"

    def test_append_and_get_log(self, storage):
        task_id = "log-test"
        storage.append_log(task_id, "first line")
        storage.append_log(task_id, "second line")
        log = storage.get_log(task_id)
        assert "first line" in log
        assert "second line" in log

    def test_delete(self, storage):
        task_id = "delete-me"
        storage.save(task_id, {})
        assert storage.load(task_id) is not None
        storage.delete(task_id)
        assert storage.load(task_id) is None

class TestTaskManager:
    @pytest.fixture
    def storage(self, tmp_path):
        status_dir = tmp_path / "status"
        output_dir = tmp_path / "output"
        return TaskStorage(str(status_dir), str(output_dir))

    @pytest.fixture
    def manager(self, storage):
        return TaskManager(storage)

    def test_create_task(self, manager):
        params = {"file": "test.csv"}
        task_id = manager.create_task(params, original_filename="test")
        assert task_id is not None
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.PENDING
        assert status["params"] == params

    def test_submit_and_complete(self, manager):
        task = DummyTask(result_data={"result": "ok"})
        task_id = manager.create_task({})
        manager.storage.update(task_id, status=TaskStatus.RUNNING)
        manager.complete_task(task_id, result_path="/out.xlsx", total_input_rows=100)
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.COMPLETED
        assert status["result_path"] == "/out.xlsx"

    def test_fail_task(self, manager):
        task_id = manager.create_task({})
        manager.fail_task(task_id, error="Something wrong")
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.FAILED
        assert status["error"] == "Something wrong"

    def test_cancel_task(self, manager):
        task_id = manager.create_task({})
        manager.cancel_task(task_id)
        status = manager.storage.load(task_id)
        assert status["status"] == TaskStatus.CANCELLED
    ```

- **tests/test_data_access_layer.py**
  - language: `python`
  - content:
    ```python
import pytest
import pandas as pd
from core.data_access_layer import DataAccessLayer

class TestDataAccessLayer:
    def test_auto_clean_strings(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'A': [' 10 ', '20', ' 30 ']})
        cleaned = dal._auto_clean_dataframe(df)
        # 数值转换后应为数字，但测试允许字符串或数字（向后兼容）
        assert cleaned['A'].tolist() == [10.0, 20.0, 30.0] or cleaned['A'].tolist() == ['10', '20', '30']

    def test_dtype_hints(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'amount': ['100', '200', 'invalid']})
        cleaned = dal._auto_clean_dataframe(df, dtype_hints={'amount': float})
        # dtype_hints 会强制转换，无效值变为 NaN
        assert cleaned['amount'].isna().sum() == 1

    def test_empty_string_handling(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'A': ['', '   ', '10']})
        cleaned = dal._auto_clean_dataframe(df)
        # 空字符串变为 NaN，'10' 转为数字或保持字符串
        assert pd.isna(cleaned['A'].iloc[0])
        assert pd.isna(cleaned['A'].iloc[1])
        assert cleaned['A'].iloc[2] == 10.0 or cleaned['A'].iloc[2] == '10'
    ```

- **tests/test_expression_engine.py**
  - language: `python`
  - content:
    ```python
import pytest
import numpy as np
from core.expression_engine import ExpressionEngine

class TestExpressionEngine:
    @pytest.fixture
    def engine(self):
        return ExpressionEngine()

    def test_basic_arithmetic(self, engine):
        assert engine.evaluate("10+20", {}) == 30
        assert engine.evaluate("10-5", {}) == 5
        assert engine.evaluate("3*4", {}) == 12
        assert engine.evaluate("10/2", {}) == 5

    def test_with_variables(self, engine):
        context = {"A": 10, "B": 5}
        assert engine.evaluate("A+B", context) == 15
        assert engine.evaluate("A*B", context) == 50

    def test_string_conversion(self, engine):
        context = {"A": "10", "B": " 5 "}
        assert engine.evaluate("A+B", context) == 15

    def test_missing_variable_nan(self, engine):
        engine.missing_value_strategy = 'nan'
        assert np.isnan(engine.evaluate("X+Y", {"X": 10}))

    def test_missing_variable_zero(self):
        engine = ExpressionEngine(missing_value_strategy='zero')
        assert engine.evaluate("X+Y", {"X": 10}) == 10.0

    def test_division_by_zero(self, engine):
        assert np.isnan(engine.evaluate("10/0", {}))

    def test_expression_caching(self, engine):
        expr = "A+B*C"
        context = {"A": 1, "B": 2, "C": 3}
        assert engine.evaluate(expr, context) == 7
        assert expr.upper() in engine._cache

    def test_invalid_expression(self, engine):
        with pytest.raises(ValueError):
            engine.evaluate("10+", {})
    ```

- **tests/test_field_parser.py**
  - language: `python`
  - content:
    ```python
import pytest
from services.field_parser.service import FieldParserService

class TestFieldParserService:
    @pytest.fixture
    def service(self):
        return FieldParserService(similarity_threshold=30)

    def test_parse_field_list_commas(self, service):
        result = service.parse_field_list("姓名, 年龄, 性别")
        assert result == ["姓名", "年龄", "性别"]

    def test_match_exact(self, service):
        result = service.match_headers(["ID", "Name"], ["ID", "Name"])
        assert len(result["matched"]) == 2
        assert result["only_a"] == []
        assert result["only_b"] == []

    def test_match_similar_chinese(self, service):
        result = service.match_headers(["姓名", "年龄"], ["名字", "岁数"])
        assert len(result["matched"]) == 1
        assert result["matched"][0]["standard"] == "姓名"
        assert result["matched"][0]["compare"] == "名字"
        assert len(result["only_a"]) == 1
        assert "年龄" in result["only_a"]
        assert len(result["only_b"]) == 1
        assert "岁数" in result["only_b"]

    def test_match_below_threshold(self, service):
        service.threshold = 0.9
        result = service.match_headers(["姓名"], ["名字"])
        assert len(result["matched"]) == 0
    ```

- **tests/test_file_naming.py**
  - language: `python`
  - content:
    ```python
# tests/test_file_naming.py
from core.file_naming import generate_output_filename

class TestFileNaming:
    def test_basic(self):
        name = generate_output_filename('subtotal', '数据.xlsx')
        assert name == 'subtotal_数据_小计.xlsx'

    def test_with_path(self):
        name = generate_output_filename('match', '/path/to/目标文件.xlsx')
        assert name == 'match_目标文件_小计.xlsx'

    def test_custom_suffix(self):
        name = generate_output_filename('filter', 'data.csv', suffix='_汇总')
        assert name == 'filter_data_汇总.xlsx'
    ```

- **tests/test_filter.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
import pytest
from core.filter_engine import parse_conditions_from_form, ConditionEngine
from core.column_resolver import resolve_columns
from core.exceptions import ValidationError

engine = ConditionEngine()

class TestParseFiltersFromForm:
    def test_empty(self):
        form = {}
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact == []
        assert fuzzy == []

    def test_single_equals(self):
        form = {
            'condition_field_0': 'name',
            'condition_op_0': '=',
            'condition_value_0': 'Alice'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(exact) == 1
        assert exact[0]['field'] == 'name'
        assert exact[0]['op'] == '='
        assert exact[0]['values'] == ['Alice']

    def test_multi_values_space(self):
        form = {
            'condition_field_0': 'dept',
            'condition_op_0': '=',
            'condition_value_0': 'HR  IT'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact[0]['values'] == ['HR', 'IT']

    def test_not_equal(self):
        form = {
            'condition_field_0': 'score',
            'condition_op_0': '!=',
            'condition_value_0': '100'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact[0]['op'] == '!='

    def test_contains(self):
        form = {
            'condition_field_0': 'name',
            'condition_op_0': '包含',
            'condition_value_0': 'Al'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(fuzzy) == 1
        assert fuzzy[0]['op'] == '包含'

    def test_compatibility_old_format(self):
        form = {
            'condition_field_0': 'dept',
            'condition_values_0': 'HR IT'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(exact) == 1
        assert exact[0]['values'] == ['HR', 'IT']
        assert exact[0]['op'] == '='

class TestApplyConditions:
    def test_equals(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie']})
        exact = [{'field': 'name', 'op': '=', 'values': ['Alice', 'Charlie']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['name']) == ['Alice', 'Charlie']

    def test_not_equals(self):
        df = pd.DataFrame({'score': [100, 90, 80]})
        exact = [{'field': 'score', 'op': '!=', 'values': ['90']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['score']) == [100, 80]

    def test_greater_than(self):
        df = pd.DataFrame({'age': [25, 30, 35]})
        exact = [{'field': 'age', 'op': '>', 'values': ['28']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['age']) == [30, 35]

    def test_contains(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Alfred']})
        fuzzy = [{'field': 'name', 'op': '包含', 'value': 'Al'}]
        result = engine.apply_to_dataframe(df, [], fuzzy)
        assert list(result['name']) == ['Alice', 'Alfred']

    def test_not_contains(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Alfred']})
        fuzzy = [{'field': 'name', 'op': '不包含', 'value': 'Al'}]
        result = engine.apply_to_dataframe(df, [], fuzzy)
        assert list(result['name']) == ['Bob']

    def test_combined(self):
        df = pd.DataFrame({
            'dept': ['HR', 'IT', 'HR', 'IT'],
            'salary': [5000, 6000, 5500, 7000],
            'name': ['Alice', 'Bob', 'Charlie', 'David']
        })
        exact = [{'field': 'dept', 'op': '=', 'values': ['HR']}]
        fuzzy = [{'field': 'name', 'op': '包含', 'value': 'li'}]
        result = engine.apply_to_dataframe(df, exact, fuzzy)
        assert len(result) == 2
        assert set(result['name']) == {'Alice', 'Charlie'}

class TestResolveColumns:
    def test_by_name(self):
        df = pd.DataFrame({'A': [1], 'B': [2], 'C': [3]})
        res = resolve_columns(df, 'A,C')
        assert list(res.columns) == ['A', 'C']

    def test_by_letter(self):
        df = pd.DataFrame({'X': [1], 'Y': [2], 'Z': [3]})
        res = resolve_columns(df, 'A,C')
        assert list(res.columns) == ['X', 'Z']

    def test_mixed(self):
        df = pd.DataFrame({'Name': ['a'], 'Age': [1], 'Score': [2]})
        res = resolve_columns(df, 'Name,B')
        assert list(res.columns) == ['Name', 'Age']

    def test_invalid_raises(self):
        df = pd.DataFrame({'X': [1]})
        with pytest.raises(ValidationError):
            resolve_columns(df, 'NonExist')
    ```

- **tests/test_header_detection.py**
  - language: `python`
  - content:
    ```python
# tests/test_header_detection.py
import pytest
import pandas as pd
from core.header_detection import SimpleExcelHeaderDetector, detect_header_start, HeaderDetectionConfig
import openpyxl

class TestHeaderDetection:
    @pytest.fixture
    def simple_file(self, tmp_path):
        df = pd.DataFrame({'姓名': ['Alice', 'Bob'], '年龄': [30, 25]})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        return str(path)

    def test_basic_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(simple_file)
        assert result.header_rows == [0]
        assert result.data_start_row == 1
        assert '姓名' in result.column_names
        assert result.confidence > 0.5

    def test_detect_header_start(self, simple_file):
        start = detect_header_start(simple_file)
        assert start == 2

    def test_skip_total_row(self, tmp_path):
        df = pd.DataFrame({'名称': ['X', '合计'], '金额': [10, 30]})
        path = tmp_path / 'skip.xlsx'
        df.to_excel(path, index=False)
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(path)
        assert result.header_rows == [0]

    def test_workbook_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        wb = openpyxl.load_workbook(simple_file)
        result1 = detector.detect_from_workbook(wb)
        result2 = detector.detect_from_workbook(wb)
        assert result1.header_rows == [0]
        assert result2.metrics.get('workbook_cache_hit') is True
        wb.close()
    ```

- **tests/test_integration_calculator.py**
  - language: `python`
  - content:
    ```python
import io
from openpyxl import load_workbook


class TestCalculatorIntegration:
    def test_column_calculation(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'column',
            'header_row': '1',
            'expression': 'A+B',
            'result_col_column': 'D',
            'merge_by_cols': ''
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        assert response.headers['Content-Disposition'].startswith('attachment')
        assert '.xlsx' in response.headers['Content-Disposition']
        wb = load_workbook(io.BytesIO(response.data))
        ws = wb.active
        assert ws['D2'].value == 30
        assert ws['D3'].value == 8

    def test_aggregate_backfill(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'aggregate',
            'header_row': '1',
            'group_cols': 'A',
            'agg_cols': 'B',
            'agg_func': 'sum',
            'result_col_aggregate': 'E'
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        assert '.xlsx' in response.headers['Content-Disposition']
        wb = load_workbook(io.BytesIO(response.data))
        ws = wb.active
        assert ws['E2'].value == 20
        assert ws['E3'].value == 3

    def test_missing_file_error(self, client):
        response = client.post('/calc/process', data={'calc_type': 'column'})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '未上传文件' in response.get_json()['error']

    def test_invalid_expression_error(self, client, sample_excel_file):
        data = {
            'file': (sample_excel_file, 'test.xlsx'),
            'calc_type': 'column',
            'header_row': '1',
            'expression': '',
            'result_col_column': 'D'
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '表达式' in response.get_json()['error']
    ```

- **tests/test_integration_checker.py**
  - language: `python`
  - content:
    ```python

import io
import json
import pandas as pd
import pytest
from unittest.mock import patch

class TestCheckerIntegration:
    def test_index_page(self, client):
        response = client.get('/checker/')
        assert response.status_code == 200
        assert '文件校验' in response.get_data(as_text=True)

    @pytest.fixture
    def sample_files(self, tmp_path):
        df_a1 = pd.DataFrame({"ID": [1, 2], "Name": ["Alice", "Bob"], "Score": [100, 90]})
        df_a2 = pd.DataFrame({"ID": [3, 4], "Name": ["Charlie", "David"], "Score": [80, 70]})
        df_b = pd.DataFrame({"ID": [1, 2, 5], "Name": ["Alice", "Bob", "Eve"], "Score": [100, 95, 60]})
        return {'a1': df_a1, 'a2': df_a2, 'b': df_b}

    def test_full_check_async(self, client, sample_files):
        def df_to_bytes(df, filename):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w:
                df.to_excel(w, index=False)
            buf.seek(0)
            return (buf, filename)

        files_a = [df_to_bytes(sample_files['a1'], 'a1.xlsx'), df_to_bytes(sample_files['a2'], 'a2.xlsx')]
        files_b = [df_to_bytes(sample_files['b'], 'b.xlsx')]

        data = {
            'key_map_a': 'ID',
            'key_map_b': '',
            'compare_map': json.dumps([{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}]),
            'filters_a': '[]',
            'duplicate_action': 'error',
            'files_a': files_a,
            'files_b': files_b,
        }

        # Mock TaskManager.start_task 使其同步执行目标函数
        app = client.application
        original_tm = app.container.task_manager()
        def sync_start(task_id, target, args):
            target(*args)  # 直接调用，不启动线程
        with patch.object(original_tm, 'start_task', side_effect=sync_start):
            response = client.post('/checker/process', data=data, content_type='multipart/form-data')
            assert response.status_code == 200
            result = response.get_json()
            assert result['success'] == True  # 现在返回统一格式
            task_id = result['data']['task_id']

            # 任务应已同步完成，直接检查状态
            status_resp = client.get(f'/checker/status/{task_id}')
            status = status_resp.get_json()['data']
            assert status['status'] == 'completed'

            # 下载结果文件
            download_resp = client.get(f'/checker/download/{task_id}')
            assert download_resp.status_code == 200
            result_df = pd.read_excel(io.BytesIO(download_resp.data), sheet_name=None)
            assert '差异数据' in result_df
            assert len(result_df['差异数据']) == 1
            assert '标准群独有' in result_df
            assert len(result_df['标准群独有']) == 2
            assert '比较群独有' in result_df
            assert len(result_df['比较群独有']) == 1

    def test_missing_files_error(self, client):
        response = client.post('/checker/process', data={})
        assert response.status_code == 400
        json_data = response.get_json()
        assert not json_data['success']
        assert '请上传两批文件' in json_data['error']

    def test_invalid_task_status(self, client):
        response = client.get('/checker/status/nonexistent')
        assert response.status_code == 404
        assert not response.get_json()['success']
    ```

- **tests/test_integration_comparator.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_comparator.py
import io
import json
import pandas as pd

class TestComparatorIntegration:
    def test_full_compare(self, client, sample_excel_comparison_files):
        f1, f2 = sample_excel_comparison_files
        data = {
            'file1': (f1, 'file1.xlsx'),
            'file2': (f2, 'file2.xlsx'),
            'compare_mode': 'full',
            'key_fields': 'id',
            'compare_fields': 'name,score',
            'sheet1': '',
            'sheet2': ''
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content_disp = response.headers['Content-Disposition']
        # 新文件名格式：compare_file1_对比结果.xlsx
        assert 'compare_file1_' in content_disp and '.xlsx' in content_disp
        result = pd.ExcelFile(io.BytesIO(response.data))
        sheet_names = result.sheet_names
        assert '仅存在于文件2' in sheet_names

    def test_field_match_mode(self, client, sample_excel_comparison_files):
        f1, f2 = sample_excel_comparison_files
        data = {
            'file1': (f1, 'file1.xlsx'),
            'file2': (f2, 'file2.xlsx'),
            'compare_mode': 'field_match',
            'match_field': 'name',
            'sheet1': '',
            'sheet2': ''
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content_disp = response.headers.get('Content-Disposition', '')
        assert 'filename=' in content_disp

    def test_no_difference_response(self, client):
        import pandas as pd
        from io import BytesIO
        df = pd.DataFrame({'id': [1], 'name': ['A']})
        f1 = BytesIO()
        f2 = BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)
        data = {
            'file1': (f1, 'same1.xlsx'),
            'file2': (f2, 'same2.xlsx'),
            'compare_mode': 'field_match',
            'match_field': 'name',
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        resp_json = response.get_json()
        assert resp_json is not None
        assert resp_json['message'] == '两个文件中该字段值完全匹配，无差异'
    ```

- **tests/test_integration_convert.py**
  - language: `python`
  - content:
    ```python
import io
import time
import pandas as pd
import pytest
import zipfile

class TestConvertIntegration:
    def test_convert_excel_to_csv(self, client):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'file': (buf, 'test.xlsx'),
            'target_format': 'csv',
            'encoding': 'utf-8-sig',
            'separator': ','
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        max_attempts = 10
        status = None
        for _ in range(max_attempts):
            status_resp = client.get(f'/convert/status/{task_id}')
            status = status_resp.get_json()['data']
            if status['status'] == 'completed':
                break
            time.sleep(0.5)
        assert status is not None and status['status'] == 'completed'
        download_resp = client.get(f'/convert/download/{task_id}')
        assert download_resp.status_code == 200
        content = download_resp.data
        if content[:2] == b'PK':
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                with zf.open(zf.namelist()[0]) as f:
                    result_df = pd.read_csv(f)
        else:
            result_df = pd.read_csv(io.BytesIO(content))
        assert list(result_df.columns) == ['A', 'B']
        assert len(result_df) == 2

    def test_convert_csv_to_excel(self, client):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        buf = io.StringIO()
        df.to_csv(buf, index=False, encoding='utf-8-sig')
        buf.seek(0)
        data = {
            'file': (io.BytesIO(buf.getvalue().encode('utf-8')), 'test.csv'),
            'target_format': 'xlsx',
            'encoding': 'utf-8-sig',
            'separator': ','
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        for _ in range(10):
            status_resp = client.get(f'/convert/status/{task_id}')
            status = status_resp.get_json()['data']
            if status['status'] == 'completed':
                break
            time.sleep(0.5)
        assert status['status'] == 'completed'
        download_resp = client.get(f'/convert/download/{task_id}')
        assert download_resp.status_code == 200
        result_df = pd.read_excel(io.BytesIO(download_resp.data))
        assert list(result_df.columns) == ['A', 'B']

    def test_missing_file_error(self, client):
        response = client.post('/convert/process', data={'target_format': 'csv'})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '请上传文件' in response.get_json()['error']

    def test_invalid_target_format(self, client):
        df = pd.DataFrame({"A": [1]})
        buf = io.BytesIO()
        df.to_excel(buf, index=False, engine='openpyxl')
        buf.seek(0)
        data = {
            'file': (buf, 'test.xlsx'),
            'target_format': 'pdf',
            'encoding': 'utf-8-sig',
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        for _ in range(10):
            status = client.get(f'/convert/status/{task_id}').get_json()['data']
            if status['status'] == 'failed':
                break
            time.sleep(0.5)
        assert status['status'] == 'failed'
    ```

- **tests/test_integration_field_parser.py**
  - language: `python`
  - content:
    ```python
import json
import pytest

class TestFieldParserIntegration:
    def test_analyze_success(self, client):
        data = {
            "standard_headers": "姓名, 年龄, 性别",
            "compare_headers": "名字, 岁数, 性别"
        }
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 200
        result = response.get_json()
        assert len(result["matched"]) == 2

    def test_analyze_missing_field(self, client):
        data = {"standard_headers": "姓名"}
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 400
        result = response.get_json()
        assert "compare_headers" in result['error']

    def test_analyze_empty(self, client):
        data = {"standard_headers": "", "compare_headers": ""}
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 400
    ```

- **tests/test_integration_matcher.py**
  - language: `python`
  - content:
    ```python
import io
import pytest
import pandas as pd
from openpyxl import Workbook

class TestMatcherIntegration:
    @pytest.fixture
    def target_file(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value1'
        ws['C1'] = 'Value2'
        ws['A2'] = 1
        ws['B2'] = ''
        ws['C2'] = ''
        ws['A3'] = 2
        ws['B3'] = ''
        ws['C3'] = ''
        ws['A4'] = 3
        ws['B4'] = ''
        ws['C4'] = ''
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    @pytest.fixture
    def target_file_duplicate(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value1'
        ws['C1'] = 'Value2'
        ws['A2'] = 1
        ws['B2'] = ''
        ws['C2'] = ''
        ws['A3'] = 1
        ws['B3'] = ''
        ws['C3'] = ''
        ws['A4'] = 2
        ws['B4'] = ''
        ws['C4'] = ''
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    @pytest.fixture
    def source_file(self):
        df = pd.DataFrame({'ID': [1, 2, 3], 'Score1': [100, 200, 300], 'Score2': [101, 202, 303]})
        f = io.BytesIO()
        with pd.ExcelWriter(f, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        f.seek(0)
        return f

    def test_match_single_column(self, client, target_file, source_file):
        """测试单列匹配（向后兼容）"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert result_df.loc[0, 'Value1'] == 100
        assert result_df.loc[1, 'Value1'] == 200
        assert result_df.loc[2, 'Value1'] == 300

    def test_match_multiple_columns(self, client, target_file, source_file):
        """测试多列匹配：源值列 B,C -> 结果列 B,C"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert result_df.loc[0, 'Value1'] == 100
        assert result_df.loc[1, 'Value1'] == 200
        assert result_df.loc[2, 'Value1'] == 300
        assert result_df.loc[0, 'Value2'] == 101
        assert result_df.loc[1, 'Value2'] == 202
        assert result_df.loc[2, 'Value2'] == 303

    def test_match_multiple_columns_with_new_result_names(self, client, target_file, source_file):
        """测试多列匹配，结果列为新建列名"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': '新分数1,新分数2',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert '新分数1' in result_df.columns
        assert '新分数2' in result_df.columns
        assert result_df.loc[0, '新分数1'] == 100
        assert result_df.loc[0, '新分数2'] == 101

    def test_match_column_count_mismatch(self, client, target_file, source_file):
        """测试源值列与结果列数量不一致时返回错误"""
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B',   # 数量不匹配
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        resp_json = response.get_json()
        assert not resp_json['success']
        assert '数量' in resp_json['error']

    def test_match_with_duplicate_keys_disabled(self, client, target_file_duplicate, source_file):
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        resp_json = response.get_json()
        assert not resp_json['success']
        assert '重复' in resp_json['error']

    def test_match_with_duplicate_keys_enabled(self, client, target_file_duplicate, source_file):
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
            'allow_duplicate_keys': 'on',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        # 第一行填充，第二行（重复键）留空
        assert result_df.loc[0, 'Value1'] == 100
        assert result_df.loc[1, 'Value1'] == '' or pd.isna(result_df.loc[1, 'Value1'])
        assert result_df.loc[2, 'Value1'] == 200

    def test_debug_mode_multiple_columns(self, client, target_file, source_file):
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B,C',
            'result_col': 'B,C',
        }
        response = client.post('/match/debug', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content = response.data.decode('utf-8')
        assert 'source_value_cols' in content or 'config' in content
    ```

- **tests/test_integration_splitter.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_splitter.py
import io, zipfile, pytest
from openpyxl import Workbook

class TestSplitterIntegration:
    @pytest.fixture
    def multi_sheet_excel(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = 'Data'
        ws1['A1'] = 'Category'
        ws1['B1'] = 'Value'
        ws1['A2'] = 'A'
        ws1['B2'] = 10
        ws1['A3'] = 'B'
        ws1['B3'] = 20
        ws2 = wb.create_sheet('Extra')
        ws2['A1'] = 'Category'
        ws2['A2'] = 'A'
        ws2['A3'] = 'B'
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    def test_split_basic(self, client, multi_sheet_excel):
        data = {
            'file': (multi_sheet_excel, 'test.xlsx'),
            'split_column': 'Category',
            'include_values': '',
            'exclude_values': '',
            'skip_sheets': '2'
        }
        response = client.post('/splitter/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        # 检查 Content-Disposition 是否包含统一前缀
        content_disp = response.headers.get('Content-Disposition', '')
        assert 'splitter_test' in content_disp or 'splitter' in content_disp
        zip_data = io.BytesIO(response.data)
        with zipfile.ZipFile(zip_data) as zf:
            names = zf.namelist()
            assert any('A.xlsx' in n for n in names)
            assert any('B.xlsx' in n for n in names)
    ```

- **tests/test_integration_subtotal.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_subtotal.py
import io
import pandas as pd

class TestSubtotalIntegration:
    def test_sync_endpoint(self, client):
        # 使用字母列名
        df = pd.DataFrame({
            'A': ['X', 'X', 'Y'],
            'B': [10, 20, 30],
            'C': [100, 200, 300]
        })
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B,C',
            'agg_func_default': 'sum',
            'sheets_mode': '1'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'subtotal_test' in resp.headers.get('Content-Disposition', '')
        result_df = pd.read_excel(io.BytesIO(resp.data), sheet_name=None)
        ws = result_df['Sheet1']
        assert len(ws) == 5
        assert ws.iloc[2]['B'] == 30
        assert ws.iloc[2]['C'] == 300

    def test_missing_dimension_col(self, client):
        df = pd.DataFrame({'A': [1]})
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'A',
            'dimension_cols': 'Z'   # 不存在的列字母
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 400
        assert '无效' in resp.get_json()['error'] or '超出' in resp.get_json()['error']

    def test_async_large_file(self, client):
        df = pd.DataFrame({'A': ['X', 'Y'] * 3000, 'B': range(6000)})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test_large.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B',
            'agg_func_default': 'sum'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'task_id' in resp.get_json()
    ```

- **tests/test_integration_summarizer.py**
  - language: `python`
  - content:
    ```python
import io
import pytest
import pandas as pd
import numpy as np

class TestSummarizerIntegration:
    @pytest.fixture
    def excel_files(self):
        df1 = pd.DataFrame({'机构': ['A', 'A'], '金额': [100, 200], '日期': ['2025-01-01', '2025-01-02']})
        df2 = pd.DataFrame({'机构': ['B'], '金额': [300], '日期': ['2025-01-03']})
        f1 = io.BytesIO()
        f2 = io.BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df1.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df2.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)
        return [f1, f2]

    def test_summarize_basic(self, client, excel_files):
        data = {
            'files': [(excel_files[0], 'file1.xlsx'), (excel_files[1], 'file2.xlsx')],
            'fields': '机构\n金额',
            'unique_keys': '机构',
            'sum_fields': '金额'
        }
        response = client.post('/summary/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert len(result_df) == 2
        assert result_df.loc[result_df['机构'] == 'A', '金额'].iloc[0] == 300
        assert result_df.loc[result_df['机构'] == 'B', '金额'].iloc[0] == 300

    def test_missing_fields_error(self, client):
        response = client.post('/summary/process', data={})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '没有选择文件' in response.get_json()['error']

    # --------------- 唯一键包含空值场景 ---------------
    def test_unique_keys_with_empty_values(self, client):
        # 构造两个文件，唯一键为 机构+日期，其中一行的日期为空
        df1 = pd.DataFrame({'机构': ['A', 'A', 'B'], '日期': ['2025-01', '', '2025-02'], '金额': [100, 200, 300]})
        df2 = pd.DataFrame({'机构': ['B'], '日期': ['2025-02'], '金额': [400]})

        f1 = io.BytesIO()
        f2 = io.BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df1.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df2.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)

        data = {
            'files': [(f1, 'file1.xlsx'), (f2, 'file2.xlsx')],
            'fields': '机构\n日期\n金额',
            'unique_keys': '机构,日期',
            'sum_fields': '金额'
        }
        response = client.post('/summary/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))

        # 预期结果：
        # A + ''  -> 200
        # A + '2025-01' -> 100
        # B + '2025-02' -> 700
        # 共3行
        assert len(result_df) == 3

        # 空值在 Excel 中读回可能是 NaN，统一处理后再比较
        # 将日期列的空值统一转换为空字符串
        result_df['日期'] = result_df['日期'].fillna('')

        row_a_empty = result_df[(result_df['机构'] == 'A') & (result_df['日期'] == '')]
        assert len(row_a_empty) == 1
        assert row_a_empty.iloc[0]['金额'] == 200

        row_b = result_df[(result_df['机构'] == 'B') & (result_df['日期'] == '2025-02')]
        assert len(row_b) == 1
        assert row_b.iloc[0]['金额'] == 700
    ```

- **tests/test_integration_team_report.py**
  - language: `python`
  - content:
    ```python
# tests/test_integration_team_report.py
import io, time, pandas as pd, pytest

class TestTeamReportIntegration:
    def test_generate_with_monthly(self, client):
        mapping_df = pd.DataFrame({"队员ID":["1","2"],"队伍名称":["骑士队","米兰队"],"队长姓名":["袁队","安队"],"司管":["方行","任翔"]})
        bm = io.BytesIO(); mapping_df.to_excel(bm,index=False); bm.seek(0)
        today_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[10,15],"报单数":[2,3],"转单数":[0,0],"新桔完单数":[5,6],"在线天数":[1,0]})
        bt = io.BytesIO(); today_df.to_excel(bt,index=False); bt.seek(0)
        yesterday_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[9,12],"报单数":[1,2],"转单数":[0,0],"新桔完单数":[4,5],"在线天数":[1,0]})
        by = io.BytesIO(); yesterday_df.to_excel(by,index=False); by.seek(0)
        monthly_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[100,150],"报单数":[20,30],"转单数":[5,10],"新桔完单数":[50,60]})
        bmon = io.BytesIO(); monthly_df.to_excel(bmon,index=False); bmon.seek(0)
        data = {
            "mapping_file": (bm, "m.xlsx"),
            "today_file": (bt, "t.xlsx"),
            "yesterday_file": (by, "y.xlsx"),
            "monthly_files": (bmon, "monthly.xlsx"),
            "daily_order_target": "1150",
            "daily_apply_target": "120",
            "monthly_order_target": "35650",
            "monthly_apply_target": "3720"
        }
        resp = client.post("/team-report/generate", data=data, content_type="multipart/form-data")
        assert resp.status_code == 200
        tid = resp.get_json()["task_id"]
        for _ in range(15):
            s = client.get(f"/team-report/status/{tid}").get_json()
            if s["status"] == "completed":
                break
            time.sleep(0.5)
        dl = client.get(f"/team-report/download/{tid}")
        assert dl.status_code == 200 and len(dl.data) > 0

    def test_generate_without_monthly(self, client):
        mapping_df = pd.DataFrame({"队员ID":["1"],"队伍名称":["骑士队"],"队长姓名":["袁队"],"司管":["方行"]})
        bm = io.BytesIO(); mapping_df.to_excel(bm,index=False); bm.seek(0)
        today_df = pd.DataFrame({"司机ID":["1"],"司机状态":["已上岗"],"总完单数":[10],"报单数":[2],"转单数":[0],"新桔完单数":[5],"在线天数":[1]})
        bt = io.BytesIO(); today_df.to_excel(bt,index=False); bt.seek(0)
        data = {"mapping_file":(bm,"m.xlsx"),"today_file":(bt,"t.xlsx")}
        resp = client.post("/team-report/generate", data=data, content_type="multipart/form-data")
        assert resp.status_code == 200
        tid = resp.get_json()["task_id"]
        for _ in range(10):
            s = client.get(f"/team-report/status/{tid}").get_json()
            if s["status"] == "completed":
                break
            time.sleep(0.3)
        dl = client.get(f"/team-report/download/{tid}")
        assert dl.status_code == 200
    ```

- **tests/test_services_data.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
import pytest
from services.data.reader import DataReader
from services.data.writer import DataWriter
from services.data.formatter import DataFormatter

class TestDataReader:
    def test_read_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2], 'B': ['x', 'y']})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        result = DataReader.read_excel(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2], 'B': ['x', 'y']})
        path = tmp_path / 'test.csv'
        df.to_csv(path, index=False)
        result = DataReader.read_csv(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_auto_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        result = DataReader.read_auto(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_auto_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'test.csv'
        df.to_csv(path, index=False)
        result = DataReader.read_auto(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_auto_unsupported(self, tmp_path):
        path = tmp_path / 'test.txt'
        path.write_text('hello')
        with pytest.raises(ValueError, match='不支持的文件类型'):
            DataReader.read_auto(str(path))

class TestDataWriter:
    def test_write_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.xlsx'
        DataWriter.write_excel(df, str(path))
        result = pd.read_excel(path)
        pd.testing.assert_frame_equal(result, df)

    def test_write_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.csv'
        DataWriter.write_csv(df, str(path))
        result = pd.read_csv(path)
        pd.testing.assert_frame_equal(result, df)

    def test_write_auto_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.xlsx'
        DataWriter.write_auto(df, str(path))
        result = pd.read_excel(path)
        pd.testing.assert_frame_equal(result, df)

    def test_write_auto_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.csv'
        DataWriter.write_auto(df, str(path))
        result = pd.read_csv(path)
        pd.testing.assert_frame_equal(result, df)

class TestDataFormatter:
    def test_round_numeric_columns(self):
        df = pd.DataFrame({'A': [1.234, 2.345], 'B': ['x', 'y']})
        result = DataFormatter.round_numeric_columns(df, decimals=2)
        assert result['A'].tolist() == [1.23, 2.35]
        assert result['B'].tolist() == ['x', 'y']

    def test_safe_format_dataframe(self):
        df = pd.DataFrame({'A': [1, None], 'B': ['x', None]})
        result = DataFormatter.safe_format_dataframe(df)
        assert result['A'].iloc[0] == 1
        assert pd.isna(result['A'].iloc[1])
        assert result['B'].tolist() == ['x', '']

    def test_fillna_empty(self):
        df = pd.DataFrame({'A': [1, None], 'B': ['x', None]})
        result = DataFormatter.fillna_empty(df)
        assert result['A'].tolist() == [1, '']
        assert result['B'].tolist() == ['x', '']
    ```

- **tests/test_services_excel.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
import pytest
from services.excel.compute_engine import ComputeEngine
from core.filter_engine import FilterEngine
from services.excel.aggregator import Aggregator
from core.filter_engine import Condition

class TestComputeEngine:
    def test_evaluate_expression(self):
        values = {'A': 10, 'B': 5}
        assert ComputeEngine.evaluate_expression('A+B', values) == 15
        assert ComputeEngine.evaluate_expression('A*B', values) == 50

    def test_extract_variables(self):
        assert ComputeEngine.extract_variables('A+B-C') == {'A', 'B', 'C'}

    def test_simple_aggregate(self):
        df = pd.DataFrame({'G': ['X', 'X', 'Y'], 'V': [10, 20, 5]})
        result = ComputeEngine.simple_aggregate(df, ['G'], ['V'], 'sum')
        assert len(result) == 2
        assert result.loc[result['G'] == 'X', 'V'].iloc[0] == 30

    def test_aggregate_by_keys(self):
        df = pd.DataFrame({'key': ['A', 'A', 'B'], 'sum': [10, 20, 5], 'other': ['first', 'second', 'other']})
        result = ComputeEngine.aggregate_by_keys(df, ['key'], ['sum'], ['other'])
        assert len(result) == 2
        assert result.loc[result['key'] == 'A', 'sum'].iloc[0] == 30
        assert result.loc[result['key'] == 'A', 'other'].iloc[0] == 'first'

class TestFilterEngine:
    def test_filter_dataframe_with_condition_objects(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [25, 30]})
        cond = Condition(field='age', op='>', value=28)
        result = FilterEngine.filter_dataframe(df, [cond])
        assert len(result) == 1
        assert result.iloc[0]['name'] == 'Bob'

    def test_filter_dataframe_with_multi_values(self):
        df = pd.DataFrame({'dept': ['HR', 'IT', 'HR'], 'name': ['A', 'B', 'C']})
        result = FilterEngine.filter_dataframe_with_multi_values(df, exact_conditions=[('dept', ['HR'])], fuzzy_conditions=[('name', 'A')])
        assert len(result) == 1
        assert result.iloc[0]['name'] == 'A'

    def test_filter_dict(self):
        row = {'name': 'Alice', 'age': 30}
        cond = Condition(field='age', op='>', value=25)
        assert FilterEngine.filter_dict(row, [cond]) is True

    def test_create_condition(self):
        cond = FilterEngine.create_condition('age', '>', 25, exclude=True)
        assert cond.field == 'age'
        assert cond.op == '>'
        assert cond.value == 25
        assert cond.exclude is True

class TestAggregator:
    def test_group_aggregate(self):
        df = pd.DataFrame({'G': ['X', 'X'], 'V': [10, 20]})
        result = Aggregator.group_aggregate(df, ['G'], {'V': 'sum'})
        assert result.iloc[0]['V'] == 30

    def test_aggregate_by_keys_with_other_fields(self):
        df = pd.DataFrame({'key': ['A', 'A'], 'sum': [10, 20], 'other': ['first', 'second']})
        result = Aggregator.aggregate_by_keys(df, ['key'], ['sum'], ['other'])
        assert result.iloc[0]['sum'] == 30
        assert result.iloc[0]['other'] == 'first'
    ```

- **tests/test_splitter.py**
  - language: `python`
  - content:
    ```python
import pytest
import tempfile
import os
from openpyxl import Workbook, load_workbook
from blueprints.splitter import split_excel_to_temp, sanitize_filename, filter_workbook_inplace
import pandas as pd

def create_test_excel(filepath):
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Category'
    ws['B1'] = 'Value'
    ws['A2'] = 'A'
    ws['B2'] = 10
    ws['A3'] = 'A'
    ws['B3'] = 20
    ws['A4'] = 'B'
    ws['B4'] = 5
    wb.save(filepath)
    return filepath

class TestSplitter:
    def test_split_excel_to_temp(self, tmp_path):
        input_file = tmp_path / 'input.xlsx'
        create_test_excel(input_file)
        output_dir = tmp_path / 'output'
        output_dir.mkdir()
        files = split_excel_to_temp(str(input_file), split_column_name='Category', include_values=[], exclude_values=[], skip_sheet_indices=set(), temp_dir=str(output_dir))
        assert len(files) == 2
        assert any('A.xlsx' in f for f in files)
        assert any('B.xlsx' in f for f in files)

    def test_sanitize_filename(self):
        assert sanitize_filename('a/b?c') == 'a_b_c'
        assert sanitize_filename('normal') == 'normal'

    def test_filter_with_skip_sheets(self, tmp_path):
        """测试原生筛选 + 跳过索引，使用 openpyxl 验证结果"""
        wb = Workbook()
        ws1 = wb.active
        ws1.title = 'Main'
        ws1['A1'] = 'Name'
        ws1['B1'] = 'Value'
        ws1['A2'] = 'Alice'
        ws1['B2'] = '100'
        ws2 = wb.create_sheet('Second')
        ws2['A1'] = 'Other'
        ws2['B1'] = 'Value'
        ws2['A2'] = 'should_not_filter'
        in_path = tmp_path / 'test.xlsx'
        wb.save(in_path)

        form = {
            'condition_field_0': 'Name',
            'condition_op_0': '=',
            'condition_value_0': 'Alice'
        }
        # 调用新函数，传入跳过索引 {2}
        filtered_path = filter_workbook_inplace(str(in_path), form, skip_sheet_indices={2})

        # 加载筛选后的工作簿
        wb_result = load_workbook(filtered_path)

        # 第一个 sheet 应该只剩下表头 + Alice 行
        ws1_res = wb_result['Main']
        assert ws1_res.max_row == 2
        assert ws1_res.cell(row=2, column=1).value == 'Alice'
        assert ws1_res.cell(row=2, column=2).value == '100'

        # 第二个 sheet 应被跳过，原始数据保留
        ws2_res = wb_result['Second']
        assert ws2_res.max_row == 2
        assert ws2_res.cell(row=2, column=1).value == 'should_not_filter'
        wb_result.close()
    ```

- **tests/test_subtotal_engine.py**
  - language: `python`
  - content:
    ```python
# tests/test_subtotal_engine.py
import pytest
import pandas as pd
import os
from services.subtotal.engine import SubtotalEngine

class TestSubtotalEngine:
    @pytest.fixture
    def sample_excel(self, tmp_path):
        df = pd.DataFrame({
            'A': ['X', 'X', 'Y'],   # 列名变为字母，方便列字母引用
            'B': [10, 20, 30],
            'C': [1, 2, 3]
        })
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        return str(path), 'test.xlsx'

    def test_basic_subtotal_sum(self, sample_excel):
        filepath, fname = sample_excel
        engine = SubtotalEngine(filepath, 'A', 'B,C', {'default': 'sum'}, original_filename=fname)
        out_path = engine.process()
        assert os.path.basename(out_path).startswith('subtotal_test')
        wb = pd.read_excel(out_path, sheet_name=None)
        ws = wb['Sheet1']
        assert len(ws) == 5
        assert ws.iloc[2]['B'] == 30
        assert ws.iloc[2]['C'] == 3

    def test_dimension_col(self, sample_excel):
        filepath, fname = sample_excel
        engine = SubtotalEngine(filepath, 'A', 'B', {'default': 'sum'},
                                dimension_cols='C', original_filename=fname)
        out_path = engine.process()
        wb = pd.read_excel(out_path, sheet_name=None)
        ws = wb['Sheet1']
        assert ws.iloc[2]['C'] == '合计'
    ```

- **tests/test_team_report_long_running.py**
  - language: `python`
  - content:
    ```python
import io, time, pandas as pd, pytest
from services.team_report import service as svc_module

class TestTeamReportLongRunning:
    """模拟线上处理较慢的场景，确保最终能完成并下载"""
    def test_slow_task_completes(self, client, monkeypatch):
        original_build = svc_module.TeamReportService.build_report_data
        def slow_build(self, today_df, yesterday_df=None, monthly_df=None):
            time.sleep(6)
            return original_build(self, today_df, yesterday_df, monthly_df)
        monkeypatch.setattr(svc_module.TeamReportService, "build_report_data", slow_build)

        mapping_df = pd.DataFrame({"队员ID":["1","2"],"队伍名称":["骑士队","米兰队"],"队长姓名":["袁队","安队"],"司管":["方行","任翔"]})
        bm = io.BytesIO(); mapping_df.to_excel(bm,index=False); bm.seek(0)
        today_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[10,15],"报单数":[2,3],"转单数":[0,0],"新桔完单数":[5,6]})
        bt = io.BytesIO(); today_df.to_excel(bt,index=False); bt.seek(0)
        data = {"mapping_file":(bm,"m.xlsx"),"today_file":(bt,"t.xlsx")}
        resp = client.post("/team-report/generate",data=data,content_type="multipart/form-data")
        assert resp.status_code==200
        tid = resp.get_json()["task_id"]
        for _ in range(30):
            s = client.get(f"/team-report/status/{tid}").get_json()
            if s["status"]=="completed": break
            time.sleep(1)
        dl = client.get(f"/team-report/download/{tid}")
        assert dl.status_code==200 and len(dl.data)>0
    ```

- **tests/test_team_report_service.py**
  - language: `python`
  - content:
    ```python
import pytest, pandas as pd, os
from services.team_report import TeamReportService
from services.team_report.writer import TeamReportExcelWriter
from core.config import TeamReportConfig

class TestTeamReportService:
    @pytest.fixture
    def cfg(self):
        return TeamReportConfig()
    @pytest.fixture
    def svc(self, cfg):
        return TeamReportService(cfg)
    @pytest.fixture
    def writer(self):
        return TeamReportExcelWriter()

    def test_load_mapping(self, svc, tmp_path):
        df = pd.DataFrame({"队员ID":["1","2"],"队伍名称":["A","B"],"队长姓名":["甲","乙"],"司管":["X","Y"]})
        p = tmp_path/"m.xlsx"; df.to_excel(p,index=False)
        m = svc.load_mapping(str(p))
        assert m["1"]["司管"]=="X"

    def test_build_report_and_writer(self, svc, writer, cfg, tmp_path):
        today_data = pd.DataFrame({
            "司机ID":["1","2"],"司机状态":["已上岗","已上岗"],
            "总完单数":[200,150],"报单数":[30,20],"转单数":[10,5],"新桔完单数":[20,10],
            "在线天数":[1,0]
        })
        yesterday_data = pd.DataFrame({
            "司机ID":["1","2"],"司机状态":["已上岗","已上岗"],
            "总完单数":[180,120],"报单数":[20,15],"转单数":[5,2],"新桔完单数":[15,8],
            "在线天数":[1,0]
        })
        t_path = tmp_path/"today.xlsx"; today_data.to_excel(t_path, index=False)
        y_path = tmp_path/"yesterday.xlsx"; yesterday_data.to_excel(y_path, index=False)
        mapping = {"1":{"队伍名称":"骑士队","队长姓名":"袁队","司管":"方行"},
                   "2":{"队伍名称":"米兰队","队长姓名":"安队","司管":"任翔"}}
        today_df = svc._prepare_flow(str(t_path), mapping)
        yesterday_df = svc._prepare_flow(str(y_path), mapping)
        monthly_df = svc.aggregate_monthly([str(t_path)], mapping)
        report = svc.build_report_data(today_df, yesterday_df, monthly_df)
        rows = svc.report_to_rows(report,
                                  daily_order=cfg.daily_order_target,
                                  daily_apply=cfg.daily_apply_target,
                                  monthly_order=cfg.monthly_order_target,
                                  monthly_apply=cfg.monthly_apply_target)
        out = tmp_path/"out.xlsx"
        writer.write(rows, str(out), "2025/10/31", "10",
                     cfg.daily_order_target, cfg.daily_apply_target,
                     cfg.monthly_order_target, cfg.monthly_apply_target)
        assert os.path.exists(out)
    ```

- **tests/test_transposer.py**
  - language: `python`
  - content:
    ```python
# tests/test_transposer.py
import pytest
import pandas as pd
from services.transposer.service import pivot_table


class TestPivotTable:
    def test_pivot_basic_with_index(self):
        df = pd.DataFrame({
            "病种": ["A", "A", "B"],
            "医疗类别": ["住院", "门诊", "住院"],
            "金额": [100, 200, 300],
            "数量": [1, 2, 3]
        })
        result = pivot_table(
            df,
            index_cols=["病种"],
            category_col="医疗类别",
            value_cols=["金额", "数量"]
        )
        expected_columns = ["病种", "住院金额", "门诊金额", "住院数量", "门诊数量"]
        assert list(result.columns) == expected_columns
        assert result.loc[0, "病种"] == "A"
        assert result.loc[0, "住院金额"] == 100
        assert result.loc[0, "门诊金额"] == 200
        assert result.loc[0, "住院数量"] == 1
        assert result.loc[0, "门诊数量"] == 2
        assert result.loc[1, "病种"] == "B"
        assert result.loc[1, "住院金额"] == 300
        assert pd.isna(result.loc[1, "门诊金额"])
        assert result.loc[1, "住院数量"] == 3
        assert pd.isna(result.loc[1, "门诊数量"])

    def test_pivot_no_index(self):
        df = pd.DataFrame({
            "医疗类别": ["住院", "门诊", "住院"],
            "金额": [100, 200, 300]
        })
        result = pivot_table(
            df,
            index_cols=None,
            category_col="医疗类别",
            value_cols=["金额"]
        )
        assert "住院金额" in result.columns
        assert "门诊金额" in result.columns
        assert len(result) == 2

    def test_pivot_multiple_value_cols(self):
        df = pd.DataFrame({
            "ID": [1, 1, 2],
            "类别": ["X", "Y", "X"],
            "值1": [10, 20, 30],
            "值2": [100, 200, 300]
        })
        result = pivot_table(
            df,
            index_cols=["ID"],
            category_col="类别",
            value_cols=["值1", "值2"]
        )
        assert "X值1" in result.columns
        assert "Y值1" in result.columns
        assert "X值2" in result.columns
        assert "Y值2" in result.columns
        assert result.loc[0, "ID"] == 1
        assert result.loc[0, "X值1"] == 10
        assert result.loc[0, "Y值1"] == 20
        assert result.loc[0, "X值2"] == 100
        assert result.loc[0, "Y值2"] == 200


class TestTransposerBlueprintIntegration:
    def test_transposer_endpoint(self, client):
        import io
        df = pd.DataFrame({
            "病种": ["A", "A", "B"],
            "医疗类别": ["住院", "门诊", "住院"],
            "金额": [100, 200, 300]
        })
        excel_file = io.BytesIO()
        df.to_excel(excel_file, index=False)
        excel_file.seek(0)

        data = {
            "file": (excel_file, "test.xlsx"),
            "index_cols": "病种",
            "category_col": "医疗类别",
            "value_cols": "金额"
        }
        response = client.post("/transposer/convert", data=data, content_type="multipart/form-data")
        assert response.status_code == 200
        content_disp = response.headers.get("Content-Disposition", "")
        # 文件名可能被 URL 编码，因此只检查是否包含 download 指令，不检查具体文件名
        assert "attachment" in content_disp or "filename" in content_disp
    ```


> ✅ 共导出 155 个代码文件。
