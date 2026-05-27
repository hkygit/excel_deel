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
import tempfile
import pandas as pd
from flask import Blueprint, render_template, request, send_file, current_app, after_this_request
from openpyxl.utils.exceptions import InvalidFileException

from services import FileService, ComputeEngine
from core.decorators import handle_exceptions
from core.response import error_response
from core.file_naming import generate_output_filename
from core.header_detection import detect_header_start
from core.exceptions import ValidationError
from services.excel.calculator_processor import ColumnCalculatorProcessor, AggregateBackfillProcessor

logger = logging.getLogger(__name__)
calc_bp = Blueprint('calc', __name__, url_prefix='/calc', template_folder='../templates/calculator')

def convert_xls_to_xlsx(filepath):
    """将 .xls 文件转换为 .xlsx，返回新文件路径（临时文件）"""
    # 读取原始 .xls 文件
    try:
        df = pd.read_excel(filepath, sheet_name=None, dtype=str)
        # 创建临时 .xlsx 文件
        fd, new_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            for sheet_name, data in df.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info(f"已将 .xls 文件转换为 .xlsx: {new_path}")
        return new_path
    except Exception as e:
        raise ValidationError(f"无法转换 .xls 文件: {str(e)}")

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
    converted = False
    # 如果是 .xls 文件，自动转换为 .xlsx
    if ext == '.xls':
        try:
            new_path = convert_xls_to_xlsx(filepath)
            # 删除原始 .xls 文件，使用新文件
            file_service.cleanup([filepath])
            filepath = new_path
            converted = True
        except Exception as e:
            file_service.cleanup([filepath])
            raise ValidationError(f"旧版 Excel 文件转换失败: {str(e)}")

    calc_type = request.form.get('calc_type')
    if calc_type not in ('column', 'aggregate'):
        file_service.cleanup([filepath])
        return error_response('无效的操作类型', 400)

    # 表头行号
    header_row_str = request.form.get('header_row', '').strip()
    if header_row_str:
        try:
            header_row = int(header_row_str)
            if header_row < 1:
                raise ValueError
        except ValueError:
            return error_response('表头行号必须为正整数', 400)
    else:
        try:
            data_start = detect_header_start(filepath)
            header_row = data_start - 1
            if header_row < 1:
                header_row = 1
        except Exception:
            header_row = 1

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
        raise ValidationError(f"文件格式错误，无法打开: {str(e)}。请确保文件是有效的 Excel 文件（.xlsx/.xls）且未被损坏。")
    except Exception as e:
        file_service.cleanup([filepath])
        raise e

    @after_this_request
    def cleanup_files(response):
        try:
            if filepath and os.path.exists(filepath):
                os.remove(filepath)
            if result_path and os.path.exists(result_path):
                os.remove(result_path)
        except Exception as e:
            logger.warning(f"清理临时文件失败: {e}")
        return response

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
    ```

- **core/decorators.py**
  - language: `python`
  - content:
    ```python
# core/decorators.py
import logging
from functools import wraps
from flask import jsonify, request
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
                # 预期内的参数/文件格式错误，返回具体消息
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
    """从表达式中提取大写列字母"""
    return set(re.findall(r'[A-Z]+', expr.upper()))

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
        ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
        if ext not in allowed_extensions:
            raise ValueError(f"不支持的文件类型: {ext}")
        filename = secure_filename(file.filename)
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
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 列计算与聚合回填处理器
支持 .xls 文件自动转换为 .xlsx
"""

import os
import gc
import logging
import tempfile
import openpyxl
import pandas as pd
import re
from core.excel_utils import ExcelFile, copy_merge_structure
from core.data_access_layer import get_data_access_layer
from core.csv_utils import build_column_mapping
from services.excel.compute_engine import ComputeEngine

logger = logging.getLogger(__name__)

def convert_xls_to_xlsx(filepath):
    """将 .xls 文件转换为临时 .xlsx 文件"""
    if not filepath.lower().endswith('.xls') or filepath.lower().endswith('.xlsx'):
        return filepath
    logger.info(f"正在将 .xls 文件转换为 .xlsx: {filepath}")
    try:
        # 使用 pandas 读取 .xls 并写入临时 .xlsx
        df = pd.read_excel(filepath, sheet_name=None, dtype=str)
        fd, temp_xlsx = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
            for sheet_name, sheet_df in df.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info(f"转换成功: {temp_xlsx}")
        return temp_xlsx
    except Exception as e:
        logger.error(f"转换 .xls 文件失败: {e}")
        raise ValueError(f"无法转换 .xls 文件，请确保文件未损坏且可读: {e}")

class ColumnCalculatorProcessor:
    """列计算处理器（自动过滤结果列，支持 .xls）"""
    def __init__(self, filepath, header_row, expression, result_col_input, merge_by_cols_str, compute_engine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.original_expression = expression
        self.result_col_input = result_col_input
        self.merge_by_cols_str = merge_by_cols_str
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_file = None

    def _remove_result_column_from_expr(self, expression, result_col_letter):
        """彻底从表达式中移除结果列字母"""
        if not result_col_letter or result_col_letter not in expression:
            return expression
        pattern = r'(?:[+\-*/]?)' + re.escape(result_col_letter) + r'(?:[+\-*/]?)'
        new_expr = re.sub(pattern, '', expression)
        new_expr = re.sub(r'[+\-*/]{2,}', '', new_expr)
        new_expr = re.sub(r'^[+\-*/]', '', new_expr)
        new_expr = re.sub(r'[+\-*/]$', '', new_expr)
        if new_expr != expression:
            logger.warning(f"表达式中包含了结果列 {result_col_letter}，已自动移除: {expression} -> {new_expr}")
        return new_expr

    def process(self):
        # 0. 如果是 .xls 文件，先转换为临时 .xlsx
        ext = os.path.splitext(self.original_filepath)[1].lower()
        if ext == '.xls':
            self.temp_file = convert_xls_to_xlsx(self.original_filepath)
            filepath = self.temp_file
        else:
            filepath = self.original_filepath

        try:
            # 1. 读取 DataFrame 并获取列名
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = self.dal._auto_clean_dataframe(df)
            headers = list(df.columns)

            # 2. 构建列字母到列名的映射
            letter_to_name, name_to_letter = build_column_mapping(headers)
            logger.debug(f"列字母映射: {letter_to_name}")

            # 3. 获取结果列字母
            result_col_letter = None
            try:
                col_idx = openpyxl.utils.column_index_from_string(self.result_col_input)
                result_col_letter = openpyxl.utils.get_column_letter(col_idx)
            except:
                pass

            # 4. 自动移除结果列字母
            expression = self.original_expression
            if result_col_letter:
                expression = self._remove_result_column_from_expr(expression, result_col_letter)
                if not expression.strip():
                    raise ValueError(f"表达式在移除结果列 {result_col_letter} 后为空")

            # 5. 提取并映射变量
            variables = self.compute_engine.extract_variables(expression)
            logger.info(f"表达式中的列字母: {variables}")
            mapped_vars = {}
            missing = []
            for letter in variables:
                if letter in letter_to_name:
                    mapped_vars[letter] = letter_to_name[letter]
                else:
                    missing.append(letter)
            if missing:
                raise ValueError(f"表达式中的列字母 {missing} 无法映射到文件中的任何列（可用列字母: {list(letter_to_name.keys())}）")

            # 6. 计算
            def calc_row(row):
                context = {letter: row[col_name] for letter, col_name in mapped_vars.items()}
                return self.compute_engine.evaluate_expression(expression, context)
            df["Result"] = df.apply(calc_row, axis=1)
            logger.info("表达式计算完成")

            # 7. 打开原始 Excel 写入（注意：如果是转换后的临时文件，原始文件无法写入，需要将结果写回原始文件？）
            # 策略：如果原文件是 .xls，则最终输出新的 .xlsx 文件，保留在原文件同目录
            if ext == '.xls':
                # 原文件是 .xls，不能直接用 openpyxl 修改，我们直接基于转换后的临时文件写入结果，保存到新文件
                output_base = os.path.splitext(os.path.basename(self.original_filepath))[0]
                output_dir = os.path.dirname(self.original_filepath)
                output_filename = f"{output_base}_处理版.xlsx"
                result_path = os.path.join(output_dir, output_filename)
                # 将计算结果写入临时文件并另存
                with ExcelFile(filepath) as ef:
                    ws = ef.active_sheet
                    # 处理结果列...
                    # 这里需要与下面逻辑类似，但为了避免重复，复用同一份代码块较复杂。
                    # 鉴于时间，我们简单说明：可以调用一个通用写入函数。
                    # 为了快速修复，我们假设原文件不是 .xls，但实际需要处理。这里提供完善方案。
                    # 由于代码量较大，我将在下面统一写入代码块中处理两种情况。
                pass

            # 统一写入逻辑（适配原始文件和临时文件）
            with ExcelFile(filepath) as ef:
                ws = ef.active_sheet
                # 处理结果列
                try:
                    result_col_idx = openpyxl.utils.column_index_from_string(self.result_col_input)
                    is_new_col = False
                    result_col_name = self.result_col_input
                except:
                    is_new_col = True
                    result_col_name = self.result_col_input.strip()
                    max_col = ws.max_column
                    result_col_idx = max_col + 1
                    ws.insert_cols(result_col_idx)
                    self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

                if not is_new_col:
                    existing = self._safe_get_cell_value(ws, self.header_row, result_col_idx)
                    if existing is None or str(existing).strip() == "":
                        self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

                start_row = self.header_row + 1
                end_row = ws.max_row

                # 清除结果列合并
                to_remove = []
                for merged in ws.merged_cells.ranges:
                    if merged.min_col <= result_col_idx <= merged.max_col:
                        to_remove.append(merged)
                for merged in to_remove:
                    ws.unmerge_cells(str(merged))

                # 写入计算结果
                for i, val in enumerate(df["Result"]):
                    ws.cell(start_row + i, result_col_idx).value = val

                # 处理合并依据列
                if self.merge_by_cols_str:
                    merge_keys = [openpyxl.utils.column_index_from_string(c.strip()) for c in self.merge_by_cols_str.split(",") if c.strip()]
                    if merge_keys:
                        ref_col = merge_keys[0]
                        has_merge = any(m.min_col <= ref_col <= m.max_col for m in ws.merged_cells.ranges)
                        if has_merge:
                            try:
                                copy_merge_structure(ws, ref_col, ws, result_col_idx, start_row, end_row)
                                logger.info(f"已复制合并结构 from {self.merge_by_cols_str} to {self.result_col_input}")
                            except Exception as e:
                                logger.warning(f"复制合并结构失败: {e}")
                        else:
                            logger.info(f"源列 {self.merge_by_cols_str} 无合并单元格，跳过")

                # 确定输出路径
                if ext == '.xls':
                    # 对于 .xls 文件，输出到原目录下，文件名加“_处理版.xlsx”
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")
                else:
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")

                ef.save(result_path)
                ef._workbook.close()

            gc.collect()
            return result_path

        finally:
            # 清理临时文件（如果创建了）
            if self.temp_file and os.path.exists(self.temp_file):
                try:
                    os.unlink(self.temp_file)
                except:
                    pass

    @staticmethod
    def _safe_get_cell_value(ws, row, col):
        cell = ws.cell(row, col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                    return ws.cell(merged_range.min_row, merged_range.min_col).value
        return cell.value

    @staticmethod
    def _safe_write_cell(ws, row, col, value):
        cell = ws.cell(row, col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            to_unmerge = []
            for merged_range in ws.merged_cells.ranges:
                if (merged_range.min_row <= row <= merged_range.max_row and
                    merged_range.min_col <= col <= merged_range.max_col):
                    to_unmerge.append(str(merged_range))
            for rng in to_unmerge:
                ws.unmerge_cells(rng)
        ws.cell(row, col).value = value


class AggregateBackfillProcessor:
    """聚合回填处理器（同样支持 .xls 转换）"""
    def __init__(self, filepath, header_row, group_cols_raw, agg_cols_raw, agg_func, result_col_input, compute_engine):
        self.original_filepath = filepath
        self.header_row = header_row
        self.group_cols_raw = group_cols_raw
        self.agg_cols_raw = agg_cols_raw
        self.agg_func = agg_func
        self.result_col_input = result_col_input
        self.compute_engine = compute_engine
        self.dal = get_data_access_layer()
        self.temp_file = None

    def process(self):
        ext = os.path.splitext(self.original_filepath)[1].lower()
        if ext == '.xls':
            self.temp_file = convert_xls_to_xlsx(self.original_filepath)
            filepath = self.temp_file
        else:
            filepath = self.original_filepath

        try:
            # 读取表头获取列名
            df_header = pd.read_excel(filepath, header=self.header_row - 1, nrows=0)
            headers = list(df_header.columns)
            letter_to_name, name_to_letter = build_column_mapping(headers)

            def resolve_col(col):
                col = col.strip()
                if col in headers:
                    return col
                upper = col.upper()
                if upper in letter_to_name:
                    return letter_to_name[upper]
                raise ValueError(f"列 '{col}' 无法映射到实际列名")

            group_cols = [resolve_col(c) for c in self.group_cols_raw]
            agg_cols = [resolve_col(c) for c in self.agg_cols_raw]

            # 读取数据并清洗
            df = pd.read_excel(filepath, header=self.header_row - 1, dtype=str)
            df = self.dal._auto_clean_dataframe(df)

            missing = [c for c in group_cols + agg_cols if c not in df.columns]
            if missing:
                raise ValueError(f"列不存在: {missing}")

            result_df = self.compute_engine.simple_aggregate(df, group_cols, agg_cols, self.agg_func)

            # 回填到原始 Excel
            with ExcelFile(filepath) as ef:
                ws = ef.active_sheet
                # 建立列名到列索引的映射
                name_to_col_idx = {}
                for idx, cell in enumerate(ws[self.header_row], start=1):
                    if cell.value is not None:
                        name_to_col_idx[str(cell.value).strip()] = idx

                def get_col_index(col):
                    col = col.strip()
                    if col in name_to_col_idx:
                        return name_to_col_idx[col]
                    try:
                        return openpyxl.utils.column_index_from_string(col)
                    except ValueError:
                        raise ValueError(f"列 '{col}' 在原始表中不存在")

                group_col_indices = [get_col_index(c) for c in group_cols]

                # 处理结果列
                try:
                    result_col_idx = openpyxl.utils.column_index_from_string(self.result_col_input)
                    is_new_col = False
                    result_col_name = self.result_col_input
                except:
                    is_new_col = True
                    result_col_name = self.result_col_input.strip()
                    result_col_idx = ws.max_column + 1
                    ws.insert_cols(result_col_idx)
                    self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

                if not is_new_col:
                    existing = self._safe_get_cell_value(ws, self.header_row, result_col_idx)
                    if existing is None or str(existing).strip() == "":
                        self._safe_write_cell(ws, self.header_row, result_col_idx, result_col_name)

                # 构建聚合结果映射
                result_map = {}
                for _, row in result_df.iterrows():
                    key = tuple(str(row[c]) for c in group_cols)
                    result_map[key] = row[agg_cols[0]]

                start_row = self.header_row + 1
                end_row = ws.max_row
                for row in range(start_row, end_row + 1):
                    key = tuple(str(ws.cell(row, idx).value) for idx in group_col_indices)
                    ws.cell(row, result_col_idx).value = result_map.get(key, "")

                if ext == '.xls':
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")
                else:
                    base_name = os.path.splitext(os.path.basename(self.original_filepath))[0]
                    result_path = os.path.join(os.path.dirname(self.original_filepath), f"{base_name}_处理版.xlsx")

                ef.save(result_path)
                ef._workbook.close()

            gc.collect()
            return result_path
        finally:
            if self.temp_file and os.path.exists(self.temp_file):
                try:
                    os.unlink(self.temp_file)
                except:
                    pass

    @staticmethod
    def _safe_get_cell_value(ws, row, col):
        cell = ws.cell(row, col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                    return ws.cell(merged_range.min_row, merged_range.min_col).value
        return cell.value

    @staticmethod
    def _safe_write_cell(ws, row, col, value):
        cell = ws.cell(row, col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            to_unmerge = []
            for merged_range in ws.merged_cells.ranges:
                if (merged_range.min_row <= row <= merged_range.max_row and
                    merged_range.min_col <= col <= merged_range.max_col):
                    to_unmerge.append(str(merged_range))
            for rng in to_unmerge:
                ws.unmerge_cells(rng)
        ws.cell(row, col).value = value
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
            'expression': '',
            'result_col_column': 'D'
        }
        response = client.post('/calc/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '表达式' in response.get_json()['error']
    ```


> ✅ 共导出 25 个代码文件。
