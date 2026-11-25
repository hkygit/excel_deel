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
"""依赖注入容器。

设计要点：
- config = Singleton(lambda: settings) 兼容旧调用点：
    container.config.provided.checker()
- 规范化子配置 provider：checker_config / team_report_config /
  field_parser_config / splitter_config / convert_config
- TaskStorage 的 status_dir 独立于 output_dir（output_folder/task_status）
- Container(config=...) 支持在实例化时覆盖配置，便于测试
"""
import os

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
from services.excel.comparator import ExcelComparator
from services.excel.filter import ExcelFilter
from services.excel.aggregator import Aggregator
from services.field_parser import FieldParserService
from services.checker import CheckerService
from services.convert.service import ConvertService
from services.team_report import TeamReportService
from services.team_report.writer import TeamReportExcelWriter


def _task_status_dir(s) -> str:
    """任务状态目录：与 output_folder 分离，避免与结果文件混放。"""
    return os.path.join(s.output_folder, "task_status")


class Container(containers.DeclarativeContainer):
    # ==================== 配置层 ====================
    # 保留 config = Singleton(lambda: settings) 以兼容旧调用点：
    #   container.config.provided.checker()
    # 新代码请优先使用下方规范化的子配置 provider。
    config = providers.Singleton(lambda: settings)

    # ---- 规范化子配置 provider ----
    splitter_config = providers.Callable(lambda s: s.splitter, config)
    checker_config = providers.Callable(lambda s: s.checker, config)
    field_parser_config = providers.Callable(lambda s: s.field_parser, config)
    team_report_config = providers.Callable(lambda s: s.team_report, config)
    convert_config = providers.Callable(lambda s: s.convert, config)

    # ==================== 任务存储 ====================
    task_storage = providers.Singleton(
        TaskStorage,
        status_dir=providers.Callable(_task_status_dir, config),
        output_dir=providers.Callable(lambda s: s.output_folder, config),
    )
    task_manager = providers.Singleton(
        TaskManager,
        storage=task_storage,
        max_concurrent=3,
    )

    # ==================== 数据层 ====================
    data_reader = providers.Singleton(DataReader)
    data_writer = providers.Singleton(DataWriter)
    data_formatter = providers.Singleton(DataFormatter)
    file_service = providers.Singleton(FileService)

    # ==================== Excel 服务 ====================
    compute_engine = providers.Singleton(ComputeEngine)
    filter_engine = providers.Singleton(FilterEngine)
    excel_comparator = providers.Singleton(ExcelComparator)
    excel_filter = providers.Singleton(ExcelFilter)
    aggregator = providers.Singleton(Aggregator)

    # ==================== 业务服务 ====================
    team_report_service = providers.Factory(
        TeamReportService,
        config=team_report_config,
    )
    team_report_writer = providers.Singleton(TeamReportExcelWriter)

    field_parser = providers.Singleton(
        FieldParserService,
        similarity_threshold=providers.Callable(
            lambda c: c.similarity_threshold,
            field_parser_config,
        ),
    )

    checker_service = providers.Factory(
        CheckerService,
        config=checker_config,
    )
    convert_service = providers.Factory(ConvertService)
    ```

- **core/excel_helpers.py**
  - language: `python`
  - content:
    ```python
"""
Excel 文件辅助函数。

P0 变更：
- normalize_excel_file 对已经是有效 .xlsx 的文件直接返回原路径，
  避免无意义的"复制到临时文件再读回"。
- 新增 _validate_xlsx 做轻量校验（文件头 + ZIP 结构 + openpyxl 尺寸检查）。
"""
import os
import logging
import tempfile
import zipfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

logger = logging.getLogger(__name__)


# Excel 单表硬限制（参考 Office Open XML 规范）
MAX_EXCEL_ROWS = 1_048_576   # 2^20
MAX_EXCEL_COLS = 16_384      # 2^14


class ExcelNormalizationError(ValueError):
    """
    Excel 文件规范化过程中无法安全处理时抛出。

    当前触发场景：
    - 单表行数 / 列数超过设定的上限（不再静默截断）

    继承自 ValueError 以保持向后兼容：调用方若只 except ValueError 仍能捕获。
    """


def _check_size_limits(rows: int, cols: int, sheet_label: str,
                       max_rows: int, max_cols: int) -> None:
    """尺寸超限时抛 ExcelNormalizationError。"""
    if rows and rows > max_rows:
        raise ExcelNormalizationError(
            f"{sheet_label} 行数 {rows:,} 超过上限 {max_rows:,}；"
            f"为避免静默截断数据，处理已中止。"
            f"建议先拆分/抽样文件，或通过 max_rows 参数显式提高上限。"
        )
    if cols and cols > max_cols:
        raise ExcelNormalizationError(
            f"{sheet_label} 列数 {cols:,} 超过上限 {max_cols:,}；"
            f"为避免静默截断数据，处理已中止。"
            f"建议先拆分/抽样文件，或通过 max_cols 参数显式提高上限。"
        )


def _validate_xlsx(filepath: str, max_rows: int, max_cols: int) -> bool:
    """快速校验 .xlsx 是否有效且尺寸合规。

    返回:
        True  - 文件有效且尺寸合规（可安全地返回原路径）
        False - 文件无效（需走转换路径）
    抛出:
        ExcelNormalizationError - 尺寸超限（不进入转换路径）

    不创建任何临时文件，不修改原文件。
    """
    # 1) 文件头检查（ZIP magic）
    try:
        with open(filepath, 'rb') as f:
            if f.read(4) != b'PK\x03\x04':
                return False
    except OSError:
        return False

    # 2) ZIP 结构检查（比 openpyxl 更轻量，不解析样式 / 关系）
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            names = zf.namelist()
            # 有效 xlsx 至少包含 xl/ 目录下的内容
            if not any(n.startswith('xl/') for n in names):
                return False
    except (zipfile.BadZipFile, OSError):
        return False

    # 3) openpyxl 打开 + 尺寸检查（read_only 模式下 max_row / max_column
    #    从 XML 元数据读取，不会真的全量扫描单元格）
    try:
        wb = load_workbook(
            filename=filepath,
            data_only=True,
            read_only=True,
            keep_vba=False,
        )
    except Exception:
        return False

    try:
        for ws in wb.worksheets:
            rows = ws.max_row or 0
            cols = ws.max_column or 0
            if rows > max_rows or cols > max_cols:
                _check_size_limits(
                    rows, cols,
                    sheet_label=f"工作表 '{ws.title}'",
                    max_rows=max_rows, max_cols=max_cols,
                )
    finally:
        wb.close()

    return True


def normalize_excel_file(
    filepath: str,
    max_rows: int = MAX_EXCEL_ROWS,
    max_cols: int = MAX_EXCEL_COLS,
) -> str:
    """确保文件是可被下游正常读写的 .xlsx。

    P0 语义变化：
    - 对**有效的 .xlsx**：直接返回原路径（零复制），调用方需保证只读不改
    - 对 .xls / .xlsm / 损坏 .xlsx：走转换路径，返回新的临时文件
    - 尺寸超限：抛 ExcelNormalizationError（不静默截断）

    Args:
        filepath: 原始文件路径
        max_rows: 单表允许的最大行数（默认 Excel 硬限制）
        max_cols: 单表允许的最大列数（默认 Excel 硬限制）

    Returns:
        可安全读写的 .xlsx 路径（可能是原路径，也可能是新的临时文件）

    Raises:
        FileNotFoundError: 源文件不存在
        ExcelNormalizationError: 超过行/列上限
        ValueError: 所有读取方式都失败
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()

    # ===== P0：有效 .xlsx 直接返回原路径，零复制 =====
    if ext == '.xlsx':
        try:
            if _validate_xlsx(filepath, max_rows, max_cols):
                logger.info(
                    f"文件已是有效 .xlsx，直接使用原路径: {filepath}"
                )
                return filepath
        except ExcelNormalizationError:
            # 尺寸超限：立即向上抛，**不走转换路径**
            # 原因：pandas / openpyxl 的全量读更容易 OOM，及早拦截更安全
            raise
        except Exception as e:  # noqa: BLE001
            logger.warning(
                f".xlsx 有效性检测异常，降级为转换路径: {e}"
            )

    # ===== 以下为需要转换的场景（.xls / .xlsm / 损坏 .xlsx）=====
    temp_dir = tempfile.gettempdir()
    base_name = os.path.basename(filepath)
    output_name = (
        f"normalized_{os.path.splitext(base_name)[0]}_{os.urandom(4).hex()}.xlsx"
    )
    output_path = os.path.join(temp_dir, output_name)

    def _safe_unlink(path: str) -> None:
        if path and os.path.exists(path):
            try:
                os.unlink(path)
            except Exception as cleanup_err:  # noqa: BLE001
                logger.debug(f"清理临时文件失败: {path} - {cleanup_err}")

    try:
        # 方法1: 用 openpyxl 读取并另存
        try:
            logger.info(f"尝试用openpyxl读取文件: {filepath}")
            wb = load_workbook(
                filename=filepath,
                data_only=True,
                read_only=True,
                keep_vba=False,
            )

            # 尺寸检查（覆盖 .xls / .xlsm 等 _validate_xlsx 未校验的场景）
            for ws in wb.worksheets:
                rows = ws.max_row or 0
                cols = ws.max_column or 0
                if rows > max_rows or cols > max_cols:
                    wb.close()
                    _safe_unlink(output_path)
                    _check_size_limits(
                        rows, cols,
                        sheet_label=f"工作表 '{ws.title}'",
                        max_rows=max_rows, max_cols=max_cols,
                    )

            # 完整复制所有行列（read_only 模式下用 iter_rows 最快）
            from openpyxl import Workbook
            wb_new = Workbook()
            wb_new.remove(wb_new.active)

            for ws in wb.worksheets:
                ws_new = wb_new.create_sheet(title=ws.title)
                for row in ws.iter_rows(values_only=True):
                    ws_new.append(list(row))

            wb.close()
            wb_new.save(output_path)
            wb_new.close()

            logger.info(f"已用openpyxl转换文件: {output_path}")
            return output_path

        except ExcelNormalizationError:
            raise
        except Exception as e:  # noqa: BLE001
            logger.warning(f"openpyxl 转换失败: {e}，回退到 pandas 方式")

            # 方法2: pandas 兜底
            try:
                engine = None
                if ext in ['.xls', '.xlsm']:
                    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

                df = pd.read_excel(
                    filepath,
                    dtype=str,
                    na_filter=False,
                    engine=engine,
                    header=None,
                )

                n_rows = df.shape[0]
                n_cols = df.shape[1]
                if n_rows > max_rows or n_cols > max_cols:
                    _safe_unlink(output_path)
                    _check_size_limits(
                        n_rows, n_cols,
                        sheet_label=f"文件 '{os.path.basename(filepath)}'",
                        max_rows=max_rows, max_cols=max_cols,
                    )

                df = df.fillna('')
                df.to_excel(
                    output_path, index=False, header=False,
                    engine='openpyxl',
                )

                logger.info(f"已用pandas转换文件: {output_path}")
                return output_path

            except ExcelNormalizationError:
                raise
            except Exception as pd_e:  # noqa: BLE001
                logger.error(f"pandas 读取也失败: {pd_e}")

                # 方法3: 生成占位文件（保留原有兜底行为）
                try:
                    from openpyxl import Workbook
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Data"
                    ws['A1'] = "错误: 无法读取原始文件"
                    ws['A2'] = f"原始错误: {str(pd_e)[:100]}"
                    wb.save(output_path)
                    logger.warning(f"创建了空的修复文件: {output_path}")
                    return output_path
                except Exception as create_e:  # noqa: BLE001
                    logger.error(f"创建修复文件也失败: {create_e}")
                    raise ValueError(
                        f"无法读取或转换 Excel 文件。请确保文件是有效的 Excel 文件。"
                        f"原始错误: {pd_e}"
                    )

    except Exception as final_e:
        logger.error(f"所有转换方法都失败: {final_e}")
        _safe_unlink(output_path)
        raise


def safe_cleanup_dataframe(df: pd.DataFrame,
                           auto_convert_numeric: bool = True) -> pd.DataFrame:
    """安全清理 DataFrame。"""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace('', pd.NA)
            df[col] = df[col].replace('nan', pd.NA)
            df[col] = df[col].replace('None', pd.NA)
            df[col] = df[col].replace('null', pd.NA)

        if auto_convert_numeric and df[col].dtype == 'object':
            try:
                numeric_series = pd.to_numeric(df[col], errors='coerce')
                if numeric_series.notna().any():
                    df[col] = numeric_series
            except Exception:
                pass

    return df


def build_column_name_mapping(ws, header_row: int):
    """构建列名到列索引的映射。"""
    mapping = {}
    max_col = min(ws.max_column, 100)

    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        if cell and cell.value:
            col_name = str(cell.value).strip()
            if col_name:
                mapping[col_name] = col_idx

    return mapping


def safe_get_cell_value(ws, row: int, col: int):
    """安全获取单元格值。"""
    try:
        if row > ws.max_row or col > ws.max_column:
            return None
        cell = ws.cell(row=row, column=col)
        return cell.value
    except Exception:
        return None


def safe_write_cell(ws, row: int, col: int, value):
    """安全写入单元格值。"""
    try:
        cell = ws.cell(row=row, column=col)
        cell.value = value
    except Exception as e:  # noqa: BLE001
        logger.warning(f"写入单元格失败 ({row}, {col}): {e}")


def clear_merges_in_column(ws, col_idx: int):
    """清除指定列的合并单元格。"""
    try:
        ranges_to_remove = []
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_col <= col_idx <= merged_range.max_col:
                ranges_to_remove.append(str(merged_range))

        for range_str in ranges_to_remove:
            ws.unmerge_cells(range_str)
    except Exception as e:  # noqa: BLE001
        logger.warning(f"清除合并单元格失败: {e}")


def safe_delete_file(filepath: str, max_retries: int = 3):
    """安全删除文件，处理文件被占用的情况。"""
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
                time.sleep(0.1 * (i + 1))
                continue
            logger.warning(f"删除文件失败(被占用): {filepath}, 错误: {e}")
            return False
        except Exception as e:  # noqa: BLE001
            logger.warning(f"删除文件失败: {filepath}, 错误: {e}")
            return False

    return False
    ```

- **core/excel_utils.py**
  - language: `python`
  - content:
    ```python
# core/excel_utils.py
"""统一的 Excel 工具：列转换、文件包装、合并单元格复制。"""
import re
import logging
from typing import Union, List, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 模块级工具函数
# ---------------------------------------------------------------------------
def col_letter_to_index(letter: str) -> int:
    return column_index_from_string(letter.upper())


def col_index_to_letter(index: int) -> str:
    return get_column_letter(index)


def validate_column_letters(letters: List[str]) -> bool:
    pattern = re.compile(r'^[A-Z]+$')
    return all(pattern.match(l.upper()) for l in letters)


# ---------------------------------------------------------------------------
# ExcelFile 上下文管理器
# ---------------------------------------------------------------------------
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
        self._workbook = openpyxl.load_workbook(
            self.filepath, data_only=self.data_only
        )
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

    def copy_merge_structure(self, src_col: Union[int, str],
                             dst_col: Union[int, str],
                             start_row: int, end_row: int,
                             sheet=None):
        """把 src_col 的合并单元格结构复制到 dst_col。"""
        ws = sheet or self.active_sheet
        if isinstance(src_col, str):
            src_col = col_letter_to_index(src_col)
        if isinstance(dst_col, str):
            dst_col = col_letter_to_index(dst_col)

        # 收集源列上与 [start_row, end_row] 相交的合并区间
        src_ranges = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= src_col <= merged.max_col:
                if merged.min_row <= end_row and merged.max_row >= start_row:
                    src_ranges.append((merged.min_row, merged.max_row))

        # 清除目标列上已有的合并
        to_remove = []
        for merged in ws.merged_cells.ranges:
            if merged.min_col <= dst_col <= merged.max_col:
                to_remove.append(merged)
        for merged in to_remove:
            ws.unmerge_cells(str(merged))

        # 在目标列重建相同区间
        for min_r, max_r in src_ranges:
            if min_r < max_r:
                ws.merge_cells(
                    start_row=min_r, start_column=dst_col,
                    end_row=max_r, end_column=dst_col,
                )

    @staticmethod
    def read_dataframe(filepath: str,
                       sheet: Union[str, int] = 0,
                       **kwargs) -> pd.DataFrame:
        try:
            return pd.read_excel(filepath, sheet_name=sheet, **kwargs)
        except Exception as e:
            logger.error(f"读取Excel失败 {filepath}: {e}")
            raise ValueError(f"无法读取Excel文件: {e}")

    @staticmethod
    def write_dataframe(df: pd.DataFrame, filepath: str,
                        sheet_name: str = 'Sheet1'):
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return filepath


# ---------------------------------------------------------------------------
# 模块级向后兼容函数（独立于 ExcelFile 类，供旧调用方使用）
# ---------------------------------------------------------------------------
def copy_merge_structure(src_ws, src_col, dst_ws, dst_col, start_row, end_row):
    """把 src_ws 的 src_col 合并结构复制到 dst_ws 的 dst_col。

    这是 ExcelFile.copy_merge_structure 的模块级包装，供 services/excel/matcher.py
    等以函数式调用的旧调用方使用。
    """
    ef = ExcelFile('')
    ef._workbook = src_ws.parent
    ef.copy_merge_structure(
        src_col, dst_col, start_row, end_row, sheet=dst_ws
    )
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
        # 修复 #3: 删除 eval 注入风险，改为原生字符串比较
        if op == '>':   return cell > target
        if op == '<':   return cell < target
        if op == '>=':  return cell >= target
        if op == '<=':  return cell <= target
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

