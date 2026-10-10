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
