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
from services.excel.filter_engine import FilterEngine
from services.excel.excel_service import ExcelService
from services.excel.calculator import ColumnCalculator
from services.excel.comparator import ExcelComparator
from services.excel.filter import ExcelFilter
from services.excel.aggregator import Aggregator
from services.simplify.service import SimplifyService
from services.field_parser import FieldParserService
from services.checker import CheckerService
from services.convert.service import ConvertService
from services.team_report import TeamReportService
from services.team_report.writer import TeamReportExcelWriter

class Container(containers.DeclarativeContainer):
    config = providers.Singleton(lambda: settings)
    task_storage = providers.Singleton(TaskStorage, status_dir=config.provided.simplify.task_status_folder, output_dir=config.provided.simplify.output_folder)
    task_manager = providers.Singleton(TaskManager, storage=task_storage, max_concurrent=3)
    data_reader = providers.Singleton(DataReader)
    data_writer = providers.Singleton(DataWriter)
    data_formatter = providers.Singleton(DataFormatter)
    file_service = providers.Singleton(FileService)
    compute_engine = providers.Singleton(ComputeEngine)
    filter_engine = providers.Singleton(FilterEngine)
    excel_service = providers.Singleton(ExcelService)
    aggregator = providers.Singleton(Aggregator)
    column_calculator = providers.Singleton(ColumnCalculator)
    excel_comparator = providers.Singleton(ExcelComparator)
    excel_filter = providers.Singleton(ExcelFilter)
    simplify_service = providers.Factory(SimplifyService)
    team_report_service = providers.Factory(
        TeamReportService,
        config=config.provided.team_report
    )
    team_report_writer = providers.Singleton(TeamReportExcelWriter)
    field_parser = providers.Singleton(FieldParserService, similarity_threshold=config.provided.field_parser.similarity_threshold)
    checker_service = providers.Factory(CheckerService, config=config.provided.checker)
    convert_service = providers.Factory(ConvertService)
