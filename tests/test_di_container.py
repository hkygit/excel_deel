# tests/test_di_container.py
"""DI 容器（P1-2）回归测试。

覆盖：
- Container 可以实例化，且默认读取全局 settings
- task_storage 的 status_dir 与 output_dir 分离
- task_storage / task_manager 为 Singleton（同一实例）
- checker_service / team_report_service 为 Factory（不同实例）
- 规范化的子配置 provider 返回正确的配置对象
- 向后兼容：config.provided.xxx 依旧可用
- Container(config=...) 覆盖配置后所有依赖同步更新
- field_parser 的 similarity_threshold 来自 field_parser_config
"""
import os

import pytest
from dependency_injector import providers

from core.config import Settings, settings as global_settings
from core.di import Container, _task_status_dir
from core.task.storage import TaskStorage
from core.task.manager import TaskManager
from services.checker.service import CheckerService
from services.team_report.service import TeamReportService


def _make_settings(tmp_path):
    """构造一个纯净的 Settings（不读 .env），输出目录指向临时目录。"""
    return Settings(
        _env_file=None,
        output_folder=str(tmp_path / "outputs"),
        upload_folder=str(tmp_path / "uploads"),
    )


@pytest.fixture
def container(tmp_path):
    """使用临时目录的 Settings 覆盖 container 的 config provider。"""
    s = _make_settings(tmp_path)
    return Container(config=providers.Singleton(lambda: s))


# ============================================================================
# 基础实例化
# ============================================================================
class TestContainerInstantiation:
    def test_container_can_be_instantiated(self):
        c = Container()
        assert c is not None
        # 不覆盖时，config 返回全局 settings
        assert c.config() is global_settings

    def test_default_config_returns_global_settings(self):
        c = Container()
        assert c.config() is global_settings


# ============================================================================
# task_storage 目录分离（本次修复的核心）
# ============================================================================
class TestTaskStorageDirs:
    def test_status_dir_separated_from_output_dir(self, container, tmp_path):
        """关键回归：status_dir 必须与 output_dir 不同。"""
        storage = container.task_storage()
        output_dir = str(tmp_path / "outputs")
        expected_status_dir = os.path.join(output_dir, "task_status")

        assert storage.output_dir == output_dir
        assert storage.status_dir == expected_status_dir
        assert storage.status_dir != storage.output_dir

    def test_status_dir_created_on_init(self, container, tmp_path):
        """首次访问 task_storage 时，两个目录都应被创建。"""
        storage = container.task_storage()
        assert os.path.isdir(storage.status_dir)
        assert os.path.isdir(storage.output_dir)

    def test_task_status_dir_helper(self, tmp_path):
        """_task_status_dir 的纯函数行为。"""
        s = _make_settings(tmp_path)
        result = _task_status_dir(s)
        assert result == os.path.join(
            str(tmp_path / "outputs"), "task_status"
        )

    def test_db_file_lives_in_status_dir(self, container):
        storage = container.task_storage()
        # tasks.db 应位于 status_dir 下
        assert storage.db_path.startswith(storage.status_dir)
        assert os.path.basename(storage.db_path) == "tasks.db"


# ============================================================================
# Singleton / Factory 语义
# ============================================================================
class TestProviderScopes:
    def test_task_storage_is_singleton(self, container):
        assert container.task_storage() is container.task_storage()

    def test_task_manager_is_singleton(self, container):
        assert container.task_manager() is container.task_manager()

    def test_task_manager_shares_storage(self, container):
        storage = container.task_storage()
        manager = container.task_manager()
        assert manager.storage is storage

    def test_checker_service_is_factory(self, container):
        assert container.checker_service() is not container.checker_service()

    def test_team_report_service_is_factory(self, container):
        assert (
            container.team_report_service()
            is not container.team_report_service()
        )

    def test_field_parser_is_singleton(self, container):
        assert container.field_parser() is container.field_parser()


# ============================================================================
# 规范化子配置 provider
# ============================================================================
class TestSubConfigProviders:
    def test_checker_config(self, container):
        cfg = container.checker_config()
        assert cfg is not None
        assert hasattr(cfg, "similarity_threshold")
        assert hasattr(cfg, "duplicate_action")

    def test_team_report_config(self, container):
        cfg = container.team_report_config()
        assert cfg is not None
        assert cfg.daily_order_target == 1150
        assert cfg.daily_apply_target == 120

    def test_field_parser_config(self, container):
        cfg = container.field_parser_config()
        assert cfg is not None
        assert hasattr(cfg, "similarity_threshold")

    def test_splitter_config(self, container):
        assert container.splitter_config() is not None

    def test_convert_config(self, container):
        assert container.convert_config() is not None

    def test_sub_config_stable_between_calls(self, container):
        """Callable provider 每次返回 settings 上的同一子配置实例。"""
        assert container.checker_config() is container.checker_config()


# ============================================================================
# 向后兼容：config.provided.xxx
# ============================================================================
class TestBackwardCompat:
    """保留 container.config.provided.xxx 的旧调用方式。"""

    def test_config_provided_checker(self, container):
        cfg = container.config.provided.checker()
        assert cfg is not None
        assert hasattr(cfg, "similarity_threshold")

    def test_config_provided_team_report(self, container):
        cfg = container.config.provided.team_report()
        assert cfg is not None

    def test_provided_and_new_provider_same_threshold(self, container):
        """两种访问方式应反映同一份配置。"""
        old = container.config.provided.checker()
        new = container.checker_config()
        assert old.similarity_threshold == new.similarity_threshold
        assert old.duplicate_action == new.duplicate_action


# ============================================================================
# 依赖注入的 wiring
# ============================================================================
class TestWiring:
    def test_checker_service_uses_checker_config(self, tmp_path):
        """checker_service 的 config 应来自 checker_config。"""
        s = _make_settings(tmp_path)
        s.checker.similarity_threshold = 999
        c = Container(config=providers.Singleton(lambda: s))
        svc = c.checker_service()
        assert isinstance(svc, CheckerService)
        assert svc.config.similarity_threshold == 999

    def test_team_report_service_uses_team_report_config(self, tmp_path):
        """team_report_service 的 config 应来自 team_report_config。"""
        s = _make_settings(tmp_path)
        s.team_report.daily_order_target = 777
        c = Container(config=providers.Singleton(lambda: s))
        svc = c.team_report_service()
        assert isinstance(svc, TeamReportService)
        assert svc.config.daily_order_target == 777

    def test_field_parser_uses_field_parser_config(self, tmp_path):
        """field_parser 的 similarity_threshold 来自 field_parser_config。"""
        s = _make_settings(tmp_path)
        s.field_parser.similarity_threshold = 42
        c = Container(config=providers.Singleton(lambda: s))
        parser = c.field_parser()
        # FieldParserService.__init__ 会做 / 100
        assert parser.threshold == 0.42

    def test_override_config_propagates(self, tmp_path):
        """覆盖 config 后，所有子 provider 应反映新配置。"""
        s = _make_settings(tmp_path)
        s.checker.duplicate_action = "keep_first"
        c = Container(config=providers.Singleton(lambda: s))

        assert c.checker_config().duplicate_action == "keep_first"
        assert c.config.provided.checker().duplicate_action == "keep_first"
