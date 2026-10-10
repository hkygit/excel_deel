# tests/test_no_legacy_compat.py
"""回归测试：确认历史兼容层已被清除，且公开 API 未受影响。

设计意图：
- 防止未来有人"顺手"把删除的类/方法加回来
- 同时保证删除后真正在用的公开 API 依然可用
"""
import ast
import inspect

import pytest


class TestRemovedLegacyLayers:
    # ---------- 兼容类 ----------
    def test_aggregate_calculator_removed(self):
        """AggregateCalculator 兼容包装应已删除。"""
        import services.excel.aggregator as agg_mod
        assert not hasattr(agg_mod, 'AggregateCalculator'), (
            "AggregateCalculator 是历史兼容层，不应存在于 services.excel.aggregator"
        )

    # ---------- 兼容静态方法 ----------
    def test_excel_file_legacy_static_methods_removed(self):
        """ExcelFile 的 3 个兼容静态方法应已删除。"""
        from core.excel_utils import ExcelFile
        legacy_methods = [
            'copy_cell_style',
            'copy_merge_structure_static',
            'get_cell_value_static',
        ]
        for name in legacy_methods:
            assert not hasattr(ExcelFile, name), (
                f"ExcelFile.{name} 是历史兼容层，应已被删除"
            )

    # ---------- blueprints/__init__ ----------
    def test_blueprints_init_does_not_import_blueprints(self):
        """blueprints/__init__.py 不应包含任何 import 语句。

        不能用 hasattr(blueprints, 'xxx') 检查：
            Python 包的机制决定了任何 ``import blueprints.calculator``
            之后，该子模块会自动成为父包的属性。运行 pytest 时
            conftest 会加载 app.py，而 app.py 又 import 了
            blueprints.* 子模块，所以 hasattr 必然为 True——与
            __init__.py 是否主动 import 无关。

        改为通过 AST 检查源码里是否存在 Import / ImportFrom 节点，
        这才是真正能反映"__init__.py 是否主动 import"的判据。
        """
        import blueprints as bp_module

        src = inspect.getsource(bp_module)
        tree = ast.parse(src)
        import_nodes = [
            node for node in ast.walk(tree)
            if isinstance(node, (ast.Import, ast.ImportFrom))
        ]
        assert import_nodes == [], (
            f"blueprints/__init__.py 不应包含 import 语句，"
            f"但发现了 {len(import_nodes)} 个：{src!r}"
        )

    # ---------- DI provider ----------
    def test_di_container_no_placeholder_providers(self):
        """DI Container 不应保留无用的占位 provider。"""
        from core.di import Container
        legacy_providers = ['excel_service', 'column_calculator']
        for name in legacy_providers:
            assert not hasattr(Container, name), (
                f"Container.{name} 是占位 provider，应已被删除"
            )


class TestPublicApiStillAvailable:
    """删除兼容层后，真正被使用的 API 必须保持可用。"""

    # ---------- core/excel_utils.py ----------
    def test_excel_file_public_methods(self):
        from core.excel_utils import ExcelFile
        for name in [
            'load', 'close', 'save', 'active_sheet', 'get_sheet',
            'get_headers', 'get_cell_value', 'copy_merge_structure',
            'read_dataframe', 'write_dataframe',
        ]:
            assert hasattr(ExcelFile, name), f"ExcelFile.{name} 不应被删除"

    def test_excel_utils_module_functions(self):
        from core.excel_utils import (
            col_letter_to_index, col_index_to_letter,
            validate_column_letters, copy_merge_structure,
        )
        assert callable(col_letter_to_index)
        assert callable(col_index_to_letter)
        assert callable(validate_column_letters)
        assert callable(copy_merge_structure)

    def test_copy_merge_structure_signature(self):
        """模块级 copy_merge_structure 的签名应保持 6 参数。"""
        from core.excel_utils import copy_merge_structure
        sig = inspect.signature(copy_merge_structure)
        params = list(sig.parameters.keys())
        assert params == [
            'src_ws', 'src_col', 'dst_ws', 'dst_col', 'start_row', 'end_row'
        ]

    # ---------- services/excel/aggregator.py ----------
    def test_aggregator_public_methods(self):
        from services.excel.aggregator import Aggregator
        for name in ['group_aggregate', 'aggregate_by_keys', 'simple_aggregate']:
            assert hasattr(Aggregator, name), f"Aggregator.{name} 不应被删除"

    # ---------- DI Container 公开 provider ----------
    def test_di_container_core_providers(self):
        from core.di import Container
        for name in [
            'config', 'task_storage', 'task_manager',
            'file_service', 'compute_engine',
            'team_report_service', 'checker_service', 'convert_service',
            'field_parser',
            'splitter_config', 'checker_config', 'field_parser_config',
            'team_report_config', 'convert_config',
        ]:
            assert hasattr(Container, name), f"Container.{name} 不应被删除"


class TestFunctionalityUnchanged:
    """删除兼容层不应影响功能。"""

    def test_aggregator_still_works(self):
        import pandas as pd
        from services.excel.aggregator import Aggregator
        df = pd.DataFrame({'G': ['X', 'X', 'Y'], 'V': [1, 2, 3]})
        result = Aggregator.simple_aggregate(df, ['G'], ['V'], 'sum')
        assert len(result) == 2
        assert result.loc[result['G'] == 'X', 'V'].iloc[0] == 3

    def test_excel_file_roundtrip(self, tmp_path):
        import pandas as pd
        from core.excel_utils import ExcelFile
        df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        src = tmp_path / 'in.xlsx'
        dst = tmp_path / 'out.xlsx'
        df.to_excel(src, index=False)

        with ExcelFile(src) as ef:
            headers = ef.get_headers()
            assert headers == ['A', 'B']
            v = ef.get_cell_value(2, 'A')
            assert v == 1
            ef.save(dst)

        assert dst.exists()
        roundtrip = pd.read_excel(dst)
        pd.testing.assert_frame_equal(roundtrip, df)
