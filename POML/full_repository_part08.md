- **tests/test_no_legacy_compat.py**
  - language: `python`
  - content:
    ```python
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
    ```

- **tests/test_remove_result_column_expr.py**
  - language: `python`
  - content:
    ```python
# tests/test_remove_result_column_expr.py
"""
针对 ColumnCalculatorProcessor._remove_result_column_from_expr 的回归测试。

核心目标：
    验证"用 token 匹配替换子串 replace"这一修复，确保：
    - 恰好等于结果列的独立标识符被移除
    - 只是包含结果列字母的标识符（如 'AA', 'AB'）不被破坏
    - 移除时连带清理邻近的二元运算符，不产生语法上非法的表达式
"""
import pytest

from services.excel.calculator_processor import ColumnCalculatorProcessor
from services.excel.compute_engine import ComputeEngine


def _make_processor(expression: str) -> ColumnCalculatorProcessor:
    """构造一个仅用于测试 _remove_result_column_from_expr 的处理器实例。"""
    return ColumnCalculatorProcessor(
        filepath="dummy.xlsx",      # 不用于本测试
        header_row=1,
        expression=expression,
        result_col_input="Z",        # 不用于本测试
        merge_by_cols_str="",
        compute_engine=ComputeEngine(),
    )


class TestRemoveResultColumnFromExpr:
    # ---------- 快速路径：无匹配 ----------
    def test_no_match_returns_original(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("D") == "A+B"

    def test_empty_expression(self):
        p = _make_processor("")
        assert p._remove_result_column_from_expr("A") == ""

    def test_empty_result_col(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("") == "A+B"

    # ---------- 关键回归：不得破坏子串标识符 ----------
    def test_aa_identifier_not_corrupted_by_removing_a(self):
        """'AA+B' 移除 'A'，'AA' 必须完整保留。"""
        p = _make_processor("AA+B")
        assert p._remove_result_column_from_expr("A") == "AA+B"

    def test_ab_identifier_not_corrupted_by_removing_a(self):
        """'AB+C' 移除 'A'，'AB' 必须完整保留。"""
        p = _make_processor("AB+C")
        assert p._remove_result_column_from_expr("A") == "AB+C"

    def test_ab_identifier_not_corrupted_by_removing_b(self):
        """'AB+C' 移除 'B'，'AB' 必须完整保留。"""
        p = _make_processor("AB+C")
        assert p._remove_result_column_from_expr("B") == "AB+C"

    def test_multi_letter_identifiers_untouched(self):
        """'AB+AC' 里没有任何 token 恰好等于 'A'。"""
        p = _make_processor("AB+AC")
        assert p._remove_result_column_from_expr("A") == "AB+AC"

    # ---------- 标准移除：独立标识符 ----------
    def test_removes_standalone_at_start(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("A") == "B"

    def test_removes_standalone_at_end(self):
        p = _make_processor("A+B")
        assert p._remove_result_column_from_expr("B") == "A"

    def test_removes_standalone_in_middle(self):
        p = _make_processor("A+B-C")
        assert p._remove_result_column_from_expr("B") == "A-C"

    def test_removes_in_longer_chain(self):
        p = _make_processor("A+B+C+D")
        assert p._remove_result_column_from_expr("C") == "A+B+D"

    # ---------- 括号与一元运算符 ----------
    def test_removes_with_parens(self):
        p = _make_processor("(A+B)*C")
        assert p._remove_result_column_from_expr("B") == "(A)*C"

    def test_removes_inside_nested_parens(self):
        p = _make_processor("((A+B)-C)")
        assert p._remove_result_column_from_expr("B") == "((A)-C)"

    # ---------- 边界情况 ----------
    def test_whole_expression_is_result_column(self):
        """表达式只有一个 token，且正是结果列本身。"""
        p = _make_processor("A")
        assert p._remove_result_column_from_expr("A") == "0"

    def test_lowercase_input_is_normalized(self):
        """小写输入应被正确识别并移除。"""
        p = _make_processor("a+b")
        assert p._remove_result_column_from_expr("A") == "B"

    def test_number_literals_are_not_mistaken_for_identifiers(self):
        """数字 token 不应被当作标识符。'A*2+B' 移除 'A' → '2+B'。"""
        p = _make_processor("A*2+B")
        assert p._remove_result_column_from_expr("A") == "2+B"

    def test_decimal_number_literal_untouched(self):
        p = _make_processor("A*1.5+B")
        assert p._remove_result_column_from_expr("A") == "1.5+B"
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


> ✅ 共导出 142 个代码文件。
