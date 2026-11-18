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
