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
