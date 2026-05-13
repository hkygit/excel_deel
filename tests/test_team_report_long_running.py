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
