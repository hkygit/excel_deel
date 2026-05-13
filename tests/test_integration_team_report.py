# tests/test_integration_team_report.py
import io, time, pandas as pd, pytest

class TestTeamReportIntegration:
    def test_generate_with_monthly(self, client):
        mapping_df = pd.DataFrame({"队员ID":["1","2"],"队伍名称":["骑士队","米兰队"],"队长姓名":["袁队","安队"],"司管":["方行","任翔"]})
        bm = io.BytesIO(); mapping_df.to_excel(bm,index=False); bm.seek(0)
        today_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[10,15],"报单数":[2,3],"转单数":[0,0],"新桔完单数":[5,6],"在线天数":[1,0]})
        bt = io.BytesIO(); today_df.to_excel(bt,index=False); bt.seek(0)
        yesterday_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[9,12],"报单数":[1,2],"转单数":[0,0],"新桔完单数":[4,5],"在线天数":[1,0]})
        by = io.BytesIO(); yesterday_df.to_excel(by,index=False); by.seek(0)
        monthly_df = pd.DataFrame({"司机ID":["1","2"],"司机状态":["已上岗","已上岗"],"总完单数":[100,150],"报单数":[20,30],"转单数":[5,10],"新桔完单数":[50,60]})
        bmon = io.BytesIO(); monthly_df.to_excel(bmon,index=False); bmon.seek(0)
        data = {
            "mapping_file": (bm, "m.xlsx"),
            "today_file": (bt, "t.xlsx"),
            "yesterday_file": (by, "y.xlsx"),
            "monthly_files": (bmon, "monthly.xlsx"),
            "daily_order_target": "1150",
            "daily_apply_target": "120",
            "monthly_order_target": "35650",
            "monthly_apply_target": "3720"
        }
        resp = client.post("/team-report/generate", data=data, content_type="multipart/form-data")
        assert resp.status_code == 200
        tid = resp.get_json()["task_id"]
        for _ in range(15):
            s = client.get(f"/team-report/status/{tid}").get_json()
            if s["status"] == "completed":
                break
            time.sleep(0.5)
        dl = client.get(f"/team-report/download/{tid}")
        assert dl.status_code == 200 and len(dl.data) > 0

    def test_generate_without_monthly(self, client):
        mapping_df = pd.DataFrame({"队员ID":["1"],"队伍名称":["骑士队"],"队长姓名":["袁队"],"司管":["方行"]})
        bm = io.BytesIO(); mapping_df.to_excel(bm,index=False); bm.seek(0)
        today_df = pd.DataFrame({"司机ID":["1"],"司机状态":["已上岗"],"总完单数":[10],"报单数":[2],"转单数":[0],"新桔完单数":[5],"在线天数":[1]})
        bt = io.BytesIO(); today_df.to_excel(bt,index=False); bt.seek(0)
        data = {"mapping_file":(bm,"m.xlsx"),"today_file":(bt,"t.xlsx")}
        resp = client.post("/team-report/generate", data=data, content_type="multipart/form-data")
        assert resp.status_code == 200
        tid = resp.get_json()["task_id"]
        for _ in range(10):
            s = client.get(f"/team-report/status/{tid}").get_json()
            if s["status"] == "completed":
                break
            time.sleep(0.3)
        dl = client.get(f"/team-report/download/{tid}")
        assert dl.status_code == 200
