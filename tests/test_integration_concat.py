# tests/test_integration_concat.py
"""分组拼接 HTTP 集成测试。"""
import io
import pandas as pd
import pytest
from openpyxl import Workbook


def _make_xlsx_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


class TestConcatIntegration:
    def test_index_page(self, client):
        resp = client.get("/concat/")
        assert resp.status_code == 200
        assert "分组拼接" in resp.get_data(as_text=True)

    def test_basic_process(self, client):
        df = pd.DataFrame({
            "身份证号": ["340721195905140923"] * 3,
            "疾病名称": ["高血压", "高血压伴并发症", "结核病"],
        })
        data = {
            "file": (_make_xlsx_bytes(df), "test.xlsx"),
            "group_col": "身份证号",
            "concat_cols": "疾病名称",
            "separator": ",",
            "sort_ascending": "true",
            "dup_keep": "first",
        }
        resp = client.post(
            "/concat/process", data=data, content_type="multipart/form-data"
        )
        assert resp.status_code == 200, resp.data
        disp = resp.headers.get("Content-Disposition", "")
        assert "test_分组拼接.xlsx" in disp or "_" in disp
        result = pd.read_excel(io.BytesIO(resp.data))
        assert len(result) == 1

    def test_missing_file(self, client):
        resp = client.post(
            "/concat/process",
            data={"group_col": "A", "concat_cols": "B"},
        )
        assert resp.status_code == 400
        assert "请上传文件" in resp.get_json()["error"]

    def test_missing_group_col(self, client):
        df = pd.DataFrame({"A": ["x"], "B": ["y"]})
        data = {
            "file": (_make_xlsx_bytes(df), "t.xlsx"),
            "group_col": "",
            "concat_cols": "B",
        }
        resp = client.post(
            "/concat/process", data=data, content_type="multipart/form-data"
        )
        assert resp.status_code == 400

    def test_missing_concat_cols(self, client):
        df = pd.DataFrame({"A": ["x"], "B": ["y"]})
        data = {
            "file": (_make_xlsx_bytes(df), "t.xlsx"),
            "group_col": "A",
            "concat_cols": "",
        }
        resp = client.post(
            "/concat/process", data=data, content_type="multipart/form-data"
        )
        assert resp.status_code == 400

    def test_group_col_with_empty_value(self, client):
        df = pd.DataFrame({
            "ID": ["A", "", "B"],
            "值": ["x", "y", "z"],
        })
        data = {
            "file": (_make_xlsx_bytes(df), "t.xlsx"),
            "group_col": "ID",
            "concat_cols": "值",
        }
        resp = client.post(
            "/concat/process", data=data, content_type="multipart/form-data"
        )
        assert resp.status_code == 400
        assert "包含空值" in resp.get_json()["error"]

    def test_large_file_rejected(self, client):
        # 构造 > 10MB 的假 xlsx（大小检查在读之前）
        big = b"PK\x03\x04" + b"\x00" * (11 * 1024 * 1024)
        data = {
            "file": (io.BytesIO(big), "big.xlsx"),
            "group_col": "A",
            "concat_cols": "B",
        }
        resp = client.post(
            "/concat/process", data=data, content_type="multipart/form-data"
        )
        assert resp.status_code == 400
        assert "大文件" in resp.get_json()["error"]

    def test_invalid_extension(self, client):
        data = {
            "file": (io.BytesIO(b"hello"), "test.txt"),
            "group_col": "A",
            "concat_cols": "B",
        }
        resp = client.post(
            "/concat/process", data=data, content_type="multipart/form-data"
        )
        assert resp.status_code == 400
        assert "不支持" in resp.get_json()["error"]

    def test_separator_escape(self, client):
        df = pd.DataFrame({"ID": ["A", "A"], "值": ["x", "y"]})
        data = {
            "file": (_make_xlsx_bytes(df), "t.xlsx"),
            "group_col": "ID",
            "concat_cols": "值",
            "separator": "\\t",  # 字面量 \t，后端应转成 Tab
        }
        resp = client.post(
            "/concat/process", data=data, content_type="multipart/form-data"
        )
        assert resp.status_code == 200
        result = pd.read_excel(io.BytesIO(resp.data))
        assert result.iloc[0]["值"] == "x\ty"
