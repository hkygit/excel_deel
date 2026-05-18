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
