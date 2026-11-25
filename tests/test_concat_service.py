# tests/test_concat_service.py
"""ConcatService 单元测试。"""
import os
import pandas as pd
import pytest

from core.config import ConcatConfig
from core.exceptions import ValidationError
from services.concat import ConcatService


@pytest.fixture
def cfg():
    return ConcatConfig()


@pytest.fixture
def service(cfg):
    return ConcatService(cfg)


def _write_xlsx(path, df):
    df.to_excel(path, index=False)
    return str(path)


class TestBasic:
    def test_basic_concat(self, service, tmp_path):
        df = pd.DataFrame({
            "身份证号": ["340721195905140923"] * 3,
            "疾病名称": ["高血压", "高血压伴并发症", "结核病"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "身份证号", ["疾病名称"])
        # 用 dtype=str 读回：避免 pandas 把 18 位身份证号推断成 int64；
        # 同时能真正验证精度——若服务内部丢了精度（如存成 float64），
        # dtype=str 读回也会得到 ...924 而非 ...923，测试依然会失败。
        result = pd.read_excel(out, dtype=str)
        assert len(result) == 1
        assert result.iloc[0]["身份证号"] == "340721195905140923"
        # 自然排序：结核病 < 高血压 < 高血压伴并发症（按码点，国 < 高）
        assert isinstance(result.iloc[0]["疾病名称"], str)
        assert "," in result.iloc[0]["疾病名称"]

    def test_multi_column_independent(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A", "A", "A"],
            "名称": ["x", "y", "x"],
            "代码": ["1", "2", "3"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["名称", "代码"])
        result = pd.read_excel(out)
        assert len(result) == 1
        # 名称去重后应为 x,y
        assert result.iloc[0]["名称"] == "x,y"
        # 代码自然排序 1,2,3
        assert result.iloc[0]["代码"] == "1,2,3"

    def test_output_filename(self, service, tmp_path):
        df = pd.DataFrame({"ID": ["A"], "V": ["x"]})
        p = _write_xlsx(tmp_path / "original_name.xlsx", df)
        out = service.execute(p, "original_name.xlsx", "ID", ["V"])
        assert os.path.basename(out) == "original_name_分组拼接.xlsx"


class TestNaturalSort:
    def test_numeric_sort(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A"] * 3,
            "值": ["10", "9", "100"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"])
        result = pd.read_excel(out)
        # 自然排序：9,10,100
        assert result.iloc[0]["值"] == "9,10,100"

    def test_descending(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A"] * 3,
            "值": ["10", "9", "100"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(
            p, "in.xlsx", "ID", ["值"], sort_ascending=False
        )
        result = pd.read_excel(out)
        assert result.iloc[0]["值"] == "100,10,9"


class TestDedup:
    def test_keep_first(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A"] * 3,
            "值": ["b", "a", "b"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"], dup_keep="first")
        result = pd.read_excel(out)
        # 排序后: a, b, b；去重保留首次: a, b；再排序: a, b
        assert result.iloc[0]["值"] == "a,b"

    def test_keep_last(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A"] * 3,
            "值": ["a", "b", "a"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"], dup_keep="last")
        result = pd.read_excel(out)
        # 排序后: a, a, b；保留末次: a(位置1), b(位置2)；再排序: a, b
        assert result.iloc[0]["值"] == "a,b"


class TestEmpty:
    def test_empty_string_filtered(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A", "A", "A"],
            "值": ["x", "", "y"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"])
        result = pd.read_excel(out)
        assert result.iloc[0]["值"] == "x,y"

    def test_all_empty_returns_empty(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A", "A"],
            "值": ["", ""],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"])
        result = pd.read_excel(out)
        # 全空 -> 空串（Excel 读回可能是 NaN）
        assert pd.isna(result.iloc[0]["值"]) or result.iloc[0]["值"] == ""

    def test_literal_nan_not_filtered(self, service, tmp_path):
        """字符串 'nan'/'None' 不应被当作空值过滤。"""
        df = pd.DataFrame({
            "ID": ["A", "A"],
            "值": ["nan", "None"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"])
        result = pd.read_excel(out)
        assert "nan" in result.iloc[0]["值"]
        assert "None" in result.iloc[0]["值"]


class TestGroupEmpty:
    def test_group_col_empty_raises(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A", "", "B"],
            "值": ["x", "y", "z"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        with pytest.raises(ValidationError, match="包含空值"):
            service.execute(p, "in.xlsx", "ID", ["值"])


class TestValidation:
    def test_missing_group_col(self, service, tmp_path):
        df = pd.DataFrame({"ID": ["A"], "V": ["x"]})
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        with pytest.raises(ValidationError):
            service.execute(p, "in.xlsx", "", ["V"])

    def test_missing_concat_cols(self, service, tmp_path):
        df = pd.DataFrame({"ID": ["A"], "V": ["x"]})
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        with pytest.raises(ValidationError):
            service.execute(p, "in.xlsx", "ID", [])

    def test_nonexistent_col(self, service, tmp_path):
        df = pd.DataFrame({"ID": ["A"], "V": ["x"]})
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        with pytest.raises(ValidationError, match="不存在"):
            service.execute(p, "in.xlsx", "ID", ["NotExist"])

    def test_cell_overflow(self, cfg, tmp_path):
        cfg.max_cell_chars = 5
        service = ConcatService(cfg)
        df = pd.DataFrame({
            "ID": ["A"] * 3,
            "值": ["aaaa", "bbbb", "cccc"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        with pytest.raises(ValidationError, match="超过Excel单单元格上限"):
            service.execute(p, "in.xlsx", "ID", ["值"])


class TestColumnResolution:
    def test_letter_resolution(self, service, tmp_path):
        df = pd.DataFrame({
            "身份证号": ["A", "A"],
            "疾病名称": ["x", "y"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "A", ["B"])
        result = pd.read_excel(out)
        # 输出列名应为解析后的原列名，而非 A/B
        assert "身份证号" in result.columns
        assert "疾病名称" in result.columns

    def test_separator_multi_char(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A", "A"],
            "值": ["x", "y"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"], separator=" | ")
        result = pd.read_excel(out)
        assert result.iloc[0]["值"] == "x | y"

    def test_separator_empty(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A", "A"],
            "值": ["x", "y"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"], separator="")
        result = pd.read_excel(out)
        assert result.iloc[0]["值"] == "xy"


class TestOtherColumns:
    def test_other_col_takes_first_row(self, service, tmp_path):
        df = pd.DataFrame({
            "ID": ["A", "A"],
            "其他": ["first", "second"],
            "值": ["x", "y"],
        })
        p = _write_xlsx(tmp_path / "in.xlsx", df)
        out = service.execute(p, "in.xlsx", "ID", ["值"])
        result = pd.read_excel(out)
        assert result.iloc[0]["其他"] == "first"
