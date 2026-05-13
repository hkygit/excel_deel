import pytest
import pandas as pd
import os
from services.checker.service import CheckerService
from core.exceptions import ValidationError
from core.config import CheckConfig  # 引入配置类


class TestCheckerService:
    @pytest.fixture
    def sample_files_a(self, tmp_path):
        df1 = pd.DataFrame({"ID": [1, 2], "Name": ["Alice", "Bob"], "Score": [100, 90]})
        df2 = pd.DataFrame({"ID": [3, 4], "Name": ["Charlie", "David"], "Score": [80, 70]})
        f1 = tmp_path / "a1.xlsx"
        f2 = tmp_path / "a2.xlsx"
        df1.to_excel(f1, index=False)
        df2.to_excel(f2, index=False)
        return [str(f1), str(f2)]

    @pytest.fixture
    def sample_files_b(self, tmp_path):
        df = pd.DataFrame({"ID": [1, 2, 5], "Name": ["Alice", "Bob", "Eve"], "Score": [100, 95, 60]})
        f = tmp_path / "b.xlsx"
        df.to_excel(f, index=False)
        return [str(f)]

    def test_execute_basic(self, sample_files_a, sample_files_b):
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            duplicate_action="error"
        )
        result_path = service.execute()
        assert os.path.exists(result_path)
        result = pd.read_excel(result_path, sheet_name=None)
        assert "差异数据" in result
        assert len(result["差异数据"]) == 1
        assert "标准群独有" in result
        assert len(result["标准群独有"]) == 2
        assert "比较群独有" in result
        assert len(result["比较群独有"]) == 1
        os.unlink(result_path)

    def test_duplicate_key_error(self, sample_files_a, sample_files_b, tmp_path):
        df_dup = pd.DataFrame({"ID": [1, 1], "Name": ["Alice", "Alice"], "Score": [100, 200]})
        dup_file = tmp_path / "dup.xlsx"
        df_dup.to_excel(dup_file, index=False)
        service = CheckerService(
            files_a=[str(dup_file)],
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}],
            duplicate_action="error"
        )
        with pytest.raises(ValidationError, match="标准群主键重复"):
            service.execute()

    def test_filter_in_standard(self, sample_files_a, sample_files_b):
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            filters_a=[{"field": "Score", "op": ">", "value": 90}],
            duplicate_action="error"
        )
        result_path = service.execute()
        result = pd.read_excel(result_path, sheet_name=None)
        assert "比较群独有" in result
        assert len(result["比较群独有"]) == 2
        os.unlink(result_path)

    def test_execute_with_sqlite_fallback(self, sample_files_a, sample_files_b, monkeypatch):
        """将 max_memory_rows 设为 0 强制使用 SQLite 回退模式"""
        # 构造一个配置使 max_memory_rows=0，触发 SQLite 路径
        low_mem_config = CheckConfig(max_memory_rows=0)
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            duplicate_action="error",
            config=low_mem_config
        )
        result_path = service.execute()
        assert os.path.exists(result_path)
        result = pd.read_excel(result_path, sheet_name=None)
        assert "差异数据" in result
        os.unlink(result_path)
