import pytest
import pandas as pd
import tempfile
import os
from services.convert.service import ConvertService
from core.exceptions import ValidationError

class TestConvertService:
    @pytest.fixture
    def sample_xlsx(self, tmp_path):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False)
        return str(path)

    @pytest.fixture
    def sample_multi_sheet_xlsx(self, tmp_path):
        df1 = pd.DataFrame({"ID": [1], "Name": ["Alice"]})
        df2 = pd.DataFrame({"ID": [2], "Name": ["Bob"]})
        path = tmp_path / "multi.xlsx"
        with pd.ExcelWriter(path, engine='openpyxl') as w:
            df1.to_excel(w, sheet_name='Sheet1', index=False)
            df2.to_excel(w, sheet_name='Sheet2', index=False)
        return str(path)

    @pytest.fixture
    def sample_csv(self, tmp_path):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        path = tmp_path / "test.csv"
        df.to_csv(path, index=False, encoding='utf-8-sig')
        return str(path)

    def test_excel_to_csv_single_sheet(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='csv')
        result_path = service.execute()
        assert result_path.endswith('.csv')
        df = pd.read_csv(result_path)
        assert list(df.columns) == ['A', 'B']
        assert len(df) == 2
        os.unlink(result_path)

    def test_excel_to_csv_multi_sheet(self, sample_multi_sheet_xlsx):
        service = ConvertService(file_path=sample_multi_sheet_xlsx, target_format='csv')
        result_path = service.execute()
        assert result_path.endswith('.zip')
        import zipfile
        with zipfile.ZipFile(result_path, 'r') as zf:
            names = zf.namelist()
            assert len(names) == 2
            assert any('Sheet1.csv' in n for n in names)
            assert any('Sheet2.csv' in n for n in names)
        os.unlink(result_path)

    def test_csv_to_excel(self, sample_csv):
        service = ConvertService(file_path=sample_csv, target_format='xlsx')
        result_path = service.execute()
        assert result_path.endswith('.xlsx')
        df = pd.read_excel(result_path)
        assert list(df.columns) == ['A', 'B']
        assert len(df) == 2
        os.unlink(result_path)

    def test_unsupported_conversion(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='xlsx')
        with pytest.raises(ValidationError, match='不支持的转换'):
            service.execute()

    def test_invalid_file(self, tmp_path):
        invalid_file = tmp_path / "invalid.xlsx"
        invalid_file.write_text("not an excel file")
        service = ConvertService(file_path=str(invalid_file), target_format='csv')
        with pytest.raises(Exception):
            service.execute()

    def test_precheck(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='csv')
        result = service.precheck()
        assert 'total_rows' in result
        assert result['total_rows'] == 2
