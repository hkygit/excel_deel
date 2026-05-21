# tests/test_integration_splitter.py
import io, zipfile, pytest
from openpyxl import Workbook

class TestSplitterIntegration:
    @pytest.fixture
    def multi_sheet_excel(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = 'Data'
        ws1['A1'] = 'Category'
        ws1['B1'] = 'Value'
        ws1['A2'] = 'A'
        ws1['B2'] = 10
        ws1['A3'] = 'B'
        ws1['B3'] = 20
        ws2 = wb.create_sheet('Extra')
        ws2['A1'] = 'Category'
        ws2['A2'] = 'A'
        ws2['A3'] = 'B'
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    def test_split_basic(self, client, multi_sheet_excel):
        data = {
            'file': (multi_sheet_excel, 'test.xlsx'),
            'split_column': 'Category',
            'include_values': '',
            'exclude_values': '',
            'skip_sheets': '2'
        }
        response = client.post('/splitter/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        # 检查 Content-Disposition 是否包含统一前缀
        content_disp = response.headers.get('Content-Disposition', '')
        assert 'splitter_test' in content_disp or 'splitter' in content_disp
        zip_data = io.BytesIO(response.data)
        with zipfile.ZipFile(zip_data) as zf:
            names = zf.namelist()
            assert any('A.xlsx' in n for n in names)
            assert any('B.xlsx' in n for n in names)
