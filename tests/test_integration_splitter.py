import io
import zipfile
import pytest
from openpyxl import Workbook
import pandas as pd

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
        assert 'application/zip' in response.content_type
        zip_data = io.BytesIO(response.data)
        with zipfile.ZipFile(zip_data) as zf:
            names = zf.namelist()
            assert any('A.xlsx' in n for n in names)
            assert any('B.xlsx' in n for n in names)

    def test_include_values_filter(self, client, multi_sheet_excel):
        data = {
            'file': (multi_sheet_excel, 'test.xlsx'),
            'split_column': 'Category',
            'include_values': 'A',
            'exclude_values': '',
        }
        response = client.post('/splitter/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        zip_data = io.BytesIO(response.data)
        with zipfile.ZipFile(zip_data) as zf:
            names = zf.namelist()
            assert any('A.xlsx' in n for n in names)
            assert not any('B.xlsx' in n for n in names)

    def test_missing_split_column_error(self, client, multi_sheet_excel):
        data = {
            'file': (multi_sheet_excel, 'test.xlsx'),
            'split_column': 'NonExistent',
        }
        response = client.post('/splitter/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '不在第一个工作表的表头中' in response.get_json()['error']

    def test_filter_and_skip_sheets(self, client, multi_sheet_excel):
        data = {
            'file': (multi_sheet_excel, 'test.xlsx'),
            'split_column': 'Category',
            'condition_field_0': 'Category',
            'condition_op_0': '=',
            'condition_value_0': 'A',
            'include_values': '',
            'exclude_values': '',
            'skip_sheets': '2'
        }
        response = client.post('/splitter/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        zip_data = io.BytesIO(response.data)
        with zipfile.ZipFile(zip_data) as zf:
            names = zf.namelist()
            assert len(names) == 1
            assert 'A.xlsx' in names[0]
