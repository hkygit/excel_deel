import io
import pytest
import pandas as pd
from openpyxl import Workbook

class TestMatcherIntegration:
    @pytest.fixture
    def target_file(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value'
        ws['A2'] = 1
        ws['B2'] = ''
        ws['A3'] = 2
        ws['B3'] = ''
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    @pytest.fixture
    def target_file_duplicate(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value'
        ws['A2'] = 1
        ws['B2'] = ''
        ws['A3'] = 1
        ws['B3'] = ''
        ws['A4'] = 2
        ws['B4'] = ''
        f = io.BytesIO()
        wb.save(f)
        f.seek(0)
        return f

    @pytest.fixture
    def source_file(self):
        df = pd.DataFrame({'ID': [1, 2], 'Score': [100, 200]})
        f = io.BytesIO()
        with pd.ExcelWriter(f, engine='openpyxl') as w: df.to_excel(w, index=False)
        f.seek(0)
        return f

    def test_match_process(self, client, target_file, source_file):
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content_disp = response.headers.get('Content-Disposition', '')
        assert 'attachment' in content_disp
        # 新格式：match_原文件名_匹配结果.xlsx
        assert 'match_target_' in content_disp or 'match_target.xlsx' in content_disp
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert result_df.loc[0, 'Value'] == 100
        assert result_df.loc[1, 'Value'] == 200

    def test_debug_mode(self, client, target_file, source_file):
        data = {
            'target_file': (target_file, 'target.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
        }
        response = client.post('/match/debug', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        assert 'config' in response.data.decode('utf-8')

    def test_match_with_duplicate_keys_disabled(self, client, target_file_duplicate, source_file):
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        resp_json = response.get_json()
        assert resp_json is not None
        assert not resp_json['success']
        assert '重复' in resp_json['error']

    def test_match_with_duplicate_keys_enabled(self, client, target_file_duplicate, source_file):
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            'allow_duplicate_keys': 'on',
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert result_df.loc[0, 'Value'] == 100
        val_second = result_df.loc[1, 'Value']
        assert val_second == '' or pd.isna(val_second)
        assert result_df.loc[2, 'Value'] == 200
