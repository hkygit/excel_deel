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
        assert 'matched_' in response.headers['Content-Disposition']
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
