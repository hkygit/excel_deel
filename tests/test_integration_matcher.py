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
        """目标文件有重复键"""
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'ID'
        ws['B1'] = 'Value'
        ws['A2'] = 1
        ws['B2'] = ''
        ws['A3'] = 1   # 重复键
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

    def test_match_with_duplicate_keys_disabled(self, client, target_file_duplicate, source_file):
        """默认不勾选，所有匹配的行均应填充相同值"""
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            # 不传 allow_duplicate_keys，默认 off
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        # 两行 ID=1 都应该被填充 100
        assert result_df.loc[0, 'Value'] == 100
        assert result_df.loc[1, 'Value'] == 100
        assert result_df.loc[2, 'Value'] == 200   # ID=2

    def test_match_with_duplicate_keys_enabled(self, client, target_file_duplicate, source_file):
        """勾选允许重复键，仅首行填充，其余行为原值（空）"""
        data = {
            'target_file': (target_file_duplicate, 'target_dup.xlsx'),
            'source_file': (source_file, 'source.xlsx'),
            'target_key_cols': 'A',
            'source_match_cols': 'A',
            'source_value_col': 'B',
            'result_col': 'B',
            'allow_duplicate_keys': 'on',   # 勾选
        }
        response = client.post('/match/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        # 第一行 ID=1 填充 100
        assert result_df.loc[0, 'Value'] == 100
        # 第二行 ID=1 应保持为空（原有空值）
        # 注意 pandas 读回空值可能是 NaN，判断用 pd.isna
        val_second = result_df.loc[1, 'Value']
        assert val_second == '' or pd.isna(val_second)
        # 第三行 ID=2 填充 200
        assert result_df.loc[2, 'Value'] == 200
