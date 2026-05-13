import io
import pytest
import pandas as pd
import numpy as np

class TestSummarizerIntegration:
    @pytest.fixture
    def excel_files(self):
        df1 = pd.DataFrame({'机构': ['A', 'A'], '金额': [100, 200], '日期': ['2025-01-01', '2025-01-02']})
        df2 = pd.DataFrame({'机构': ['B'], '金额': [300], '日期': ['2025-01-03']})
        f1 = io.BytesIO()
        f2 = io.BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df1.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df2.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)
        return [f1, f2]

    def test_summarize_basic(self, client, excel_files):
        data = {
            'files': [(excel_files[0], 'file1.xlsx'), (excel_files[1], 'file2.xlsx')],
            'fields': '机构\n金额',
            'unique_keys': '机构',
            'sum_fields': '金额'
        }
        response = client.post('/summary/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))
        assert len(result_df) == 2
        assert result_df.loc[result_df['机构'] == 'A', '金额'].iloc[0] == 300
        assert result_df.loc[result_df['机构'] == 'B', '金额'].iloc[0] == 300

    def test_missing_fields_error(self, client):
        response = client.post('/summary/process', data={})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '没有选择文件' in response.get_json()['error']

    # --------------- 唯一键包含空值场景 ---------------
    def test_unique_keys_with_empty_values(self, client):
        # 构造两个文件，唯一键为 机构+日期，其中一行的日期为空
        df1 = pd.DataFrame({'机构': ['A', 'A', 'B'], '日期': ['2025-01', '', '2025-02'], '金额': [100, 200, 300]})
        df2 = pd.DataFrame({'机构': ['B'], '日期': ['2025-02'], '金额': [400]})

        f1 = io.BytesIO()
        f2 = io.BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df1.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df2.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)

        data = {
            'files': [(f1, 'file1.xlsx'), (f2, 'file2.xlsx')],
            'fields': '机构\n日期\n金额',
            'unique_keys': '机构,日期',
            'sum_fields': '金额'
        }
        response = client.post('/summary/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        result_df = pd.read_excel(io.BytesIO(response.data))

        # 预期结果：
        # A + ''  -> 200
        # A + '2025-01' -> 100
        # B + '2025-02' -> 700
        # 共3行
        assert len(result_df) == 3

        # 空值在 Excel 中读回可能是 NaN，统一处理后再比较
        # 将日期列的空值统一转换为空字符串
        result_df['日期'] = result_df['日期'].fillna('')

        row_a_empty = result_df[(result_df['机构'] == 'A') & (result_df['日期'] == '')]
        assert len(row_a_empty) == 1
        assert row_a_empty.iloc[0]['金额'] == 200

        row_b = result_df[(result_df['机构'] == 'B') & (result_df['日期'] == '2025-02')]
        assert len(row_b) == 1
        assert row_b.iloc[0]['金额'] == 700
