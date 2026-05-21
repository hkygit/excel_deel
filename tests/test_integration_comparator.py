# tests/test_integration_comparator.py
import io
import json
import pandas as pd

class TestComparatorIntegration:
    def test_full_compare(self, client, sample_excel_comparison_files):
        f1, f2 = sample_excel_comparison_files
        data = {
            'file1': (f1, 'file1.xlsx'),
            'file2': (f2, 'file2.xlsx'),
            'compare_mode': 'full',
            'key_fields': 'id',
            'compare_fields': 'name,score',
            'sheet1': '',
            'sheet2': ''
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content_disp = response.headers['Content-Disposition']
        # 新文件名格式：compare_file1_对比结果.xlsx
        assert 'compare_file1_' in content_disp and '.xlsx' in content_disp
        result = pd.ExcelFile(io.BytesIO(response.data))
        sheet_names = result.sheet_names
        assert '仅存在于文件2' in sheet_names

    def test_field_match_mode(self, client, sample_excel_comparison_files):
        f1, f2 = sample_excel_comparison_files
        data = {
            'file1': (f1, 'file1.xlsx'),
            'file2': (f2, 'file2.xlsx'),
            'compare_mode': 'field_match',
            'match_field': 'name',
            'sheet1': '',
            'sheet2': ''
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        content_disp = response.headers.get('Content-Disposition', '')
        assert 'filename=' in content_disp

    def test_no_difference_response(self, client):
        import pandas as pd
        from io import BytesIO
        df = pd.DataFrame({'id': [1], 'name': ['A']})
        f1 = BytesIO()
        f2 = BytesIO()
        with pd.ExcelWriter(f1, engine='openpyxl') as w: df.to_excel(w, index=False)
        with pd.ExcelWriter(f2, engine='openpyxl') as w: df.to_excel(w, index=False)
        f1.seek(0); f2.seek(0)
        data = {
            'file1': (f1, 'same1.xlsx'),
            'file2': (f2, 'same2.xlsx'),
            'compare_mode': 'field_match',
            'match_field': 'name',
        }
        response = client.post('/compare/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        resp_json = response.get_json()
        assert resp_json is not None
        assert resp_json['message'] == '两个文件中该字段值完全匹配，无差异'
