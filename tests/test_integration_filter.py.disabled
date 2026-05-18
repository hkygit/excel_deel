import io
import zipfile
import pandas as pd

class TestFilterIntegration:
    def test_column_and_row_limit(self, client):
        df = pd.DataFrame({'Name': ['Alice', 'Bob', 'Charlie'], 'Age': [25, 30, 35], 'Score': [90, 80, 85]})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'files': (buf, 'test.xlsx'),
            'action': 'generate_excel',
            'selected_columns': 'Name,Score',
            'uploadRowsLimit': '2'
        }
        resp = client.post('/filter/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        result = pd.read_excel(io.BytesIO(resp.data), sheet_name='筛选数据')
        assert list(result.columns) == ['Name', 'Score']
        assert len(result) == 2

    def test_operator_conditions(self, client):
        df = pd.DataFrame({'Product': ['A', 'B', 'C', 'D'], 'Price': [10, 20, 15, 30], 'Category': ['X', 'Y', 'X', 'Y']})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'files': (buf, 'test.xlsx'),
            'action': 'generate_excel',
            'condition_field_0': 'Price',
            'condition_op_0': '>',
            'condition_value_0': '15',
            'condition_field_1': 'Category',
            'condition_op_1': '!=',
            'condition_value_1': 'Y'
        }
        resp = client.post('/filter/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        result = pd.read_excel(io.BytesIO(resp.data), sheet_name='筛选数据')
        assert len(result) == 0

    def test_contains_and_get_values(self, client):
        df = pd.DataFrame({'Name': ['Alice', 'Bob', 'Charlie'], 'Dept': ['HR', 'IT', 'HR']})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'files': (buf, 'test.xlsx'),
            'action': 'get_values',
            'target_field': 'Name',
            'condition_field_0': 'Name',
            'condition_op_0': '包含',
            'condition_value_0': 'li'
        }
        resp = client.post('/filter/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        text = resp.data.decode('utf-8')
        assert 'Alice' in text
        assert 'Charlie' in text
        assert 'Bob' not in text

    def test_old_format_still_works(self, client):
        df = pd.DataFrame({'A': [10, 20, 10], 'B': ['x', 'y', 'z']})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'files': (buf, 'test.xlsx'),
            'action': 'generate_excel',
            'condition_field_0': 'A',
            'condition_values_0': '10'
        }
        resp = client.post('/filter/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        result = pd.read_excel(io.BytesIO(resp.data), sheet_name='筛选数据')
        assert len(result) == 2
        assert list(result['A']) == [10, 10]

    def test_truncated_content_mode(self, client):
        csv_text = "Name,Age,Score\nAlice,25,90\nBob,30,80\nCharlie,35,85"
        data = {
            'truncated_content': csv_text,
            'action': 'generate_excel',
            'selected_columns': 'Name,Score',
            'uploadRowsLimit': '2',
            'condition_field_0': 'Age',
            'condition_op_0': '>',
            'condition_value_0': '28'
        }
        resp = client.post('/filter/process', data=data)
        assert resp.status_code == 200
        result = pd.read_excel(io.BytesIO(resp.data), sheet_name='筛选数据')
        assert len(result) == 1
        assert list(result.columns) == ['Name', 'Score']
        assert result.iloc[0]['Name'] == 'Bob'

    def test_multi_file_zip_download(self, client):
        df1 = pd.DataFrame({'A': [1, 2, 3], 'B': ['x', 'y', 'z']})
        df2 = pd.DataFrame({'A': [4, 5, 6], 'B': ['a', 'b', 'c']})
        buf1 = io.BytesIO()
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf1, engine='openpyxl') as w: df1.to_excel(w, index=False)
        with pd.ExcelWriter(buf2, engine='openpyxl') as w: df2.to_excel(w, index=False)
        buf1.seek(0)
        buf2.seek(0)
        data = {
            'files': [(buf1, 'file1.xlsx'), (buf2, 'file2.xlsx')],
            'action': 'generate_excel',
            'selected_columns': 'A',
            'uploadRowsLimit': '2',
            'condition_field_0': 'A',
            'condition_op_0': '>',
            'condition_value_0': '1'
        }
        resp = client.post('/filter/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert resp.mimetype == 'application/zip'
        with zipfile.ZipFile(io.BytesIO(resp.data), 'r') as zf:
            names = zf.namelist()
            assert len(names) == 2
            with zf.open(names[0]) as f:
                res_df = pd.read_excel(f)
                assert len(res_df) == 1
                assert list(res_df['A'])[0] == 2
