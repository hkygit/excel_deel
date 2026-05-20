# tests/test_integration_subtotal.py
import io
import pandas as pd

class TestSubtotalIntegration:
    def test_sync_endpoint(self, client):
        # 使用字母列名
        df = pd.DataFrame({
            'A': ['X', 'X', 'Y'],
            'B': [10, 20, 30],
            'C': [100, 200, 300]
        })
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B,C',
            'agg_func_default': 'sum',
            'sheets_mode': '1'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'subtotal_test' in resp.headers.get('Content-Disposition', '')
        result_df = pd.read_excel(io.BytesIO(resp.data), sheet_name=None)
        ws = result_df['Sheet1']
        assert len(ws) == 5
        assert ws.iloc[2]['B'] == 30
        assert ws.iloc[2]['C'] == 300

    def test_missing_dimension_col(self, client):
        df = pd.DataFrame({'A': [1]})
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'A',
            'dimension_cols': 'Z'   # 不存在的列字母
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 400
        assert '无效' in resp.get_json()['error'] or '超出' in resp.get_json()['error']

    def test_async_large_file(self, client):
        df = pd.DataFrame({'A': ['X', 'Y'] * 3000, 'B': range(6000)})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test_large.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B',
            'agg_func_default': 'sum'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'task_id' in resp.get_json()
