import io
import time
import pandas as pd
import pytest
import zipfile

class TestConvertIntegration:
    def test_convert_excel_to_csv(self, client):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
        buf.seek(0)
        data = {
            'file': (buf, 'test.xlsx'),
            'target_format': 'csv',
            'encoding': 'utf-8-sig',
            'separator': ','
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        max_attempts = 10
        status = None
        for _ in range(max_attempts):
            status_resp = client.get(f'/convert/status/{task_id}')
            status = status_resp.get_json()['data']
            if status['status'] == 'completed':
                break
            time.sleep(0.5)
        assert status is not None and status['status'] == 'completed'
        download_resp = client.get(f'/convert/download/{task_id}')
        assert download_resp.status_code == 200
        content = download_resp.data
        if content[:2] == b'PK':
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                with zf.open(zf.namelist()[0]) as f:
                    result_df = pd.read_csv(f)
        else:
            result_df = pd.read_csv(io.BytesIO(content))
        assert list(result_df.columns) == ['A', 'B']
        assert len(result_df) == 2

    def test_convert_csv_to_excel(self, client):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        buf = io.StringIO()
        df.to_csv(buf, index=False, encoding='utf-8-sig')
        buf.seek(0)
        data = {
            'file': (io.BytesIO(buf.getvalue().encode('utf-8')), 'test.csv'),
            'target_format': 'xlsx',
            'encoding': 'utf-8-sig',
            'separator': ','
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        for _ in range(10):
            status_resp = client.get(f'/convert/status/{task_id}')
            status = status_resp.get_json()['data']
            if status['status'] == 'completed':
                break
            time.sleep(0.5)
        assert status['status'] == 'completed'
        download_resp = client.get(f'/convert/download/{task_id}')
        assert download_resp.status_code == 200
        result_df = pd.read_excel(io.BytesIO(download_resp.data))
        assert list(result_df.columns) == ['A', 'B']

    def test_missing_file_error(self, client):
        response = client.post('/convert/process', data={'target_format': 'csv'})
        assert response.status_code == 400
        assert not response.get_json()['success']
        assert '请上传文件' in response.get_json()['error']

    def test_invalid_target_format(self, client):
        df = pd.DataFrame({"A": [1]})
        buf = io.BytesIO()
        df.to_excel(buf, index=False, engine='openpyxl')
        buf.seek(0)
        data = {
            'file': (buf, 'test.xlsx'),
            'target_format': 'pdf',
            'encoding': 'utf-8-sig',
        }
        response = client.post('/convert/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 200
        task_id = response.get_json()['data']['task_id']
        for _ in range(10):
            status = client.get(f'/convert/status/{task_id}').get_json()['data']
            if status['status'] == 'failed':
                break
            time.sleep(0.5)
        assert status['status'] == 'failed'
