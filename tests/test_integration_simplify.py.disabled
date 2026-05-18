import json
import pytest
from unittest.mock import patch

class TestSimplifyIntegration:
    def test_submit_task_and_poll(self, client, sample_csv_header_file, sample_csv_data_file):
        params = {
            'key_cols': 'key',
            'keep_cols': 'value',
            'non_sum_fields': 'other',
            'filters': [],
            'output_format': 'xlsx',
            'column_mode': 'name'
        }
        data = {
            'header_file': (sample_csv_header_file, 'header.csv'),
            'data_files': (sample_csv_data_file, 'data.csv'),
            'params': json.dumps(params)
        }

        app = client.application
        original_tm = app.container.task_manager()
        def sync_start(task_id, target, args):
            target(*args)
        with patch.object(original_tm, 'start_task', side_effect=sync_start):
            response = client.post('/simplify/process', data=data, content_type='multipart/form-data')
            assert response.status_code == 200
            result = response.get_json()
            assert result['success'] == True
            task_id = result['data']['task_id']

            # 任务已同步完成
            status_resp = client.get(f'/simplify/status/{task_id}')
            status = status_resp.get_json()
            assert status['success'] == True
            assert status['data']['status'] == 'completed'
            assert status['data']['total_input_rows'] > 0
            assert status['data']['total_output_rows'] > 0

    def test_download_result(self, client, sample_csv_header_file, sample_csv_data_file):
        params = {
            'key_cols': 'key',
            'keep_cols': 'value',
            'non_sum_fields': 'other',
            'filters': [],
            'output_format': 'xlsx',
            'column_mode': 'name'
        }
        data = {
            'header_file': (sample_csv_header_file, 'header.csv'),
            'data_files': (sample_csv_data_file, 'data.csv'),
            'params': json.dumps(params)
        }
        app = client.application
        original_tm = app.container.task_manager()
        def sync_start(task_id, target, args):
            target(*args)
        with patch.object(original_tm, 'start_task', side_effect=sync_start):
            response = client.post('/simplify/process', data=data)
            result = response.get_json()
            assert result['success'] == True
            task_id = result['data']['task_id']

            status = client.get(f'/simplify/status/{task_id}').get_json()
            assert status['data']['status'] == 'completed'

            download_resp = client.get(f'/simplify/download/{task_id}')
            assert download_resp.status_code == 200
            assert 'attachment' in download_resp.headers['Content-Disposition']

    def test_invalid_csv_format(self, client):
        from io import BytesIO
        fake_txt = BytesIO(b'not a csv')
        data = {'header_file': (fake_txt, 'header.txt'), 'params': '{}'}
        response = client.post('/simplify/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        assert 'CSV 格式' in response.get_json()['error']
