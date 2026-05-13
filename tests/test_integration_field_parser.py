import json
import pytest

class TestFieldParserIntegration:
    def test_analyze_success(self, client):
        data = {
            "standard_headers": "姓名, 年龄, 性别",
            "compare_headers": "名字, 岁数, 性别"
        }
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 200
        result = response.get_json()
        assert len(result["matched"]) == 2

    def test_analyze_missing_field(self, client):
        data = {"standard_headers": "姓名"}
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 400
        result = response.get_json()
        assert "compare_headers" in result['error']

    def test_analyze_empty(self, client):
        data = {"standard_headers": "", "compare_headers": ""}
        response = client.post('/field-parser/analyze', data=json.dumps(data), content_type='application/json')
        assert response.status_code == 400
