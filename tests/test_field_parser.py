import pytest
from services.field_parser.service import FieldParserService

class TestFieldParserService:
    @pytest.fixture
    def service(self):
        return FieldParserService(similarity_threshold=30)

    def test_parse_field_list_commas(self, service):
        result = service.parse_field_list("姓名, 年龄, 性别")
        assert result == ["姓名", "年龄", "性别"]

    def test_match_exact(self, service):
        result = service.match_headers(["ID", "Name"], ["ID", "Name"])
        assert len(result["matched"]) == 2
        assert result["only_a"] == []
        assert result["only_b"] == []

    def test_match_similar_chinese(self, service):
        result = service.match_headers(["姓名", "年龄"], ["名字", "岁数"])
        assert len(result["matched"]) == 1
        assert result["matched"][0]["standard"] == "姓名"
        assert result["matched"][0]["compare"] == "名字"
        assert len(result["only_a"]) == 1
        assert "年龄" in result["only_a"]
        assert len(result["only_b"]) == 1
        assert "岁数" in result["only_b"]

    def test_match_below_threshold(self, service):
        service.threshold = 0.9
        result = service.match_headers(["姓名"], ["名字"])
        assert len(result["matched"]) == 0
