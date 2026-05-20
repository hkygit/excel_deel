# tests/test_header_detection.py
import pytest
import pandas as pd
from core.header_detection import SimpleExcelHeaderDetector, detect_header_start, HeaderDetectionConfig
import openpyxl

class TestHeaderDetection:
    @pytest.fixture
    def simple_file(self, tmp_path):
        df = pd.DataFrame({'姓名': ['Alice', 'Bob'], '年龄': [30, 25]})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        return str(path)

    def test_basic_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(simple_file)
        assert result.header_rows == [0]
        assert result.data_start_row == 1
        assert '姓名' in result.column_names
        assert result.confidence > 0.5

    def test_detect_header_start(self, simple_file):
        start = detect_header_start(simple_file)
        assert start == 2

    def test_skip_total_row(self, tmp_path):
        df = pd.DataFrame({'名称': ['X', '合计'], '金额': [10, 30]})
        path = tmp_path / 'skip.xlsx'
        df.to_excel(path, index=False)
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(path)
        assert result.header_rows == [0]

    def test_workbook_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        wb = openpyxl.load_workbook(simple_file)
        result1 = detector.detect_from_workbook(wb)
        result2 = detector.detect_from_workbook(wb)
        assert result1.header_rows == [0]
        assert result2.metrics.get('workbook_cache_hit') is True
        wb.close()
