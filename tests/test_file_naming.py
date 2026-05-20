# tests/test_file_naming.py
from core.file_naming import generate_output_filename

class TestFileNaming:
    def test_basic(self):
        name = generate_output_filename('subtotal', '数据.xlsx')
        assert name == 'subtotal_数据_小计.xlsx'

    def test_with_path(self):
        name = generate_output_filename('match', '/path/to/目标文件.xlsx')
        assert name == 'match_目标文件_小计.xlsx'

    def test_custom_suffix(self):
        name = generate_output_filename('filter', 'data.csv', suffix='_汇总')
        assert name == 'filter_data_汇总.xlsx'
