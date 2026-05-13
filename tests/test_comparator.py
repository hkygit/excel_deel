import pandas as pd
import tempfile
from services.excel.comparator import ExcelComparator

def create_test_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    return filename

class TestExcelComparator:
    def test_full_compare_identical(self):
        data = [{'id': 1, 'name': 'A', 'score': 100}, {'id': 2, 'name': 'B', 'score': 90}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data, f1.name)
            create_test_excel(data, f2.name)
            result = ExcelComparator.full_compare(f1.name, f2.name, ['id'], ['name', 'score'])
            assert result['only_in_1'].empty
            assert result['only_in_2'].empty
            assert result['diff'] is None

    def test_full_compare_diff(self):
        data1 = [{'id': 1, 'name': 'A', 'score': 100}]
        data2 = [{'id': 1, 'name': 'A', 'score': 90}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data1, f1.name)
            create_test_excel(data2, f2.name)
            result = ExcelComparator.full_compare(f1.name, f2.name, ['id'], ['score'])
            assert result['diff'] is not None
            assert len(result['diff']) == 1
            assert result['diff'].iloc[0]['文件1_score'] == 100
            assert result['diff'].iloc[0]['文件2_score'] == 90

    def test_field_match(self):
        data1 = [{'name': 'A'}, {'name': 'B'}, {'name': 'C'}]
        data2 = [{'name': 'A'}, {'name': 'B'}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data1, f1.name)
            create_test_excel(data2, f2.name)
            only1, only2 = ExcelComparator.field_match(f1.name, f2.name, 'name')
            assert len(only1) == 1
            assert only1.iloc[0]['name'] == 'C'
            assert only2.empty
