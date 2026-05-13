import pandas as pd
import pytest
from services.data.reader import DataReader
from services.data.writer import DataWriter
from services.data.formatter import DataFormatter

class TestDataReader:
    def test_read_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2], 'B': ['x', 'y']})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        result = DataReader.read_excel(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2], 'B': ['x', 'y']})
        path = tmp_path / 'test.csv'
        df.to_csv(path, index=False)
        result = DataReader.read_csv(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_auto_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        result = DataReader.read_auto(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_auto_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'test.csv'
        df.to_csv(path, index=False)
        result = DataReader.read_auto(str(path))
        pd.testing.assert_frame_equal(result, df)

    def test_read_auto_unsupported(self, tmp_path):
        path = tmp_path / 'test.txt'
        path.write_text('hello')
        with pytest.raises(ValueError, match='不支持的文件类型'):
            DataReader.read_auto(str(path))

class TestDataWriter:
    def test_write_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.xlsx'
        DataWriter.write_excel(df, str(path))
        result = pd.read_excel(path)
        pd.testing.assert_frame_equal(result, df)

    def test_write_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.csv'
        DataWriter.write_csv(df, str(path))
        result = pd.read_csv(path)
        pd.testing.assert_frame_equal(result, df)

    def test_write_auto_excel(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.xlsx'
        DataWriter.write_auto(df, str(path))
        result = pd.read_excel(path)
        pd.testing.assert_frame_equal(result, df)

    def test_write_auto_csv(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2]})
        path = tmp_path / 'out.csv'
        DataWriter.write_auto(df, str(path))
        result = pd.read_csv(path)
        pd.testing.assert_frame_equal(result, df)

class TestDataFormatter:
    def test_round_numeric_columns(self):
        df = pd.DataFrame({'A': [1.234, 2.345], 'B': ['x', 'y']})
        result = DataFormatter.round_numeric_columns(df, decimals=2)
        assert result['A'].tolist() == [1.23, 2.35]
        assert result['B'].tolist() == ['x', 'y']

    def test_safe_format_dataframe(self):
        df = pd.DataFrame({'A': [1, None], 'B': ['x', None]})
        result = DataFormatter.safe_format_dataframe(df)
        assert result['A'].iloc[0] == 1
        assert pd.isna(result['A'].iloc[1])
        assert result['B'].tolist() == ['x', '']

    def test_fillna_empty(self):
        df = pd.DataFrame({'A': [1, None], 'B': ['x', None]})
        result = DataFormatter.fillna_empty(df)
        assert result['A'].tolist() == [1, '']
        assert result['B'].tolist() == ['x', '']
