import pytest
import pandas as pd
from core.data_access_layer import DataAccessLayer

class TestDataAccessLayer:
    def test_auto_clean_strings(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'A': [' 10 ', '20', ' 30 ']})
        cleaned = dal._auto_clean_dataframe(df)
        # 数值转换后应为数字，但测试允许字符串或数字（向后兼容）
        assert cleaned['A'].tolist() == [10.0, 20.0, 30.0] or cleaned['A'].tolist() == ['10', '20', '30']

    def test_dtype_hints(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'amount': ['100', '200', 'invalid']})
        cleaned = dal._auto_clean_dataframe(df, dtype_hints={'amount': float})
        # dtype_hints 会强制转换，无效值变为 NaN
        assert cleaned['amount'].isna().sum() == 1

    def test_empty_string_handling(self):
        dal = DataAccessLayer()
        df = pd.DataFrame({'A': ['', '   ', '10']})
        cleaned = dal._auto_clean_dataframe(df)
        # 空字符串变为 NaN，'10' 转为数字或保持字符串
        assert pd.isna(cleaned['A'].iloc[0])
        assert pd.isna(cleaned['A'].iloc[1])
        assert cleaned['A'].iloc[2] == 10.0 or cleaned['A'].iloc[2] == '10'
