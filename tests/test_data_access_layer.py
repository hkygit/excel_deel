import pytest
import pandas as pd
import numpy as np
from core.data_access_layer import DataAccessLayer

class TestDataAccessLayer:
    @pytest.fixture
    def dal(self):
        return DataAccessLayer()

    def test_auto_clean_strings(self, dal):
        df = pd.DataFrame({
            'A': [' 10 ', '20', '30'],
            'B': ['x', 'y', 'z']
        })
        cleaned = dal._auto_clean_dataframe(df)
        assert cleaned['A'].tolist() == [10.0, 20.0, 30.0]
        assert cleaned['B'].tolist() == ['x', 'y', 'z']

    def test_dtype_hints(self, dal):
        df = pd.DataFrame({'amount': ['100', '200', 'invalid'], 'count': ['1', '2', '3']})
        cleaned = dal._auto_clean_dataframe(df, dtype_hints={'count': int})
        assert cleaned['amount'].isna().sum() == 1
        assert cleaned['count'].dtype == int

    def test_empty_string_handling(self, dal):
        df = pd.DataFrame({'A': ['', '  ', '10']})
        cleaned = dal._auto_clean_dataframe(df)
        assert pd.isna(cleaned['A'].iloc[0])
        assert pd.isna(cleaned['A'].iloc[1])
        assert cleaned['A'].iloc[2] == 10.0
