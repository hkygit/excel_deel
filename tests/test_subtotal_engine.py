# tests/test_subtotal_engine.py
import pytest
import pandas as pd
import os
from services.subtotal.engine import SubtotalEngine

class TestSubtotalEngine:
    @pytest.fixture
    def sample_excel(self, tmp_path):
        df = pd.DataFrame({
            'A': ['X', 'X', 'Y'],   # 列名变为字母，方便列字母引用
            'B': [10, 20, 30],
            'C': [1, 2, 3]
        })
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        return str(path), 'test.xlsx'

    def test_basic_subtotal_sum(self, sample_excel):
        filepath, fname = sample_excel
        engine = SubtotalEngine(filepath, 'A', 'B,C', {'default': 'sum'}, original_filename=fname)
        out_path = engine.process()
        assert os.path.basename(out_path).startswith('subtotal_test')
        wb = pd.read_excel(out_path, sheet_name=None)
        ws = wb['Sheet1']
        assert len(ws) == 5
        assert ws.iloc[2]['B'] == 30
        assert ws.iloc[2]['C'] == 3

    def test_dimension_col(self, sample_excel):
        filepath, fname = sample_excel
        engine = SubtotalEngine(filepath, 'A', 'B', {'default': 'sum'},
                                dimension_cols='C', original_filename=fname)
        out_path = engine.process()
        wb = pd.read_excel(out_path, sheet_name=None)
        ws = wb['Sheet1']
        assert ws.iloc[2]['C'] == '合计'
