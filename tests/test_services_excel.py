import pandas as pd
import pytest
from services.excel.compute_engine import ComputeEngine
from core.filter_engine import FilterEngine
from services.excel.aggregator import Aggregator
from core.filter_engine import Condition

class TestComputeEngine:
    def test_evaluate_expression(self):
        values = {'A': 10, 'B': 5}
        assert ComputeEngine.evaluate_expression('A+B', values) == 15
        assert ComputeEngine.evaluate_expression('A*B', values) == 50

    def test_extract_variables(self):
        assert ComputeEngine.extract_variables('A+B-C') == {'A', 'B', 'C'}

    def test_simple_aggregate(self):
        df = pd.DataFrame({'G': ['X', 'X', 'Y'], 'V': [10, 20, 5]})
        result = ComputeEngine.simple_aggregate(df, ['G'], ['V'], 'sum')
        assert len(result) == 2
        assert result.loc[result['G'] == 'X', 'V'].iloc[0] == 30

    def test_aggregate_by_keys(self):
        df = pd.DataFrame({'key': ['A', 'A', 'B'], 'sum': [10, 20, 5], 'other': ['first', 'second', 'other']})
        result = ComputeEngine.aggregate_by_keys(df, ['key'], ['sum'], ['other'])
        assert len(result) == 2
        assert result.loc[result['key'] == 'A', 'sum'].iloc[0] == 30
        assert result.loc[result['key'] == 'A', 'other'].iloc[0] == 'first'

class TestFilterEngine:
    def test_filter_dataframe_with_condition_objects(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [25, 30]})
        cond = Condition(field='age', op='>', value=28)
        result = FilterEngine.filter_dataframe(df, [cond])
        assert len(result) == 1
        assert result.iloc[0]['name'] == 'Bob'

    def test_filter_dataframe_with_multi_values(self):
        df = pd.DataFrame({'dept': ['HR', 'IT', 'HR'], 'name': ['A', 'B', 'C']})
        result = FilterEngine.filter_dataframe_with_multi_values(df, exact_conditions=[('dept', ['HR'])], fuzzy_conditions=[('name', 'A')])
        assert len(result) == 1
        assert result.iloc[0]['name'] == 'A'

    def test_filter_dict(self):
        row = {'name': 'Alice', 'age': 30}
        cond = Condition(field='age', op='>', value=25)
        assert FilterEngine.filter_dict(row, [cond]) is True

    def test_create_condition(self):
        cond = FilterEngine.create_condition('age', '>', 25, exclude=True)
        assert cond.field == 'age'
        assert cond.op == '>'
        assert cond.value == 25
        assert cond.exclude is True

class TestAggregator:
    def test_group_aggregate(self):
        df = pd.DataFrame({'G': ['X', 'X'], 'V': [10, 20]})
        result = Aggregator.group_aggregate(df, ['G'], {'V': 'sum'})
        assert result.iloc[0]['V'] == 30

    def test_aggregate_by_keys_with_other_fields(self):
        df = pd.DataFrame({'key': ['A', 'A'], 'sum': [10, 20], 'other': ['first', 'second']})
        result = Aggregator.aggregate_by_keys(df, ['key'], ['sum'], ['other'])
        assert result.iloc[0]['sum'] == 30
        assert result.iloc[0]['other'] == 'first'
