import pandas as pd
import pytest
from core.filter_engine import parse_conditions_from_form, ConditionEngine
from core.column_resolver import resolve_columns
from core.exceptions import ValidationError

engine = ConditionEngine()

class TestParseFiltersFromForm:
    def test_empty(self):
        form = {}
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact == []
        assert fuzzy == []

    def test_single_equals(self):
        form = {
            'condition_field_0': 'name',
            'condition_op_0': '=',
            'condition_value_0': 'Alice'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(exact) == 1
        assert exact[0]['field'] == 'name'
        assert exact[0]['op'] == '='
        assert exact[0]['values'] == ['Alice']

    def test_multi_values_space(self):
        form = {
            'condition_field_0': 'dept',
            'condition_op_0': '=',
            'condition_value_0': 'HR  IT'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact[0]['values'] == ['HR', 'IT']

    def test_not_equal(self):
        form = {
            'condition_field_0': 'score',
            'condition_op_0': '!=',
            'condition_value_0': '100'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert exact[0]['op'] == '!='

    def test_contains(self):
        form = {
            'condition_field_0': 'name',
            'condition_op_0': '包含',
            'condition_value_0': 'Al'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(fuzzy) == 1
        assert fuzzy[0]['op'] == '包含'

    def test_compatibility_old_format(self):
        form = {
            'condition_field_0': 'dept',
            'condition_values_0': 'HR IT'
        }
        exact, fuzzy = parse_conditions_from_form(form)
        assert len(exact) == 1
        assert exact[0]['values'] == ['HR', 'IT']
        assert exact[0]['op'] == '='

class TestApplyConditions:
    def test_equals(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie']})
        exact = [{'field': 'name', 'op': '=', 'values': ['Alice', 'Charlie']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['name']) == ['Alice', 'Charlie']

    def test_not_equals(self):
        df = pd.DataFrame({'score': [100, 90, 80]})
        exact = [{'field': 'score', 'op': '!=', 'values': ['90']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['score']) == [100, 80]

    def test_greater_than(self):
        df = pd.DataFrame({'age': [25, 30, 35]})
        exact = [{'field': 'age', 'op': '>', 'values': ['28']}]
        result = engine.apply_to_dataframe(df, exact, [])
        assert list(result['age']) == [30, 35]

    def test_contains(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Alfred']})
        fuzzy = [{'field': 'name', 'op': '包含', 'value': 'Al'}]
        result = engine.apply_to_dataframe(df, [], fuzzy)
        assert list(result['name']) == ['Alice', 'Alfred']

    def test_not_contains(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Alfred']})
        fuzzy = [{'field': 'name', 'op': '不包含', 'value': 'Al'}]
        result = engine.apply_to_dataframe(df, [], fuzzy)
        assert list(result['name']) == ['Bob']

    def test_combined(self):
        df = pd.DataFrame({
            'dept': ['HR', 'IT', 'HR', 'IT'],
            'salary': [5000, 6000, 5500, 7000],
            'name': ['Alice', 'Bob', 'Charlie', 'David']
        })
        exact = [{'field': 'dept', 'op': '=', 'values': ['HR']}]
        fuzzy = [{'field': 'name', 'op': '包含', 'value': 'li'}]
        result = engine.apply_to_dataframe(df, exact, fuzzy)
        assert len(result) == 2
        assert set(result['name']) == {'Alice', 'Charlie'}

class TestResolveColumns:
    def test_by_name(self):
        df = pd.DataFrame({'A': [1], 'B': [2], 'C': [3]})
        res = resolve_columns(df, 'A,C')
        assert list(res.columns) == ['A', 'C']

    def test_by_letter(self):
        df = pd.DataFrame({'X': [1], 'Y': [2], 'Z': [3]})
        res = resolve_columns(df, 'A,C')
        assert list(res.columns) == ['X', 'Z']

    def test_mixed(self):
        df = pd.DataFrame({'Name': ['a'], 'Age': [1], 'Score': [2]})
        res = resolve_columns(df, 'Name,B')
        assert list(res.columns) == ['Name', 'Age']

    def test_invalid_raises(self):
        df = pd.DataFrame({'X': [1]})
        with pytest.raises(ValidationError):
            resolve_columns(df, 'NonExist')
