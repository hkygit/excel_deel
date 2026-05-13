import pandas as pd
from utils.condition_filter import ConditionFilter, Condition

class TestConditionFilter:
    def test_apply_to_dict_eq(self):
        row = {'name': 'Alice', 'age': '30'}
        cond = Condition(field='name', op='=', value='Alice')
        assert ConditionFilter.apply_to_dict(row, [cond]) == True

    def test_apply_to_dict_contains(self):
        row = {'name': 'Alice'}
        cond = Condition(field='name', op='contains', value='Ali')
        assert ConditionFilter.apply_to_dict(row, [cond]) == True

    def test_apply_to_dict_exclude(self):
        row = {'name': 'Alice'}
        cond = Condition(field='name', op='=', value='Alice', exclude=True)
        assert ConditionFilter.apply_to_dict(row, [cond]) == False

    def test_apply_to_dataframe(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie'], 'age': [25, 30, 35]})
        cond = Condition(field='name', op='contains', value='Ali')
        mask = ConditionFilter.apply_to_dataframe(df, [cond])
        assert mask.tolist() == [True, False, False]

    def test_apply_to_dataframe_multiple(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie'], 'age': [25, 30, 35]})
        cond1 = Condition(field='age', op='>', value='30')
        cond2 = Condition(field='name', op='contains', value='Char')
        mask = ConditionFilter.apply_to_dataframe(df, [cond1, cond2])
        assert mask.tolist() == [False, False, True]
