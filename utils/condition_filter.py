# utils/condition_filter.py
"""向后兼容层：委托给 core.filter_engine.ConditionEngine"""
import pandas as pd
from typing import List, Union, Tuple
from core.filter_engine import ConditionEngine, parse_conditions_from_form

class Condition:
    def __init__(self, field, op, value, exclude=False):
        self.field = field
        self.op = op
        self.value = value
        self.exclude = exclude

class ConditionFilter:
    @staticmethod
    def apply_to_dict(row, conditions):
        engine = ConditionEngine()
        for c in conditions:
            if isinstance(c, Condition):
                op = c.op
                if op in ('contains', '包含'): op = '包含'
                elif op in ('not_contains', '不包含'): op = '不包含'
                if op in ('包含', '不包含'):
                    fuzzy = [{'field': c.field, 'op': op, 'value': str(c.value)}]
                    exact = []
                else:
                    exact = [{'field': c.field, 'op': op, 'values': [str(c.value)]}]
                    fuzzy = []
                result = engine.apply_to_row(row, exact, fuzzy)
                if c.exclude:
                    result = not result
                if not result:
                    return False
        return True

    @staticmethod
    def apply_to_dataframe(df, conditions):
        """返回布尔 mask"""
        mask = pd.Series(True, index=df.index)
        engine = ConditionEngine()
        for c in conditions:
            if isinstance(c, Condition):
                op = c.op
                if op in ('contains', '包含'): op = '包含'
                elif op in ('not_contains', '不包含'): op = '不包含'
                if op in ('包含', '不包含'):
                    fuzzy = [{'field': c.field, 'op': op, 'value': str(c.value)}]
                    exact = []
                else:
                    exact = [{'field': c.field, 'op': op, 'values': [str(c.value)]}]
                    fuzzy = []
                filtered = engine.apply_to_dataframe(df, exact, fuzzy)
                cond_mask = pd.Series(False, index=df.index)
                cond_mask[filtered.index] = True
                if c.exclude:
                    cond_mask = ~cond_mask
                mask &= cond_mask
        return mask
