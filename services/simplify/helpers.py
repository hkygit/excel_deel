# services/simplify/helpers.py
from core.filter_engine import ConditionEngine
from core.csv_utils import build_column_mapping
from typing import List, Dict, Set, Tuple

class RowFilter:
    """基于统一引擎的 RowFilter"""
    def __init__(self, filters: List[Dict], name_to_letter: Dict[str, str],
                 headers: List[str], date_columns: Set[str]):
        self.engine = ConditionEngine(date_columns)
        self.exact = []
        self.fuzzy = []
        for f in filters:
            field = f.get('field', '').strip()
            op = f.get('op', '').strip()
            value = f.get('value', '').strip()
            exclude = f.get('exclude', False)
            if not field or not value:
                continue
            if op in ('包含', '不包含'):
                self.fuzzy.append({'field': field, 'op': op, 'value': value})
            else:
                if ',' in value:
                    vals = [v.strip() for v in value.split(',') if v.strip()]
                else:
                    vals = [value]
                self.exact.append({'field': field, 'op': op, 'values': vals})

    def apply(self, row: Dict[str, str]) -> bool:
        return self.engine.apply_to_row(row, self.exact, self.fuzzy)
