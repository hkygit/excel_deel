# core/filter_engine.py
"""统一筛选引擎，基于 simplify/RowFilter 成熟实现，支持 DataFrame 和字典行"""
from typing import Dict, List, Union, Optional
from datetime import datetime
import pandas as pd
from core.exceptions import ValidationError

VALID_OPS = {'=', '!=', '>', '<', '>=', '<=', '包含', '不包含'}

def _split_by_delimiters(text: str) -> List[str]:
    if not text:
        return []
    parts = []
    for seg in text.split(','):
        parts.extend(seg.split())
    return [p.strip() for p in parts if p.strip()]

def parse_conditions_from_form(form: Dict) -> (List[Dict], List[Dict]):
    """从表单字典解析精确条件和模糊条件"""
    exact = []
    fuzzy = []
    i = 0
    while True:
        field = form.get(f'condition_field_{i}')
        if field is None:
            break
        op = form.get(f'condition_op_{i}')
        val = form.get(f'condition_value_{i}')
        if val is None:
            val = form.get(f'condition_values_{i}')   # 旧格式
        if not field or not val:
            i += 1
            continue
        field = field.strip()
        val = val.strip()
        if not field or not val:
            i += 1
            continue
        if op is None:
            op = '='
        op = op.strip()
        if op not in VALID_OPS:
            raise ValidationError(f"不支持的操作符: {op}")
        if op in ('包含', '不包含'):
            fuzzy.append({'field': field, 'op': op, 'value': val})
        else:
            vals = _split_by_delimiters(val)
            if vals:
                exact.append({'field': field, 'op': op, 'values': vals})
        i += 1
    return exact, fuzzy

class ConditionEngine:
    """统一条件引擎，支持 DataFrame 和行字典"""

    def __init__(self, date_columns: Optional[set] = None):
        self.date_columns = date_columns or set()

    def _try_number(self, s: str) -> Optional[float]:
        try:
            return float(s)
        except (ValueError, TypeError):
            return None

    def _try_datetime(self, s: str) -> Optional[datetime]:
        for fmt in ('%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M:%S',
                    '%Y/%m/%d %H:%M', '%Y-%m-%d %H:%M',
                    '%Y/%m/%d', '%Y-%m-%d', '%Y%m%d'):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return None

    def _compare(self, cell: str, op: str, target: str) -> bool:
        # 多值 OR 仅对 = 操作符
        if op == '=' and ',' in target:
            return any(cell == v.strip() for v in target.split(','))

        # 数值比较
        cell_num = self._try_number(cell)
        target_num = self._try_number(target)
        if cell_num is not None and target_num is not None:
            if op == '=':   return cell_num == target_num
            if op == '!=':  return cell_num != target_num
            if op == '>':   return cell_num > target_num
            if op == '<':   return cell_num < target_num
            if op == '>=':  return cell_num >= target_num
            if op == '<=':  return cell_num <= target_num

        # 日期比较
        cell_dt = self._try_datetime(cell)
        target_dt = self._try_datetime(target)
        if cell_dt is not None and target_dt is not None:
            if op == '=':   return cell_dt == target_dt
            if op == '!=':  return cell_dt != target_dt
            if op == '>':   return cell_dt > target_dt
            if op == '<':   return cell_dt < target_dt
            if op == '>=':  return cell_dt >= target_dt
            if op == '<=':  return cell_dt <= target_dt

        # 字符串比较
        if op == '=':   return cell == target
        if op == '!=':  return cell != target
        if op == 'contains' or op == '包含': return target in cell
        if op == 'not_contains' or op == '不包含': return target not in cell
        if op in ('>', '<', '>=', '<='):
            try:
                return eval(f'"{cell}" {op} "{target}"')
            except:
                return False
        return False

    def apply_to_row(self, row: Dict[str, str], exact: List[Dict], fuzzy: List[Dict]) -> bool:
        for c in exact:
            field = c['field']
            op = c['op']
            for v in c['values']:
                cell = str(row.get(field, ''))
                if not self._compare(cell, op, v):
                    return False
        for c in fuzzy:
            field = c['field']
            op = c['op']
            value = c['value']
            cell = str(row.get(field, ''))
            if not self._compare(cell, op, value):
                return False
        return True

    def apply_to_dataframe(self, df: pd.DataFrame, exact: List[Dict], fuzzy: List[Dict]) -> pd.DataFrame:
        mask = pd.Series(True, index=df.index)
        for c in exact:
            f = c['field']
            if f not in df.columns:
                raise ValidationError(f"字段 {f} 不存在")
            op = c['op']
            sub = pd.Series(False, index=df.index)
            for v in c['values']:
                # 尝试数值/日期/字符串
                num_series = pd.to_numeric(df[f], errors='coerce')
                num_val = self._try_number(v)
                if num_val is not None and num_series.notna().any():
                    if op == '=':   sub |= (num_series == num_val)
                    elif op == '!=': sub |= (num_series != num_val)
                    elif op == '>': sub |= (num_series > num_val)
                    elif op == '<': sub |= (num_series < num_val)
                    elif op == '>=': sub |= (num_series >= num_val)
                    elif op == '<=': sub |= (num_series <= num_val)
                    continue
                # 日期比较
                dt_series = pd.to_datetime(df[f], errors='coerce')
                dt_val = self._try_datetime(v)
                if dt_val is not None and dt_series.notna().any():
                    if op == '=':   sub |= (dt_series == dt_val)
                    elif op == '!=': sub |= (dt_series != dt_val)
                    elif op == '>': sub |= (dt_series > dt_val)
                    elif op == '<': sub |= (dt_series < dt_val)
                    elif op == '>=': sub |= (dt_series >= dt_val)
                    elif op == '<=': sub |= (dt_series <= dt_val)
                    continue
                # 字符串
                str_series = df[f].astype(str)
                if op == '=':   sub |= (str_series == v)
                elif op == '!=': sub |= (str_series != v)
                else:           # 其他比较，尝试在字符串上进行
                    try:
                        if op == '>': sub |= (str_series > v)
                        elif op == '<': sub |= (str_series < v)
                        elif op == '>=': sub |= (str_series >= v)
                        elif op == '<=': sub |= (str_series <= v)
                    except:
                        pass
            mask &= sub
        for c in fuzzy:
            f = c['field']
            if f not in df.columns:
                raise ValidationError(f"字段 {f} 不存在")
            op = c['op']
            value = c['value']
            if op in ('包含', 'contains'):
                mask &= df[f].astype(str).str.contains(value, na=False)
            elif op in ('不包含', 'not_contains'):
                mask &= ~df[f].astype(str).str.contains(value, na=False)
        return df[mask]

# 全局默认引擎实例
default_engine = ConditionEngine()
