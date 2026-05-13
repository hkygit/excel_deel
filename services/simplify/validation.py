# services/simplify/validation.py
import csv
import os
import re
from typing import List, Dict, Tuple, Optional, Set
import pandas as pd

from core.csv_utils import read_csv_header, get_column_letter, split_fields

__all__ = ['read_csv_header', 'get_column_letter', 'split_fields', 'ColumnMapper', 'CsvFileValidator',
           'MixedTypeChecker', 'ValidationEngine', 'ValidationResult', 'ColumnTypeDetector']

class ColumnMapper:
    def __init__(self, headers: List[str]):
        self.headers = headers
        self._letter_to_name = {}
        self._name_to_letter = {}
        for i, name in enumerate(headers):
            letter = get_column_letter(i + 1)
            self._letter_to_name[letter] = name
            self._name_to_letter[name] = letter

    def name_to_letter(self, name: str) -> Optional[str]:
        return self._name_to_letter.get(name)

    def letter_to_name(self, letter: str) -> Optional[str]:
        return self._letter_to_name.get(letter.upper())

    def resolve_user_input(self, user_input: str, mode: str) -> List[str]:
        items = split_fields(user_input)
        if mode == 'letter':
            result = []
            for letter in items:
                name = self.letter_to_name(letter)
                if name is None:
                    raise ValueError(f"列字母 '{letter}' 不存在")
                result.append(name)
            return result
        else:
            result = []
            for name in items:
                if name not in self._name_to_letter:
                    matched = False
                    for header in self.headers:
                        if name.strip().lower() == header.strip().lower():
                            result.append(header)
                            matched = True
                            break
                    if not matched:
                        raise ValueError(f"列名 '{name}' 不存在")
                else:
                    result.append(name)
            return result

class CsvFileValidator:
    def __init__(self, filepath: str, preview_rows: int = 100):
        self.filepath = filepath
        self.preview_rows = preview_rows
        self._headers = None
        self._sample_rows = None

    def validate(self) -> Tuple[bool, Optional[str]]:
        if not os.path.exists(self.filepath):
            return False, "文件不存在"
        if os.path.getsize(self.filepath) == 0:
            return False, "文件为空"
        try:
            with open(self.filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader)
                samples = []
                for i, row in enumerate(reader):
                    if i >= self.preview_rows:
                        break
                    samples.append(row)
        except UnicodeDecodeError:
            return False, "文件编码不是 UTF-8，请转换为 UTF-8 后重试"
        except StopIteration:
            return False, "文件为空或没有表头行"
        except Exception as e:
            return False, f"读取文件失败: {str(e)}"

        cleaned = []
        empty_count = 0
        for col in header:
            col = col.strip().strip('"\ufeff').strip()
            if col == '':
                empty_count += 1
                col = f"空列{empty_count}"
            cleaned.append(col)
        self._headers = cleaned
        self._sample_rows = samples

        seen = set()
        duplicates = []
        for h in self._headers:
            if h in seen:
                duplicates.append(h)
            else:
                seen.add(h)
        if duplicates:
            return False, f"表头存在重复列名: {', '.join(duplicates)}"
        return True, None

    def get_headers(self) -> List[str]:
        if self._headers is None:
            raise RuntimeError("请先调用 validate()")
        return self._headers

    def get_sample_rows(self) -> List[List[str]]:
        if self._sample_rows is None:
            raise RuntimeError("请先调用 validate()")
        return self._sample_rows

class MixedTypeChecker:
    def __init__(self, headers: List[str], sample_rows: List[List[str]]):
        self.headers = headers
        self.sample_rows = sample_rows

    def check(self, keep_cols: List[str]) -> List[str]:
        errors = []
        for col in keep_cols:
            if col not in self.headers:
                continue
            col_idx = self.headers.index(col)
            numeric_count = 0
            string_count = 0
            total = 0
            for row in self.sample_rows:
                if len(row) <= col_idx:
                    continue
                val = row[col_idx].strip()
                if val == '':
                    continue
                total += 1
                try:
                    cleaned = val.replace(',', '').strip()
                    if re.match(r'^-?\d*\.?\d+$', cleaned):
                        numeric_count += 1
                    else:
                        string_count += 1
                except:
                    string_count += 1
            if total > 0 and string_count > 0:
                errors.append(f"列 '{col}' 包含 {string_count} 个非数值数据，不能作为聚合列。")
        return errors

class ValidationResult:
    def __init__(self, is_valid: bool, error_msg: Optional[str] = None):
        self.is_valid = is_valid
        self.error_msg = error_msg

    @classmethod
    def success(cls):
        return cls(True)

    @classmethod
    def failure(cls, msg: str):
        return cls(False, msg)

class ValidationEngine:
    def __init__(self, header_path: str, data_paths: List[str], preview_rows: int = 100):
        self.header_path = header_path
        self.data_paths = data_paths
        self.preview_rows = preview_rows
        self._headers = None
        self._sample_rows = None

    def validate(self, config: 'ProcessConfig') -> ValidationResult:
        valid, msg = self._validate_header_file()
        if not valid:
            return ValidationResult.failure(f"表头文件无效: {msg}")

        valid, msg = self._validate_data_files()
        if not valid:
            return ValidationResult.failure(f"数据文件无效: {msg}")

        valid, msg = self._validate_columns_exist(config)
        if not valid:
            return ValidationResult.failure(msg)

        if config.keep_cols and self._sample_rows:
            checker = MixedTypeChecker(self._headers, self._sample_rows)
            errors = checker.check(config.keep_cols)
            if errors:
                return ValidationResult.failure("以下列包含非数值数据，不能作为聚合列：\n" + "\n".join(errors))

        return ValidationResult.success()

    def _validate_header_file(self) -> Tuple[bool, str]:
        validator = CsvFileValidator(self.header_path, self.preview_rows)
        valid, msg = validator.validate()
        if valid:
            self._headers = validator.get_headers()
            self._sample_rows = validator.get_sample_rows()
        return valid, msg

    def _validate_data_files(self) -> Tuple[bool, str]:
        expected_col_count = len(self._headers)
        for data_path in self.data_paths:
            try:
                with open(data_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline()
                    if not first_line:
                        return False, f"数据文件 {os.path.basename(data_path)} 为空"
                    first_row = next(csv.reader([first_line]))
                    if len(first_row) != expected_col_count:
                        return False, f"数据文件 {os.path.basename(data_path)} 列数不一致"
                    f.seek(0)
                    reader = csv.reader(f)
                    for i, row in enumerate(reader):
                        if i >= self.preview_rows:
                            break
                        if len(row) != expected_col_count:
                            return False, f"数据文件 {os.path.basename(data_path)} 第 {i + 1} 行列数不正确"
            except UnicodeDecodeError:
                return False, f"数据文件 {os.path.basename(data_path)} 编码不是 UTF-8"
            except Exception as e:
                return False, f"读取数据文件 {os.path.basename(data_path)} 失败: {str(e)}"
        return True, ""

    def _validate_columns_exist(self, config) -> Tuple[bool, str]:
        all_cols = set(config.key_cols) | set(config.keep_cols) | set(config.non_sum_fields)
        for f in config.filters:
            all_cols.add(f.get('field', ''))
        missing = [c for c in all_cols if c and c not in self._headers]
        if missing:
            return False, f"列不存在: {', '.join(missing)}"
        return True, ""

    @property
    def headers(self) -> List[str]:
        return self._headers

    @property
    def sample_rows(self) -> List[List[str]]:
        return self._sample_rows

class ColumnTypeDetector:
    """检测列的数据类型（数值、字符串、混合、日期）"""
    def __init__(self, headers: List[str], sample_data: List[Dict[str, str]]):
        self.headers = headers
        self.sample_data = sample_data

    @classmethod
    def from_csv(cls, filepath: str, sample_rows: int = 1000) -> 'ColumnTypeDetector':
        sample = []
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for i, row in enumerate(reader):
                if i >= sample_rows:
                    break
                sample.append(row)
        if not sample:
            raise ValueError("无法读取样本数据")
        headers = list(sample[0].keys())
        return cls(headers, sample)

    def detect_numeric_and_string(self, exclude_cols: Set[str] = None) -> Tuple[List[str], List[str], List[str]]:
        exclude = exclude_cols or set()
        numeric = []
        string = []
        mixed = []
        for col in self.headers:
            if col in exclude:
                continue
            numeric_count = 0
            string_count = 0
            total = 0
            for row in self.sample_data:
                val = row.get(col, '')
                if val == '':
                    continue
                total += 1
                try:
                    cleaned = str(val).replace(',', '').strip()
                    if re.match(r'^-?\d*\.?\d+$', cleaned):
                        numeric_count += 1
                    else:
                        string_count += 1
                except:
                    string_count += 1
            if total == 0:
                continue
            numeric_ratio = numeric_count / total
            string_ratio = string_count / total
            if numeric_ratio > 0.95:
                numeric.append(col)
            elif string_ratio > 0.95:
                string.append(col)
            elif numeric_ratio > 0.05 and string_ratio > 0.05:
                mixed.append(col)
            else:
                if numeric_count > 0:
                    mixed.append(col)
                else:
                    string.append(col)
        return numeric, string, mixed

    def detect_date_columns(self, threshold: float = 0.8) -> Set[str]:
        date_cols = set()
        for col in self.headers:
            values = [row.get(col, '') for row in self.sample_data if row.get(col, '')]
            if self._is_date_column(values, threshold):
                date_cols.add(col)
        return date_cols

    @staticmethod
    def _is_date_column(values: List[str], threshold: float) -> bool:
        if not values:
            return False
        success = 0
        for v in values:
            try:
                pd.to_datetime(v, errors='raise')
                success += 1
            except:
                pass
        return (success / len(values)) >= threshold
