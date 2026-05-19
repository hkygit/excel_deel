import os
import pandas as pd
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment
from core.excel_utils import ExcelFile, col_letter_to_index
from core.excel_utils import copy_merge_structure
from services.data.reader import DataReader

class ExcelMatcher:
    """Excel 数据匹配填充核心类（支持依赖注入 DataReader）"""

    def __init__(self, target_path, source_path, config, sheet_indices=None, data_reader=None):
        self.target_path = target_path
        self.source_path = source_path
        self.config = config
        self.sheet_indices = sheet_indices
        self.errors = []
        self.target_wb = None
        self.target_ws = None
        self.source_df = None
        self.data_reader = data_reader or DataReader()
        self.allow_duplicate_keys = config.get('allow_duplicate_keys', False)

    def _validate_column_letters(self):
        try:
            target_key_letters = [c.strip() for c in self.config['target_key_cols'].split(',')]
            source_match_letters = [c.strip() for c in self.config['source_match_cols'].split(',')]
            source_value_letter = self.config['source_value_col'].strip()

            self.target_key_cols = [column_index_from_string(letter) for letter in target_key_letters]
            self.source_match_cols = [col_letter_to_index(letter) for letter in source_match_letters]
            self.source_value_col = col_letter_to_index(source_value_letter)

            self.target_key_letters = target_key_letters
            self.source_match_letters = source_match_letters
            self.source_value_letter = source_value_letter
        except Exception as e:
            self.errors.append(f"列字母解析错误: {str(e)}")
            return False

        try:
            self.result_col = column_index_from_string(self.config['result_col'])
            self.result_is_new = False
        except:
            self.result_is_new = True
            self.result_col_name = self.config['result_col'].strip()
            self.result_col = None
        return True

    def _load_source(self):
        try:
            if self.sheet_indices:
                all_sheets = pd.read_excel(self.source_path, sheet_name=None, dtype=str)
                sheet_names = list(all_sheets.keys())
                selected_sheets = []
                for idx in self.sheet_indices:
                    if idx < 1 or idx > len(sheet_names):
                        self.errors.append(f"Sheet索引 {idx} 超出范围（共 {len(sheet_names)} 个sheet）")
                        return False
                    selected_sheets.append(sheet_names[idx-1])
                df_list = []
                for sheet_name in selected_sheets:
                    df = all_sheets[sheet_name]
                    df_list.append(df)
                self.source_df = pd.concat(df_list, ignore_index=True, sort=False)
            else:
                self.source_df = self.data_reader.read_excel(self.source_path, dtype=str)

            if self.source_df.empty:
                self.errors.append("源文件无数据")
                return False

            self.source_col_names = list(self.source_df.columns)
            max_cols = len(self.source_col_names)

            for letter, idx in zip(self.source_match_letters, self.source_match_cols):
                if idx > max_cols:
                    self.errors.append(f"源文件匹配列 {letter} 超出列范围（共 {max_cols} 列）")
                    return False

            if self.source_value_col > max_cols:
                self.errors.append(f"源文件值列 {self.source_value_letter} 超出列范围（共 {max_cols} 列）")
                return False
            return True
        except Exception as e:
            self.errors.append(f"读取源文件失败: {str(e)}")
            return False

    def _load_target(self):
        try:
            self.target_wb = ExcelFile(self.target_path)
            self.target_wb.load()
            self.target_ws = self.target_wb.active_sheet
            max_col = self.target_ws.max_column
            for letter, idx in zip(self.target_key_letters, self.target_key_cols):
                if idx > max_col:
                    self.errors.append(f"目标文件键列 {letter} 超出最大列数 {max_col}")
                    return False
            return True
        except Exception as e:
            msg = str(e)
            if 'does not support' in msg or 'file format' in msg:
                self.errors.append("目标文件格式不支持，请使用 .xlsx 格式文件。如为 .xls 文件，请先用 Excel 另存为 .xlsx。")
            else:
                self.errors.append(f"读取目标文件失败: {msg}")
            return False

    def _extract_target_keys(self):
        start_row = 2
        end_row = self.target_ws.max_row
        keys = []
        for row in range(start_row, end_row + 1):
            key = []
            for col in self.target_key_cols:
                val = self.target_wb.get_cell_value(row, col, sheet=self.target_ws)
                key.append(str(val).strip() if val else '')
            keys.append(tuple(key))
        return keys, start_row, end_row

    def _build_lookup_dict(self):
        match_col_names = [self.source_col_names[col - 1] for col in self.source_match_cols]
        value_col_name = self.source_col_names[self.source_value_col - 1]
        match_len = min(len(self.target_key_cols), len(match_col_names))
        match_col_names = match_col_names[:match_len]

        lookup = {}
        for _, row in self.source_df.iterrows():
            key = tuple(str(row[col]).strip() if pd.notna(row[col]) else '' for col in match_col_names)
            if key not in lookup:
                lookup[key] = row[value_col_name]
        return lookup, match_len

    def _write_result_column(self, result_col_idx, keys, values, start_row):
        to_remove = []
        for merged in self.target_ws.merged_cells.ranges:
            if merged.min_col <= result_col_idx <= merged.max_col:
                to_remove.append(merged)
        for merged in to_remove:
            self.target_ws.unmerge_cells(str(merged))

        seen_keys = set()
        for i, key in enumerate(keys):
            row = start_row + i
            if self.allow_duplicate_keys:
                if key in seen_keys:
                    continue
                seen_keys.add(key)

            val = values.get(key, 0)
            try:
                val = float(val)
            except:
                pass
            self.target_ws.cell(row, result_col_idx).value = val

        if self.target_key_cols:
            copy_merge_structure(self.target_ws, self.target_key_cols[0],
                                 self.target_ws, result_col_idx,
                                 start_row, self.target_ws.max_row)

        for row in range(start_row, self.target_ws.max_row + 1):
            cell = self.target_ws.cell(row, result_col_idx)
            cell.alignment = Alignment(horizontal='right')

    def _generate_report(self, used_target_keys, result_values, start_row, match_len):
        target_keys_full, _, _ = self._extract_target_keys()
        target_keys_used = used_target_keys
        lookup, _ = self._build_lookup_dict()
        source_keys_set = set(lookup.keys())
        source_items = list(lookup.items())

        header_row = 1
        target_col_names = []
        for col_1based in self.target_key_cols:
            col_name = self.target_ws.cell(header_row, col_1based).value
            if not col_name:
                col_name = f"列{chr(64 + col_1based)}"
            target_col_names.append(col_name)
        target_key_label = "|".join(target_col_names)
        target_file_name = os.path.basename(self.target_path)

        matched_pairs = []
        only_target_pairs = []
        only_source_pairs = []

        for i, (full_key, used_key) in enumerate(zip(target_keys_full, target_keys_used)):
            row_num = start_row + i
            if used_key in source_keys_set:
                matched_val = lookup.get(used_key)
                matched_pairs.append((row_num, full_key, used_key, matched_val))
            else:
                only_target_pairs.append((row_num, full_key))

        for key, val in source_items:
            if key not in set(target_keys_used):
                only_source_pairs.append((key, val))

        report_sheet = self.target_wb._workbook.create_sheet("校验报告")
        headers = ["行号", f"目标文件键列值\n({target_file_name}【{target_key_label}】)", "源文件匹配键", "源文件值", "匹配状态", "备注"]
        for col, header in enumerate(headers, 1):
            report_sheet.cell(1, col).value = header

        row_idx = 2
        for row_num, full_key, source_key, val in matched_pairs:
            report_sheet.cell(row_idx, 1).value = row_num
            report_sheet.cell(row_idx, 2).value = "|".join(str(v) for v in full_key)
            report_sheet.cell(row_idx, 3).value = "|".join(str(v) for v in source_key)
            report_sheet.cell(row_idx, 4).value = val
            report_sheet.cell(row_idx, 5).value = 1
            row_idx += 1

        for row_num, full_key in only_target_pairs:
            report_sheet.cell(row_idx, 1).value = row_num
            report_sheet.cell(row_idx, 2).value = "|".join(str(v) for v in full_key)
            report_sheet.cell(row_idx, 5).value = 0
            row_idx += 1

        for key, val in only_source_pairs:
            report_sheet.cell(row_idx, 1).value = "-"
            report_sheet.cell(row_idx, 3).value = "|".join(str(v) for v in key)
            report_sheet.cell(row_idx, 4).value = val
            report_sheet.cell(row_idx, 5).value = 0
            row_idx += 1

        for col in range(1, 7):
            for row in range(2, row_idx):
                report_sheet.cell(row, col).alignment = Alignment(horizontal='right')

    def match(self):
        if not self._validate_column_letters():
            return False, self.errors
        if not self._load_source():
            return False, self.errors
        if not self._load_target():
            return False, self.errors

        target_keys, start_row, end_row = self._extract_target_keys()
        lookup, match_len = self._build_lookup_dict()
        used_target_keys = [key[:match_len] for key in target_keys]

        result_values = {key: lookup.get(key, 0) for key in set(used_target_keys)}

        if self.result_is_new:
            max_col = self.target_ws.max_column
            result_col = max_col + 1
            self.target_ws.insert_cols(result_col)
            self.target_ws.cell(1, result_col).value = self.result_col_name
        else:
            result_col = self.result_col

        self._write_result_column(result_col, used_target_keys, result_values, start_row)
        self._generate_report(used_target_keys, result_values, start_row, match_len)

        out_path = os.path.join(os.path.dirname(self.target_path), f"matched_{os.path.basename(self.target_path)}")
        self.target_wb.save(out_path)
        return True, out_path

    def debug_info(self):
        if not self._validate_column_letters():
            return {'error': self.errors}
        if not self._load_source():
            return {'error': self.errors}
        if not self._load_target():
            return {'error': self.errors}

        target_keys, start_row, end_row = self._extract_target_keys()
        lookup, match_len = self._build_lookup_dict()
        used_target_keys = [key[:match_len] for key in target_keys]

        target_set = set(used_target_keys)
        source_set = set(lookup.keys())
        common = target_set & source_set
        only_target = target_set - source_set
        only_source = source_set - target_set

        matched_pairs = []
        unmatched_pairs = []
        for i, (full_key, used_key) in enumerate(zip(target_keys, used_target_keys)):
            row_num = start_row + i
            matched_val = lookup.get(used_key)
            if matched_val is not None:
                matched_pairs.append((row_num, full_key, used_key, matched_val))
            else:
                unmatched_pairs.append((row_num, full_key, used_key))
            if len(matched_pairs) >= 20 and len(unmatched_pairs) >= 20:
                break

        return {
            'config': self.config,
            'target_rows': len(target_keys),
            'target_keys_sample': [(start_row + i, key) for i, key in enumerate(target_keys[:20])],
            'source_keys_sample': list(source_set)[:20],
            'source_values_sample': [(k, lookup[k]) for k in list(source_set)[:20]],
            'common_count': len(common),
            'only_target_count': len(only_target),
            'only_source_count': len(only_source),
            'matched_count': len(common),
            'unmatched_keys': list(only_target)[:20],
            'start_row': start_row,
            'end_row': end_row,
            'matched_pairs': matched_pairs,
            'unmatched_pairs': unmatched_pairs
        }
