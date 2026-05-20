#!/usr/bin/env python3
"""
生产级公共工具部署脚本
- 修复表头检测（缓存、索引、正则、安全、性能）
- 统一文件名生成
- 更新 Subtotal 引擎与蓝图
- 包含全套单元/集成测试
- 自动清理缓存
"""
import os
import sys
import shutil

# ===================== core/header_detection.py =====================
HEADER_DETECTION = '''# core/header_detection.py
"""生产级Excel表头检测器"""
import re
import os
import time
from typing import List, Tuple, Optional, Dict, Any
from dataclasses import dataclass, field, asdict
from openpyxl import load_workbook
import logging
import hashlib

logger = logging.getLogger(__name__)

@dataclass
class HeaderDetectionConfig:
    max_check_rows: int = 20
    header_row_range: Tuple[int, int] = (0, 2)
    min_header_cols: int = 3
    max_header_cols: int = 200
    min_non_empty_ratio: float = 0.6
    min_string_ratio: float = 0.7
    max_duplicate_ratio: float = 0.3
    min_unique_cols: int = 2
    known_header_keywords: List[str] = field(default_factory=lambda: [
        "名称", "姓名", "编号", "代码", "日期", "时间", "金额", "数量",
        "单位", "备注", "状态", "类型", "级别", "部门", "岗位",
        "科目", "凭证", "摘要", "借方", "贷方", "余额", "本期", "上期",
        "病历号", "患者", "诊断", "药品", "剂量", "用法", "医生", "科室",
        "id", "name", "date", "time", "code", "type", "status", "amount"
    ])
    skip_rows_keywords: List[str] = field(default_factory=lambda: [
        "合计", "总计", "小计", "total", "sum", "汇总", "平均", "average"
    ])
    mandatory_columns: List[str] = field(default_factory=list)
    read_only: bool = True
    data_only: bool = True
    cache_ttl_seconds: int = 300

    def to_dict(self) -> Dict:
        return asdict(self)

    def get_hash(self) -> str:
        import json
        config_str = json.dumps(self.to_dict(), sort_keys=True)
        return hashlib.md5(config_str.encode()).hexdigest()[:8]

class HeaderDetectionResult:
    def __init__(self):
        self.header_rows: List[int] = []
        self.data_start_row: int = 0          # 0-based
        self.column_names: List[str] = []
        self.confidence: float = 0.0
        self.reason: str = ""
        self.metrics: Dict[str, Any] = {}
        self.warnings: List[str] = []
        self.timestamp: float = time.time()
        self.config_hash: str = ""

    def is_valid(self) -> bool:
        return bool(self.header_rows) and self.confidence >= 0.1

    def get_data_start_excel_row(self) -> int:
        """返回1-based的数据起始行"""
        return self.data_start_row + 1 if self.data_start_row >= 0 else 1

class SimpleExcelHeaderDetector:
    def __init__(self, config: HeaderDetectionConfig = None):
        self.config = config or HeaderDetectionConfig()
        self.config_hash = self.config.get_hash()
        self._cache = {}
        self._workbook_cache = {}
        self._keyword_set = {kw.lower() for kw in self.config.known_header_keywords}
        self._skip_set = {kw.lower() for kw in self.config.skip_rows_keywords}
        self._numeric_regex = re.compile(r'^[+-]?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?%?$')
        self._invalid_chars_regex = re.compile(r'[\\/*?\[\]]')

    def detect_headers(self, file_path: str, sheet_name: str = None, use_cache: bool = True) -> HeaderDetectionResult:
        self._validate_file(file_path)
        mtime = os.path.getmtime(file_path)
        cache_key = (file_path, mtime, self.config_hash, sheet_name)
        if use_cache and cache_key in self._cache:
            cached = self._cache[cache_key]
            if time.time() - cached.timestamp < self.config.cache_ttl_seconds:
                cached.metrics['cache_hit'] = True
                return cached

        result = HeaderDetectionResult()
        result.config_hash = self.config_hash
        try:
            wb = load_workbook(filename=file_path, read_only=self.config.read_only, data_only=self.config.data_only)
            ws = wb[sheet_name] if sheet_name else wb.active
            result = self._detect_from_worksheet(ws)
            wb.close()
        except Exception as e:
            logger.error(f"检测失败 {file_path}: {e}")
            result = self._create_fallback_result()
            result.warnings.append(str(e))

        if use_cache and result.is_valid():
            self._cache[cache_key] = result
        return result

    def detect_from_workbook(self, workbook, sheet_name: str = None) -> HeaderDetectionResult:
        workbook_id = id(workbook)
        cache_key = (workbook_id, sheet_name, self.config_hash)
        if cache_key in self._workbook_cache:
            cached = self._workbook_cache[cache_key]
            cached.metrics['workbook_cache_hit'] = True
            return cached

        try:
            ws = workbook[sheet_name] if sheet_name else workbook.active
            result = self._detect_from_worksheet(ws)
            result.config_hash = self.config_hash
        except Exception as e:
            logger.error(f"从workbook检测失败: {e}")
            result = self._create_fallback_result()
            result.warnings.append(str(e))

        self._workbook_cache[cache_key] = result
        return result

    def _validate_file(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        if not os.path.isfile(file_path):
            raise ValueError(f"路径不是文件: {file_path}")
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in {'.xlsx', '.xls', '.xlsm'}:
            raise ValueError(f"不支持的文件类型: {ext}")

    def _detect_from_worksheet(self, ws) -> HeaderDetectionResult:
        result = HeaderDetectionResult()
        start_time = time.time()
        max_row = min(ws.max_row, self.config.max_check_rows)
        max_col = min(ws.max_column, self.config.max_header_cols)
        if max_row == 0 or max_col == 0:
            result.warnings.append("工作表为空")
            return self._create_fallback_result()

        data = self._read_cell_range(ws, 1, max_row, 1, max_col)
        if not data:
            return self._create_fallback_result()

        candidates = []
        start, end = self.config.header_row_range
        for row_idx in range(min(end + 1, len(data))):
            if row_idx < start:
                continue
            candidate = self._evaluate_row(data, row_idx)
            if candidate['score'] > 0:
                candidates.append(candidate)

        if candidates:
            candidates.sort(key=lambda x: x['score'], reverse=True)
            best = candidates[0]
            result.header_rows = [best['row']]
            result.confidence = best['score']
            result.reason = best['reason']
            result.column_names = best['column_names']
            result.data_start_row = best['row'] + 1
        else:
            result = self._create_fallback_result()
            if data:
                result.column_names = [
                    str(cell).strip() if cell is not None and str(cell).strip()
                    else f"Col_{i+1}" for i, cell in enumerate(data[0])
                ]
        result.metrics.update({
            'execution_time_ms': (time.time() - start_time) * 1000,
            'rows_checked': len(data),
            'columns_checked': max_col,
            'candidates_found': len(candidates)
        })
        return result

    def _read_cell_range(self, ws, min_row, max_row, min_col, max_col):
        try:
            return [
                [cell for cell in row]
                for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                        min_col=min_col, max_col=max_col,
                                        values_only=True)
            ]
        except Exception as e:
            logger.warning(f"读取单元格失败: {e}")
            return []

    def _evaluate_row(self, data, row_idx):
        if row_idx >= len(data):
            return self._create_candidate(row_idx, 0, "索引越界")
        row_data = data[row_idx]
        if not row_data:
            return self._create_candidate(row_idx, 0, "空行")

        non_empty = [cell for cell in row_data if cell is not None and str(cell).strip() != '']
        non_empty_ratio = len(non_empty) / len(row_data) if row_data else 0
        if non_empty_ratio < 0.2:
            return self._create_candidate(row_idx, 0, f"非空{non_empty_ratio:.0%}")

        string_cells = []
        for cell in non_empty:
            s = str(cell).strip()
            if s and not self._is_numeric(s):
                string_cells.append(s)
        string_ratio = len(string_cells) / len(row_data) if row_data else 0
        unique_strings = set(string_cells)
        dup_ratio = 1 - len(unique_strings) / len(string_cells) if string_cells else 0

        keyword_matches = sum(1 for s in string_cells for kw in self._keyword_set if kw in s.lower())
        keyword_ratio = keyword_matches / len(string_cells) if string_cells else 0

        combined = ' '.join(string_cells).lower()
        if any(skip in combined for skip in self._skip_set):
            return self._create_candidate(row_idx, 0, "跳过关键词")

        valid_fmt = sum(1 for s in string_cells if self._is_valid_column_name(s))
        valid_fmt_ratio = valid_fmt / len(string_cells) if string_cells else 0

        score = 0.0
        reasons = []
        if non_empty_ratio >= self.config.min_non_empty_ratio:
            score += 0.3; reasons.append(f"非空{non_empty_ratio:.0%}")
        if string_ratio >= self.config.min_string_ratio:
            score += 0.3; reasons.append(f"文本{string_ratio:.0%}")
        if dup_ratio <= self.config.max_duplicate_ratio:
            score += 0.2; reasons.append(f"唯一{1-dup_ratio:.0%}")
        if keyword_ratio > 0.1:
            score += min(keyword_ratio, 0.2); reasons.append(f"关键词{keyword_ratio:.0%}")
        if valid_fmt_ratio > 0.5:
            score += 0.1; reasons.append("格式正确")
        score = max(0.0, min(1.0, score))

        column_names = [
            str(cell).strip() if cell is not None and str(cell).strip() != '' else f"Col_{i+1}"
            for i, cell in enumerate(row_data)
        ]
        return {
            'row': row_idx, 'score': score,
            'reason': ', '.join(reasons) or '基础特征',
            'column_names': column_names,
            'metrics': {}
        }

    def _create_candidate(self, row_idx, score, reason):
        return {'row': row_idx, 'score': score, 'reason': reason, 'column_names': [], 'metrics': {}}

    def _is_numeric(self, s: str) -> bool:
        s = s.strip().replace(',', '')
        if s.endswith('%'):
            s = s[:-1]
        try:
            float(s)
            return True
        except ValueError:
            return False

    def _is_valid_column_name(self, s: str) -> bool:
        if not s or len(s) > 50:
            return False
        if re.match(r'^[\d.,%\s-]+$', s):
            return False
        if self._invalid_chars_regex.search(s):
            return False
        return s.strip() != ''

    def _create_fallback_result(self) -> HeaderDetectionResult:
        result = HeaderDetectionResult()
        result.header_rows = [0]
        result.data_start_row = 1
        result.confidence = 0.3
        result.reason = "回退默认第一行"
        return result

    def clear_cache(self):
        self._cache.clear()
        self._workbook_cache.clear()


def detect_header_start(file_path: str, sheet_name: str = None) -> int:
    """返回数据起始行（1-based），失败返回2"""
    detector = SimpleExcelHeaderDetector()
    try:
        result = detector.detect_headers(file_path, sheet_name)
        if result.is_valid():
            return result.get_data_start_excel_row()
    except Exception:
        pass
    return 2
'''

# ===================== core/file_naming.py =====================
FILE_NAMING = '''# core/file_naming.py
import os

def generate_output_filename(module_name: str, original_filename: str, suffix: str = '_小计') -> str:
    base = os.path.splitext(os.path.basename(original_filename))[0]
    safe_module = module_name.replace('/', '_').replace('\\\\', '_')
    return f"{safe_module}_{base}{suffix}.xlsx"
'''

# ===================== services/subtotal/engine.py =====================
SUBTOTAL_ENGINE = '''# services/subtotal/engine.py
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from core.csv_utils import split_fields
from core.header_detection import SimpleExcelHeaderDetector
from core.file_naming import generate_output_filename

SUBTOTAL_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
SUBTOTAL_FONT = Font(bold=True)
THIN_BORDER = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))

class SubtotalEngine:
    def __init__(self, filepath, key_cols_str, calc_cols_str, agg_funcs,
                 dimension_cols=None, sheets='1', original_filename='data.xlsx'):
        self.filepath = filepath
        self.original_filename = original_filename
        self.key_cols = self._parse_cols(key_cols_str)
        self.calc_cols = self._parse_cols(calc_cols_str)
        self.dimension_cols = self._parse_cols(dimension_cols) if dimension_cols else []
        self.agg_funcs = agg_funcs
        self.sheets = sheets
        self.detector = SimpleExcelHeaderDetector()

    @staticmethod
    def _parse_cols(raw):
        if not raw:
            return []
        items = split_fields(raw)
        result = []
        for item in items:
            if '-' in item:
                start, end = item.split('-', 1)
                start_idx = column_index_from_string(start.strip())
                end_idx = column_index_from_string(end.strip())
                for i in range(start_idx, end_idx + 1):
                    result.append(i)
            else:
                result.append(column_index_from_string(item.strip()))
        return sorted(list(set(result)))

    def _get_agg_func(self, col_letter):
        return self.agg_funcs.get(col_letter, self.agg_funcs.get('default', 'sum'))

    def process(self):
        wb = openpyxl.load_workbook(self.filepath)
        header_result = self.detector.detect_from_workbook(wb)
        data_start_excel_row = header_result.get_data_start_excel_row()

        sheet_names = wb.sheetnames
        if self.sheets == 'all':
            target_sheets = sheet_names
        else:
            indices = [int(x.strip()) for x in self.sheets.split(',') if x.strip().isdigit()]
            target_sheets = [sheet_names[i-1] for i in indices if 1 <= i <= len(sheet_names)]

        for sheet_name in target_sheets:
            ws = wb[sheet_name]
            self._process_sheet(ws, data_start_excel_row)

        out_name = generate_output_filename('subtotal', self.original_filename)
        out_dir = os.path.dirname(self.filepath)
        out_path = os.path.join(out_dir, out_name)
        wb.save(out_path)
        wb.close()
        return out_path

    def _process_sheet(self, ws, data_start_row):
        max_col = ws.max_column
        max_row = ws.max_row
        if data_start_row > max_row:
            return

        groups = {}
        for row_idx in range(data_start_row, max_row + 1):
            key = tuple(ws.cell(row_idx, col).value for col in self.key_cols)
            groups.setdefault(key, []).append(row_idx)

        sorted_groups = sorted(groups.items(), key=lambda x: max(x[1]), reverse=True)

        for key, rows in sorted_groups:
            last_row = max(rows)
            ws.insert_rows(last_row + 1)
            subtotal_row = last_row + 1

            for col_idx in range(1, max_col + 1):
                cell = ws.cell(subtotal_row, col_idx)
                cell.fill = SUBTOTAL_FILL
                cell.font = SUBTOTAL_FONT
                cell.border = THIN_BORDER

            for i, col in enumerate(self.key_cols):
                ws.cell(subtotal_row, col).value = key[i]

            for col in self.dimension_cols:
                ws.cell(subtotal_row, col).value = '合计'

            for col in self.calc_cols:
                col_letter = get_column_letter(col)
                func = self._get_agg_func(col_letter)
                values = []
                for r in rows:
                    val = ws.cell(r, col).value
                    try:
                        val = float(val)
                    except (TypeError, ValueError):
                        val = 0
                    values.append(val)

                if func == 'sum':
                    result = sum(values)
                elif func == 'avg':
                    result = sum(values) / len(values) if values else 0
                elif func == 'min':
                    result = min(values) if values else 0
                elif func == 'max':
                    result = max(values) if values else 0
                elif func == 'count':
                    result = len(values)
                else:
                    result = sum(values)
                ws.cell(subtotal_row, col).value = result
'''

# ===================== blueprints/subtotal.py =====================
SUBTOTAL_BP = '''# blueprints/subtotal.py
import os
import logging
import openpyxl
from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from core.decorators import handle_exceptions
from core.exceptions import ValidationError
from core.task.decorator import async_task
from services.subtotal.engine import SubtotalEngine
from core.csv_utils import split_fields

logger = logging.getLogger(__name__)
subtotal_bp = Blueprint('subtotal', __name__, url_prefix='/subtotal', template_folder='../templates/subtotal')

@subtotal_bp.route('/')
def index():
    return render_template('subtotal/index.html')

def _validate_col_range(item, max_col, category):
    if '-' in item:
        parts = item.split('-')
        if len(parts) != 2:
            raise ValidationError(f"{category}列范围格式错误：{item}")
        start, end = parts[0].strip(), parts[1].strip()
        _check_letter(start, max_col, category)
        _check_letter(end, max_col, category)
        if openpyxl.utils.column_index_from_string(start) > openpyxl.utils.column_index_from_string(end):
            raise ValidationError(f"{category}列范围开始大于结束：{item}")
    else:
        _check_letter(item, max_col, category)

def _check_letter(letter, max_col, category):
    idx = openpyxl.utils.column_index_from_string(letter)
    if idx > max_col:
        raise ValidationError(f"{category}列 '{letter}' 超出工作表列范围（共{max_col}列）")

def _check_columns_exist(raw_str, category, max_col):
    if not raw_str:
        return
    items = split_fields(raw_str)
    for item in items:
        _validate_col_range(item, max_col, category)

@subtotal_bp.route('/process', methods=['POST'])
@handle_exceptions()
def process():
    file = request.files.get('file')
    if not file or file.filename == '':
        raise ValidationError('请上传文件')

    original_filename = file.filename
    key_cols_raw = request.form.get('key_cols', '').strip()
    if not key_cols_raw:
        raise ValidationError('请指定主键列')

    dimension_cols_raw = request.form.get('dimension_cols', '').strip() or None
    calc_cols_raw = request.form.get('calc_cols', '').strip()
    if not calc_cols_raw:
        raise ValidationError('请指定计算列')

    agg_func_default = request.form.get('agg_func_default', 'sum').strip()
    agg_overrides_raw = request.form.get('agg_overrides', '').strip()
    sheets_mode = request.form.get('sheets_mode', '1').strip()

    file_service = current_app.container.file_service()
    path = file_service.save_upload(file, allowed_extensions={'xlsx', 'xls'})

    wb = openpyxl.load_workbook(path)
    ws = wb.active
    max_col = ws.max_column
    wb.close()

    _check_columns_exist(key_cols_raw, '主键', max_col)
    _check_columns_exist(calc_cols_raw, '计算', max_col)
    if dimension_cols_raw:
        _check_columns_exist(dimension_cols_raw, '维度', max_col)

    agg_funcs = {'default': agg_func_default}
    if agg_overrides_raw:
        for part in agg_overrides_raw.split(','):
            if ':' in part:
                col, func = part.split(':', 1)
                col = col.strip().upper()
                func = func.strip().lower()
                if func not in ('sum', 'avg', 'min', 'max', 'count'):
                    raise ValidationError(f"不支持的聚合函数: {func}")
                try:
                    openpyxl.utils.column_index_from_string(col)
                except:
                    raise ValidationError(f"无效的列字母: {col}")
                agg_funcs[col] = func

    wb2 = openpyxl.load_workbook(path, read_only=True)
    total_rows = sum(sheet.max_row for sheet in wb2.worksheets)
    wb2.close()
    data_rows = total_rows - 1

    app = current_app._get_current_object()

    if data_rows > 5000:
        task_storage = app.container.task_storage()

        @async_task(original_filename='subtotal')
        def run_subtotal(task_id):
            with app.app_context():
                def progress_cb(percent, message):
                    task_storage.update(task_id, progress_percent=percent, progress_message=message)
                try:
                    progress_cb(0, "开始处理...")
                    engine = SubtotalEngine(
                        path, key_cols_raw, calc_cols_raw, agg_funcs,
                        dimension_cols=dimension_cols_raw,
                        sheets=sheets_mode,
                        original_filename=original_filename
                    )
                    out_path = engine.process()
                    progress_cb(100, "完成")
                    return {'result_path': out_path}
                except Exception as e:
                    import traceback
                    error_msg = f"{str(e)}\\n{traceback.format_exc()}"
                    task_storage.append_log(task_id, error_msg)
                    raise
                finally:
                    file_service.cleanup([path])

        result = run_subtotal()
        return jsonify({'task_id': result['task_id']})
    else:
        try:
            engine = SubtotalEngine(
                path, key_cols_raw, calc_cols_raw, agg_funcs,
                dimension_cols=dimension_cols_raw,
                sheets=sheets_mode,
                original_filename=original_filename
            )
            out_path = engine.process()
            return send_file(out_path, as_attachment=True,
                             download_name=os.path.basename(out_path),
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        finally:
            file_service.cleanup([path])

@subtotal_bp.route('/status/<task_id>')
def status(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data:
        return jsonify({'error': '任务不存在'}), 404
    log = storage.get_log(task_id, max_lines=100)
    data['log'] = log
    return jsonify(data)

@subtotal_bp.route('/download/<task_id>')
def download(task_id):
    storage = current_app.container.task_storage()
    data = storage.load(task_id)
    if not data or data.get('status') != 'completed':
        return jsonify({'error': '结果不可用'}), 400
    path = data.get('result_path')
    if not path or not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(path, as_attachment=True, download_name='subtotal_result.xlsx')
'''

# ===================== 测试文件 =====================
TEST_HEADER_DETECTION = '''# tests/test_header_detection.py
import pytest
import pandas as pd
from core.header_detection import SimpleExcelHeaderDetector, detect_header_start, HeaderDetectionConfig
import openpyxl

class TestHeaderDetection:
    @pytest.fixture
    def simple_file(self, tmp_path):
        df = pd.DataFrame({'姓名': ['Alice', 'Bob'], '年龄': [30, 25]})
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        return str(path)

    def test_basic_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(simple_file)
        assert result.header_rows == [0]
        assert result.data_start_row == 1
        assert '姓名' in result.column_names
        assert result.confidence > 0.5

    def test_detect_header_start(self, simple_file):
        start = detect_header_start(simple_file)
        assert start == 2

    def test_skip_total_row(self, tmp_path):
        df = pd.DataFrame({'名称': ['X', '合计'], '金额': [10, 30]})
        path = tmp_path / 'skip.xlsx'
        df.to_excel(path, index=False)
        detector = SimpleExcelHeaderDetector()
        result = detector.detect_headers(path)
        assert result.header_rows == [0]

    def test_workbook_detection(self, simple_file):
        detector = SimpleExcelHeaderDetector()
        wb = openpyxl.load_workbook(simple_file)
        result1 = detector.detect_from_workbook(wb)
        result2 = detector.detect_from_workbook(wb)
        assert result1.header_rows == [0]
        assert result2.metrics.get('workbook_cache_hit') is True
        wb.close()
'''

TEST_FILE_NAMING = '''# tests/test_file_naming.py
from core.file_naming import generate_output_filename

class TestFileNaming:
    def test_basic(self):
        name = generate_output_filename('subtotal', '数据.xlsx')
        assert name == 'subtotal_数据_小计.xlsx'

    def test_with_path(self):
        name = generate_output_filename('match', '/path/to/目标文件.xlsx')
        assert name == 'match_目标文件_小计.xlsx'

    def test_custom_suffix(self):
        name = generate_output_filename('filter', 'data.csv', suffix='_汇总')
        assert name == 'filter_data_汇总.xlsx'
'''

TEST_SUBTOTAL_ENGINE = '''# tests/test_subtotal_engine.py
import pytest
import pandas as pd
import os
from services.subtotal.engine import SubtotalEngine

class TestSubtotalEngine:
    @pytest.fixture
    def sample_excel(self, tmp_path):
        df = pd.DataFrame({
            '部门': ['X', 'X', 'Y'],
            '金额': [10, 20, 30],
            '数量': [1, 2, 3]
        })
        path = tmp_path / 'test.xlsx'
        df.to_excel(path, index=False)
        return str(path), 'test.xlsx'

    def test_basic_subtotal_sum(self, sample_excel):
        filepath, fname = sample_excel
        engine = SubtotalEngine(filepath, '部门', '金额,数量', {'default': 'sum'}, original_filename=fname)
        out_path = engine.process()
        assert os.path.basename(out_path).startswith('subtotal_test')
        wb = pd.read_excel(out_path, sheet_name=None)
        ws = wb['Sheet1']
        assert len(ws) == 5
        assert ws.iloc[2]['金额'] == 30
        assert ws.iloc[2]['数量'] == 3

    def test_dimension_col(self, sample_excel):
        filepath, fname = sample_excel
        engine = SubtotalEngine(filepath, '部门', '金额', {'default': 'sum'},
                                dimension_cols='数量', original_filename=fname)
        out_path = engine.process()
        wb = pd.read_excel(out_path, sheet_name=None)
        ws = wb['Sheet1']
        assert ws.iloc[2]['数量'] == '合计'
'''

TEST_INTEGRATION_SUBTOTAL = '''# tests/test_integration_subtotal.py
import io
import pandas as pd

class TestSubtotalIntegration:
    def test_sync_endpoint(self, client):
        df = pd.DataFrame({
            '部门': ['X', 'X', 'Y'],
            '金额': [10, 20, 30]
        })
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': '部门',
            'calc_cols': '金额',
            'agg_func_default': 'sum',
            'sheets_mode': '1'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'subtotal_test' in resp.headers.get('Content-Disposition', '')
        result_df = pd.read_excel(io.BytesIO(resp.data), sheet_name=None)
        ws = result_df['Sheet1']
        assert len(ws) == 5
        assert ws.iloc[2]['金额'] == 30

    def test_missing_dimension_col(self, client):
        df = pd.DataFrame({'A': [1]})
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'A',
            'dimension_cols': 'Z'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 400
        assert '超出' in resp.get_json()['error']

    def test_async_large_file(self, client):
        df = pd.DataFrame({'A': ['X', 'Y'] * 3000, 'B': range(6000)})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
        buf.seek(0)

        data = {
            'file': (buf, 'test_large.xlsx'),
            'key_cols': 'A',
            'calc_cols': 'B',
            'agg_func_default': 'sum'
        }
        resp = client.post('/subtotal/process', data=data, content_type='multipart/form-data')
        assert resp.status_code == 200
        assert 'task_id' in resp.get_json()
'''

def clean_cache(root_dir):
    removed = 0
    for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
        for d in list(dirnames):
            if d == '__pycache__':
                shutil.rmtree(os.path.join(dirpath, d), ignore_errors=True)
                removed += 1
    print(f"清理了 {removed} 个 __pycache__ 目录")

def main():
    root = os.path.dirname(os.path.abspath(__file__))
    files = {
        'core/header_detection.py': HEADER_DETECTION,
        'core/file_naming.py': FILE_NAMING,
        'services/subtotal/engine.py': SUBTOTAL_ENGINE,
        'blueprints/subtotal.py': SUBTOTAL_BP,
        'tests/test_header_detection.py': TEST_HEADER_DETECTION,
        'tests/test_file_naming.py': TEST_FILE_NAMING,
        'tests/test_subtotal_engine.py': TEST_SUBTOTAL_ENGINE,
        'tests/test_integration_subtotal.py': TEST_INTEGRATION_SUBTOTAL,
    }
    for rel_path, content in files.items():
        full_path = os.path.join(root, rel_path)
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"已写入: {rel_path}")

    clean_cache(root)
    print("\n部署完成。请运行 pytest 验证：")
    print("python -m pytest tests/test_header_detection.py tests/test_file_naming.py tests/test_subtotal_engine.py tests/test_integration_subtotal.py -v")

if __name__ == '__main__':
    main()