- **services/team_report/service.py**
  - language: `python`
  - content:
    ```python
# services/team_report/service.py
"""团队报表服务：数据聚合与计算"""
import pandas as pd
import numpy as np
from typing import Dict, List, Optional
from core.exceptions import ValidationError
import logging

logger = logging.getLogger(__name__)

class TeamReportService:
    def __init__(self, config):
        self.config = config

    def _validate_columns(self, df, required_cols, label):
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValidationError(f"{label}缺少列: {missing}")

    # ---------- 映射加载 ----------
    def load_mapping(self, path: str) -> Dict[str, Dict[str, str]]:
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        core = [self.config.mapping_driver_id_col, self.config.mapping_team_name_col,
                self.config.mapping_leader_name_col]
        self._validate_columns(df, core, "队伍明细")
        df[self.config.mapping_driver_id_col] = df[self.config.mapping_driver_id_col].astype(str).str.strip().str.replace('.0', '', regex=False)
        has_s = self.config.mapping_supervisor_col in df.columns
        static = getattr(self.config, "static_supervisor_mapping", {}) if not has_s else {}
        mapping = {}
        for _, row in df.iterrows():
            did = row[self.config.mapping_driver_id_col]
            leader = row[self.config.mapping_leader_name_col]
            team = row[self.config.mapping_team_name_col]
            sup = row[self.config.mapping_supervisor_col] if has_s else static.get(leader, "")
            mapping[did] = {"队伍名称": team, "队长姓名": leader, "司管": sup}
        return mapping

    # ---------- 流水数据预处理 ----------
    def _prepare_flow(self, path: str, mapping: Dict) -> pd.DataFrame:
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        dc = self.config.flow_driver_id_col
        if dc not in df.columns:
            raise ValidationError(f"缺少司机ID列: {dc}")
        df[dc] = df[dc].astype(str).str.strip().str.replace('.0', '', regex=False)
        df['队伍名称'] = df[dc].map(lambda x: mapping.get(x, {}).get('队伍名称', '未分配'))
        df['队长姓名'] = df[dc].map(lambda x: mapping.get(x, {}).get('队长姓名', ''))
        df['司管'] = df[dc].map(lambda x: mapping.get(x, {}).get('司管', ''))
        return df

    # ---------- 按队伍聚合 ----------
    def _aggregate_by_team(self, df: pd.DataFrame):
        status_col = self.config.flow_status_col
        active = df[df[status_col] == self.config.flow_active_status].copy() if status_col in df.columns else df.copy()
        team_size = df.groupby('队伍名称').size().rename('小队人数').reset_index()

        online_col = self.config.flow_online_days_col
        if online_col in active.columns:
            active[online_col] = pd.to_numeric(active[online_col], errors='coerce').fillna(0)

        value_cols = [
            self.config.flow_order_col,
            self.config.flow_apply_col,
            self.config.flow_transfer_col,
            self.config.flow_new_ju_col,
        ]
        for c in value_cols:
            if c in active.columns:
                active[c] = pd.to_numeric(active[c], errors='coerce').fillna(0)

        agg = {self.config.flow_driver_id_col: 'count'}
        for c in value_cols:
            if c in active.columns:
                agg[c] = 'sum'
        if online_col in active.columns:
            agg[online_col] = 'sum'

        grp = ['队伍名称', '队长姓名', '司管']
        avail = [c for c in grp if c in active.columns]
        summary = active.groupby(avail, as_index=False).agg(agg)
        summary.rename(columns={self.config.flow_driver_id_col: '可上线人数'}, inplace=True)
        if online_col in active.columns:
            summary.rename(columns={online_col: '上线人数'}, inplace=True)
        else:
            summary['上线人数'] = 0

        summary = summary.merge(team_size, on='队伍名称', how='left')
        return summary

    # ---------- 月度数据聚合（多文件） ----------
    def aggregate_monthly(self, file_paths: List[str], mapping: Dict) -> pd.DataFrame:
        """合并多个月度流水文件并聚合队伍月度数据"""
        dfs = []
        for path in file_paths:
            df = self._prepare_flow(path, mapping)
            dfs.append(df)
        if not dfs:
            return pd.DataFrame()
        combined = pd.concat(dfs, ignore_index=True)
        # 按队伍聚合月度数值
        monthly = self._aggregate_by_team(combined)
        return monthly

    # ---------- 生成对比报表数据 ----------
    def build_report_data(self, today_df: pd.DataFrame,
                          yesterday_df: Optional[pd.DataFrame],
                          monthly_df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        today_agg = self._aggregate_by_team(today_df)
        if yesterday_df is not None:
            yesterday_agg = self._aggregate_by_team(yesterday_df)
            report = pd.merge(today_agg, yesterday_agg,
                              on=['队伍名称', '队长姓名', '司管', '小队人数'],
                              how='outer', suffixes=('_今日', '_昨日'))
            # 填充缺失
            for c in report.columns:
                if '_今日' in c or '_昨日' in c:
                    report[c] = report[c].fillna(0).astype(int)
            if '可上线人数_今日' in report.columns and '可上线人数_昨日' in report.columns:
                report['可上线人数'] = report['可上线人数_今日'].fillna(0) + report['可上线人数_昨日'].fillna(0)
            elif '可上线人数_今日' in report.columns:
                report['可上线人数'] = report['可上线人数_今日']
            else:
                report['可上线人数'] = report.get('可上线人数_昨日', 0)
        else:
            report = today_agg.copy()
            report['可上线人数'] = report.get('可上线人数', 0)

        # 上线率
        today_online = '上线人数_今日' if yesterday_df is not None else '上线人数'
        yesterday_online = '上线人数_昨日' if yesterday_df is not None else None
        can_today = '可上线人数_今日' if yesterday_df is not None else '可上线人数'
        can_yesterday = '可上线人数_昨日' if yesterday_df is not None else None
        report['今日上线率'] = np.where(report[can_today] > 0, report[today_online] / report[can_today], 0)
        if yesterday_df is not None:
            report['昨日上线率'] = np.where(report[can_yesterday] > 0, report[yesterday_online] / report[can_yesterday], 0)
            report['上线率同比'] = report['今日上线率'] - report['昨日上线率']
        else:
            report['昨日上线率'] = 0.0
            report['上线率同比'] = 0.0

        # 日完单/报单完成率
        order_today = self.config.flow_order_col + '_今日'
        apply_today = self.config.flow_apply_col + '_今日'
        newju_today = self.config.flow_new_ju_col + '_今日'
        if order_today in report.columns:
            report['日完单目标完成率'] = report[order_today] / self.config.daily_order_target
        if apply_today in report.columns:
            report['日报单目标完成率'] = report[apply_today] / self.config.daily_apply_target
        if newju_today in report.columns:
            safe_online = report[can_today].replace(0, 1)
            report['新桔渗透率'] = report[newju_today] / safe_online
        else:
            report['新桔渗透率'] = 0.0

        # 月度数据合并
        if monthly_df is not None and not monthly_df.empty:
            # 确保月度 aggreg 有队伍名称等键
            monthly_renames = {
                self.config.flow_order_col: '月完单',
                self.config.flow_apply_col: '月报单',
                self.config.flow_transfer_col: '月转单',
                self.config.flow_new_ju_col: '月新桔',
            }
            for old, new in monthly_renames.items():
                if old in monthly_df.columns:
                    monthly_df[new] = monthly_df[old]
            monthly_df = monthly_df[['队伍名称'] + [v for v in monthly_renames.values() if v in monthly_df.columns]]
            report = pd.merge(report, monthly_df, on='队伍名称', how='left')
        # 填充月度列默认值
        for col in ['月完单', '月报单', '月转单', '月新桔']:
            if col not in report.columns:
                report[col] = 0
            else:
                report[col] = report[col].fillna(0).astype(int)

        report.sort_values(['司管', '队伍名称'], inplace=True)
        return report

    # 为 writer 提供标准化的行字典
    def report_to_rows(self, report: pd.DataFrame,
                       daily_order: int, daily_apply: int,
                       monthly_order: int, monthly_apply: int) -> List[Dict]:
        rows = []
        for _, r in report.iterrows():
            row = {
                '司管': r['司管'],
                '队长姓名': r['队长姓名'],
                '队伍名称': r['队伍名称'],
                '小队人数': r['小队人数'],
                '可上线人数': r['可上线人数'],
                '今日上线率': r['今日上线率'],
                '昨日上线率': r['昨日上线率'],
                '上线率同比': r['上线率同比'],
                'daily_order_today': r.get(self.config.flow_order_col+'_今日', 0),
                'daily_apply_today': r.get(self.config.flow_apply_col+'_今日', 0),
                'daily_transfer_today': r.get(self.config.flow_transfer_col+'_今日', 0),
                'daily_newju_today': r.get(self.config.flow_new_ju_col+'_今日', 0),
                'daily_order_yesterday': r.get(self.config.flow_order_col+'_昨日', 0),
                'daily_apply_yesterday': r.get(self.config.flow_apply_col+'_昨日', 0),
                'daily_newju_yesterday': r.get(self.config.flow_new_ju_col+'_昨日', 0),
                '上线人数_今日': r.get('上线人数_今日', 0),
                '上线人数_昨日': r.get('上线人数_昨日', 0),
                '可上线人数_今日': r.get('可上线人数_今日', 0),
                '可上线人数_昨日': r.get('可上线人数_昨日', 0),
                'monthly_order': r.get('月完单', 0),
                'monthly_apply': r.get('月报单', 0),
                'monthly_transfer': r.get('月转单', 0),
                'monthly_newju': r.get('月新桔', 0),
            }
            rows.append(row)
        return rows
    ```

- **services/team_report/writer.py**
  - language: `python`
  - content:
    ```python
# services/team_report/writer.py
"""团队报表 Excel 渲染器（修正完成率）"""
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from services.common.excel_styles import (
    TITLE_FONT, HEADER_FONT_BOLD, HEADER_FONT_RED, NORMAL_FONT, BOLD_FONT,
    GRAY_FILL, GREEN_FILL, RED_FILL, LIGHT_GRAY_FILL, THIN_BORDER,
    CENTER_ALIGN, LEFT_ALIGN
)
import logging

logger = logging.getLogger(__name__)

class ColumnMapping:
    """列定义，支持动态标题和数据提取"""
    def __init__(self, letter: str, title_fn, width: float,
                 data_key: str = None,
                 is_target: bool = False,
                 value_fn: callable = None):
        self.letter = letter
        self._title_fn = title_fn
        self.width = width
        self.data_key = data_key
        self.is_target = is_target
        self._value_fn = value_fn

    @property
    def title(self):
        return self._title_fn() if callable(self._title_fn) else self._title_fn

    def get_value(self, row, extra_context=None):
        if self._value_fn:
            return self._value_fn(row)
        if self.data_key:
            return row.get(self.data_key, 0)
        return 0

class TeamReportExcelWriter:
    """生成车队日报 Excel 文件（输出与原版完全一致）"""
    def __init__(self):
        self._column_defs = []

    def build_columns(self, report_date: str, month: str,
                      daily_order_target: int, daily_apply_target: int,
                      monthly_order_target: int, monthly_apply_target: int):
        day = report_date.split('/')[-1]
        cols = []
        # A 司管
        cols.append(ColumnMapping('A', lambda: '司管', 6.5, data_key='司管'))
        # B 小队长
        cols.append(ColumnMapping('B', lambda: '小队长', 13.88, data_key='队长姓名'))
        # C 队伍
        cols.append(ColumnMapping('C', lambda: '队伍', 16.25, data_key='队伍名称'))
        # D 小队人数
        cols.append(ColumnMapping('D', lambda: '小队人数', 7.63, data_key='小队人数'))
        # E 可上线人数
        cols.append(ColumnMapping('E', lambda: '可上线人数', 7.63, data_key='可上线人数'))
        # F 今日上线率
        cols.append(ColumnMapping('F', lambda: '今日上线率', 7.63, data_key='今日上线率'))
        # G 昨日上线率
        cols.append(ColumnMapping('G', lambda: '昨日上线率', 7.63, data_key='昨日上线率'))
        # H 同比趋势（上线率）
        cols.append(ColumnMapping('H', lambda: '同比趋势（上线率）', 7.63, data_key='上线率同比'))
        # I {day}号完单
        cols.append(ColumnMapping('I', lambda: f'{day}号完单', 6, data_key='daily_order_today'))
        # J 日完单目标完成率（前值/目标值）
        cols.append(ColumnMapping('J', lambda: f'日完单目标值({daily_order_target}单)', 13.5,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('daily_order_today', 0) / daily_order_target)
                                  if daily_order_target else 0))
        # K {day}号报单
        cols.append(ColumnMapping('K', lambda: f'{day}号报单', 6, data_key='daily_apply_today'))
        # L 日报单目标完成率
        cols.append(ColumnMapping('L', lambda: f'日报单目标值({daily_apply_target}单)', 13.5,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('daily_apply_today', 0) / daily_apply_target)
                                  if daily_apply_target else 0))
        # M {day}号转单
        cols.append(ColumnMapping('M', lambda: f'{day}号转单', 6, data_key='daily_transfer_today'))
        # N {day}号新桔在线数
        cols.append(ColumnMapping('N', lambda: f'{day}号新桔在线数', 8.5, data_key='daily_newju_today'))
        # O 新桔渗透率
        cols.append(ColumnMapping('O', lambda: f'{day}号新桔渗透率', 12,
                                  value_fn=lambda r: (r['daily_newju_today'] / r['daily_order_today'])
                                  if r.get('daily_order_today') else 0))
        # P 昨天完单
        cols.append(ColumnMapping('P', lambda: '昨天完单', 6, data_key='daily_order_yesterday'))
        # Q 同比趋势（完单）
        cols.append(ColumnMapping('Q', lambda: '同比趋势（完单）', 12,
                                  value_fn=lambda r: r['daily_order_today'] - r['daily_order_yesterday']))
        # R 昨天报单
        cols.append(ColumnMapping('R', lambda: '昨天报单', 6, data_key='daily_apply_yesterday'))
        # S 同比趋势（报单）
        cols.append(ColumnMapping('S', lambda: '同比趋势（报单）', 12,
                                  value_fn=lambda r: r['daily_apply_today'] - r['daily_apply_yesterday']))
        # T 昨天新桔
        cols.append(ColumnMapping('T', lambda: '昨天新桔', 6, data_key='daily_newju_yesterday'))
        # U 同比趋势（新桔）
        cols.append(ColumnMapping('U', lambda: '同比趋势（新桔）', 12,
                                  value_fn=lambda r: r['daily_newju_today'] - r['daily_newju_yesterday']))
        # V {month}月完单
        cols.append(ColumnMapping('V', lambda: f'{month}月完单', 8.4, data_key='monthly_order'))
        # W 月完单目标完成率
        cols.append(ColumnMapping('W', lambda: f'{month}月完单目标值({monthly_order_target}单)', 8.4,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('monthly_order', 0) / monthly_order_target)
                                  if monthly_order_target else 0))
        # X {month}月报单
        cols.append(ColumnMapping('X', lambda: f'{month}月报单', 8.4, data_key='monthly_apply'))
        # Y 月报单目标完成率
        cols.append(ColumnMapping('Y', lambda: f'{month}月报单目标值({monthly_apply_target}单)', 8.4,
                                  is_target=True,
                                  value_fn=lambda r: (r.get('monthly_apply', 0) / monthly_apply_target)
                                  if monthly_apply_target else 0))
        # Z {month}月转单
        cols.append(ColumnMapping('Z', lambda: f'{month}月转单', 12, data_key='monthly_transfer'))
        # AA {month}月新桔在线数
        cols.append(ColumnMapping('AA', lambda: f'{month}月新桔在线数', 12, data_key='monthly_newju'))
        self._column_defs = cols

    def write(self, rows, output_path: str,
              report_date: str, month: str,
              daily_order: int, daily_apply: int,
              monthly_order: int, monthly_apply: int):
        self.build_columns(report_date, month, daily_order, daily_apply, monthly_order, monthly_apply)
        wb = Workbook()
        ws = wb.active
        ws.title = "每日数据对比"

        # 列宽
        for col_def in self._column_defs:
            ws.column_dimensions[col_def.letter].width = col_def.width

        # 第1行标题
        ws.row_dimensions[1].height = 27
        ws.merge_cells('A1:AA1')
        c = ws['A1']
        c.value = f"日期\t{report_date}\t杭州新享{month}月三大队订单完成情况"
        c.font = TITLE_FONT
        c.alignment = CENTER_ALIGN
        c.fill = GRAY_FILL

        # 第2行子标题
        ws.row_dimensions[2].height = 27
        ws.merge_cells('A2:B2')
        ws['A2'].value = f"月份\t{month}月"
        ws['A2'].font = HEADER_FONT_BOLD
        ws['A2'].alignment = CENTER_ALIGN
        ws['A2'].fill = GRAY_FILL
        ws.merge_cells('C2:E2')
        ws['C2'].fill = GRAY_FILL
        ws.merge_cells('F2:H2')
        ws['F2'].fill = GRAY_FILL
        ws.merge_cells('I2:O2')
        ws['I2'].value = "今日数据"
        ws['I2'].font = HEADER_FONT_BOLD
        ws['I2'].alignment = CENTER_ALIGN
        ws['I2'].fill = GREEN_FILL
        ws.merge_cells('P2:U2')
        ws['P2'].value = "昨日数据"
        ws['P2'].font = HEADER_FONT_BOLD
        ws['P2'].alignment = CENTER_ALIGN
        ws['P2'].fill = RED_FILL
        ws.merge_cells('V2:AA2')
        ws['V2'].value = f"{month}月数据"
        ws['V2'].font = HEADER_FONT_BOLD
        ws['V2'].alignment = CENTER_ALIGN
        ws['V2'].fill = GRAY_FILL

        # 第3行表头
        ws.row_dimensions[3].height = 54
        for idx, col_def in enumerate(self._column_defs):
            col_idx = idx + 1
            cell = ws.cell(row=3, column=col_idx, value=col_def.title)
            cell.font = HEADER_FONT_RED if col_def.is_target else HEADER_FONT_BOLD
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER
            if col_idx <= 5:
                cell.fill = GRAY_FILL
            elif 6 <= col_idx <= 8:
                cell.fill = GRAY_FILL
            elif 9 <= col_idx <= 15:
                cell.fill = GREEN_FILL
            elif 16 <= col_idx <= 21:
                cell.fill = RED_FILL
            else:
                cell.fill = GRAY_FILL

        current_row = 4
        start_data_row = 4
        subtotal_rows = []
        supervisors = sorted(list({r.get('司管', '') for r in rows}))
        global_on_today = 0
        global_on_yesterday = 0
        global_can_today = 0
        global_can_yesterday = 0

        # 数据行
        for sup in supervisors:
            group = [r for r in rows if r.get('司管', '') == sup]
            first_group_row = current_row
            sum_on_today = 0
            sum_on_yesterday = 0
            sum_can_today = 0
            sum_can_yesterday = 0
            for row in group:
                for col_idx, col_def in enumerate(self._column_defs, start=1):
                    val = col_def.get_value(row)
                    cell = ws.cell(row=current_row, column=col_idx, value=val)
                    if col_idx in (6, 7, 8, 10, 12, 15, 23, 25):
                        cell.number_format = '0.00%'
                    cell.font = NORMAL_FONT
                    cell.alignment = CENTER_ALIGN if col_idx > 3 else LEFT_ALIGN
                    cell.border = THIN_BORDER
                sum_on_today += row.get('上线人数_今日', 0)
                sum_on_yesterday += row.get('上线人数_昨日', 0)
                sum_can_today += row.get('可上线人数_今日', 0)
                sum_can_yesterday += row.get('可上线人数_昨日', 0)
                current_row += 1

            global_on_today += sum_on_today
            global_on_yesterday += sum_on_yesterday
            global_can_today += sum_can_today
            global_can_yesterday += sum_can_yesterday

            if first_group_row < current_row:
                ws.merge_cells(start_row=first_group_row, start_column=1,
                               end_row=current_row-1, end_column=1)

            # 司管小计行
            subtotal_rows.append(current_row)
            for col_idx in range(1, 28):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.font = BOLD_FONT
                cell.border = THIN_BORDER
                cell.fill = GRAY_FILL
                cell.alignment = CENTER_ALIGN
            ws.cell(row=current_row, column=1, value="总计")

            # 求和列
            sum_cols_indices = [4, 5, 9, 11, 13, 14, 16, 17, 18, 19, 20, 21, 22, 24, 26, 27]
            for ci in sum_cols_indices:
                total = 0
                for r in range(first_group_row, current_row):
                    val = ws.cell(row=r, column=ci).value
                    if isinstance(val, (int, float)):
                        total += val
                ws.cell(row=current_row, column=ci, value=total)

            # 小计行完成率
            total_on_today = sum_on_today
            total_on_yesterday = sum_on_yesterday
            total_can_today = sum_can_today if sum_can_today else 1
            total_can_yesterday = sum_can_yesterday if sum_can_yesterday else 1
            ws.cell(row=current_row, column=6, value=total_on_today / total_can_today).number_format = '0.00%'
            ws.cell(row=current_row, column=7, value=total_on_yesterday / total_can_yesterday).number_format = '0.00%'
            ws.cell(row=current_row, column=8, value=(total_on_today / total_can_today) - (total_on_yesterday / total_can_yesterday)).number_format = '0.00%'
            total_order_today = ws.cell(row=current_row, column=9).value or 0
            ws.cell(row=current_row, column=10, value=total_order_today / daily_order if daily_order else 0).number_format = '0.00%'
            total_apply_today = ws.cell(row=current_row, column=11).value or 0
            ws.cell(row=current_row, column=12, value=total_apply_today / daily_apply if daily_apply else 0).number_format = '0.00%'
            total_newju = ws.cell(row=current_row, column=14).value or 0
            ws.cell(row=current_row, column=15, value=total_newju / total_order_today if total_order_today else 0).number_format = '0.00%'
            total_monthly_order = ws.cell(row=current_row, column=22).value or 0
            ws.cell(row=current_row, column=23, value=total_monthly_order / monthly_order if monthly_order else 0).number_format = '0.00%'
            total_monthly_apply = ws.cell(row=current_row, column=24).value or 0
            ws.cell(row=current_row, column=25, value=total_monthly_apply / monthly_apply if monthly_apply else 0).number_format = '0.00%'
            current_row += 1

        # 全部总计行
        grand_total_row = current_row
        for col_idx in range(1, 28):
            cell = ws.cell(row=grand_total_row, column=col_idx)
            cell.font = BOLD_FONT
            cell.border = THIN_BORDER
            cell.fill = GRAY_FILL
            cell.alignment = CENTER_ALIGN
        ws.cell(row=grand_total_row, column=1, value="总计")

        sum_cols_indices_all = [4, 5, 9, 11, 13, 14, 16, 17, 18, 19, 20, 21, 22, 24, 26, 27]
        for ci in sum_cols_indices_all:
            total = 0
            for r in range(start_data_row, grand_total_row):
                if r in subtotal_rows:
                    continue
                val = ws.cell(row=r, column=ci).value
                if isinstance(val, (int, float)):
                    total += val
            ws.cell(row=grand_total_row, column=ci, value=total)

        # 全部总计完成率
        all_can_today = global_can_today if global_can_today else 1
        all_can_yesterday = global_can_yesterday if global_can_yesterday else 1
        ws.cell(row=grand_total_row, column=6, value=global_on_today / all_can_today).number_format = '0.00%'
        ws.cell(row=grand_total_row, column=7, value=global_on_yesterday / all_can_yesterday).number_format = '0.00%'
        ws.cell(row=grand_total_row, column=8, value=(global_on_today / all_can_today) - (global_on_yesterday / all_can_yesterday)).number_format = '0.00%'
        total_order_all = ws.cell(row=grand_total_row, column=9).value or 0
        ws.cell(row=grand_total_row, column=10, value=total_order_all / daily_order if daily_order else 0).number_format = '0.00%'
        total_apply_all = ws.cell(row=grand_total_row, column=11).value or 0
        ws.cell(row=grand_total_row, column=12, value=total_apply_all / daily_apply if daily_apply else 0).number_format = '0.00%'
        total_newju_all = ws.cell(row=grand_total_row, column=14).value or 0
        ws.cell(row=grand_total_row, column=15, value=total_newju_all / total_order_all if total_order_all else 0).number_format = '0.00%'
        total_monthly_order_all = ws.cell(row=grand_total_row, column=22).value or 0
        ws.cell(row=grand_total_row, column=23, value=total_monthly_order_all / monthly_order if monthly_order else 0).number_format = '0.00%'
        total_monthly_apply_all = ws.cell(row=grand_total_row, column=24).value or 0
        ws.cell(row=grand_total_row, column=25, value=total_monthly_apply_all / monthly_apply if monthly_apply else 0).number_format = '0.00%'
        current_row = grand_total_row + 1

        # 负值标红
        for r in range(start_data_row, current_row):
            for ci in [8, 9, 11, 13, 14, 16, 17, 18, 19, 20, 21, 22, 24, 26, 27]:
                cell = ws.cell(row=r, column=ci)
                if isinstance(cell.value, (int, float)) and cell.value < 0:
                    cell.font = Font(name='微软雅黑', size=10, color="FF0000")

        # 备注行
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=27)
        note_cell = ws.cell(row=current_row, column=1)
        note_cell.value = "表格中标红色的表示数值低于该小队当日该数据的平均值，需要根据数值所在的单元格注意队员的接单情况。"
        note_cell.font = NORMAL_FONT
        note_cell.fill = LIGHT_GRAY_FILL
        note_cell.alignment = CENTER_ALIGN

        wb.save(output_path)
        return output_path
    ```

- **services/transposer/service.py**
  - language: `python`
  - content:
    ```python
# services/transposer/service.py
import pandas as pd
from typing import List, Optional


def pivot_table(df: pd.DataFrame,
                index_cols: Optional[List[str]],
                category_col: str,
                value_cols: List[str]) -> pd.DataFrame:
    """
    将分类列的值转置为新的列名，同时保留其他固定列。

    :param df: 原始数据
    :param index_cols: 索引列列表；若为空则自动生成组内序号（此时固定列不会被保留）
    :param category_col: 分类列名（如“医疗类别”）
    :param value_cols: 值列名列表（如“a数据”，“b数据”）
    :return: 转置后的 DataFrame，固定列在左侧，拼接后的新值列严格按“类别+值列名”排列（无空格）
    """
    df = df.copy()

    # 如果未提供索引列，自动生成组内序号，并在此情况下不保留固定列
    auto_index = False
    if not index_cols:
        df['_idx'] = df.groupby(category_col).cumcount()
        index_cols = ['_idx']
        auto_index = True

    # 区分固定列：所有列 - category_col - value_cols
    # 注意：value_cols 中可能包含已被选为索引的列，但用户应避免错误输入
    other_cols = [c for c in df.columns 
                  if c not in value_cols and c != category_col]

    # 如果使用自动序号，则不保留任何固定列（因为固定列与自动序号没有实际关联）
    if auto_index:
        other_cols = []
    else:
        # 确保索引列被包含在 other_cols 中（通常已包含）
        # 并去除重复
        other_cols = list(dict.fromkeys(other_cols))

    # 取出固定列的唯一值（按 index_cols 去重）
    if other_cols:
        fixed_df = df[other_cols].drop_duplicates(subset=index_cols)
        # 将固定列设置为索引以便后续合并
        fixed_df = fixed_df.set_index(index_cols)
    else:
        fixed_df = pd.DataFrame()

    # 执行 pivot：包括索引列、分类列、值列
    pivot_input = df[index_cols + [category_col] + value_cols]
    pivoted = pivot_input.pivot(index=index_cols, columns=category_col, values=value_cols)

    # 扁平化多层列名并去除空格，连接方式为“类别+值列名”，不包含任何分隔符
    new_columns = []
    for col in pivoted.columns:
        if isinstance(col, tuple):
            # 元组顺序可能是 (value, category) 或 (category, value)，需要统一
            # 通过检查 col 中哪个是 value_cols 中的列来确定顺序
            cat, val = None, None
            for part in col:
                if part in value_cols:
                    val = str(part).strip()
                else:
                    cat = str(part).strip()
            if cat and val:
                new_col = f"{cat}{val}"
                new_columns.append(new_col)
            else:
                # 退化情况
                new_columns.append(''.join(str(x).strip() for x in col))
        else:
            new_columns.append(str(col).strip())
    pivoted.columns = new_columns

    # 重置索引以便与固定列合并
    pivoted.reset_index(inplace=True)
    # 删除可能由自动索引添加的 '_idx' 列
    if '_idx' in pivoted.columns:
        pivoted.drop(columns=['_idx'], inplace=True)

    # 合并固定列
    if not fixed_df.empty:
        # 将 pivoted 的索引列也设为索引以执行合并
        pivoted_merge = pivoted.set_index(index_cols)
        result = fixed_df.join(pivoted_merge, how='right')
        result.reset_index(inplace=True)
    else:
        result = pivoted

    # 重新排序列：固定列（other_cols 中出现在 result 中的列）→ 新的值列
    fixed_cols_in_result = [c for c in other_cols if c in result.columns]
    new_value_cols = [c for c in result.columns if c not in fixed_cols_in_result]
    result = result[fixed_cols_in_result + new_value_cols]

    # 对新的值列按“类别+值列名”排序（例如：门诊a数据、门诊b数据、住院a数据、住院b数据、其他a数据、其他b数据）
    # 此处保持原顺序，仅确保无空白
    return result
    ```

- **static/css/checker.css**
  - language: `css`
  - content:
    ```css
/* 文件校验模块样式 */
.mapping-table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
}
.mapping-table th, .mapping-table td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
}
.compare-pair {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
}
.compare-pair input {
    flex: 1;
    max-width: 200px;
}
.compare-pair span {
    font-weight: bold;
    font-size: 1.2em;
}
.btn-remove-pair {
    padding: 5px 10px;
    background: #ef476f;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
}
.precheck-result {
    background: #f8f9fa;
    padding: 10px;
    border-radius: 8px;
    margin: 10px 0;
}
    ```

- **static/css/style.css**
  - language: `css`
  - content:
    ```css
/* ========== 全局变量与重置 ========== */
:root {
    --primary: #4361ee;
    --primary-dark: #3a56d4;
    --primary-light: #eef2ff;
    --secondary: #2ec4b6;
    --success: #06d6a0;
    --danger: #ef476f;
    --warning: #ffd166;
    --dark: #1e293b;
    --gray: #64748b;
    --light: #f8fafc;
    --border: #e2e8f0;
    --shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
    --shadow-lg: 0 10px 30px rgba(0, 0, 0, 0.1);
    --radius: 12px;
    --radius-sm: 8px;
    --transition: all 0.2s ease;
}

* {
    box-sizing: border-box;
}

body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, sans-serif;
    margin: 0;
    padding: 24px;
    background: linear-gradient(145deg, #f1f5f9 0%, #e6edf5 100%);
    min-height: 100vh;
    color: var(--dark);
    line-height: 1.6;
}

.container {
    max-width: 1280px;
    margin: 0 auto;
    background: rgba(255, 255, 255, 0.85);
    backdrop-filter: blur(10px);
    padding: 32px 36px;
    border-radius: 28px;
    box-shadow: var(--shadow-lg);
    border: 1px solid rgba(255, 255, 255, 0.7);
}

h1 {
    font-size: 2.2rem;
    font-weight: 700;
    margin-top: 0;
    margin-bottom: 28px;
    color: var(--dark);
    display: flex;
    align-items: center;
    gap: 12px;
    border-bottom: 2px solid var(--primary-light);
    padding-bottom: 16px;
}

h2 {
    font-size: 1.6rem;
    font-weight: 600;
    margin-top: 0;
    margin-bottom: 24px;
    color: var(--dark);
}

h3 {
    font-size: 1.2rem;
    font-weight: 600;
    margin-bottom: 16px;
    color: var(--gray);
}

/* 导航栏 */
nav ul {
    list-style: none;
    padding: 0;
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 8px;
    margin-bottom: 32px;
    background: var(--light);
    padding: 8px;
    border-radius: 48px;
    border: 1px solid var(--border);
}

nav ul li a {
    display: inline-block;
    padding: 10px 20px;
    text-decoration: none;
    color: var(--gray);
    font-weight: 500;
    border-radius: 40px;
    transition: var(--transition);
    background: transparent;
}

nav ul li a:hover {
    background: white;
    color: var(--primary);
    box-shadow: var(--shadow);
}

/* 表单通用 */
.form-section {
    background: white;
    border-radius: var(--radius);
    padding: 24px;
    margin-bottom: 28px;
    box-shadow: var(--shadow);
    border: 1px solid var(--border);
}

.form-row {
    margin-bottom: 20px;
}

label {
    font-weight: 600;
    display: block;
    margin-bottom: 8px;
    color: var(--dark);
    font-size: 0.95rem;
}

input, select, textarea {
    width: 100%;
    padding: 12px 16px;
    border: 1.5px solid var(--border);
    border-radius: var(--radius-sm);
    font-size: 0.95rem;
    transition: var(--transition);
    background: white;
}

input:focus, select:focus, textarea:focus {
    outline: none;
    border-color: var(--primary);
    box-shadow: 0 0 0 3px rgba(67, 97, 238, 0.15);
}

.hint, .help-text {
    font-size: 0.85rem;
    color: var(--gray);
    margin-top: 6px;
}

/* 按钮 */
.btn, button {
    background: var(--primary);
    color: white;
    padding: 12px 24px;
    border: none;
    border-radius: 40px;
    cursor: pointer;
    font-size: 1rem;
    font-weight: 600;
    transition: var(--transition);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    box-shadow: 0 2px 6px rgba(67, 97, 238, 0.3);
}

.btn:hover, button:hover {
    background: var(--primary-dark);
    transform: translateY(-2px);
    box-shadow: 0 6px 12px rgba(67, 97, 238, 0.25);
}

.btn-success {
    background: var(--success);
    box-shadow: 0 2px 6px rgba(6, 214, 160, 0.3);
}
.btn-success:hover { background: #05b384; }

.btn-danger {
    background: var(--danger);
    box-shadow: 0 2px 6px rgba(239, 71, 111, 0.3);
}
.btn-danger:hover { background: #d63e5e; }

.btn-secondary {
    background: var(--secondary);
    box-shadow: 0 2px 6px rgba(46, 196, 182, 0.3);
}
.btn-secondary:hover { background: #25a99a; }

/* 消息与状态 */
.message {
    margin-top: 20px;
    padding: 14px 18px;
    border-radius: var(--radius-sm);
    font-weight: 500;
}
.message.error { background: #fee2e2; color: #b91c1c; }
.message.success { background: #d1fae5; color: #065f46; }
.message.info { background: #dbeafe; color: #1e40af; }

.status-area {
    margin-top: 24px;
    padding: 24px;
    background: white;
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    border: 1px solid var(--border);
}

.progress {
    width: 100%;
    background: var(--border);
    border-radius: 20px;
    height: 10px;
    overflow: hidden;
}
.progress-bar {
    height: 10px;
    background: var(--primary);
    width: 0%;
    transition: width 0.3s;
}

.log-area {
    background: #0f172a;
    color: #a5f3fc;
    padding: 16px;
    border-radius: var(--radius-sm);
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.85rem;
    max-height: 240px;
    overflow-y: auto;
    white-space: pre-wrap;
}

/* 筛选条件行 */
.filter-row {
    display: flex;
    gap: 12px;
    margin-bottom: 16px;
    align-items: center;
    flex-wrap: wrap;
}
.filter-field { width: 180px; }
.filter-op { width: 120px; }
.filter-value { width: 240px; }

/* 表格样式（用于匹配填充等） */
.table-container {
    overflow-x: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    border-radius: var(--radius);
    overflow: hidden;
    box-shadow: var(--shadow);
}
th, td {
    padding: 12px 16px;
    text-align: left;
    border-bottom: 1px solid var(--border);
}
th {
    background: var(--primary-light);
    font-weight: 600;
}

/* 响应式 */
@media (max-width: 768px) {
    body { padding: 12px; }
    .container { padding: 20px; }
    nav ul { flex-direction: column; align-items: stretch; }
    nav ul li a { display: block; text-align: center; }
}
    ```

- **static/js/checker.js**
  - language: `javascript`
  - content:
    ```javascript
// checker.js
(function() {
    // 确保 DOM 加载完成后再绑定事件
    document.addEventListener('DOMContentLoaded', function() {
        // 定义全局 addFilter 函数
        window.addFilter = function() {
            const container = document.getElementById('filtersContainer');
            const template = document.getElementById('filterRowTemplate');
            if (template) container.appendChild(template.content.cloneNode(true));
        };

        // 解析列名
        function parseColumns(text) {
            if (!text) return [];
            const lines = text.split(/[\n\r]+/);
            const cols = [];
            lines.forEach(line => {
                const parts = line.split(/[,，\s]+/);
                parts.forEach(p => {
                    const trimmed = p.trim();
                    if (trimmed) cols.push(trimmed);
                });
            });
            return cols;
        }

        // 收集对比列映射
        function collectCompareMap() {
            const stdText = document.getElementById('compareStandardText').value;
            const cmpText = document.getElementById('compareCompareText').value;
            const stdCols = parseColumns(stdText);
            let cmpCols = parseColumns(cmpText);
            if (cmpCols.length === 0) cmpCols = [...stdCols];
            if (stdCols.length === 0) {
                alert('请填写标准对比列');
                return null;
            }
            if (stdCols.length !== cmpCols.length) {
                alert(`标准对比列数量 (${stdCols.length}) 与比较对比列数量 (${cmpCols.length}) 不一致，请检查输入。`);
                return null;
            }
            const map = [];
            for (let i = 0; i < stdCols.length; i++) {
                map.push({ standard: stdCols[i], compare: cmpCols[i] });
            }
            return map;
        }

        // 获取 DOM 元素
        const btnPrecheck = document.getElementById('btnPrecheck');
        const btnStart = document.getElementById('btnStart');
        const statusArea = document.getElementById('statusArea');
        const taskStatus = document.getElementById('taskStatus');
        const progressBar = document.getElementById('progressBar');
        const progressMessage = document.getElementById('progressMessage');
        const precheckResult = document.getElementById('precheckResult');
        const downloadLink = document.getElementById('downloadLink');

        // 预检按钮点击事件
        btnPrecheck.addEventListener('click', async function() {
            const filesA = document.getElementById('filesA').files;
            const filesB = document.getElementById('filesB').files;
            if (filesA.length === 0 || filesB.length === 0) {
                alert('请上传文件');
                return;
            }

            btnPrecheck.disabled = true;
            btnPrecheck.textContent = '⏳ 预检中...';
            statusArea.style.display = 'block';
            taskStatus.textContent = '正在预检，请稍后...';
            progressBar.style.width = '10%';
            progressMessage.textContent = '正在上传文件并扫描...';
            precheckResult.innerHTML = '';

            const formData = new FormData();
            for (let f of filesA) formData.append('files_a', f);
            for (let f of filesB) formData.append('files_b', f);

            try {
                const resp = await fetch('/checker/precheck', { method: 'POST', body: formData });
                const result = await resp.json();
                if (resp.ok) {
                    precheckResult.innerHTML = `
                        <p>标准文件: ${result.total_files_a} 个, 约 ${result.total_rows_a.toLocaleString()} 行</p>
                        <p>比较文件: ${result.total_files_b} 个, 约 ${result.total_rows_b.toLocaleString()} 行</p>
                        <p>预估内存: ${result.estimated_memory_mb} MB | 预估时间: ${result.estimated_time_seconds} 秒</p>
                        <p style="color:${result.advice.includes('安全') ? 'green' : 'orange'}">${result.advice}</p>
                    `;
                    taskStatus.textContent = '预检完成，可以开始校验';
                    progressBar.style.width = '100%';
                    btnStart.disabled = false;
                } else {
                    taskStatus.textContent = '预检失败';
                    precheckResult.innerHTML = `<p style="color:red;">错误: ${result.message || result.error}</p>`;
                    btnStart.disabled = true;
                }
            } catch (err) {
                taskStatus.textContent = '网络错误';
                precheckResult.innerHTML = `<p style="color:red;">网络请求失败: ${err.message}</p>`;
                btnStart.disabled = true;
            } finally {
                btnPrecheck.disabled = false;
                btnPrecheck.textContent = '🔎 预检';
            }
        });

        // 开始校验按钮点击事件
        btnStart.addEventListener('click', async function() {
            const filesA = document.getElementById('filesA').files;
            const filesB = document.getElementById('filesB').files;
            const keyA = document.getElementById('keyMapA').value.trim();
            const keyB = document.getElementById('keyMapB').value.trim();
            const duplicateAction = document.getElementById('duplicateAction').value;
            const compareMap = collectCompareMap();
            if (!compareMap) return;

            const filters = [];
            document.querySelectorAll('#filtersContainer .filter-row').forEach(row => {
                const field = row.querySelector('input[name="filter_field"]')?.value.trim();
                const op = row.querySelector('select[name="filter_op"]')?.value;
                const value = row.querySelector('input[name="filter_value"]')?.value.trim();
                if (field && value) filters.push({field, op, value});
            });

            if (!keyA) { alert('请填写标准群主键'); return; }

            const formData = new FormData();
            for (let f of filesA) formData.append('files_a', f);
            for (let f of filesB) formData.append('files_b', f);
            formData.append('key_map_a', keyA);
            formData.append('key_map_b', keyB || keyA);
            formData.append('compare_map', JSON.stringify(compareMap));
            formData.append('filters_a', JSON.stringify(filters));
            formData.append('duplicate_action', duplicateAction);

            btnStart.disabled = true;
            btnStart.textContent = '提交中...';
            statusArea.style.display = 'block';
            taskStatus.textContent = '提交中...';
            progressBar.style.width = '0%';
            progressMessage.textContent = '';

            const resp = await fetch('/checker/process', { method: 'POST', body: formData });
            const result = await resp.json();
            if (result.error) {
                taskStatus.textContent = '错误: ' + result.error;
                btnStart.disabled = false;
                btnStart.textContent = '▶ 开始校验';
                return;
            }
            const taskId = result.task_id;
            taskStatus.textContent = '任务已提交，处理中...';

            const poll = setInterval(async () => {
                try {
                    const statusResp = await fetch('/checker/status/' + taskId);
                    const status = await statusResp.json();
                    if (status.error) {
                        clearInterval(poll);
                        taskStatus.textContent = '错误: ' + status.error;
                        btnStart.disabled = false;
                        btnStart.textContent = '▶ 开始校验';
                        return;
                    }
                    if (status.progress_percent !== undefined && status.progress_percent !== null) {
                        progressBar.style.width = status.progress_percent + '%';
                    }
                    if (status.progress_message) {
                        progressMessage.textContent = status.progress_message;
                    }
                    if (status.processed_rows !== undefined && status.total_rows) {
                        taskStatus.textContent = `扫描中 (${status.processed_rows.toLocaleString()} / ${status.total_rows.toLocaleString()} 行)`;
                    }
                    if (status.status === 'completed') {
                        clearInterval(poll);
                        taskStatus.textContent = '完成！';
                        progressBar.style.width = '100%';
                        downloadLink.innerHTML = `<a href="/checker/download/${taskId}" class="btn btn-success">下载结果</a>`;
                        btnStart.disabled = true;
                    } else if (status.status === 'failed') {
                        clearInterval(poll);
                        taskStatus.textContent = '失败: ' + (status.error || '');
                        btnStart.disabled = false;
                        btnStart.textContent = '▶ 开始校验';
                    }
                } catch (err) {
                    console.error('轮询失败', err);
                }
            }, 1000);
        });
    });
})();
    ```

- **static/js/convert.js**
  - language: `javascript`
  - content:
    ```javascript
// convert.js
(function() {
    document.getElementById('convertForm').addEventListener('submit', async function(e) {
        e.preventDefault();
        const fileInput = document.getElementById('fileInput');
        const targetFormat = document.getElementById('targetFormat').value;
        const encoding = document.getElementById('encoding').value;
        const separator = document.getElementById('separator').value;
        if (!fileInput.files.length) { alert('请选择文件'); return; }

        const btn = document.getElementById('btnSubmit');
        btn.disabled = true;
        btn.textContent = '提交中...';
        document.getElementById('statusArea').style.display = 'block';
        document.getElementById('taskStatus').textContent = '提交中...';
        document.getElementById('progressBar').style.width = '0%';

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('target_format', targetFormat);
        formData.append('encoding', encoding);
        formData.append('separator', separator);

        try {
            const resp = await fetch('/convert/process', { method: 'POST', body: formData });
            const result = await resp.json();
            if (result.error) {
                document.getElementById('taskStatus').textContent = '错误: ' + result.error;
                btn.disabled = false;
                return;
            }
            const taskId = result.task_id;
            document.getElementById('taskStatus').textContent = '转换中...';

            const poll = setInterval(async () => {
                const statusResp = await fetch('/convert/status/' + taskId);
                const status = await statusResp.json();
                if (status.error) {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '错误: ' + status.error;
                    btn.disabled = false;
                    return;
                }
                if (status.progress_percent !== undefined) {
                    document.getElementById('progressBar').style.width = status.progress_percent + '%';
                }
                if (status.progress_message) {
                    document.getElementById('progressMessage').textContent = status.progress_message;
                }
                if (status.status === 'completed') {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '完成！';
                    document.getElementById('downloadLink').innerHTML = `<a href="/convert/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    btn.disabled = false;
                    btn.textContent = '开始转换';
                } else if (status.status === 'failed') {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '失败: ' + (status.error || '');
                    btn.disabled = false;
                    btn.textContent = '开始转换';
                }
            }, 1000);
        } catch (err) {
            document.getElementById('taskStatus').textContent = '网络错误: ' + err.message;
            btn.disabled = false;
        }
    });
})();
    ```

- **static/js/download_utils.js**
  - language: `javascript`
  - content:
    ```javascript
function getFilenameFromContentDisposition(disposition) {
    if (!disposition) return null;
    const filenameStarMatch = disposition.match(/filename\*=(?:UTF-8'')?([^;]+)/i);
    if (filenameStarMatch) {
        try {
            return decodeURIComponent(filenameStarMatch[1].trim().replace(/['"]+/g, ''));
        } catch (e) {
            return filenameStarMatch[1].trim().replace(/['"]+/g, '');
        }
    }
    const filenameMatch = disposition.match(/filename="?([^";]+)"?/i);
    return filenameMatch ? filenameMatch[1] : null;
}
    ```

- **static/js/field_parser.js**
  - language: `javascript`
  - content:
    ```javascript
// field_parser.js
(function() {
    let lastResult = null;   // 保存最近分析结果（含 matched, only_a, only_b, cols_a, cols_b）

    document.getElementById('btnAnalyze').addEventListener('click', async function() {
        const textA = document.getElementById('textA').value.trim();
        const textB = document.getElementById('textB').value.trim();
        if (!textA || !textB) {
            alert('请填写两边的列名');
            return;
        }

        const btn = document.getElementById('btnAnalyze');
        btn.disabled = true;
        btn.textContent = '分析中...';

        try {
            const resp = await fetch('/field-parser/analyze', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    standard_headers: textA,
                    compare_headers: textB
                })
            });

            const result = await resp.json();
            if (resp.ok) {
                lastResult = result;   // 保存完整结果
                displayResult(result);
                document.getElementById('btnExport').style.display = 'inline-block';
            } else {
                alert('分析失败: ' + (result.message || result.error || '未知错误'));
            }
        } catch (err) {
            alert('网络请求失败: ' + err.message);
        } finally {
            btn.disabled = false;
            btn.textContent = '开始分析';
        }
    });

    document.getElementById('btnExport').addEventListener('click', async function() {
        if (!lastResult) {
            alert('请先进行分析');
            return;
        }
        // 发送完整映射结果和原始列名
        const payload = {
            matched: lastResult.matched,
            only_a: lastResult.only_a,
            only_b: lastResult.only_b,
            cols_a: lastResult.cols_a,
            cols_b: lastResult.cols_b
        };

        try {
            const resp = await fetch('/field-parser/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (!resp.ok) {
                const err = await resp.json();
                alert('导出失败: ' + (err.message || '未知错误'));
                return;
            }
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const filename = getFilenameFromContentDisposition(resp.headers.get('Content-Disposition')) || '字段映射结果.xlsx';
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (err) {
            alert('导出请求失败: ' + err.message);
        }
    });

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function displayResult(mapping) {
        const tbody = document.querySelector('#mappingTable tbody');
        tbody.innerHTML = '';
        mapping.matched.forEach(m => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${escapeHtml(m.standard)}</td><td>${escapeHtml(m.compare)}</td><td>${m.score}</td>`;
            tbody.appendChild(tr);
        });
        mapping.only_a.forEach(h => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${escapeHtml(h)}</td><td>-</td><td>仅标准</td>`;
            tbody.appendChild(tr);
        });
        mapping.only_b.forEach(h => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>-</td><td>${escapeHtml(h)}</td><td>仅比较</td>`;
            tbody.appendChild(tr);
        });
        document.getElementById('resultArea').style.display = 'block';
    }
})();
    ```

- **static/js/main.js**
  - language: `javascript`
  - content:
    ```javascript
// 全局交互逻辑
document.addEventListener('DOMContentLoaded', () => {
    console.log('✨ Excel 智能工具集已加载');

    // 为所有提交按钮添加加载状态（可选）
    document.querySelectorAll('form').forEach(form => {
        form.addEventListener('submit', () => {
            const btn = form.querySelector('button[type="submit"]');
            if (btn) {
                btn.disabled = true;
                btn.innerHTML = '<span class="spinner"></span> 处理中...';
            }
        });
    });
});
    ```

- **static/js/subtotal.js**
  - language: `javascript`
  - content:
    ```javascript
// static/js/subtotal.js
(function() {
    const form = document.getElementById('subtotalForm');
    const statusArea = document.getElementById('statusArea');
    const addOverrideBtn = document.getElementById('addOverrideBtn');
    const overrideContainer = document.getElementById('overrideContainer');
    const sheetsModeRadios = document.getElementsByName('sheets_mode');
    const sheetsCustom = document.getElementById('sheetsCustom');

    addOverrideBtn.addEventListener('click', function() {
        const template = document.getElementById('overrideRowTemplate');
        const clone = template.content.cloneNode(true);
        overrideContainer.appendChild(clone);
    });

    sheetsModeRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            sheetsCustom.style.display = this.value === 'custom' ? 'inline-block' : 'none';
        });
    });

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        // 收集特殊列函数
        const overrides = [];
        document.querySelectorAll('.override-row').forEach(row => {
            const col = row.querySelector('.override-col').value.trim().toUpperCase();
            const func = row.querySelector('.override-func').value;
            if (col) {
                overrides.push(col + ':' + func);
            }
        });

        const formData = new FormData(form);
        if (overrides.length > 0) {
            formData.set('agg_overrides', overrides.join(','));
        }
        const mode = formData.get('sheets_mode');
        if (mode === 'custom') {
            const customVal = sheetsCustom.value.trim();
            formData.set('sheets_mode', customVal || '1');
        }

        const submitBtn = form.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        statusArea.style.display = 'block';
        document.getElementById('taskStatus').textContent = '提交中...';
        document.getElementById('progressBar').style.width = '0%';
        document.getElementById('downloadLink').innerHTML = '';

        try {
            const resp = await fetch('/subtotal/process', { method: 'POST', body: formData });
            if (!resp.ok) {
                const err = await resp.json();
                document.getElementById('taskStatus').textContent = '错误: ' + (err.error || err.message);
                submitBtn.disabled = false;
                return;
            }
            const contentType = resp.headers.get('Content-Type');
            if (contentType && contentType.includes('application/json')) {
                // 异步任务
                const result = await resp.json();
                const taskId = result.task_id;
                pollTask(taskId, submitBtn);
            } else {
                // 同步下载
                    const blob = await resp.blob();
                const url = URL.createObjectURL(blob);
                const filename = getFilenameFromContentDisposition(resp.headers.get('Content-Disposition')) || 'subtotal_result.xlsx';
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                document.getElementById('taskStatus').textContent = '处理完成，已下载。';
                submitBtn.disabled = false;
            }
        } catch (err) {
            document.getElementById('taskStatus').textContent = '网络错误: ' + err.message;
            submitBtn.disabled = false;
        }
    });

    function pollTask(taskId, submitBtn) {
        const interval = setInterval(async () => {
            try {
                const resp = await fetch('/subtotal/status/' + taskId);
                const status = await resp.json();
                if (status.error) {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '错误: ' + status.error;
                    submitBtn.disabled = false;
                    return;
                }
                if (status.progress_percent !== undefined) {
                    document.getElementById('progressBar').style.width = status.progress_percent + '%';
                }
                document.getElementById('taskStatus').textContent =
                    status.status === 'completed' ? '完成' : (status.status || '处理中...');
                if (status.status === 'completed') {
                    clearInterval(interval);
                    document.getElementById('downloadLink').innerHTML =
                        `<a href="/subtotal/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    submitBtn.disabled = false;
                } else if (status.status === 'failed') {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '失败';
                    submitBtn.disabled = false;
                }
            } catch (err) {
                console.error(err);
            }
        }, 2000);
    }
})();
    ```

- **static/js/team_report.js**
  - language: `javascript`
  - content:
    ```javascript
// team_report.js (updated to include monthly files)
(function() {
    const form = document.getElementById('reportForm');
    if (!form) return;

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const submitBtn = document.getElementById('submitBtn');
        const statusArea = document.getElementById('statusArea');
        const taskStatus = document.getElementById('taskStatus');
        const progressBar = document.getElementById('progressBar');
        const progressMessage = document.getElementById('progressMessage');

        submitBtn.disabled = true;
        statusArea.style.display = 'block';
        taskStatus.textContent = '提交中...';
        progressBar.style.width = '0%';
        progressMessage.textContent = '';

        const formData = new FormData(form);

        try {
            const resp = await fetch('/team-report/generate', {
                method: 'POST',
                body: formData
            });
            if (!resp.ok) {
                taskStatus.textContent = '提交失败 (HTTP ' + resp.status + ')';
                submitBtn.disabled = false;
                return;
            }
            const result = await resp.json();
            if (result.error) {
                taskStatus.textContent = '错误: ' + result.error;
                submitBtn.disabled = false;
                return;
            }
            const taskId = result.task_id;
            taskStatus.textContent = '任务已提交，正在处理...';
            progressBar.style.width = '10%';

            const poll = setInterval(async () => {
                try {
                    const statusResp = await fetch('/team-report/status/' + taskId);
                    if (!statusResp.ok) {
                        if (statusResp.status === 524 || statusResp.status === 502) {
                            console.warn('收到 ' + statusResp.status + '，继续等待...');
                            return;
                        }
                        clearInterval(poll);
                        taskStatus.textContent = '查询状态失败 (HTTP ' + statusResp.status + ')';
                        submitBtn.disabled = false;
                        return;
                    }
                    const status = await statusResp.json();
                    if (status.error) {
                        clearInterval(poll);
                        taskStatus.textContent = '错误: ' + status.error;
                        submitBtn.disabled = false;
                        return;
                    }
                    if (status.progress_percent !== undefined) {
                        progressBar.style.width = status.progress_percent + '%';
                    }
                    if (status.progress_message) {
                        progressMessage.textContent = status.progress_message;
                    }
                    if (status.status === 'completed') {
                        clearInterval(poll);
                        taskStatus.textContent = '处理完成，正在下载...';
                        progressBar.style.width = '100%';
                        window.location.href = '/team-report/download/' + taskId;
                        setTimeout(() => { submitBtn.disabled = false; }, 2000);
                    } else if (status.status === 'failed') {
                        clearInterval(poll);
                        taskStatus.textContent = '处理失败: ' + (status.error || '未知错误');
                        submitBtn.disabled = false;
                    }
                } catch (err) {
                    console.error('轮询错误:', err);
                }
            }, 4000);
        } catch (err) {
            taskStatus.textContent = '网络错误: ' + err.message;
            submitBtn.disabled = false;
        }
    });
})();
    ```

- **static/js/transposer.js**
  - language: `javascript`
  - content:
    ```javascript
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('transposerForm');
    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        const statusArea = document.getElementById('statusArea');
        const statusMsg = document.getElementById('statusMessage');
        const submitBtn = document.getElementById('submitBtn');

        statusArea.style.display = 'block';
        statusMsg.textContent = '处理中，请稍候...';
        submitBtn.disabled = true;

        const formData = new FormData(form);
        try {
            const response = await fetch('/transposer/convert', {
                method: 'POST',
                body: formData
            });
            if (!response.ok) {
                const errData = await response.json();
                statusMsg.textContent = '错误：' + (errData.error || '未知错误');
                submitBtn.disabled = false;
                return;
            }
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const filename = getFilenameFromContentDisposition(response.headers.get('Content-Disposition')) || '转置结果.xlsx';
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            statusMsg.textContent = '转换成功，文件已下载。';
        } catch (err) {
            statusMsg.textContent = '网络错误：' + err.message;
        } finally {
            submitBtn.disabled = false;
        }
    });
});
    ```

- **static/style.css**
  - language: `css`
  - content:
    ```css
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    margin: 0;
    padding: 20px;
    background: #f5f7fa;
}
.container {
    max-width: 1200px;
    margin: 0 auto;
    background: white;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}
nav ul {
    list-style: none;
    padding: 0;
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 20px;
    margin-bottom: 30px;
}
nav a {
    text-decoration: none;
    color: #007bff;
    font-weight: 500;
}
nav a:hover {
    text-decoration: underline;
}
.message {
    margin-top: 20px;
    padding: 12px;
    border-radius: 6px;
}
.message.error {
    background: #fee;
    color: #c00;
}
.message.success {
    background: #efe;
    color: #060;
}
button, .btn {
    background: #007bff;
    color: white;
    padding: 10px 20px;
    border: none;
    border-radius: 6px;
    cursor: pointer;
    font-size: 16px;
}
button:hover, .btn:hover {
    background: #0056b3;
}
input, select, textarea {
    padding: 8px;
    border: 1px solid #ddd;
    border-radius: 6px;
    width: 100%;
    max-width: 400px;
    margin-bottom: 15px;
}
label {
    font-weight: 600;
    display: block;
    margin-bottom: 5px;
}
    ```

- **templates/base.html**
  - language: `html`
  - content:
    ```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{% block title %}Excel 智能工具集{% endblock %}</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
    <script src="{{ url_for('static', filename='js/main.js') }}" defer></script>
    <script src="{{ url_for('static', filename='js/download_utils.js') }}"></script>
    {% block head %}{% endblock %}
</head>
<body>
    <div class="container">
        <h1>📊 Excel 智能工具集</h1>
        <nav>
            <ul>
                <li><a href="{{ url_for('index') }}">🏠 首页</a></li>
                <li><a href="{{ url_for('calc.index') }}">🧮 列计算/聚合</a></li>
                <li><a href="{{ url_for('compare.index') }}">🔍 文件对比</a></li>
                <li><a href="{{ url_for('match.index') }}">🔗 匹配填充</a></li>
                <li><a href="{{ url_for('summary.index') }}">📋 汇总工具</a></li>
                <li><a href="{{ url_for('splitter.index') }}">✂️ 文件拆分</a></li>
                <li><a href="{{ url_for('field_parser.index') }}">🔗 字段映射</a></li>
                <li><a href="{{ url_for('checker.index') }}">✅ 文件校验</a></li>
                <li><a href="{{ url_for('convert.index') }}">🔄 格式转换</a></li>
                <li><a href="{{ url_for('subtotal.index') }}">📊 数据小计</a></li>
                <li><a href="{{ url_for('team_report.index') }}">📊 车队日报</a></li>
                <li><a href="{{ url_for('transposer.index') }}">🔄 行列转置</a></li>
            </ul>
        </nav>
        {% block content %}{% endblock %}
        <div id="message" class="message"></div>
    </div>
    {% block scripts %}{% endblock %}
</body>
</html>
    ```

- **templates/calculator/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}列计算 / 聚合回填{% endblock %}
{% block content %}
<h2>🧮 列计算 / 聚合回填</h2>
<form action="{{ url_for('calc.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
        <div class="form-row">
            <label>⚙️ 操作类型</label>
            <select name="calc_type" id="calcType">
                <option value="column">列计算</option>
                <option value="aggregate">聚合回填</option>
            </select>
        </div>
        <div class="form-row">
            <label>表头所在行号（从1开始，默认1）</label>
            <input type="number" name="header_row" value="1" min="1" style="max-width:100px;">
            <div class="hint">如果你的 Excel 顶部有标题行，请填写实际列名所在的行号</div>
        </div>
    </div>
    <div id="columnConfig" class="form-section">
        <h3>📐 列计算设置</h3>
        <div class="form-row">
            <label>表达式 (如 A+B-C)</label>
            <input type="text" name="expression" placeholder="例如 A+B">
        </div>
        <div class="form-row">
            <label>结果列 (字母或新列名)</label>
            <input type="text" name="result_col_column" placeholder="K 或 '差值'">
        </div>
        <div class="form-row">
            <label>合并依据列 (可选，逗号分隔)</label>
            <input type="text" name="merge_by_cols" placeholder="A,B">
            <div class="hint">相同组的结果列只保留第一行，其余合并。</div>
        </div>
    </div>
    <div id="aggregateConfig" class="form-section" style="display:none;">
        <h3>📊 聚合回填设置</h3>
        <div class="form-row">
            <label>分组列 (逗号分隔，支持列名或列字母)</label>
            <input type="text" name="group_cols" placeholder="A,B">
        </div>
        <div class="form-row">
            <label>聚合列 (逗号分隔)</label>
            <input type="text" name="agg_cols" placeholder="C,D">
        </div>
        <div class="form-row">
            <label>聚合函数</label>
            <select name="agg_func">
                <option value="sum">求和</option>
                <option value="mean">平均值</option>
                <option value="max">最大值</option>
                <option value="min">最小值</option>
                <option value="count">计数</option>
            </select>
        </div>
        <div class="form-row">
            <label>结果列名</label>
            <input type="text" name="result_col_aggregate" placeholder="K 或 '聚合结果'">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始处理</button>
</form>
<script>
    const typeSelect = document.getElementById('calcType');
    const colCfg = document.getElementById('columnConfig');
    const aggCfg = document.getElementById('aggregateConfig');
    function toggle() {
        if (typeSelect.value === 'column') {
            colCfg.style.display = 'block';
            aggCfg.style.display = 'none';
        } else {
            colCfg.style.display = 'none';
            aggCfg.style.display = 'block';
        }
    }
    typeSelect.addEventListener('change', toggle);
    toggle();
</script>
{% endblock %}
    ```

- **templates/checker/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}文件校验{% endblock %}
{% block head %}
<link rel="stylesheet" href="{{ url_for('static', filename='css/checker.css') }}">
{% endblock %}
{% block content %}
<h2>✅ 文件校验</h2>
<p>上传标准文件群和比较文件群，系统将根据您指定的主键和对比列进行比较，并生成差异报告。</p>

<form id="checkerForm" enctype="multipart/form-data">
    <div class="form-section">
        <h3>📂 文件上传</h3>
        <div class="form-row">
            <label>标准文件群（可多选）</label>
            <input type="file" id="filesA" name="files_a" accept=".xlsx,.xls,.csv" multiple required>
        </div>
        <div class="form-row">
            <label>比较文件群（可多选）</label>
            <input type="file" id="filesB" name="files_b" accept=".xlsx,.xls,.csv" multiple required>
        </div>
    </div>

    <div class="form-section">
        <h3>🔑 主键设置</h3>
        <div class="form-row">
            <label>标准群主键列（逗号或空格分隔）</label>
            <input type="text" id="keyMapA" placeholder="例如 ID,日期">
        </div>
        <div class="form-row">
            <label>比较群主键列（逗号或空格分隔，默认与标准相同）</label>
            <input type="text" id="keyMapB" placeholder="留空则与标准群主键相同">
        </div>
        <div class="form-row">
            <label>主键重复处理</label>
            <select id="duplicateAction">
                <option value="error" selected>报错</option>
                <option value="keep_first">保留第一条</option>
            </select>
        </div>
    </div>

    <div class="form-section">
        <h3>📊 对比列</h3>
        <div class="parser-container" style="display: flex; gap: 20px;">
            <div class="parser-box" style="flex: 1;">
                <label>标准群对比列（每行一个或用逗号/空格分隔）</label>
                <textarea id="compareStandardText" rows="5" placeholder="例如：医疗类别&#10;入院日期&#10;出院日期&#10;医疗费总额&#10;结算时间&#10;定点医药机构名称"></textarea>
            </div>
            <div class="parser-box" style="flex: 1;">
                <label>比较群对比列（留空则与标准相同）</label>
                <textarea id="compareCompareText" rows="5" placeholder="例如：医疗类别&#10;入院日期&#10;出院日期&#10;医疗费总额&#10;结算时间&#10;定点医疗机构名称"></textarea>
            </div>
        </div>
        <div class="hint">按顺序一一对应，支持多列；比较群留空则默认与标准群列名相同。</div>
    </div>

    <div class="form-section">
        <h3>🔍 标准群筛选（可选）</h3>
        <div id="filtersContainer"></div>
        <button type="button" class="btn btn-secondary" onclick="addFilter()">+ 添加筛选条件</button>
        <div class="hint">支持多值OR：在“值”中用英文逗号分隔多个值</div>
    </div>

    <div style="text-align:center; margin: 30px 0;">
        <button type="button" class="btn" id="btnPrecheck">🔎 预检</button>
        <button type="button" class="btn btn-success" id="btnStart" disabled>▶ 开始校验</button>
    </div>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar" style="width:0%"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <p id="progressMessage" style="color:#666;"></p>
    <div id="precheckResult" class="precheck-result"></div>
    <div id="downloadLink" style="margin-top:15px;"></div>
</div>

<template id="filterRowTemplate">
    <div class="filter-row">
        <input type="text" name="filter_field" placeholder="字段名" class="filter-field">
        <select name="filter_op" class="filter-op">
            <option value="=">=</option><option value="!=">!=</option><option value=">">></option>
            <option value="<"><</option><option value=">=">>=</option><option value="<="><=</option>
            <option value="contains">包含</option><option value="not_contains">不包含</option>
        </select>
        <input type="text" name="filter_value" placeholder="值（多值用英文逗号）" class="filter-value">
        <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">删除</button>
    </div>
</template>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/checker.js') }}"></script>
{% endblock %}
    ```

