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
