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
