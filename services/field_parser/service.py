# services/field_parser/service.py
import re
import logging
from typing import List, Dict
from difflib import SequenceMatcher

logger = logging.getLogger(__name__)

COMMON_REPLACEMENTS = {
    'id': '标识',
    'tel': '电话',
    'phone': '电话',
    'mobile': '手机',
    'name': '姓名',
    'addr': '地址',
    'no': '编号',
    'num': '数量',
}

class FieldParserService:
    """字段解析与分层匹配服务"""

    def __init__(self, similarity_threshold: int = 30):
        self.threshold = similarity_threshold / 100.0

    def parse_field_list(self, text: str) -> List[str]:
        parts = re.split(r'[,，\s\n\r\t]+', text.strip())
        return [p.strip() for p in parts if p.strip()]

    def _normalize(self, s: str) -> str:
        s = s.strip().lower()
        s = re.sub(r'[\s\-_]+', '', s)
        for abbr, full in COMMON_REPLACEMENTS.items():
            if s == abbr:
                return full
        return s

    def match_headers(self, headers_a: List[str], headers_b: List[str]) -> Dict:
        norm_a = [self._normalize(h) for h in headers_a]
        norm_b = [self._normalize(h) for h in headers_b]

        # ---------- 第一步：完全匹配（标准化后相等） ----------
        exact_matched = []
        remaining_a = []
        remaining_b_idx = set(range(len(norm_b)))  # 还未匹配的 B 列索引

        for i, ha in enumerate(headers_a):
            na = norm_a[i]
            matched = False
            for j in list(remaining_b_idx):
                if na == norm_b[j]:
                    exact_matched.append({
                        "standard": ha,
                        "compare": headers_b[j],
                        "score": 1.0
                    })
                    remaining_b_idx.remove(j)
                    matched = True
                    break
            if not matched:
                remaining_a.append(i)

        # ---------- 第二步：相似度匹配（按分数从高到低） ----------
        # 收集所有可能配对的相似度
        candidates = []
        for i in remaining_a:
            for j in remaining_b_idx:
                score = self._similarity_score(norm_a[i], norm_b[j])
                if score >= self.threshold:
                    candidates.append((score, i, j))

        # 按分数降序
        candidates.sort(key=lambda x: x[0], reverse=True)

        used_a = set()
        used_b = set()
        fuzzy_matched = []

        for score, i, j in candidates:
            if i in used_a or j in used_b:
                continue
            fuzzy_matched.append({
                "standard": headers_a[i],
                "compare": headers_b[j],
                "score": round(score, 2)
            })
            used_a.add(i)
            used_b.add(j)

        # ---------- 汇总结果 ----------
        all_matched = exact_matched + fuzzy_matched
        matched_a_idx = set(range(len(headers_a))) - {i for i in remaining_a} | used_a
        matched_b_idx = set(range(len(headers_b))) - remaining_b_idx | used_b

        only_a = [headers_a[i] for i in range(len(headers_a)) if i not in matched_a_idx]
        only_b = [headers_b[j] for j in range(len(headers_b)) if j not in matched_b_idx]

        return {"matched": all_matched, "only_a": only_a, "only_b": only_b}

    def _similarity_score(self, a: str, b: str) -> float:
        if a == b:
            return 1.0
        if not a or not b:
            return 0.0

        # SequenceMatcher
        seq_sim = SequenceMatcher(None, a, b).ratio()

        # 归一化编辑距离
        len_a, len_b = len(a), len(b)
        dp = [[0]*(len_b+1) for _ in range(len_a+1)]
        for i in range(len_a+1):
            dp[i][0] = i
        for j in range(len_b+1):
            dp[0][j] = j
        for i in range(1, len_a+1):
            for j in range(1, len_b+1):
                cost = 0 if a[i-1] == b[j-1] else 1
                dp[i][j] = min(dp[i-1][j]+1, dp[i][j-1]+1, dp[i-1][j-1]+cost)
        edit_sim = 1.0 - (dp[len_a][len_b] / max(len_a, len_b))

        # Jaccard
        set_a, set_b = set(a), set(b)
        common = len(set_a & set_b)
        total = len(set_a | set_b)
        jaccard = common / total if total > 0 else 0.0

        return 0.4 * seq_sim + 0.3 * edit_sim + 0.3 * jaccard
