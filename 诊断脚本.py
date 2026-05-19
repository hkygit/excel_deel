#!/usr/bin/env python3
"""
诊断“匹配填充”操作：检查目标文件(双通道.xlsx)与源文件(居民财务支付.xlsx)的匹配情况。
前端参数：目标键列 A,B；源匹配列 A,B；源值列 C；结果列 M；允许重复键。
"""
import os
import sys
from datetime import datetime

# 确保项目路径可导入
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from services.excel.matcher import ExcelMatcher

# 目标文件和源文件路径
root_dir = os.path.dirname(os.path.abspath(__file__))
target_file = os.path.join(root_dir, '最后是双通道.xlsx')
source_file = os.path.join(root_dir, '居民财务支付.xlsx')

# 参数
config = {
    'target_key_cols': 'A,B',
    'source_match_cols': 'A,B',
    'source_value_col': 'C',
    'result_col': 'M',
    'allow_duplicate_keys': True,   # 勾选
}

# 日志输出目录
log_dir = os.path.join(root_dir, 'log')
os.makedirs(log_dir, exist_ok=True)
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
log_file = os.path.join(log_dir, f'diagnose_matcher_{timestamp}.log')

def log(msg):
    print(msg)
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(msg + '\n')

def main():
    log(f'诊断开始: {datetime.now().isoformat()}')
    log(f'目标文件: {target_file}')
    log(f'源文件:   {source_file}')
    log(f'参数: {config}\n')

    # 1. 文件存在性检查
    if not os.path.exists(target_file):
        log(f'❌ 目标文件不存在: {target_file}')
        return
    if not os.path.exists(source_file):
        log(f'❌ 源文件不存在: {source_file}')
        return
    log('✅ 文件存在检查通过\n')

    # 2. 创建匹配器实例
    matcher = ExcelMatcher(target_file, source_file, config)

    # 3. 列字母校验
    log('=== 列字母校验 ===')
    if not matcher._validate_column_letters():
        for err in matcher.errors:
            log(f'❌ {err}')
        return
    log('✅ 列字母解析成功')
    log(f'目标键列索引: {matcher.target_key_cols}')
    log(f'源匹配列索引: {matcher.source_match_cols}')
    log(f'源值列索引:   {matcher.source_value_col}')
    log(f'结果列:       {matcher.result_col} ({"新列" if matcher.result_is_new else "已有列"})')
    log('')

    # 4. 加载源文件
    log('=== 加载源文件 ===')
    if not matcher._load_source():
        for err in matcher.errors:
            log(f'❌ {err}')
        return
    log(f'✅ 源文件读取成功，总行数: {len(matcher.source_df)}')
    log(f'源文件列名: {list(matcher.source_df.columns)}')
    # 匹配列和值列的实际名称
    try:
        match_cols = [matcher.source_col_names[i-1] for i in matcher.source_match_cols]
        val_col = matcher.source_col_names[matcher.source_value_col-1]
        log(f'匹配列: {match_cols}，值列: {val_col}')
    except Exception as e:
        log(f'⚠️ 列索引获取异常: {e}')

    # 构建查找字典，显示部分样例
    lookup, match_len = matcher._build_lookup_dict()
    log(f'源数据去重后唯一键数量: {len(lookup)}')
    sample_keys = list(lookup.keys())[:5]
    log(f'源数据样例键值对:')
    for k in sample_keys:
        log(f'  {k} -> {lookup[k]}')
    log('')

    # 5. 加载目标文件
    log('=== 加载目标文件 ===')
    if not matcher._load_target():
        for err in matcher.errors:
            log(f'❌ {err}')
        return
    log('✅ 目标文件读取成功')
    # 提取目标键
    target_keys, start_row, end_row = matcher._extract_target_keys()
    log(f'目标数据行范围: 第 {start_row} 行 至 第 {end_row} 行')
    log(f'目标数据总行数: {len(target_keys)}')
    # 显示前5个键
    log('目标键示例（前5个）:')
    for i, k in enumerate(target_keys[:5]):
        log(f'  行{start_row + i}: {k}')
    # 检查重复键
    from collections import Counter
    key_counts = Counter(target_keys)
    dup_keys = {k: v for k, v in key_counts.items() if v > 1}
    if dup_keys:
        log(f'⚠️ 目标文件中存在重复键（共 {len(dup_keys)} 组）:')
        for k, v in list(dup_keys.items())[:10]:
            log(f'  键 {k} 出现 {v} 次')
    else:
        log('✅ 目标键无重复')
    log('')

    # 6. 匹配情况分析（不使用 allow_duplicate_keys 的写入逻辑，仅诊断匹配）
    log('=== 匹配结果分析 ===')
    # 构建目标键的截短版本（match_len）
    used_keys = [key[:match_len] for key in target_keys]
    source_key_set = set(lookup.keys())
    target_key_set = set(used_keys)
    matched_in_target = target_key_set & source_key_set
    only_in_target = target_key_set - source_key_set
    only_in_source = source_key_set - target_key_set

    log(f'目标文件中出现的唯一键数量: {len(target_key_set)}')
    log(f'源文件中出现的唯一键数量: {len(source_key_set)}')
    log(f'成功匹配的键数量: {len(matched_in_target)}')
    log(f'仅在目标中出现的键数量: {len(only_in_target)}')
    log(f'仅在源中出现的键数量: {len(only_in_source)}')

    if only_in_target:
        log('仅在目标中的键示例（前10个）:')
        for k in list(only_in_target)[:10]:
            log(f'  {k}')
    if only_in_source:
        log('仅在源中的键示例（前10个）:')
        for k in list(only_in_source)[:10]:
            log(f'  {k}')

    # 7. 模拟填充结果（仅首行填充）
    log('\n=== 模拟填充结果（allow_duplicate_keys=True） ===')
    seen = set()
    filled_rows = []
    skipped_rows = []
    for i, key in enumerate(used_keys):
        row_num = start_row + i
        if key in seen:
            skipped_rows.append((row_num, key))
            continue
        seen.add(key)
        value = lookup.get(key, 0)
        filled_rows.append((row_num, key, value))
    log(f'实际填充的行数: {len(filled_rows)}')
    log(f'因重复键跳过的行数: {len(skipped_rows)}')
    if filled_rows:
        log('填充行示例（前10行）:')
        for row_num, key, val in filled_rows[:10]:
            log(f'  行{row_num}: 键{key} -> 值 {val}')
    if skipped_rows:
        log('跳过的重复行示例（前10行）:')
        for row_num, key in skipped_rows[:10]:
            log(f'  行{row_num}: 键{key}（已在前面的行中填充过）')

    log(f'\n✅ 诊断完成，报告已保存至: {log_file}')

if __name__ == '__main__':
    main()