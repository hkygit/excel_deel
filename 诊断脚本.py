#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
完全模拟 calc 蓝图列计算处理流程（包括列字母映射）
输入文件：诊断文件/医保基金确认核对_匹配结果.xlsx
表达式：B - E
结果列：H
合并依据列：A
"""

import os
import sys
import gc
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment

# 切换到脚本所在目录
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.getcwd())

from core.expression_engine import get_expression_engine, extract_column_letters
from core.excel_utils import ExcelFile, copy_merge_structure
from core.header_detection import detect_header_start
from core.csv_utils import build_column_mapping
from core.data_access_layer import get_data_access_layer

# ========== 配置 ==========
INPUT_FILE = os.path.join("诊断文件", "医保基金确认核对_匹配结果.xlsx")
OUTPUT_FILE = os.path.join("诊断文件", "医保基金确认核对_calc完整模拟.xlsx")
EXPRESSION = "B-E"
RESULT_COL_INPUT = "H"
MERGE_BY_COLS_STR = "A"
HEADER_ROW = None   # 自动检测

# ========== 辅助函数 ==========
def auto_clean_dataframe(df):
    """复制自 DataAccessLayer._auto_clean_dataframe"""
    df = df.copy()
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(r'^\s*$', np.nan, regex=True)
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        if numeric_series.notna().any():
            df[col] = numeric_series
    return df

# ========== 主流程 ==========
def main():
    print("=" * 60)
    print("完全模拟 calc 蓝图列计算处理流程（含列字母映射）")
    print("=" * 60)
    print(f"当前工作目录: {os.getcwd()}")
    print(f"输入文件: {INPUT_FILE}")
    print(f"文件存在: {os.path.exists(INPUT_FILE)}")
    print(f"表达式: {EXPRESSION}")
    print(f"结果列: {RESULT_COL_INPUT}")
    print(f"合并依据列: {MERGE_BY_COLS_STR}")

    if not os.path.exists(INPUT_FILE):
        print("错误: 文件不存在，请检查路径。")
        return

    # 1. 确定表头行号
    if HEADER_ROW is None:
        try:
            data_start = detect_header_start(INPUT_FILE)
            header_row = data_start - 1
            if header_row < 1:
                header_row = 1
            print(f"自动检测: 数据开始行 {data_start}, 表头行号 = {header_row}")
        except Exception as e:
            print(f"表头检测失败: {e}，使用默认表头行 1")
            header_row = 1
    else:
        header_row = HEADER_ROW
        print(f"使用指定表头行: {header_row}")

    # 2. 读取 Excel (pandas)，跳过前 header_row-1 行，第 header_row 行作为列名
    try:
        df = pd.read_excel(INPUT_FILE, header=header_row - 1, dtype=str)
        print(f"读取成功，形状: {df.shape}")
        headers = list(df.columns)
        print(f"原始列名: {headers}")
    except Exception as e:
        print(f"pandas 读取失败: {e}")
        return

    # 3. 构建列字母到列名的映射（关键步骤！）
    letter_to_name, name_to_letter = build_column_mapping(headers)
    print(f"列字母映射 (前6个): {dict(list(letter_to_name.items())[:6])}")

    # 4. 将表达式中的列字母转换为列名
    variables = extract_column_letters(EXPRESSION)
    print(f"提取的列字母: {variables}")
    mapped_vars = {}
    for letter in variables:
        if letter in letter_to_name:
            mapped_vars[letter] = letter_to_name[letter]
        else:
            print(f"错误: 列字母 {letter} 无法映射到任何列名")
            return
    print(f"映射后列名: {mapped_vars}")

    # 5. 重命名 DataFrame 列，使表达式可以直接使用原始列名（可选，但为了计算方便，我们直接用映射后的列名计算）
    # 这里我们创建一个新的表达式，将字母替换为实际列名（但表达式引擎不支持中文列名中的特殊字符，所以最好直接使用字母）。
    # 实际上 calc 蓝图的计算引擎直接使用字母从行中取值，但它依赖的是列名本身就是字母？不，它通过 context 传入字母对应的值。
    # 在 ColumnCalculatorProcessor 中，它使用变量直接作为 DataFrame 的列名，但前提是列名已经是字母。
    # 但我们的列名是中文，所以我们必须将表达式中的字母替换为实际列名，才能正确地调用 engine.evaluate。
    # 更简单的方法：创建一个临时的 DataFrame，列名为字母 A,B,C...，将原数据按位置映射过去。
    # 我们在这里采用最符合实际蓝图逻辑的方式：使用原始列名，但构建 context 时用字母映射到对应列的值。
    engine = get_expression_engine()

    # 计算每一行
    def calc_row(row):
        # 构建 context: { 'B': row['城乡居民应支付'], 'E': row['城镇职工应支付'] }
        context = {}
        for letter, col_name in mapped_vars.items():
            context[letter] = row[col_name]
        return engine.evaluate(EXPRESSION, context)

    df["Result"] = df.apply(calc_row, axis=1)
    print("表达式计算完成。")

    # 6. 打开原始 Excel (openpyxl)
    print("打开原始工作簿...")
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    ws = wb.active
    max_col = ws.max_column

    # 7. 处理结果列
    try:
        result_col_idx = column_index_from_string(RESULT_COL_INPUT)
        is_new_col = False
        result_col_name = RESULT_COL_INPUT
    except:
        is_new_col = True
        result_col_name = RESULT_COL_INPUT.strip()
        result_col_idx = max_col + 1
        ws.insert_cols(result_col_idx)
        ws.cell(header_row, result_col_idx).value = result_col_name
        print(f"新建结果列 {RESULT_COL_INPUT}，索引 {result_col_idx}")

    if not is_new_col:
        existing = ws.cell(header_row, result_col_idx).value
        if existing is None or str(existing).strip() == "":
            ws.cell(header_row, result_col_idx).value = result_col_name
            print(f"结果列已存在，表头为空，已设置列名。")
        else:
            print(f"结果列已存在，将覆盖数据。")

    # 8. 清除结果列原有合并
    to_remove = []
    for merged in ws.merged_cells.ranges:
        if merged.min_col <= result_col_idx <= merged.max_col:
            to_remove.append(merged)
    for merged in to_remove:
        ws.unmerge_cells(str(merged))
    print(f"已清除结果列中的 {len(to_remove)} 个合并区域。")

    # 9. 写入计算结果
    start_row = header_row + 1
    end_row = ws.max_row
    print(f"数据行范围: {start_row} 到 {end_row}")
    for i, val in enumerate(df["Result"]):
        row_num = start_row + i
        if isinstance(val, float) and val != val:  # NaN
            val = None
        ws.cell(row_num, result_col_idx).value = val
    print("计算结果已写入。")

    # 10. 处理合并依据列
    if MERGE_BY_COLS_STR:
        merge_cols = [c.strip() for c in MERGE_BY_COLS_STR.split(',') if c.strip()]
        if merge_cols:
            first_merge_col = merge_cols[0]
            try:
                merge_idx = column_index_from_string(first_merge_col)
                print(f"复制 {first_merge_col} 列 (索引 {merge_idx}) 的合并结构到结果列...")
                copy_merge_structure(ws, merge_idx, ws, result_col_idx, start_row, end_row)
            except Exception as e:
                print(f"复制合并结构失败: {e}")
                import traceback
                traceback.print_exc()
        else:
            print("合并依据列为空，跳过。")
    else:
        print("未提供合并依据列，跳过。")

    # 11. 设置结果列右对齐
    for row in range(start_row, end_row + 1):
        ws.cell(row, result_col_idx).alignment = Alignment(horizontal='right')
    print("已设置右对齐。")

    # 12. 保存结果
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    wb.save(OUTPUT_FILE)
    print(f"\n结果已保存: {OUTPUT_FILE}")

    # 13. 验证合并单元格
    print("\n验证结果列中的合并单元格:")
    found = 0
    for merged in ws.merged_cells.ranges:
        if merged.min_col <= result_col_idx <= merged.max_col:
            print(f"  {merged}")
            found += 1
            if found >= 10:
                break
    if found == 0:
        print("  未发现任何合并单元格。")

    wb.close()
    gc.collect()
    print("\n模拟完成。")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n发生未捕获异常:")
        import traceback
        traceback.print_exc()
    input("\n按 Enter 键退出...")