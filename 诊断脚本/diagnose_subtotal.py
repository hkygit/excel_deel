#!/usr/bin/env python3
"""
Subtotal 功能诊断脚本
用法: python diagnose_subtotal.py <文件名.xlsx> [主键列] [计算列] [聚合函数]
示例: python diagnose_subtotal.py 测试文件.xlsx A B sum
"""
import os
import sys
import shutil
import copy
import pandas as pd
from datetime import datetime

# 确保项目根目录在 sys.path 中
ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from services.subtotal.engine import SubtotalEngine
from core.file_naming import generate_output_filename

def main():
    # 1. 解析参数
    if len(sys.argv) < 2:
        print("错误: 请指定输入文件名。")
        print("用法: python diagnose_subtotal.py <文件名.xlsx> [主键列] [计算列] [聚合函数]")
        sys.exit(1)

    input_filename = sys.argv[1]
    input_path = os.path.join(ROOT, input_filename)
    if not os.path.exists(input_path):
        print(f"错误: 文件不存在 - {input_path}")
        sys.exit(1)

    key_cols = sys.argv[2] if len(sys.argv) >= 3 else 'A'
    calc_cols = sys.argv[3] if len(sys.argv) >= 4 else 'B'
    agg_func_default = sys.argv[4] if len(sys.argv) >= 5 else 'sum'

    # 2. 创建诊断文件夹
    diag_dir = os.path.join(ROOT, '../diagnosis')
    os.makedirs(diag_dir, exist_ok=True)

    # 3. 复制输入文件到诊断文件夹（避免污染原文件）
    diag_input_path = os.path.join(diag_dir, input_filename)
    shutil.copy2(input_path, diag_input_path)

    # 4. 初始化引擎并执行
    print("=" * 60)
    print("Subtotal 诊断")
    print("=" * 60)
    print(f"输入文件: {input_filename}")
    print(f"主键列:   {key_cols}")
    print(f"计算列:   {calc_cols}")
    print(f"聚合函数: {agg_func_default}")
    print("-" * 60)

    engine = SubtotalEngine(
        filepath=diag_input_path,
        key_cols_str=key_cols,
        calc_cols_str=calc_cols,
        agg_funcs={'default': agg_func_default},
        dimension_cols=None,          # 不设置维度列
        sheets='1',                   # 仅第一个 sheet
        original_filename=input_filename  # 传入原始文件名，影响命名
    )

    try:
        output_path = engine.process()
    except Exception as e:
        print(f"处理失败: {e}")
        sys.exit(1)

    # 5. 分析输出
    expected_name = generate_output_filename('subtotal', input_filename, suffix='_小计')
    actual_name = os.path.basename(output_path)

    print("处理完成。")
    print(f"预期输出文件名: {expected_name}")
    print(f"实际输出文件名: {actual_name}")
    print(f"实际输出路径:   {output_path}")

    # 6. 验证
    if actual_name == expected_name:
        print("\n✅ 文件名符合统一命名规范！")
    else:
        print("\n❌ 文件名不符合预期。")
        print("可能原因:")
        print("  1. SubtotalEngine.process() 中未正确调用 generate_output_filename")
        print("  2. 传入的 original_filename 参数有误")
        print("  3. generate_output_filename 函数实现与预期不一致")

    # 7. 输出文件内容摘要（前几行）
    try:
        df = pd.read_excel(output_path, sheet_name=0)
        print(f"\n输出文件包含 {len(df)} 行数据，列名: {list(df.columns)}")
        print("前5行预览:")
        print(df.head(5).to_string(index=False))
    except Exception:
        print("\n无法读取输出文件进行预览。")

    # 8. 保存诊断报告
    report_path = os.path.join(diag_dir, 'diagnostic_report.txt')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(f"诊断时间: {datetime.now().isoformat()}\n")
        f.write(f"输入文件: {input_filename}\n")
        f.write(f"参数: key_cols={key_cols}, calc_cols={calc_cols}, agg={agg_func_default}\n")
        f.write(f"预期输出文件名: {expected_name}\n")
        f.write(f"实际输出文件名: {actual_name}\n")
        f.write(f"命名符合预期: {actual_name == expected_name}\n")
    print(f"\n诊断报告已保存至: {report_path}")

if __name__ == '__main__':
    main()