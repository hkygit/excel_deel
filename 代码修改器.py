#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修复 test_basic_concat：pandas 读回时把 18 位身份证号推断成 int64。

将 pd.read_excel(out) 改为 pd.read_excel(out, dtype=str)，以验证字符串
精度保护。该修改幂等，重复运行不会重复修改。
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TEST_FILE = ROOT / "tests" / "test_concat_service.py"


OLD = '''        out = service.execute(p, "in.xlsx", "身份证号", ["疾病名称"])
        result = pd.read_excel(out)'''

NEW = '''        out = service.execute(p, "in.xlsx", "身份证号", ["疾病名称"])
        # 用 dtype=str 读回：避免 pandas 把 18 位身份证号推断成 int64；
        # 同时能真正验证精度——若服务内部丢了精度（如存成 float64），
        # dtype=str 读回也会得到 ...924 而非 ...923，测试依然会失败。
        result = pd.read_excel(out, dtype=str)'''


def main():
    if not TEST_FILE.exists():
        print(f"错误：未找到 {TEST_FILE}")
        sys.exit(1)

    text = TEST_FILE.read_text(encoding="utf-8")

    if NEW in text:
        print("已修复，无需修改")
        return

    if OLD not in text:
        print("错误：未找到预期代码片段，请人工检查 test_basic_concat")
        # 打印一下方便排查
        idx = text.find("def test_basic_concat")
        if idx >= 0:
            print("--- 当前 test_basic_concat 上下文 ---")
            print(text[idx:idx + 500])
        sys.exit(1)

    text = text.replace(OLD, NEW, 1)
    TEST_FILE.write_text(text, encoding="utf-8")
    print("完成：test_basic_concat 已改为 dtype=str 读回")


if __name__ == "__main__":
    main()