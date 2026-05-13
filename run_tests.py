#!/usr/bin/env python3
"""
测试运行脚本
- 自动执行 pytest 运行所有测试
python -m pytest tests/ -v
- 若全部通过：仅打印成功信息，不生成任何日志文件
- 若有失败：将完整输出保存到 log/ 目录，并包含失败测试列表
"""

import subprocess
import sys
import re
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
LOG_DIR = BASE_DIR / "log"


def run_tests():
    """运行 pytest 并返回 (returncode, output)"""
    cmd = [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"]
    print(f"🚀 运行测试: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    output = result.stdout + result.stderr
    return result.returncode, output


def extract_failed_tests(output: str) -> list:
    """从 pytest 输出中提取失败的测试名称"""
    failed_tests = []
    # 匹配类似 "tests/test_xxx.py::TestClass::test_method FAILED" 的行
    pattern = re.compile(r"^(.+?\.py::[^ ]+)\s+FAILED", re.MULTILINE)
    for match in pattern.finditer(output):
        failed_tests.append(match.group(1))
    return failed_tests


def save_failure_log(output: str, failed_tests: list):
    """保存失败日志到 log/ 目录，包含失败测试列表"""
    LOG_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"test_failures_{timestamp}.log"

    # 构建日志内容
    header = f"测试运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    header += f"失败测试数量: {len(failed_tests)}\n"
    header += "-" * 60 + "\n"
    if failed_tests:
        header += "失败测试列表:\n"
        for test in failed_tests:
            header += f"  - {test}\n"
        header += "-" * 60 + "\n\n"

    full_log = header + output

    log_file.write_text(full_log, encoding="utf-8")
    print(f"\n📁 失败日志已保存至: {log_file}")
    print("\n⚠️ 部分测试失败，失败测试列表：")
    for test in failed_tests:
        print(f"   ❌ {test}")
    print(f"\n详细信息请查看日志文件。")


def main():
    returncode, output = run_tests()

    if returncode == 0:
        print("\n✅ 所有测试通过！未生成日志文件。")
    else:
        failed_tests = extract_failed_tests(output)
        save_failure_log(output, failed_tests)
        sys.exit(returncode)


if __name__ == "__main__":
    main()