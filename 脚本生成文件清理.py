#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理项目生成文件、缓存和临时数据
- 删除 outputs/, uploads/, task_status/, logs/, __pycache__/ 等目录
- 删除根目录下匹配模式的文件：team_report_*.xlsx, *.tmp, *.bak, *.log, *.db, *.pyc 等
- 虚拟环境目录（.venv/, venv/, env/）不会删除
"""

import os
import shutil
import fnmatch
import sys

# ---------- 需要清理的目录（相对于项目根目录），不会清除虚拟环境 ----------
DIRS_TO_CLEAN = [
    "outputs",
    "uploads",
    "task_status",
    "logs",
    "__pycache__",
    ".pytest_cache",
    "htmlcov",
    ".git.backup",
    "task_status_test",
]

# 需要清理的根目录文件模式（使用 fnmatch）
FILE_PATTERNS = [
    "*.tmp", "*.bak", "*.log", "*.cache",
    "*.db", "*.sqlite", "*.pyc", "*.pyo",
    "coverage.xml", ".coverage",
    "team_report_*.xlsx",          # 车队日报生成文件
    "comparison_result.xlsx",      # 对比结果
    "匹配结果.xlsx",
    "filtered_*.xlsx",
    "汇总结果.xlsx",
    "字段映射结果.xlsx",
    "筛选结果.xlsx",
    "*.zip",                       # 可能生成的临时压缩包
]

def confirm(prompt):
    """获取用户确认"""
    answer = input(prompt).strip().lower()
    return answer in ('y', 'yes')

def remove_dir(path):
    """递归删除目录（如果存在）"""
    if os.path.isdir(path):
        try:
            shutil.rmtree(path)
            print(f"  ✓ 已删除目录: {path}")
        except Exception as e:
            print(f"  ✗ 删除失败 {path}: {e}")

def remove_file(path):
    """删除文件（如果存在）"""
    if os.path.isfile(path):
        try:
            os.remove(path)
            print(f"  ✓ 已删除文件: {path}")
        except Exception as e:
            print(f"  ✗ 删除失败 {path}: {e}")

def list_matching_files(root_dir, patterns):
    """在根目录下查找匹配模式的文件（不递归）"""
    matches = []
    for pattern in patterns:
        for f in os.listdir(root_dir):
            if fnmatch.fnmatch(f, pattern):
                full = os.path.join(root_dir, f)
                if os.path.isfile(full):
                    matches.append(full)
    return matches

def main():
    print("=" * 60)
    print("项目清理工具 - 删除生成文件、缓存和临时数据")
    print("=" * 60)
    print("注意：此操作将删除 outputs/、uploads/、task_status/ 等目录")
    print("以及所有 team_report_*.xlsx 等报告文件。")
    print("源代码、.git 目录、虚拟环境（.venv 等）将**保留**。\n")

    if not confirm("是否继续？(y/N): "):
        print("已取消")
        sys.exit(0)

    root_dir = os.getcwd()
    items_to_delete = []  # 存储 (类型, 路径)

    # 收集要删除的目录
    for d in DIRS_TO_CLEAN:
        path = os.path.join(root_dir, d)
        if os.path.exists(path):
            items_to_delete.append(("目录", path))

    # 收集要删除的文件
    pattern_files = list_matching_files(root_dir, FILE_PATTERNS)
    for f in pattern_files:
        items_to_delete.append(("文件", f))

    if not items_to_delete:
        print("未发现需要清理的项目。")
        return

    print("\n即将删除以下内容：")
    for typ, path in items_to_delete:
        print(f"  [{typ}] {path}")

    if not confirm("\n确认删除以上所有文件/目录？(y/N): "):
        print("已取消")
        sys.exit(0)

    print("\n开始清理...")
    for typ, path in items_to_delete:
        if typ == "目录":
            remove_dir(path)
        else:
            remove_file(path)

    print("\n✅ 清理完成！")

if __name__ == "__main__":
    main()