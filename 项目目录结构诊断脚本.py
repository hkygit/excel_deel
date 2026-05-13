#!/usr/bin/env python3
"""
项目目录结构诊断脚本
运行方式：python inspect_project.py
"""

import os
import sys
from pathlib import Path


def print_tree(directory, prefix="", max_depth=4,
               exclude_dirs={'.venv', '__pycache__', '.git', 'node_modules', '.idea', '.vscode', 'outputs', 'uploads',
                             'task_status'}):
    """打印目录树，排除某些文件夹"""
    try:
        items = sorted(os.listdir(directory))
    except PermissionError:
        return
    paths = [item for item in items if not item.startswith('.') or item == '__init__.py']
    for i, item in enumerate(paths):
        path = os.path.join(directory, item)
        is_last = (i == len(paths) - 1)
        current_prefix = "└── " if is_last else "├── "
        print(prefix + current_prefix + item)
        if os.path.isdir(path) and item not in exclude_dirs and max_depth > 0:
            next_prefix = prefix + ("    " if is_last else "│   ")
            print_tree(path, next_prefix, max_depth - 1, exclude_dirs)


def main():
    print("=" * 60)
    print("项目目录结构诊断报告")
    print("=" * 60)

    # 当前工作目录
    cwd = os.getcwd()
    print(f"\n1. 当前工作目录: {cwd}")

    # 脚本所在目录（通常是项目根）
    script_dir = Path(__file__).parent.absolute()
    print(f"\n2. 脚本所在目录: {script_dir}")

    # Python 路径
    print("\n3. sys.path:")
    for i, p in enumerate(sys.path):
        print(f"   {i}: {p}")

    # 检查关键文件/目录是否存在
    print("\n4. 关键路径检查:")
    critical_paths = [
        "app.py",
        "config.py",
        "blueprints",
        "utils",
        "templates",
        "static",
        "requirements.txt",
        "blueprints/simplify.py",
        "utils/simplify_helpers.py",
        "utils/simplify_large.py",
        "utils/simplify_async.py",
        "utils/format_utils.py",
        "utils/file_utils.py",
        "utils/__init__.py",
        "blueprints/__init__.py",
    ]
    for rel_path in critical_paths:
        full = script_dir / rel_path
        exists = full.exists()
        is_dir = full.is_dir() if exists else False
        status = "✓ 目录" if exists and is_dir else ("✓ 文件" if exists else "✗ 缺失")
        print(f"   {rel_path:30} {status}")

    # 项目根目录的完整树结构（限制深度）
    print("\n5. 项目目录树（深度3，排除 .venv 等）:")
    print_tree(str(script_dir), max_depth=3)

    # 检查 utils 包的导入路径问题
    print("\n6. 导入路径测试（模拟）:")
    utils_init = script_dir / "utils" / "__init__.py"
    if utils_init.exists():
        with open(utils_init, 'r', encoding='utf-8') as f:
            content = f.read()
        print("   utils/__init__.py 内容:")
        for line in content.splitlines()[:10]:
            print(f"      {line}")
    else:
        print("   ✗ utils/__init__.py 不存在")

    # 输出建议
    print("\n" + "=" * 60)
    print("请将以上完整输出复制反馈给我，我将根据实际结构修正导入错误。")
    print("=" * 60)


if __name__ == "__main__":
    main()