#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
彻底移除 utils 目录，将所有 from utils.xxx 导入迁移到 core 对应模块。
"""

import os
import shutil
import datetime

# 映射关系：utils 模块名 -> core 模块名
MAPPING = {
    "validators": "core.validators",
    "excel_helpers": "core.excel_utils",
    "condition_filter": "core.filter_engine",
    "exceptions": "core.exceptions",
}

def replace_import_in_file(filepath):
    """替换文件中的导入语句"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    original = content
    for utils_mod, core_mod in MAPPING.items():
        content = content.replace(f"from utils.{utils_mod} import", f"from {core_mod} import")
        content = content.replace(f"import utils.{utils_mod}", f"import {core_mod}")
    content = content.replace("from core import", "from core import")

    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

def walk_and_replace(root_dir):
    count = 0
    ignore_dirs = {".git", ".venv", "venv", "__pycache__", ".trash", "log", "outputs", "uploads", "task_status"}
    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames[:] = [d for d in dirnames if d not in ignore_dirs]
        for filename in filenames:
            if filename.endswith(".py"):
                full_path = os.path.join(dirpath, filename)
                if replace_import_in_file(full_path):
                    print(f"已更新: {full_path}")
                    count += 1
    return count

def remove_utils_directory(root_dir):
    utils_path = os.path.join(root_dir, "utils")
    if not os.path.exists(utils_path):
        print("utils 目录不存在，跳过删除")
        return
    trash_dir = os.path.join(root_dir, ".trash")
    os.makedirs(trash_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(trash_dir, f"utils_{timestamp}")
    shutil.move(utils_path, dest)
    print(f"utils 目录已移动到 {dest}")

def main():
    root = os.getcwd()
    print("开始迁移 utils 导入到 core...")
    updated = walk_and_replace(root)
    print(f"共更新了 {updated} 个文件")
    remove_utils_directory(root)
    print("完成。请运行测试确保一切正常。")

if __name__ == "__main__":
    main()