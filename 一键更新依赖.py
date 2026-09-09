#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键更新项目依赖（自动修正 pandas 和 numpy 版本）
"""

import subprocess
import sys
import os
import re

def run_cmd(cmd, capture=False):
    try:
        if capture:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='utf-8', errors='replace')
            return result.returncode, result.stdout.strip(), result.stderr.strip()
        else:
            result = subprocess.run(cmd, shell=True)
            return result.returncode, "", ""
    except Exception as e:
        return 1, "", str(e)

def fix_versions(req_file):
    """修正 requirements.txt 中的 pandas 和 numpy 版本"""
    with open(req_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    modified = False
    new_lines = []
    for line in lines:
        # 修正 pandas==2.0.3 -> pandas==2.3.3
        if re.match(r'^pandas\s*==\s*2\.0\.3\s*$', line.strip()):
            new_lines.append('pandas==2.3.3\n')
            modified = True
            print("已将 pandas 版本从 2.0.3 升级为 2.3.3（有预编译包）")
        # 修正 numpy==1.24.3 -> numpy==2.4.3
        elif re.match(r'^numpy\s*==\s*1\.24\.3\s*$', line.strip()):
            new_lines.append('numpy==2.4.3\n')
            modified = True
            print("已将 numpy 版本从 1.24.3 升级为 2.4.3（有预编译包）")
        else:
            new_lines.append(line)
    if modified:
        with open(req_file, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
    return modified

def main():
    print("="*60)
    print("一键更新项目依赖（自动修正 pandas 和 numpy）")
    print("="*60)

    req_file = "requirements.txt"
    if not os.path.isfile(req_file):
        print("错误：未找到 requirements.txt")
        sys.exit(1)

    ret, _, _ = run_cmd("pip --version", capture=True)
    if ret != 0:
        print("错误：pip 不可用")
        sys.exit(1)

    use_mirror = input("是否使用清华镜像加速下载？(y/n，默认 y): ").strip().lower()
    mirror = "-i https://pypi.tuna.tsinghua.edu.cn/simple" if (use_mirror == '' or use_mirror == 'y') else ""

    upgrade = input("是否升级所有包到最新符合版本约束的版本？(y/n，默认 n): ").strip().lower()
    upgrade_flag = "--upgrade" if upgrade == 'y' else ""

    print("\n正在升级 pip、setuptools、wheel...")
    cmd_base = f"pip install --upgrade pip setuptools wheel {mirror}".strip()
    run_cmd(cmd_base, capture=False)

    fix_versions(req_file)

    cmd = f"pip install -r {req_file} {upgrade_flag} --only-binary :all: {mirror}".strip()
    print(f"\n执行命令: {cmd}\n")
    ret, _, err = run_cmd(cmd, capture=False)
    if ret != 0:
        print(f"安装失败，请检查错误信息。")
        sys.exit(1)
    else:
        print("\n✅ 依赖更新完成！")

if __name__ == "__main__":
    main()