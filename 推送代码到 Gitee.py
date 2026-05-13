#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
推送代码到 Gitee（增强版）
- 自动完善 .gitignore（忽略 Excel、exe、大文件、临时文件、输出目录等）
- 自动排除所有文件名包含中文的文件及中文目录
- 支持强制推送（失败后询问）
"""

import os
import sys
import subprocess
import re

# ========== 忽略规则（会合并到 .gitignore） ==========
EXTRA_IGNORE = """
# 运行时生成目录
uploads/
outputs/
task_status/
logs/

# 备份目录
.git.backup/

# 虚拟环境
.venv/
venv/
ENV/
env/

# Python 缓存
__pycache__/
*.pyc
*.pyo

# IDE
.vscode/
.idea/
*.swp
.DS_Store

# 临时文件
*.tmp
*.bak
*.log
*.cache

# 大文件（数据/模型/压缩包）
*.zip
*.tar.gz
*.7z
*.rar
*.iso
*.bin
*.dat
*.pkl
*.joblib
*.parquet
*.feather
*.pth
*.ckpt
*.pb
*.onnx
*.h5
*.npy

# 测试覆盖
.pytest_cache/
.coverage
htmlcov/

# 环境变量
.env
.env.*

# 可执行文件
*.exe
*.msi

# Excel 文件（所有格式）
*.xlsx
*.xls
*.xlsm
*.xlsb
*.xltx
*.xltm

# 其他媒体大文件（可选，按需）
*.pdf
*.mp4
*.avi
*.mov
*.mkv
*.mp3
*.wav
"""

def ensure_gitignore():
    """确保 .gitignore 包含必要的忽略规则（不覆盖已有）"""
    gitignore_path = ".gitignore"
    # 读取已有内容
    existing_lines = []
    if os.path.exists(gitignore_path):
        with open(gitignore_path, "r", encoding="utf-8") as f:
            existing_lines = f.read().splitlines()
    # 合并新规则（去重）
    new_rules = []
    for line in EXTRA_IGNORE.splitlines():
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        if line not in existing_lines:
            new_rules.append(line)
    if new_rules:
        with open(gitignore_path, "a", encoding="utf-8") as f:
            f.write("\n" + "\n".join(new_rules) + "\n")
        print(f"✓ 已向 .gitignore 追加 {len(new_rules)} 条规则")
    else:
        print("✓ .gitignore 已包含所有必要规则")

def run_git_command(cmd):
    """执行 Git 命令，返回 (returncode, stdout, stderr)"""
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def is_git_repo():
    return os.path.isdir(".git")

def get_current_branch():
    ret, branch, _ = run_git_command("git rev-parse --abbrev-ref HEAD")
    return branch if ret == 0 else None

def has_changes():
    ret, status, _ = run_git_command("git status --porcelain")
    return ret == 0 and bool(status)

def add_all():
    ret, _, err = run_git_command("git add .")
    if ret != 0:
        print("添加文件失败：", err)
        return False
    print("✓ 已添加所有更改（已遵守 .gitignore）")
    return True

def commit(message):
    ret, _, err = run_git_command(f'git commit -m "{message}"')
    if ret != 0:
        print("提交失败：", err)
        return False
    print("✓ 提交成功")
    return True

def remove_chinese_items():
    """从暂存区移除所有文件名包含中文的文件，以及目录名包含中文的目录"""
    ret, files, _ = run_git_command("git ls-files --cached")
    if ret != 0 or not files:
        return

    chinese_files = []
    chinese_dirs = set()
    for f in files.splitlines():
        f = f.strip()
        if not f:
            continue
        # 检查文件名（不含路径）是否含中文
        basename = os.path.basename(f)
        if re.search(r'[\u4e00-\u9fff]', basename):
            chinese_files.append(f)
            continue
        # 检查路径中的每一级目录是否含中文
        parts = f.split('/')
        for i in range(len(parts)-1):
            if re.search(r'[\u4e00-\u9fff]', parts[i]):
                prefix = '/'.join(parts[:i+1])
                chinese_dirs.add(prefix)
                break

    if chinese_files:
        print(f"发现 {len(chinese_files)} 个文件名含中文的文件，正在从索引移除...")
        for f in chinese_files:
            run_git_command(f'git rm --cached --ignore-unmatch "{f}"')
        print("✓ 已移除中文名文件")
    if chinese_dirs:
        print(f"发现 {len(chinese_dirs)} 个目录名含中文的目录，正在递归移除...")
        for d in chinese_dirs:
            run_git_command(f'git rm -r --cached --ignore-unmatch "{d}/"')
        print("✓ 已移除中文目录下的所有内容")

def push(branch, force=False):
    cmd = f"git push origin {branch}"
    if force:
        cmd += " --force"
        print("⚠️ 使用强制推送（--force）")
    print(f"执行: {cmd}")
    ret, _, err = run_git_command(cmd)
    if ret != 0:
        if "rejected" in err.lower() or "failed" in err.lower():
            print("\n推送失败：远程仓库可能有更新。")
            if not force:
                print("提示：可以尝试强制推送（--force）覆盖远程。")
        else:
            print("推送失败：", err)
        return False
    print("✓ 推送成功")
    return True

def main():
    print("=== 推送代码到 Gitee（增强版） ===")
    if not is_git_repo():
        print("错误：当前目录不是 Git 仓库")
        sys.exit(1)

    ensure_gitignore()

    branch = get_current_branch()
    if not branch:
        print("无法获取当前分支")
        sys.exit(1)
    print(f"当前分支：{branch}")

    # 显示状态
    ret, status, _ = run_git_command("git status")
    if ret == 0:
        print("\n当前 Git 状态：")
        print(status)

    if not has_changes():
        print("\n没有需要提交的更改，程序退出。")
        return

    commit_msg = input("\n请输入提交信息（直接回车使用默认）: ").strip()
    if not commit_msg:
        commit_msg = "更新代码"

    if not add_all():
        sys.exit(1)

    # 移除中文名文件/目录
    remove_chinese_items()

    if not commit(commit_msg):
        sys.exit(1)

    # 询问是否强制推送
    print("\n准备推送到远程仓库...")
    force = False
    choice = input("是否使用强制推送（--force）？这会覆盖远程历史，请确认 (y/N): ").strip().lower()
    if choice == 'y':
        force = True
        print("\n⚠️ 您已确认强制推送，本地版本将完全覆盖远程仓库。")
        confirm = input("再次确认，输入 'YES' 继续: ").strip()
        if confirm != 'YES':
            print("已取消强制推送，程序退出。")
            sys.exit(0)

    if not push(branch, force=force):
        sys.exit(1)

    print("\n🎉 所有操作完成！")

if __name__ == "__main__":
    main()