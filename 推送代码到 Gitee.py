#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增强版 Git 推送工具
- 自动完善 .gitignore（忽略 Excel、exe、大文件、临时文件、输出目录等）
- 自动排除所有文件名包含中文的文件及中文目录
- 支持用户选择远程仓库（自动列出所有 remote）
- 支持普通推送或强制推送（--force），带二次确认
- 在推送失败时提供重试或拉取建议
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


def run_git_command(cmd, capture_output=True):
    """执行 Git 命令，返回 (returncode, stdout, stderr)"""
    try:
        if capture_output:
            result = subprocess.run(
                cmd,
                shell=True,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace'
            )
            return result.returncode, result.stdout.strip(), result.stderr.strip()
        else:
            # 实时输出（用于 git push）
            result = subprocess.run(cmd, shell=True)
            return result.returncode, "", ""
    except Exception as e:
        return 1, "", str(e)


def ensure_gitignore():
    """确保 .gitignore 包含必要的忽略规则（不覆盖已有）"""
    gitignore_path = ".gitignore"
    existing_lines = []
    if os.path.exists(gitignore_path):
        with open(gitignore_path, "r", encoding="utf-8") as f:
            existing_lines = f.read().splitlines()
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
        basename = os.path.basename(f)
        if re.search(r'[\u4e00-\u9fff]', basename):
            chinese_files.append(f)
            continue
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


def get_remote_list():
    """获取所有远程仓库名称列表"""
    ret, remotes, _ = run_git_command("git remote")
    if ret != 0 or not remotes:
        return []
    return [r.strip() for r in remotes.splitlines() if r.strip()]


def select_remote(remote_list):
    """交互式选择远程仓库，返回选择的远程名称"""
    if not remote_list:
        print("错误：未找到任何远程仓库。请先使用 `git remote add` 添加。")
        sys.exit(1)

    print("\n可用的远程仓库：")
    for i, name in enumerate(remote_list, 1):
        # 显示远程 URL（便于区分）
        ret, url, _ = run_git_command(f"git remote get-url {name}")
        url_str = f" ({url})" if ret == 0 and url else ""
        print(f"  {i}. {name}{url_str}")

    while True:
        choice = input(f"\n请选择要推送的远程仓库 (1-{len(remote_list)}，默认: {remote_list[0]}): ").strip()
        if not choice:
            return remote_list[0]
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(remote_list):
                return remote_list[idx]
            else:
                print(f"无效选择，请输入 1-{len(remote_list)} 之间的数字。")
        except ValueError:
            print("请输入数字。")


def push(remote_name, branch, force=False):
    """推送代码到指定远程仓库"""
    cmd = f"git push {remote_name} {branch}"
    if force:
        cmd += " --force"
        print("⚠️ 使用强制推送（--force）")
    print(f"执行: {cmd}")
    # 使用实时输出，以便看到推送进度
    ret, _, err = run_git_command(cmd, capture_output=False)
    if ret != 0:
        # 捕获错误信息（因为 capture_output=False 无法捕获，重新执行捕获）
        _, _, err = run_git_command(cmd, capture_output=True)
        print("\n推送失败。")
        if "rejected" in err.lower():
            print("原因：远程仓库有新的提交，本地落后于远程。")
            if not force:
                print("建议：先执行 `git pull --rebase` 合并远程更改，或使用强制推送（--force）。")
            else:
                print("提示：强制推送被拒绝，可能是远程仓库开启了保护（禁止强制推送）。")
        elif "failed" in err.lower():
            print("原因：网络或权限问题。")
        else:
            print("错误详情：", err)
        return False
    print("✓ 推送成功")
    return True


def main():
    print("=== 增强版 Git 推送工具 ===")

    if not is_git_repo():
        print("错误：当前目录不是 Git 仓库")
        sys.exit(1)

    ensure_gitignore()

    branch = get_current_branch()
    if not branch:
        print("无法获取当前分支")
        sys.exit(1)
    print(f"当前分支：{branch}")

    # 显示当前状态
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

    # 获取远程列表并让用户选择
    remote_list = get_remote_list()
    if not remote_list:
        print("错误：没有配置任何远程仓库。")
        sys.exit(1)
    remote_name = select_remote(remote_list)
    print(f"将推送到远程：{remote_name}")

    # 询问是否强制推送
    print("\n准备推送到远程仓库...")
    force = False
    choice = input("是否使用强制推送（--force）？这会覆盖远程历史，请确认 (y/N): ").strip().lower()
    if choice == 'y':
        force = True
        print("\n⚠️ 您已确认强制推送，本地版本将完全覆盖远程仓库。")
        confirm = input("再次确认，输入 'YES' 或 'yes' 继续: ").strip().lower()
        if confirm not in ('yes', 'y'):
            print("已取消强制推送，程序退出。")
            sys.exit(0)

    if not push(remote_name, branch, force=force):
        sys.exit(1)

    print("\n🎉 所有操作完成！")


if __name__ == "__main__":
    main()