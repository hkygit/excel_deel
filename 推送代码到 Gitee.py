#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
推送代码到Gitee.py - 将当前项目的所有更改提交并推送到 Gitee 远程仓库
功能：自动 add/commit/push，要求用户输入提交信息，检测冲突并提示
增强：自动忽略 Excel/ 目录（添加到 .gitignore）
"""

import os
import sys
import subprocess
import re


def ensure_gitignore_ignore_excel():
    """确保 .gitignore 文件中包含 Excel/ 忽略规则"""
    gitignore_path = ".gitignore"
    ignore_line = "Excel/"

    # 如果 .gitignore 不存在，创建并写入
    if not os.path.exists(gitignore_path):
        with open(gitignore_path, "w", encoding="utf-8") as f:
            f.write(f"{ignore_line}\n")
        print("✓ 已创建 .gitignore 并添加 Excel/ 忽略规则")
        return

    # 检查是否已存在该规则
    with open(gitignore_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    if any(line.strip() == ignore_line.rstrip('/') or line.strip() == ignore_line for line in lines):
        print("✓ .gitignore 已包含 Excel/ 忽略规则")
        return

    # 追加规则
    with open(gitignore_path, "a", encoding="utf-8") as f:
        f.write(f"\n{ignore_line}\n")
    print("✓ 已向 .gitignore 添加 Excel/ 忽略规则")


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
        stdout = result.stdout.strip() if result.stdout is not None else ""
        stderr = result.stderr.strip() if result.stderr is not None else ""
        return result.returncode, stdout, stderr
    except Exception as e:
        return 1, "", str(e)


def is_git_repo():
    """检查当前目录是否为 Git 仓库"""
    return os.path.isdir(".git")


def get_current_branch():
    """获取当前分支名"""
    ret, branch, err = run_git_command("git rev-parse --abbrev-ref HEAD")
    if ret != 0:
        return None
    return branch


def show_status():
    """显示当前 Git 状态"""
    ret, status, err = run_git_command("git status")
    if ret == 0:
        print("\n当前 Git 状态：")
        print(status)
    else:
        print("无法获取 Git 状态：", err)


def has_changes():
    """检查是否有未提交的更改（排除 Excel/ 目录）"""
    ret, status, _ = run_git_command("git status --porcelain")
    if ret != 0:
        return False
    # 如果 status 为空，则无更改
    if not status:
        return False
    # 可选：进一步过滤掉 Excel/ 目录的更改（但 .gitignore 已经处理，无需额外过滤）
    return True


def add_all():
    """添加所有更改到暂存区（Excel/ 目录已被 .gitignore 忽略）"""
    ret, _, err = run_git_command("git add .")
    if ret != 0:
        print("添加文件失败：", err)
        return False
    print("✓ 已添加所有更改（已忽略 Excel/ 目录）")
    return True


def commit(message):
    """提交更改"""
    ret, _, err = run_git_command(f'git commit -m "{message}"')
    if ret != 0:
        print("提交失败：", err)
        return False
    print("✓ 提交成功")
    return True


def push(branch):
    """推送到远程仓库"""
    print(f"正在推送到远程仓库（分支 {branch}）...")
    ret, out, err = run_git_command(f"git push origin {branch}")
    if ret != 0:
        if "failed to push" in err.lower() or "rejected" in err.lower():
            print("\n推送失败！远程仓库可能有更新。")
            print("建议先执行 `git pull` 合并远程更改，再重新推送。")
        else:
            print("推送失败：", err)
        return False
    print("✓ 推送成功")
    return True


def main():
    print("=== 推送代码到 Gitee ===")

    if not is_git_repo():
        print("错误：当前目录不是 Git 仓库，请在项目根目录运行")
        sys.exit(1)

    # 确保 Excel/ 目录被 Git 忽略
    ensure_gitignore_ignore_excel()

    branch = get_current_branch()
    if not branch:
        print("无法获取当前分支，请检查仓库状态")
        sys.exit(1)
    print(f"当前分支：{branch}")

    show_status()

    if not has_changes():
        print("\n没有需要提交的更改，程序退出。")
        return

    # 获取提交消息
    commit_msg = input("\n请输入提交信息（可直接回车使用默认消息）: ").strip()
    if not commit_msg:
        commit_msg = "更新代码"

    # 执行 add, commit, push
    if not add_all():
        sys.exit(1)
    if not commit(commit_msg):
        sys.exit(1)
    if not push(branch):
        sys.exit(1)

    print("\n所有操作完成！")


if __name__ == "__main__":
    main()