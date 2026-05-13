#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
回溯版本.py - 查看提交历史并回退到指定版本
功能：列出所有提交（从最新到最旧），让用户选择版本并选择重置类型（软/硬重置）
"""

import os
import sys
import subprocess
import re

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

def get_commit_history():
    """获取提交历史（hash+消息），倒序（最新在前）"""
    ret, output, err = run_git_command("git log --oneline")
    if ret != 0:
        print("获取提交历史失败：", err)
        return []
    lines = output.split('\n')
    # 过滤空行
    commits = [line.strip() for line in lines if line.strip()]
    return commits

def show_commits(commits):
    """显示提交列表（带序号）"""
    print("\n提交历史（从最新到最旧）：")
    for idx, commit in enumerate(commits, start=1):
        print(f"{idx:3d}. {commit}")

def select_commit(commits):
    """让用户选择要回退的版本（序号或hash前缀）"""
    while True:
        choice = input("\n请选择要回退到的版本（输入序号或 commit hash 前几位）: ").strip()
        if not choice:
            print("输入不能为空，请重新选择。")
            continue
        # 尝试按序号匹配
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(commits):
                return commits[idx-1].split()[0]  # 返回完整 hash
            else:
                print(f"序号超出范围（1-{len(commits)}），请重新选择。")
                continue
        # 尝试按 hash 前缀匹配
        matches = [c for c in commits if c.startswith(choice)]
        if len(matches) == 1:
            return matches[0].split()[0]
        elif len(matches) > 1:
            print("找到多个匹配的提交，请输入更长的前缀：")
            for m in matches:
                print(f"  {m}")
        else:
            print("未找到匹配的提交，请重新选择。")

def select_reset_type():
    """选择重置类型（软重置或硬重置）"""
    print("\n重置类型：")
    print("1. 软重置 (--soft) - 保留工作目录和暂存区的变更，只移动 HEAD")
    print("2. 硬重置 (--hard) - 丢弃所有变更，工作目录会回退到目标版本（危险）")
    while True:
        choice = input("请选择重置类型（1/2）: ").strip()
        if choice == "1":
            return "--soft"
        elif choice == "2":
            confirm = input("⚠️ 硬重置将丢弃所有未提交的更改，确定继续？(yes/no): ").strip().lower()
            if confirm == "yes":
                return "--hard"
            else:
                print("已取消硬重置。")
                return None
        else:
            print("无效选择，请输入 1 或 2。")

def reset_commit(commit_hash, reset_type):
    """执行 git reset"""
    if reset_type is None:
        return False
    print(f"\n正在执行 git reset {reset_type} {commit_hash} ...")
    ret, out, err = run_git_command(f"git reset {reset_type} {commit_hash}")
    if ret != 0:
        print("重置失败：", err)
        return False
    print("✓ 重置成功")
    # 显示当前状态
    run_git_command("git status")
    return True

def main():
    print("=== Git 版本回溯工具 ===")

    if not is_git_repo():
        print("错误：当前目录不是 Git 仓库，请在项目根目录运行")
        sys.exit(1)

    # 获取提交历史
    commits = get_commit_history()
    if not commits:
        print("当前仓库没有任何提交，无法回溯。")
        return

    show_commits(commits)

    # 选择目标版本
    target_hash = select_commit(commits)
    print(f"\n已选择版本: {target_hash}")

    # 选择重置类型
    reset_type = select_reset_type()
    if reset_type is None:
        print("操作已取消。")
        return

    # 执行重置
    if reset_commit(target_hash, reset_type):
        print("\n操作完成！")
    else:
        print("\n操作失败。")

if __name__ == "__main__":
    main()