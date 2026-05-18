#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从远程 Git 仓库拉取代码，支持合并、强制覆盖本地等选项。
"""

import os
import sys
import subprocess

def run_cmd(cmd, capture=True):
    """执行命令，返回 (returncode, stdout, stderr)"""
    try:
        if capture:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='utf-8')
            return result.returncode, result.stdout.strip(), result.stderr.strip()
        else:
            result = subprocess.run(cmd, shell=True)
            return result.returncode, "", ""
    except Exception as e:
        return 1, "", str(e)

def is_git_repo():
    return os.path.isdir(".git")

def has_uncommitted_changes():
    """检查是否有未提交的更改（包括未跟踪文件）"""
    ret, out, _ = run_cmd("git status --porcelain")
    return ret == 0 and bool(out)

def get_current_branch():
    ret, branch, _ = run_cmd("git rev-parse --abbrev-ref HEAD")
    return branch if ret == 0 else None

def stash_changes():
    print("正在暂存本地更改...")
    ret, _, err = run_cmd("git stash push -u -m 'auto-stash before pull'")
    if ret != 0:
        print(f"暂存失败: {err}")
        return False
    print("已暂存本地更改")
    return True

def drop_stash():
    ret, _, err = run_cmd("git stash drop")
    if ret != 0:
        print(f"删除暂存失败: {err}")
    else:
        print("已删除暂存")

def apply_stash():
    ret, _, err = run_cmd("git stash pop")
    if ret != 0:
        print(f"恢复暂存失败: {err}")
        return False
    print("已恢复本地更改")
    return True

def force_reset(branch):
    """强制重置到远程分支，丢弃所有本地更改"""
    print(f"强制重置到 origin/{branch}，将丢弃所有本地更改")
    ret, _, err = run_cmd(f"git reset --hard origin/{branch}")
    if ret != 0:
        print(f"重置失败: {err}")
        return False
    print("本地已强制同步为远程版本")
    return True

def normal_pull(branch):
    """普通 pull（合并）"""
    print("执行 git pull...")
    ret, out, err = run_cmd(f"git pull origin {branch}")
    if ret != 0:
        print(f"拉取失败: {err}")
        return False
    print(out)
    print("拉取成功")
    return True

def main():
    print("="*60)
    print("从远程仓库拉取代码")
    print("="*60)

    if not is_git_repo():
        print("错误：当前目录不是 Git 仓库")
        sys.exit(1)

    branch = get_current_branch()
    if not branch:
        print("无法获取当前分支")
        sys.exit(1)
    print(f"当前分支: {branch}")

    # 检查远程是否存在
    ret, _, _ = run_cmd(f"git ls-remote --heads origin {branch}")
    if ret != 0:
        print(f"警告：远程分支 origin/{branch} 不存在，请先推送或切换分支")
        sys.exit(1)

    # 检查是否有未提交的更改
    if has_uncommitted_changes():
        print("\n⚠️ 检测到本地有未提交的更改。")
        print("1. 暂存本地更改（stash），拉取后恢复")
        print("2. 强制覆盖本地（丢弃所有本地更改）")
        print("3. 放弃拉取，手动处理")
        choice = input("请选择 (1/2/3): ").strip()
        if choice == '1':
            if not stash_changes():
                sys.exit(1)
            stashed = True
        elif choice == '2':
            if force_reset(branch):
                print("拉取完成，本地已与远程同步")
                sys.exit(0)
            else:
                sys.exit(1)
        else:
            print("已取消拉取")
            sys.exit(0)
    else:
        stashed = False

    # 执行普通 pull
    if normal_pull(branch):
        if stashed:
            apply_stash()
        print("\n✅ 拉取成功！")
    else:
        # pull 失败（可能冲突）
        if stashed:
            print("拉取失败，尝试恢复本地更改...")
            apply_stash()
        print("\n请手动解决冲突后重新运行脚本。")
        sys.exit(1)

if __name__ == "__main__":
    main()