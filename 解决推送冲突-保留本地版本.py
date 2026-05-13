#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键解决推送冲突 + 自动处理大文件拦截（保留本地版本）
功能：
  1. 当 git push 因远程有新提交被拒时，创建合并提交（以本地版本为准），再推送。
  2. 若推送时远程 pre-receive hook 提示文件超过限额，自动查找并清理大文件，
     然后重新执行合并 + 推送（如有必要会强制推送）。
"""

import subprocess
import sys
import os
import re
import shutil

# ---------- 配置 ----------
REMOTE_NAME = "origin"          # 远程仓库名
SIZE_LIMIT_MB = 100            # 远程允许的单文件最大 MB（用于本地扫描）
# --------------------------

def run(cmd, check=True):
    """执行命令，返回 (returncode, stdout, stderr)"""
    try:
        result = subprocess.run(
            cmd, shell=True, capture_output=True, text=True,
            encoding='utf-8', errors='replace'
        )
        if check and result.returncode != 0:
            print(f"⚠️ 命令执行异常: {cmd}\n{result.stderr.strip()}")
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def is_repo_clean():
    """检查工作区是否干净（没有未提交的更改）"""
    ret, out, _ = run("git status --porcelain", check=False)
    return ret == 0 and out == ""

def get_current_branch():
    """获取当前分支名"""
    ret, branch, err = run("git rev-parse --abbrev-ref HEAD")
    if ret != 0:
        print(f"❌ 无法获取当前分支：{err}")
        sys.exit(1)
    return branch

def find_large_file_in_history(blob_id=None):
    """
    查找历史中的大文件：
    - 如果提供了 blob_id（从远程错误消息中提取），直接定位该文件路径。
    - 否则扫描所有对象，列出超过 100MB 的文件。
    返回 (文件路径, blob_id) 或 None。
    """
    if blob_id:
        # 通过日志找到该 blob 首次出现的提交和路径
        ret, out, _ = run(
            f'git log --all --find-object={blob_id} --format="%H" --name-only',
            check=False
        )
        if ret == 0 and out:
            lines = out.splitlines()
            # 第一行是commit hash，后面是文件路径
            for line in lines[1:]:
                if line.strip():
                    return line.strip(), blob_id
        # 如果 log 失败，尝试 rev-list 直接查对象
        ret2, out2, _ = run(f"git rev-list --objects --all | grep {blob_id}", check=False)
        if ret2 == 0 and out2:
            # 输出格式：blob_id 文件路径
            parts = out2.split(maxsplit=1)
            if len(parts) > 1:
                return parts[1], blob_id
    else:
        # 扫描当前 HEAD 及全部历史，列出大文件的 blob
        print("正在扫描仓库历史中的大文件（>{}MB）...".format(SIZE_LIMIT_MB))
        # 使用 git rev-list 结合 --objects 获取所有对象及其大小
        ret, out, _ = run(
            "git rev-list --objects --all | "
            "git cat-file --batch-check='%(objecttype) %(objectsize) %(objectname) %(rest)' | "
            "awk '/^blob/ {print $2, $3, $4}' | "
            f"sort -nr | awk '$1 > {SIZE_LIMIT_MB * 1024 * 1024} {{print $2, $3}}'",
            check=False
        )
        if ret == 0 and out:
            lines = out.strip().splitlines()
            if lines:
                # 返回第一个大文件的 (路径, blob_id)
                parts = lines[0].split()
                if len(parts) >= 2:
                    return parts[1], parts[0]
    return None

def install_filter_repo():
    """提示安装 git-filter-repo 并尝试自动安装"""
    print("🔧 需要安装 git-filter-repo 以安全清理历史...")
    choice = input("是否现在安装？[Y/n] ").strip().lower()
    if choice in ('', 'y', 'yes'):
        code, out, err = run(f"{sys.executable} -m pip install git-filter-repo", check=False)
        if code == 0:
            print("✅ 安装成功")
            return True
        else:
            print(f"❌ 安装失败：{err}")
    else:
        print("跳过安装，将使用 git filter-branch（备用方案）")
    return False

def clean_history_with_filter_repo(file_path):
    """使用 git filter-repo 从整个历史中删除指定文件"""
    print(f"正在使用 git filter-repo 删除历史中的文件：{file_path}")
    # filter-repo 会重写历史，需要先备份（脚本外建议）
    ret, _, err = run(f"git filter-repo --path {file_path} --invert-paths --force", check=False)
    if ret != 0:
        print(f"❌ filter-repo 执行失败：{err}")
        return False
    print("✅ 历史已重写，大文件已移除")
    # 重写后 remote 信息可能丢失，重新添加
    run(f"git remote add {REMOTE_NAME} https://gitee.com/hky_gitee/excel-tool.git", check=False)
    return True

def clean_history_with_filter_branch(file_path):
    """使用 git filter-branch 删除指定文件（备用方案）"""
    print(f"正在使用 git filter-branch 删除历史中的文件：{file_path}")
    cmd = (f'git filter-branch --force --index-filter '
           f'"git rm --cached --ignore-unmatch {file_path}" '
           f'--prune-empty --tag-name-filter cat -- --all')
    ret, _, err = run(cmd, check=False)
    if ret != 0:
        print(f"❌ filter-branch 执行失败：{err}")
        return False
    # 清理旧引用
    run("git for-each-ref --format='%(refname)' refs/original/ | xargs -r git update-ref -d", check=False)
    run("git reflog expire --expire=now --all", check=False)
    run("git gc --prune=now --aggressive", check=False)
    print("✅ 历史已重写，大文件已移除")
    return True

def handle_large_file_error(error_msg):
    """
    解析推送失败的错误，查找大文件并清理。
    返回 True 表示清理完成，可以继续推送；False 表示用户放弃或失败。
    """
    # 尝试从错误信息中提取 blob id (形如 40位hex)
    blob_match = re.search(r'[0-9a-fA-F]{40}', error_msg)
    blob_id = blob_match.group(0) if blob_match else None

    # 查找具体的大文件
    result = find_large_file_in_history(blob_id)
    if not result:
        print("❌ 未能定位到大文件，请手动检查。")
        # 可以手动扫描
        result = find_large_file_in_history()  # 不传blob_id扫描
        if not result:
            return False

    file_path, found_blob = result
    print(f"🔍 发现大文件：{file_path} (blob: {found_blob})")

    # 确认是否清理
    confirm = input("是否要从所有历史提交中永久删除此文件？[y/N] ").strip().lower()
    if confirm != 'y':
        print("操作已取消")
        return False

    # 检查 filter-repo 是否可用
    filter_repo_available = shutil.which("git-filter-repo") or (
        run("git filter-repo --version", check=False)[0] == 0
    )
    if not filter_repo_available:
        installed = install_filter_repo()
        if installed:
            filter_repo_available = True
        else:
            print("⚠️ 将使用降级方案 git filter-branch（较慢且需手动清理引用）")

    if filter_repo_available:
        success = clean_history_with_filter_repo(file_path)
    else:
        success = clean_history_with_filter_branch(file_path)

    if success:
        print("⚠️ 历史已重写。若远程已有相关提交，推送时需使用 --force。")
        return True
    return False

def main():
    print("=== 增强版解决推送冲突（保留本地版本） ===")

    # 0. 环境检查
    if not os.path.isdir(".git"):
        print("❌ 错误：当前目录不是 Git 仓库，请在项目根目录运行。")
        sys.exit(1)

    if not is_repo_clean():
        print("❌ 工作区有未提交的更改，请先暂存或提交后再运行。")
        sys.exit(1)

    branch = get_current_branch()
    print(f"当前分支：{branch}")

    # 1. 拉取远程最新信息
    print("正在获取远程仓库最新信息...")
    ret, _, err = run(f"git fetch {REMOTE_NAME}")
    if ret != 0:
        print(f"❌ 获取远程信息失败：{err}")
        sys.exit(1)
    print("✓ 远程信息已更新")

    # 2. 创建合并提交（保留本地版本）
    print(f"正在将 {REMOTE_NAME}/{branch} 合并到本地（保留本地版本）...")
    merge_cmd = f'git merge -s ours {REMOTE_NAME}/{branch} -m "合并远程更新，保留当前本地版本"'
    ret, out, err = run(merge_cmd, check=False)
    if ret != 0:
        if "Already up to date" in out or "已经是最新的" in out:
            print("ℹ️ 本地已经是最新的，无需合并。")
        else:
            print(f"❌ 合并失败：{err}")
            sys.exit(1)
    else:
        print("✓ 合并提交已创建（本地文件未改动）")

    # 3. 尝试推送（可能被拒绝）
    print(f"正在推送到 {REMOTE_NAME}/{branch} ...")
    ret, out, err = run(f"git push {REMOTE_NAME} {branch}")
    if ret == 0:
        print("✓ 推送成功！远程仓库已更新。")
        return

    # 4. 分析推送失败原因
    combined_msg = out + "\n" + err
    # 检查是否因文件大小被拒
    if ("exceeds quota" in combined_msg or
        "pre-receive hook declined" in combined_msg or
        "remote rejected" in combined_msg):
        print("\n⚠️ 推送被远程拒绝，似乎是由于历史中存在大文件。")
        cleaned = handle_large_file_error(combined_msg)
        if not cleaned:
            sys.exit(1)

        # 清理完成后，重新进行合并与推送
        print("\n由于历史已重写，现在重新执行合并与推送...")
        # 重新 fetch（远程可能有新变化）
        run(f"git fetch {REMOTE_NAME}")
        # 再次执行 ours 合并
        merge_cmd = f'git merge -s ours {REMOTE_NAME}/{branch} -m "合并远程更新（清理大文件后）"'
        ret, out, _ = run(merge_cmd, check=False)
        if ret != 0 and "Already up to date" not in out:
            print(f"❌ 重新合并失败：{out}")
            sys.exit(1)

        # 由于历史重写，通常需要强制推送
        print("历史已重写，将执行强制推送（--force）...")
        # 二次确认
        force_confirm = input(f"确认强制推送到 {REMOTE_NAME}/{branch}？[y/N] ").strip().lower()
        if force_confirm != 'y':
            print("操作取消")
            sys.exit(0)

        ret, out, err = run(f"git push --force {REMOTE_NAME} {branch}")
        if ret == 0:
            print("✅ 强制推送成功！")
        else:
            print(f"❌ 推送仍然失败：{err}")
            print("可能原因：远程分支受保护，不允许强制推送。请在 Gitee 仓库设置中解除保护后重试。")
            sys.exit(1)
    else:
        print(f"❌ 推送失败：{err}")
        sys.exit(1)

if __name__ == "__main__":
    main()