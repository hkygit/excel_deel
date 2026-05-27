#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键修复 GitHub 大文件问题（使用 Git LFS）
- 自动检测并迁移所有超过 100MB 的 .xlsx 文件到 LFS
- 保留完整历史（仅改变存储方式，不丢失文件内容）
- 自动安装 LFS 跟踪规则
- 最后询问是否强制推送到 GitHub
"""

import os
import sys
import subprocess

def run_cmd(cmd, capture=True, check=False):
    """执行命令"""
    try:
        if capture:
            result = subprocess.run(
                cmd, shell=True, capture_output=True, text=True,
                encoding='utf-8', errors='replace'
            )
            return result.returncode, result.stdout.strip(), result.stderr.strip()
        else:
            result = subprocess.run(cmd, shell=True)
            return result.returncode, "", ""
    except Exception as e:
        return 1, "", str(e)

def is_git_repo():
    return os.path.isdir(".git")

def check_git_lfs():
    """检查 git-lfs 是否安装并初始化"""
    ret, _, _ = run_cmd("git lfs version")
    if ret != 0:
        print("❌ git-lfs 未安装。请先安装：https://git-lfs.github.com/")
        return False
    # 检查当前仓库是否已初始化 LFS
    ret, _, _ = run_cmd("git lfs track")
    if ret != 0:
        print("正在初始化 git-lfs...")
        run_cmd("git lfs install")
    return True

def is_working_tree_clean():
    """检查工作区是否有未提交的更改"""
    ret, status, _ = run_cmd("git status --porcelain")
    if ret != 0:
        return True  # 无法判断，假设干净
    return len(status.strip()) == 0

def commit_all_changes(commit_msg="自动提交未跟踪更改（LFS 迁移前）"):
    """添加并提交所有当前更改"""
    print("检测到未提交的更改，正在自动提交...")
    run_cmd("git add .")
    ret, _, err = run_cmd(f'git commit -m "{commit_msg}"')
    if ret != 0:
        print("自动提交失败，请手动提交后再运行脚本。")
        return False
    print("✅ 已提交当前更改。")
    return True

def find_large_files():
    """找出超过 100MB 的文件路径（基于当前和历史）"""
    # 使用 git lfs migrate info 获取大文件列表
    ret, output, _ = run_cmd("git lfs migrate info --everything --top=10")
    if ret != 0:
        print("无法获取大文件列表。")
        return []
    large_files = []
    for line in output.splitlines():
        # 输出格式如：*.xlsx   231 MB  100%
        if 'MB' in line and ('xlsx' in line or 'xls' in line):
            try:
                parts = line.split()
                if len(parts) >= 2 and parts[1].endswith('MB'):
                    size_mb = float(parts[1].replace('MB', ''))
                    if size_mb > 100:
                        large_files.append(parts[0])  # 文件名模式
            except:
                pass
    # 如果没有找到，手动添加已知问题文件
    if not large_files:
        # 检查是否存在诊断文件
        ret, _, _ = run_cmd("git ls-tree HEAD -- '诊断文件/2022年结算总数据.xlsx'")
        if ret == 0:
            large_files.append("诊断文件/2022年结算总数据.xlsx")
    return large_files

def migrate_to_lfs(patterns):
    """迁移匹配的文件到 LFS"""
    # 先添加 LFS 跟踪规则
    for pat in patterns:
        print(f"设置 LFS 跟踪: {pat}")
        run_cmd(f"git lfs track '{pat}'")
    # 提交 .gitattributes
    run_cmd("git add .gitattributes")
    run_cmd('git commit -m "Add LFS tracking"')
    # 执行迁移
    include_pattern = ",".join(patterns)
    print(f"正在迁移历史中的文件（模式: {include_pattern}）...")
    cmd = f"git lfs migrate import --include='{include_pattern}' --everything --verbose"
    ret, _, err = run_cmd(cmd)
    if ret != 0:
        print("迁移失败：", err)
        return False
    print("✅ 迁移成功！")
    return True

def force_push():
    """强制推送到远程仓库（用户选择）"""
    # 获取远程仓库列表
    ret, remotes, _ = run_cmd("git remote")
    if ret != 0 or not remotes:
        print("没有找到任何远程仓库。")
        return False
    remote_list = remotes.splitlines()
    print("可用的远程仓库：")
    for i, r in enumerate(remote_list, 1):
        print(f"  {i}. {r}")
    choice = input(f"请选择要推送的远程仓库 (1-{len(remote_list)}，默认: {remote_list[0]}): ").strip()
    if choice.isdigit():
        idx = int(choice) - 1
        if 0 <= idx < len(remote_list):
            remote = remote_list[idx]
        else:
            remote = remote_list[0]
    else:
        remote = remote_list[0] if not choice else choice
    branch = run_cmd("git rev-parse --abbrev-ref HEAD")[1]
    if not branch:
        branch = "master"
    print(f"\n准备强制推送 {remote}/{branch}")
    print("⚠️ 警告：强制推送会覆盖远程历史，请确保团队成员知晓。")
    confirm = input("确认继续？(y/N): ").strip().lower()
    if confirm != 'y':
        print("已取消。")
        return False
    confirm2 = input("再次确认，输入 'YES' 继续: ").strip().lower()
    if confirm2 != 'yes':
        print("已取消。")
        return False
    cmd = f"git push {remote} {branch} --force"
    print(f"执行: {cmd}")
    ret, _, err = run_cmd(cmd, capture=False)
    if ret != 0:
        print("推送失败。", err)
        return False
    print("✅ 推送成功！")
    return True

def main():
    print("=== Git LFS 一键修复大文件问题 ===")
    if not is_git_repo():
        print("错误：当前目录不是 Git 仓库。")
        sys.exit(1)

    if not check_git_lfs():
        sys.exit(1)

    # 确保工作区干净
    if not is_working_tree_clean():
        if not commit_all_changes():
            sys.exit(1)

    large_files = find_large_files()
    if not large_files:
        print("未发现超过 100MB 的大文件，无需修复。")
        # 可选：仍然询问是否推送
        ans = input("是否仍要强制推送当前分支？(y/N): ").strip().lower()
        if ans == 'y':
            force_push()
        return

    print(f"发现大文件模式: {large_files}")
    print("将使用 Git LFS 迁移这些文件的历史，保留完整内容。")
    print("注意：此操作会重写 Git 历史，但文件内容不会丢失。")
    ans = input("是否继续？(y/N): ").strip().lower()
    if ans != 'y':
        print("已取消。")
        return

    if not migrate_to_lfs(large_files):
        print("迁移失败，请检查错误信息。")
        sys.exit(1)

    # 清理无用对象（可选）
    run_cmd("git reflog expire --expire=now --all")
    run_cmd("git gc --prune=now --aggressive")

    # 推送
    if force_push():
        print("\n🎉 修复完成！GitHub 仓库已成功推送，不再有大文件错误。")
    else:
        print("\n推送失败，请手动执行强制推送。")

if __name__ == "__main__":
    main()