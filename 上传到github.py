#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增强版 Git 推送工具（v5）
- 自动完善 .gitignore
- 自动排除中文名文件/目录
- 支持选择远程仓库、强制推送（带确认）
- 推送失败给出建议
- 先 add 再检查变更，列出被忽略文件
- 自动修复 dubious ownership，使用错误信息中的路径
"""

import os
import sys
import subprocess
import re

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

# 大文件
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

# 测试
.pytest_cache/
.coverage
htmlcov/

# 环境变量
.env
.env.*

# 可执行文件
*.exe
*.msi

# Excel 文件
*.xlsx
*.xls
*.xlsm
*.xlsb
*.xltx
*.xltm

# 媒体大文件
*.pdf
*.mp4
*.avi
*.mov
*.mkv
*.mp3
*.wav
"""


def run_git_command(cmd, capture_output=True):
    try:
        if capture_output:
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


def ensure_gitignore():
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
    ret, out, err = run_git_command("git symbolic-ref --short HEAD")
    if ret == 0 and out:
        return out
    ret, out, err = run_git_command("git branch --show-current")
    if ret == 0 and out:
        return out
    head_path = os.path.join(".git", "HEAD")
    if os.path.isfile(head_path):
        try:
            with open(head_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
            if content.startswith("ref: refs/heads/"):
                return content.split("/")[-1]
            elif re.match(r'^[0-9a-f]{40}$', content):
                print("⚠️ 当前处于 detached HEAD 状态，无法获取分支名。")
                return None
        except Exception:
            pass
    print("无法自动获取当前分支。")
    if err:
        print("错误详情：", err)
    return None


def has_commits():
    ret, _, _ = run_git_command("git rev-parse --verify HEAD")
    return ret == 0


def has_staged_changes():
    ret, out, _ = run_git_command("git diff --cached --name-only")
    return ret == 0 and bool(out)


def add_all():
    """执行 git add .，并在遇到 dubious ownership 时自动修复"""
    ret, _, err = run_git_command("git add .")
    if ret == 0:
        print("✓ 已添加所有更改（已遵守 .gitignore）")
        return True

    # 检查是否为 dubious ownership 错误
    if "dubious ownership" in err.lower():
        print("\n⚠️ Git 检测到仓库目录所有权异常（dubious ownership）。")
        # 从错误信息中提取路径（Git 报告的是正斜杠格式）
        match = re.search(r"repository at '([^']+)'", err)
        if match:
            repo_path_from_git = match.group(1)
        else:
            # 如果提取失败，用当前路径转换
            repo_path_from_git = os.path.abspath('.').replace('\\', '/')
        print(f"仓库路径（Git 识别）: {repo_path_from_git}")
        # 添加安全目录
        cmd = f'git config --global --add safe.directory "{repo_path_from_git}"'
        ret2, _, err2 = run_git_command(cmd)
        if ret2 != 0:
            print(f"自动添加安全目录失败: {err2}")
            print("请手动执行以下命令后重新运行：")
            print(f'  git config --global --add safe.directory "{repo_path_from_git}"')
            return False
        print("✓ 安全目录已添加，正在重试 git add ...")
        # 重试添加
        ret3, _, err3 = run_git_command("git add .")
        if ret3 == 0:
            print("✓ 已添加所有更改")
            return True
        else:
            print(f"重试添加失败: {err3}")
            return False
    else:
        print("添加文件失败：", err)
        return False


def commit(message):
    ret, _, err = run_git_command(f'git commit -m "{message}"')
    if ret != 0:
        print("提交失败：", err)
        return False
    print("✓ 提交成功")
    return True


def remove_chinese_items():
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


def list_ignored_files():
    ret, out, _ = run_git_command("git ls-files --others --ignored --exclude-standard")
    if ret == 0 and out:
        return out.splitlines()
    return []


def get_remote_list():
    ret, remotes, _ = run_git_command("git remote")
    if ret != 0 or not remotes:
        return []
    return [r.strip() for r in remotes.splitlines() if r.strip()]


def select_remote(remote_list):
    if not remote_list:
        print("错误：未找到任何远程仓库。请先使用 `git remote add` 添加。")
        sys.exit(1)

    print("\n可用的远程仓库：")
    for i, name in enumerate(remote_list, 1):
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
    cmd = f"git push {remote_name} {branch}"
    if force:
        cmd += " --force"
        print("⚠️ 使用强制推送（--force）")
    print(f"执行: {cmd}")
    ret, _, err = run_git_command(cmd, capture_output=False)
    if ret != 0:
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
    print("=== 增强版 Git 推送工具 (v5) ===")

    if not is_git_repo():
        print("错误：当前目录不是 Git 仓库")
        sys.exit(1)

    ensure_gitignore()

    # ---- 分支处理 ----
    branch = get_current_branch()
    if branch is None:
        if not has_commits():
            print("\n⚠️ 当前仓库没有任何提交。")
            print("请先添加文件并创建首次提交。")
            branch = "main"
            print(f"将使用分支名：{branch}")
            ret, _, _ = run_git_command(f"git checkout -b {branch} 2>/dev/null || true")
            if ret != 0:
                print("自动创建分支失败，请手动执行：git checkout -b main")
                sys.exit(1)
        else:
            print("\n⚠️ 无法自动获取当前分支名称。")
            while True:
                branch = input("请手动输入当前分支名称（例如 main、master）: ").strip()
                if branch:
                    break
                print("分支名不能为空。")
            ret, _, _ = run_git_command(f"git checkout {branch} 2>/dev/null || git checkout -b {branch}")
            if ret != 0:
                print(f"无法切换到或创建分支 {branch}，请手动处理。")
                sys.exit(1)

    print(f"当前分支：{branch}")

    # ---- 显示当前状态 ----
    ret, status, _ = run_git_command("git status")
    if ret == 0:
        print("\n当前 Git 状态：")
        print(status)

    # ---- 先执行 git add . ----
    if not add_all():
        sys.exit(1)

    # ---- 移除中文名文件/目录（在 add 之后处理暂存区） ----
    remove_chinese_items()

    # ---- 检查暂存区是否有变更 ----
    if not has_staged_changes():
        print("\n⚠️ 暂存区没有变更（所有更改可能已被 .gitignore 忽略，或文件未修改）。")
        ignored = list_ignored_files()
        if ignored:
            print("\n以下文件被 .gitignore 忽略，未被添加：")
            for f in ignored[:20]:
                print(f"  - {f}")
            if len(ignored) > 20:
                print(f"  ... 还有 {len(ignored)-20} 个文件")
            print("\n提示：若您希望提交这些文件，请修改 .gitignore 规则后重新运行。")
        else:
            print("当前没有未忽略的新文件或修改。")
        choice = input("\n是否仍然继续提交（将创建空提交）？(y/N): ").strip().lower()
        if choice != 'y':
            print("已退出。")
            sys.exit(0)
        ret, _, err = run_git_command('git commit --allow-empty -m "空提交（无变更）"')
        if ret != 0:
            print("提交失败：", err)
            sys.exit(1)
        print("✓ 已创建空提交（无变更）")
    else:
        commit_msg = input("\n请输入提交信息（直接回车使用默认）: ").strip()
        if not commit_msg:
            commit_msg = "更新代码"
        if not commit(commit_msg):
            sys.exit(1)

    # ---- 推送 ----
    remote_list = get_remote_list()
    if not remote_list:
        print("错误：没有配置任何远程仓库。")
        sys.exit(1)
    remote_name = select_remote(remote_list)
    print(f"将推送到远程：{remote_name}")

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