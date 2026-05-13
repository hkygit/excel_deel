#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
终极版：重建 Git 仓库，自动检测并排除大文件，强制推送
"""

import os
import sys
import subprocess
import time
import shutil

def run_cmd(cmd, capture=True, check=False, shell=True):
    try:
        if capture:
            result = subprocess.run(cmd, shell=shell, capture_output=True, text=True,
                                    encoding='utf-8', errors='replace', check=check)
            return result.returncode, result.stdout.strip(), result.stderr.strip()
        else:
            result = subprocess.run(cmd, shell=shell, check=check)
            return result.returncode, "", ""
    except subprocess.CalledProcessError as e:
        return e.returncode, e.stdout.strip() if e.stdout else "", e.stderr.strip() if e.stderr else str(e)

def get_file_size_mb(path):
    return os.path.getsize(path) / (1024 * 1024)

def scan_large_files(root='.', limit_mb=90):
    """扫描当前目录（排除 .git）下所有超过 limit_mb 的文件"""
    large = []
    for dirpath, dirnames, filenames in os.walk(root):
        # 跳过 .git 目录
        if '.git' in dirnames:
            dirnames.remove('.git')
        if '.git.backup' in dirnames:
            dirnames.remove('.git.backup')
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            try:
                size_mb = get_file_size_mb(full)
                if size_mb > limit_mb:
                    rel = os.path.relpath(full, root)
                    large.append((rel, size_mb))
            except OSError:
                pass
    return large

def force_delete_git():
    if not os.path.exists('.git'):
        return True
    for attempt in range(1, 4):
        ret, _, _ = run_cmd('rmdir /s /q .git', capture=True)
        if ret == 0 or not os.path.exists('.git'):
            return True
        time.sleep(2)
        ret, _, _ = run_cmd('rd /s /q .git', capture=True)
        if ret == 0 or not os.path.exists('.git'):
            return True
        time.sleep(2)
    # 最后尝试 shutil
    try:
        shutil.rmtree('.git', ignore_errors=False, onerror=lambda f, p, e: os.chmod(p, 0o777) or os.unlink(p))
        return not os.path.exists('.git')
    except:
        return False

def ensure_clean_working_dir():
    """确保工作区没有大文件，并删除 .git.backup"""
    # 删除 .git.backup 如果存在
    if os.path.exists('.git.backup'):
        print("发现残留备份 .git.backup，正在删除...")
        try:
            shutil.rmtree('.git.backup', ignore_errors=False)
            print("✓ 已删除 .git.backup")
        except Exception as e:
            print(f"删除 .git.backup 失败: {e}，请手动删除。")
            return False

    # 扫描大文件
    large = scan_large_files(limit_mb=95)
    if large:
        print("\n❌ 错误：当前工作区存在超过 95MB 的大文件，这会导致推送被拒绝：")
        for path, size in large:
            print(f"  - {path} ({size:.2f} MB)")
        print("\n请手动删除或移动这些文件，然后再运行本脚本。")
        print("如果这些文件是必要的，请使用 Git LFS 管理。")
        return False
    return True

def main():
    print("=" * 60)
    print("终极版：重建 Git 仓库并排除大文件，强制推送")
    print("=" * 60)
    print("注意：会删除 .git 和 .git.backup 目录，并检查工作区大文件。")
    confirm = input("是否继续？(y/N): ").strip().lower()
    if confirm != 'y':
        return

    # 1. 删除旧 git
    if not force_delete_git():
        print("无法删除 .git 目录，请关闭占用程序后重试。")
        sys.exit(1)
    print("✓ 已删除 .git")

    # 2. 清理工作区（删除 .git.backup，检查大文件）
    if not ensure_clean_working_dir():
        sys.exit(1)

    # 3. 重新初始化
    run_cmd("git init", capture=False)
    print("✓ Git 仓库已初始化")

    # 4. 创建 .gitignore（关键！排除备份和临时文件）
    gitignore = """# 备份目录
.git.backup/
*.backup

# 大文件
*.zip
*.tar.gz
*.7z
*.rar
*.iso
*.exe
*.msi
*.bin
*.dat
*.pkl
*.joblib
*.parquet
*.feather
*.pack
*.idx

# Python
__pycache__/
*.py[cod]
*.so
.Python
env/
venv/
.venv/
ENV/
*.log
*.tmp

# IDE
.vscode/
.idea/
*.swp
.DS_Store

# 项目特定
Excel/
"""
    with open(".gitignore", "w", encoding="utf-8") as f:
        f.write(gitignore)
    print("✓ .gitignore 已创建（包含备份目录排除）")

    # 5. 添加所有文件
    print("\n添加文件到暂存区...")
    run_cmd("git add .", capture=False)

    # 6. 提交
    commit_msg = input("\n请输入提交信息（直接回车使用默认）: ").strip()
    if not commit_msg:
        commit_msg = "Initial commit: clean project"
    ret, _, err = run_cmd(f'git commit -m "{commit_msg}"')
    if ret != 0:
        print("提交失败:", err)
        sys.exit(1)
    print("✓ 提交成功")

    # 7. 远程 URL
    remote_url = input("\n请输入远程仓库 URL（直接回车使用 https://gitee.com/hky_gitee/data_deel.git）: ").strip()
    if not remote_url:
        remote_url = "https://gitee.com/hky_gitee/data_deel.git"

    # 8. 配置远程
    if run_cmd(f"git remote get-url origin", capture=True)[0] == 0:
        run_cmd(f"git remote set-url origin {remote_url}", capture=False)
        print("✓ 已更新 origin URL")
    else:
        run_cmd(f"git remote add origin {remote_url}", capture=False)
        print("✓ 已添加 origin")

    # 9. 强制推送
    branch = "master"
    print(f"\n准备强制推送到 {remote_url} ...")
    final = input("是否执行强制推送？(y/N): ").strip().lower()
    if final != 'y':
        print("已取消推送")
        return

    ret, _, err = run_cmd(f"git push -u origin {branch} --force")
    if ret != 0:
        print("推送失败:", err)
        sys.exit(1)

    print("\n🎉 推送成功！远程仓库已完全同步。")

if __name__ == "__main__":
    main()