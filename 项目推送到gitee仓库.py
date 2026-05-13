#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
强制重建 Git 仓库并推送到指定 Gitee 仓库
- 删除所有历史
- 忽略 .exe、所有 Excel 文件、中文名文件/目录
- 自动处理大文件
"""

import os
import sys
import subprocess
import time
import re

# ========== 配置 ==========
REMOTE_URL = "https://gitee.com/hky_gitee/data_deel"
REMOTE_NAME = "origin"
BRANCH = "master"

# ========== 工具函数 ==========
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

def force_delete_git():
    """强制删除 .git 目录，自动重试"""
    if not os.path.exists('.git'):
        print("没有 .git 目录，跳过删除。")
        return True

    print("正在强制删除 .git 目录...")
    for attempt in range(1, 4):
        ret, _, err = run_cmd('rmdir /s /q .git', capture=True)
        if ret == 0 or not os.path.exists('.git'):
            print(f"✓ 删除成功（尝试 {attempt}）")
            return True
        print(f"删除失败（尝试 {attempt}），等待 2 秒...")
        time.sleep(2)
        # 备用命令
        ret, _, err = run_cmd('rd /s /q .git', capture=True)
        if ret == 0 or not os.path.exists('.git'):
            print(f"✓ 删除成功（尝试 {attempt} 使用 rd）")
            return True
        time.sleep(2)

    print("❌ 无法删除 .git 目录，请手动关闭占用程序后重试。")
    return False

def create_gitignore():
    """创建完善的 .gitignore"""
    content = """# === 可执行文件 ===
*.exe
*.msi

# === Excel 文件（所有格式） ===
*.xlsx
*.xls
*.xlsm
*.xlsb
*.xltx
*.xltm

# === 大文件（防止超限） ===
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

# === 临时文件 ===
*.tmp
*.bak
*.log
*.cache
*.pyc
__pycache__/

# === IDE ===
.vscode/
.idea/
*.swp
.DS_Store

# === 虚拟环境 ===
.venv/
venv/
ENV/
env/

# === 项目特定 ===
Excel/
*.ipynb_checkpoints/

# === 其他可能的大文件 ===
*.pdf
*.mp4
*.avi
*.mov
*.mkv
*.mp3
*.wav

# === 压缩备份 ===
*.backup
*.old
"""
    with open(".gitignore", "w", encoding="utf-8") as f:
        f.write(content)
    print("✓ .gitignore 已创建/覆盖")

def has_chinese(text):
    """判断字符串是否包含中文"""
    return bool(re.search(r'[\u4e00-\u9fff]', text))

def remove_chinese_items_from_index():
    """
    从 Git 索引中移除所有文件名包含中文的文件，以及目录名包含中文的目录（递归）
    本地文件不受影响
    """
    # 获取当前暂存区所有文件（包括已添加的）
    ret, files, err = run_cmd("git ls-files --cached")
    if ret != 0:
        print("无法获取暂存区文件列表")
        return

    chinese_files = []
    chinese_dirs = set()

    for f in files.splitlines():
        f = f.strip()
        if not f:
            continue
        # 检查文件名（不含路径）是否含中文
        basename = os.path.basename(f)
        if has_chinese(basename):
            chinese_files.append(f)
            continue
        # 检查路径中的每一级目录是否含中文
        parts = f.split('/')
        for i in range(len(parts)-1):  # 目录部分
            if has_chinese(parts[i]):
                # 标记该目录前缀
                prefix = '/'.join(parts[:i+1])
                chinese_dirs.add(prefix)
                break

    if chinese_files:
        print(f"发现 {len(chinese_files)} 个文件名含中文的文件，正在从索引移除...")
        for f in chinese_files:
            run_cmd(f'git rm --cached --ignore-unmatch "{f}"', capture=False, check=False)
        print("✓ 已移除中文名文件")

    if chinese_dirs:
        print(f"发现 {len(chinese_dirs)} 个目录名含中文的目录，正在递归移除...")
        for d in chinese_dirs:
            run_cmd(f'git rm -r --cached --ignore-unmatch "{d}/"', capture=False, check=False)
        print("✓ 已移除中文目录下的所有内容")

def main():
    print("=" * 60)
    print("强制重建 Git 仓库并推送到 Gitee")
    print("目标仓库:", REMOTE_URL)
    print("=" * 60)
    print("注意：此操作将删除所有 Git 历史记录，仅保留当前文件内容。")
    confirm = input("是否继续？(y/N): ").strip().lower()
    if confirm != 'y':
        print("已取消")
        sys.exit(0)

    # 1. 删除 .git
    if not force_delete_git():
        sys.exit(1)

    # 2. 初始化
    print("\n初始化新仓库...")
    run_cmd("git init", capture=False)

    # 3. 创建 .gitignore
    create_gitignore()

    # 4. 添加所有文件（此时 .gitignore 会生效）
    print("\n添加文件到暂存区（.gitignore 自动生效）...")
    run_cmd("git add .", capture=False)

    # 5. 移除中文名文件/目录（从暂存区）
    print("\n检查并移除中文名文件/目录...")
    remove_chinese_items_from_index()

    # 6. 检查是否有需要提交的内容
    ret, status, _ = run_cmd("git status --porcelain")
    if ret != 0 or not status:
        print("\n⚠️ 没有需要提交的文件！请检查 .gitignore 规则或项目目录是否为空。")
        sys.exit(1)

    # 7. 提交
    commit_msg = input("\n请输入提交信息（直接回车使用默认）: ").strip()
    if not commit_msg:
        commit_msg = "Initial commit: clean project without ignored files"
    ret, _, err = run_cmd(f'git commit -m "{commit_msg}"')
    if ret != 0:
        print("提交失败:", err)
        sys.exit(1)
    print("✓ 提交成功")

    # 8. 配置远程
    print(f"\n配置远程仓库 {REMOTE_NAME} -> {REMOTE_URL}")
    ret, _, _ = run_cmd(f"git remote get-url {REMOTE_NAME}", capture=True, check=False)
    if ret == 0:
        run_cmd(f"git remote set-url {REMOTE_NAME} {REMOTE_URL}", capture=False)
        print("已更新 origin URL")
    else:
        run_cmd(f"git remote add {REMOTE_NAME} {REMOTE_URL}", capture=False)
        print("已添加远程仓库")

    # 9. 强制推送
    print(f"\n准备强制推送到 {REMOTE_NAME}/{BRANCH}...")
    final = input("是否执行强制推送？(y/N): ").strip().lower()
    if final != 'y':
        print("已取消推送，本地仓库已重建但未推送。")
        sys.exit(0)

    ret, _, err = run_cmd(f"git push -u {REMOTE_NAME} {BRANCH} --force")
    if ret != 0:
        print("推送失败:", err)
        sys.exit(1)

    print("\n🎉 推送成功！远程仓库已更新，历史已清空。")

if __name__ == "__main__":
    main()