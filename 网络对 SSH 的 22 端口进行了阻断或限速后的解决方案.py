#!/usr/bin/env python3
"""
push_https.py
使用 HTTPS + 个人访问令牌 推送代码到 GitHub
运行：python push_https.py
"""

import subprocess
import sys
from pathlib import Path

def run_cmd(cmd, capture=True):
    try:
        if capture:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            return result.stdout.strip(), result.stderr.strip(), result.returncode
        else:
            result = subprocess.run(cmd, shell=True)
            return "", "", result.returncode
    except Exception as e:
        return "", str(e), -1

def main():
    print("="*60)
    print("GitHub HTTPS 推送助手")
    print("="*60)

    # 1. 检查是否在 Git 仓库中
    if not (Path.cwd() / ".git").exists():
        print("当前目录不是 Git 仓库，正在初始化...")
        run_cmd("git init", capture=False)
        run_cmd("git add .", capture=False)
        run_cmd('git commit -m "Initial commit"', capture=False)
        print("✓ 已初始化并提交")

    # 2. 获取仓库地址
    remote_url = input("\n请输入你的 GitHub HTTPS 仓库地址（例如 https://github.com/用户名/仓库名.git）：").strip()
    if not remote_url:
        print("未输入地址，退出。")
        return

    # 3. 设置远程仓库
    run_cmd("git remote remove origin")
    stdout, stderr, code = run_cmd(f"git remote add origin {remote_url}")
    if code != 0:
        print(f"设置远程仓库失败：{stderr}")
        return
    print(f"✓ 已设置远程仓库为 {remote_url}")

    # 4. 提示生成个人访问令牌
    print("\n" + "="*60)
    print("⚠️  重要提示：")
    print("请先在 GitHub 上生成个人访问令牌（repo 权限）")
    print("步骤：")
    print("  1. 登录 GitHub，点击头像 → Settings → Developer settings → Personal access tokens → Tokens (classic)")
    print("  2. 点击 Generate new token (classic)，勾选 'repo'，生成后复制令牌（只显示一次）")
    print("="*60)
    input("\n准备好后，按 Enter 键继续...")

    # 5. 推送代码
    print("\n正在推送代码...")
    print("（首次推送会提示输入用户名和密码，密码请粘贴个人访问令牌）")
    stdout, stderr, code = run_cmd("git push -u origin main", capture=False)
    if code != 0:
        # 尝试 master 分支
        print("尝试推送 master 分支...")
        stdout2, stderr2, code2 = run_cmd("git push -u origin master", capture=False)
        if code2 == 0:
            print("\n✅ 推送成功！")
        else:
            print("\n❌ 推送失败，请检查：")
            print("   - 仓库地址是否正确")
            print("   - 个人访问令牌是否有效（需要 repo 权限）")
            print("   - 分支名称是否正确（手动查看本地分支：git branch）")
    else:
        print("\n✅ 推送成功！")

if __name__ == "__main__":
    main()