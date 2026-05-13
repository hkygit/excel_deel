#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健壮的依赖安装脚本
- 自动轮换多个镜像源（阿里云、豆瓣、华为、清华、官方）
- 跳过 pip 升级，避免 403 错误
- 支持超时和重试
"""

import subprocess
import sys
import time
import os

REQUIREMENTS_FILE = "requirements.txt"
MAX_RETRIES = 2          # 每个镜像重试次数
RETRY_DELAY = 5          # 重试间隔（秒）
TIMEOUT = 180            # 安装超时（秒）

# 镜像源列表（按优先级排序）
MIRRORS = [
    "https://mirrors.aliyun.com/pypi/simple/",
    "https://pypi.douban.com/simple/",
    "https://repo.huaweicloud.com/repository/pypi/simple/",
    "https://pypi.tuna.tsinghua.edu.cn/simple/",
    "https://pypi.org/simple/",          # 官方源，最后尝试
]

def run_pip_install(mirror):
    """尝试使用指定镜像源安装依赖"""
    host = mirror.split("//")[1].split("/")[0]
    cmd = [
        sys.executable, "-m", "pip", "install",
        "-r", REQUIREMENTS_FILE,
        "--timeout", str(TIMEOUT),
        "--retries", str(MAX_RETRIES),
        "-i", mirror,
        "--trusted-host", host,
        "--exists-action", "w",
    ]
    print(f"\n>>> 使用镜像源: {mirror}")
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            result = subprocess.run(cmd, timeout=TIMEOUT, check=False, capture_output=True, text=True)
            if result.returncode == 0:
                print("✅ 安装成功！")
                print(result.stdout)
                return True
            else:
                print(f"安装失败 (尝试 {attempt}/{MAX_RETRIES})，返回码 {result.returncode}")
                if result.stderr:
                    print("错误详情:", result.stderr[-500:])  # 只打印最后500字符
                if attempt < MAX_RETRIES:
                    print(f"等待 {RETRY_DELAY} 秒后重试...")
                    time.sleep(RETRY_DELAY)
        except subprocess.TimeoutExpired:
            print(f"安装超时（{TIMEOUT}秒），尝试 {attempt}/{MAX_RETRIES}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)
    return False

def main():
    print("=" * 60)
    print("Python 依赖安装工具（多镜像源自动轮换）")
    print(f"Python 版本: {sys.version}")
    print(f"依赖文件: {REQUIREMENTS_FILE}")
    if not os.path.exists(REQUIREMENTS_FILE):
        print(f"错误: 找不到 {REQUIREMENTS_FILE} 文件")
        sys.exit(1)
    print("=" * 60)

    # 不升级 pip，避免 403 错误
    print("跳过 pip 升级，使用当前版本。")

    for mirror in MIRRORS:
        success = run_pip_install(mirror)
        if success:
            print("\n🎉 所有依赖安装成功！")
            return
        print(f"镜像源 {mirror} 失败，切换到下一个...")
        time.sleep(RETRY_DELAY)

    print("\n❌ 所有镜像源均失败，请手动安装缺失的包。")
    print("建议手动执行：pip install -r requirements.txt")
    sys.exit(1)

if __name__ == "__main__":
    main()