#!/usr/bin/env python3
"""
开发环境重启脚本：
- 强制终止占用 5000 端口的旧进程
- 清空所有 Python 缓存
- 直接运行 app.py 启动服务（保留原有启动逻辑）
"""
import os
import sys
import shutil
import signal
import subprocess
import time

PORT = 5000

def kill_port(port):
    """跨平台终止占用指定端口的进程"""
    if os.name == 'nt':  # Windows
        result = subprocess.run(
            f'netstat -ano | findstr :{port}',
            shell=True, capture_output=True, text=True
        )
        for line in result.stdout.splitlines():
            if f':{port}' in line and 'LISTENING' in line:
                pid = line.strip().split()[-1]
                subprocess.run(f'taskkill /PID {pid} /F', shell=True)
                print(f"已终止 Windows 进程 PID {pid} (端口 {port})")
    else:  # Linux/macOS
        result = subprocess.run(
            f'lsof -ti :{port}',
            shell=True, capture_output=True, text=True
        )
        pids = result.stdout.strip().split()
        for pid in pids:
            os.kill(int(pid), signal.SIGKILL)
            print(f"已终止进程 PID {pid} (端口 {port})")

def clean_cache(root_dir):
    """删除所有 __pycache__, .pyc, .pytest_cache"""
    removed_dirs = 0
    removed_files = 0
    for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
        for d in list(dirnames):
            if d in ('__pycache__', '.pytest_cache'):
                full = os.path.join(dirpath, d)
                shutil.rmtree(full, ignore_errors=True)
                removed_dirs += 1
        for f in filenames:
            if f.endswith('.pyc') or f == '.DS_Store':
                full = os.path.join(dirpath, f)
                os.remove(full)
                removed_files += 1
    print(f"已清理 {removed_dirs} 个缓存目录, {removed_files} 个缓存文件")

if __name__ == '__main__':
    root = os.path.dirname(os.path.abspath(__file__))

    print("1. 终止旧进程...")
    kill_port(PORT)
    time.sleep(1)

    print("2. 清理缓存...")
    clean_cache(root)

    print("3. 启动服务（app.py）...")
    # 直接执行 app.py，保留其内部的 waitress.serve 调用
    subprocess.run([sys.executable, os.path.join(root, 'app.py')], cwd=root)