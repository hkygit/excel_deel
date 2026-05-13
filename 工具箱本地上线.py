# start_share.py - 一键启动并生成分享链接（DNS 已系统修复）
import subprocess
import re
import sys
import os
import time
import urllib.request
import urllib.error

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def start_flask():
    print("🚀 启动 Flask 应用...")
    return subprocess.Popen(
        [sys.executable, os.path.join(BASE_DIR, "app.py")],
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding="utf-8", errors="replace",
        text=True, bufsize=1
    )

def start_tunnel():
    exe = os.path.join(BASE_DIR, "cloudflared-windows-amd64.exe")
    print("🔗 启动 Cloudflare 隧道...")
    return subprocess.Popen(
        [exe, "tunnel", "--url", "http://localhost:5000"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding="utf-8", errors="replace",
        text=True, bufsize=1
    )

def wait_for_http(url, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        try:
            urllib.request.urlopen(url, timeout=2)
            return True
        except (urllib.error.URLError, ConnectionRefusedError, OSError):
            time.sleep(1)
    return False

def main():
    flask_proc = start_flask()

    print("⏳ 等待 Flask 启动并监听端口 5000...")
    if not wait_for_http("http://localhost:5000", timeout=20):
        print("❌ 错误：无法连接到 http://localhost:5000，请检查 Flask 是否启动成功。")
        flask_proc.terminate()
        input("按回车键退出...")
        return
    print("✅ Flask 已成功启动，本地访问正常。")

    tunnel_proc = start_tunnel()
    share_url = None
    print("⏳ 正在生成分享链接...\n")
    try:
        for line in tunnel_proc.stdout:
            print(line, end='')
            if not share_url:
                match = re.search(r'https://[a-zA-Z0-9\-]+\.trycloudflare\.com', line)
                if match:
                    share_url = match.group(0)
                    print("\n" + "="*60)
                    print("✅ 分享链接生成成功！")
                    print(f"   👉 {share_url}")
                    print("="*60)
                    print("📌 将该链接发给朋友即可访问。")
                    print("📌 按 Ctrl+C 停止所有服务并退出。\n")
    except KeyboardInterrupt:
        print("\n⏹️  用户中断...")
    finally:
        tunnel_proc.terminate()
        flask_proc.terminate()
        print("👋 服务已停止。")

if __name__ == "__main__":
    main()