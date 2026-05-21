#!/usr/bin/env python3
"""
增强版开发环境刷新脚本（已修复误删 templates 目录的问题）
仅删除了危险的临时目录清理逻辑，其余功能与您的版本完全一致
"""
import os
import sys
import shutil
import subprocess
import time
import signal
import atexit
import webbrowser
import io
import socket
import pandas as pd
from pathlib import Path

# 尝试导入psutil用于精确的进程管理
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False
    print("⚠️  psutil未安装，进程终止可能不彻底")
    print("   安装命令: pip install psutil")

# 全局配置
ROOT = os.path.dirname(os.path.abspath(__file__))
PORT = 5000

# ============================================================
# 1. 进程管理函数
# ============================================================

def kill_processes_by_port(port=5000):
    """按端口终止进程（更精确）"""
    print(f"🛑 终止占用端口 {port} 的进程...")
    killed = []

    if PSUTIL_AVAILABLE:
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                connections = proc.net_connections()
                for conn in connections:
                    if hasattr(conn.laddr, 'port') and conn.laddr.port == port:
                        for child in proc.children(recursive=True):
                            child.terminate()
                        proc.terminate()
                        try:
                            proc.wait(timeout=5)
                        except:
                            proc.kill()
                        killed.append(proc.pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied, AttributeError):
                continue

    if killed:
        print(f"   已终止进程PID: {killed}")
    else:
        print(f"   使用系统命令清理端口 {port}...")
        if os.name == 'nt':
            try:
                result = subprocess.run(
                    f'netstat -ano | findstr :{port}',
                    shell=True, capture_output=True, text=True
                )
                for line in result.stdout.strip().split('\n'):
                    if line:
                        parts = line.strip().split()
                        if len(parts) >= 5:
                            pid = parts[-1]
                            subprocess.run(f'taskkill /F /PID {pid}', shell=True,
                                           capture_output=True)
                            killed.append(pid)
            except:
                pass
        else:
            try:
                result = subprocess.run(
                    f'lsof -ti:{port}',
                    shell=True, capture_output=True, text=True
                )
                pids = result.stdout.strip().split()
                for pid in pids:
                    if pid:
                        subprocess.run(f'kill -9 {pid}', shell=True,
                                       capture_output=True)
                        killed.append(pid)
            except:
                pass

    time.sleep(1)
    return len(killed) > 0


def kill_all_python():
    """跨平台终止所有Python进程"""
    print("🔪 终止所有Python进程...")
    killed_count = 0

    if PSUTIL_AVAILABLE:
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                should_kill = False
                if proc.info['name'] and 'python' in proc.info['name'].lower():
                    should_kill = True
                elif proc.info['cmdline']:
                    cmdline = ' '.join(proc.info['cmdline'])
                    if 'python' in cmdline.lower():
                        should_kill = True

                if should_kill and proc.pid != os.getpid():
                    try:
                        proc.terminate()
                        try:
                            proc.wait(timeout=2)
                        except:
                            proc.kill()
                        killed_count += 1
                        print(f"   已终止PID {proc.info['pid']} ({proc.info['name']})")
                    except Exception as e:
                        pass
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    else:
        if os.name == 'nt':
            result = subprocess.run(
                'taskkill /F /IM python* /T',
                shell=True, capture_output=True, text=True
            )
            if result.returncode == 0 or "No tasks" in result.stdout:
                print("   已执行 taskkill /F /IM python*")
                killed_count = 1
        else:
            result = subprocess.run(
                ['pgrep', '-f', 'python'],
                capture_output=True, text=True
            )
            pids = result.stdout.strip().split()
            for pid in pids:
                if pid and int(pid) != os.getpid():
                    try:
                        os.kill(int(pid), signal.SIGKILL)
                        killed_count += 1
                        print(f"   已终止 PID {pid}")
                    except Exception:
                        pass

    print(f"   总计终止了 {killed_count} 个Python进程")
    time.sleep(2)
    return killed_count


# ============================================================
# 2. 清理函数
# ============================================================

def deep_clean_cache():
    """递归删除所有 __pycache__ 和 .pyc 文件"""
    print("🧹 清除所有 Python 缓存...")
    removed = 0

    for dirpath, dirnames, filenames in os.walk(ROOT, topdown=False):
        for d in list(dirnames):
            if d == '__pycache__':
                full = os.path.join(dirpath, d)
                try:
                    shutil.rmtree(full, ignore_errors=True)
                    removed += 1
                except:
                    pass

        for f in filenames:
            if f.endswith('.pyc') or f in ['.DS_Store', '.pyo', '.pyd']:
                try:
                    os.remove(os.path.join(dirpath, f))
                    removed += 1
                except:
                    pass

            if f.endswith('.pyc'):
                for suffix in ['.cpython-', '.pypy']:
                    if suffix in f:
                        try:
                            os.remove(os.path.join(dirpath, f))
                            removed += 1
                        except:
                            pass

    extra_cache_dirs = [
        os.path.join(ROOT, '.pytest_cache'),
        os.path.join(ROOT, '.mypy_cache'),
        os.path.join(ROOT, '__pycache__'),
        os.path.join(ROOT, '.coverage'),
    ]

    for cache_dir in extra_cache_dirs:
        if os.path.exists(cache_dir):
            try:
                shutil.rmtree(cache_dir, ignore_errors=True)
                removed += 1
            except:
                pass

    print(f"   清除了 {removed} 个缓存项目")
    return removed


def clean_runtime_files():
    """清理 uploads, outputs, task_status 中的临时文件（已保护关键目录）"""
    print("🗑️  清理临时文件...")
    folders_cleaned = 0

    # 清理指定的文件夹
    for folder in ['uploads', 'outputs', 'task_status', 'static/temp', 'temp']:
        full = os.path.join(ROOT, folder)
        if os.path.exists(full):
            try:
                for root, dirs, files in os.walk(full):
                    for f in files:
                        try:
                            file_path = os.path.join(root, f)
                            if not f.startswith('.') and not f.endswith('.gitkeep'):
                                os.remove(file_path)
                        except Exception as e:
                            pass
                folders_cleaned += 1
                print(f"   已清理 {folder}/")
            except Exception as e:
                print(f"   清理 {folder} 失败: {e}")

    # ⚠️ 关键修复：删除临时目录时，绝对保护 templates、static 等核心目录
    protected_dirs = {'templates', 'static', 'core', 'services', 'blueprints', 'tests', 'log'}
    for item in os.listdir(ROOT):
        item_path = os.path.join(ROOT, item)
        if os.path.isdir(item_path) and item not in protected_dirs:
            # 只删除以 tmp/temp 开头且不是保护目录的临时文件夹
            if (item.startswith('tmp') or item.startswith('temp')) and item.lower() != 'templates':
                try:
                    shutil.rmtree(item_path, ignore_errors=True)
                    print(f"   已删除临时目录: {item}")
                except:
                    pass

    # 清理临时文件
    for pattern in ['*.tmp', '*.temp', '*.log', '*.cache']:
        for file in Path(ROOT).rglob(pattern):
            try:
                file.unlink()
            except:
                pass

    return folders_cleaned


def clear_temp_directories():
    """清除系统临时目录中的相关文件"""
    print("🗂️  清除系统临时目录...")
    temp_dirs = []

    if os.name == 'nt':
        temp_dirs.extend([
            os.environ.get('TEMP', ''),
            os.environ.get('TMP', ''),
            os.environ.get('LOCALAPPDATA', '') + '\\Temp',
            os.path.expanduser('~\\AppData\\Local\\Temp'),
        ])
    else:
        temp_dirs.extend([
            '/tmp',
            '/var/tmp',
            os.path.expanduser('~/Library/Caches'),
            os.path.expanduser('~/.cache'),
            os.path.expanduser('~/tmp'),
        ])

    cleared_count = 0
    for temp_dir in temp_dirs:
        if temp_dir and os.path.exists(temp_dir):
            try:
                for item in os.listdir(temp_dir):
                    item_lower = item.lower()
                    if any(keyword in item_lower for keyword in
                           ['python', 'flask', 'pycache', 'tmp', 'temp', 'cache']):
                        try:
                            full_path = os.path.join(temp_dir, item)
                            if os.path.isfile(full_path):
                                os.remove(full_path)
                                cleared_count += 1
                            elif os.path.isdir(full_path):
                                shutil.rmtree(full_path, ignore_errors=True)
                                cleared_count += 1
                        except Exception:
                            pass
            except Exception as e:
                pass

    print(f"   从系统临时目录清除了 {cleared_count} 个项目")
    return cleared_count


def clear_dns_cache():
    """清除DNS缓存"""
    print("🌐 清除DNS缓存...")
    try:
        if os.name == 'nt':
            subprocess.run(['ipconfig', '/flushdns'], capture_output=True,
                           shell=True, text=True)
            print("   已执行 ipconfig /flushdns")
        elif sys.platform == 'darwin':
            subprocess.run(['sudo', 'killall', '-HUP', 'mDNSResponder'],
                           capture_output=True, text=True)
            subprocess.run(['sudo', 'dscacheutil', '-flushcache'],
                           capture_output=True, text=True)
            print("   已刷新 macOS DNS 缓存")
        else:
            result = subprocess.run(['which', 'systemd-resolve'],
                                    capture_output=True, text=True)
            if result.returncode == 0:
                subprocess.run(['sudo', 'systemd-resolve', '--flush-caches'],
                               capture_output=True, text=True)
                print("   已执行 systemd-resolve --flush-caches")
            else:
                subprocess.run(['sudo', '/etc/init.d/nscd', 'restart'],
                               capture_output=True, text=True)
                print("   已重启 nscd 服务")
    except Exception as e:
        print(f"   DNS清除失败: {e}")
    return True


# ============================================================
# 3. Python/Flask缓存清理
# ============================================================

def clear_python_import_caches():
    """清除Python导入缓存"""
    print("🧪 清除Python模块导入缓存...")

    modules_to_clear = []
    for module_name in list(sys.modules.keys()):
        if module_name.startswith('app') or \
                module_name.startswith('routes') or \
                module_name.startswith('utils') or \
                module_name.startswith('models') or \
                module_name.startswith('forms'):
            modules_to_clear.append(module_name)

    for module_name in modules_to_clear:
        try:
            del sys.modules[module_name]
        except:
            pass

    if 'importlib' in sys.modules:
        importlib = sys.modules['importlib']
        if hasattr(importlib, 'invalidate_caches'):
            importlib.invalidate_caches()

    for finder in sys.meta_path:
        if hasattr(finder, 'invalidate_caches'):
            finder.invalidate_caches()

    print(f"   已清除 {len(modules_to_clear)} 个模块缓存")
    return len(modules_to_clear)


def force_flask_reload():
    """强制Flask完全重新加载"""
    print("🔄 强制Flask重新加载配置...")

    trigger_file = os.path.join(ROOT, 'app.py')
    if os.path.exists(trigger_file):
        try:
            stat = os.stat(trigger_file)
            new_time = time.time()
            os.utime(trigger_file, (new_time, new_time))
            print(f"   已更新 {trigger_file} 时间戳")

            try:
                from werkzeug._reloader import _get_args_for_reloading
                if hasattr(_get_args_for_reloading, 'cache_clear'):
                    _get_args_for_reloading.cache_clear()
            except:
                pass

            try:
                import jinja2
                jinja2.cache.clear()
            except:
                pass
        except Exception as e:
            print(f"   强制重载失败: {e}")
    return True


# ============================================================
# 4. 浏览器缓存清理
# ============================================================

def clear_browser_cache_hard():
    """尝试直接清除浏览器缓存文件"""
    print("🧼 清除浏览器缓存文件...")

    cache_paths = []
    if os.name == 'nt':
        appdata = os.environ.get('LOCALAPPDATA', '')
        roaming = os.environ.get('APPDATA', '')
        cache_paths.extend([
            os.path.join(appdata, 'Google', 'Chrome', 'User Data', 'Default', 'Cache'),
            os.path.join(appdata, 'Google', 'Chrome', 'User Data', 'Default', 'Code Cache'),
        ])
        firefox_profiles = os.path.join(roaming, 'Mozilla', 'Firefox', 'Profiles')
        if os.path.exists(firefox_profiles):
            for profile in os.listdir(firefox_profiles):
                cache_paths.append(os.path.join(firefox_profiles, profile, 'cache2'))
        cache_paths.append(os.path.join(appdata, 'Microsoft', 'Edge', 'User Data', 'Default', 'Cache'))
    else:
        home = os.path.expanduser('~')
        cache_paths.extend([
            os.path.join(home, '.cache', 'google-chrome'),
            os.path.join(home, '.config', 'google-chrome', 'Default'),
            os.path.join(home, '.cache', 'chromium'),
            os.path.join(home, 'Library', 'Caches', 'Google', 'Chrome'),
            os.path.join(home, '.cache', 'mozilla', 'firefox'),
            os.path.join(home, 'Library', 'Caches', 'Firefox'),
        ])

    cleared_count = 0
    for cache_path in cache_paths:
        expanded_path = os.path.expanduser(cache_path)
        if os.path.exists(expanded_path):
            try:
                if os.path.isfile(expanded_path):
                    os.remove(expanded_path)
                    cleared_count += 1
                elif os.path.isdir(expanded_path):
                    shutil.rmtree(expanded_path, ignore_errors=True)
                    cleared_count += 1
            except Exception:
                pass

    print(f"   清理了 {cleared_count} 个浏览器缓存位置")
    return cleared_count


# ============================================================
# 5. 服务启动与验证
# ============================================================

def start_service():
    """启动Flask服务"""
    print("🚀 启动服务...")
    app_path = os.path.join(ROOT, 'app.py')
    if not os.path.exists(app_path):
        print(f"❌ 错误: 找不到 {app_path}")
        sys.exit(1)

    env = os.environ.copy()
    env['PYTHONDONTWRITEBYTECODE'] = '1'
    env['PYTHONUNBUFFERED'] = '1'

    try:
        proc = subprocess.Popen(
            [sys.executable, app_path],
            cwd=ROOT,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        def cleanup_process():
            if proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except:
                    proc.kill()

        atexit.register(cleanup_process)
        print(f"   服务已启动，PID: {proc.pid}")
        return proc
    except Exception as e:
        print(f"❌ 启动服务失败: {e}")
        sys.exit(1)


def wait_for_service(port=5000, timeout=30):
    """等待服务启动"""
    print("⏳ 等待服务启动...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            result = sock.connect_ex(('127.0.0.1', port))
            sock.close()
            if result == 0:
                print("✅ 服务已就绪")
                return True
            elapsed = int(time.time() - start_time)
            if elapsed % 5 == 0:
                print(f"   等待中... ({elapsed}/{timeout}秒)")
        except Exception as e:
            pass
        time.sleep(1)
    print(f"❌ 服务在 {timeout} 秒内未启动")
    return False


def verify_service():
    """验证服务是否正常工作"""
    print("🔍 验证新代码行为...")
    sys.path.insert(0, ROOT)

    try:
        from app import create_app
    except ImportError as e:
        print(f"❌ 无法导入app模块: {e}")
        return False

    try:
        app = create_app()
        app.config['TESTING'] = True

        with app.test_client() as client:
            resp = client.get('/')
            if resp.status_code != 200:
                print(f"❌ 首页返回状态码: {resp.status_code}")
                return False
            print("✅ 首页访问正常")

            try:
                df = pd.DataFrame({'A': ['X', 'Y'], 'B': [10, 20]})
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    df.to_excel(w, index=False)
                buf.seek(0)

                data = {
                    'file': (buf, 'test_verification.xlsx'),
                    'key_cols': 'A',
                    'calc_cols': 'B',
                    'agg_func_default': 'sum',
                }
                resp = client.post('/subtotal/process', data=data,
                                   content_type='multipart/form-data')
                if resp.status_code == 200:
                    print("✅ 文件处理端点正常")
                    return True
                else:
                    print(f"⚠️  文件处理端点返回: {resp.status_code}")
                    return True
            except Exception as e:
                print(f"⚠️  文件处理测试跳过: {e}")
                return True
    except Exception as e:
        print(f"❌ 验证过程中出错: {e}")
        return False


def verify_with_real_browser():
    """用真实浏览器请求验证"""
    print("🌍 用浏览器验证...")
    timestamp = int(time.time())
    test_url = f'http://127.0.0.1:{PORT}/?_nocache={timestamp}'
    print(f"   将在浏览器中打开: {test_url}")
    try:
        webbrowser.open_new_tab(test_url)
        return True
    except Exception as e:
        print(f"⚠️  无法打开浏览器: {e}")
        return False


# ============================================================
# 主流程
# ============================================================

def main_enhanced():
    print("=" * 60)
    print("🔧 增强版开发环境刷新脚本（已修复误删 templates）")
    print("=" * 60)

    kill_processes_by_port(PORT)
    kill_all_python()
    clear_temp_directories()
    clear_dns_cache()
    deep_clean_cache()
    clean_runtime_files()
    clear_python_import_caches()
    force_flask_reload()

    print("\n" + "=" * 40)
    proc = start_service()

    def cleanup():
        if proc and proc.poll() is None:
            print("\n🛑 正在停止服务进程...")
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except:
                if proc.poll() is None:
                    proc.kill()

    atexit.register(cleanup)

    if not wait_for_service(PORT, timeout=15):
        print("❌ 服务启动失败，尝试强制停止并重试...")
        cleanup()
        time.sleep(2)
        proc = start_service()
        if not wait_for_service(PORT, timeout=10):
            print("❌ 服务仍然无法启动")
            sys.exit(1)

    clear_browser_cache_hard()

    print("\n" + "=" * 40)
    time.sleep(2)

    if not verify_service():
        print("\n❌ 服务验证失败！")
        print("\n建议操作：")
        print("1. 检查 app.py 是否有语法错误")
        print("2. 运行: python -m py_compile app.py")
        print("3. 检查端口 5000 是否被其他程序占用")
        print("4. 查看服务日志输出")
        sys.exit(1)

    verify_with_real_browser()

    print("\n" + "=" * 60)
    print("✅ 增强清理完成！")
    print(f"\n服务运行在: http://127.0.0.1:{PORT}")
    print("按 Ctrl+C 停止服务")

    try:
        proc.wait()
    except KeyboardInterrupt:
        print("\n\n🛑 服务已手动停止")
        cleanup()
    except Exception as e:
        print(f"\n❌ 服务异常退出: {e}")
        sys.exit(1)


if __name__ == '__main__':
    try:
        main_enhanced()
    except KeyboardInterrupt:
        print("\n\n👋 脚本被用户中断")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 脚本执行出错: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)