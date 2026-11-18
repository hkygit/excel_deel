#!/usr/bin/env python3
"""
测试运行脚本

⭐ 跳过不常用模块：
   默认跳过 SKIP_TEST_FILES 中列出的测试文件（当前：车队日报）。
   命令行加 --all 时跑全量，加 --skip 追加跳过文件。

- 全部通过：仅打印成功信息，不生成日志
- 有失败或异常：落盘到 log/ 目录
- 区分"pytest 启动失败"与"真正 0 失败"

用法：
    python run_tests.py                 # 默认（跳过车队日报）
    python run_tests.py --all           # 全量
    python run_tests.py --skip tests/test_xxx.py,tests/test_yyy.py  # 追加跳过
    python run_tests.py --show-skip     # 显示当前跳过列表
"""

import os
import sys
import subprocess
import re
import traceback
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# 环境守卫：强制切换到项目 .venv
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
VENV_PY = PROJECT_ROOT / ".venv" / "Scripts" / "python.exe"
LOG_DIR = PROJECT_ROOT / "log"
TEST_TIMEOUT_SECONDS = 30 * 60


def _ensure_project_venv():
    if not VENV_PY.exists():
        print(f"⚠️ 项目 venv 不存在: {VENV_PY}")
        print("   请先运行 fix_env.py 或手动创建 .venv")
        os.environ.pop("PYTHONPATH", None)
        return

    try:
        current = Path(sys.executable).resolve()
        target = VENV_PY.resolve()
    except OSError:
        return

    if current == target:
        if os.environ.get("PYTHONPATH"):
            print(f"[环境清理] PYTHONPATH 已移除")
            os.environ.pop("PYTHONPATH", None)
        return

    print(f"[环境切换] 当前解释器: {current}")
    print(f"[环境切换] 项目 venv:  {target}")
    print(f"[环境切换] 重新启动 run_tests.py 到项目 venv ...")
    sys.stdout.flush()
    sys.stderr.flush()

    clean_env = os.environ.copy()
    clean_env.pop("PYTHONPATH", None)
    clean_env.pop("PYTHONHOME", None)
    path_parts = clean_env.get("PATH", "").split(os.pathsep)
    filtered = []
    for p in path_parts:
        pl = p.lower()
        if ".venv" in pl and "scripts" in pl and str(PROJECT_ROOT).lower() not in pl:
            continue
        filtered.append(p)
    clean_env["PATH"] = os.pathsep.join(filtered)

    try:
        result = subprocess.run(
            [str(target), str(Path(__file__).resolve())] + sys.argv[1:],
            env=clean_env,
        )
        sys.exit(result.returncode)
    except Exception as e:
        print(f"[环境切换] 失败: {e}")
        os.environ.pop("PYTHONPATH", None)


_ensure_project_venv()


# ---------------------------------------------------------------------------
# 跳过列表（默认跳过的测试文件）
# 要恢复某个文件：从列表里删掉即可
# 要新增跳过：往列表里加相对路径
# ---------------------------------------------------------------------------
SKIP_TEST_FILES = [
    # ---- 车队日报（不常用，跳过） ----
    "tests/test_team_report_long_running.py",
    "tests/test_integration_team_report.py",
    "tests/test_team_report_service.py",
]


# ---------------------------------------------------------------------------
# 参数解析
# ---------------------------------------------------------------------------
def parse_args():
    run_all = "--all" in sys.argv
    show_skip = "--show-skip" in sys.argv

    extra_skip = []
    for i, arg in enumerate(sys.argv):
        if arg == "--skip" and i + 1 < len(sys.argv):
            extra_skip = [
                s.strip() for s in sys.argv[i + 1].split(",") if s.strip()
            ]
        elif arg.startswith("--skip="):
            extra_skip = [
                s.strip() for s in arg[len("--skip="):].split(",") if s.strip()
            ]

    return run_all, show_skip, extra_skip


# ---------------------------------------------------------------------------
# 构造命令
# ---------------------------------------------------------------------------
def build_cmd(run_all, extra_skip):
    cmd = [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"]

    if run_all:
        print("🔍 全量模式（不跳过任何测试）")
        return cmd

    skip_files = list(SKIP_TEST_FILES) + [s for s in extra_skip if s not in SKIP_TEST_FILES]

    # 过滤不存在的文件（避免 pytest 因为 --ignore 不存在的路径报错）
    valid_skip = []
    for f in skip_files:
        p = PROJECT_ROOT / f
        if p.exists():
            valid_skip.append(f)
        else:
            print(f"⚠️ 跳过列表中文件不存在，忽略: {f}")

    if valid_skip:
        print(f"⚡ 跳过 {len(valid_skip)} 个文件（共约 {len(skip_files)} 个请求跳过）")
        for f in valid_skip:
            cmd.extend(["--ignore", f])
    else:
        print("⚡ 跳过列表为空，等价于全量")

    return cmd


# ---------------------------------------------------------------------------
# 运行
# ---------------------------------------------------------------------------
def run_tests(cmd):
    print(f"🚀 运行测试: {' '.join(cmd)}")
    print(f"   解释器: {sys.executable}")
    print(f"   版本:   {sys.version.split()[0]}")

    started = datetime.now()
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=TEST_TIMEOUT_SECONDS,
        )
        output = (result.stdout or "") + (result.stderr or "")
        elapsed = (datetime.now() - started).total_seconds()
        return result.returncode, output, None, elapsed
    except subprocess.TimeoutExpired as e:
        partial = ""
        for stream in (e.stdout, e.stderr):
            if not stream:
                continue
            partial += stream if isinstance(stream, str) else stream.decode("utf-8", "replace")
        elapsed = (datetime.now() - started).total_seconds()
        return -1, partial, f"pytest 运行超时（>{TEST_TIMEOUT_SECONDS}s）", elapsed
    except Exception as e:
        elapsed = (datetime.now() - started).total_seconds()
        return -1, "", f"无法启动 pytest: {e}\n{traceback.format_exc()}", elapsed


def extract_failed_tests(output: str) -> list:
    tests = []
    for pat in (r"^(.+?\.py::[^ ]+)\s+FAILED", r"^(.+?\.py::[^ ]+)\s+ERROR"):
        for m in re.finditer(pat, output or "", re.MULTILINE):
            name = m.group(1)
            if name not in tests:
                tests.append(name)
    return tests


def looks_like_pytest_bootstrap_failure(output: str) -> bool:
    if not output:
        return True
    if "test session starts" in output:
        return False
    markers = ("ModuleNotFoundError", "ImportError:",
               "Traceback (most recent call last)", "No module named")
    return any(m in output for m in markers)


def save_failure_log(output, failed_tests, returncode=None,
                     error_msg=None, elapsed=None, bootstrap_failed=False):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"test_failures_{ts}.log"

    lines = [f"测试运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"]
    lines.append(f"解释器: {sys.executable}")
    lines.append(f"版本:   {sys.version.split()[0]}")
    if elapsed is not None:
        lines.append(f"耗时: {elapsed:.2f}s")
    if returncode is not None:
        lines.append(f"pytest 退出码: {returncode}")
    if error_msg:
        lines.append(f"运行异常: {error_msg}")
    if bootstrap_failed:
        lines.append("⚠️ pytest 启动失败（未能正常 collect）")
    lines.append(f"失败/错误测试数量: {len(failed_tests)}")
    lines.append("-" * 60)

    if failed_tests:
        lines.append("失败/错误测试列表:")
        for t in failed_tests:
            lines.append(f"  - {t}")
        lines.append("-" * 60)

    log_file.write_text("\n".join(lines) + "\n\n" + (output or "(无输出)"), encoding="utf-8")

    print("\n" + "=" * 60)
    if bootstrap_failed:
        print("❌ pytest 启动失败（还没开始 collect）")
    else:
        print("⚠️ 测试未通过（或运行异常），完整输出如下：")
    print("=" * 60)
    print(output if output else "(无输出)")
    print("=" * 60)
    print(f"\n📁 失败日志已保存至: {log_file}")

    if bootstrap_failed:
        print("\n💡 诊断建议：")
        print(f"   1) 确认项目 venv 依赖齐全：")
        print(f"      {VENV_PY} -m pip list")
        print(f"   2) 若缺包，重装：")
        print(f"      {VENV_PY} -m pip install -r requirements.txt")
    elif failed_tests:
        print("\n失败/错误测试列表：")
        for t in failed_tests:
            print(f"   ❌ {t}")


def main():
    run_all, show_skip, extra_skip = parse_args()

    if show_skip:
        print("当前默认跳过的测试文件：")
        for f in SKIP_TEST_FILES:
            exists = "✓" if (PROJECT_ROOT / f).exists() else "✗ (不存在)"
            print(f"  {exists} {f}")
        return 0

    cmd = build_cmd(run_all, extra_skip)
    returncode, output, error_msg, elapsed = run_tests(cmd)

    if returncode == 0 and error_msg is None:
        print(f"\n✅ 所有测试通过！耗时 {elapsed:.2f}s。未生成日志文件。")
        return 0

    failed_tests = extract_failed_tests(output)
    bootstrap_failed = (not failed_tests) and looks_like_pytest_bootstrap_failure(output)

    save_failure_log(output, failed_tests, returncode=returncode,
                     error_msg=error_msg, elapsed=elapsed,
                     bootstrap_failed=bootstrap_failed)
    sys.exit(returncode if returncode != 0 else 1)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        tb = traceback.format_exc()
        try:
            save_failure_log(tb, [], returncode=-1,
                             error_msg="run_tests.py 自身异常",
                             bootstrap_failed=True)
        except Exception:
            print(tb, file=sys.stderr)
        sys.exit(1)