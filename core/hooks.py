# core/hooks.py
import functools
import logging
import time
import threading
import os
from datetime import datetime, timedelta
from typing import Dict, Any, Callable, List, Optional
from blinker import Signal

logger = logging.getLogger(__name__)

class HookManager:
    _signals: Dict[str, Signal] = {}
    _async_pool = {}
    _module_loggers: Dict[str, logging.Logger] = {}

    @classmethod
    def _get_module_name(cls, hook_name: str) -> str:
        for prefix in ("HOOK_BEFORE_", "HOOK_AFTER_"):
            if hook_name.startswith(prefix):
                name = hook_name[len(prefix):]
                break
        else:
            name = hook_name
        parts = name.lower().split('_')
        return parts[0] if parts else "general"

    @classmethod
    def _get_module_logger(cls, hook_name: str) -> logging.Logger:
        module = cls._get_module_name(hook_name)
        if module in cls._module_loggers:
            return cls._module_loggers[module]

        log_dir = os.path.join("logs", "hooks", module)
        os.makedirs(log_dir, exist_ok=True)

        module_logger = logging.getLogger(f"hook.{module}")
        module_logger.setLevel(logging.INFO)
        module_logger.propagate = False

        # 使用普通 FileHandler，避免滚动问题
        log_file = os.path.join(log_dir, "hooks.log")
        handler = logging.FileHandler(log_file, encoding='utf-8')
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        module_logger.addHandler(handler)

        # 控制台输出（可选）
        console = logging.StreamHandler()
        console.setFormatter(formatter)
        module_logger.addHandler(console)

        # 强制写入一条初始化日志，确保文件创建
        module_logger.info(f"钩子日志初始化成功，模块: {module}")

        cls._module_loggers[module] = module_logger
        return module_logger

    @classmethod
    def get_signal(cls, name: str) -> Signal:
        if name not in cls._signals:
            cls._signals[name] = Signal(name)
        return cls._signals[name]

    @classmethod
    def register(cls, name: str, func: Callable, priority: int = 0, async_mode: bool = False):
        signal = cls.get_signal(name)
        def wrapper(*args, **kwargs):
            if async_mode:
                threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True).start()
            else:
                try:
                    func(*args, **kwargs)
                except Exception as e:
                    module_logger = cls._get_module_logger(name)
                    module_logger.error(f"Hook '{name}' 执行失败: {e}", exc_info=True)
                    logger.error(f"Hook '{name}' 执行失败: {e}", exc_info=True)
        if not hasattr(signal, '_priorities'):
            signal._priorities = {}
        signal._priorities[func] = priority
        signal.connect(wrapper)

    @classmethod
    def trigger(cls, name: str, **kwargs):
        signal = cls.get_signal(name)
        receivers = signal.receivers
        if hasattr(signal, '_priorities'):
            sorted_receivers = sorted(receivers, key=lambda r: signal._priorities.get(r, 0))
        else:
            sorted_receivers = list(receivers)
        for receiver in sorted_receivers:
            try:
                receiver(**kwargs)
            except Exception as e:
                module_logger = cls._get_module_logger(name)
                module_logger.error(f"Hook '{name}' receiver 异常: {e}", exc_info=True)
                logger.error(f"Hook '{name}' receiver 异常: {e}", exc_info=True)

    @classmethod
    def clear(cls, name: str = None):
        if name:
            cls.get_signal(name).receivers.clear()
        else:
            for sig in cls._signals.values():
                sig.receivers.clear()

    @classmethod
    def clean_old_logs(cls, days: int = 7):
        base_dir = os.path.join("logs", "hooks")
        if not os.path.exists(base_dir):
            return
        cutoff = datetime.now() - timedelta(days=days)
        for root, dirs, files in os.walk(base_dir):
            for f in files:
                file_path = os.path.join(root, f)
                try:
                    mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if mtime < cutoff:
                        os.remove(file_path)
                        logger.info(f"清理过期日志文件: {file_path}")
                except Exception as e:
                    logger.warning(f"清理文件失败 {file_path}: {e}")


HOOK_BEFORE_FILE_UPLOAD = "HOOK_BEFORE_FILE_UPLOAD"
HOOK_AFTER_FILE_UPLOAD = "HOOK_AFTER_FILE_UPLOAD"
HOOK_BEFORE_DATA_READ = "HOOK_BEFORE_DATA_READ"
HOOK_AFTER_DATA_READ = "HOOK_AFTER_DATA_READ"
HOOK_BEFORE_PIVOT_AGGREGATE = "HOOK_BEFORE_PIVOT_AGGREGATE"
HOOK_AFTER_PIVOT_AGGREGATE = "HOOK_AFTER_PIVOT_AGGREGATE"
HOOK_BEFORE_EXCEL_WRITE = "HOOK_BEFORE_EXCEL_WRITE"
HOOK_AFTER_EXCEL_WRITE = "HOOK_AFTER_EXCEL_WRITE"

def hook(name: str, priority: int = 0, async_mode: bool = False):
    def decorator(func):
        HookManager.register(name, func, priority, async_mode)
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        return wrapper
    return decorator

def start_log_cleaner(interval_hours: int = 24):
    def cleaner_loop():
        while True:
            time.sleep(interval_hours * 3600)
            HookManager.clean_old_logs(days=7)
    thread = threading.Thread(target=cleaner_loop, daemon=True)
    thread.start()
    return thread
