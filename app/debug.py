# -*- coding: utf-8 -*-
"""AI4Word 隐藏调试模式：以 -debug 启动时记录程序的全部操作、处理过程与异常。

触发方式（对用户完全隐藏：无菜单、无设置项、无界面变化）：
- 命令行参数：AI4Word.exe -debug / python ai4word.pyw -debug / python -m app -debug
- 程序名后缀：AI4Word-debug.exe / ai4word-debug.pyw
  （sys.argv[0] 或打包后 sys.executable 的 basename 含 -debug）

非调试模式下所有调用都是零开销空操作：不创建任何文件、不格式化字符串、
不安装钩子。调试日志写入 %APPDATA%\\AI4Word\\debug.log（回退链与
app.settings 一致：APPDATA -> 用户主目录 -> TEMP），单文件超过 5MB
滚动为 debug.log.bak（只保留一个备份）。

本模块只做记录，绝不改变任何调用方的行为与控制流。
"""
import json
import os
import sys
import threading
import time
import traceback

MAX_FIELD_LEN = 300               # 字段值默认截断长度（字符）
MAX_LOG_SIZE = 5 * 1024 * 1024    # 超过则滚动为 .bak（只留一个备份）

_SENSITIVE = ("api_key", "apikey", "key", "token", "secret", "password")

_lock = threading.RLock()
_state = {"debug": None, "path": None, "hooks": False, "failed": False}


# ---------- 检测 ----------

def _detect(argv=None):
    src = sys.argv if argv is None else list(argv)
    for a in src:
        if str(a).strip().lower() in ("-debug", "--debug"):
            return True
    base = os.path.basename(str(src[0])) if src else ""
    if "-debug" in base.lower():
        return True
    if getattr(sys, "frozen", False):
        exe = os.path.basename(str(getattr(sys, "executable", "") or ""))
        if "-debug" in exe.lower():
            return True
    return False


def _log_dir():
    """日志目录：与 settings.json / crash.log 同处，回退链一致。"""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "AI4Word")
    try:
        os.makedirs(d, exist_ok=True)
        return d
    except OSError:
        pass
    d = os.path.join(os.environ.get("TEMP") or os.path.expanduser("~") or ".",
                     "AI4Word")
    try:
        os.makedirs(d, exist_ok=True)
        return d
    except OSError:
        return None


def init(argv=None):
    """初始化调试模式（幂等，进程级）。返回是否处于调试模式。

    传入 argv 时按该参数检测；默认读 sys.argv。首次调用写运行头，
    并安装未捕获异常钩子（sys.excepthook / threading.excepthook）。
    """
    with _lock:
        if _state["debug"] is not None:
            return _state["debug"]
        enabled = _detect(argv)
        if not enabled:
            _state["debug"] = False
            _state["path"] = None
            return False
        _state["debug"] = True
        d = _log_dir()
        if d is None:
            _state["path"] = None
            _state["failed"] = True
        else:
            _state["path"] = os.path.join(d, "debug.log")
        _install_hooks()
        _write_header()
        return True


def is_debug():
    """是否处于调试模式（未初始化时自动初始化一次）。"""
    if _state["debug"] is None:
        init()
    return bool(_state["debug"])


def log_path():
    """当前日志文件路径；未启用调试模式时为 None。"""
    return _state["path"]


# ---------- 写入 ----------

def _mask(name, value):
    low = str(name).lower()
    if low in _SENSITIVE and isinstance(value, str):
        if len(value) > 12:
            return value[:6] + "..." + value[-4:]
        return "***"
    return value


def _fmt_field(name, value, full=False):
    value = _mask(name, value)
    if not isinstance(value, str):
        try:
            return name + "=" + json.dumps(value, ensure_ascii=False, default=str)
        except Exception:
            value = str(value)
    if not full and len(value) > MAX_FIELD_LEN:
        value = value[:MAX_FIELD_LEN] + "..."
    if any(c.isspace() for c in value) or '"' in value or "=" in value:
        escaped = value.replace("\\", "\\\\").replace('"', '\\"')
        return name + '="' + escaped + '"'
    return name + "=" + value


def _emit(level, event, fields, full=False):
    if not is_debug():
        return
    path = _state["path"]
    if path is None:
        return
    now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    ms = int((time.time() % 1.0) * 1000)
    thread = threading.current_thread().name
    line = "%s.%03d [%s] %s %s" % (now, ms, thread, level, event)
    if fields:
        line += " " + " ".join(_fmt_field(k, v, full=full)
                               for k, v in fields.items())
    _write(line + "\n")


def _write(text):
    """追加写一行；主路径失败时回退一次到 TEMP，再失败则静默关闭日志。"""
    with _lock:
        path = _state["path"]
        if not path:
            return
        try:
            try:
                if os.path.getsize(path) > MAX_LOG_SIZE:
                    try:
                        os.replace(path, path + ".bak")
                    except OSError:
                        pass
            except OSError:
                pass
            with open(path, "a", encoding="utf-8") as f:
                f.write(text)
            return
        except OSError:
            pass
        if _state["failed"]:
            return
        _state["failed"] = True
        alt = os.path.join(os.environ.get("TEMP") or ".", "ai4word_debug.log")
        _state["path"] = alt
        try:
            with open(alt, "a", encoding="utf-8") as f:
                f.write("[fallback] 主日志路径不可写，切换到 " + alt + "\n" + text)
        except OSError:
            _state["path"] = None


def _write_header():
    version = "?"
    try:
        from app import __version__ as v
        version = v
    except Exception:
        pass
    build = 0
    try:
        build = sys.getwindowsversion().build
    except Exception:
        pass
    _emit("INFO", "run_start", {
        "version": version,
        "frozen": bool(getattr(sys, "frozen", False)),
        "python": sys.version.split()[0],
        "executable": sys.executable,
        "argv": list(sys.argv),
        "win_build": build,
        "pid": os.getpid(),
        "log_path": _state["path"],
    })


# ---------- 未捕获异常钩子 ----------

def _install_hooks():
    if _state["hooks"]:
        return
    _state["hooks"] = True
    _orig = sys.excepthook

    def _hook(etype, value, tb):
        try:
            _emit("ERROR", "uncaught_exception", {
                "type": getattr(etype, "__name__", str(etype)),
                "value": str(value),
                "traceback": "".join(traceback.format_exception(etype, value, tb)),
            }, full=True)
        except Exception:
            pass
        _orig(etype, value, tb)

    sys.excepthook = _hook
    _orig_thread = threading.excepthook

    def _thook(args):
        try:
            _emit("ERROR", "uncaught_thread_exception", {
                "thread": getattr(args.thread, "name", "?"),
                "type": getattr(args.exc_type, "__name__", str(args.exc_type)),
                "value": str(args.exc_value),
                "traceback": "".join(traceback.format_exception(
                    args.exc_type, args.exc_value, args.exc_traceback)),
            }, full=True)
        except Exception:
            pass
        _orig_thread(args)

    threading.excepthook = _thook


# ---------- 对外 API ----------

def log(event, *, full=False, **fields):
    """记录一条普通操作 / 过程日志。full=True 时字段不截断。"""
    _emit("INFO", event, fields, full=full)


def warn(event, *, full=False, **fields):
    _emit("WARN", event, fields, full=full)


def error(event, *, full=False, **fields):
    _emit("ERROR", event, fields, full=full)


def exc(event, **fields):
    """在 except 块内调用：记录异常类型、值与完整堆栈。非调试模式零开销。"""
    if not is_debug():
        return
    fields.setdefault("traceback", traceback.format_exc())
    _emit("ERROR", event, fields, full=True)
