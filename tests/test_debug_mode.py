# -*- coding: utf-8 -*-
"""离线测试：隐藏调试模式（app.debug）。

覆盖：argv / 程序名后缀触发与误报排除、非调试模式零副作用（不建文件、
不装钩子）、幂等初始化、行格式、敏感字段打码、超长截断与 full=True、
exc() 写 traceback、5MB 滚动、级别区分、未捕获异常钩子仅调试时安装、
多线程并发写不交错。
"""
import os
import re
import sys
import threading

try:
    import pytest
except ImportError:  # .venv 无 pytest 时可用 python tests/test_debug_mode.py 直接跑
    pytest = None


if pytest:
    _fixture = pytest.fixture
else:

    def _fixture(*args, **kwargs):
        if args and callable(args[0]) and not kwargs:
            return args[0]

        def deco(fn):
            return fn

        return deco


sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import debug


@_fixture(autouse=True)
def _restore_exception_hooks():
    saved = (sys.excepthook, threading.excepthook)
    yield
    sys.excepthook, threading.excepthook = saved


def _scratch_dir():
    """日志目录：直接写 TEMP 根下的 debug.log。

    与既有测试套件（test_settings 等）一致；workspace-write 沙箱下
    也只有 TEMP 根可直接写文件。完全不可写时回退到测试目录。"""
    base = os.environ.get("TEMP") or os.path.expanduser("~") or "."
    probe = os.path.join(base, "debug.log")
    try:
        with open(probe, "a", encoding="utf-8"):
            pass
        os.remove(probe)
        return base
    except OSError:
        pass
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_debug_scratch")
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


def _make(monkeypatch):
    d = _scratch_dir()
    for name in ("debug.log", "debug.log.bak"):  # 保证每个测试从空白日志开始
        try:
            os.remove(os.path.join(d, name))
        except OSError:
            pass
    monkeypatch.setattr(debug, "_state",
                        {"debug": None, "path": None, "hooks": False, "failed": False})
    monkeypatch.setattr(debug, "_log_dir", lambda: d)
    return d


def _read_log(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def test_argv_debug_flag_triggers(monkeypatch):
    d = _make(monkeypatch)
    assert debug.init(["AI4Word.exe", "-debug"]) is True
    path = os.path.join(d, "debug.log")
    assert os.path.exists(path)
    text = _read_log(path)
    assert "run_start" in text
    assert "frozen" in text  # 运行头字段齐全


def test_argv_double_dash(monkeypatch):
    _make(monkeypatch)
    assert debug.init(["python", "ai4word.pyw", "--debug"]) is True


def test_argv_flag_case_insensitive(monkeypatch):
    _make(monkeypatch)
    assert debug.init(["AI4Word.exe", "-DEBUG"]) is True


def test_exe_name_suffix_triggers(monkeypatch):
    _make(monkeypatch)
    assert debug.init([r"D:\out\AI4Word-debug.exe"]) is True
    _make(monkeypatch)
    assert debug.init(["ai4word-debug.pyw"]) is True


def test_no_false_positive_from_dir_name(monkeypatch):
    """目录名含 -debug 不算：只看 basename。"""
    _make(monkeypatch)
    assert debug.init([r"D:\debug-tools\x\AI4Word.exe"]) is False


def test_no_debug_no_file_no_hooks(monkeypatch):
    d = _make(monkeypatch)
    saved = (sys.excepthook, threading.excepthook)
    assert debug.init(["AI4Word.exe"]) is False
    assert debug.log_path() is None
    assert not os.path.exists(os.path.join(d, "debug.log"))
    # 各种写入口都应是零开销空操作
    debug.log("x", a=1)
    debug.warn("x")
    debug.error("x")
    try:
        raise ValueError("boom")
    except ValueError:
        debug.exc("x")
    assert not os.path.exists(os.path.join(d, "debug.log"))
    assert sys.excepthook is saved[0]
    assert threading.excepthook is saved[1]


def test_init_is_idempotent(monkeypatch):
    d = _make(monkeypatch)
    debug.init(["AI4Word.exe", "-debug"])
    first = debug.log_path()
    assert debug.init(["AI4Word.exe", "-debug"]) is True  # 已初始化：argv 不再生效
    assert debug.log_path() == first
    text = _read_log(os.path.join(d, "debug.log"))
    assert text.count("run_start") == 1


def test_line_format(monkeypatch):
    d = _make(monkeypatch)
    debug.init(["AI4Word", "-debug"])
    debug.log("evt_one", count=3, name="hello world")
    line = [ln for ln in _read_log(os.path.join(d, "debug.log")).splitlines()
            if "evt_one" in ln][0]
    pat = (r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} "
           r"\[.*\] INFO evt_one count=3 name=\"hello world\"$")
    assert re.match(pat, line), line


def test_sensitive_fields_masked(monkeypatch):
    d = _make(monkeypatch)
    debug.init(["AI4Word", "-debug"])
    secret = "sk-abcdefgh1234567890"
    debug.log("api_call", api_key=secret, token="tok-1234567890", model="m")
    text = _read_log(os.path.join(d, "debug.log"))
    assert secret not in text
    assert "tok-1234567890" not in text
    # 掩码格式：value[:6] + "..." + value[-4:]
    assert "sk-abc...7890" in text
    assert "tok-12...7890" in text


def test_truncation_and_full(monkeypatch):
    d = _make(monkeypatch)
    debug.init(["AI4Word", "-debug"])
    long_val = "x" * 1000
    debug.log("trunc_evt", big=long_val)
    debug.log("full_evt", big=long_val, full=True)
    text = _read_log(os.path.join(d, "debug.log"))
    trunc_line = [ln for ln in text.splitlines() if "trunc_evt" in ln][0]
    full_line = [ln for ln in text.splitlines() if "full_evt" in ln][0]
    assert len(trunc_line) < len(long_val)
    assert trunc_line.rstrip().endswith("...")
    assert long_val in full_line


def test_exc_records_traceback(monkeypatch):
    d = _make(monkeypatch)
    debug.init(["AI4Word", "-debug"])
    try:
        raise RuntimeError("deliberate failure")
    except RuntimeError:
        debug.exc("op_failed", stage="middle")
    text = _read_log(os.path.join(d, "debug.log"))
    assert re.search(r"ERROR op_failed", text)
    assert "Traceback" in text
    assert "deliberate failure" in text
    assert "RuntimeError" in text


def test_levels_warn_and_error(monkeypatch):
    d = _make(monkeypatch)
    debug.init(["AI4Word", "-debug"])
    debug.warn("warm_evt")
    debug.error("bad_evt")
    text = _read_log(os.path.join(d, "debug.log"))
    assert re.search(r"WARN warm_evt", text)
    assert re.search(r"ERROR bad_evt", text)


def test_rotation_at_max_size(monkeypatch):
    d = _make(monkeypatch)
    debug.init(["AI4Word", "-debug"])
    monkeypatch.setattr(debug, "MAX_LOG_SIZE", 2048)
    path = os.path.join(d, "debug.log")
    bak = path + ".bak"
    debug.log("first", payload="z" * 4000, full=True)   # 字段默认截断 300 字符，这里用 full=True 才能写超阈值
    assert not os.path.exists(bak)
    debug.log("second", payload="y" * 4000, full=True)  # 写入前超阈值 → 滚动
    assert os.path.exists(bak)
    assert "first" in _read_log(bak)
    assert "second" in _read_log(path)
    debug.log("third", payload="x" * 4000, full=True)   # 再次滚动：旧 .bak 被替换，只留一个
    assert "second" in _read_log(bak)
    assert "third" in _read_log(path)


def test_hooks_installed_only_in_debug(monkeypatch):
    _make(monkeypatch)
    saved = (sys.excepthook, threading.excepthook)
    debug.init(["AI4Word", "-debug"])
    assert debug._state["hooks"] is True
    assert sys.excepthook is not saved[0]
    assert threading.excepthook is not saved[1]
    # 未捕获异常也能写进日志
    try:
        sys.excepthook(ValueError, ValueError("uncaught!"), None)
    except Exception:
        pass
    text = _read_log(debug.log_path())
    assert "uncaught_exception" in text
    assert "uncaught!" in text


def test_concurrent_writers_interleave_safely(monkeypatch):
    """多线程并发写：行数严格等于写入次数，没有半行拼凑。"""
    d = _make(monkeypatch)
    debug.init(["AI4Word", "-debug"])
    n = 60
    errs = []

    def worker(i):
        try:
            for k in range(n):
                debug.log("thread_evt", worker=i, seq=k)
        except Exception as e:  # pragma: no cover - 防御
            errs.append(e)

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert not errs
    lines = [ln for ln in _read_log(os.path.join(d, "debug.log")).splitlines()
             if "thread_evt" in ln]
    assert len(lines) == 4 * n
    for ln in lines:
        assert re.match(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} ", ln), ln


if __name__ == "__main__":
    import inspect
    import traceback

    class _MiniMonkey:
        def __init__(self):
            self._undo = []

        def setattr(self, obj, name, value):
            self._undo.append((obj, name, getattr(obj, name, None)))
            setattr(obj, name, value)

        def undo(self):
            while self._undo:
                obj, name, old = self._undo.pop()
                setattr(obj, name, old)

    failed = 0
    tests = [(k, v) for k, v in sorted(globals().items())
             if k.startswith("test_") and callable(v)]
    for name, fn in tests:
        mp = _MiniMonkey()
        kwargs = {}
        params = inspect.signature(fn).parameters
        if "monkeypatch" in params:
            kwargs["monkeypatch"] = mp
        hooks = (sys.excepthook, threading.excepthook)
        try:
            fn(**kwargs)
            print("ok   %s" % name)
        except Exception:
            failed += 1
            print("FAIL %s" % name)
            traceback.print_exc()
        finally:
            sys.excepthook, threading.excepthook = hooks
            mp.undo()
    print("%d failed, %d passed" % (failed, len(tests) - failed))
    sys.exit(1 if failed else 0)