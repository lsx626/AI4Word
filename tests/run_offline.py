import sys, traceback, importlib.util, pathlib, inspect, os


class _MiniMonkey:
    """pytest 风格测试（签名带 monkeypatch）的离线替身：setattr/setenv/
    delenv 记录旧值，调用结束统一还原，不让上轮测试的打桩污染同进程
    的后续模块（run_offline 单进程串跑全部 test_*）。"""

    def __init__(self):
        self._undo = []

    def setattr(self, obj, name, value):
        self._undo.append(("attr", obj, name, getattr(obj, name, None)))
        setattr(obj, name, value)

    def setenv(self, name, value):
        self._undo.append(("env", name, os.environ.get(name)))
        os.environ[name] = value

    def delenv(self, name, raising=True):
        self._undo.append(("env", name, os.environ.get(name)))
        os.environ.pop(name, None)

    def undo(self):
        while self._undo:
            kind, a, b, old = self._undo.pop()
            if kind == "env":
                if old is None:
                    os.environ.pop(b, None)
                else:
                    os.environ[b] = old
            else:
                setattr(a, b, old)


root = pathlib.Path(__file__).parent
for f in sorted(root.glob("test_*.py")):
    spec = importlib.util.spec_from_file_location(f.stem, f)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    passed = failed = 0
    for name in dir(mod):
        if name.startswith("test_"):
            fn = getattr(mod, name)
            if callable(fn):
                mp = _MiniMonkey()
                kwargs = {}
                # pytest 风格测试按签名注入（与 test_debug_mode.py 的
                # __main__ 块同一手法），否则裸调用
                try:
                    params = inspect.signature(fn).parameters
                except (TypeError, ValueError):
                    params = {}
                if "monkeypatch" in params:
                    kwargs["monkeypatch"] = mp
                try:
                    fn(**kwargs)
                    passed += 1
                except Exception:
                    failed += 1
                    print(f"FAIL {f.name}::{name}")
                    traceback.print_exc()
                finally:
                    mp.undo()
    print(f"{f.name}: {passed} passed, {failed} failed")
