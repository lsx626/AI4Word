# -*- coding: utf-8 -*-
"""AI4Word 启动入口（开发态与打包后共用）。

不在 app.__main__ 里直接跑，是为了在这里兜底未捕获异常：窗口化打包
应用的异常不会打印到任何控制台，统一写入 crash.log 便于排障。
"""
import os
import sys
import traceback


def _crash_log():
    try:
        from app import debug
        debug.exc("crash")
    except Exception:
        pass
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "AI4Word", "crash.log")
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(traceback.format_exc() + "\n")
    except Exception:
        path = os.path.join(os.environ.get("TEMP") or ".", "ai4word_crash.log")
        try:
            with open(path, "a", encoding="utf-8") as f:
                f.write(traceback.format_exc() + "\n")
        except Exception:
            pass


def _strip_alien_qt_dirs():
    """打包后防止 PATH 里的外来 Qt6*.dll（如 Anaconda 自带的旧 Qt）串扰。

    PyQt 的插件加载走 LoadLibrary 的默认搜索（含 PATH），Anaconda 等
    环境的旧 Qt6Core.dll 会先被命中，报“找不到指定的程序”。处理：
    把捆绑的 _internal\PySide6 注册为 DLL 目录并提到 PATH 最前，
    同时从 PATH 剔除其他含 Qt6*.dll 的目录。
    """
    if not getattr(sys, "frozen", False):
        return
    base = os.path.dirname(sys.executable)
    internal = os.path.join(base, "_internal")
    bundled = [d for d in (os.path.join(internal, "PySide6"), internal, base)
               if os.path.isdir(d)]
    for d in bundled:
        try:
            os.add_dll_directory(d)
        except Exception:
            pass

    def has_qt(d):
        try:
            return any(f.lower().startswith("qt6") and f.lower().endswith(".dll")
                       for f in os.listdir(d))
        except OSError:
            return False

    real_bundled = {os.path.realpath(d).lower() for d in bundled}
    kept = []
    for d in os.environ.get("PATH", "").split(os.pathsep):
        if not d:
            continue
        try:
            if os.path.realpath(d).lower() in real_bundled or not has_qt(d):
                kept.append(d)
        except OSError:
            kept.append(d)
    os.environ["PATH"] = os.pathsep.join(bundled + kept)


def main():
    _strip_alien_qt_dirs()
    from app import debug
    debug.init()
    debug.log("startup", entry="ai4word.pyw", argv=list(sys.argv),
              log_path=debug.log_path())
    from app.__main__ import main as _run
    sys.exit(_run(sys.argv))


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException:
        _crash_log()
        raise
