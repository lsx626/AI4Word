# -*- coding: utf-8 -*-
"""离线测试：设置对话框生命周期事件接线（settings_open / settings_closed）。

回归缺口：settings_open 由 SettingsDialog.__init__ 打，而 settings_closed
只挂在 MainWindow._open_settings 的包装路径里（dlg.exec() 之后）。一旦调用方
直接构造 SettingsDialog 并 accept/reject（调试演练场景就是如此），日志里
settings_open 有、settings_closed 永远没有——单路径事件缺口。

现在 SettingsDialog.done() 是统一收口：accept / reject / Esc / 点 X 都发一次
settings_closed；accept 在保存失败时提前 return（done 不被调用），对话框没关
就不报「已关闭」。本文件同时验证「直接构造」与「经 MainWindow._open_settings
真实入口」两条路径的事件计数都恰好为 1。
"""
import contextlib
import os
import shutil
import sys
import tempfile
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QTimer  # noqa: E402
from PySide6.QtWidgets import (QApplication, QDialog, QMessageBox,
                               QWidget)  # noqa: E402

import app.settings_dialog as sd  # noqa: E402
from app import debug  # noqa: E402
from app.engine import AgentWorker  # noqa: E402
from app.main_window import MainWindow  # noqa: E402
from app.settings import Settings  # noqa: E402
from app.settings_dialog import SettingsDialog  # noqa: E402

_app = QApplication.instance() or QApplication(sys.argv)


def _settings(tag="sd_events"):
    path = os.path.join(os.environ.get("TEMP") or ".",
                        "ai4word_test_%s.json" % tag)
    s = Settings(path)
    s.load()
    return s


def _line(text, event):
    """取包含该事件名的第一行（已含级别与前缀）。"""
    lines = [ln for ln in text.splitlines() if (" " + event + " ") in ln
             or ln.endswith(" " + event)]
    assert lines, "日志里没有事件 %r" % event
    return lines[0]


def _count(text, event):
    return sum(1 for ln in text.splitlines()
               if (" " + event + " ") in ln or ln.endswith(" " + event))


@contextlib.contextmanager
def _debug_log():
    """临时启用调试模式并指向独立日志文件；退出时完整还原全局状态。

    走 app.debug 的真实写入路径（而不是 mock 掉 debug.log）：堵死
    「发射点在、但 debug.log 是空操作 / 写不进文件」这一类接线缺口。
    """
    tmp = tempfile.mkdtemp(prefix="ai4word_sd_test_")
    path = os.path.join(tmp, "debug.log")
    saved_state = dict(debug._state)
    saved_log_dir = debug._log_dir
    saved_excepthook = sys.excepthook
    saved_thread_hook = threading.excepthook
    saved_env = dict(os.environ)
    debug._state.clear()
    debug._state.update({"debug": None, "path": None, "hooks": False,
                         "failed": False})
    debug._log_dir = lambda: tmp
    try:
        debug.init(["AI4Word.exe", "-debug"])
        yield path
    finally:
        debug._state.clear()
        debug._state.update(saved_state)
        debug._log_dir = saved_log_dir
        sys.excepthook = saved_excepthook
        threading.excepthook = saved_thread_hook
        os.environ.clear()
        os.environ.update(saved_env)
        shutil.rmtree(tmp, ignore_errors=True)


@contextlib.contextmanager
def _no_side_effects():
    """拦截开机自启注册表写入、系统警告框与 apply_env 的环境变量外泄。"""
    real = (sd.is_enabled, sd.enable, sd.disable)
    warns = []
    sd.is_enabled = lambda: False
    sd.enable = lambda: True
    sd.disable = lambda: None
    orig_warning = QMessageBox.warning

    def fake_warning(*args, **kwargs):
        warns.append(args)
        return QMessageBox.Ok

    QMessageBox.warning = fake_warning
    try:
        yield warns
    finally:
        sd.is_enabled, sd.enable, sd.disable = real
        QMessageBox.warning = orig_warning


def _visible(cls, parent):
    for w in parent.findChildren(QWidget):
        if isinstance(w, cls) and w.isVisible():
            return w
    return None


def test_accept_emits_closed_once():
    with _debug_log() as path, _no_side_effects():
        s = _settings()
        s.set("api_key", "")
        dlg = SettingsDialog(s, AgentWorker())
        dlg.key_edit.setText("sk-abcdefgh1234567890")
        dlg.accept()
        text = open(path, encoding="utf-8").read()
        assert _count(text, "settings_open") == 1, "对话框打开事件恰好一次"
        assert _count(text, "settings_closed") == 1, "关闭事件恰好一次"
        assert "accepted=true" in _line(text, "settings_closed")
        assert int(dlg.result()) == int(QDialog.Accepted)
        assert s.get("api_key") == "sk-abcdefgh1234567890"


def test_reject_emits_closed_but_no_accept():
    with _debug_log() as path, _no_side_effects():
        s = _settings()
        s.set("api_key", "sk-keep-me123456")
        dlg = SettingsDialog(s, AgentWorker())
        dlg.reject()
        text = open(path, encoding="utf-8").read()
        assert _count(text, "settings_open") == 1
        assert _count(text, "settings_closed") == 1
        assert "accepted=false" in _line(text, "settings_closed")
        assert _count(text, "settings_accept") == 0, "取消不应提交任何修改"
        assert int(dlg.result()) == int(QDialog.Rejected)
        assert s.get("api_key") == "sk-keep-me123456"


def test_save_failure_does_not_report_closed():
    """保存失败时对话框不关闭：不能发 settings_closed。修好路径后仍只发一次。"""
    with _debug_log() as path, _no_side_effects():
        s = _settings()
        s.set("api_key", "")
        bad = os.path.join(os.path.dirname(s.path), "no_such_dir_ai4word",
                           "settings.json")
        saved_path = s.path
        dlg = SettingsDialog(s, AgentWorker())
        dlg.key_edit.setText("sk-first-try12345")
        s.path = bad
        dlg.accept()                       # 保存失败 -> 提前 return
        s.path = saved_path
        dlg.accept()                       # 路径恢复后真正关闭
        text = open(path, encoding="utf-8").read()
        assert _count(text, "settings_save_failed") >= 1
        assert int(dlg.result()) == int(QDialog.Accepted)
        # 关键：对话框只关了一次，closed 也只许有一次
        assert _count(text, "settings_closed") == 1, \
            "保存失败未关闭不能发 closed；最终关闭只发一次"
        assert "accepted=true" in _line(text, "settings_closed")


def test_main_window_entry_emits_pair_exactly_once():
    """真实入口（托盘/按钮 -> _open_settings）仍成对、且各只一次。

    直接构造对话框的调用方与包装路径共用同一份（对话框内的）接线，
    是本组修复的核心断言。exec() 是模态阻塞调用，必须在它的局部事件
    循环里用定时器把对话框确定掉，否则单测挂死。
    """
    with _debug_log() as path, _no_side_effects():
        s = _settings()
        s.set("geometry", [60, 60])
        s.set("expanded", False)
        win = MainWindow(AgentWorker(), s)
        fired = []

        def _auto_ok():
            dlg = _visible(SettingsDialog, win)
            if dlg is None:
                return
            fired.append(True)
            dlg.key_edit.setText("sk-from-entry987")
            dlg.accept()

        def _watchdog():
            if fired:                     # exec() 已返回，无需兜底
                return
            dlg = _visible(SettingsDialog, win)
            if dlg is not None:
                dlg.reject()              # 兜底：防止 accept 没生效导致挂死

        QTimer.singleShot(0, _auto_ok)
        QTimer.singleShot(3000, _watchdog)
        win._open_settings()
        text = open(path, encoding="utf-8").read()
        assert fired, "模态对话框在 exec() 期间被找到并确定"
        assert _count(text, "settings_open") == 1, \
            "包装路径不应再各打一遍 settings_open"
        assert _count(text, "settings_closed") == 1, \
            "包装路径不应再各打一遍 settings_closed"
        assert "accepted=true" in _line(text, "settings_closed")
        assert s.get("api_key") == "sk-from-entry987"


if __name__ == "__main__":
    import inspect
    import traceback

    failed = 0
    tests = [(k, v) for k, v in sorted(globals().items())
             if k.startswith("test_") and callable(v)]
    for name, fn in tests:
        try:
            fn()
            print("ok   %s" % name)
        except Exception:
            failed += 1
            print("FAIL %s" % name)
            traceback.print_exc()
    print("%d failed, %d passed" % (failed, len(tests) - failed))
    sys.exit(1 if failed else 0)
