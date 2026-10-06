# -*- coding: utf-8 -*-
"""离线回归：开机自启 disable() 的「无值删除」语义与设置面板取消勾选。

禁止项不存在时，取消自启是 no-op 成功（不是失败）——把 [WinError 2]
当告警会与真实写入失败（杀软 / 组策略 / 权限）混为一谈。锁定：
disable() 遇到不存在的值静默成功、真实失败仍走 autostart_disable_failed；
accept() 未勾选时先问 is_enabled()，不以 KEY_SET_VALUE 打开注册表。

注册表副作用隔离：RUN_KEY / VALUE_NAME 指向 HKCU\\Software\\
AI4WordTest\\Run 下的临时键，退出时整体删除，不碰真实 Run 键。
"""
import contextlib
import os
import shutil
import sys
import tempfile
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    import winreg
except ImportError:  # 非 Windows：auto_start 本就降级，无注册表可测
    winreg = None

if winreg is None:  # pragma: no cover
    if __name__ == "__main__":
        print("skip: 非 Windows 平台，无注册表语义可测")
    else:
        import pytest
        pytest.skip("winreg unavailable", allow_module_level=True)

from PySide6.QtWidgets import QApplication  # noqa: E402

import app.auto_start as auto_start  # noqa: E402
from app import debug  # noqa: E402
from app.settings import Settings  # noqa: E402
from app.settings_dialog import SettingsDialog  # noqa: E402

_app = QApplication.instance() or QApplication(sys.argv)

SCRATCH_KEY = r"Software\AI4WordTest\Run"
SCRATCH_VALUE = "AI4WordTest"


@contextlib.contextmanager
def _debug_log():
    """临时启用调试模式并指向独立日志；退出时完整还原全局状态。

    走 app.debug 的真实写入路径而不是 mock 掉 debug.log，堵死
    「发射点在、但日志写不进文件」这一类接线缺口。
    """
    tmp = tempfile.mkdtemp(prefix="ai4word_as_test_")
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


def _count(text, event):
    return sum(1 for ln in text.splitlines()
               if (" " + event + " ") in ln or ln.endswith(" " + event))


@contextlib.contextmanager
def _scratch_reg(create_key=True):
    """把 auto_start 指向临时注册表键，退出时删除（含残留值）。"""
    saved = (auto_start.RUN_KEY, auto_start.VALUE_NAME)
    auto_start.RUN_KEY = SCRATCH_KEY
    auto_start.VALUE_NAME = SCRATCH_VALUE
    _drop_key()
    if create_key:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, SCRATCH_KEY):
            pass
    try:
        yield
    finally:
        _drop_key()
        auto_start.RUN_KEY, auto_start.VALUE_NAME = saved


def _drop_key():
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, SCRATCH_KEY)
    except OSError:
        pass


def _value_present():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SCRATCH_KEY, 0,
                            winreg.KEY_READ) as k:
            winreg.QueryValueEx(k, SCRATCH_VALUE)
        return True
    except OSError:
        return False


def test_disable_without_value_is_silent_noop():
    """Run 键在、值不在：disable() 返回 True 且不打 autostart_disable_failed。

    这就是真机日志里两条噪声告警的复现场景（未勾自启 + 本机无 Run 值）。
    """
    with _debug_log() as log, _scratch_reg():
        assert not _value_present()
        ok = auto_start.disable()
        text = open(log, encoding="utf-8").read()
    assert ok is True, "值不存在时取消自启应成功（no-op）"
    assert _count(text, "autostart_disable_failed") == 0, \
        "值不存在不是失败，不应打 autostart_disable_failed：" + text
    assert _count(text, "autostart_disabled") == 1
    assert "value_absent" in text, "应标注本次是无值 no-op，便于排障区分"
    assert not _value_present()


def test_enable_disable_roundtrip_on_scratch_key():
    """enable -> is_enabled -> disable 全链路在临时键上走真实注册表。"""
    with _debug_log() as log, _scratch_reg():
        assert auto_start.enable() is True
        assert _value_present()
        assert auto_start.is_enabled() is True
        assert auto_start.disable() is True
        assert not _value_present()
        assert auto_start.is_enabled() is False
        text = open(log, encoding="utf-8").read()
    assert _count(text, "autostart_enabled") == 1
    assert _count(text, "autostart_disabled") == 1
    assert _count(text, "autostart_disable_failed") == 0


def test_disable_with_run_key_missing_always_succeeds():
    """Run 键整体不存在（OpenKey 即抛 [WinError 2]）同样是 no-op 成功。"""
    with _debug_log() as log, _scratch_reg(create_key=False):
        ok = auto_start.disable()
        text = open(log, encoding="utf-8").read()
    assert ok is True
    assert _count(text, "autostart_disable_failed") == 0, \
        "Run 键缺失也不是失败告警：" + text
    assert _count(text, "autostart_disabled") == 1


def test_disable_real_registry_failure_still_warns():
    """真实失败（权限拒绝）必须保留告警：不能因为加了静默分支而吞掉。"""
    saved_mod = auto_start.winreg

    class _BlockedReg:
        """模拟 KEY_SET_VALUE 被拒绝（OSError 但非文件不存在）。"""

        HKEY_CURRENT_USER = winreg.HKEY_CURRENT_USER
        KEY_SET_VALUE = winreg.KEY_SET_VALUE

        def OpenKey(self, *args, **kwargs):
            raise PermissionError(5, "拒绝访问")

    auto_start.winreg = _BlockedReg()
    try:
        with _debug_log() as log:
            ok = auto_start.disable()
            text = open(log, encoding="utf-8").read()
    finally:
        auto_start.winreg = saved_mod
    assert ok is True
    assert _count(text, "autostart_disable_failed") == 1, \
        "真实注册表失败必须告警：" + text
    assert "拒绝访问" in text


def test_settings_accept_unchecked_autostart_does_not_warn():
    """SettingsDialog.accept 未勾自启的真实路径：不再产出噪声告警。

    不 stub auto_start（区别于 test_settings_dialog.py 的 _no_side_effects
    桩），走 disable() 的真实分支——本组修复的落点。
    """
    settings_path = os.path.join(tempfile.mkdtemp(prefix="ai4word_as_cfg_"),
                                 "settings.json")
    s = Settings(settings_path)
    s.load()
    s.set("api_key", "")
    with _debug_log() as log, _scratch_reg():
        assert not auto_start.is_enabled()
        dlg = SettingsDialog(s, worker=None)
        assert not dlg.auto_check.isChecked(), \
            "无 Run 值时复选框应初始为未勾选"
        dlg.key_edit.setText("sk-unchecked-autostart")
        dlg.accept()
        text = open(log, encoding="utf-8").read()
        assert int(dlg.result()) == 1  # QDialog.Accepted
        assert s.get("auto_start") is False
        assert not auto_start.is_enabled()
    assert _count(text, "autostart_disable_failed") == 0, \
        "accept 未勾自启 + 无 Run 值不应告警：" + text
    assert _count(text, "settings_accept") == 1
    try:
        shutil.rmtree(os.path.dirname(settings_path), ignore_errors=True)
    except OSError:
        pass


if __name__ == "__main__":
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
