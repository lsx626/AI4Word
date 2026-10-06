# -*- coding: utf-8 -*-
"""开机自启动：写 / 删 HKCU Run 键。"""
import os
import sys

from app import debug

try:
    import winreg
except ImportError:  # 非 Windows：安全降级
    winreg = None

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
VALUE_NAME = "AI4Word"


def launcher_path():
    """开机要执行的命令行。打包后是 exe；开发态是 pythonw + 启动脚本。"""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    here = os.path.dirname(os.path.abspath(__file__))
    launcher = os.path.normpath(os.path.join(here, "..", "ai4word.pyw"))
    return f'"{sys.executable}" "{launcher}"'


def is_enabled():
    if winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_READ) as k:
            winreg.QueryValueEx(k, VALUE_NAME)
        return True
    except OSError:
        return False
    except Exception:
        return False


def enable():
    if winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, VALUE_NAME, 0, winreg.REG_SZ, launcher_path())
        debug.log("autostart_enabled", path=launcher_path())
        return True
    except OSError as e:
        debug.warn("autostart_enable_failed", error=str(e)[:200])
        return False


def disable():
    if winreg is None:
        return True
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, VALUE_NAME)
        debug.log("autostart_disabled")
    except FileNotFoundError:
        # 值（或 Run 键）本来就不存在：取消自启是 no-op 成功，不是失败。
        # 过去它与权限拒绝等真实失败一并被 OSError 捕获打成
        # autostart_disable_failed（真机日志 21:38:08 两条「待确认告警」
        # 即此噪声），噪声告警与真实写入失败混在一起，排障无从分辨。
        debug.log("autostart_disabled", note="value_absent")
    except OSError as e:
        debug.warn("autostart_disable_failed", error=str(e)[:200])
    return True
