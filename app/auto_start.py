# -*- coding: utf-8 -*-
"""开机自启动：写 / 删 HKCU Run 键。"""
import os
import sys

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
        return True
    except OSError:
        return False


def disable():
    if winreg is None:
        return True
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, VALUE_NAME)
    except OSError:
        pass
    return True
