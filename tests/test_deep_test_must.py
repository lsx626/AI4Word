# -*- coding: utf-8 -*-
"""回归：deep_test 的 esc_hide_hint MUST 表必须等于场景实际产出。

2026-10-05 真机 slot1 发现的规格错配：deep_test 的 esc_hide_hint MUST 写
hide_via_esc=2，但场景 esc_hide_and_hint 只产生 1 次——第一下 Esc 由展开态
收起（main_window.keyPressEvent 在展开态走 _apply_expanded，不打日志），
第二下 Esc 才隐藏（app/main_window.py 的 keyPressEvent -> debug.log(
"hide_via_esc")，恰好 1 条）。MUST 检查把这条纯规格错配报成
「hide_via_esc(1/2) 缺失」。

本测试从两侧把不变量锁住：
- app 层：复现场景的按键序列，断言 hide_via_esc 恰好 1 条
  （首下 Esc 收起 = 0 条，第二下 Esc 隐藏 = 1 条，summon 不再增加）。
- spec 层：重放 debug_walkthrough.esc_hide_and_hint 真实场景，把产出的事件
  数与 tests/deep_test.py 场景清单里的 MUST 表对比，不一致即失败。

deep_test.py 本身不能 import（导入即锁定 QT_QPA_PLATFORM=windows 并向
sys.argv 追加 -debug，会污染同一进程里的离线套件），MUST 表只做 ast 只读
抽取。
"""
import ast
import os
import re
import sys
import tempfile
import threading

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)

sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEvent, Qt  # noqa: E402
from PySide6.QtGui import QKeyEvent  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from app import debug  # noqa: E402
from app.engine import AgentWorker  # noqa: E402
from app.main_window import MainWindow  # noqa: E402
from app.settings import Settings  # noqa: E402
from app.tray import Tray  # noqa: E402

_app = QApplication.instance() or QApplication(["AI4Word"])

# 与 deep_test._LINE_RE 同格式：时间 [线程] 级别 事件 [字段...]
_LOG_LINE = re.compile(
    r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} \[[^\]]*\] "
    r"(INFO|WARN|ERROR) (\S+)(?:\s+(.*))?$")


class _OfflineDebugLog:
    """把 app.debug 的日志接到本测试的临时文件（用完还原，不碰
    %APPDATA%\\AI4Word\\debug.log）。"""

    def __init__(self):
        self._dir = tempfile.mkdtemp(prefix="ai4word_must_test_")
        self.path = os.path.join(self._dir, "debug.log")

    def __enter__(self):
        self._saved_state = debug._state
        self._saved_hooks = (sys.excepthook, threading.excepthook)
        self._saved_env = os.environ.get("AI4WORD_DEBUG_LOG")
        debug._state = {"debug": None, "path": None, "hooks": False,
                        "failed": False}
        os.environ["AI4WORD_DEBUG_LOG"] = self.path
        debug.init(["must_regression", "-debug"])
        return self

    def __exit__(self, *exc):
        sys.excepthook, threading.excepthook = self._saved_hooks
        debug._state = self._saved_state
        if self._saved_env is None:
            os.environ.pop("AI4WORD_DEBUG_LOG", None)
        else:
            os.environ["AI4WORD_DEBUG_LOG"] = self._saved_env
        return False

    def count(self, event):
        """日志中某事件名的行数（同 deep_test._count 的口径）。"""
        try:
            with open(self.path, encoding="utf-8", errors="replace") as f:
                lines = f.read().splitlines()
        except OSError:
            return 0
        n = 0
        for ln in lines:
            m = _LOG_LINE.match(ln)
            if m and m.group(2) == event:
                n += 1
        return n


def _make_window_tray():
    """与 debug_walkthrough.assemble 同款接线：窗口 + 托盘 + 显示。"""
    settings = Settings(os.path.join(
        tempfile.gettempdir(), "ai4word_test_must_settings.json"))
    settings.load()
    settings.set("geometry", [60, 60])
    settings.set("expanded", False)
    window = MainWindow(AgentWorker(), settings)
    tray = Tray(window, settings)
    window.set_tray(tray)
    tray.show()
    window.show()
    QApplication.processEvents()
    return window, tray


def _tidy(window, tray):
    """收掉窗口与托盘；顺便注销 showEvent 路径可能注册的测试用全局热键。"""
    try:
        window._unregister_hotkey()
    except Exception:
        pass
    try:
        window.hide()
        tray.hide()
    except Exception:
        pass
    window.deleteLater()
    tray.deleteLater()
    QApplication.processEvents()


def _esc_key():
    return QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier)


def test_esc_sequence_emits_exactly_one_hide_via_esc():
    """app 层不变量：展开态首下 Esc 只收起（0 条 hide_via_esc），第二下 Esc
    才隐藏（恰好 1 条），托盘首隐藏提示弹出，summon 不再增加计数。"""
    with _OfflineDebugLog() as log:
        window, tray = _make_window_tray()
        try:
            window.set_expanded(True)
            QApplication.processEvents()
            esc = _esc_key()

            window.keyPressEvent(esc)      # 展开态：收起，不隐藏、不打日志
            assert not window._expanded, "展开态首下 Esc 应收起"
            assert window.isVisible(), "收起不隐藏窗口"
            assert log.count("hide_via_esc") == 0, \
                "收起不应打 hide_via_esc（实际 %d 条）" % log.count(
                    "hide_via_esc")

            window.keyPressEvent(esc)      # 紧凑态：隐藏 + 托盘首次提示
            assert not window.isVisible(), "紧凑态第二下 Esc 应隐藏到托盘"
            assert tray._hinted, "首次隐藏应弹出托盘召回提示"
            assert log.count("hide_via_esc") == 1, \
                "整个序列只应打 1 条 hide_via_esc（实际 %d 条）" % log.count(
                    "hide_via_esc")

            window.summon()                # 召回：不产生新的 hide_via_esc
            assert window.isVisible(), "summon 应重新显示窗口"
            assert log.count("hide_via_esc") == 1, \
                "summon 不应增加 hide_via_esc（实际 %d 条）" % log.count(
                    "hide_via_esc")
        finally:
            _tidy(window, tray)


def _esc_hide_hint_must():
    """ast 只读抽取 tests/deep_test.py 场景清单里 esc_hide_hint 的 MUST 表。"""
    src = os.path.join(_TESTS, "deep_test.py")
    with open(src, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Tuple) and len(node.elts) == 3):
            continue
        name, body, must = node.elts
        if (isinstance(name, ast.Constant) and name.value == "esc_hide_hint"
                and isinstance(body, ast.Lambda) and isinstance(must, ast.Dict)
                and all(isinstance(k, ast.Constant) for k in must.keys)
                and all(isinstance(v, ast.Constant) for v in must.values)):
            return {k.value: v.value
                    for k, v in zip(must.keys, must.values)}
    raise AssertionError("在 %s 中找不到 esc_hide_hint 场景的 MUST 表"
                         "（场景清单结构变了？）" % src)


def test_deep_test_esc_hide_hint_must_matches_scenario():
    """spec 层不变量：deep_test 的 esc_hide_hint MUST 表计数必须等于重放
    debug_walkthrough.esc_hide_and_hint 真实场景产出的 hide_via_esc 数。"""
    import debug_walkthrough as W          # noqa: E402 -- 函数级导入：
    # 避免模块导入时顺带 load_dotenv / win32com 污染离线套件收集阶段
    with _OfflineDebugLog() as log:
        window, tray = _make_window_tray()
        fail0 = len(W.FAIL)
        try:
            W.esc_hide_and_hint(window, tray)
        finally:
            _tidy(window, tray)
    new_fails = W.FAIL[fail0:]
    assert not new_fails, "esc_hide_and_hint 自身 check 失败: %s" % new_fails

    produced = log.count("hide_via_esc")
    must = _esc_hide_hint_must()
    assert must.get("hide_via_esc") == produced, (
        "deep_test.py 的 esc_hide_hint MUST 表要求 hide_via_esc=%s，但场景"
        " esc_hide_and_hint 实际只产生 %d 次（首下 Esc 收起不打日志、第二"
        "下 Esc 才隐藏）——MUST 表必须与场景实际产出一致，改场景或改 MUST"
        " 请同步本测试" % (must.get("hide_via_esc"), produced))


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
