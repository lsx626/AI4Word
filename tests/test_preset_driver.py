# -*- coding: utf-8 -*-
"""离线回归：预设驱动的选中路径必须在两种绑定下都投递恰好一条命令。

事故重放（deep_test review_presets_save）：场景驱动用
`preset_combo.setCurrentIndex(idx)`，而 MainWindow 曾把 _on_preset 连在
QComboBox.activated 上——该信号只在用户真实弹窗/键盘选择时触发，
setCurrentIndex 永远不发它（PySide6 6.11.0 实测确认），_on_preset 被
静默跳过：组合框不复位、preset 命令不投递，四个预设全部报
preset_selected/preset_applied(0/4)，wait_handled 空转并把缺事件掩盖成
「命令未处理完成」。

修复后的契约（本文件钉死）：
1. 驱动侧（deep_test 场景与压测同款代码）：setCurrentIndex 之后追补
   activated.emit(idx)——当前绑定（currentIndexChanged）下 activated 无
   接收端、且 emit 不改变 currentIndex，不产生第二条命令；万一绑定被
   改回 activated-only，驱动仍照常选中；
2. 应用侧：四个预设每个恰好投递一条 preset 命令、组合框复位到占位项、
   引擎处理完打出 preset_applied（即 deep_test MUST 里 4/4 的来源）；
3. 重复选同一项（currentIndex 未变、setCurrentIndex 无信号）也必须
   投递——activated.emit 兜底，wait_handled 不烧超时。
"""
import os
import sys
import tempfile
import traceback

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication  # noqa: E402

_APP = QApplication.instance() or QApplication([])

_argv = list(sys.argv)
import debug_walkthrough as W  # noqa: E402
import test_event_wiring as _ew  # noqa: E402  复用同一份调试日志捕获
import test_fake_word as _fake  # noqa: E402
from app.engine import AgentWorker  # noqa: E402
from app.main_window import MainWindow  # noqa: E402
from app.settings import Settings  # noqa: E402

# deep_test 在 import 时强制 windows 平台 + 往 sys.argv 塞 -debug；
# 驱动代码只需要它的场景函数本体，import 完立刻还原本进程环境
import deep_test  # noqa: E402
sys.argv[:] = _argv

_PRESETS = ("论文", "公文", "简历", "博客")


class _AppShell:
    """FakeApp 没有 Documents / ActiveDocument，补上引擎连接所需的壳
    （与 tests/test_wait_handled.py 的 _AppShell 同实现）。"""

    def __init__(self, app):
        self._app = app

    @property
    def Documents(self):
        return self

    @property
    def Count(self):
        return 1

    def Add(self):
        return self._app.doc

    @property
    def ActiveDocument(self):
        return self._app.doc

    @property
    def Selection(self):
        return self._app.Selection

    def __getattr__(self, name):
        return getattr(self._app, name)


def _start_window_with_worker():
    """真实 MainWindow + 真实 worker 线程 + FakeApp Word（全离线）。"""
    app, _writer, _model = _fake.make()
    _writer.write_block("甲段\n\n乙段", animate=False)  # 文档非空，预设才有内容可改
    shell = _AppShell(app)
    worker = AgentWorker(word_factory=lambda: shell)
    worker.set_api_key("test-key")
    W.patch_handle(worker)
    W.SENT.update(n=0, per_cmd={})
    W.DONE.update(n=0, per_cmd={})
    worker.start()
    W.wait_for(lambda: worker.isRunning(), 10)

    s = Settings(os.path.join(tempfile.mkdtemp(prefix="ai4word_pd_"),
                              "settings.json"))
    s.load()
    s.set("geometry", [60, 60])
    s.set("expanded", False)
    window = MainWindow(worker, s)
    W.pump(0.1)
    # 构造时投递的首条 speed 命令不计入本测试的 preset 计数；清零对齐
    W.SENT.update(n=0, per_cmd={})
    W.DONE.update(n=0, per_cmd={})
    return window, worker, app


def _finish(worker, window):
    worker.send("quit")
    worker.wait(4000)
    if window is not None:
        window.close()
    W.pump(0.1)


def test_preset_driver_queues_exactly_one_command_per_preset():
    """驱动 + 应用的完整链路：每预设一条 preset 命令、复位、preset_applied。

    用 deep_test 场景的同款驱动代码（setCurrentIndex + activated.emit）
    在 FakeApp Word 上真跑，断言 deep_test MUST 表的来源事件数（4 条
    preset_selected / 4 条 preset_applied）。
    """
    window, worker, app = _start_window_with_worker()
    try:
        worker.send("save")  # 先装配（与场景顺序一致：文档有块才应用预设）
        assert W.wait_handled(worker, 30, cmd="save"), "save 未完成"
        assert worker._doc is not None, "Word 未连接，preset 命令无处落地"
        W.pump(0.1)
        W.SENT.update(n=0, per_cmd={})
        W.DONE.update(n=0, per_cmd={})

        with _ew._debug_capture() as events:
            sent_before = 0
            for name in _PRESETS:
                idx = window.preset_combo.findText(name)
                assert idx > 0, "预设「%s」不在组合框中" % name
                # deep_test 场景的同款驱动代码
                window.preset_combo.setCurrentIndex(idx)
                window.preset_combo.activated.emit(idx)
                assert window.preset_combo.currentIndex() == 0, \
                    "应用后组合框应复位到占位项（%s）" % name
                assert W.wait_handled(worker, 30, cmd="preset"), \
                    "预设「%s」命令未处理完成" % name
                # 计数是全局累计，用增量断言：本步恰好 +1（不多投、不漏投）
                sent_now = W.SENT["per_cmd"].get("preset", 0)
                done_now = W.DONE["per_cmd"].get("preset", 0)
                assert sent_now == sent_before + 1, \
                    "单个预设应只投递一条命令（累计 %d -> %d）" % (
                        sent_before, sent_now)
                assert done_now == sent_now, \
                    "preset 命令发送与处理计数不对称：sent=%d done=%d" % (
                        sent_now, done_now)
                sent_before = sent_now
                W.pump(0.1)

        assert events.count("preset_selected") == 4, \
            "preset_selected 应为 4（每个预设一次选择）：%s" % events
        assert events.count("preset_applied") == 4, \
            "preset_applied 应为 4（deep_test MUST 的来源）：%s" % events
        assert W.SENT["per_cmd"]["preset"] == 4, W.SENT["per_cmd"]
        assert W.DONE["per_cmd"]["preset"] == 4, W.DONE["per_cmd"]
        # 驱动全程恰好四条 preset 命令，无重复投递（activated 不触发二次）
        assert W.SENT["n"] == W.DONE["n"] == 4, (W.SENT["n"], W.DONE["n"])
    finally:
        _finish(worker, window)
    print("ok: 预设驱动每项恰好一条 preset 命令，复位 + preset_applied×4")


def test_preset_driver_reselecting_same_preset_still_sends():
    """currentIndex 未变时 setCurrentIndex 无信号：activated.emit 兜底，
    驱动第二次选同一预设仍投递一条命令（压测场景可能随机重复选同一项）。"""
    window, worker, app = _start_window_with_worker()
    try:
        worker.send("save")
        assert W.wait_handled(worker, 30, cmd="save"), "save 未完成"
        W.pump(0.1)
        W.SENT.update(n=0, per_cmd={})
        W.DONE.update(n=0, per_cmd={})

        name = "公文"
        idx = window.preset_combo.findText(name)
        window.preset_combo.setCurrentIndex(idx)        # 选中
        assert W.wait_handled(worker, 30, cmd="preset")  # 复位回 0
        W.pump(0.1)
        W.SENT.update(n=0, per_cmd={})
        W.DONE.update(n=0, per_cmd={})

        window.preset_combo.setCurrentIndex(idx)        # 再次选中同一项
        window.preset_combo.activated.emit(idx)         # 强信号兜底
        ok = W.wait_handled(worker, 30, cmd="preset")
        assert ok, "重复选同一预设时命令未投递/未处理（wait_handled 烧超时）"
        assert W.SENT["per_cmd"].get("preset", 0) == 1, W.SENT["per_cmd"]
    finally:
        _finish(worker, window)
    print("ok: 重复选同一预设仍投递一条命令（activated 兜底，不依赖状态变化）")


if __name__ == "__main__":
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
