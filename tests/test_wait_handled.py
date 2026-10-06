# -*- coding: utf-8 -*-
"""离线回归测试：harness 计数（SENT/DONE）与 wait_handled 的语义等待。

背景（V9.5 deep_test 真机日志，doc_persistence 场景）：refresh_map 的
「重连装配完成」检查误报 FAIL，而真实链路 2.2 秒就完成了（handle_start
cmd=refresh_map → load_blocks → word_assembled restored=64）。根因是
harness 计数失同步——deep_test._patch_doc_vars 的合成命令 _read_doc_vars
在计数包装之外直接 return（DONE 不 +1），而 wrapped_send 的排除表里没有
它（SENT 照计），SENT 恒大于 DONE，之后每条 wait_handled 都判失败并烧
满超时。

本测试在纯离线环境（FakeApp Word + 真实 AgentWorker 线程）按
deep_test.main() 的装配顺序（patch_handle → _patch_doc_vars）复现整条
链路：探针 _read_doc_vars 两侧均不计计数、真实命令的 wait_handled 快速
成功、全局 SENT/DONE 保持对称；并验证 wait_handled(cmd=...) 按命令语义
等待不受全局失步影响。
"""
import os
import sys
import time
import traceback

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)

# 先把平台钉成 offscreen 并建好 QApplication，再 import deep_test：
# deep_test 在 import 时会把 QT_QPA_PLATFORM 强制为 windows（真实平台）
# 并往 sys.argv 塞 -debug；本进程的 QApplication 已建成（offscreen），
# import 后立刻还原，run_offline 同进程加载的其它 test_* 模块不受副作用。
os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication  # noqa: E402

_APP = QApplication.instance() or QApplication([])

_argv = list(sys.argv)
import deep_test  # noqa: E402
sys.argv[:] = _argv

import debug_walkthrough as W  # noqa: E402
import test_fake_word as _fake  # noqa: E402
from app.engine import AgentWorker  # noqa: E402


class _AppShell:
    """FakeApp 没有 Documents / ActiveDocument，补上引擎连接所需的壳
    （与 tests/test_engine_offline.py 的 _AppShell 同实现）。"""

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


def _reset_counters():
    """每个测试从零计数（W 的计数器是模块级全局，跨测试会残留）。"""
    W.SENT.update(n=0, per_cmd={})
    W.DONE.update(n=0, per_cmd={})
    deep_test.DOC_VARS.update(req=0, rsp=None, text=None)
    W.DOC_BOX.update(req=0, rsp=0, text=None)


def _start_worker():
    """装配顺序与 deep_test.main() 一致：W.patch_handle（W.assemble 内）
    之后接 deep_test._patch_doc_vars，然后才启动线程。返回 (worker, app)。"""
    app, _writer, _model = _fake.make()
    _writer.write_block("甲段\n\n乙段", animate=False)  # 文档非空 → 装配时导入 2 块
    shell = _AppShell(app)
    worker = AgentWorker(word_factory=lambda: shell)
    worker.set_api_key("test-key")
    worker._attach_factory = lambda: shell  # refresh_map 的只附加路径
    W.patch_handle(worker)              # ← W.assemble() 里的同一行
    deep_test._patch_doc_vars(worker)   # ← deep_test.main() 里的下一行
    _reset_counters()
    worker.start()
    W.wait_for(lambda: worker.isRunning(), 10)
    return worker, app


def _finish(worker):
    worker.send("quit")  # quit 本来就不计入 SENT/DONE
    worker.wait(4000)
    W.pump(0.1)


def test_read_doc_vars_probe_keeps_counters_balanced():
    """根因回归：_read_doc_vars 探针两侧都不计计数；之后真实命令的
    wait_handled 必须快速成功，全局 SENT/DONE 保持对称。"""
    worker, app = _start_worker()
    try:
        worker.send("save")  # 装配 + 把块索引写进文档变量
        assert W.wait_handled(worker, 30), "save 命令未完成"
        assert W.SENT["n"] == W.DONE["n"] == 1, (W.SENT["n"], W.DONE["n"])

        # 1) 探针：在 worker 线程读文档变量（与 scenario_doc_persistence 同法）
        deep_test.DOC_VARS["req"] += 1
        deep_test.DOC_VARS["rsp"] = None
        worker.send("_read_doc_vars", {"req": deep_test.DOC_VARS["req"]})
        assert W.wait_for(
            lambda: deep_test.DOC_VARS["rsp"] == deep_test.DOC_VARS["req"],
            15), "探针未应答"
        stored = deep_test.DOC_VARS["text"]
        assert isinstance(stored, str) and stored.strip().isdigit() \
            and int(stored.strip()) >= 1, repr(stored)
        # 探针不计入全局计数：SENT/DONE 原地不动
        # （修复前这里 SENT+1 而 DONE 不动 → 永久失步）
        assert W.SENT["n"] == W.DONE["n"] == 1, (W.SENT["n"], W.DONE["n"])

        # 2) 抹掉引擎绑定，发 refresh_map（查看类命令触发重连装配）
        worker._model = None
        worker._doc = None
        worker._writer = None
        worker.send("refresh_map")
        t0 = time.time()
        ok = W.wait_handled(worker, 30, cmd="refresh_map")
        elapsed = time.time() - t0
        assert ok, "refresh_map 未完成"
        assert elapsed < 10, "wait_handled 烧满超时 %.1fs" % elapsed
        # 3) 全局计数仍对称（refresh_map 是一条真实命令，两侧各 +1）
        assert W.SENT["n"] == W.DONE["n"] == 2, (W.SENT["n"], W.DONE["n"])
        # 重连装配真的发生：从文档变量恢复出 2 块
        assert worker._model is not None, "重连后块模型未恢复"
        assert len(worker._model.blocks) == 2, \
            [b.text for b in worker._model.blocks]
    finally:
        _finish(worker)
    print("ok: _read_doc_vars 探针对称排除（SENT/DONE 不失步，"
          "wait_handled 不烧超时）")


def test_wait_handled_semantic_ignores_global_desync():
    """按命令语义等待：人为制造全局失步（模拟将来某个忘了对称排除的
    探针命令），全局计数已不可能追平，但 wait_handled(cmd=...) 仍按
    该命令自己的发送/处理计数快速成功。"""
    worker, app = _start_worker()
    try:
        worker.send("save")
        assert W.wait_handled(worker, 30, cmd="save"), "save 命令未完成"
        W.SENT["n"] += 7  # 人为失步：SENT 恒大于 DONE
        worker.send("speed", "slow")
        t0 = time.time()
        ok = W.wait_handled(worker, 30, cmd="speed")
        elapsed = time.time() - t0
        assert ok, "speed 命令未完成"
        assert elapsed < 10, "按命令等待烧满超时 %.1fs" % elapsed
        # 全局模式此时确实追不平（这正是语义等待存在的理由）
        assert not W.wait_handled(worker, 1), "全局失步时全局模式竟成功"
        assert W.SENT["n"] > W.DONE["n"]
    finally:
        _finish(worker)
    print("ok: wait_handled(cmd=...) 按命令语义等待，不受全局失步影响")


def test_sync_doc_text_probe_keeps_counters_balanced():
    """同一类的既有探针 _sync_doc_text（W.doc_text 内部发送）也必须两侧
    对称：探针后发真实命令，全局 wait_handled 仍然成功且 SENT==DONE。"""
    worker, app = _start_worker()
    try:
        worker.send("save")
        assert W.wait_handled(worker, 30), "save 命令未完成"
        text = W.doc_text(worker)  # 内部发 _sync_doc_text 并等 rsp
        assert text and "甲段" in text, repr(text)
        worker.send("speed", "fast")
        t0 = time.time()
        ok = W.wait_handled(worker, 30)
        elapsed = time.time() - t0
        assert ok, "speed 命令未完成"
        assert elapsed < 10, "wait_handled 烧满超时 %.1fs" % elapsed
        assert W.SENT["n"] == W.DONE["n"] == 2, (W.SENT["n"], W.DONE["n"])
    finally:
        _finish(worker)
    print("ok: _sync_doc_text 探针同样对称排除（全局计数平衡）")


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
