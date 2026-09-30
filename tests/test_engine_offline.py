# -*- coding: utf-8 -*-
"""离线测试：GUI 后台引擎 AgentWorker。

用 tests/test_fake_word.py 的 FakeApp 注入 word_factory，patch
ai_client.ai_stream / ai_client.ai_request（引擎与 agent 都是调用时才
import，patch 直接生效），在 offscreen QApplication 下验证：
- arrange 全流程（生成 -> 执行 -> 空闲）与自我修复（坏码 -> 修正 -> 应用）
- write 流式写入（块流 -> 落盘 -> 空闲）
- write 中断 -> 回滚（快照恢复，后半段文本不入文档）
"""
import importlib.util
import os
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ["QT_QPA_PLATFORM"] = "offscreen"

_spec = importlib.util.spec_from_file_location(
    "fake_word_mod", os.path.join(os.path.dirname(__file__), "test_fake_word.py"))
_fake = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_fake)  # 顺带把 time.sleep 置为 no-op，测试更快

import ai_client
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from app.engine import AgentWorker
_app = QApplication([])


class _AppShell:
    """FakeApp 没有 Documents / ActiveDocument，补上引擎连接所需的壳。"""

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


class _Recorder:
    """主线程收集 worker 信号。"""

    def __init__(self, worker):
        self.messages = []      # [(kind, text)]
        self.states = []
        self.chunks = []
        self.maps = []
        self.interrupted = 0
        self.streams = 0
        worker.message.connect(lambda k, t: self.messages.append((k, t)))
        worker.stateChanged.connect(lambda s: self.states.append(s))
        worker.streamChunk.connect(lambda t: self.chunks.append(t))
        worker.streamStart.connect(lambda: self.__setattr__("streams", self.streams + 1))
        worker.blockMap.connect(lambda t: self.maps.append(t))
        worker.interrupted.connect(lambda: self.__setattr__("interrupted", self.interrupted + 1))

    def texts(self, kind):
        return [t for k, t in self.messages if k == kind]

    def wait_for(self, pred, timeout_ms=10000):
        t0 = time.time()
        while not pred():
            QTest.qWait(20)
            if (time.time() - t0) * 1000 > timeout_ms:
                raise AssertionError("等待超时: " + repr(self.messages[-6:]))


_GATE = threading.Event()


def _streamer(chunks, tail=None):
    """生成一个可在末尾阻塞的假 ai_stream。"""
    def fake(prompt, api_key, system):
        for c in chunks:
            yield c
        if tail is not None:
            _GATE.wait(timeout=8)
            yield tail
    return fake


def _start_worker(rec_cls=None):
    app, _w, _m = _fake.make()
    worker = AgentWorker(word_factory=lambda: _AppShell(app))
    rec = _Recorder(worker)
    worker.start()
    QTest.qWait(100)
    return worker, rec, app


def _finish(worker):
    worker.send("quit")
    worker.wait(4000)


def test_arrange_full_flow():
    ai_client.ai_request = lambda prompt, key, system, model=None: \
        "insert_at_end('排版插入的正文')"
    worker, rec, app = _start_worker()
    try:
        worker.send("arrange", "在末尾加一段")
        rec.wait_for(lambda: any("代码执行完毕" in t for t in rec.texts("info"))
                       and rec.states[-1] == "idle")
        assert any("AI 生成的代码" in t for t in rec.texts("code"))
        assert "排版插入的正文" in app.doc.content
        assert rec.states[-1] == "idle"
        assert "arranging" in rec.states
        assert rec.maps, "应至少广播一次块地图"
    finally:
        _finish(worker)


def test_arrange_self_heal():
    calls = {"n": 0}

    def fake_request(prompt, key, system, model=None):
        calls["n"] += 1
        if calls["n"] == 1:
            return "1 / 0"  # 故意失败
        return "insert_at_end('修复后的正文')"

    ai_client.ai_request = fake_request
    worker, rec, app = _start_worker()
    try:
        worker.send("arrange", "加一段")
        rec.wait_for(lambda: any("修正代码执行完毕" in t for t in rec.texts("info"))
                       and rec.states[-1] == "idle")
        assert "修复后的正文" in app.doc.content
        assert calls["n"] == 2
        assert rec.states[-1] == "idle"
    finally:
        _finish(worker)


def test_write_streams_into_doc():
    ai_client.ai_stream = _streamer(["你好", "，世界"])
    worker, rec, app = _start_worker()
    try:
        worker.send("write", "写一句问候")
        rec.wait_for(lambda: any("已写入 Word" in t for t in rec.texts("info"))
                       and rec.states[-1] == "idle")
        assert "".join(rec.chunks) == "你好，世界"
        assert "你好" in app.doc.content
        assert "writing" in rec.states and rec.states[-1] == "idle"
        assert rec.streams == 1
    finally:
        _finish(worker)


def test_write_interrupt_rollback():
    ai_client.ai_stream = _streamer(["前半段"], tail="后半段")
    worker, rec, app = _start_worker()
    try:
        worker.send("write", "可被打断的生成")
        rec.wait_for(lambda: rec.chunks and rec.chunks[-1] == "前半段")
        worker.interrupt()
        _GATE.set()
        rec.wait_for(lambda: rec.interrupted >= 1)
        worker.choose("rollback")
        rec.wait_for(lambda: any("已回滚" in t for t in rec.texts("info"))
                       and rec.states[-1] == "idle")
        assert "前半段" not in app.doc.content, "回滚后文档不应含本次生成的内容"
        assert "后半段" not in app.doc.content
        assert "".join(rec.chunks) == "前半段"  # 尾部片段在循环顶部被丢弃
        assert rec.states[-1] == "idle"
    finally:
        _GATE.clear()
        _finish(worker)
