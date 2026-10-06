# -*- coding: utf-8 -*-
"""离线测试：GUI 后台引擎 AgentWorker。

用 tests/test_fake_word.py 的 FakeApp 注入 word_factory，patch
ai_client.ai_stream / ai_client.ai_request（引擎与 agent 都是调用时才
import，patch 直接生效），在 offscreen QApplication 下验证：
- arrange 全流程（生成 -> 执行 -> 空闲）与自我修复（坏码 -> 修正 -> 应用，
  且修复成功必须发出 arrange_applied 终态事件）
- write 流式写入（块流 -> 落盘 -> 空闲）
- write 中断 -> 回滚（快照恢复，后半段文本不入文档）
- set_speed 直连路径发出 speed_set 调试事件（无 writer 时也可观测）
"""
import importlib.util
import os
import re
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

from app import debug
from app.engine import AgentWorker
# run_offline 单进程串跑全部 test_*：先到的模块建 QApplication，后来的
# 必须复用单例（否则 libshiboken 抛「先销毁单例」）。idempotent 守卫与
# test_fix_code_sandbox.py / test_wait_handled.py 等同款
_app = QApplication.instance() or QApplication([])


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
    worker.set_api_key("test-key")  # 空密钥守卫会拦截写入流
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


# 与 deep_test._LINE_RE 同一格式：时间戳 [线程] 级别 事件 字段...
_LINE_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} \[[^\]]*\] "
    r"(INFO|WARN|ERROR) (\S+)(?:\s+(.*))?$")

_ARRANGE_TRACE_EVENTS = ("arrange_start", "arrange_applied", "arrange_failed",
                         "run_code_result", "run_code2_result")


def _arrange_trace(log_file):
    """读调试日志，按时间序返回排版链路相关事件 (事件名, 字段串)。"""
    try:
        with open(log_file, encoding="utf-8", errors="replace") as f:
            text = f.read()
    except OSError:
        return []
    out = []
    for ln in text.splitlines():
        m = _LINE_RE.match(ln)
        if m and m.group(2) in _ARRANGE_TRACE_EVENTS:
            out.append((m.group(2), m.group(3) or ""))
    return out


def test_arrange_self_heal_emits_arrange_applied():
    """回归：自我修复成功分支必须发出 arrange_applied 终态事件。

    事故重放（tests/deep_runs 多轮 slot2 的 triage 全报
    arrange_applied(0/1)）：_cmd_arrange 的首次成功分支在 if ok: 内打
    arrange_applied，而自我修复成功分支 if ok2: 执行完全相同的收敛动作
    （消息「修正代码执行完毕，已应用。」/ record_turn / _persist /
    _emit_map / idle）却漏掉该事件：日志里 run_code2_result ok=true 之后
    直达 save_blocks，既无 arrange_applied 也无 arrange_failed，事件日志
    无法判定该排版指令的终态。此前被「修复从未成功」与「同轮后续又走了
    首次成功分支」双重掩盖（DEEP_TEST_WORKFLOW.md 的 arrange_repair、
    tests/experiment_lab.py 的 MUST_EVENTS、tests/deep_test.py 的场景
    MUST 三处都要求 arrange_applied，本测试把它钉进离线套件）。
    """
    log_dir = os.environ.get("TEMP") or os.path.expanduser("~") or "."
    log_file = os.path.join(log_dir, "ai4word_test_arrange_terminal.log")
    saved_state = dict(debug._state)
    saved_env = os.environ.get("AI4WORD_DEBUG_LOG")
    saved_hooks = (sys.excepthook, threading.excepthook)
    try:
        for p in (log_file, log_file + ".bak"):
            try:
                os.remove(p)
            except OSError:
                pass
        os.environ["AI4WORD_DEBUG_LOG"] = log_file
        debug._state.update({"debug": None, "path": None, "hooks": False,
                             "failed": False})
        assert debug.init(["AI4Word", "-debug"]) is True

        calls = {"n": 0}

        def fake_request(prompt, key, system, model=None):
            calls["n"] += 1
            if calls["n"] == 1:
                return "1 / 0"  # 首段代码必失败 -> 进入自我修复分支
            return "insert_at_end('修复后的正文')"  # 修复代码合规且可执行

        ai_client.ai_request = fake_request
        worker, rec, app = _start_worker()
        try:
            worker.send("arrange", "加一段")
            rec.wait_for(lambda: any("修正代码执行完毕" in t for t in rec.texts("info"))
                           and rec.states[-1] == "idle")
            assert "修复后的正文" in app.doc.content
            assert calls["n"] == 2
        finally:
            _finish(worker)  # 线程退出后日志已全部落盘，再读不竞态

        trace = _arrange_trace(log_file)
        names = [ev for ev, _ in trace]
        assert names.count("arrange_start") == 1, names
        # 确实走了自我修复分支：第一轮执行失败、修复代码第二轮执行成功
        first = [f for ev, f in trace if ev == "run_code_result"]
        second = [f for ev, f in trace if ev == "run_code2_result"]
        assert first and re.search(r"\bok=false\b", first[0]), first
        assert second and re.search(r"\bok=true\b", second[0]), second
        # 终态事件：run_code2_result ok=true 之后紧接着必须就是
        # arrange_applied（真机日志里缺的正是这一行：之后直达 save_blocks），
        # 且整条指令不得出现 arrange_failed
        idx = names.index("run_code2_result")
        assert idx + 1 < len(trace) and trace[idx + 1][0] == "arrange_applied", (
            "run_code2_result ok=true 之后缺少 arrange_applied 终态事件：" + str(trace))
        assert "arrange_failed" not in names, names
    finally:
        debug._state.clear()
        debug._state.update(saved_state)
        if saved_env is None:
            os.environ.pop("AI4WORD_DEBUG_LOG", None)
        else:
            os.environ["AI4WORD_DEBUG_LOG"] = saved_env
        sys.excepthook, threading.excepthook = saved_hooks
        for p in (log_file, log_file + ".bak"):
            try:
                os.remove(p)
            except OSError:
                pass
    print("ok: 自我修复成功分支在 run_code2_result ok=true 后发出 arrange_applied 终态事件")


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


class _SpeedSpy:
    """假 writer：记录 set_speed 转发（无 COM 语义，仅验证调用）。"""

    def __init__(self):
        self.calls = []

    def set_speed(self, mode):
        self.calls.append(mode)


def _speed_set_events(log_file):
    """读日志，返回 speed_set 事件行的 mode 字段列表（按时间序）。"""
    try:
        with open(log_file, encoding="utf-8") as f:
            text = f.read()
    except OSError:
        return []
    modes = []
    for ln in text.splitlines():
        if not re.search(r"\bINFO speed_set( |$)", ln):
            continue
        m = re.search(r"\bmode=(auto|slow|fast)\b", ln)
        if m:
            modes.append(m.group(1))
    return modes


def test_set_speed_logs_event_without_writer():
    """回归：UI 直连换档（AgentWorker.set_speed）必须发出 speed_set 事件。

    真机场景 speed_gears 曾出现 speed_set(0/3)：直连路径不写日志，而
    writer 侧的 set_speed 日志在尚未连接 Word（_writer 为 None）时又不
    会发出，档位变更完全不可观测。要求与队列路径（_set_speed）事件名
    一致，且无论 writer 是否存在都可见。
    """
    log_dir = os.environ.get("TEMP") or os.path.expanduser("~") or "."
    log_file = os.path.join(log_dir, "ai4word_test_engine_speed.log")
    saved_state = dict(debug._state)
    saved_env = os.environ.get("AI4WORD_DEBUG_LOG")
    saved_hooks = (sys.excepthook, threading.excepthook)
    try:
        for p in (log_file, log_file + ".bak"):
            try:
                os.remove(p)
            except OSError:
                pass
        os.environ["AI4WORD_DEBUG_LOG"] = log_file
        debug._state.update({"debug": None, "path": None, "hooks": False,
                             "failed": False})
        assert debug.init(["AI4Word", "-debug"]) is True

        app, _w, _m = _fake.make()
        worker = AgentWorker(word_factory=lambda: _AppShell(app))
        worker.set_api_key("test-key")
        assert worker._writer is None  # 尚未连接 Word：UI 换档早于建表

        for mode in ("slow", "auto", "fast"):
            worker.set_speed(mode)
            assert worker._speed == mode
        worker.set_speed("bogus")  # 非法档位：属性不变、也不发事件
        assert worker._speed == "fast"

        # 有 writer 时：事件照常发出，且仍转发给 writer
        spy = _SpeedSpy()
        worker._writer = spy
        worker.set_speed("slow")
        assert spy.calls == ["slow"]

        assert _speed_set_events(log_file) == ["slow", "auto", "fast", "slow"]
    finally:
        debug._state.clear()
        debug._state.update(saved_state)
        if saved_env is None:
            os.environ.pop("AI4WORD_DEBUG_LOG", None)
        else:
            os.environ["AI4WORD_DEBUG_LOG"] = saved_env
        sys.excepthook, threading.excepthook = saved_hooks
        for p in (log_file, log_file + ".bak"):
            try:
                os.remove(p)
            except OSError:
                pass
