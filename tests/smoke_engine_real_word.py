# -*- coding: utf-8 -*-
"""真实 Word × GUI 引擎冒烟：AgentWorker 线程里的 COM 连接与流式写入。

不调用 Atria API（ai_stream/ai_request 打桩），只验证：
- worker 线程 CoInitialize 后能连上真实 Word（与 GUI 运行时同一代码路径）
- write 流式落盘到真实文档
- arrange（生成代码 -> run_code -> 应用）到真实文档
- write 中断 -> 回滚在真实文档上成立

需要本机装有 Word。在新建的临时文档上操作，结束关闭不保存。
"""
import os
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ["QT_QPA_PLATFORM"] = "offscreen"

import win32com.client

import ai_client
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from app.agent import get_word as _unused  # noqa: F401  (确认可导入)
from app.engine import AgentWorker

_app = QApplication([])
_GATE = threading.Event()


class _Rec:
    def __init__(self, w):
        self.msgs, self.states, self.chunks = [], [], []
        w.message.connect(lambda k, t: self.msgs.append((k, t)))
        w.stateChanged.connect(lambda s: self.states.append(s))
        w.streamChunk.connect(lambda t: self.chunks.append(t))

    def texts(self, kind):
        return [t for k, t in self.msgs if k == kind]

    def wait_for(self, pred, timeout_ms=30000):
        t0 = time.time()
        while not pred():
            QTest.qWait(20)
            if (time.time() - t0) * 1000 > timeout_ms:
                raise AssertionError("超时: " + repr(self.msgs[-8:]))


def _word():
    """在 worker 线程内创建/连接 Word（COM 对象必须在调用线程内取得）。"""
    return win32com.client.Dispatch("Word.Application")


def _verify_content():
    """主线程用独立连接读取当前文档文本（不跨线程共享接口代理）。"""
    v = win32com.client.Dispatch("Word.Application")
    if v.Documents.Count == 0:
        return ""
    return v.ActiveDocument.Content.Text


def _close_word():
    v = win32com.client.Dispatch("Word.Application")
    for _ in range(40):
        try:
            while v.Documents.Count > 0:
                v.Documents(1).Close(SaveChanges=0)
            v.Quit()
            return
        except Exception:
            time.sleep(0.25)


def main():
    # 先用主线程连接清理残留文档，然后释放引用；真正的 Word 对象由
    # worker 线程在 _ensure_ready 里 Dispatch 取得，避免跨线程调用。
    prep = win32com.client.Dispatch("Word.Application")
    prep.Visible = True
    while prep.Documents.Count > 0:
        prep.Documents(1).Close(SaveChanges=0)
    prep = None

    worker = AgentWorker(word_factory=_word)
    rec = _Rec(worker)
    worker.start()
    QTest.qWait(100)

    try:
        # 1) 流式写入
        ai_client.ai_stream = lambda prompt, key, system: iter(["真实流式写入的", "第二块内容"])
        worker.send("write", "写一段测试文字")
        rec.wait_for(lambda: any("已写入 Word" in t for t in rec.texts("info")))
        content = _verify_content()
        assert "真实流式写入的" in content, content
        assert "第二块内容" in content, content
        print("ok: GUI 引擎连接真实 Word 并流式写入")

        # 2) 排版：生成阶段打桩，返回块编辑代码
        ai_client.ai_request = lambda prompt, key, system, model=None: \
            "insert_at_end('引擎排版插入项')"
        worker.send("arrange", "在末尾追加一行")
        rec.wait_for(lambda: any("代码执行完毕" in t for t in rec.texts("info")))
        assert "引擎排版插入项" in _verify_content()
        print("ok: GUI 引擎 arrange（生成 -> 执行 -> 应用）")

        # 3) 中断 -> 回滚（真实文档上恢复快照）
        def slow_stream(prompt, key, system):
            yield "待回滚内容一"
            _GATE.wait(timeout=10)
            yield "待回滚内容二"
        ai_client.ai_stream = slow_stream
        worker.send("write", "马上要被打断")
        rec.wait_for(lambda: rec.chunks and rec.chunks[-1] == "待回滚内容一")
        worker.interrupt()
        _GATE.set()
        rec.wait_for(lambda: any("已中断" in t for t in rec.texts("info")))
        worker.choose("rollback")
        rec.wait_for(lambda: any("已回滚" in t for t in rec.texts("info")))
        content = _verify_content()
        assert "待回滚内容一" not in content, content
        assert "引擎排版插入项" in content, "回滚不应丢掉中断前的内容: " + content
        print("ok: 中断 -> 回滚（真实文档快照恢复）")
    finally:
        worker.send("quit")
        worker.wait(4000)
        _GATE.clear()
        _close_word()
    print("ALL OK")


if __name__ == "__main__":
    main()
