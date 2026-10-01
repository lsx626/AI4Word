# -*- coding: utf-8 -*-
"""真实 Word：用户切换活动文档后，引擎自动重绑到新文档。"""
import os
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ["QT_QPA_PLATFORM"] = "offscreen"

import win32com.client

import ai_client
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from app.engine import AgentWorker

_app = QApplication([])


class _Rec:
    def __init__(self, w):
        self.msgs = []
        w.message.connect(lambda k, t: self.msgs.append((k, t)))

    def texts(self, kind):
        return [t for k, t in self.msgs if k == kind]

    def wait_for(self, pred, timeout_ms=30000):
        import time
        t0 = time.time()
        while not pred():
            QTest.qWait(20)
            if (time.time() - t0) * 1000 > timeout_ms:
                raise AssertionError("超时: " + repr(self.msgs[-8:]))


def _word():
    # 由 WORKER 线程附加到主线程启动的 Word（与 GUI 真实运行路径一致）
    return win32com.client.GetObject(None, "Word.Application")


def _main():
    import time
    # 主线程先启动 Word 并准备好两个文档（COM 对象留在主线程公寓里）
    app = win32com.client.Dispatch("Word.Application")
    app.Visible = True
    time.sleep(0.8)
    doc1 = app.Documents.Add()
    doc2 = app.Documents.Add()
    doc1.Activate()
    time.sleep(0.5)

    ai_client.ai_stream = lambda prompt, key, system: iter(["文档一的内容"])
    worker = AgentWorker(word_factory=_word)
    worker.set_api_key("test-key")
    rec = _Rec(worker)
    worker.start()
    QTest.qWait(150)

    # 1) 在 doc1（当前活动）上写
    worker.send("write", "测试")
    rec.wait_for(lambda: any("已写入 Word" in t for t in rec.texts("info")), 60000)
    print("doc1:", repr(doc1.Content.Text[:40]))
    assert "文档一的内容" in doc1.Content.Text

    # 2) 用户切换到 doc2
    doc2.Activate()
    time.sleep(0.5)

    # 3) 再发写命令：引擎应检测到活动文档变化并重绑到 doc2
    ai_client.ai_stream = lambda prompt, key, system: iter(["文档二的内容"])
    rec.msgs.clear()
    worker.send("write", "测试二")
    rec.wait_for(lambda: any("重新绑定" in t for t in rec.texts("info")), 60000)
    rec.wait_for(lambda: any("已写入 Word" in t for t in rec.texts("info")), 60000)
    print("doc1 after:", repr(doc1.Content.Text[:40]))
    print("doc2 after:", repr(doc2.Content.Text[:40]))
    assert "文档一的内容" in doc1.Content.Text, "doc1 不应被改动"
    assert "文档一的内容" not in doc2.Content.Text, "doc2 不应含 doc1 的内容"
    assert "文档二的内容" in doc2.Content.Text, "新内容应写进 doc2"

    worker.send("quit")
    worker.wait(4000)
    doc1.Close(False)
    doc2.Close(False)
    app.Quit()
    print("REAL WORD DOC-SWITCH TEST OK")


if __name__ == "__main__":
    _main()
