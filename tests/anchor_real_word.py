# -*- coding: utf-8 -*-
"""真实 Word 上的写入锚点防护验证（不需要 API）。

直接驱动 StreamingWriter 往真实 Word 里流式写入，中途把选区挪到文档
别处（模拟用户点击），验证后续内容仍然接着已写内容。
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client

from doc_model import DocModel
from streaming_writer import StreamingWriter


def main():
    app = win32com.client.Dispatch("Word.Application")
    app.Visible = True
    time.sleep(1.5)  # Word 启动/重分页期间 COM 调用会被拒绝，先让它缓过劲
    doc = app.Documents.Add()
    time.sleep(0.8)
    _ = doc.Content.Text  # 预热一次读取
    sel = app.Selection
    # 这里的 feed 之间没有网络延迟，用一个真打字节奏的写入器
    writer = StreamingWriter(app, doc, sel, model=None, char_delay=0.005)
    model = DocModel(app, doc, sel, writer)
    writer.model = model

    writer.reset_anchor()
    pieces = ["第一段开头", "继续写几个字", "，第一段收尾。"]
    for p in pieces:
        writer.feed(p)
        time.sleep(0.05)
    writer.flush()
    text_mid = doc.Content.Text
    print("写入第一波后:", repr(text_mid[:80]))

    # === 模拟用户在生成期间点击文档别处 ===
    # 在文档最前面插入锚点文字并把选区挪过去
    doc.Range(0, 0).InsertBefore("[用户点这里]\r")
    sel.SetRange(0, 0)   # 光标被用户挪到文档最前面

    writer.feed("第二段紧接第一波")
    writer.feed("继续追加。")
    writer.flush()
    final = doc.Content.Text
    print("最终内容:", repr(final[:160]))

    ok = ("第二段紧接第一波继续追加。" in final.replace("\r", "")
          and final.index("第一段开头") < final.index("第二段紧接"))
    # 用户插入的文字应当在最前面，AI 的续写在后面
    assert final.startswith("[用户点这里]"), f"用户插入位置异常: {final[:30]!r}"
    assert ok, f"光标漂移防护失败，续写未接在已写内容后: {final!r}"
    print("REAL WORD ANCHOR TEST OK")
    # 不保存关闭
    doc.Close(False)
    app.Quit()


if __name__ == "__main__":
    main()
