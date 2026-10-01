# -*- coding: utf-8 -*-
"""真实 Word：import_document 读取已有非空文档（含中文标题样式）。"""
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
    time.sleep(1.2)
    doc = app.Documents.Add()
    time.sleep(0.6)
    sel = app.Selection

    # 手工造一个「用户早已写好」的文档
    sel.TypeText("已有第一段正文内容")
    sel.TypeParagraph()
    sel.TypeText("已有第二段正文内容")
    sel.TypeParagraph()
    doc.Paragraphs(3).Range.Style = doc.Styles(-2)  # 标题 1
    sel.TypeText("已有标题文字")
    time.sleep(0.3)

    writer = StreamingWriter(app, doc, sel, model=None, char_delay=0.005)
    model = DocModel(app, doc, sel, writer)
    writer.model = writer and model
    n = model.import_document()
    print(f"imported: {n}")
    print(model.block_map())
    assert n == 3, f"应登记 3 段，实际 {n}"
    kinds = [b.kind for b in model.blocks]
    assert kinds == ["paragraph", "paragraph", "heading1"], kinds
    texts = [b.text.strip() for b in model.blocks]
    assert texts[0] == "已有第一段正文内容" and texts[2] == "已有标题文字", texts

    # 导入后块编辑可用
    model.replace_block(0, "替换后的新内容")
    final = doc.Content.Text
    print("final:", repr(final[:120]))
    assert "替换后的新内容" in final
    assert "已有第二段正文内容" in final
    print("REAL WORD IMPORT TEST OK")
    doc.Close(False)
    app.Quit()


if __name__ == "__main__":
    main()
