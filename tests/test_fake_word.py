"""离线测试：用模拟的 Word COM 对象验证流式写入与块编辑。

真实 Word 的 Range 是动态的（文档变动时自动平移），这里用字符串 + 区间平移
模拟同样的语义，从而可以在不启动 Word 的情况下验证核心逻辑。
"""
import sys
import os
import time as _time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_time.sleep = lambda s: None  # 测试中关闭打字机延迟

from doc_model import DocModel
from streaming_writer import StreamingWriter


class FakeFont:
    def __init__(self):
        self.bold = False
        self.italic = False


class FakeStyle:
    def __init__(self, sid):
        self.sid = sid


class FakeParagraph:
    def __init__(self, doc, rng):
        self._doc = doc
        self.rng = rng


class FakeParagraphs:
    def __init__(self, doc, rng):
        self._doc = doc
        self.rng = rng

    def __call__(self, n=1):
        return self  # 只支持 Paragraphs(1)

    @property
    def Range(self):
        doc = self._doc
        content = doc.content
        p = self.rng.start
        line_start = content.rfind("\r", 0, max(p, 0)) + 1 if p > 0 else 0
        nxt = content.find("\r", line_start)
        end = nxt + 1 if nxt != -1 else len(content)
        return doc.Range(line_start, end)


class FakeRange:
    def __init__(self, doc, start, end):
        self._doc = doc
        self.start = start
        self.end = end
        self.style = None
        doc._ranges.append(self)

    @property
    def Start(self):
        return self.start

    @property
    def End(self):
        return self.end

    @property
    def Text(self):
        return self._doc.content[self.start:self.end]

    @property
    def Document(self):
        return self._doc

    def Delete(self):
        self._doc._delete(self.start, self.end)

    def InsertBefore(self, text):
        self._doc._insert(self.start, text)

    def Select(self):
        self._doc.app.Selection.pos = self.start

    @property
    def Paragraphs(self):
        return FakeParagraphs(self._doc, self)

    def __repr__(self):
        return f"Range[{self.start},{self.end})={self.Text!r}"


class FakeDoc:
    def __init__(self, app):
        self.app = app
        # 真实 Word 中任何段落都带标记，空文档的内容就是 "\r"
        self.content = "\r"
        self._ranges = []

    def Range(self, a, b):
        return FakeRange(self, a, b)

    @property
    def Content(self):
        return self

    def _insert(self, pos, text):
        self.content = self.content[:pos] + text + self.content[pos:]
        for r in self._ranges:
            if pos <= r.start:
                r.start += len(text)
                r.end += len(text)
            elif pos < r.end:
                r.end += len(text)

    def _delete(self, a, b):
        length = b - a
        self.content = self.content[:a] + self.content[b:]
        for r in self._ranges:
            if r.start >= b:
                r.start -= length
                r.end -= length
            elif r.start > a:
                r.start = a
                if r.end > b:
                    r.end -= length
                elif r.end > a:
                    r.end = a
            else:
                if r.end > b:
                    r.end -= length
                elif r.end > a:
                    r.end = a
            if r.end < r.start:
                r.end = r.start

    def Styles(self, sid):
        return FakeStyle(sid)


class FakeSelection:
    def __init__(self, doc):
        self._doc = doc
        self.pos = 0
        self.font = FakeFont()

    def TypeText(self, c):
        self._doc._insert(self.pos, c)
        self.pos += len(c)

    def TypeParagraph(self):
        self.TypeText("\r")

    @property
    def Range(self):
        return self._doc.Range(self.pos, self.pos)

    @property
    def Start(self):
        return self.pos

    @property
    def Font(self):
        return self.font

    @property
    def Document(self):
        return self._doc


class FakeApp:
    def __init__(self):
        self.doc = FakeDoc(self)
        self.sel = FakeSelection(self.doc)

    @property
    def Selection(self):
        return self.sel


def make():
    app = FakeApp()
    writer = StreamingWriter(app, app.doc, app.sel, model=None)
    model = DocModel(app, app.doc, app.sel, writer)
    writer.model = model
    return app, writer, model


def feed_in_pieces(writer, text, sizes=(1, 2, 3)):
    i = 0
    while i < len(text):
        step = sizes[i % len(sizes)]
        writer.feed(text[i:i + step])
        i += step


def test_streaming_write():
    app, writer, model = make()
    md = "# 标题一\n\n正文段落，含**粗体**与*斜体*。\n\n- 项目一\n- 项目二\n\n尾段落"
    feed_in_pieces(writer, md)
    writer.flush()
    assert app.doc.content == "标题一\r正文段落，含粗体与斜体。\r项目一\r项目二\r尾段落\r", \
        f"内容不符: {app.doc.content!r}"
    kinds = [b.kind for b in model.blocks]
    assert kinds == ["heading1", "paragraph", "list", "paragraph"], kinds
    print("ok: 流式写入（内容/块类型/粗体斜体标记剥离）")


def test_block_split_mid_token():
    app, writer, model = make()
    # 缓冲区切在 ** 粗体 ** 标记中间
    for ch in "正**粗**体":
        writer.feed(ch)
    writer.flush()
    assert app.doc.content == "正粗体\r", app.doc.content
    assert model.blocks[0].text == "正粗体"
    print("ok: 标记跨片到达不残留")


def test_replace_block():
    app, writer, model = make()
    feed_in_pieces(writer, "# 标题\n\n旧内容段落\n\n尾段", sizes=(4,))
    writer.flush()
    new_blocks = model.replace_block(1, "全新段落")
    assert app.doc.content == "标题\r全新段落\r尾段\r", app.doc.content
    assert len(model.blocks) == 3
    assert model.blocks[1].text == "全新段落"
    assert model.blocks[2].text == "尾段"  # 后续块 Range 自动平移
    print("ok: replace_block（重写并保持后续块定位）")


def test_replace_block_with_markdown():
    app, writer, model = make()
    writer.write_block("段落甲", animate=False)
    model.replace_block(0, "## 新标题\n\n新正文")
    assert app.doc.content == "新标题\r新正文\r", app.doc.content
    assert [b.kind for b in model.blocks] == ["heading2", "paragraph"]
    print("ok: replace_block 多块 markdown")


def test_insert_after():
    app, writer, model = make()
    feed_in_pieces(writer, "第一段\n\n第三段")
    writer.flush()
    model.insert_after(0, "第二段")
    assert app.doc.content == "第一段\r第二段\r第三段\r", app.doc.content
    assert [b.text for b in model.blocks] == ["第一段", "第二段", "第三段"]
    print("ok: insert_after（中间插入，块顺序与定位正确）")


def test_insert_after_last():
    app, writer, model = make()
    feed_in_pieces(writer, "第一段")
    writer.flush()
    model.insert_at_end("最后一段")
    assert app.doc.content == "第一段\r最后一段\r", app.doc.content
    print("ok: insert_at_end")


def test_delete_block():
    app, writer, model = make()
    writer.write_block("甲\n\n乙\n\n丙".replace("\n\n", "\r"), animate=False) if False else None
    app2, writer2, model2 = make()
    feed_in_pieces(writer2, "甲\n\n乙\n\n丙")
    writer2.flush()
    model2.delete_block(1)
    assert app2.doc.content == "甲\r丙\r", app2.doc.content
    assert [b.text for b in model2.blocks] == ["甲", "丙"]
    model2.delete_block(0)
    assert app2.doc.content == "丙\r", app2.doc.content
    print("ok: delete_block（含段落标记删除，后续块平移）")


def test_list_write():
    app, writer, model = make()
    writer.write_block("- 甲\n- 乙\n- 丙", animate=False)
    assert app.doc.content == "甲\r乙\r丙\r", app.doc.content
    assert len(model.blocks) == 1 and model.blocks[0].kind == "list"
    print("ok: 列表块（多项目写为连续段落并登记为单个块）")


if __name__ == "__main__":
    test_streaming_write()
    test_block_split_mid_token()
    test_replace_block()
    test_replace_block_with_markdown()
    test_insert_after()
    test_insert_after_last()
    test_delete_block()
    test_list_write()
    print("\n全部测试通过。")
