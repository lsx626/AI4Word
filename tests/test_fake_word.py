"""离线测试：用模拟的 Word COM 对象验证流式写入与块编辑。

真实 Word 的 Range 是动态的（文档变动时自动平移），这里用字符串 + 区间平移
模拟同样的语义，从而可以在不启动 Word 的情况下验证核心逻辑。

fake 的段落模型与真实 Word 对齐：
- 任何段落都带标记，文档永远以 \r 结尾，空文档内容就是 "\r"；
- FakeDoc.Paragraphs 是整篇文档的段落集合——这样 DocModel.rebuild_ranges
  在离线也会真正执行（旧版 fake 缺这个属性，异常被吞掉，块对齐从未被测到）。

字符属性模型（草稿态/内联格式用）：
- 每个 TypeText 调用产生一个 _Seg（start, end, bold, italic, name, color），
  记录打出时的字体状态；
- 对 Range.Font 的属性赋值会切分并改写覆盖到的 segs（真实 Word 同语义），
  草稿提交时的"灰字回填"和规范化都走这条路径，离线可断言；
- ParagraphFormat.LeftIndent 的赋值被记录到 doc._indents，供嵌套列表断言；
- 表格不做模拟（Word 表格语义复杂，半吊子模拟等于制造假阳性），
  fake 下 write_block 的表格路径走的是"创建失败退化为文本"分支，表格的
  真路径由 tests/smoke_real_word.py 覆盖。
"""
import sys
import os
import time as _time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_time.sleep = lambda s: None  # 测试中关闭打字机延迟

from doc_model import DocModel
from streaming_writer import StreamingWriter


class _Seg:
    """一段连续的字符属性（对应一次 TypeText 批次）。"""

    __slots__ = ("start", "end", "bold", "italic", "name", "color")

    def __init__(self, start, end, bold=False, italic=False, name=None, color=None):
        self.start = start
        self.end = end
        self.bold = bold
        self.italic = italic
        self.name = name
        self.color = color

    def copy(self):
        return _Seg(self.start, self.end, self.bold, self.italic, self.name, self.color)


class FakeSelectionFont:
    """打字机状态：折叠选区下 Font 赋值只影响接下来要打的字。"""

    def __init__(self):
        self.bold = False
        self.italic = False
        self.name = None
        self.colorindex = None

    # COM 属性大小写不敏感，Python 属性敏感：Writer 用 Font.Name（大写），
    # 内部字段用小写 name，靠 property 桥接，TypeText 才能登记到正确字体
    @property
    def Name(self):
        return self.name

    @Name.setter
    def Name(self, v):
        self.name = v


class FakeStyle:
    def __init__(self, sid):
        self.sid = sid


class FakeParagraph:
    """一个段落：Range 含段落标记，与真实 Word 的 Paragraph 一致。"""

    def __init__(self, doc, rng):
        self._doc = doc
        self.rng = rng

    @property
    def Range(self):
        return self.rng


def _paragraph_ranges(doc):
    """把整篇文档按 \r 切成段落区间，每个区间含末尾的段落标记。"""
    content = doc.content
    ranges = []
    s = 0
    n = len(content)
    while s < n:
        nl = content.find("\r", s)
        if nl == -1:
            ranges.append((s, n))
            break
        ranges.append((s, nl + 1))
        s = nl + 1
    return ranges


class FakeParagraphs:
    """段落集合：支持 .Count 与 Paragraphs(n).Range，可按任意区间过滤。"""

    def __init__(self, doc, rng):
        self._doc = doc
        self.rng = rng

    def _overlapping(self):
        a, b = self.rng.Start, self.rng.End
        out = []
        for (s, e) in _paragraph_ranges(self._doc):
            if a == b:
                if s <= a < e:  # 折叠区间：取包含该位置的段落
                    out.append((s, e))
            elif s < b and e > a:  # 任意重叠
                out.append((s, e))
        return out

    @property
    def Count(self):
        return len(self._overlapping())

    def __call__(self, n=1):
        ranges = self._overlapping()
        if n < 1 or n > len(ranges):
            raise IndexError(f"Paragraphs({n}) 越界（共 {len(ranges)} 个段落）")
        s, e = ranges[n - 1]
        return FakeParagraph(self._doc, self._doc.Range(s, e))


class FakeParagraphFormat:
    """段落格式：LeftIndent 赋值被记录；Borders 故意不可用（走 hr 降级）。"""

    def __init__(self, doc, rng):
        self._doc = doc
        self._rng = rng

    @property
    def LeftIndent(self):
        return None

    @LeftIndent.setter
    def LeftIndent(self, v):
        self._doc._indents.append((self._rng.Start, self._rng.End, v))

    def Borders(self, n):
        raise AttributeError("fake 不支持段落边框")


def _set_seg_attr(doc, a, b, attr, value):
    """把 [a, b) 区间内的字符段切分后统一赋属性（Range.Font 的真实语义）。"""
    if a >= b:
        return
    segs = doc._segs
    i = 0
    while i < len(segs):
        s = segs[i]
        if s.start < a and s.end > a:
            left = s.copy()
            left.end = a
            right = s.copy()
            right.start = a
            segs[i:i + 1] = [left, right]
            i += 1
            continue
        if s.start < b and s.end > b:
            left = s.copy()
            left.end = b
            right = s.copy()
            right.start = b
            segs[i:i + 1] = [left, right]
            i += 1
            continue
        i += 1
    for s in segs:
        if s.start >= a and s.end <= b:
            setattr(s, attr, value)


class FakeRangeFont:
    """Range.Font：赋值改写覆盖到的字符段；取值返回唯一值或 None。"""

    def __init__(self, doc, rng):
        self._doc = doc
        self._rng = rng

    def _overlap(self):
        a, b = self._rng.Start, self._rng.End
        out = []
        for s in self._doc._segs:
            if a == b:
                if s.start <= a < s.end:
                    out.append(s)
            elif s.start < b and s.end > a:
                out.append(s)
        return out

    def _apply(self, attr, value):
        _set_seg_attr(self._doc, self._rng.Start, self._rng.End, attr, value)

    @property
    def Bold(self):
        vals = {s.bold for s in self._overlap()}
        return vals.pop() if len(vals) == 1 else None

    @Bold.setter
    def Bold(self, v):
        self._apply("bold", bool(v))

    @property
    def Italic(self):
        vals = {s.italic for s in self._overlap()}
        return vals.pop() if len(vals) == 1 else None

    @Italic.setter
    def Italic(self, v):
        self._apply("italic", bool(v))

    @property
    def Name(self):
        vals = {s.name for s in self._overlap()}
        return vals.pop() if len(vals) == 1 else None

    @Name.setter
    def Name(self, v):
        self._apply("name", v)

    @property
    def ColorIndex(self):
        vals = {s.color for s in self._overlap()}
        return vals.pop() if len(vals) == 1 else None

    @ColorIndex.setter
    def ColorIndex(self, v):
        self._apply("color", v)


class FakeRange:
    def __init__(self, doc, start, end):
        self._doc = doc
        self.start = start
        self.end = end
        self._style = None
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

    @property
    def Style(self):
        return self._style

    @Style.setter
    def Style(self, value):
        self._style = value
        self._doc._styles.append((self.start, self.end, value))

    @property
    def Font(self):
        return FakeRangeFont(self._doc, self)

    @property
    def Paragraphs(self):
        return FakeParagraphs(self._doc, self)

    @property
    def ParagraphFormat(self):
        return FakeParagraphFormat(self._doc, self)

    def Delete(self):
        self._doc._delete(self.start, self.end)

    def InsertBefore(self, text):
        self._doc._insert(self.start, text)

    def InsertAfter(self, text):
        self._doc._insert(self.end, text)

    def Select(self):
        self._doc.app.Selection.pos = self.start

    def __repr__(self):
        return f"Range[{self.start},{self.end})={self.Text!r}"


class FakeDoc:
    def __init__(self, app):
        self.app = app
        # 真实 Word 中任何段落都带标记，空文档的内容就是 "\r"
        self.content = "\r"
        self._ranges = []
        self._styles = []
        self._segs = []     # 字符属性段
        self._indents = []  # LeftIndent 赋值记录
        self._undo_count = 0

    def Range(self, a, b):
        return FakeRange(self, a, b)

    @property
    def Content(self):
        return self

    @property
    def Paragraphs(self):
        return FakeParagraphs(self, self.Range(0, len(self.content)))

    def Styles(self, sid):
        return FakeStyle(sid)

    def Undo(self, times=1):
        self._undo_count += times
        return True

    def _insert(self, pos, text):
        self.content = self.content[:pos] + text + self.content[pos:]
        for r in self._ranges:
            if pos <= r.start:
                r.start += len(text)
                r.end += len(text)
            elif pos < r.end:
                r.end += len(text)
        for s in self._segs:
            if pos <= s.start:
                s.start += len(text)
                s.end += len(text)
            elif pos < s.end:
                s.end += len(text)

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
        for s in self._segs:
            if s.start >= b:
                s.start -= length
                s.end -= length
            elif s.start > a:
                s.start = a
                if s.end > b:
                    s.end -= length
                elif s.end > a:
                    s.end = a
            else:
                if s.end > b:
                    s.end -= length
                elif s.end > a:
                    s.end = a
            if s.end < s.start:
                s.end = s.start


class FakeSelection:
    def __init__(self, doc):
        self._doc = doc
        self.pos = 0
        self.font = FakeSelectionFont()

    def TypeText(self, text):
        f = self.font
        # 先插入再登记 seg：否则 _insert 的移位循环会把刚追加的 seg 自身右移
        self._doc._insert(self.pos, text)
        self._doc._segs.append(_Seg(self.pos, self.pos + len(text),
                                    f.bold, f.italic, f.name, f.colorindex))
        self.pos += len(text)

    def TypeParagraph(self):
        self.TypeText("\r")

    @property
    def Range(self):
        return self._doc.Range(self.pos, self.pos)

    @property
    def Start(self):
        return self.pos

    @property
    def End(self):
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


def assert_aligned(model, app):
    """段落数对齐断言：sum(各块 n_paras) 必须等于文档段落总数。

    rebuild_ranges 在不匹配时会静默放弃；这个断言让"块与段落失去对应"
    在离线测试里直接失败，而不是延迟到真实 Word 里才暴露。
    """
    total = app.doc.Paragraphs.Count
    assert sum(b.n_paras for b in model.blocks) == total, \
        f"段落数不对齐: blocks={[(b.kind, b.n_paras) for b in model.blocks]} doc={total}"


def segs_of(doc):
    return [s for s in doc._segs if s.end > 0]


def test_streaming_write():
    app, writer, model = make()
    md = "# 标题一\n\n正文段落，含**粗体**与*斜体*。\n\n- 项目一\n- 项目二\n\n尾段落"
    feed_in_pieces(writer, md)
    writer.flush()
    assert app.doc.content == "标题一\r正文段落，含粗体与斜体。\r项目一\r项目二\r尾段落\r", \
        f"内容不符: {app.doc.content!r}"
    kinds = [b.kind for b in model.blocks]
    assert kinds == ["heading1", "paragraph", "list", "paragraph"], kinds
    assert model.blocks[2].n_paras == 2  # 列表两个项目 = 两个段落
    assert_aligned(model, app)
    print("ok: 流式写入（内容/块类型/粗体斜体标记剥离/段落数对齐）")


def test_block_split_mid_token():
    app, writer, model = make()
    # 缓冲区切在 ** 粗体 ** 标记中间；草稿态逐字到达也能正确剥离并加粗
    for ch in "正**粗**体":
        writer.feed(ch)
    writer.flush()
    assert app.doc.content == "正粗体\r", app.doc.content
    assert model.blocks[0].text == "正粗体"
    # 草稿态断言：提交后灰色回填、粗体段落属性正确
    segs = [s for s in app.doc._segs if s.end > s.start]
    assert all(s.color == 0 for s in segs), f"灰字未回填: {[(s.color) for s in segs]}"
    bold_text = "".join(app.doc.content[s.start:s.end] for s in segs if s.bold)
    assert bold_text == "粗", f"粗体段落错: {bold_text!r}"
    assert_aligned(model, app)
    print("ok: 标记跨片到达不残留（草稿态：颜色回填、粗体段正确）")


def test_draft_streaming_multiline():
    app, writer, model = make()
    # 段内软换行不输入；空行切分两个段落块
    feed_in_pieces(writer, "第一段\n第二行\n\n第二段")
    writer.flush()
    assert app.doc.content == "第一段第二行\r第二段\r", app.doc.content
    assert [b.kind for b in model.blocks] == ["paragraph", "paragraph"]
    assert_aligned(model, app)
    print("ok: 草稿态多行段落（软换行不输入、空行切块）")


def test_draft_structure_interrupt():
    app, writer, model = make()
    # 草稿写到一半遇到结构行：立即提交草稿，结构行另成块
    feed_in_pieces(writer, "前言一句\n- 列表甲\n- 列表乙", sizes=(3,))
    writer.flush()
    assert app.doc.content == "前言一句\r列表甲\r列表乙\r", app.doc.content
    assert [b.kind for b in model.blocks] == ["paragraph", "list"]
    assert model.blocks[1].n_paras == 2
    assert_aligned(model, app)
    print("ok: 草稿被结构行打断（立即提交、pending 截到结构行起点）")


def test_replace_block():
    app, writer, model = make()
    feed_in_pieces(writer, "# 标题\n\n旧内容段落\n\n尾段", sizes=(4,))
    writer.flush()
    new_blocks = model.replace_block(1, "全新段落")
    assert app.doc.content == "标题\r全新段落\r尾段\r", app.doc.content
    assert len(model.blocks) == 3
    assert model.blocks[1].text == "全新段落"
    assert model.blocks[2].text == "尾段"  # 后续块 Range 自动平移
    assert_aligned(model, app)
    print("ok: replace_block（重写并保持后续块定位）")


def test_replace_block_with_markdown():
    app, writer, model = make()
    writer.write_block("段落甲", animate=False)
    model.replace_block(0, "## 新标题\n\n新正文")
    assert app.doc.content == "新标题\r新正文\r", app.doc.content
    assert [b.kind for b in model.blocks] == ["heading2", "paragraph"]
    assert_aligned(model, app)
    print("ok: replace_block 多块 markdown")


def test_insert_after():
    app, writer, model = make()
    feed_in_pieces(writer, "第一段\n\n第三段")
    writer.flush()
    model.insert_after(0, "第二段")
    assert app.doc.content == "第一段\r第二段\r第三段\r", app.doc.content
    assert [b.text for b in model.blocks] == ["第一段", "第二段", "第三段"]
    assert_aligned(model, app)
    print("ok: insert_after（中间插入，块顺序与定位正确）")


def test_insert_after_last():
    app, writer, model = make()
    feed_in_pieces(writer, "第一段")
    writer.flush()
    model.insert_at_end("最后一段")
    assert app.doc.content == "第一段\r最后一段\r", app.doc.content
    assert_aligned(model, app)
    print("ok: insert_at_end")


def test_insert_after_list():
    app, writer, model = make()
    writer.write_block("甲\n\n- 项一\n- 项二", animate=False)
    model.insert_after(1, "乙")
    # 插入点必须在列表块末尾（而非第一项之后）
    assert app.doc.content == "甲\r项一\r项二\r乙\r", app.doc.content
    assert [b.text for b in model.blocks] == ["甲", "项一\r项二", "乙"]
    assert_aligned(model, app)
    print("ok: 列表块后插入（定位在块末尾而非第一项后）")


def test_delete_block():
    app, writer, model = make()
    feed_in_pieces(writer, "甲\n\n乙\n\n丙")
    writer.flush()
    model.delete_block(1)
    assert app.doc.content == "甲\r丙\r", app.doc.content
    assert [b.text for b in model.blocks] == ["甲", "丙"]
    model.delete_block(0)
    assert app.doc.content == "丙\r", app.doc.content
    assert_aligned(model, app)
    print("ok: delete_block（含段落标记删除，后续块平移）")


def test_delete_list_block_no_orphan():
    app, writer, model = make()
    writer.write_block("甲\n\n- 项一\n- 项二\n\n乙", animate=False)
    model.delete_block(1)
    # 旧实现只删块 Range（不含末尾段落标记），会留下孤儿空段落
    assert app.doc.content == "甲\r乙\r", f"疑似孤儿段落: {app.doc.content!r}"
    assert [b.text for b in model.blocks] == ["甲", "乙"]
    assert_aligned(model, app)
    print("ok: 删除多项目列表块（无孤儿空段落、段落对齐）")


def test_list_write():
    app, writer, model = make()
    writer.write_block("- 甲\n- 乙\n- 丙", animate=False)
    assert app.doc.content == "甲\r乙\r丙\r", app.doc.content
    assert len(model.blocks) == 1 and model.blocks[0].kind == "list"
    assert model.blocks[0].n_paras == 3
    sids = [s.sid for (_, _, s) in app.doc._styles]
    assert -49 in sids, f"无序列表未应用项目符号样式: {sids}"
    assert_aligned(model, app)
    print("ok: 无序列表块（多项目写为连续段落、登记为单个块、应用符号样式）")


def test_ordered_list():
    app, writer, model = make()
    writer.write_block("1. 甲\n2. 乙\n3. 丙", animate=False)
    assert app.doc.content == "甲\r乙\r丙\r", f"内容不符: {app.doc.content!r}"
    b = model.blocks[0]
    assert b.kind == "list" and b.n_paras == 3, (b.kind, b.n_paras)
    sids = [s.sid for (_, _, s) in app.doc._styles]
    assert -50 in sids, f"有序列表未应用编号样式: {sids}"
    assert_aligned(model, app)
    print("ok: 有序列表（编号样式、每项一段、无双倍空段落）")


def test_nested_list():
    app, writer, model = make()
    # 旧实现把嵌套列表劈成多个块、n_paras 全乱；现在整块登记、子项缩进
    writer.write_block("- 甲\n  - 甲一\n- 乙", animate=False)
    assert app.doc.content == "甲\r甲一\r乙\r", f"内容不符: {app.doc.content!r}"
    assert len(model.blocks) == 1, [b.kind for b in model.blocks]
    b = model.blocks[0]
    assert b.kind == "list" and b.n_paras == 3, (b.kind, b.n_paras)
    indents = [v for (_, _, v) in app.doc._indents]
    assert 0 in indents and 21 in indents, f"嵌套层级缩进缺失: {indents}"
    assert_aligned(model, app)
    print("ok: 嵌套列表（整块登记、n_paras 精确、按层级缩进）")


def test_quote():
    app, writer, model = make()
    writer.write_block("> 引用甲\n>\n> 引用乙", animate=False)
    assert app.doc.content == "引用甲\r引用乙\r", f"内容不符: {app.doc.content!r}"
    b = model.blocks[0]
    assert b.kind == "quote" and b.n_paras == 2, (b.kind, b.n_paras)
    assert 21 in [v for (_, _, v) in app.doc._indents], "引用块未缩进"
    assert_aligned(model, app)
    print("ok: 引用块（斜体、缩进、每段一段、整块登记）")


def test_hr():
    app, writer, model = make()
    # fake 无边框支持，走字符降级
    writer.write_block("前言\n\n---\n\n后语", animate=False)
    assert app.doc.content == "前言\r――――――――――――\r后语\r", f"内容不符: {app.doc.content!r}"
    kinds = [b.kind for b in model.blocks]
    assert kinds == ["paragraph", "hr", "paragraph"], kinds
    assert model.blocks[1].n_paras == 1
    assert_aligned(model, app)
    print("ok: 水平线（无边框时降级为字符横线、占一段）")


def test_link_and_code_inline():
    app, writer, model = make()
    # fake 无 Hyperlinks 支持：链接文字不能丢；行内代码走 Consolas 字体
    writer.write_block("看[链接文字](https://example.com)与`code`结尾", animate=False)
    assert app.doc.content == "看链接文字与code结尾\r", f"内容不符: {app.doc.content!r}"
    segs = [s for s in app.doc._segs if s.end > s.start]
    code_segs = [s for s in segs if s.name == "Consolas"]
    assert code_segs and "".join(app.doc.content[s.start:s.end] for s in code_segs) == "code", \
        f"行内代码字体错: {[(app.doc.content[s.start:s.end], s.name) for s in segs]}"
    assert_aligned(model, app)
    print("ok: 超链接降级（文字不丢）+ 行内代码字体")


def test_code_block_streaming():
    app, writer, model = make()
    md = "```\nprint(1)\n\n空行在代码块内\n```"
    feed_in_pieces(writer, md, sizes=(2,))
    writer.flush()
    # 围栏标记绝不能泄漏为正文；代码块内的空行不能把块从中间切断
    assert "```" not in app.doc.content, f"围栏标记泄漏: {app.doc.content!r}"
    assert app.doc.content == "print(1)\r\r空行在代码块内\r", f"内容不符: {app.doc.content!r}"
    b = model.blocks[0]
    assert b.kind == "code", b.kind
    assert b.n_paras == 3, b.n_paras  # 三行 = 三个段落
    assert_aligned(model, app)
    print("ok: 围栏代码块（跨分片完整、空行不切断、按行占段）")


def test_table_fallback_keeps_content():
    app, writer, model = make()
    # fake 没有表格支持，走"创建失败退化为逐行文本"分支，验证内容不丢
    writer.write_block("| 甲 | 乙 |\n| --- | --- |\n| 丙 | 丁 |", animate=False)
    content = app.doc.content
    for cell in "甲乙丙丁":
        assert cell in content, f"单元格内容丢失: {content!r}"
    b = model.blocks[-1]
    assert b.kind == "table", b.kind
    assert_aligned(model, app)
    print("ok: 表格降级路径（无表格支持时内容不丢、段落对齐）")


def test_greedy_alignment_and_undo():
    app, writer, model = make()
    feed_in_pieces(writer, "甲\n\n乙\n\n丙")
    writer.flush()
    # 手动制造漂移：块模型与文档段落数不一致
    model.blocks[0].n_paras = 99
    model.rebuild_ranges()
    assert model.alignment_warning, "漂移时应有对齐警告"
    # 贪心前缀对齐：前两块仍有有效 Range
    assert model.blocks[1].text == "乙"
    # undo 走 doc.Undo 封装
    assert model.undo(2)
    assert app.doc._undo_count == 2
    print("ok: 段落漂移贪心对齐 + 警告 + undo")


def test_apply_and_preview_edit():
    app, writer, model = make()
    feed_in_pieces(writer, "旧段落一\n\n旧段落二")
    writer.flush()
    # 非法 spec 必须抛 ValueError（不执行）
    for bad in ({"op": "noop"}, {"op": "replace", "index": 0},
                {"op": "delete", "index": 99}):
        try:
            model.apply_edit(bad)
        except ValueError:
            continue
        raise AssertionError(f"非法 spec 未拒绝: {bad}")
    # 预览不落盘
    spec = {"op": "replace", "index": 1, "md": "新段落二"}
    pv = model.preview_edit(spec)
    assert pv["before"] == "旧段落二" and pv["after"] == "新段落二", pv
    assert app.doc.content == "旧段落一\r旧段落二\r", "preview 不应改动文档"
    # 执行
    r = model.apply_edit(spec)
    assert r["created"] == 1 and app.doc.content == "旧段落一\r新段落二\r", app.doc.content
    # 插入与删除
    model.apply_edit({"op": "insert_after", "index": 0, "md": "插入段"})
    assert app.doc.content == "旧段落一\r插入段\r新段落二\r", app.doc.content
    model.apply_edit({"op": "delete", "index": 1})
    assert app.doc.content == "旧段落一\r新段落二\r", app.doc.content
    model.apply_edit({"op": "insert_at_end", "md": "尾段"})
    assert app.doc.content == "旧段落一\r新段落二\r尾段\r", app.doc.content
    assert_aligned(model, app)
    print("ok: 声明式编辑（apply/preview：校验、预览不落盘、四种操作）")


def test_watch_detects_user_edit():
    app, writer, model = make()
    feed_in_pieces(writer, "段落一")
    writer.flush()
    assert model.drain_events() == []  # 自身写入不误报
    # 模拟用户在停顿期间手动改动
    app.doc.content = app.doc.content + "用户加的\r"
    model.watch._last_poll = 0.0  # 绕过 1s 轮询间隔，模拟下一次空闲点
    model.watch.poll()
    events = model.drain_events()
    assert events and "段落数" in events[0], events
    print("ok: 文档看门狗（自身写入不误报、用户改动入队）")


if __name__ == "__main__":
    test_streaming_write()
    test_block_split_mid_token()
    test_draft_streaming_multiline()
    test_draft_structure_interrupt()
    test_replace_block()
    test_replace_block_with_markdown()
    test_insert_after()
    test_insert_after_last()
    test_insert_after_list()
    test_delete_block()
    test_delete_list_block_no_orphan()
    test_list_write()
    test_ordered_list()
    test_nested_list()
    test_quote()
    test_hr()
    test_link_and_code_inline()
    test_code_block_streaming()
    test_table_fallback_keeps_content()
    test_greedy_alignment_and_undo()
    test_apply_and_preview_edit()
    test_watch_detects_user_edit()
    print("\n全部测试通过。")