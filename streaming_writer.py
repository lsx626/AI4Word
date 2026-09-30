"""流式 Markdown 写入器：块级缓冲 + 打字机动画 + 块登记。

为什么是"块级缓冲"而不是逐字直接写？
inline 格式（**粗体**、*斜体*）的标记必须等一个块解析完才能确定结构，
所以按空行 / 标题行切分块：块完整后立即解析、应用格式并以动画逐字写入。
块内延迟极小，体感上仍然是 AI 边想边写、Word 边出字。

每写完一个块，把它的 Word Range 登记进 DocModel。Word 的 Range 是动态对象——
文档别处变动时位置自动跟随，这让后续的"改写第 N 段"可以精确定位。
"""
import re
import time

from markdown_it import MarkdownIt

# --- Word COM 常量 ---
WD_STYLE_NORMAL = -1      # 普通段落
WD_STYLE_HEADING_1 = -2   # 一级标题
WD_STYLE_HEADING_2 = -3   # 二级标题
WD_STYLE_HEADING_3 = -4   # 三级标题
WD_STYLE_LIST_BULLET = -49  # 无序列表

_HEADING_RE = re.compile(r"^#{1,3}\s")


class _BlockDraft:
    """正在写入的一个块的临时记录。"""

    def __init__(self, kind, level, start):
        self.kind = kind
        self.level = level
        self.start = start
        self.end = None
        self.text_parts = []
        self.n_paras = 1  # 该块占用的段落数（列表 = 项目数）

    def finish(self, doc, md, model, registered):
        rng = doc.Range(self.start, self.end)
        block = model.register(self.kind, self.level, md, "".join(self.text_parts), rng, self.n_paras)
        registered.append(block)


class StreamingWriter:
    def __init__(self, app, doc, sel, model, char_delay=0.01):
        self.app = app
        self.doc = doc
        self.sel = sel
        self.model = model
        self.char_delay = char_delay
        self.pending = ""
        self._first_block = True
        self._md = MarkdownIt()

    # ---------- 流式输入 ----------

    def feed(self, piece):
        """喂入一个流式片段（delta），内部自动切分并写出完整的块。"""
        if not piece:
            return
        self.pending += piece
        while True:
            block, rest = self._split_pending()
            if block is None:
                break
            self.pending = rest
            self._write_top_level(block)

    def flush(self):
        """流结束后调用，把缓冲区里的剩余部分写出。"""
        rest = self.pending.strip()
        self.pending = ""
        if rest:
            self._write_top_level(rest)

    def _split_pending(self):
        """从缓冲区里切出一个完整块。

        块边界 = 空行（段落之间），或标题行的行尾（标题只需一行即完整）。
        返回 (block, rest)；没有完整块时 block 为 None。
        """
        buf = self.pending
        idx = buf.find("\n\n")
        if idx != -1:
            return buf[:idx], buf[idx + 2:]
        if _HEADING_RE.match(buf) and "\n" in buf:
            nl = buf.find("\n")
            return buf[:nl], buf[nl + 1:]
        return None, None

    def _write_top_level(self, md_text):
        md_text = md_text.strip()
        if not md_text:
            return
        if not self._first_block:
            self.sel.TypeParagraph()
        self._first_block = False
        self.write_block(md_text, animate=True)

    # ---------- 块写入 ----------

    def write_block(self, md_text, animate=True):
        """在当前光标位置写入任意 Markdown（可含多个块），返回登记的块列表。

        调用方负责保证光标已经位于一个可用（通常是空）的段落起点；
        本方法只管解析、应用格式、逐字写入并登记。
        """
        tokens = self._md.parse(md_text)
        registered = []
        started = False   # 本次调用是否已写出首个块（用于块间 TypeParagraph）
        in_list = False
        first_item = True
        current = None
        for tok in tokens:
            t = tok.type
            if t == "paragraph_open" and in_list:
                # 松散列表内部的段落：换行与样式都由 list_item 统一处理
                continue
            if t in ("paragraph_open", "heading_open", "bullet_list_open"):
                if current is not None:
                    # 前一个块（段落/标题）没有被 close 事件收尾，这里补登
                    current.end = self.sel.Range.Start
                    current.finish(self.doc, md_text, self.model, registered)
                    current = None
                if started:
                    self.sel.TypeParagraph()
                started = True
                if t == "heading_open":
                    level = int(tok.tag[1])
                    style = {1: WD_STYLE_HEADING_1, 2: WD_STYLE_HEADING_2,
                             3: WD_STYLE_HEADING_3}.get(level, WD_STYLE_NORMAL)
                    self.sel.Range.Style = self.sel.Document.Styles(style)
                    current = _BlockDraft(f"heading{level}", level, self.sel.Range.Start)
                elif t == "paragraph_open":
                    self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_NORMAL)
                    current = _BlockDraft("paragraph", 0, self.sel.Range.Start)
                else:  # bullet_list_open
                    in_list = True
                    first_item = True
                    current = _BlockDraft("list", 0, self.sel.Range.Start)
            elif t == "list_item_open":
                if not first_item:
                    self.sel.TypeParagraph()
                first_item = False
                self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_LIST_BULLET)
            elif t == "inline" and tok.children and current is not None:
                current.text_parts.append(self._emit_inline(tok.children, animate))
                current.end = self.sel.Range.Start
            elif t == "bullet_list_close":
                in_list = False
                if current is not None:
                    current.end = self.sel.Range.Start
                    current.finish(self.doc, md_text, self.model, registered)
                    current = None
        if current is not None:
            current.end = self.sel.Range.Start
            current.finish(self.doc, md_text, self.model, registered)
        return registered

    def _emit_inline(self, children, animate):
        """逐字写入一个 inline 片段，处理粗体/斜体开关；返回纯文本。"""
        parts = []
        for ch in children:
            ct = ch.type
            if ct in ("text", "code_inline"):
                for c in ch.content:
                    self.sel.TypeText(c)
                    if animate:
                        time.sleep(self.char_delay)
                parts.append(ch.content)
            elif ct in ("softbreak", "hardbreak"):
                pass  # 段内换行暂按跨行连续处理
            elif ct == "strong_open":
                self.sel.Font.Bold = True
            elif ct == "strong_close":
                self.sel.Font.Bold = False
            elif ct == "em_open":
                self.sel.Font.Italic = True
            elif ct == "em_close":
                self.sel.Font.Italic = False
            # link_open / link_close / html_inline：忽略标记，链接文字照常输入
        return "".join(parts)
