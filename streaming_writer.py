"""流式 Markdown 写入器：块级缓冲 + 打字机动画 + 块登记。

为什么是"块级缓冲"而不是逐字直接写？
inline 格式（**粗体**、*斜体*）的标记必须等一个块解析完才能确定结构，
所以按空行 / 标题行 / 围栏闭合处切分块：块完整后立即解析、应用格式、
以动画逐字写入。块内延迟极小，体感上仍然是 AI 边想边写、Word 边出字。

支持的块级 markdown：段落、1-3 级标题、有序与无序列表、围栏代码块、表格。
- 列表：项目符号 / 自动编号样式，每个项目一个段落，整块登记为单个块；
- 代码块：等宽字体逐行写入，占用的段落数 = 行数；
- 表格：写成真正的 Word 表格（含边框与单元格），块 Range 覆盖表格与其容器段落。
每个块占用的段落数 n_paras 都被精确登记，供 DocModel 的块对齐使用。

每写完一个块，把它的 Word Range 登记进 DocModel。Word 的 Range 是动态对象——
文档别处变动时位置自动跟随，这让后续的"改写第 N 段"可以精确定位。
"""
import re
import time

from markdown_it import MarkdownIt

# --- Word COM 常量 ---
WD_STYLE_NORMAL = -1        # 普通段落
WD_STYLE_HEADING_1 = -2     # 一级标题
WD_STYLE_HEADING_2 = -3     # 二级标题
WD_STYLE_HEADING_3 = -4     # 三级标题
WD_STYLE_LIST_BULLET = -49  # 无序（项目符号）列表
WD_STYLE_LIST_NUMBER = -50  # 有序（编号）列表
WD_IN_TABLE = 12            # Range.Information 参数：位置是否在表格内
CODE_FONT_NAME = "Consolas"

_HEADING_RE = re.compile(r"^#{1,3}\s")
_FENCE_LINE_RE = re.compile(r"^\s*```")


class _BlockDraft:
    """正在写入的一个块的临时记录。"""

    def __init__(self, kind, level, start, n_paras=1):
        self.kind = kind
        self.level = level
        self.start = start
        self.end = None
        self.text_parts = []
        self.n_paras = n_paras  # 该块占用的段落数（列表 = 项目数，代码 = 行数）

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
        self._md = MarkdownIt().enable("table")  # 启用 GFM 表格规则

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
        围栏代码块内部可能包含空行，因此遇到围栏开头时必须等到围栏闭合，
        绝不能按空行从中间切分。
        返回 (block, rest)；没有完整块时 block 为 None。
        """
        buf = self.pending
        pos = 0
        in_fence = False
        while pos < len(buf):
            nl = buf.find("\n", pos)
            line_end = nl if nl != -1 else len(buf)
            line = buf[pos:line_end]
            if in_fence:
                if _FENCE_LINE_RE.match(line):
                    if nl == -1:
                        return None, None  # 闭合行尚未完整
                    return buf[:nl + 1], buf[nl + 1:]
                pos = nl + 1 if nl != -1 else len(buf)
                continue
            if _FENCE_LINE_RE.match(line):
                in_fence = True
                pos = nl + 1 if nl != -1 else len(buf)
                continue
            if line.strip() == "" and nl != -1:
                # 空行：块在本空行之前
                return buf[:pos], buf[nl + 1:]
            if _HEADING_RE.match(line) and nl != -1:
                return buf[:nl], buf[nl + 1:]
            pos = nl + 1 if nl != -1 else len(buf)
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
        md_text = (md_text or "").strip()
        if not md_text:
            return []
        tokens = self._md.parse(md_text)
        registered = []
        started = False   # 本次调用是否已写出首个块（用于块间 TypeParagraph）
        in_list = False
        list_style = WD_STYLE_LIST_BULLET
        first_item = True
        current = None

        def close_current():
            nonlocal current
            if current is not None:
                current.end = self.sel.Range.Start
                current.finish(self.doc, md_text, self.model, registered)
                current = None

        i = 0
        n_tokens = len(tokens)
        while i < n_tokens:
            tok = tokens[i]
            t = tok.type
            if t == "paragraph_open" and in_list:
                # 列表项内部的段落：换行与样式都由 list_item 统一处理
                i += 1
                continue
            if t in ("paragraph_open", "heading_open", "bullet_list_open", "ordered_list_open"):
                close_current()
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
                else:  # bullet_list_open / ordered_list_open
                    in_list = True
                    first_item = True
                    list_style = (WD_STYLE_LIST_BULLET if t == "bullet_list_open"
                                  else WD_STYLE_LIST_NUMBER)
                    current = _BlockDraft("list", 0, self.sel.Range.Start, n_paras=0)
                i += 1
                continue
            if t == "list_item_open":
                if not first_item:
                    self.sel.TypeParagraph()
                first_item = False
                if current is not None:
                    current.n_paras += 1  # 每个项目占用一个段落
                self.sel.Range.Style = self.sel.Document.Styles(list_style)
                i += 1
                continue
            if t in ("bullet_list_close", "ordered_list_close"):
                in_list = False
                close_current()
                i += 1
                continue
            if t == "inline" and tok.children and current is not None:
                current.text_parts.append(self._emit_inline(tok.children, animate))
                current.end = self.sel.Range.Start
                i += 1
                continue
            if t == "fence":
                close_current()
                if started:
                    self.sel.TypeParagraph()
                started = True
                self._write_code(tok, md_text, registered)
                i += 1
                continue
            if t == "table_open":
                close_current()
                if started:
                    self.sel.TypeParagraph()
                started = True
                rows, i = self._collect_table_rows(tokens, i + 1)
                self._write_table(rows, md_text, registered)
                continue
            i += 1

        close_current()
        return registered

    # ---------- 各类块的写入 ----------

    def _write_code(self, tok, md_text, registered):
        """围栏代码块：等宽字体逐行写入，每行一个段落。"""
        self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_NORMAL)
        start = self.sel.Range.Start
        lines = tok.content.rstrip("\n").split("\n")
        for li, line in enumerate(lines):
            if li:
                self.sel.TypeParagraph()
            self._type_text(line, animate=True)
        end = self.sel.Range.Start
        try:
            self.doc.Range(start, end).Font.Name = CODE_FONT_NAME
        except Exception:
            pass  # 字体设置失败不影响内容写入
        draft = _BlockDraft("code", 0, start, n_paras=max(1, len(lines)))
        draft.end = end
        draft.text_parts.append(tok.content)
        draft.finish(self.doc, md_text, self.model, registered)

    def _collect_table_rows(self, tokens, i):
        """从 table_open 之后收集单元格文本，返回 (rows, table_close 后的索引)。"""
        rows = []
        cur_row = None
        n = len(tokens)
        while i < n and tokens[i].type != "table_close":
            tt = tokens[i].type
            if tt in ("tr_open",):
                cur_row = []
                rows.append(cur_row)
            elif tt in ("th_open", "td_open"):
                if cur_row is None:
                    cur_row = []
                    rows.append(cur_row)
                cell = ""
                if i + 1 < n and tokens[i + 1].type == "inline":
                    nxt = tokens[i + 1]
                    # 单元格里的 inline 只取文本，丢掉 ** 等标记
                    cell = "".join(ch.content for ch in (nxt.children or [])
                                   if ch.type in ("text", "code_inline"))
                cur_row.append(cell)
            i += 1
        return rows, i + 1  # 跳过 table_close

    def _write_table(self, rows, md_text, registered):
        """把收集到的表格行写成一个真正的 Word 表格并登记为块。

        Word 语义（实测）：在光标所在空段落插入表格后，原段落被劈开成表格后的
        容器段落；表格自身占用的段落数 = tbl.Range.Paragraphs.Count，
        加上容器段落即为本块占用的段落数。
        """
        rows = [r for r in rows if r] or [[]]
        ncols = max(len(r) for r in rows)
        rows = [r + [""] * (ncols - len(r)) for r in rows]
        insert_at = self.sel.Range.Start
        try:
            tbl = self.doc.Tables.Add(self.doc.Range(insert_at, insert_at), len(rows), ncols)
            try:
                tbl.Borders.Enable = True  # 网格边框
            except Exception:
                pass
            for r in range(len(rows)):
                for c in range(ncols):
                    try:
                        tbl.Cell(r + 1, c + 1).Range.Text = rows[r][c]
                    except Exception:
                        pass
        except Exception:
            # 表格创建失败时退化为逐行文本，保证内容不丢
            for r, row in enumerate(rows):
                if r:
                    self.sel.TypeParagraph()
                self._type_text(" | ".join(row), animate=True)
            draft = _BlockDraft("table", 0, insert_at, n_paras=len(rows))
            draft.end = self.sel.Range.Start
            draft.text_parts.append(" / ".join(" | ".join(r) for r in rows))
            draft.finish(self.doc, md_text, self.model, registered)
            return

        # 定位容器段落（表格后第一个非表格位置），把光标移到它末尾
        pos = tbl.Range.End
        block_end = tbl.Range.End
        guard = 0
        while guard < 200:
            try:
                if not self.doc.Range(pos, pos).Information(WD_IN_TABLE):
                    break
            except Exception:
                break
            pos += 1
            guard += 1
        try:
            container = self.doc.Range(pos, pos).Paragraphs(1)
            block_end = container.Range.End
            self.doc.Range(block_end, block_end).Select()
        except Exception:
            pass

        try:
            table_paras = tbl.Range.Paragraphs.Count
        except Exception:
            table_paras = 1
        draft = _BlockDraft("table", 0, tbl.Range.Start, n_paras=table_paras + 1)
        draft.end = block_end
        draft.text_parts.append(" / ".join(" | ".join(r) for r in rows))
        draft.finish(self.doc, md_text, self.model, registered)

    # ---------- inline 逐字写入 ----------

    def _type_text(self, text, animate):
        for c in text:
            self.sel.TypeText(c)
            if animate:
                time.sleep(self.char_delay)

    def _emit_inline(self, children, animate):
        """逐字写入一个 inline 片段，处理粗体/斜体开关；返回纯文本。"""
        parts = []
        for ch in children:
            ct = ch.type
            if ct in ("text", "code_inline"):
                self._type_text(ch.content, animate)
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
