"""流式 Markdown 写入器：草稿态实时打字 + 块级缓冲 + 块登记。

两层实时性：
1. 草稿态（draft）：段落型文本一到达就以灰色"打字机"立即写入 Word，内联标记
   （**、*、`）由增量状态机实时切换字体属性、跳过标记字符本身；块边界（空行 /
   结构行 / flush）到达时提交：颜色回填为自动色、按完整解析规范化字体属性、
   登记为块。草稿态只改时序，不改最终内容。
2. 结构块（标题 / 列表 / 引用 / 代码 / 表格 / hr）：结构必须完整才能解析，按块
   缓冲后一次写入。

支持的块级 markdown：段落、1-3 级标题、有序与无序列表（含嵌套，按层级缩进）、
引用块、围栏代码块、表格、水平线；内联支持粗体 / 斜体 / 行内代码 / 超链接。

体感细节：
- 停顿即思考：流式片段间隔超过阈值时，下一块开头先停顿，模拟斟酌。
- 打字按批次（默认 6 字符）一次 COM 调用，减少往返又保留动画感。
- 写操作前后通知 DocWatch 置忙并重置基线，避免把 AI 自己的写入误报成用户改动。

每个块占用的段落数 n_paras 都被精确登记，供 DocModel 的块对齐使用。
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
WD_COLOR_AUTO = 0           # 字体颜色：自动
WD_COLOR_GRAY25 = 16        # 草稿态灰色
WD_BORDER_BOTTOM = -3       # 段落底边框 wdBorderBottom（用作水平线）
INDENT_STEP = 21            # 引用 / 嵌套列表每级缩进（磅）

_HEADING_RE = re.compile(r"^#{1,6}\s")
_FENCE_LINE_RE = re.compile(r"^\s*```")
# 结构行：标题、列表项、有序项、引用、围栏、水平线、表格行
_STRUCT_LINE_RE = re.compile(
    r"^\s*(?:#{1,6}\s|[-*+]\s|\d+[.)]\s|>\s|```|[-*_]{3,}\s*$|\|)"
)
# 歧义前缀：可能是结构行（# 标题、- 列表、> 引用、``` 代码围栏）的开头，
# 但分片尚不完整——先缓冲，等更多字符到达再定性
_AMBIGUOUS_RE = re.compile(r"^\s*(?:#{1,6}|[-*+]|>|`{1,2}|\d+[.)]?|)$")
_MARKER_CHARS = "*/`"       # 可能构成内联标记的字符
_TYPE_BATCH = 6             # 每次 COM 调用批量输入的字符数
_THINK_GAP = 0.5            # 超过这么多秒的片段间隔视为"思考"
_THINK_MAX = 1.2            # 思考停顿上限（秒）


def _find_boundary(buf):
    """返回段落草稿边界 '\n' 的索引；None 表示尚未到达边界。

    边界 = 空行，或下一行是结构行（标题 / 列表 / 引用 / 围栏 / 表格 / hr）。
    末尾单个 '\n' 后续内容未知时不算边界（可能是软换行）。
    """
    pos = 0
    while pos < len(buf):
        nl = buf.find("\n", pos)
        if nl == -1:
            return None
        rest = buf[nl + 1:]
        if not rest:
            return None          # 末尾换行：后续未知，暂不定性
        if rest[0] == "\n":
            return nl            # 空行
        if _STRUCT_LINE_RE.match(rest):
            return nl            # 下一行是结构行
        pos = nl + 1
    return None


class _BlockDraft:
    """正在写入的一个块的临时记录。"""

    def __init__(self, kind, level, start, n_paras=1):
        self.kind = kind
        self.level = level
        self.start = start
        self.end = None
        self.text_parts = []
        self.n_paras = n_paras  # 该块占用的段落数（列表 = 项目数，代码 = 行数）
        self.indents = []  # 列表每个项目段落的缩进值（按层级）

    def finish(self, doc, md, model, registered):
        rng = doc.Range(self.start, self.end)
        block = model.register(self.kind, self.level, md, "".join(self.text_parts), rng, self.n_paras)
        registered.append(block)


class _DraftState:
    """正在以灰色打字机实时写入的段落草稿。"""

    def __init__(self, start, base_name=None):
        self.start = start       # 草稿在文档中的起点
        self.raw = ""            # 原始文本（含 markdown 标记）
        self.emitted = 0         # raw 中已处理到的位置
        self.hold = ""           # 末尾尚未定性的标记字符
        self.buf = ""            # 待批量输出的纯字符
        self.plain = ""          # 已输出的纯文本
        self.bold = False
        self.italic = False
        self.name = None         # 当前字体名（None = 默认）
        self.base_name = base_name  # 草稿起点的基础字体（出代码段时恢复）
        self._name_stack = []    # 行内代码的字体名栈


class StreamingWriter:
    def __init__(self, app, doc, sel, model, char_delay=0.01):
        self.app = app
        self.doc = doc
        self.sel = sel
        self.model = model
        self.char_delay = char_delay
        self.pending = ""
        self._first_block = True
        self._draft = None
        self._think = 0.0
        self._last_feed = None
        self._md = MarkdownIt().enable("table")  # 启用 GFM 表格规则

    # ---------- 流式输入 ----------

    def feed(self, piece):
        """喂入一个流式片段（delta），内部自动切分并写出完整的块。"""
        if not piece:
            return
        self._note_gap()
        watch = getattr(self.model, "watch", None)
        if watch is not None:
            watch.poll()          # 先检测停顿期间的用户改动
            watch.begin_write()   # 自身写入期间屏蔽看门狗
        try:
            if self._draft is not None:
                self._draft.raw += piece
                self._draft_advance()
                return
            self.pending += piece
            self._pump()
        finally:
            if watch is not None:
                watch.end_write()  # 重置基线，吞掉自身变更

    def flush(self):
        """流结束后调用，把缓冲区里的剩余部分写出。"""
        watch = getattr(self.model, "watch", None)
        if watch is not None:
            watch.begin_write()
        try:
            if self._draft is not None:
                d = self._draft
                self._draft_process(len(d.raw), final=True)
                self._commit_draft(d.raw.strip())
                self._draft = None
            rest = self.pending.strip()
            self.pending = ""
            if rest:
                self._write_top_level(rest)
        finally:
            if watch is not None:
                watch.end_write()

    def _note_gap(self):
        """记录片段间隔：超过阈值则攒一次"思考"停顿，下一块开头消费。"""
        now = time.time()
        if self._last_feed is not None:
            gap = now - self._last_feed
            if gap > _THINK_GAP:
                self._think = min(gap * 0.5, _THINK_MAX)
        self._last_feed = now

    def _pump(self):
        """调度循环：段落走草稿态，结构行走块缓冲。"""
        while True:
            if self._draft is not None:
                self._draft_advance()
                if self._draft is not None:
                    return  # 草稿仍在等待更多输入
                continue     # 草稿已在边界处提交，pending 可能还有新内容
            if not self.pending:
                return
            if self._starts_paragraph(self.pending):
                self._start_draft()
                continue
            block, rest = self._split_pending()
            if block is None:
                return
            self.pending = rest
            self._write_top_level(block)

    def _starts_paragraph(self, buf):
        """缓冲区首行是否为段落型（非结构行、非空行、非歧义前缀）。"""
        nl = buf.find("\n")
        line = buf if nl == -1 else buf[:nl]
        if line.strip() == "":
            return False  # 空行交给 _split_pending 当空块吞掉
        if _STRUCT_LINE_RE.match(line):
            return False
        if _AMBIGUOUS_RE.match(line):
            return False  # 歧义前缀：先缓冲，等下一片到达再定性
        return True

    def _start_draft(self):
        """开始一段灰色实时草稿。"""
        self._ensure_new_paragraph()
        try:
            base_name = self.sel.Font.Name  # 基础字体：行内代码结束后恢复用
        except Exception:
            base_name = None
        d = _DraftState(self.sel.Range.Start, base_name)
        d.raw = self.pending
        self.pending = ""
        self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_NORMAL)
        try:
            self.sel.Font.ColorIndex = WD_COLOR_GRAY25
        except Exception:
            pass  # 颜色设置失败不影响内容
        self._draft = d
        self._draft_advance()

    # ---------- 草稿态 ----------

    def _draft_advance(self):
        """处理草稿中已确定的部分；到达边界则提交。"""
        d = self._draft
        raw = d.raw
        b = _find_boundary(raw)
        limit = len(raw) if b is None else b
        # 末尾待定字符先攒着：标记字符（*/`）与可能是边界前兆的换行
        hold_back = 0
        while limit - hold_back > d.emitted and raw[limit - 1 - hold_back] in _MARKER_CHARS:
            hold_back += 1
        if b is None and limit - hold_back > d.emitted:
            # 换行后若是歧义结构前缀（-、#、>、` 等），把换行与前缀整体攒住，
            # 避免把可能成为列表标记的字符按正文打出
            ln = raw.rfind("\n", d.emitted, limit - hold_back)
            if ln != -1 and _AMBIGUOUS_RE.match(raw[ln + 1:]):
                hold_back = limit - ln
        self._draft_process(limit - hold_back, final=False)
        if b is None:
            return
        rest = raw[b + 1:]
        if rest.startswith("\n"):
            rest = rest[1:]  # 空行整体吞掉
        self._commit_draft(raw[:b])
        self.pending = rest
        self._draft = None

    def _draft_process(self, limit, final):
        """把 raw[emitted:limit] 过一遍增量标记状态机并打出。"""
        d = self._draft
        raw = d.raw
        while d.emitted < limit:
            ch = raw[d.emitted]
            if ch in _MARKER_CHARS:
                d.hold += ch
                d.emitted += 1
                continue
            if d.hold:
                self._resolve_markers(d, ch)
                d.hold = ""
            if ch == "\n":
                d.emitted += 1
                continue  # 软换行不输入（与批量解析的 softbreak 处理一致）
            self._draft_type(ch)
            d.emitted += 1
        if final and d.hold:
            self._resolve_markers(d, None)
            d.hold = ""
        self._flush_draft_buffer()

    def _resolve_markers(self, d, next_ch):
        """判定攒着的标记字符是标记（切换属性）还是普通字符（原样打出）。"""
        run = d.hold
        if next_ch is not None and next_ch.isspace():
            # 后接空白：markdown 规则下不构成标记，原样输出
            for c in run:
                self._draft_type(c)
            return
        c = run[0]
        if c == "`":
            if d.name == CODE_FONT_NAME and d._name_stack:
                d.name = d._name_stack.pop()
            else:
                d._name_stack.append(d.name)
                d.name = CODE_FONT_NAME
        else:
            n = min(len(run), 3)
            if n >= 2:
                d.bold = not d.bold
            if n == 1 or n == 3:
                d.italic = not d.italic

    def _draft_type(self, ch):
        d = self._draft
        if not d.buf:
            # 批次起点：先把当前字体状态应用到选区
            self.sel.Font.Bold = d.bold
            self.sel.Font.Italic = d.italic
            # 出代码段后恢复基础字体（草稿起点捕获的字体名，真实 Word 中必为字体名）
            nm = d.name if d.name is not None else d.base_name
            try:
                self.sel.Font.Name = nm
            except Exception:
                pass
        d.buf += ch
        if len(d.buf) >= _TYPE_BATCH:
            self._flush_draft_buffer()

    def _flush_draft_buffer(self):
        d = self._draft
        if d.buf:
            self._type_text(d.buf, animate=True)
            d.plain += d.buf
            d.buf = ""

    def _commit_draft(self, md):
        """块边界到达：草稿转正——回填颜色、规范化字体属性、登记为块。"""
        d = self._draft
        self._flush_draft_buffer()
        end = self.sel.Range.Start
        rng = self.doc.Range(d.start, end)
        try:
            rng.Font.ColorIndex = WD_COLOR_AUTO  # 灰字回填为自动色
        except Exception:
            pass
        self._normalize_inline(rng, md, d.plain, d.start)
        self.model.register("paragraph", 0, md, d.plain, rng, 1)

    def _normalize_inline(self, rng, md, plain, start):
        """用完整解析校验草稿的字体属性：文本一致时按解析结果重设粗体/斜体/字体。

        增量状态机在正常配对标记下与解析一致；未闭合标记等边缘情形下解析更
        可信，此处统一兜底。文本长度不符（存在被剥掉的字面标记）时放弃，避免
        错位。
        """
        try:
            tokens = self._md.parse(md)
        except Exception:
            return
        for tok in tokens:
            if tok.type != "inline" or not tok.children:
                continue
            runs = []
            bold = italic = False
            for ch in tok.children:
                ct = ch.type
                if ct in ("text", "code_inline"):
                    nm = CODE_FONT_NAME if ct == "code_inline" else None
                    runs.append((ch.content, bold, italic, nm))
                elif ct == "strong_open":
                    bold = True
                elif ct == "strong_close":
                    bold = False
                elif ct == "em_open":
                    italic = True
                elif ct == "em_close":
                    italic = False
            text = "".join(r[0] for r in runs)
            if text != plain:
                return  # 长度不符：保留状态机的结果
            off = 0
            for t, bv, iv, nm in runs:
                if t:
                    try:
                        r = self.doc.Range(start + off, start + off + len(t))
                        r.Font.Bold = bv
                        r.Font.Italic = iv
                        if nm is not None:
                            r.Font.Name = nm
                    except Exception:
                        pass
                    off += len(t)
            break  # 草稿只含一个段落，处理第一个 inline 即可

    # ---------- 块缓冲切分（结构块） ----------

    def _split_pending(self):
        """从缓冲区里切出一个完整结构块。

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
        md_text = (md_text or "").strip()
        if not md_text:
            return
        self._ensure_new_paragraph()
        self.write_block(md_text, animate=True)

    def _ensure_new_paragraph(self):
        """新块开头：先消费"思考"停顿，非首块插入段落分隔。"""
        if self._think:
            time.sleep(self._think)
            self._think = 0.0
        if not self._first_block:
            self.sel.TypeParagraph()
        self._first_block = False

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
        list_depth = 0    # 列表嵌套深度（0 = 不在列表中）
        quote_depth = 0   # 引用嵌套深度
        list_style = WD_STYLE_LIST_BULLET
        current = None

        def close_current():
            nonlocal current
            if current is not None:
                current.end = self.sel.Range.Start
                if current.kind == "list" and current.indents:
                    # 实测：TypeText / TypeParagraph 会把列表段落的直接缩进
                    # 归一化回 List 样式内置值（如 22 磅），逐段赋值会被后续
                    # 打字冲掉；只能在整块打完后统一重设一次（此后稳定）。
                    try:
                        blk = self.doc.Range(current.start, current.end)
                        paras = blk.Paragraphs
                        for k, ind in enumerate(current.indents[:paras.Count]):
                            paras(k + 1).Range.ParagraphFormat.LeftIndent = ind
                    except Exception:
                        pass
                current.finish(self.doc, md_text, self.model, registered)
                current = None

        i = 0
        n_tokens = len(tokens)
        while i < n_tokens:
            tok = tokens[i]
            t = tok.type
            if t == "paragraph_open":
                if list_depth > 0:
                    pass  # 列表项的段落由 list_item_open 统一处理
                elif quote_depth > 0:
                    # 引用内部的段落：每段一行，缩进随引用深度
                    if current is not None and current.kind == "quote":
                        if current.n_paras > 0:
                            self.sel.TypeParagraph()
                        current.n_paras += 1
                        self._quote_para_format(quote_depth)
                else:
                    close_current()
                    if started:
                        self.sel.TypeParagraph()
                    started = True
                    self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_NORMAL)
                    current = _BlockDraft("paragraph", 0, self.sel.Range.Start)
                i += 1
                continue
            if t == "heading_open":
                close_current()
                if started:
                    self.sel.TypeParagraph()
                started = True
                level = int(tok.tag[1])
                style = {1: WD_STYLE_HEADING_1, 2: WD_STYLE_HEADING_2,
                         3: WD_STYLE_HEADING_3}.get(level, WD_STYLE_NORMAL)
                self.sel.Range.Style = self.sel.Document.Styles(style)
                current = _BlockDraft(f"heading{level}", level, self.sel.Range.Start)
                i += 1
                continue
            if t in ("bullet_list_open", "ordered_list_open"):
                list_depth += 1
                if list_depth == 1:
                    # 顶层列表：结束上一块，开始列表块；嵌套列表只是深度变化
                    close_current()
                    if started:
                        self.sel.TypeParagraph()
                    started = True
                    list_style = (WD_STYLE_LIST_BULLET if t == "bullet_list_open"
                                  else WD_STYLE_LIST_NUMBER)
                    current = _BlockDraft("list", 0, self.sel.Range.Start, n_paras=0)
                i += 1
                continue
            if t in ("bullet_list_close", "ordered_list_close"):
                list_depth -= 1
                if list_depth == 0:
                    close_current()
                i += 1
                continue
            if t == "list_item_open":
                # 任意深度的列表项都占一个段落，缩进随深度（ListLevelNumber 在
                # pywin32 下对选中空段落设置会失败，改用 LeftIndent 表达层级；
                # 缩进赋值见 paragraph_close 与 close_current 的注释）
                if current is not None and current.kind == "list":
                    if current.n_paras > 0:
                        self.sel.TypeParagraph()
                    current.n_paras += 1
                    current.indents.append(INDENT_STEP * max(0, list_depth - 1))
                    self.sel.Range.Style = self.sel.Document.Styles(list_style)
                i += 1
                continue
            if t == "paragraph_close":
                # 列表项文字打完的瞬间先设一次缩进（实时可视）；
                # 整块结束前 close_current 会再统一重设（Word 会归一化中途的赋值）
                if list_depth > 0 and current is not None and current.kind == "list":
                    try:
                        self.sel.Range.ParagraphFormat.LeftIndent = \
                            INDENT_STEP * max(0, list_depth - 1)
                    except Exception:
                        pass
                i += 1
                continue
            if t == "blockquote_open":
                quote_depth += 1
                if quote_depth == 1:
                    close_current()
                    if started:
                        self.sel.TypeParagraph()
                    started = True
                    self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_NORMAL)
                    current = _BlockDraft("quote", 0, self.sel.Range.Start, n_paras=0)
                i += 1
                continue
            if t == "blockquote_close":
                quote_depth -= 1
                if quote_depth == 0:
                    close_current()
                    try:
                        self.sel.Font.Italic = False
                    except Exception:
                        pass
                i += 1
                continue
            if t == "hr":
                close_current()
                if started:
                    self.sel.TypeParagraph()
                started = True
                self._write_hr(md_text, registered)
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
            if t == "inline" and tok.children and current is not None:
                current.text_parts.append(self._emit_inline(tok.children, animate))
                current.end = self.sel.Range.Start
                i += 1
                continue
            i += 1

        close_current()
        if quote_depth > 0:
            # 引用未闭合（markdown 不完整）：兜底恢复斜体
            try:
                self.sel.Font.Italic = False
            except Exception:
                pass
        return registered

    def _quote_para_format(self, depth):
        """引用内部段落的格式：普通样式、按深度缩进、斜体。"""
        self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_NORMAL)
        try:
            self.sel.Range.ParagraphFormat.LeftIndent = INDENT_STEP * depth
        except Exception:
            pass
        try:
            self.sel.Font.Italic = True
        except Exception:
            pass

    # ---------- 各类块的写入 ----------

    def _write_hr(self, md_text, registered):
        """水平线：优先用段落底边框；无边框支持时降级为字符横线。"""
        self.sel.Range.Style = self.sel.Document.Styles(WD_STYLE_NORMAL)
        start = self.sel.Range.Start
        try:
            self.sel.Range.ParagraphFormat.Borders(WD_BORDER_BOTTOM).LineStyle = 1
            draft = _BlockDraft("hr", 0, start, n_paras=1)
            draft.end = self.sel.Range.Start
            draft.text_parts.append("")
            draft.finish(self.doc, md_text, self.model, registered)
            return
        except Exception:
            pass  # 无边框支持，走字符降级
        line = "―" * 12
        self._type_text(line, animate=True)
        draft = _BlockDraft("hr", 0, start, n_paras=1)
        draft.end = self.sel.Range.Start
        draft.text_parts.append(line)
        draft.finish(self.doc, md_text, self.model, registered)

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
        """批量打字：每次 COM 调用输入约 _TYPE_BATCH 个字符，停顿按比例。"""
        if not text:
            return
        for i in range(0, len(text), _TYPE_BATCH):
            batch = text[i:i + _TYPE_BATCH]
            self.sel.TypeText(batch)
            if animate:
                time.sleep(self.char_delay * len(batch))

    def _emit_inline(self, children, animate):
        """逐字写入一个 inline 片段，处理粗体/斜体/行内代码/链接；返回纯文本。"""
        parts = []
        link_stack = []  # (href, 链接文字起点)
        for ch in children:
            ct = ch.type
            if ct == "text":
                self._type_text(ch.content, animate)
                parts.append(ch.content)
            elif ct == "code_inline":
                prev = None
                try:
                    prev = self.sel.Font.Name
                    self.sel.Font.Name = CODE_FONT_NAME
                except Exception:
                    prev = None
                self._type_text(ch.content, animate)
                parts.append(ch.content)
                try:
                    self.sel.Font.Name = prev  # 打完代码恢复原字体（真实 Word 中 prev 必为字体名）
                except Exception:
                    pass
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
            elif ct == "link_open":
                href = None
                try:
                    href = ch.attrGet("href")
                except Exception:
                    href = None
                link_stack.append((href, self.sel.Range.Start))
            elif ct == "link_close":
                if link_stack:
                    href, ls = link_stack.pop()
                    if href:
                        try:
                            self.doc.Hyperlinks.Add(
                                self.doc.Range(ls, self.sel.Range.Start), Address=href)
                        except Exception:
                            pass  # 无超链接支持时退化为纯文本（文字不丢）
        return "".join(parts)