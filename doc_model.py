"""块级文档模型：把每个写入的块登记为 Word Range，支持精确编辑。

核心思路：Word 的 Range 对象是动态的——文档别处的插入/删除会自动平移已经
存在的 Range。因此每个块记住自己的 Range 后，"改写第 i 块 / 在第 i 块后插入 /
删除第 i 块" 都能精确定位，不需要维护任何偏移量。

所有编辑操作完成后会把选区移动到受影响区域并闪烁高亮，让排版过程可见。
"""
import time

WD_NO_HIGHLIGHT = 0
WD_YELLOW = 7


class Block:
    def __init__(self, kind, level, md, text, rng, n_paras=1):
        self.kind = kind      # paragraph / heading1..3 / list
        self.level = level
        self.md = md          # 原始 markdown
        self.text = text      # 写入时的纯文本（编辑后由 refresh_texts 刷新）
        self.range = rng      # 动态 Range，随文档变动自动跟随
        self.n_paras = n_paras  # 该块占用的段落数（列表 = 项目数）

    def __repr__(self):
        t = (self.text or "").replace("\r", " ").replace("\n", " ")[:18]
        return f"<Block {self.kind} [{t}]>"


class DocModel:
    def __init__(self, app, doc, sel, writer=None):
        self.app = app
        self.doc = doc
        self.sel = sel
        self.writer = writer
        self.blocks = []

    def register(self, kind, level, md, text, rng, n_paras=1):
        b = Block(kind, level, md, text, rng, n_paras)
        self.blocks.append(b)
        return b

    def rebuild_ranges(self):
        """编辑后按段落扫描重建每块的 Range。

        Word 的 Range 自动平移在「恰好在块边界处插入」时会把新文字划进相邻块，
        不可靠；而块与段落的对应关系是我们自己写出来的（每块 n_paras 个段落），
        所以编辑后按段落重新对齐最稳妥。
        """
        try:
            paras = self.doc.Paragraphs
            total = paras.Count
        except Exception:
            return
        if sum(b.n_paras for b in self.blocks) != total:
            return  # 段落数不匹配（用户手动改过文档），放弃重建
        idx = 0
        for b in self.blocks:
            start = paras(idx + 1).Range.Start
            end = paras(idx + b.n_paras).Range.End - 1  # 去掉末尾段落标记
            if end < start:
                end = start
            b.range = self.doc.Range(start, end)
            idx += b.n_paras
        self.refresh_texts()

    # ---------- 查询 ----------

    def refresh_texts(self):
        for b in self.blocks:
            try:
                b.text = b.range.Text
            except Exception:
                pass

    def block_map(self):
        """返回文档结构概览（给 AI 看的块索引地图）。"""
        self.refresh_texts()
        lines = []
        for i, b in enumerate(self.blocks):
            preview = (b.text or "").replace("\r", " ").replace("\n", " ").strip()[:30]
            lines.append(f"[{i}] {b.kind}: {preview}")
        return "\n".join(lines)

    def get_block_text(self, i):
        b = self.blocks[i]
        try:
            return b.range.Text
        except Exception:
            return b.text

    # ---------- 编辑原语（会注入到 AI 生成代码的全局变量里） ----------

    def replace_block(self, i, md):
        """用新的 markdown 改写第 i 块（原段落保留，清空内容后重写）。"""
        b = self.blocks[i]
        if b.kind == "table":
            # 表格块：删掉表格本身，保留其后的容器段落作为新内容的写入点
            try:
                b.range.Tables(1).Delete()
            except Exception:
                b.range.Delete()
        else:
            b.range.Delete()
        pos = b.range.Start  # 动态 Range 在删除后塌缩到删除点
        self.doc.Range(pos, pos).Select()
        # write_block 会把新块 register 到 blocks 末尾，先记下越界点再搬运到目标位置
        before_count = len(self.blocks)
        self.writer.write_block(md, animate=True)
        new_blocks = self.blocks[before_count:]
        del self.blocks[before_count:]
        self.blocks[i:i + 1] = new_blocks
        self.rebuild_ranges()
        self._flash(new_blocks[-1].range if new_blocks else None)
        return new_blocks

    def insert_after(self, i, md):
        """在第 i 块之后插入新内容。"""
        b = self.blocks[i]
        pos = None
        if b.kind == "table":
            # 表格块的 Range 只覆盖到表格本身（不含其后的容器段落），
            # 块内最后一个段落是表格的行尾段落——不能在行尾上插入。
            # 容器段落紧随表格之后，新内容造在容器之后。
            pos = b.range.End
            try:
                container = self.doc.Range(pos, pos).Paragraphs(1)
                pos = container.Range.End
                container.Range.InsertAfter("\r")
            except Exception:
                pass
        else:
            # 插入点在块占用的最后一个段落末尾（含段落标记），多段落块不会插进块中间
            paras = b.range.Paragraphs
            last_para = paras(paras.Count)
            pos = last_para.Range.End
            last_para.Range.InsertAfter("\r")
        # 上面统一用段落的 InsertAfter 而不是 doc.Range(pos, pos).InsertBefore：
        # 当块在文档末尾时 pos 等于 Content.End，Word 无法构造折叠在文档末端的
        # Range（报"数值超出范围"）；InsertAfter 先把文档撑长，pos 随即合法。
        if pos is not None:
            self.doc.Range(pos, pos).Select()
        before_count = len(self.blocks)
        self.writer.write_block(md, animate=True)
        new_blocks = self.blocks[before_count:]
        del self.blocks[before_count:]
        self.blocks[i + 1:i + 1] = new_blocks
        self.rebuild_ranges()
        self._flash(new_blocks[-1].range if new_blocks else None)
        return new_blocks

    def insert_at_end(self, md):
        """在文档末尾追加新内容。"""
        if not self.blocks:
            return self.writer.write_block(md, animate=True)
        return self.insert_after(len(self.blocks) - 1, md)

    def delete_block(self, i):
        """删除第 i 块（连同它占用的全部段落与段落标记）。"""
        b = self.blocks[i]
        if b.kind == "table":
            self._delete_table_block(b)
        else:
            try:
                # 删除块占用的全部段落（含末尾段落标记），避免留下孤儿空段落
                paras = b.range.Paragraphs
                end = paras(paras.Count).Range.End
                self.doc.Range(b.range.Start, end).Delete()
            except Exception:
                b.range.Delete()
        self.blocks.pop(i)
        self.rebuild_ranges()
        return True

    def _delete_table_block(self, b):
        """删除表格块。

        实测：对覆盖表格的 Range 调 Delete 只删掉文字，会留下一堆单元格
        标记（\x07）的孤儿结构，Tables.Count 不变。必须先 Tables(1).Delete()
        删掉表格结构本身，再把表格删除后残留的容器段落删掉。
        """
        try:
            b.range.Tables(1).Delete()
        except Exception:
            pass
        try:
            # 表格删除后动态 Range 塌缩到原位置，容器段落即该位置所在段落
            p = self.doc.Range(b.range.Start, b.range.Start).Paragraphs(1)
            p.Range.Delete()
        except Exception:
            pass

    def replace_text(self, old, new):
        """全文查找替换。"""
        f = self.doc.Content.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        f.Text = old
        f.Replacement.Text = new
        # 注意：pywin32 动态分发下 Execute 的关键字参数不可靠，必须按位置传全部 11 个参数
        # (FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        #  MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
        f.Execute(old, False, False, False, False, False, True, 1, False, new, 2)
        self.rebuild_ranges()
        return True

    def select_block(self, i):
        """滚动到第 i 块并闪烁高亮（纯可视化用）。"""
        b = self.blocks[i]
        self._flash(b.range)
        return True

    # ---------- 可视化 ----------

    def _flash(self, rng, seconds=0.4):
        """让一个区域滚动进视野并短暂高亮，操作过程肉眼可见。"""
        if rng is None:
            return
        try:
            rng.Select()
            rng.HighlightColorIndex = WD_YELLOW
            time.sleep(seconds)
            rng.HighlightColorIndex = WD_NO_HIGHLIGHT
        except Exception:
            pass
