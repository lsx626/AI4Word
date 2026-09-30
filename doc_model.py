"""块级文档模型：把每个写入的块登记为 Word Range，支持精确编辑。

核心思路：Word 的 Range 对象是动态的——文档别处的插入/删除会自动平移已经
存在的 Range。因此每个块记住自己的 Range 后，"改写第 i 块 / 在第 i 块后插入 /
删除第 i 块" 都能精确定位，不需要维护任何偏移量。

编辑能力分两层：
- 命令式原语（replace_block / insert_after / delete_block …）；
- 声明式编辑 apply_edit(spec) + preview_edit(spec)：spec 为
  {"op": "replace|insert_after|insert_at_end|delete", "index": i, "md": ...}，
  先预览文本级前后差异再执行，适合 AI 生成的结构化编辑指令。

段落对齐采用贪心前缀策略：段落数不一致时（用户手动改过文档），尽量对齐前面
的块并记录 alignment_warning，而不是整体放弃。
undo() 封装 Word 编辑栈；文档看门狗 DocWatch 检测流式停顿期间的用户改动。

所有编辑操作完成后会把选区移动到受影响区域并闪烁高亮，让排版过程可见。
"""
import time

import pywintypes
from markdown_it import MarkdownIt

from doc_watch import DocWatch

WD_NO_HIGHLIGHT = 0
WD_YELLOW = 7
RPC_CALL_REJECTED = -2147418111  # Word 忙（重分页/界面刷新）时拒绝调用


def _com_retry(fn, attempts=6):
    """重试被 Word 拒绝的 COM 调用（排版刚结束常常还没缓过来）。"""
    for k in range(attempts):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult != RPC_CALL_REJECTED or k == attempts - 1:
                raise
            time.sleep(0.15 * (k + 1))


def md_to_text(md):
    """把 markdown 展开为纯文本（preview_edit 的前后对比用）。

    与写入器最终落盘的文本近似一致（段落/标题/列表项按段拼接）。
    """
    md = (md or "").strip()
    if not md:
        return ""
    try:
        tokens = MarkdownIt().enable("table").parse(md)
    except Exception:
        return md
    out = []
    for tok in tokens:
        t = tok.type
        if t == "inline" and tok.children:
            parts = [ch.content for ch in tok.children
                     if ch.type in ("text", "code_inline")]
            out.append("".join(parts))
        elif t == "fence":
            out.append(tok.content.rstrip("\n"))
        elif t == "hr":
            out.append("―" * 12)
    return "\r".join(out)


class Block:
    def __init__(self, kind, level, md, text, rng, n_paras=1):
        self.kind = kind      # paragraph / heading1..3 / list / quote / code / table / hr
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
        self.alignment_warning = None  # 段落漂移时的对齐警告（给 AI 与用户看）
        self.watch = DocWatch(app, doc)

    def register(self, kind, level, md, text, rng, n_paras=1):
        b = Block(kind, level, md, text, rng, n_paras)
        self.blocks.append(b)
        return b

    def rebuild_ranges(self):
        """编辑后按段落扫描重建每块的 Range。

        Word 的 Range 自动平移在「恰好在块边界处插入」时会把新文字划进相邻块，
        不可靠；而块与段落的对应关系是我们自己写出来的（每块 n_paras 个段落），
        所以编辑后按段落重新对齐最稳妥。

        段落数不一致（用户手动改过文档）时不再整体放弃：贪心对齐尽量多的前缀
        块，并记录 alignment_warning；block_map 会把警告暴露给 AI。
        """
        try:
            paras = self.doc.Paragraphs
            total = paras.Count
        except Exception:
            return
        if not self.blocks:
            return
        idx = 0
        warning = None
        for bi, b in enumerate(self.blocks):
            if idx >= total:
                warning = f"文档比块模型短：块 {bi}（{b.kind}）起已无对应段落"
                break
            n = min(b.n_paras, total - idx)
            start = paras(idx + 1).Range.Start
            end = paras(idx + n).Range.End - 1  # 去掉末尾段落标记
            if end < start:
                end = start
            b.range = _com_retry(lambda: self.doc.Range(start, end))
            idx += n
        if warning is None and idx != total:
            warning = f"段落漂移：块模型覆盖 {idx} 段，文档实际 {total} 段"
        self.alignment_warning = warning
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
        if self.alignment_warning:
            lines.append(f"[!] {self.alignment_warning}")
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

    def undo(self, times=1):
        """撤销 Word 编辑栈中的最近操作（TypeText/Delete/InsertAfter 都在栈里）。

        块模型随后按段落贪心重建；若撤销改变了文档结构，alignment_warning
        会被设置并在 block_map 中可见。
        """
        ok = None
        try:
            ok = self.doc.Undo(times)
        except Exception:
            try:
                ok = self.doc.Undo()
            except Exception:
                try:
                    ok = self.app.Undo()
                except Exception:
                    return False
        self.rebuild_ranges()
        return bool(ok)

    # ---------- 声明式编辑 ----------

    def _require_index(self, i):
        if not isinstance(i, int) or isinstance(i, bool):
            raise ValueError(f"index 必须是整数（收到 {i!r}）")
        if i < 0 or i >= len(self.blocks):
            raise ValueError(f"索引越界: {i}（当前共 {len(self.blocks)} 块）")

    def _require_md(self, spec, op):
        md = spec.get("md")
        if not isinstance(md, str):
            raise ValueError(f"操作 {op!r} 需要字符串 md 字段")
        return md

    def preview_edit(self, spec):
        """不执行，只返回受影响块的文本级前后对比。"""
        if not isinstance(spec, dict):
            raise ValueError("spec 必须是字典")
        op = spec.get("op")
        if op == "insert_at_end":
            md = self._require_md(spec, op)
            tail = len(self.blocks) - 1
            before = self.get_block_text(tail) if tail >= 0 else ""
            return {"op": op, "affected": [tail], "before": before, "after": md_to_text(md)}
        self._require_index(spec.get("index"))
        i = spec["index"]
        if op == "delete":
            return {"op": op, "affected": [i], "before": self.get_block_text(i), "after": "（删除）"}
        if op == "replace":
            md = self._require_md(spec, op)
            return {"op": op, "affected": [i],
                    "before": self.get_block_text(i), "after": md_to_text(md)}
        if op == "insert_after":
            md = self._require_md(spec, op)
            return {"op": op, "affected": [i, i + 1],
                    "before": self.get_block_text(i), "after": md_to_text(md)}
        raise ValueError(f"未知操作: {op!r}（应为 replace/insert_after/insert_at_end/delete）")

    def apply_edit(self, spec):
        """执行声明式编辑 spec，返回执行摘要。"""
        if not isinstance(spec, dict):
            raise ValueError("spec 必须是字典")
        op = spec.get("op")
        if op not in ("replace", "insert_after", "insert_at_end", "delete"):
            raise ValueError(f"未知操作: {op!r}（应为 replace/insert_after/insert_at_end/delete）")
        if op == "insert_at_end":
            md = self._require_md(spec, op)
            blocks = self.insert_at_end(md)
            return {"op": op, "index": len(self.blocks) - len(blocks), "created": len(blocks)}
        self._require_index(spec.get("index"))
        i = spec["index"]
        if op == "delete":
            self.delete_block(i)
            return {"op": op, "index": i, "created": 0}
        md = self._require_md(spec, op)
        if op == "replace":
            blocks = self.replace_block(i, md)
        else:
            blocks = self.insert_after(i, md)
        return {"op": op, "index": i, "created": len(blocks)}

    # ---------- 文档看门狗 ----------

    def drain_events(self):
        """取出文档看门狗检测到的用户改动事件（流式停顿期间的漂移）。"""
        return self.watch.drain_events()

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