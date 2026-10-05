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
undo() 封装 Word 编辑栈；model_undo/model_redo 是块级快照撤销（整篇重写）。
事务 begin_txn/commit/rollback_txn 保证多步编辑的原子性；修订模式把 AI 的编辑
变成 Word 修订（accept/reject 由用户拍板）；块模型可序列化进文档变量（doc.Variables），
重开同一文档即恢复；段落数漂移时 realign_blocks 按文本相似度重对齐；
insert_toc/set_header/set_footer/insert_page_break 补齐长文档布局能力。

所有编辑操作完成后会把选区移动到受影响区域并闪烁高亮，让排版过程可见。
"""
import difflib
import json
import time

import pywintypes
from markdown_it import MarkdownIt

from app import debug
from doc_watch import DocWatch

WD_NO_HIGHLIGHT = 0
WD_YELLOW = 7
RPC_CALL_REJECTED = -2147418111  # Word 忙（重分页/界面刷新）时拒绝调用
BLOCKS_PROP = "AI4WordBlocks"
_PROP_CHUNK = 8000                # doc.Variables 单值实测 ≥50000 字符可用，按 8000 字符分片留余量  # 块模型持久化存入的自定义文档属性名
WD_HEADER_FOOTER_PRIMARY = 1   # wdHeaderFooterPrimary
WD_PAGE_BREAK = 7              # wdPageBreak


_real_sleep = time.sleep  # 测试会 monkeypatch time.sleep 关闭动画延迟；COM 重试的退避必须真实等待


def _com_retry(fn, attempts=6):
    """重试被 Word 拒绝的 COM 调用（排版刚结束常常还没缓过来）。"""
    for k in range(attempts):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult != RPC_CALL_REJECTED or k == attempts - 1:
                raise
            _real_sleep(0.15 * (k + 1))


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


def _block_pos(b):
    """Document position of a block range; None when unavailable."""
    try:
        return int(b.range.Start)
    except Exception:
        return None


class DocModel:
    def __init__(self, app, doc, sel, writer=None):
        self.app = app
        self.doc = doc
        self.sel = sel
        self.writer = writer
        self.blocks = []
        self.alignment_warning = None  # 段落漂移时的对齐警告（给 AI 与用户看）
        self.watch = DocWatch(app, doc)
        # 块级快照撤销 / 事务
        self._undo_stack = []
        self._redo_stack = []
        self._suppress_push = False  # 事务或 apply_edit 期间，原语不再单独入快照栈
        self._txn_active = False
        self._txn_snap = None
        self._track_prev = False

    def register(self, kind, level, md, text, rng, n_paras=1):
        b = Block(kind, level, md, text, rng, n_paras)
        at = self._insert_index(b.range)
        if at >= len(self.blocks):
            self.blocks.append(b)
        else:
            self.blocks.insert(at, b)
        debug.log("block_register", kind=kind, level=level, n_paras=n_paras,
                  total=len(self.blocks), at=at, preview=md)
        return b

    def _insert_index(self, rng):
        """Document-order insertion point for a new block range.

        Mid-document writes (select_block moves the cursor into the
        document body) must slot into the model at their document
        position, otherwise rebuild_ranges and snapshot replay diverge
        from the real layout.
        """
        try:
            start = int(rng.Start)
        except Exception:
            return len(self.blocks)
        for i in range(len(self.blocks) - 1, -1, -1):
            try:
                s = int(self.blocks[i].range.Start)
            except Exception:
                return len(self.blocks)
            if s <= start:
                return i + 1
        return 0

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
                # 失配块必须失活：保留过时的活 Range，refresh_texts 会读出
                # 错误文本喂给 block_map，之后的 delete/replace 会在错位
                # 内容上操作（对照 realign_blocks 的安全做法）
                b.range = None
                continue
            n = min(b.n_paras, total - idx)
            try:
                start = paras(idx + 1).Range.Start
                end = paras(idx + n).Range.End - 1  # 去掉末尾段落标记
            except Exception:
                # 排版后重分页期 COM 常报 RPC_CALL_REJECTED，与 _com_retry 同源
                start = _com_retry(lambda: paras(idx + 1).Range.Start)
                end = _com_retry(lambda: paras(idx + n).Range.End) - 1
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
            if b.range is None:
                continue  # 重对齐中失活的块
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
        self._require_index(i)
        b = self.blocks[i]
        try:
            if b.range is not None:
                return b.range.Text
        except Exception:
            pass
        return b.text

    # ---------- 编辑原语（会注入到 AI 生成代码的全局变量里） ----------

    def replace_block(self, i, md):
        """用新的 markdown 改写第 i 块（原段落保留，清空内容后重写）。"""
        debug.log("replace_block", index=i, md_preview=md)
        self._require_index(i)
        self._require_md_str(md, "replace_block")
        self._push_undo()
        b = self.blocks[i]
        if b.range is None:
            raise ValueError(f"块 {i} 的 Range 已失活（文档与块模型脱节），拒绝写入")
        if b.kind == "table":
            # 表格块：删掉表格本身，保留其后的容器段落作为新内容的写入点
            try:
                _com_retry(lambda: b.range.Tables(1).Delete())
            except Exception:
                b.range.Delete()
        else:
            b.range.Delete()
        pos = b.range.Start  # 动态 Range 在删除后塌缩到删除点
        self.doc.Range(pos, pos).Select()
        new_blocks = self.writer.write_block(md, animate=True)
        if not new_blocks:
            # write_block 对空 md 返回 [] 会留下孤儿空段落且块索引错位；
            # _require_md_str 已在入口拒绝空 md，这里防御性兜底
            raise ValueError("replace_block 写入失败：未登记任何新块")
        # register() may have inserted the new blocks at a document
        # position rather than the tail, so collect them by identity
        new_ids = {id(nb) for nb in new_blocks}
        self.blocks = [b for b in self.blocks if id(b) not in new_ids]
        self.blocks[i:i + 1] = new_blocks
        self.rebuild_ranges()
        self._flash(new_blocks[-1].range if new_blocks else None)
        return new_blocks

    def insert_after(self, i, md):
        """在第 i 块之后插入新内容。"""
        debug.log("insert_after", index=i, md_preview=md)
        self._require_index(i)
        self._require_md_str(md, "insert_after")
        self._push_undo()
        b = self.blocks[i]
        if b.range is None:
            raise ValueError(f"块 {i} 的 Range 已失活（文档与块模型脱节），拒绝写入")
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
        new_blocks = self.writer.write_block(md, animate=True)
        if not new_blocks:
            raise ValueError("insert_after 写入失败：未登记任何新块")
        new_ids = {id(nb) for nb in new_blocks}
        self.blocks = [b for b in self.blocks if id(b) not in new_ids]
        self.blocks[i + 1:i + 1] = new_blocks
        self.rebuild_ranges()
        self._flash(new_blocks[-1].range if new_blocks else None)
        return new_blocks

    def insert_at_end(self, md):
        """在文档末尾追加新内容。"""
        debug.log("insert_at_end", md_preview=md)
        self._push_undo()
        if not self.blocks:
            return self.writer.write_block(md, animate=True)
        return self.insert_after(len(self.blocks) - 1, md)

    def delete_block(self, i):
        """删除第 i 块（连同它占用的全部段落与段落标记）。"""
        debug.log("delete_block", index=i)
        self._require_index(i)
        self._push_undo()
        b = self.blocks[i]
        if b.range is None:
            raise ValueError(f"块 {i} 的 Range 已失活（文档与块模型脱节），拒绝写入")
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
        ok_table = True
        try:
            _com_retry(lambda: b.range.Tables(1).Delete())
        except Exception:
            ok_table = False
        try:
            # 表格删除后动态 Range 塌缩到原位置，容器段落即该位置所在段落
            p = self.doc.Range(b.range.Start, b.range.Start).Paragraphs(1)
            p.Range.Delete()
        except Exception:
            if not ok_table:
                # 两步都失败（COM 瞬断/Word 忙）：不能仍从模型里弹块——
                # 文档里表格还在、模型少一块，之后所有索引整体错位
                raise RuntimeError("删除表格失败：Word 拒绝了删除请求，请重试")
            return
        if not ok_table:
            # 表格删除失败但容器段落已删：仍可能留下 \x07 孤儿，明确返回而非静默
            raise RuntimeError("表格结构删除失败，留下单元格标记孤儿结构")

    def replace_text(self, old, new):
        """全文查找替换；返回是否真的替换到内容。"""
        debug.log("replace_text", old_preview=old, new_preview=new)
        if not isinstance(old, str) or not isinstance(new, str):
            raise ValueError("replace_text 的 old/new 必须是字符串")
        if not old:
            raise ValueError("replace_text 的 old 不能为空字符串")
        self._push_undo()

        def _cur_text(b):
            if b.range is None:
                return b.text
            try:
                return b.range.Text
            except Exception:
                return b.text

        old_texts = {id(b): _cur_text(b) for b in self.blocks}
        f = self.doc.Content.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        f.Text = old
        f.Replacement.Text = new
        # 注意：pywin32 动态分发下 Execute 的关键字参数不可靠，必须按位置传全部 11 个参数
        # (FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        #  MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
        replaced = f.Execute(old, False, False, False, False, False, True, 1, False, new, 2)
        self.rebuild_ranges()
        # 回写受影响块的 md：快照/恢复以 md 为准，不回写则 rollback 会把
        # replace_text 的结果静默还原（代价是丢失 md 里的 ** 等 markdown
        # 标记——换回保存 replace 的修改）
        for b in self.blocks:
            cur = _cur_text(b)
            if old_texts.get(id(b)) != cur:
                b.md = self._md_from_text(b, cur)
        return bool(replaced)

    @staticmethod
    def _md_from_text(b, text):
        if b.kind.startswith("heading") and len(b.kind) > 7:
            lv = b.kind[7:]
            return "#" * int(lv) + " " + text
        return text

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

    @staticmethod
    def _require_md_str(md, op):
        """命令式原语的 md 校验：非字符串或空白一律拒绝。

        空 md 的 write_block 返回 []：replace 会留下孤儿空段落且块索引
        错位、insert 留下孤儿段落——都是静默的数据损坏，不如明确报错
        让 AI 的自我修复路径纠正。
        """
        if not isinstance(md, str):
            raise ValueError(f"{op} 需要 markdown 字符串")
        if not md.strip():
            raise ValueError(f"{op} 的内容不能为空")
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
        """执行声明式编辑 spec，返回执行摘要（整次编辑作为一条快照撤销记录）。"""
        prev = self._suppress_push
        self._push_undo()
        self._suppress_push = True
        try:
            return self._apply_edit(spec)
        finally:
            self._suppress_push = prev

    def _apply_edit(self, spec):
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
        self._require_index(i)
        b = self.blocks[i]
        debug.log("select_block", index=i, kind=b.kind)
        self._flash(b.range)
        return True


    # ---------- 修订模式（人机协同：AI 的修改以 Word 修订呈现，可接受/拒绝） ----------

    def review_on(self):
        """开启修订模式：此后本模型的编辑都成为 Word 修订，用户可逐条接受/拒绝。"""
        try:
            self._track_prev = bool(self.doc.TrackRevisions)
        except Exception:
            self._track_prev = False
        try:
            self.doc.TrackRevisions = True
        except Exception:
            pass
        debug.log("review_on", prev_track=self._track_prev)
        return True

    def review_off(self):
        """关闭修订模式，恢复开启前的状态。"""
        try:
            self.doc.TrackRevisions = self._track_prev
        except Exception:
            pass
        debug.log("review_off")
        return True

    def has_revisions(self):
        """文档中是否还有未决修订。"""
        try:
            return int(self.doc.Revisions.Count) > 0
        except Exception:
            return False

    def accept_block_revisions(self, i):
        """接受第 i 块范围内的全部修订，随后重建块模型。"""
        return self._resolve_block_revisions(i, accept=True)

    def reject_block_revisions(self, i):
        """拒绝第 i 块范围内的全部修订（内容回滚），随后重建块模型。"""
        return self._resolve_block_revisions(i, accept=False)

    def _resolve_block_revisions(self, i, accept):
        b = self.blocks[i]
        n = 0
        try:
            revs = b.range.Revisions
            count = int(revs.Count)
            # 倒序遍历：每处理一条，集合就变一次
            for k in range(count, 0, -1):
                try:
                    r = revs(k)
                except Exception:
                    continue
                if accept:
                    r.Accept()
                else:
                    r.Reject()
                n += 1
        except Exception:
            pass
        self.rebuild_ranges()
        return n

    # ---------- 块模型持久化（随文档保存，重开同一文档可恢复） ----------

    def save_blocks(self):
        """把块模型序列化进文档变量（doc.Variables）；返回块数。

        文档变量随文档一起保存，下次打开同一文档时 load_blocks()
        按 n_paras 重新对齐段落、重建 Range，块索引语义跨会话保留。
        doc.Variables 单值实测可容 ≥50000 字符，为稳妥起见仍分片：
        BLOCKS_PROP 存分片总数（字符串），BLOCKS_PROP_0..n-1 存各分片。
        """
        data = [{"kind": b.kind, "level": b.level, "md": b.md,
                 "text": b.text, "n_paras": b.n_paras} for b in self.blocks]
        payload = json.dumps(data, ensure_ascii=False)
        n = max(1, (len(payload) + _PROP_CHUNK - 1) // _PROP_CHUNK)
        # 先写全部分片，最后写计数：计数是「提交标志」，先写计数时中途
        # COM 瞬断会留下计数与分片不一致（load_blocks 整档判废）
        for k in range(n):
            self._set_doc_variable(f"{BLOCKS_PROP}_{k}",
                                      payload[k * _PROP_CHUNK:(k + 1) * _PROP_CHUNK])
        self._set_doc_variable(BLOCKS_PROP, str(n))
        # 本次块数变少时，清理残留的编号更大的旧分片
        k = n
        while self._clear_doc_variable(f"{BLOCKS_PROP}_{k}"):
            k += 1
        debug.log("save_blocks", blocks=len(data), chunks=n)
        return len(data)

    def _set_doc_variable(self, name, value):
        """写入一个文档变量；已存在则覆盖。"""
        try:
            self.doc.Variables(name).Value = value
            return
        except Exception:
            pass  # 尚无此变量，走 Add
        try:
            # 真实签名 Variables.Add(Name, Value)；同名已存在会抛错
            self.doc.Variables.Add(name, value)
        except Exception:
            # 变量已存在但 .Value 赋值失败时，删除后重建
            try:
                self.doc.Variables(name).Delete()
            except Exception:
                pass
            self.doc.Variables.Add(name, value)

    def _clear_doc_variable(self, name):
        """删除某个文档变量；不存在时返回 False。"""
        try:
            self.doc.Variables(name).Delete()
            return True
        except Exception:
            return False

    def load_blocks(self):
        """从文档变量恢复块模型；返回恢复的块数（0 表示没有存档）。"""
        try:
            n = int(self.doc.Variables(BLOCKS_PROP).Value)
        except Exception:
            return 0
        if n <= 0:
            return 0
        parts = []
        for k in range(n):
            try:
                parts.append(self.doc.Variables(f"{BLOCKS_PROP}_{k}").Value)
            except Exception:
                return 0
        payload = "".join(parts)
        try:
            data = json.loads(payload)
        except Exception:
            return 0
        if not isinstance(data, list):
            return 0
        blocks = []
        for d in data:
            if not isinstance(d, dict):
                continue
            n_paras = d.get("n_paras", 1)
            # 损坏存档（手改 / 同步冲突）的 n_paras 可能不是正整数：
            # rebuild_ranges 的 min(b.n_paras, total-idx) 会抛 TypeError，
            # 此时不能污染 self.blocks（CLI 启动路径无 try/except）
            if not isinstance(n_paras, int) or isinstance(n_paras, bool) or n_paras <= 0:
                return 0
            blocks.append(Block(d.get("kind", "paragraph"), d.get("level", 0),
                                d.get("md", ""), d.get("text", ""),
                                None, n_paras))
        if not blocks:
            return 0
        self.blocks = blocks
        try:
            self.rebuild_ranges()
            # 存档与现状不一致（别的会话里改过文档）时按相似度重对齐
            if self.alignment_warning:
                self.realign_blocks()
        except Exception:
            self.blocks = []  # 重建失败：失效存档，交给调用方走 import_document
            return 0
        debug.log("load_blocks", blocks=len(self.blocks))
        return len(self.blocks)

    # ---------- 首次接入已有文档：把现成段落登记为块 ----------

    def import_document(self, limit=600):
        """把当前文档的现有内容登记为块模型（读取已存在的非空文档）。

        新打开的非空文档没有 AI4Word 存档（load_blocks 返回 0）时调用：
        逐段落登记——标题样式识别为 heading1-3，其余为 paragraph；
        空段落同样登记（保持块模型与文档段落数 1:1 对齐，rebuild_ranges
        不会报漂移）。之后块地图、replace_block / insert_after / delete_block
        等原语即可直接作用于已有内容。
        """
        self.blocks = []
        try:
            paras = self.doc.Paragraphs
            total = paras.Count
        except Exception:
            return 0
        if total <= 1 and self._doc_is_empty():
            return 0
        # 内建标题样式的本地名（中文 Word 是「标题 1」，英文是「Heading 1」）
        heading_names = {}
        for sid, level in ((-2, 1), (-3, 2), (-4, 3), (-5, 4), (-6, 5), (-7, 6)):
            try:
                name = str(self.doc.Styles(sid).NameLocal).strip().lower()
                if name:
                    heading_names[name] = level
            except Exception:
                pass
        n = 0
        for i in range(1, min(total, limit) + 1):
            try:
                rng = paras(i).Range
            except Exception:
                continue
            text = ""
            try:
                text = str(rng.Text or "")
            except Exception:
                pass
            raw = text.rstrip("\r\n\x07 \t　")
            kind = "paragraph"
            level = 0
            try:
                lname = str(rng.Style.NameLocal or "").strip().lower()
                lv = heading_names.get(lname)
                if lv is not None:
                    kind = f"heading{lv}"
                    level = lv
            except Exception:
                pass
            # 与流式块登记 / rebuild_ranges 一致：Range 不含末尾段落标记，
            # 否则 replace_block 删块时会连带吞掉段落分隔符
            start = rng.Start
            end = max(rng.End - 1, start)
            try:
                block_range = self.doc.Range(start, end)
            except Exception:
                block_range = rng
            try:
                self.register(kind, level, raw, raw, block_range, 1)
                n += 1
            except Exception:
                continue
        self.alignment_warning = None
        if total > limit:
            self.alignment_warning = f"文档较长（{total} 段），仅登记前 {limit} 段"
        debug.log("import_document", paras=total, imported=n)
        return n

    def _doc_is_empty(self):
        """文档是否没有任何可见内容（只有段落标记）。"""
        try:
            return not (self.doc.Content.Text or "").strip("\r\n\x07 \t　")
        except Exception:
            return True

    # ---------- 漂移后的模糊重对齐（用户手改文档，段落数已变） ----------

    def realign_blocks(self):
        """按文本相似度把块重新钉到当前段落上，返回成功对齐的块数。

        rebuild_ranges 的贪心前缀只接受"块在前、段落一一对应"的理想情况；
        用户手动增删段落后，用 difflib 为每个块找回位置，找不到的块 Range
        置空（block_map 会显示警告）。
        """
        try:
            paras = self.doc.Paragraphs
            total = paras.Count
        except Exception:
            return 0
        if not self.blocks or total <= 0:
            return 0
        para_texts = []
        for k in range(1, total + 1):
            try:
                para_texts.append(paras(k).Range.Text.replace("\r", " ").replace("\n", " ").strip())
            except Exception:
                para_texts.append("")
        matched = 0
        pk = 0
        for b in self.blocks:
            bt = (b.text or "").replace("\r", " ").replace("\n", " ").strip()
            best = None
            for start in range(pk, min(pk + 10, total)):
                for cnt in (1, 2, 3, 4):
                    if start + cnt > total:
                        break
                    cand = " ".join(t for t in para_texts[start:start + cnt] if t).strip()
                    ratio = difflib.SequenceMatcher(None, bt, cand).ratio()
                    if best is None or ratio > best[0]:
                        best = (ratio, start, cnt)
            if best is not None and best[0] >= 0.55:
                _, start, cnt = best
                s = paras(start + 1).Range.Start
                e = paras(start + cnt).Range.End - 1
                if e < s:
                    e = s
                b.range = _com_retry(lambda: self.doc.Range(s, e))
                b.n_paras = cnt
                pk = start + cnt
                matched += 1
            else:
                b.range = None
        self.alignment_warning = None if matched == len(self.blocks) else \
            f"重对齐：{matched}/{len(self.blocks)} 块找到对应段落，其余块已失去定位"
        self.refresh_texts()
        return matched

    # ---------- 块级快照 undo/redo（整篇重写级，区别于 Word 的操作栈） ----------

    def snapshot(self):
        """当前块模型的快照（各块的 md 与结构），可传给 restore_snapshot。"""
        debug.log("snapshot", blocks=len(self.blocks))
        return [{"kind": b.kind, "level": b.level, "md": b.md,
                 "text": b.text, "n_paras": b.n_paras,
                 "pos": _block_pos(b)} for b in self.blocks]

    def restore_snapshot(self, snap):
        """按快照整篇重写文档（块 md 顺序重写），返回重建的块数。

        model_undo / 事务回滚用它。注意：快照恢复会丢弃用户在 AI 块之外
        手动做的格式调整——这是"结构一致"与"保留一切手动痕迹"之间的取舍。
        任一块重写失败（COM 瞬断等）会抛 RuntimeError——文档已清空、按
        逐块写入到底写了多少是不确定状态，不能静默报「已回滚成功」。
        """
        debug.log("restore_snapshot", blocks=len(snap or []))
        if self.writer is None:
            raise RuntimeError("没有可用的写入器，无法恢复快照")
        snap = list(snap or [])
        # Replay in document order so the rebuilt layout matches the
        # pre-write one even when a block was written mid-document.
        snap.sort(key=lambda item: (item.get("pos") is None, item.get("pos") or 0))
        try:
            self.doc.Content.Delete()
        except Exception:
            pass
        try:
            self.doc.Range(0, 0).Select()
        except Exception:
            pass
        self.blocks = []
        w = self.writer
        w.pending = ""
        w._draft = None
        w._first_block = True
        failed = 0
        for idx, item in enumerate(snap):
            if idx > 0:
                try:
                    # 块之间的段落分隔：流式路径里由 _write_top_level 的 started 逻辑补，
                    # 逐块重写时调用方（此处）负责补上，否则相邻块会并进同一段落
                    w.sel.TypeParagraph()
                except Exception:
                    pass
            try:
                w.write_block(item.get("md", ""), animate=False)
            except Exception as e:
                # 不吞：调用方/engine 会把失败报给用户；数一下好给出定位信息
                failed += 1
                last_err = e
        if failed:
            self.alignment_warning = None
            self.rebuild_ranges()
            raise RuntimeError(
                f"恢复快照时 {failed}/{len(snap)} 个块写入失败"
                f"（最后错误：{last_err}），文档可能不完整")
        self.alignment_warning = None  # 全新对齐，旧的漂移警告作废
        self.rebuild_ranges()
        debug.log("restore_snapshot_done", blocks=len(self.blocks))
        return len(self.blocks)

    def model_undo(self):
        """快照撤销：回到上一个编辑前的块模型状态。"""
        debug.log("model_undo", stack=len(self._undo_stack))
        if not self._undo_stack:
            return False
        cur = self.snapshot()
        nxt = self._undo_stack.pop()
        self._redo_stack.append(cur)
        self.restore_snapshot(nxt)
        return True

    def model_redo(self):
        """快照重做（撤销之后才能用）。"""
        debug.log("model_redo", stack=len(self._redo_stack))
        if not self._redo_stack:
            return False
        cur = self.snapshot()
        nxt = self._redo_stack.pop()
        self._undo_stack.append(cur)
        self.restore_snapshot(nxt)
        return True

    def _push_undo(self):
        """编辑前压入快照；事务进行中或被上层抑制时跳过（避免双重入栈）。"""
        if self._suppress_push:
            return
        self._undo_stack.append(self.snapshot())
        if len(self._undo_stack) > 50:
            del self._undo_stack[0]
        self._redo_stack.clear()

    # ---------- 事务：多步编辑的原子性 ----------

    def begin_txn(self):
        """开启事务：到 commit/rollback 之间的编辑是一个原子单元。"""
        debug.log("begin_txn")
        if self._txn_active:
            raise RuntimeError("已有进行中的事务")
        self._txn_snap = self.snapshot()
        self._push_undo()
        self._txn_active = True
        self._suppress_push = True  # 事务内的原语不再单独入快照栈
        return True

    def commit_txn(self):
        """提交事务：编辑保留（整条事务已作为一条快照撤销记录入栈）。"""
        debug.log("commit_txn")
        if not self._txn_active:
            raise RuntimeError("没有进行中的事务")
        self._txn_active = False
        self._suppress_push = False
        self._txn_snap = None
        return True

    def is_txn_active(self):
        """事务是否进行中（engine 在编辑失败时据此主动回滚卡住的事务）。"""
        return self._txn_active

    def rollback_txn(self):
        """回滚事务：文档与块模型恢复到事务开始前。"""
        debug.log("rollback_txn")
        if not self._txn_active:
            raise RuntimeError("没有进行中的事务")
        snap = self._txn_snap
        # 先复位标志再恢复：restore_snapshot 现在会把失败抛出来，
        # 此处不能因为恢复失败就把 _txn_active 卡在 True
        # （之后所有编辑都不入快照栈、再次 begin_txn 也会报错）
        self._txn_active = False
        self._suppress_push = False
        self._txn_snap = None
        try:
            self.restore_snapshot(snap)
        except Exception:
            # 文档可能停在半写状态：快照已丢，至少把块模型重建到现状
            try:
                self.rebuild_ranges()
                self.refresh_texts()
            except Exception:
                pass
            raise
        return True

    # ---------- 长文档布局：目录 / 页眉页脚 / 分页 ----------

    def insert_toc(self):
        """在文档末尾插入目录（基于标题级别 1-3，Word 域自动生成）。"""
        try:
            end = self.doc.Content.End
        except Exception:
            end = 0
        pos = max(0, end - 1)  # 文档以段落标记结尾，末尾前一位置更稳
        try:
            self.doc.Range(pos, pos).Select()
            self.sel.TypeParagraph()
            pos = self.sel.Range.Start
        except Exception:
            pass
        toc = self.doc.TablesOfContents.Add(self.doc.Range(pos, pos), True, 1, 3)
        try:
            toc.Update()
        except Exception:
            pass
        try:
            self.sel.Range.Select()
        except Exception:
            pass
        return True

    def set_header(self, text):
        """设置第一节页眉文本。"""
        h = self.doc.Sections(1).Headers(WD_HEADER_FOOTER_PRIMARY)
        h.Range.Text = text
        return True

    def set_footer(self, text):
        """设置第一节页脚文本。"""
        f = self.doc.Sections(1).Footers(WD_HEADER_FOOTER_PRIMARY)
        f.Range.Text = text
        return True

    def insert_page_break(self):
        """在当前光标处插入分页符。"""
        self.sel.InsertBreak(WD_PAGE_BREAK)
        self.rebuild_ranges()
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
            # 闪罩后把光标恢复为块尾的折叠点：留着覆盖整个块的选区，
            # 后续流式写入会 用 TypeParagraph/TypeText 替换掉块内容
            # （含内联图片）。
            try:
                self.doc.Range(rng.End, rng.End).Select()
            except Exception:
                pass
        except Exception:
            pass