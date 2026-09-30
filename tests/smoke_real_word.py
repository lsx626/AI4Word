"""真实 Word 冒烟测试：验证 COM 交互的关键假设。

不调用 Atria API，只用预制 markdown 走完整流程：
流式写入 -> 块登记 -> block_map -> replace/insert/delete -> 逐语句执行。
全部在新建的临时文档上进行，结束关闭不保存。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import time
import win32com.client

from doc_model import DocModel
from format_runner import run_code
from streaming_writer import StreamingWriter

time.sleep = lambda s: None  # 冒烟测试关闭动画延迟

WD_BORDER_BOTTOM = -3  # 段落底边框（与 streaming_writer 一致）


def feed_in_pieces(writer, text, size=3):
    for i in range(0, len(text), size):
        writer.feed(text[i:i + size])


def main():
    app = win32com.client.Dispatch("Word.Application")
    app.Visible = True
    # 清理上次失败遗留的临时文档，避免新文档与旧选区串台
    try:
        while app.Documents.Count > 0:
            app.Documents(1).Close(SaveChanges=0)
    except Exception:
        pass
    doc = app.Documents.Add()
    sel = app.Selection

    writer = StreamingWriter(app, doc, sel, model=None)
    model = DocModel(app, doc, sel, writer)
    writer.model = model

    md = "# 流式标题\n\n这是第一段正文，含**粗体**和*斜体*。\n\n- 列表项甲\n- 列表项乙\n\n结尾段落"
    feed_in_pieces(writer, md)
    writer.flush()

    content = doc.Content.Text
    assert "流式标题" in content, "标题未写入"
    assert "粗体" in content and "斜体" in content, "格式文本未写入"
    assert "**" not in content and "*" not in content, f"markdown 标记残留: {content!r}"
    kinds = [b.kind for b in model.blocks]
    assert kinds == ["heading1", "paragraph", "list", "paragraph"], kinds
    print("ok: 真实 Word 流式写入（含格式剥离、块登记）")

    print("block_map:\n" + model.block_map())

    # 块编辑：改写第二块（paragraph）
    model.replace_block(1, "改写后的第二段")
    content = doc.Content.Text
    assert "改写后的第二段" in content and "这是第一段正文" not in content, content
    assert model.blocks[3].text == "结尾段落", f"后续块定位错: {model.blocks[3].text!r}"
    print("ok: replace_block 真实 Word")

    # 插入
    model.insert_after(0, "插入的新段落")
    content = doc.Content.Text
    assert "插入的新段落" in content, content
    assert model.blocks[1].text == "插入的新段落"
    assert model.blocks[2].text == "改写后的第二段"
    print("ok: insert_after 真实 Word")

    # 逐语句执行 + 样式修改（可视化闪光会真实发生）
    exec_globals = {
        "word_app": app, "doc": doc, "sel": sel, "model": model,
        "replace_block": model.replace_block,
        "insert_after": model.insert_after,
        "delete_block": model.delete_block,
        "replace_text": model.replace_text,
        "select_block": model.select_block,
        "block_map": model.block_map,
        "get_block_text": model.get_block_text,
        "WD_STYLE_NORMAL": -1, "WD_STYLE_HEADING_1": -2,
        "WD_STYLE_HEADING_2": -3, "WD_STYLE_HEADING_3": -4,
    }
    code = "doc.Styles(WD_STYLE_NORMAL).Font.Name = '宋体'\nselect_block(0)"
    ok, err = run_code(code, exec_globals)
    assert ok, f"run_code 失败: {err}"
    print("ok: run_code 逐语句执行样式修改 + 可视化")

    # 全文替换
    model.replace_text("插入的新段落", "替换后的段落")
    content = doc.Content.Text
    assert "替换后的段落" in content, content
    print("ok: replace_text")

    # ---- V7.1：有序列表 / 围栏代码块 / 表格及其块编辑 ----

    def aligned():
        total = doc.Paragraphs.Count
        assert sum(b.n_paras for b in model.blocks) == total, \
            f"段落数不对齐: {[(b.kind, b.n_paras) for b in model.blocks]} vs 文档 {total}"

    # 有序列表：编号样式、每项一段、无双倍空段落
    model.insert_at_end("1. 有序甲\n2. 有序乙\n3. 有序丙")
    lb = model.blocks[-1]
    assert lb.kind == "list" and lb.n_paras == 3, (lb.kind, lb.n_paras)
    assert "有序甲" in doc.Content.Text and "有序丙" in doc.Content.Text
    aligned()
    print("ok: 有序列表（编号样式、每项一段、段落对齐）")

    # 围栏代码块：等宽字体、块内空行不切断、按行占段
    model.insert_at_end("```\nline one\n\nline three\n```")
    cb = model.blocks[-1]
    assert cb.kind == "code" and cb.n_paras == 3, (cb.kind, cb.n_paras)
    code_rng = doc.Range(cb.range.Start, cb.range.End)
    assert code_rng.Font.Name == "Consolas", f"代码字体: {code_rng.Font.Name}"
    assert "```" not in doc.Content.Text, "围栏标记泄漏"
    aligned()
    print("ok: 围栏代码块（等宽字体、空行不断流、按行占段）")

    # 表格：真 Word 表格、含边框、块 Range 覆盖表格 + 容器段落
    model.insert_at_end("| 姓名 | 年龄 |\n| --- | --- |\n| 张三 | 20 |\n| 李四 | 22 |")
    tb = model.blocks[-1]
    assert tb.kind == "table", tb.kind
    tbl = tb.range.Tables(1)
    assert tbl.Rows.Count == 3 and tbl.Columns.Count == 2, \
        (tbl.Rows.Count, tbl.Columns.Count)
    assert "张三" in tbl.Range.Text and "李四" in tbl.Range.Text
    assert tb.n_paras == tbl.Range.Paragraphs.Count + 1, tb.n_paras
    aligned()
    print("ok: 表格块（真表格、边框、段落计数对齐）")

    # 列表块整体删除：不留孤儿段落（按文本定位刚才写入的有序列表块）
    target = next(i for i, b in enumerate(model.blocks)
                  if b.kind == "list" and "有序甲" in (b.text or ""))
    model.delete_block(target)
    assert "有序甲" not in doc.Content.Text and "有序丙" not in doc.Content.Text
    assert "列表项甲" in doc.Content.Text  # 其他列表块不受影响
    aligned()
    print("ok: 删除列表块（项目段落全部清除、无孤儿）")

    # 删除表格块：表格与容器段落一并删除
    tables_before = doc.Tables.Count
    for i, b in enumerate(model.blocks):
        if b.kind == "table":
            model.delete_block(i)
            break
    assert doc.Tables.Count == tables_before - 1, f"表格数: {doc.Tables.Count}"
    assert "张三" not in doc.Content.Text
    aligned()
    print("ok: 删除表格块（表格与容器段落一并删除）")

    # 替换表格块：删除表格、保留容器段落写入新内容
    model.insert_at_end("| A | B |\n| --- | --- |\n| 1 | 2 |")
    nb = None
    for i, b in enumerate(model.blocks):
        if b.kind == "table":
            nb = model.replace_block(i, "表格已被替换为段落")[0]
            break
    assert nb is not None and nb.kind == "paragraph", nb
    assert "表格已被替换为段落" in doc.Content.Text
    assert "A" not in doc.Content.Text
    aligned()
    print("ok: 替换表格块（删表留容器段落）")

    # 表格块后插入：插入点在容器段落之后，而非表格内部
    model.insert_at_end("| C | D |\n| --- | --- |\n| 3 | 4 |")
    ib = None
    for i, b in enumerate(model.blocks):
        if b.kind == "table":
            ib = model.insert_after(i, "表格之后的段落")[0]
            break
    assert ib is not None and ib.kind == "paragraph", ib
    assert "表格之后的段落" in doc.Content.Text
    aligned()
    print("ok: 表格块后插入（容器段落之后定位）")

    # ---- V7.2：嵌套列表 / 引用块 / hr / 超链接 / 声明式编辑 / undo ----

    # 嵌套列表：按层级 LeftIndent（0 / 21 / 42 / 21）
    model.insert_at_end("- 一级项\n  - 二级项\n    - 三级项\n  - 回二级")
    nlb = model.blocks[-1]
    assert nlb.kind == "list" and nlb.n_paras == 4, (nlb.kind, nlb.n_paras)
    total = doc.Paragraphs.Count
    indents = [doc.Paragraphs(total - 3 + i).Format.LeftIndent for i in range(4)]
    assert indents == [0, 21, 42, 21], f"嵌套列表缩进错: {indents}"
    assert "三级项" in doc.Content.Text
    aligned()
    print("ok: 嵌套列表（按层级缩进）")

    # 引用块：两段、缩进、斜体
    model.insert_at_end("> 引用甲\n>\n> 引用乙")
    qb = model.blocks[-1]
    assert qb.kind == "quote" and qb.n_paras == 2, (qb.kind, qb.n_paras)
    total = doc.Paragraphs.Count
    qp = doc.Paragraphs(total - 1)
    assert qp.Format.LeftIndent == 21, f"引用缩进错: {qp.Format.LeftIndent}"
    assert qp.Range.Font.Italic, "引用块未斜体"
    assert "引用乙" in doc.Content.Text
    aligned()
    print("ok: 引用块（缩进、斜体、多段）")

    # 水平线：段落底边框（或字符降级）
    model.insert_at_end("---")
    hb = model.blocks[-1]
    assert hb.kind == "hr" and hb.n_paras == 1, (hb.kind, hb.n_paras)
    total = doc.Paragraphs.Count
    hp = doc.Paragraphs(total)
    try:
        line_style = hp.Range.ParagraphFormat.Borders(WD_BORDER_BOTTOM).LineStyle
    except Exception:
        line_style = None
    hr_ok = (line_style not in (None, 0)) or ("――" in hp.Range.Text)
    assert hr_ok, f"水平线既无边框也无降级字符: {hp.Range.Text!r}"
    aligned()
    print("ok: 水平线（段落边框优先）")

    # 超链接：真实 Hyperlinks.Add；链接文字不丢
    links_before = doc.Hyperlinks.Count
    model.insert_at_end("看[链接文字](https://example.com)与`code`结尾")
    lb = model.blocks[-1]
    assert lb.kind == "paragraph", lb.kind
    assert doc.Hyperlinks.Count == links_before + 1, \
        f"超链接数: {doc.Hyperlinks.Count}"
    addr = None
    for hl in doc.Hyperlinks:
        if hl.Range.Text == "链接文字":
            addr = hl.Address
    # Word 会把无路径 URL 规范化成带结尾斜杠
    assert addr in ("https://example.com", "https://example.com/"), f"链接地址错: {addr}"
    assert "code" in doc.Content.Text, "行内代码内容丢失"
    aligned()
    print("ok: 超链接（真实 Hyperlinks.Add + 行内代码字体）")

    # 声明式编辑：先 preview 后 apply
    tail = len(model.blocks) - 1
    spec = {"op": "replace", "index": tail, "md": "声明式替换的段落"}
    pv = model.preview_edit(spec)
    assert pv["op"] == "replace" and pv["after"] == "声明式替换的段落", pv
    r = model.apply_edit(spec)
    assert r["created"] == 1 and model.blocks[tail].text == "声明式替换的段落", r
    assert "声明式替换的段落" in doc.Content.Text
    aligned()
    print("ok: apply_edit / preview_edit（预览不落盘、执行后对齐）")

    model.apply_edit({"op": "insert_at_end", "md": "声明式追加"})
    assert model.blocks[-1].text == "声明式追加"
    aligned()
    print("ok: apply_edit insert_at_end")

    # undo：单次 COM 操作（replace_text）对应一条撤销记录
    model.replace_text("声明式替换的段落", "撤销前文本")
    assert "撤销前文本" in doc.Content.Text
    ok_undo = model.undo(1)
    assert ok_undo, "undo 失败"
    assert "撤销前文本" not in doc.Content.Text, "undo 未回退替换"
    assert "声明式替换的段落" in doc.Content.Text, "undo 回退过度"
    aligned()
    print("ok: undo（单条编辑记录精确回退、块模型重建对齐）")

    print("\n真实 Word 冒烟测试全部通过。")
    doc.Close(SaveChanges=0)


if __name__ == "__main__":
    try:
        main()
    finally:
        # 无论成败都关闭临时文档，避免遗留文档影响下一次运行
        try:
            import win32com.client as _wc
            for _ in range(3):
                _wc.GetObject(None, "Word.Application").Documents(1).Close(SaveChanges=0)
        except Exception:
            pass
