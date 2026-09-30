"""真实 Word 冒烟测试：验证 COM 交互的关键假设。

不调用 DeepSeek API，只用预制 markdown 走完整流程：
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


def feed_in_pieces(writer, text, size=3):
    for i in range(0, len(text), size):
        writer.feed(text[i:i + size])


def main():
    app = win32com.client.Dispatch("Word.Application")
    app.Visible = True
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

    print("\n真实 Word 冒烟测试全部通过。")
    doc.Close(SaveChanges=0)


if __name__ == "__main__":
    try:
        main()
    finally:
        pass
