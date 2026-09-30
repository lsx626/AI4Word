# -*- coding: utf-8 -*-
"""离线测试：流式写入的光标漂移防护（写入锚点）。

用户在生成期间点击文档别处把光标挪走时，后续内容必须仍然接在已写
内容的末尾，而不是跳到用户的光标位置。用 FakeWord（动态 Range 语义）
验证草稿态与结构块两条写入路径。
"""
import importlib.util
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

_spec = importlib.util.spec_from_file_location(
    "fake_word_mod", os.path.join(os.path.dirname(__file__), "test_fake_word.py"))
_fake = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_fake)

from streaming_writer import StreamingWriter  # noqa: E402


def _hijack(app, pos):
    """模拟用户在生成期间点击文档别处：光标被挪走。"""
    app.sel.pos = pos


def test_draft_hijack_protection():
    app, writer, model = _fake.make()
    writer.reset_anchor()  # 引擎在一波写入开始前重置锚点
    _fake.feed_in_pieces(writer, "第一段正文")
    # 用户此刻点了文档开头：光标挪到 0
    _hijack(app, 0)
    _fake.feed_in_pieces(writer, "第二段续写内容")
    writer.flush()
    # 所有内容必须按顺序接在末尾，而不是插到文档开头
    assert app.doc.content == "第一段正文第二段续写内容\r", \
        f"光标漂移防护失败: {app.doc.content!r}"
    assert model.blocks[0].text == "第一段正文第二段续写内容"
    _fake.assert_aligned(model, app)
    print("ok: 草稿态光标漂移防护（续写接在已写内容后）")


def test_block_path_hijack_protection():
    app, writer, model = _fake.make()
    writer.reset_anchor()
    _fake.feed_in_pieces(writer, "# 大标题\n\n正文一段")
    _hijack(app, 0)  # 结构行到达前用户挪走光标
    _fake.feed_in_pieces(writer, "\n\n- 列表甲\n- 列表乙\n\n尾段")
    writer.flush()
    assert app.doc.content == "大标题\r正文一段\r列表甲\r列表乙\r尾段\r", \
        f"结构块路径漂移防护失败: {app.doc.content!r}"
    _fake.assert_aligned(model, app)
    print("ok: 结构块路径光标漂移防护")


def test_hijack_inside_one_feed():
    """一次 feed 内部的批次之间漂移也要防住。"""
    app, writer, model = _fake.make()
    writer.reset_anchor()
    writer.set_speed("fast")  # 批次 40 字符，一次 feed 多批
    long_text = "甲" * 50 + "乙" * 50
    writer.feed(long_text)
    # 在第一批评审之后、第二批之前把光标挪走
    _hijack(app, 0)
    writer.feed("丙" * 50)
    writer.flush()
    assert app.doc.content == long_text + "丙" * 50 + "\r", \
        f"批间漂移防护失败: {app.doc.content!r}"
    print("ok: 批次之间的漂移防护")


def test_write_block_resets_to_caller_position():
    """write_block 是原语入口：锚点重置，从调用方定位的选区开始写。"""
    app, writer, model = _fake.make()
    _fake.feed_in_pieces(writer, "已有内容")
    writer.flush()
    # 模拟 doc_model 原语把选区定位到文档开头后写入
    app.sel.pos = 0
    writer.write_block("插入到开头", animate=False)
    assert app.doc.content.startswith("插入到开头"), app.doc.content
    assert "已有内容" in app.doc.content
    print("ok: write_block 入口锚点重置（从调用方选区写入）")


def test_reset_anchor_new_session():
    """两波写入之间用户挪过光标：第二波从当前光标开始（而非旧锚点）。"""
    app, writer, model = _fake.make()
    writer.reset_anchor()
    _fake.feed_in_pieces(writer, "第一波")
    writer.flush()
    # 会话之间用户把光标挪到文档开头，并期望第二波从这里开始
    _hijack(app, 0)
    writer.reset_anchor()
    _fake.feed_in_pieces(writer, "第二波")
    writer.flush()
    # 段落分隔符会插在用户光标处（_first_block 已为 False 的既有行为），
    # 但关键性质是「第二波」写在了用户的光标位置（文档前部）而非旧锚点（末尾）
    assert app.doc.content.index("第二波") < app.doc.content.index("第一波"), \
        f"会话边界锚点重置失败: {app.doc.content!r}"
    assert app.doc.content == "\r第二波第一波\r", \
        f"第二波写入位置异常: {app.doc.content!r}"
    print("ok: 会话边界锚点重置（第二波从用户当前光标开始）")


if __name__ == "__main__":
    test_draft_hijack_protection()
    test_block_path_hijack_protection()
    test_hijack_inside_one_feed()
    test_write_block_resets_to_caller_position()
    test_reset_anchor_new_session()
    print("ALL ANCHOR TESTS OK")
