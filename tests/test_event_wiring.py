# -*- coding: utf-8 -*-
"""离线回归：提交链路的调试事件接线契约。

deep_test 的 compact_send 场景曾把 `write_block` 写进 MUST_EVENTS，但
紧凑输入框回车提交走的是流式链路：

    engine.write → writer.feed(piece) / writer.flush()
                 → _write_top_level → _write_md(keep_anchor=True)
                 → _BlockDraft.finish → DocModel.register（发 block_register）

而 `write_block` 事件只由外部入口 streaming_writer.write_block() 发出
（编辑原语 replace_block/insert_after/delete_block、存档回写恢复、
回滚恢复重放）。两条链路井水不犯河水， MUST 表却按外部入口的事件名
去断言流式链路，于是三个槽全部误报 [MUST 缺失] write_block(0/1)
（tests/deep_runs/20261005-211342 的 triage）。

修复选择「改 MUST、不碰接线」（避免 write_block 事件名二义：既是外部
入口又是落块，断言与 triage 都无法区分来源）。本文件把接线契约钉死：

1. 流式提交（reset_anchor + feed + flush）必须发 feed/flush/block_register，
   绝不发 write_block；
2. 外部入口 write_block() 必须发 write_block 事件（落块同样登记，
   也带 block_register）；
3. deep_test.py 里 compact_send 的 MUST 表必须与真实提交链路一致
   （要 feed/flush/block_register/write_done，不要 write_block）。

今后若有人改接线（让流式也打 write_block），第 1/2 条先红，提醒同步改
MUST 表与本测试；若有人只改 MUST 表，第 3 条把契约按住，防止回退到
曾经的错配。
"""
import ast
import contextlib
import importlib.util
import os
import re
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import debug  # noqa: E402

# 复用 fake Word（动态 Range / 段落模型等真实 COM 语义的离线实现）
_spec = importlib.util.spec_from_file_location(
    "fake_word_mod", os.path.join(os.path.dirname(__file__), "test_fake_word.py"))
_fake = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_fake)

# 与 deep_test._LINE_RE 同一格式：时间戳 [线程] 级别 事件 字段...
_LINE_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} \[[^\]]*\] "
    r"(INFO|WARN|ERROR) (\S+)(?:\s+(.*))?$")

_STREAM_MD = "# 标题一\n\n正文段落，含**粗体**。\n\n- 项目一\n- 项目二"


def _event_names(path):
    """从调试日志里提取全部事件名（保持顺序，含重复）。"""
    if not path or not os.path.exists(path):
        return []
    with open(path, encoding="utf-8", errors="replace") as f:
        return [m.group(2) for ln in f.read().splitlines()
                if (m := _LINE_RE.match(ln))]


@contextlib.contextmanager
def _debug_capture():
    """把本进程的调试日志重定向到临时文件，收集期间的全部事件名。

    debug 的状态是进程级的（_state / excepthook / AI4WORD_DEBUG_LOG），
    退出时全部还原，避免污染同进程的其它测试（run_offline 单进程串跑）。
    """
    saved_state = dict(debug._state)
    saved_env = os.environ.get("AI4WORD_DEBUG_LOG")
    saved_hooks = (sys.excepthook, threading.excepthook)
    base = os.environ.get("TEMP") or os.path.expanduser("~") or "."
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:
        pass
    path = os.path.join(base, "ai4word_event_wiring.log")
    for stale in (path, path + ".bak"):
        try:
            os.remove(stale)
        except OSError:
            pass
    os.environ["AI4WORD_DEBUG_LOG"] = path
    debug._state.clear()
    debug._state.update({"debug": None, "path": None, "hooks": False,
                         "failed": False})
    try:
        debug.init(["AI4Word", "-debug"])
        events = []
        yield events
        events.extend(_event_names(path))
    finally:
        debug._state.clear()
        debug._state.update(saved_state)
        if saved_env is None:
            os.environ.pop("AI4WORD_DEBUG_LOG", None)
        else:
            os.environ["AI4WORD_DEBUG_LOG"] = saved_env
        sys.excepthook, threading.excepthook = saved_hooks
        try:
            os.remove(path)
        except OSError:
            pass


def _deep_test_compact_send_must():
    """结构化解析 tests/deep_test.py 的 scenarios 清单，取 compact_send 的 MUST。"""
    src = os.path.join(os.path.dirname(__file__), "deep_test.py")
    with open(src, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    for node in ast.walk(tree):
        if (isinstance(node, ast.Assign) and len(node.targets) == 1
                and getattr(node.targets[0], "id", None) == "scenarios"
                and isinstance(node.value, (ast.List, ast.Tuple))):
            for elt in node.value.elts:
                if not (isinstance(elt, ast.Tuple) and len(elt.elts) == 3):
                    continue
                name, _fn, must_node = elt.elts
                if (isinstance(name, ast.Constant) and name.value == "compact_send"
                        and isinstance(must_node, ast.Dict)):
                    must = {}
                    for key, val in zip(must_node.keys, must_node.values):
                        if (isinstance(key, ast.Constant)
                                and isinstance(val, ast.Constant)):
                            must[key.value] = val.value
                    return must
    raise AssertionError("未在 tests/deep_test.py 中找到 compact_send 的 MUST 表")


def test_streaming_submission_emits_block_register_not_write_block():
    """流式提交链路（compact_send 真实走的路径）的事件接线。

    feed/flush 落块由 DocModel.register 登记（block_register），从头到尾
    不经过 write_block 外部入口，所以日志里不得出现 write_block。
    """
    app, writer, model = _fake.make()
    with _debug_capture() as events:
        writer.reset_anchor()           # 与 engine.write 一致：新一波先重置锚点
        _fake.feed_in_pieces(writer, _STREAM_MD)
        writer.flush()
    for required in ("feed", "flush", "block_register"):
        assert required in events, f"流式提交缺少必发事件 {required}：{events}"
    assert "write_block" not in events, (
        "流式提交链路发出了 write_block 事件：它与外部入口（编辑原语/存档回写/"
        "回滚恢复重放）的事件同名，deep_test 的 MUST 表无法区分来源"
    )
    # 同一接线下落块本身仍然完整（接线契约不是"不落块"）
    kinds = [b.kind for b in model.blocks]
    assert kinds == ["heading1", "paragraph", "list"], kinds
    _fake.assert_aligned(model, app)
    print("ok: 流式提交发 feed/flush/block_register，不发 write_block")


def test_external_write_block_entry_emits_write_block():
    """外部入口 write_block()（编辑原语 / 存档回写 / 回滚恢复重放）的事件。"""
    app, writer, model = _fake.make()
    with _debug_capture() as events:
        writer.write_block("外部入口写入的段落", animate=False)
    assert "write_block" in events, (
        "write_block 外部入口未发出 write_block 事件：" + str(events)
    )
    assert "block_register" in events, (
        "write_block 外部入口落块未登记（block_register）：" + str(events)
    )
    assert app.doc.content == "外部入口写入的段落\r", app.doc.content
    _fake.assert_aligned(model, app)
    print("ok: write_block 外部入口发 write_block 事件（落块同样登记）")


def test_compact_send_must_matches_streaming_submission():
    """deep_test 的 compact_send MUST 只能要求流式提交真实发出的事件。"""
    must = _deep_test_compact_send_must()
    # 流式链路（feed/flush → block_register）+ engine 收尾（write_done）真实发出
    for ev in ("feed", "flush", "block_register", "write_done"):
        assert ev in must, (
            f"compact_send 的 MUST 缺少流式提交必发事件 {ev}：{must}"
        )
    # write_block 只有外部入口发，流式提交永不发出，写进 MUST 必然误报缺失
    assert "write_block" not in must, (
        "compact_send 的 MUST 又要求 write_block：该事件只由 write_block 外部入口"
        "（编辑原语/存档回写/回滚恢复重放）发出，流式提交（feed/flush → "
        "_write_md(keep_anchor=True)）不发，必然报 [MUST 缺失]（见 "
        "tests/deep_runs/20261005-211342 的 triage）"
    )


if __name__ == "__main__":
    test_streaming_submission_emits_block_register_not_write_block()
    test_external_write_block_entry_emits_write_block()
    test_compact_send_must_matches_streaming_submission()
    print("ALL EVENT WIRING TESTS OK")
