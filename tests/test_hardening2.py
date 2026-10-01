# -*- coding: utf-8 -*-
"""防呆加固回归测试（doc_model / settings 部分）。"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import test_fake_word as _fake

# 引擎测试在独立 QApplication 里跑（与 test_engine_offline 同一模式）
from app.engine import AgentWorker as _AgentWorker


def _aligned():
    return _fake.assert_aligned


def test_guard_against_bad_indices():
    app, writer, model = _fake.make()
    writer.write_block("甲\n\n乙\n\n丙", animate=False)
    content_before = app.doc.content

    for bad in (-1, -2, 99, 3):
        try:
            model.delete_block(bad)
        except ValueError:
            pass
        else:
            raise AssertionError(f"delete_block({bad}) 应被拒绝")
    for bad in (-1, 99):
        for fn, args in (("replace_block", (bad, "新内容")),
                         ("insert_after", (bad, "新内容")),
                         ("select_block", (bad,))):
            try:
                getattr(model, fn)(*args)
            except ValueError:
                pass
            else:
                raise AssertionError(f"{fn}{args!r} 应被拒绝")
    # 空 md：留下孤儿空段落 + 块索引错位是最坏结果，明确报错才对
    for name, args in (("replace_block", (0, "")), ("replace_block", (0, "   ")),
                       ("insert_after", (0, ""))):
        try:
            getattr(model, name)(*args)
        except ValueError:
            pass
        else:
            raise AssertionError(f"{name}{args!r} 应被拒绝（空 md）")

    # 全部被拒：文档一个字没动
    assert app.doc.content == content_before, app.doc.content
    _fake.assert_aligned(model, app)
    print("ok: 负/越界索引与空 md 被拒（文档未被改动）")


def test_stale_ranges_deactivated():
    app, writer, model = _fake.make()
    writer.write_block("甲\n\n乙\n\n丙\n\n丁", animate=False)
    # 模拟用户从尾部删掉两段：文档段落数 2 < 块模型 4 块
    # （fake 的 TypeText 会自动补齐文档末尾的段落标记）
    app.doc.Content.Delete()
    app.sel.TypeText("甲\r乙")
    model.rebuild_ranges()
    assert model.alignment_warning is not None, "应有文档变短的警告"
    for i, b in enumerate(model.blocks):
        if i >= 2:
            assert b.range is None, f"失配块 {i} 的 Range 应被置 None"
    try:
        model.replace_block(3, "新内容")
    except ValueError:
        pass
    else:
        raise AssertionError("失活块上的 replace_block 应被拒绝")
    # 注意：本场景刻意制造「文档比块模型短」，sum(n_paras) != total 是预期，
    # assert_aligned 的对齐断言不适用；alignment_warning 携带漂移信息即
    # block_map 可见的正确行为
    print("ok: 文档变短时失配块 Range 失活（拒绝错位写入）")


def test_txn_failure_does_not_stick():
    app, writer, model = _fake.make()
    writer.write_block("甲\n\n乙", animate=False)
    model.begin_txn()
    try:
        model.replace_block(99, "不可能")  # 越界 -> 抛错，模拟 run_code 中止
    except ValueError:
        pass
    assert model.is_txn_active()
    model.rollback_txn()
    assert not model.is_txn_active()
    model.begin_txn()
    model.replace_block(0, "丙")
    model.commit_txn()
    assert not model.is_txn_active()
    assert app.doc.content == "丙\r乙\r", app.doc.content
    _fake.assert_aligned(model, app)
    print("ok: 卡住的事务被回滚后正常工作")


def test_replace_text_updates_md_and_reports():
    app, writer, model = _fake.make()
    writer.write_block("# 标题甲\n\n正文乙", animate=False)
    hit = model.replace_text("甲", "丙")
    assert hit is True, "replace_text 应回报是否真的替换到"
    miss = model.replace_text("不存在三个字", "X")
    assert miss is False, "未命中时应返回 False"
    assert [b.md for b in model.blocks][0] == "# 标题丙", \
        [b.md for b in model.blocks]
    _fake.assert_aligned(model, app)
    print("ok: replace_text 回写 md 并真实返回命中/未命中")


def test_settings_type_tolerance():
    """settings.json 语法合法但值类型错误时不能启动崩溃。"""
    import tempfile
    from app.settings import Settings
    path = tempfile.mktemp(suffix=".json")
    with open(path, "w", encoding="utf-8") as f:
        f.write('{"geometry": ["100", null], "stay_on_top": "false", '
                '"api_key": 123, "speed": "turbo", "expanded": "yes"}')
    try:
        s = Settings(path).load()
        assert s.get("geometry") is None, "非法 geometry 应回退默认"
        assert s.get("stay_on_top") is True, "字符串布尔不应被采纳"
        assert s.get("api_key") == "", "非法 api_key 应回退"
        assert s.get("speed") == "auto"
        assert s.get("expanded") is False
    finally:
        os.remove(path)
    print("ok: settings.json 类型级损坏被容忍（回退默认，不崩溃）")


def test_settings_save_returns_bool():
    """save() 返回成功状态（写盘失败时调用方需要能提示）。"""
    import tempfile
    from app.settings import Settings
    path = tempfile.mktemp(suffix=".json")
    try:
        s = Settings(path).load()
        assert s.save() is True
        s.path = os.path.join(path, "no", "dir", "settings.json")
        assert s.save() is False
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    print("ok: settings.save() 正确返回成功/失败")


def test_engine_speed_lives():
    """档位设置必须立即对进行中的生成生效（不再排队等生成结束）。"""
    import importlib.util
    import threading
    import ai_client
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QApplication
    if QApplication.instance() is None:
        QApplication([])

    spec = importlib.util.spec_from_file_location(
        "fake_word_speed_mod",
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_fake_word.py"))
    fake_mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(fake_mod)

    class _Shell:
        def __init__(self, a):
            self._a = a

        @property
        def Documents(self):
            return self

        @property
        def Count(self):
            return 1

        def Add(self):
            return self._a.doc

        @property
        def ActiveDocument(self):
            return self._a.doc

        @property
        def Selection(self):
            return self._a.Selection

        def __getattr__(self, name):
            return getattr(self._a, name)

    gate = threading.Event()

    def fake_stream(prompt, key, system):
        yield "一段内容"
        gate.wait(timeout=8)
        yield "尾段"

    shell_app, _w, _m = fake_mod.make()
    worker = _AgentWorker(word_factory=lambda: _Shell(shell_app))
    worker.set_api_key("test-key")
    rec_states = []
    worker.stateChanged.connect(lambda s: rec_states.append(s))
    worker.start()
    QTest.qWait(120)
    ai_client.ai_stream = fake_stream
    worker.send("write", "测试")
    # 等到「writing」状态出现（命令确实开始了）
    t0 = time.time()
    while "writing" not in rec_states:
        QTest.qWait(20)
        if (time.time() - t0) * 1000 > 8000:
            raise AssertionError(f"write 命令未开始，状态: {rec_states}")
    worker.set_speed("fast")  # 不走队列，应立即生效
    QTest.qWait(50)
    assert worker._writer._speed == "fast", \
        f"档位应立即生效（实际 {worker._writer._speed}）"
    gate.set()
    t0 = time.time()
    while worker._busy:
        QTest.qWait(20)
        if (time.time() - t0) * 1000 > 8000:
            break
    worker.send("quit")
    worker.wait(4000)
    print("ok: 打字档位立即对进行中的生成生效")


def test_engine_attach_only_for_map():
    """查看类命令（块地图）不应启动新 Word 进程。"""
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QApplication
    if QApplication.instance() is None:
        QApplication([])

    launched = []
    never = []

    def fake_factory():  # 模拟「启动 Word」（Dispatch）
        launched.append(1)
        return None

    def fake_attach():  # 模拟 GetObject：没在运行
        never.append(1)
        return None

    worker = _AgentWorker(word_factory=fake_factory, attach_factory=fake_attach)
    worker.start()
    QTest.qWait(120)
    worker.send("refresh_map")
    # 等待效果而非固定时长：offscreen 下线程启动耗时偶发超过 150ms
    t0 = time.time()
    while not never and (time.time() - t0) * 1000 < 8000:
        QTest.qWait(20)
    assert launched == [], f"查看块地图不应启动 Word（实际调用 {len(launched)} 次）"
    assert len(never) >= 1, "应调用附加构造器（GetObject），一次都没调"
    worker.send("quit")
    worker.wait(4000)
    print("ok: 块地图查看类命令不启动新 Word 进程（只附加）")


if __name__ == "__main__":
    test_guard_against_bad_indices()
    test_stale_ranges_deactivated()
    test_txn_failure_does_not_stick()
    test_replace_text_updates_md_and_reports()
    test_settings_type_tolerance()
    test_settings_save_returns_bool()
    test_engine_speed_lives()
    test_engine_attach_only_for_map()
    print("\ndoc_model / settings / engine 加固测试全部通过")
