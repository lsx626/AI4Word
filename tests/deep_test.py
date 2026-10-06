# -*- coding: utf-8 -*-
"""全真实驱动深度测试（deep_test）—— AI4Word 工作流「测试 → 分析日志 → 修复 → 复测」的测试层。

与 tests/debug_walkthrough.py（offscreen + 槽函数驱动）不同，本层追求**全真实**：

- **真实平台窗口**：QT_QPA_PLATFORM=windows，悬浮窗真正画出来（真实 HWND、
  DWM 圆角、真实贴边吸附、真实托盘），不走 offscreen。
- **真实输入**：发送走真实回车键（QTest.keyClick），按钮走真实鼠标点击
  （QTest.mouseClick），中断后的 回滚 / 保留 / 追加补充 点击选择条上**真实的
  按钮**，拖动 / Esc / 热键走真实事件序列。
- **真实 Word COM 全程参与**：附加或启动真实 Word，建临时空文档写入。
- **真实 Atria API**：密钥来自 .env（唯一的假是两处「注入故障」：排版模拟
  1/0 语句失败、流式抛一次错——它们验证的是真实的自我修复机制，故障本身
  是脚本点起来的，triage 会把这两类 WARN/ERROR 标成「预期故障路径」）。

每轮的产物归档到 tests/deep_runs/<时间戳>/：
- debug.log     本轮完整调试日志（-debug 模式，%APPDATA%\\AI4Word\\debug.log 的拷贝）
- summary.json  机器可读结果（场景 / FAIL / MUST_EVENTS 缺失 / 未分类告警）
- triage.txt    人类可读分类报告（真实 bug 候选 / 平台假象 / 预期故障路径）

退出码（供工作流脚本判断循环）：
- 0 = 全绿：无 FAIL、无 MUST_EVENTS 缺失、无未分类 WARN/ERROR
- 1 = 有发现（FAIL / 缺失 / 待确认告警）
- 2 = 环境错误（缺 ATRIA_API_KEY、已有实例运行、harness 异常等）

运行：
    .venv\\python.exe tests\\deep_test.py                 # 全量（单进程）
    .venv\\python.exe tests\\deep_test.py --only interrupt  # 只跑名字含 interrupt 的场景
    .venv\\python.exe tests\\deep_test.py --seed 31337      # 压力场景的确定性种子
    .venv\\python.exe tests\\deep_test.py --slot 1/3        # 并行模式：只跑分到槽 1 的场景

并行模式由 tests\\deep_parallel.py 编排（默认 3 槽，各槽独立 Word / 日志 /
单实例锁 / 临时 settings；全局热键与自启注册表场景固定属于槽 0）：
    .venv\\python.exe tests\\deep_parallel.py --slots 3

注意：跑测试期间不要用鼠标键盘（窗口会被真实移动、真实点击）；Word 会被
附加或启动一个新空文档，临时文档全程不保存。请先退出正在运行的 AI4Word
（单实例锁会在装配时拦截）。
"""
import argparse
import datetime
import json
import os
import random
import shutil
import sys
import time
import traceback

# ---- 真实平台：必须在 import debug_walkthrough 之前锁定 ----
# debug_walkthrough 在 import 时会 setdefault("QT_QPA_PLATFORM", "offscreen")，
# 先显式置为 windows，setdefault 才不会覆盖。
os.environ["QT_QPA_PLATFORM"] = "windows"

if "-debug" not in sys.argv and "--debug" not in sys.argv:
    sys.argv.append("-debug")

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
_RUNS = os.path.join(_TESTS, "deep_runs")

sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)

import debug_walkthrough as W   # 复用装配 / 等待原语 / Word 工厂 / 已验证场景
from debug_walkthrough import (check, note_platform, pump, wait_for, wait_busy,
                               wait_handled, wait_awaiting_choice, doc_text)

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QPushButton

from app import __version__
from app import debug

# ---------------------------------------------------------------- 全局状态

STATE = {"seed": 1007, "only": None, "skip": set(), "keep_word": False,
         "run_dir": None, "results": [], "t0": None,
         "slot": (0, 1), "allowed": None}

DOC_VARS = {"req": 0, "rsp": None, "text": None}   # worker 线程读到的文档变量

# ---------------------------------------------------------------- 日志解析

# 窗口解析抽到 tests/deep_log.py（纯 stdlib + app.debug）：本文件是真实驱动
# 入口，导入即锁定 QT_QPA_PLATFORM=windows 并追加 -debug，不能被离线测试套件
# import；解析逻辑放 deep_log 里才能被 tests/test_debug_window.py 直接回归。
from deep_log import _count, _field_int, _ms_floor, _parse_window


# ---------------------------------------------------------------- 场景框架

# 「预期故障路径」：场景注入的故障打出的 WARN/ERROR（不是真实 bug）。
SCENARIO_EXPECTED = {
    "arrange_repair": {"stmt_failed"},
    "arrange_gen_fail": {"gen_code_failed", "arrange_failed"},
    "stream_fault": {"stream_error"},
    "settings_dialog": {"settings_save_failed"},
    "tray_actions": {"autostart_enable_failed", "autostart_disable_failed"},
}

# 真实平台下仍可能出现、但与代码无关的环境噪音（按事件名归类）。
ENV_NOISE = {"api_request_retry", "api_stream_retry", "api_stream_failed",
             "api_request_failed", "word_connect_failed"}

# 机器级全局状态：只能固定在 slot 0 跑（并行槽互不冲突的保证）：
#   drag_snap_hotkey —— RegisterHotKey(Ctrl+Alt+Space) 是全局热键；
#   tray_actions     —— 开机自启写注册表（HKEY_CURRENT_USER\...\Run）。
SLOT0_ONLY = frozenset({"drag_snap_hotkey", "tray_actions"})

# 槽种子：每个槽都先跑一遍的场景，保证槽内依赖闭合——
#   compact_send       —— 块地图 / 后续场景需要的块模型与首次 Word 连接；
#   review_presets_save —— doc_persistence 需要先有 save_blocks 落进文档变量。
# 种子场景本身很轻（一条短 AI 写入 + 本地预设/修订/存档，无额外 API），
# 且各槽并行执行，墙钟成本 ≈ 一份。代价：合并报告里这两个场景名会出现
# K 次（每槽一份），属预期。
SLOT_SEED = ("compact_send", "review_presets_save")


def slot_assignment(ordered_names, slot, total):
    """在过滤后的场景名清单（保持顺序）上做槽分区。

    - total <= 1：全部属于本进程（单进程模式，兼容旧行为）。
    - 否则：SLOT_SEED 场景每槽都跑；其余非 SLOT0_ONLY 场景按出现顺序
      轮询摊到 0..total-1 各槽；SLOT0_ONLY 场景无条件归 slot 0。
      由此 block_map / doc_persistence 等依赖前置写入与存档的场景，
      在本槽内一定能等到它的依赖（种子）先跑。
    deep_parallel.py 复用同一函数，保证编排器与测试进程的分区完全一致。
    """
    if total <= 1:
        return set(ordered_names)
    seeds = {n for n in SLOT_SEED if n in ordered_names}
    rest = [n for n in ordered_names
            if n not in SLOT0_ONLY and n not in seeds]
    mine = set(seeds)
    mine.update(n for i, n in enumerate(rest) if (i % total) == slot)
    if slot == 0:
        mine.update(n for n in ordered_names if n in SLOT0_ONLY)
    return mine


# 场景执行顺序的权威清单（deep_parallel.py 复用来做槽分区；
# 与 main() 里的 scenarios 列表必须一致，那里有 assert 校验）。
SCENARIO_NAMES = (
    "speed_gears", "compact_send", "block_map", "interrupt_rollback",
    "interrupt_keep", "interrupt_extra", "panel_send", "keys",
    "arrange_ok", "arrange_repair", "arrange_gen_fail", "stream_fault",
    "review_presets_save", "expand_collapse", "settings_dialog",
    "tray_actions", "drag_snap_hotkey", "esc_hide_hint",
    "input_rejections", "doc_persistence", "stress_mixed",
)


def run_scenario(name, fn, must):
    """跑一个场景并记录：FAIL / 平台假象 / MUST_EVENTS 缺失 / 待确认告警。"""
    if STATE["only"] and STATE["only"] not in name:
        return
    if any(s in name for s in STATE["skip"]):
        print("--- 跳过场景 %s（--skip）" % name)
        return
    if name not in STATE["allowed"]:
        print("--- 跳过场景 %s（分到其它槽）" % name)
        return
    print("\n===== 场景：%s =====" % name)
    start = time.time()
    fail0, plat0 = len(W.FAIL), len(W.PLATFORM)
    errors = []
    try:
        fn()
    except Exception:
        errors.append("harness 异常:\n" + traceback.format_exc())
        print("  HARNESS-ERROR " + name)
        traceback.print_exc()
    end = time.time() + 2.0          # 收尾尾巴：worker 线程最后一跳日志的余量
    events = _parse_window(start, end)
    missing = ["%s(%d/%d)" % (k, _count(events, k), n)
               for k, n in must.items() if _count(events, k) < n]
    fails = list(W.FAIL[fail0:])
    notes = list(W.PLATFORM[plat0:])
    expected = SCENARIO_EXPECTED.get(name, set())
    warns, env = [], []
    for (lvl, ev, f) in events:
        if lvl == "INFO":
            continue
        if ev in expected:
            continue
        if ev in ENV_NOISE:
            env.append("%s %s %s" % (lvl, ev, f[:120]))
        else:
            warns.append("%s %s %s" % (lvl, ev, f[:160]))
    res = {"name": name, "ok": not (fails or missing or warns or errors),
           "fails": fails, "notes": notes, "missing_must": missing,
           "unclassified": warns, "environment": env, "errors": errors,
           "duration_s": round(end - 2.0 - start, 1)}
    STATE["results"].append(res)
    print("--- %s: %s（FAIL %d / 缺失 %d / 待确认 %d / 环境 %d）" % (
        name, "通过" if res["ok"] else "有发现", len(fails), len(missing),
        len(warns), len(env)))


# ---------------------------------------------------------------- 真实输入

def _click(button):
    """真实鼠标点击（完整 press/release/clicked 信号序列）。"""
    QTest.mouseClick(button, Qt.LeftButton)


def _choice_button(window, prefix):
    """中断选择条上按文字找真实按钮。"""
    for b in window._choice_bar.findChildren(QPushButton):
        if b.text().strip().startswith(prefix):
            return b
    return None


def _send_compact_enter(window, text):
    """紧凑输入框：填字 + 真实回车（returnPressed -> _on_send_compact）。"""
    window.input_c.setText(text)
    QTest.keyClick(window.input_c, Qt.Key_Return)


def _send_panel_enter(window, text):
    """面板输入框：填字 + 真实回车（InputEdit returnPressed -> _on_send_panel）。"""
    window.input_p.setPlainText(text)
    QTest.keyClick(window.input_p, Qt.Key_Return)


LONG_WRITE_PROMPT = ("写一篇关于秋天的长篇抒情散文，至少 1000 字，"
                     "要有一级标题、多个二级标题、若干段落和一个列表。")
SHORT_WRITE_PROMPT = "写一首关于秋天的短诗：一个一级标题加两段正文。"
ARRANGE_PROMPT = "把所有一级标题居中并加粗"


# ---------------------------------------------------------------- 场景矩阵

def scenario_speed_gears(window, worker, settings):
    # 直接复用：button.click() 本身就是真实点击信号路径
    W.speed_gears(window, worker, settings)


def scenario_compact_send(window, worker):
    print("--- 紧凑态发送：真实回车 + 真实流式写作 ---")
    window.set_expanded(False)
    pump(0.3)
    check(not worker._busy and worker._queue.empty(), "发送前 worker 空闲")
    _send_compact_enter(window, SHORT_WRITE_PROMPT)
    check(wait_busy(worker, 60), "worker 进入 writing（含首次连接 Word）")
    check(wait_handled(worker, 300), "流式写作完成，worker 回 idle")
    check("已连接" in window.word_label.text(),
          "Word 连接状态: " + window.word_label.text())
    text = doc_text(worker) or ""
    check(len(text) > 10 and "#" not in text, "正文真实写入（%d 字）" % len(text))


def scenario_block_map(window, worker):
    W.block_map(window, worker)


def scenario_interrupt_real(window, worker, choice, extra_text=None):
    """发送长文 -> 真实回车中断 -> 点击选择条上真实的 回滚/保留/追加 按钮。"""
    print("--- 中断（真实回车 + 真实按钮）-> %s ---" % choice)
    window.set_expanded(False)
    pump(0.3)
    window._speed_group.button(0).click()      # slow：流式足够长，来得及中断
    before = doc_text(worker) or ""
    W.CHARS["n"] = 0
    _send_compact_enter(window, LONG_WRITE_PROMPT)
    check(wait_busy(worker, 60), "worker 进入 writing")
    if not wait_for(lambda: W.CHARS["n"] > 0, 60):
        note_platform("60s 内无流式分片（AI/网络波动）")
    pump(0.5)
    # 流式自己挂掉 / 跑空的兜底：没有选择条就别按出新一轮写作
    if not worker._busy and not window._awaiting_choice:
        note_platform("流式自行结束（网络波动），跳过本轮选择")
        return
    mid = doc_text(worker, 1.0)
    QTest.keyClick(window.input_c, Qt.Key_Return)   # busy 中第二个回车 = 中断
    check(wait_awaiting_choice(window, 150), "中断后弹出 回滚/保留/追加 选择条")
    mid = doc_text(worker, 1.0)
    if choice == "rollback":
        btn = _choice_button(window, "回滚")
        check(btn is not None and btn.isEnabled(), "回滚按钮存在且可用")
        if btn is not None:
            _click(btn)
        check(wait_handled(worker, 180), "回滚分支处理完成")
        after = doc_text(worker)
        if after != before and after is not None:
            debug.log("rollback_diff", before_len=len(before),
                      after_len=len(after or ""),
                      first_diff=W._first_diff(before, after or ""),
                      full=True)
        check(doc_text(worker) == before, "回滚后文档恢复到生成前状态")
    elif choice == "keep":
        btn = _choice_button(window, "保留")
        check(btn is not None and btn.isEnabled(), "保留按钮存在且可用")
        if btn is not None:
            _click(btn)
        check(wait_handled(worker, 60), "保留分支处理完成")
        final = doc_text(worker) or ""
        if mid:
            check(len(final) >= len(mid) and len(mid) > len(before),
                  "保留住了中断时的部分正文")
        else:
            # worker 阻在 _wait_choice 时 mid 读不到：与 before 比
            check(len(final) > len(before), "保留住了部分正文（vs before）")
    elif choice == "extra":
        btn = _choice_button(window, "追加")
        check(btn is not None and btn.isEnabled(), "追加按钮存在且可用")
        if btn is not None:
            _click(btn)
        pump(0.3)
        window.input_p.setPlainText(extra_text or "再补充一段初冬清晨霜景的描写，150 字左右。")
        QTest.keyClick(window.input_p, Qt.Key_Return)   # 追加模式提交
        consumed = wait_for(lambda: not window._awaiting_choice, 8)
        if not consumed:
            check(False, "追加补充已提交（awaiting_choice 未解除）")
            worker.choose("keep")           # 兜底：解除 worker 的 _wait_choice 阻塞
            wait_handled(worker, 60)
            return
        check(consumed, "追加补充已提交，选择条消失")
        check(wait_handled(worker, 300), "追加生成完成，worker 回 idle")
        final = doc_text(worker) or ""
        if mid:
            check(len(final) > len(mid), "追加内容写入文档")
        else:
            check(len(final) > len(before), "追加内容写入文档（vs before）")


def scenario_panel_send(window, worker):
    print("--- 面板态发送：展开 + 真实回车 ---")
    window.btn_expand.click()
    pump(0.3)
    window._speed_group.button(2).click()     # fast
    _send_panel_enter(window, "写一篇关于春天的短文，含一个一级标题和三个段落。")
    check(wait_busy(worker, 60), "worker 进入 writing")
    check(wait_handled(worker, 300), "面板发送流式写作完成")
    check(len(doc_text(worker) or "") > 20, "面板输入写入文档")


def scenario_keys(window, worker, tray):
    # 已是真实事件路径：Enter / Shift+Enter / IME 合成 / closeEvent / 托盘激活 / 忙时拒绝
    W.keys_inputs_and_close(window, worker, tray)


def _arrange_real(window, worker, prompt):
    window._mode_group.button(1).click()      # 切到排版模式（真实点击）
    pump(0.2)
    _send_panel_enter(window, prompt)
    check(wait_handled(worker, 300), "排版链路完成（含可能的自我修复）")
    window._mode_group.button(0).click()
    check(window._write_mode, "切回写作模式")


def scenario_arrange_ok(window, worker):
    print("--- 排版链路：gen_code -> run_code（真实 AI 代码） ---")
    _arrange_real(window, worker, ARRANGE_PROMPT)


def scenario_arrange_repair(window, worker):
    """注入一条必然失败的 1/0 语句：验证「失败 -> fix_code -> 重跑」真实链路。"""
    print("--- 排版链路：失败 -> 自我修复 ---")
    import app.engine as engine_mod
    real_gen = engine_mod.gen_code
    counter = {"n": 0}

    def fake_gen(prompt, api_key, session=None, block_map_fn=None, sink=print,
                 model=None):
        counter["n"] += 1
        if counter["n"] == 1:
            sink("（deep_test 注入的故意失败代码）")
            return "x = 1 / 0"
        return real_gen(prompt, api_key, session=session,
                        block_map_fn=block_map_fn, sink=sink, model=model)

    real_fix = engine_mod.fix_code
    fixes = {"n": 0}

    def fake_fix(prompt, failed_code, error_msg, api_key, session=None,
                 sink=print, block_map_fn=None, model=None):
        fixes["n"] += 1
        return real_fix(prompt, failed_code, error_msg, api_key, session=session,
                        sink=sink, block_map_fn=block_map_fn, model=model)

    engine_mod.gen_code = fake_gen
    engine_mod.fix_code = fake_fix
    try:
        window._mode_group.button(1).click()
        pump(0.2)
        _send_panel_enter(window, ARRANGE_PROMPT)
        check(wait_handled(worker, 300), "失败 -> fix_code -> run_code2 全链路走完")
        check(counter["n"] == 1, "首次 gen_code 即注入的失败")
        check(fixes["n"] >= 1, "fix_code 自我修复被触发")
        window._mode_group.button(0).click()
    finally:
        engine_mod.gen_code = real_gen
        engine_mod.fix_code = real_fix


def scenario_arrange_gen_fail(window, worker):
    """注入 gen_code 抛错：验证失败上报（arrange_failed）真实链路。"""
    print("--- 排版链路：生成失败上报 ---")
    import app.engine as engine_mod
    real_gen = engine_mod.gen_code
    raised = {"n": 0}

    def fake_gen(*a, **kw):
        raised["n"] += 1
        if raised["n"] == 1:
            raise RuntimeError("deep_test 注入的 gen_code 故障")
        return real_gen(*a, **kw)

    engine_mod.gen_code = fake_gen
    try:
        window._mode_group.button(1).click()
        pump(0.2)
        _send_panel_enter(window, ARRANGE_PROMPT)
        check(wait_handled(worker, 300), "生成失败上报完成，worker 回 idle")
        window._mode_group.button(0).click()
    finally:
        engine_mod.gen_code = real_gen


def scenario_stream_fault(window, worker):
    """注入一次流式抛错：验证 stream_error 真实链路与 worker 自恢复。"""
    print("--- 流式故障注入：stream_error -> worker 自恢复 ---")
    import ai_client as ai_mod
    real_stream = ai_mod.ai_stream
    raised = {"n": 0}

    def fake_stream(*a, **kw):
        raised["n"] += 1
        if raised["n"] == 1:
            raise RuntimeError("deep_test 注入的 stream 故障")
        return real_stream(*a, **kw)

    ai_mod.ai_stream = fake_stream
    try:
        window.set_expanded(False)
        pump(0.2)
        _send_compact_enter(window, SHORT_WRITE_PROMPT)
        check(wait_busy(worker, 60), "worker 进入 writing")
        check(wait_handled(worker, 120), "故障路径处理完成，worker 回 idle")
    finally:
        ai_mod.ai_stream = real_stream


def scenario_review_presets_save(window, worker):
    print("--- 修订开关 / 四个预设 / 存档（真实点击） ---")
    window.btn_review.click()                  # 真实点击：修订开
    check(wait_handled(worker, 60), "修订开启命令处理完成")
    window.btn_review.click()                  # 修订关
    check(wait_handled(worker, 60), "修订关闭命令处理完成")
    for name in ("论文", "公文", "简历", "博客"):
        idx = window.preset_combo.findText(name)
        check(idx > 0, "预设「%s」在组合框中" % name)
        # setCurrentIndex 只发 currentIndexChanged；activated.emit 补齐
        # 「只连 activated」的旧绑定路径，同时消除「再次选同一项」时
        # setCurrentIndex 无信号被静默吞掉的隐患。当前绑定只连
        # currentIndexChanged，activated 无接收端且不改变 currentIndex，
        # 不会重复投递（tests/test_ui.py 钉死单投递契约）
        window.preset_combo.setCurrentIndex(idx)
        window.preset_combo.activated.emit(idx)
        check(wait_handled(worker, 60), "预设「%s」应用完成" % name)
        check(window.preset_combo.currentIndex() == 0, "应用后组合框复位")
    window.btn_save.click()                    # 真实点击：存档
    check(wait_handled(worker, 60), "存档命令处理完成")
    check(len(worker._model.blocks) > 0, "块模型非空，存档有内容")


def scenario_expand_collapse_avatar(window):
    print("--- 展开 / 收起 / 头像点击（真实鼠标） ---")
    window.set_expanded(False)
    pump(0.2)
    QTest.mouseClick(window.avatar_c, Qt.LeftButton)   # 紧凑头像 -> 展开
    check(window._expanded, "紧凑态头像点击 -> 展开")
    QTest.mouseClick(window.avatar_p, Qt.LeftButton)   # 面板头像 -> 收起
    pump(0.2)
    check(not window._expanded, "面板头像点击 -> 收起")
    QTest.mouseClick(window.btn_expand, Qt.LeftButton)
    check(window._expanded, "展开按钮")
    QTest.mouseClick(window.btn_collapse, Qt.LeftButton)
    pump(0.2)
    check(not window._expanded, "收起按钮")


def scenario_settings_dialog(window, worker, settings):
    # 已是真实控件操作（QLineEdit 改值 / 真实 accept/reject / 真实密码显隐按钮）
    W.settings_dialog(window, worker, settings)


def scenario_tray_actions(window, tray, worker):
    # 真实托盘菜单动作触发（act_show.trigger / 自启勾选）/ 注册表状态还原
    W.tray_actions(window, tray, worker)


def scenario_drag_snap_hotkey(window):
    # 真实平台：hotkey_register 必须成功（offscreen 的「无 HWND」借口失效）
    W.drag_snap_summon_hotkey(window)


def scenario_esc_hide_hint(window, tray):
    W.esc_hide_and_hint(window, tray)


def scenario_input_rejections(window, worker):
    W.input_rejections(window, worker)


def _patch_doc_vars(worker):
    """包一层 worker._handle：新增合成命令 _read_doc_vars（在 worker 线程、
    COM apartment 所属线程上读文档变量），其余命令原样转交 W 的计数包装。

    _read_doc_vars 是只读探针：完成经 DOC_VARS["rsp"] 标志确认，两侧都
    不进 W 的 SENT/DONE（W.wrapped_send 的 _UNCOUNTED 排除表里有它）。
    只一侧计数会造成 SENT/DONE 永久失步：之后每条 wait_handled 都判
    失败并烧满超时——本场景的 refresh_map 检查曾被这样误报 FAIL。
    """
    prev = worker._handle

    def wrapped(cmd, payload):
        if cmd == "_read_doc_vars":
            value = None
            try:
                value = worker._doc.Variables("AI4WordBlocks").Value
            except Exception:
                value = None
            DOC_VARS["text"] = value
            DOC_VARS["rsp"] = (payload.get("req")
                               if isinstance(payload, dict) else None)
            return
        prev(cmd, payload)

    worker._handle = wrapped


def scenario_doc_persistence(window, worker):
    print("--- 文档持久化：存档落盘 + 抹掉绑定 + 重连恢复 ---")
    t0 = time.time()
    check(worker._model is not None and len(worker._model.blocks) > 0,
          "前面场景已写入块（当前 %d 块）"
          % (len(worker._model.blocks) if worker._model else 0))
    # 1) 存档真实落进文档变量（在 worker 线程读，不违反 COM apartment）
    DOC_VARS["req"] += 1
    DOC_VARS["rsp"] = None
    worker.send("_read_doc_vars", {"req": DOC_VARS["req"]})
    check(wait_for(lambda: DOC_VARS["rsp"] == DOC_VARS["req"], 15),
          "worker 线程读到文档变量")
    stored = DOC_VARS["text"]
    check(isinstance(stored, str) and stored.strip().isdigit() and
          int(stored.strip()) >= 1,
          "AI4WordBlocks 计数变量存在且 ≥1：" + repr(stored)[:60])
    # 2) 抹掉引擎绑定：下一条查看类命令触发「附加运行中的 Word -> 重连装配」
    worker._model = None
    worker._doc = None
    worker._writer = None
    worker.send("refresh_map")
    # 按命令语义等 refresh_map 自己处理完（不依赖全局 SENT/DONE）：
    # 探针命令若造成全局计数失步，已完成的命令会被误判成「没完成」
    # 并烧满整个超时——refresh_map 的 FAIL 就是这么来的
    check(wait_handled(worker, 60, cmd="refresh_map"),
          "重连装配完成（refresh_map）")
    # 3) 恢复事件与恢复块数（只看本场景窗口，排除首次装配的 restored=0）
    evs = _parse_window(t0, time.time() + 2)
    restored = _field_int(evs, "word_assembled", "restored")
    check(restored is not None and restored >= 1,
          "重连从文档变量恢复出块（restored=%s）" % restored)
    check(worker._model is not None and len(worker._model.blocks) >= 1,
          "块模型恢复非空（%d 块）"
          % (len(worker._model.blocks) if worker._model else 0))
    # 4) 块地图在重连后仍然能刷出行
    window.btn_blockmap.click()
    pump(0.3)
    check(window.btn_blockmap.isChecked() and window.block_panel.isVisible(),
          "重连后块地图侧栏可用")
    ok = wait_for(lambda: window.block_panel._list.count() > 0, 30)
    check(ok, "块地图刷新出 %d 行" % window.block_panel._list.count())
    window.btn_blockmap.click()


def scenario_stress_mixed(window, worker, settings):
    print("--- 综合压力：随机组合（种子 %d） ---" % STATE["seed"])
    rng = random.Random(STATE["seed"])
    ops = rng.sample(["preset", "save", "review", "speed", "arrange",
                      "interrupt_rollback"], 3)
    print("    随机选出: %s" % ", ".join(ops))
    for op in ops:
        if op == "preset":
            name = rng.choice(("论文", "公文", "简历", "博客"))
            idx = window.preset_combo.findText(name)
            window.preset_combo.setCurrentIndex(idx)
            window.preset_combo.activated.emit(idx)  # 兼容 activated-only 绑定
            check(wait_handled(worker, 60), "压测：预设「%s」" % name)
        elif op == "save":
            window.btn_save.click()
            check(wait_handled(worker, 60), "压测：存档")
        elif op == "review":
            window.btn_review.click()
            check(wait_handled(worker, 60), "压测：修订切换")
        elif op == "speed":
            idx = rng.randrange(3)
            window._speed_group.button(idx).click()
            pump(0.3)
            check(True, "压测：速度档 %d" % idx)
        elif op == "arrange":
            _arrange_real(window, worker, ARRANGE_PROMPT)
        elif op == "interrupt_rollback":
            scenario_interrupt_real(window, worker, "rollback")


# ---------------------------------------------------------------- 归档与总结

def archive_and_summarize():
    """把 debug.log / summary.json / triage.txt 写进本轮目录。"""
    run_dir = STATE["run_dir"]
    ran = bool(STATE["results"])
    os.makedirs(run_dir, exist_ok=True)
    log_src = debug.log_path()
    if log_src and os.path.exists(log_src):
        try:
            shutil.copy2(log_src, os.path.join(run_dir, "debug.log"))
        except OSError:
            pass

    totals = {"fail": sum(len(r["fails"]) for r in STATE["results"]),
              "notes": sum(len(r["notes"]) for r in STATE["results"]),
              "missing_must": sum(len(r["missing_must"])
                                  for r in STATE["results"]),
              "unclassified": sum(len(r["unclassified"]) for r in STATE["results"]),
              "environment": sum(len(r["environment"]) for r in STATE["results"]),
              "harness_errors": sum(len(r["errors"]) for r in STATE["results"])}
    clean = (ran and not totals["fail"] and not totals["missing_must"]
             and not totals["unclassified"] and not totals["harness_errors"])
    summary = {
        "program": "AI4Word", "version": __version__,
        "python": sys.version.split()[0],
        "platform": "windows (real)", "layer": "deep_test (fully real driving)",
        "started": datetime.datetime.fromtimestamp(
            STATE["t0"]).strftime("%Y-%m-%d %H:%M:%S"),
        "seed": STATE["seed"],
        "ran_scenarios": ran,
        "scenarios": STATE["results"],
        "totals": totals,
        "clean": clean,
        "exit_code": 0 if clean else (2 if not ran else 1),
    }
    with open(os.path.join(run_dir, "summary.json"), "w",
              encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    triage = []
    triage.append("AI4Word deep_test 分类报告（triage）")
    triage.append("目录: %s    种子: %d" % (run_dir, STATE["seed"]))
    triage.append("结果: %s    场景: %d    FAIL %d    MUST 缺失 %d    "
                  "待确认告警 %d    环境告警 %d"
                  % ("全绿" if clean else "有发现", len(STATE["results"]),
                     totals["fail"], totals["missing_must"],
                     totals["unclassified"], totals["environment"]))
    triage.append("")
    triage.append("分类规则：")
    triage.append("  - 待确认（可能是真实 bug）：场景时间窗内的 WARN/ERROR，且不属于")
    triage.append("    注入故障（stmt_failed / gen_code_failed / stream_error /")
    triage.append("    settings_save_failed / autostart_*_failed）或环境重试噪音")
    triage.append("  - 环境告警：api_*_retry / api_*_failed / word_connect_failed")
    triage.append("  - 预期故障路径：场景注入的故障（已跳过，参见 SCENARIO_EXPECTED）")
    triage.append("  - uncaught_exception / uncaught_thread_exception：一律视为真实 bug")
    triage.append("")
    for r in STATE["results"]:
        triage.append("### %s   %s   (%.1fs)"
                      % (r["name"], "通过" if r["ok"] else "有发现",
                         r["duration_s"]))
        for x in r["fails"]:
            triage.append("  [FAIL]        " + x)
        for x in r["missing_must"]:
            triage.append("  [MUST 缺失]   " + x)
        for x in r["unclassified"]:
            triage.append("  [待确认]      " + x)
        for x in r["environment"]:
            triage.append("  [环境]        " + x)
        for x in r["notes"]:
            triage.append("  [平台假象]    " + x)
        for x in r["errors"]:
            triage.append("  [HARNESS 异常] " + x.splitlines()[0])
    path = os.path.join(run_dir, "triage.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(triage) + "\n")
    print("\n归档目录: " + run_dir)
    print("triage:  " + path)
    return summary


def main():
    ap = argparse.ArgumentParser(description="AI4Word 全真实驱动深度测试")
    ap.add_argument("--seed", type=int, default=1007, help="压力场景随机种子")
    ap.add_argument("--only", default=None, help="只跑名字含该串的场景")
    ap.add_argument("--skip", default="", help="跳过名字含这些串的场景（逗号分隔）")
    ap.add_argument("--keep-word", action="store_true",
                    help="结束时不清掉 Word 临时文档（便于人工查看）")
    ap.add_argument("--runs-name", default=None,
                    help="归档目录名（默认 tests/deep_runs/<时间戳>）")
    ap.add_argument("--slot", default="0/1",
                    help="槽位分区 N/M：只跑分到第 N 槽（0..M-1）的场景；"
                         "M=1 为单进程全量。全局状态场景永远属于槽 0。")
    known, _unknown = ap.parse_known_args()
    STATE["seed"] = known.seed
    STATE["only"] = known.only
    STATE["skip"] = {s.strip() for s in known.skip.split(",") if s.strip()}
    STATE["keep_word"] = known.keep_word
    try:
        _n, _, _m = known.slot.partition("/")
        slot_n, slot_m = int(_n), int(_m)
        assert 0 <= slot_n < slot_m and slot_m >= 1
    except Exception:
        print("[参数错误] --slot 格式应为 N/M（如 1/3）")
        return 2
    STATE["slot"] = (slot_n, slot_m)

    from dotenv import load_dotenv
    load_dotenv(os.path.join(_REPO, ".env"))

    print("=== deep_test 开始（全真实驱动：真实窗口 / 真实点击 / 真实 Word / "
          "真实 Atria）===")
    debug.init(["AI4Word", "-debug"])
    check(debug.is_debug() and debug.log_path(),
          "调试模式已激活，日志: " + str(debug.log_path()))
    W.delete_all(W.candidate_log_paths())
    STATE["t0"] = time.time()
    stamp = known.runs_name or datetime.datetime.fromtimestamp(
        STATE["t0"]).strftime("%Y%m%d-%H%M%S")
    # 多槽时归档到 tests/deep_runs/<时间戳>/slot<N>；单槽保持 tests/deep_runs/<时间戳>
    slot_n, slot_m = STATE["slot"]
    sub = "" if slot_m <= 1 else os.path.join("slot%d" % slot_n)
    STATE["run_dir"] = os.path.join(_RUNS, stamp, sub) if sub else \
        os.path.join(_RUNS, stamp)

    try:
        W.assemble()                       # QApplication + MainWindow + Tray + Worker（真实平台）
        worker, window, tray = (W.S["worker"], W.S["window"], W.S["tray"])
        _patch_doc_vars(worker)            # 合成命令：worker 线程读文档变量
        if slot_m > 1:
            # 并行槽：把查看类命令的重连绑定钉回本槽自己的 Word 实例
            # （_ensure_ready(launch=False) 默认 GetObject 附加到「某个」
            # 正在运行的 Word——并行时那是别的槽或用户的进程）。
            worker._attach_factory = lambda w=worker: w._app
        worker.start()
        pump(0.5)
        check(worker.isRunning(), "worker 线程启动")

        # 场景清单（顺序即执行顺序；名字与 deep_parallel.py 复用的清单一致）
        scenarios = [
            ("speed_gears",
             lambda: scenario_speed_gears(window, worker, W.S["settings"]),
             {"speed_clicked": 3, "speed_set": 3}),
            # write_block 是外部入口（编辑原语 / 存档回写 / 回滚恢复重放）的事件；
            # 紧凑框回车提交走流式链路 feed/flush → _write_md(keep_anchor=True)，
            # 落块只发 block_register。接线契约见 tests/test_event_wiring.py。
            ("compact_send", lambda: scenario_compact_send(window, worker),
             {"word_assembled": 1, "send_compact": 1, "write_sent": 1,
              "write_start": 1, "feed": 1, "flush": 1,
              "block_register": 1, "write_done": 1,
              "api_stream_done": 1}),
            ("block_map", lambda: scenario_block_map(window, worker),
             {"blockmap_toggled": 2, "block_clicked": 1,
              "select_block": 1}),
            ("interrupt_rollback",
             lambda: scenario_interrupt_real(window, worker, "rollback"),
             {"interrupt_requested": 1, "interrupted_ui": 1, "choose": 1,
              "choice_made": 1, "rollback_ok": 1, "snapshot_taken": 1}),
            ("interrupt_keep",
             lambda: scenario_interrupt_real(window, worker, "keep"),
             {"interrupted_ui": 1, "choose": 1, "choice_made": 1}),
            ("interrupt_extra",
             lambda: scenario_interrupt_real(
                 window, worker, "extra",
                 "再补充一段初冬清晨霜景的描写，150 字左右。"),
             {"interrupted_ui": 1, "enter_extra_mode": 1,
              "extra_submitted": 1, "extra_start": 1, "extra_done": 1}),
            ("panel_send", lambda: scenario_panel_send(window, worker),
             {"send_panel": 1, "write_sent": 1, "write_done": 1,
              "api_stream_done": 1}),
            ("keys",
             lambda: scenario_keys(window, worker, tray),
             {"send_compact": 1, "write_sent": 1, "write_done": 1,
              "close_to_tray": 1, "tray_activated": 2,
              "save_rejected_busy": 1, "preset_rejected_busy": 1}),
            ("arrange_ok", lambda: scenario_arrange_ok(window, worker),
             {"mode_clicked": 2, "arrange_sent": 1, "arrange_start": 1,
              "gen_code": 1, "gen_code_result": 1, "run_code_start": 1,
              "stmt_ok": 1, "run_code_done": 1, "arrange_applied": 1}),
            ("arrange_repair",
             lambda: scenario_arrange_repair(window, worker),
             {"arrange_start": 1, "stmt_failed": 1, "fix_code": 1,
              "fix_code_result": 1, "run_code2_result": 1,
              "arrange_applied": 1}),
            ("arrange_gen_fail",
             lambda: scenario_arrange_gen_fail(window, worker),
             {"arrange_start": 1, "gen_code_failed": 1,
              "arrange_failed": 1}),
            ("stream_fault", lambda: scenario_stream_fault(window, worker),
             {"write_start": 1, "stream_error": 1, "write_done": 1}),
            ("review_presets_save",
             lambda: scenario_review_presets_save(window, worker),
             {"review_toggled": 2, "review_set": 2, "preset_selected": 4,
              "preset_applied": 4, "save_clicked": 1, "save_blocks": 1}),
            ("expand_collapse",
             lambda: scenario_expand_collapse_avatar(window),
             {"toggle_expand": 2}),
            ("settings_dialog",
             lambda: scenario_settings_dialog(window, worker,
                                              W.S["settings"]),
             {"settings_open": 1, "settings_closed": 1}),
            ("tray_actions", lambda: scenario_tray_actions(window, tray,
                                                           worker),
             {"autostart_toggle": 2, "autostart_disabled": 1}),
            ("drag_snap_hotkey",
             lambda: scenario_drag_snap_hotkey(window),
             {"snap_to_edge": 1, "summon": 1, "hotkey_register": 1}),
            # esc_hide_and_hint 只产生 1 次 hide_via_esc：第一下 Esc 由
            # 展开态收起（keyPressEvent 在展开态走 _apply_expanded，不打
            # 日志），第二下 Esc 才隐藏。曾误写 2 被报成
            # hide_via_esc(1/2) 缺失；tests/test_deep_test_must.py 把
            # 「MUST 表 == 场景实际产出」锁在一起。
            ("esc_hide_hint", lambda: scenario_esc_hide_hint(window, tray),
             {"hide_via_esc": 1}),
            ("input_rejections",
             lambda: scenario_input_rejections(window, worker),
             {"send_rejected_too_long": 1}),
            ("doc_persistence",
             lambda: scenario_doc_persistence(window, worker),
             {"word_assembled": 1, "load_blocks": 1}),
            ("stress_mixed",
             lambda: scenario_stress_mixed(window, worker, W.S["settings"]),
             {"queue_put": 1}),   # 至少发出一条命令；其余 MUST 由随机 op 覆盖
        ]
        assert [n for n, _f, _m in scenarios] == list(SCENARIO_NAMES), \
            "场景清单与 SCENARIO_NAMES 不一致，deep_parallel 的分区会错"

        filtered = [n for n, _f, _m in scenarios
                    if (not STATE["only"] or STATE["only"] in n)
                    and not any(s in n for s in STATE["skip"])]
        STATE["allowed"] = slot_assignment(filtered, slot_n, slot_m)
        print("槽位 %d/%d：本槽场景 %d/%d%s"
              % (slot_n, slot_m, len(STATE["allowed"]), len(filtered),
                 "（全局状态场景固定 slot 0）" if slot_n == 0 else ""))

        for name, fn, must in scenarios:
            run_scenario(name, fn, must)

        if not STATE["keep_word"]:
            W.close_word(worker)      # 必须在 worker 存活时清场（COM apartment）
        W.quit_flow(worker, tray, W.S["settings"])
        W.log_triage()
    except SystemExit as e:
        # W.assemble 的环境失败（缺 key=2 / 单实例=3）统一归一到 2（环境错误）
        print("\n=== 环境错误，退出码 %s（归一为 2）===" % e.code)
        raise SystemExit(2)
    except Exception:
        traceback.print_exc()
        W.emergency_teardown()
        print("\n=== harness 异常，已紧急收场 ===")
        return 2
    finally:
        try:
            summary = archive_and_summarize()
        except Exception:
            traceback.print_exc()
            summary = None

    print("\n=== 结果 ===")
    for r in STATE["results"]:
        print("  %-20s %s" % (r["name"], "通过" if r["ok"] else "有发现"))
    if summary is None:
        return 2
    return summary["exit_code"]


if __name__ == "__main__":
    sys.exit(main())
