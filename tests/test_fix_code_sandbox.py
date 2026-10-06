# -*- coding: utf-8 -*-
"""离线回归：自我修复链路必须能收敛回 arrange_applied。

事故重放（tests/deep_runs 真机日志）：注入 1/0 故障后，fix_code 返回的
修复代码以 `import win32com.client as win32` 开头，被沙箱判
「禁止的语句: Import」（run_code2 ok=false）-> arrange_failed，
arrange_applied 从此再不出现。根因不是沙箱（它判得对），而是 fix_code
的提示词没有沙箱禁令清单，修复代码结构性不可能过沙箱。这里验证：
1. fix_code 提示词携带禁令清单（import / while / with / 双下划线 / 危险内建）；
2. 修复代码撞沙箱时，SandboxError 文案与被拒代码写回下一轮提示词重修；
3. 重修次数有上限，超过后原样交回执行层（run_code2 仍报精确错）；
4. 引擎级全链路：gen 故障 -> fix 撞沙箱 -> 重修成功 -> arrange_applied。
5. 修复成功分支的 arrange_applied 必须真正落盘进 debug.log：真机事故
   （20261005-213146 / 230853 / 231023 / 231406 连续 4 轮 slot2 的
   MUST 缺失）里 _cmd_arrange 的 if ok2: 分支只发用户消息、漏发终态
   事件；旧离线断言只检查消息文本，缺口在离线下永远看不到。这里开启
   真实调试日志并读文件（不是打桩 debug.log 调用），把接线锁死。
6. 符号契约（CODEGEN_SYSTEM 同源）：20261005-234719 slot2 与
   20261005-231023 slot0 的修复代码虽然过了沙箱却在运行时死掉——
   `for b in block_map(): b.get('type')`（block_map() 返回多行字符串，
   迭代出字符 -> AttributeError: 'str' object has no attribute 'get'）
   与 `doc.Styles("Heading 1")`（中文 Word 报「集合所要求的成员不存在」）。
   根因是修复提示词不含 gen_code 的 API 契约，模型只能猜 API 形状；
   且静态沙箱看不出这类错（事故代码合法），唯一防线就是提示词本身。
   这里锁死两件事：修复提示词携带契约，且契约与运行时真相一致
   （block_map() 真返字符串、行格式真是 `[索引] 类型: 预览`）。
"""
import ast
import importlib.util
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ai_client
from app import debug
from app.agent import (CODEGEN_SYSTEM, FIX_MAX_SANDBOX_RETRIES, _static_reject_reason,
                       fix_code)
from sandbox import SANDBOX_RULES_TEXT


# ---------- 单元级：fix_code 提示词与重修循环 ----------

def _patch_ai(seq):
    """按顺序返回伪造响应，并记录每次请求的 (user, system) 提示词。"""
    calls = []

    def fake(prompt, api_key, system_prompt, model=None):
        calls.append({"prompt": prompt, "system": system_prompt})
        if calls and len(calls) <= len(seq):
            return seq[len(calls) - 1]
        return "insert_at_end('兜底代码')"

    ai_client.ai_request = fake
    return calls


def test_prompt_carries_sandbox_rules():
    calls = _patch_ai(["insert_at_end('修复正文')"])
    sink_out = []
    code = fix_code("把标题居中", "x = 1 / 0",
                    "ZeroDivisionError: division by zero", "key",
                    sink=sink_out.append)
    assert code == "insert_at_end('修复正文')", code
    system = calls[0]["system"]
    # 禁令清单来自 sandbox.py 的实际禁令（单一真相源），不能各写一份
    assert SANDBOX_RULES_TEXT in system
    for word in ("import", "while", "with", "双下划线", "只输出纯 Python"):
        assert word in system, word
    assert calls[0]["prompt"] == "把标题居中"
    assert any("自我修复" in t for t in sink_out)
    print("ok: fix_code 提示词携带沙箱禁令清单（import/while/with/双下划线/危险内建）")


def test_repair_hitting_sandbox_gets_retried_with_reason():
    # 第一次修复原样重演事故：以 import 开头，必被沙箱判死
    bad = "import win32com.client as win32\napp = win32.Dispatch('Word.Application')"
    calls = _patch_ai([bad, "insert_at_end('修复正文')"])
    sink_out = []
    code = fix_code("排版", "x = 1 / 0",
                    "ZeroDivisionError: division by zero", "key",
                    sink=sink_out.append)
    assert code == "insert_at_end('修复正文')", code
    assert len(calls) == 2, len(calls)  # 撞沙箱 -> 重修一次，不多打
    system2 = calls[1]["system"]
    # SandboxError 文案与被拒代码都写回下一轮
    assert "SandboxError" in system2 and "禁止的语句: Import" in system2
    assert bad in system2
    assert any("未通过预检" in t for t in sink_out)
    print("ok: 修复代码撞沙箱后带 SandboxError 文案与被拒代码重修")


def test_syntax_error_repair_also_retried():
    calls = _patch_ai(["def broken(:", "insert_at_end('修复正文')"])
    code = fix_code("排版", "x = 1 / 0",
                    "ZeroDivisionError: division by zero", "key")
    assert code == "insert_at_end('修复正文')", code
    assert len(calls) == 2
    assert "SyntaxError" in calls[1]["system"]
    print("ok: 修复代码语法错误同样进入重修循环")


def test_retry_budget_is_bounded():
    # AI 顽固地返回 import：用尽预算后原样交回（执行层报精确错），不无限请求
    calls = _patch_ai(["import os"] * 16)
    code = fix_code("排版", "x = 1 / 0",
                    "ZeroDivisionError: division by zero", "key")
    assert "import os" in code, code
    assert len(calls) == FIX_MAX_SANDBOX_RETRIES + 1, len(calls)
    print(f"ok: 重修预算有上限（{FIX_MAX_SANDBOX_RETRIES + 1} 轮后交回执行层）")


def test_empty_ai_reply_still_returns_none():
    _patch_ai([None])
    assert fix_code("排版", "x = 1 / 0", "ZeroDivisionError", "key") is None
    print("ok: AI 空回复仍返回 None（契约不变）")


# ---------- 符号契约：修复提示词与 gen_code 同一份 API 说明 ----------

# 交给修复模型的假失败代码：正是 20261005-234719 slot2 的事故形态
# （把 block_map() 当字典列表迭代），错误信息也是事故原文
_INCIDENT_CODE = "for b in block_map():\n    if b.get('type') == 'heading1':\n        idx = b.get('index')"
_INCIDENT_ERR = "AttributeError: 'str' object has no attribute 'get'"


def test_prompt_carries_codegen_contract():
    """修复提示词必须以 CODEGEN_SYSTEM 为前缀并带上两条例外警示。

    真机事故（20261005-234719 slot2 / 20261005-231023 slot0）里修复模型
    猜 API 形状：把 block_map() 返回的多行字符串当字典列表迭代、用
    样式名字符串取样式。根因是修复提示词不含 gen_code 所用的符号契约
    —— 这里锁死契约随修复提示词下发，且与 gen_code 是同一份
    （CODEGEN_SYSTEM 本体，不是手抄的第二份，避免漂移）。
    """
    calls = _patch_ai(["insert_at_end('修复正文')"])
    fix_code("把所有一级标题加粗", _INCIDENT_CODE, _INCIDENT_ERR, "key",
             block_map_fn=lambda: "[0] heading1: 一级标题\n[1] para: 正文")
    system = calls[0]["system"]
    # 契约是 CODEGEN_SYSTEM 本体（单一真相源），不是另写一份
    assert system.startswith(CODEGEN_SYSTEM), "修复提示词未以 CODEGEN_SYSTEM 为前缀"
    # 契约的关键符号都在：块原语签名 / 常量 / 三大黄金法则
    for sym in ("block_map():", "get_block_text(i)", "replace_block(i, new_md)",
                "insert_after(i, md)", "delete_block(i)", "replace_text(old, new)",
                "WD_STYLE_HEADING_1", "WD_STYLE_NORMAL", "三大黄金法则",
                "每行 `[索引] 类型: 内容预览`", "索引从 0 开始"):
        assert sym in system, "契约缺失: " + sym
    # 事故形态的具名警示（block_map 返回字符串、样式须用数字常量）
    assert "多行字符串" in system, "未警示 block_map() 返回字符串"
    assert 'doc.Styles("Heading 1")' in system, "未警示样式名字符串在中文 Word 失效"
    # 失败上下文仍然携带（这是修复提示词的本职）
    for sym in ("FAILED CODE", "CURRENT BLOCK MAP", _INCIDENT_CODE, _INCIDENT_ERR,
                SANDBOX_RULES_TEXT):
        assert sym in system, "失败上下文缺失: " + sym
    print("ok: 修复提示词携带 CODEGEN_SYSTEM 符号契约（block_map 形状/WD_STYLE 常量/黄金法则）+ 事故警示")


def test_prompt_contract_matches_real_block_map():
    """契约必须与运行时真相一致：block_map() 真返多行字符串。

    双向锁死：提示词说「每行 `[索引] 类型: 内容预览` 的多行字符串」，
    真实的 DocModel.block_map() 就必须真是这个形状；任何一侧漂移
    （比如哪天 block_map 改返字典列表）都会让另一侧的契约变成谎言，
    修复模型照提示词写必定运行时炸 —— 这里的断言让漂移当场失败。
    同时重演事故代码本身：静态沙箱判不出（合法的 for + 属性访问），
    只有运行时才暴露 AttributeError —— 说明这条防线只能在提示词层。
    """
    app, _writer, model = _fake.make()
    model.insert_at_end("# 一级标题\n\n正文段落\n\n## 二级标题")
    bm = model.block_map()
    # 真相侧：返回值就是字符串本体，不是字典/列表
    assert isinstance(bm, str), type(bm)
    lines = [ln for ln in bm.splitlines() if ln and not ln.startswith("[!]")]
    assert lines, bm
    # 每行格式真是 `[索引] 类型: 预览`，索引从 0 开始连续
    for ln in lines:
        assert re.match(r"^\[\d+\] \S.*: ", ln), ln
    assert lines[0].startswith("[0] "), lines[0]
    kinds = {ln.split("] ", 1)[1].split(": ", 1)[0] for ln in lines}
    assert "heading1" in kinds, kinds
    # 事故代码静态过沙箱（for + .get 都合法）—— 防线不在沙箱层
    assert _static_reject_reason(_INCIDENT_CODE) is None
    # 运行时确实炸 AttributeError：这正是提示词必须警示的失败形态
    try:
        exec(compile(ast.parse(_INCIDENT_CODE), "<incident>", "exec"),
             {"block_map": model.block_map})
    except AttributeError as e:
        assert "get" in str(e), e
    else:
        raise AssertionError("事故代码竟未复现 AttributeError——block_map 形状已漂移")
    # 契约侧：提示词对形状的描述与上面验证的真相一致
    calls = _patch_ai(["insert_at_end('修复正文')"])
    fix_code("排版", _INCIDENT_CODE, _INCIDENT_ERR, "key",
             block_map_fn=model.block_map)
    system = calls[0]["system"]
    assert "多行字符串" in system
    assert "每行 `[索引] 类型: 内容预览`" in system
    assert model.block_map() in system, "修复提示词未携带当前 block_map 原文"
    print("ok: 契约与真相一致——block_map() 返回多行字符串，事故代码运行时炸 AttributeError（静态沙箱看不出）")


# ---------- 引擎级：事故全链路重放 ----------

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
_spec = importlib.util.spec_from_file_location(
    "fake_word_mod_fix", os.path.join(os.path.dirname(__file__), "test_fake_word.py"))
_fake = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_fake)

from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from app.engine import AgentWorker
_app = QApplication.instance() or QApplication([])


class _AppShell:
    """FakeApp 没有 Documents / ActiveDocument，补上引擎连接所需的壳。"""

    def __init__(self, app):
        self._app = app

    @property
    def Documents(self):
        return self

    @property
    def Count(self):
        return 1

    def Add(self):
        return self._app.doc

    @property
    def ActiveDocument(self):
        return self._app.doc

    @property
    def Selection(self):
        return self._app.Selection

    def __getattr__(self, name):
        return getattr(self._app, name)


class _Recorder:
    def __init__(self, worker):
        self.messages = []
        self.states = []
        worker.message.connect(lambda k, t: self.messages.append((k, t)))
        worker.stateChanged.connect(lambda s: self.states.append(s))

    def texts(self, kind):
        return [t for k, t in self.messages if k == kind]

    def wait_for(self, pred, timeout_ms=10000):
        t0 = time.time()
        while not pred():
            QTest.qWait(20)
            if (time.time() - t0) * 1000 > timeout_ms:
                raise AssertionError("等待超时: " + repr(self.messages[-6:]))


def _start_worker():
    app, _w, _m = _fake.make()
    worker = AgentWorker(word_factory=lambda: _AppShell(app))
    worker.set_api_key("test-key")
    rec = _Recorder(worker)
    worker.start()
    QTest.qWait(100)
    return worker, rec, app


def _finish(worker):
    worker.send("quit")
    worker.wait(4000)


def _scratch_debug_dir():
    """调试日志草稿目录：按 PID 隔离，避免与 test_debug_mode 的 TEMP 根
    debug.log 或 %APPDATA% 的真实运行日志混淆。

    必须按进程隔离：本仓库的并行修复会同时跑多份 run_offline（每份都会
    执行本文件），共享同一确定性路径会让两个进程交替 删除/追加 同一个
    debug.log——上一轮就在两个进程的并发写里撕出了半截 0xae 字节，
    读回来触发 UnicodeDecodeError。PID 目录使各进程互不可见。"""
    base = os.environ.get("TEMP") or os.path.expanduser("~") or "."
    d = os.path.join(base, "ai4word_test_fix_code_debug_%d" % os.getpid())
    try:
        os.makedirs(d, exist_ok=True)
        return d
    except OSError:
        pass
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "_fix_code_debug_scratch_%d" % os.getpid())
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


def _capture_debug_log(monkeypatch):
    """在 worker 启动前开启真实调试日志（写草稿目录），返回日志文件路径。

    深度真机测试的 MUST 事件解析的是 debug.log 文件本身，本回归同样读
    真实文件而不是打桩 debug.log 调用——离线必须覆盖到「事件真的落盘」。

    必须在 worker 线程第一次 debug.log 之前完成 init：is_debug() 的自动
    检测按真实 sys.argv 走，会先把调试模式锁死为 False，幂等的 init()
    之后就再也开不进来。hooks 预置为 True，避免本测试改动进程级的
    sys.excepthook / threading.excepthook。
    """
    d = _scratch_debug_dir()
    for name in ("debug.log", "debug.log.bak"):
        try:
            os.remove(os.path.join(d, name))
        except OSError:
            pass
    monkeypatch.setattr(debug, "_state",
                        {"debug": None, "path": None, "hooks": True,
                         "failed": False})
    monkeypatch.setattr(debug, "_log_dir", lambda: d)
    assert debug.init(["AI4Word.exe", "-debug"]) is True
    path = debug.log_path()
    assert path and os.path.dirname(path) == d, path
    return path


def _read_log(path):
    """读取 debug.log 全文（调试模块单行追加写、每行独立 open/write，
    落读即所见）。"""
    with open(path, encoding="utf-8") as f:
        return f.read()


def _log_events(text):
    """解析 'YYYY-MM-DD HH:MM:SS.mmm [thread] LEVEL event ...' 的事件名。

    exc() 的 traceback 字段会内嵌换行（_fmt_field 不转义 \n），续行按
    同规则切词后长度不足或第 4 词不是级别名，会被自然跳过。
    """
    events = []
    for line in text.splitlines():
        parts = line.split()
        if len(parts) >= 5 and parts[3] in ("INFO", "WARN", "ERROR"):
            events.append(parts[4])
    return events


def test_arrange_converges_after_sandbox_rejected_repair(monkeypatch):
    # 调试日志在 worker 启动前开启：整条自我修复链的事件都会落盘
    log_path = _capture_debug_log(monkeypatch)
    calls = {"n": 0}

    def fake_request(prompt, api_key, system, model=None):
        calls["n"] += 1
        if calls["n"] == 1:
            return "x = 1 / 0"  # 注入故障：第一段代码必失败
        if calls["n"] == 2:
            # 事故原样：修复代码以 import 开头，沙箱必判死
            return "import win32com.client as win32\nwin32.Dispatch('Word.Application')"
        return "insert_at_end('修复后的正文')"  # 第二版修复：合规且可执行

    ai_client.ai_request = fake_request
    worker, rec, app = _start_worker()
    try:
        worker.send("arrange", "加一段")
        rec.wait_for(lambda: any("修正代码执行完毕" in t for t in rec.texts("info"))
                       and rec.states[-1] == "idle")
        assert "修复后的正文" in app.doc.content
        assert calls["n"] == 3, calls["n"]  # gen + 两轮修复
        # 重修过程对用户可见（未通过预检 -> 带禁区重修）
        assert any("未通过预检" in t for t in rec.texts("code")), rec.texts("code")
        assert rec.states[-1] == "idle"
        # 修复成功分支必须发 arrange_applied（与首轮成功分支对齐）：
        # arrange_repair 场景的 MUST 靠它判定收敛，终态只能是它。
        # 读真实 debug.log 文件——20261005 连续 4 轮真机 slot2 正是
        # 该分支漏发事件而旧离线断言只看消息文本，掩盖了缺口。
        events = _log_events(_read_log(log_path))
        assert events.count("arrange_applied") == 1, events
        assert "arrange_failed" not in events, events
    finally:
        _finish(worker)
    print("ok: 引擎级全链路 gen 故障 -> 修复撞沙箱 -> 重修 -> arrange_applied 落盘 1 条，无 arrange_failed")


class _Patch:
    """直接运行本文件时的 monkeypatch 替身（run_offline / pytest 注入真的）。"""

    def __init__(self):
        self._undo = []

    def setattr(self, obj, name, value):
        self._undo.append((obj, name, getattr(obj, name, None)))
        setattr(obj, name, value)

    def undo(self):
        while self._undo:
            obj, name, old = self._undo.pop()
            setattr(obj, name, old)


if __name__ == "__main__":
    test_prompt_carries_sandbox_rules()
    test_prompt_carries_codegen_contract()
    test_prompt_contract_matches_real_block_map()
    test_repair_hitting_sandbox_gets_retried_with_reason()
    test_syntax_error_repair_also_retried()
    test_retry_budget_is_bounded()
    test_empty_ai_reply_still_returns_none()
    mp = _Patch()
    try:
        test_arrange_converges_after_sandbox_rejected_repair(mp)
    finally:
        mp.undo()
    print("\nfix_code 沙箱收敛 + 符号契约测试全部通过。")
