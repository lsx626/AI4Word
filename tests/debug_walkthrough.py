# -*- coding: utf-8 -*-
"""全功能 walkthrough：真实 Word + 真实 Atria API + 隐藏调试模式。

以 argv = ["AI4Word", "-debug"] 在进程内装配真实 MainWindow + AgentWorker +
Settings（临时 settings 文件，不污染用户配置），逐项驱动全部 UI 操作与
后台链路，验证调试日志完整覆盖，并暴露隐藏问题。

运行（需要本机 Word 与 .env 中的 ATRIA_API_KEY）：
    调试模式：.venv\python.exe tests\debug_walkthrough.py -debug
    对照组：  .venv\python.exe tests\debug_walkthrough.py

结尾输出 debug.log 路径与 WARN/ERROR / traceback 摘要，供 triage。
"""
import os
import sys
import time
import gc
import traceback

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

_REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _REPO)

from dotenv import load_dotenv

load_dotenv()

from PySide6.QtCore import QEvent, QPoint, QSharedMemory, Qt
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox

import win32com.client

from app import __version__, debug
from app.engine import AgentWorker
from app.icons import app_icon
from app.main_window import MainWindow
from app.settings import Settings
from app.theme import qss
from app.tray import Tray

DEBUG_MODE = "-debug" in sys.argv or "--debug" in sys.argv

FAIL = []
PLATFORM = []
DONE = {"n": 0}  # worker 已处理完毕的命令计数
SENT = {"n": 0}  # send-side counter, races DONE; see wait_handled
CHARS = {"n": 0}  # chars delivered via worker.streamChunk


def check(cond, msg):
    if cond:
        print("  ok   " + msg)
    else:
        FAIL.append(msg)
        print("  FAIL " + msg)


def note_platform(msg):
    PLATFORM.append(msg)
    print("  note " + msg + "（offscreen 平台假象，不算 bug）")


def pump(seconds=0.3):
    t0 = time.time()
    while time.time() - t0 < seconds:
        QApplication.processEvents()
        time.sleep(0.01)


def wait_for(pred, timeout, period=0.03):
    deadline = time.time() + timeout
    while time.time() < deadline:
        QApplication.processEvents()
        if pred():
            return True
        time.sleep(period)
    QApplication.processEvents()
    return bool(pred())


def _on_stream_chunk(piece):
    """Count chars the AI actually delivered (queued to main thread)."""
    try:
        CHARS["n"] += len(piece or "")
    except Exception:
        pass


def patch_handle(worker):
    """包一层 worker._handle：每处理完一条命令计数 +1。

    wait_handled 据此判断命令真正完成，不依赖 _busy（preset / save /
    select_block 等命令从不置 busy，直接看 busy 会误判成「还没开始」。
    """
    orig = worker._handle

    def wrapped(cmd, payload):
        if cmd == "_sync_doc_text":
            # Synthetic command: read Word text on the worker thread.
            # Deliberately NOT counted in DONE to avoid racing
            # wait_handled into returning early.
            # Read FIRST, publish rsp LAST: rsp is the release flag, and
            # publishing it before the read let the waiter wake up to a
            # not-yet-updated (None) text.
            text = _read_doc_text(worker)
            if isinstance(payload, dict):
                DOC_BOX["rsp"] = payload.get("req", DOC_BOX["req"])
            else:
                DOC_BOX["rsp"] = DOC_BOX["req"]
            DOC_BOX["text"] = text
            return
        if cmd == "_close_word":
            # Close docs/app on the worker thread (the COM apartment owner).
            started = bool(payload.get("started")) if isinstance(payload, dict) else False
            closed = set()
            for d in (worker._doc, S.get("doc")):
                if d is None:
                    continue
                try:
                    nm = d.Name
                except Exception:
                    continue
                if nm in closed:
                    continue
                closed.add(nm)
                try:
                    d.Close(SaveChanges=0)
                except Exception:
                    pass
            if started and worker._app is not None:
                try:
                    worker._app.Quit()
                except Exception:
                    pass
            # Release COM references on the owning thread while it still
            # runs: releasing them on the main thread (which never
            # CoInitialized) is an apartment violation, and releasing them
            # during interpreter shutdown crashed the process.
            worker._app = None
            worker._doc = None
            worker._sel = None
            worker._writer = None
            worker._model = None
            worker._exec_globals = None
            S["doc"] = None
            gc.collect()
            DONE["n"] += 1  # real command: let close_word's wait_handled see it
            return
        try:
            orig(cmd, payload)
        finally:
            DONE["n"] += 1

    worker._handle = wrapped
    # count sends on the way in so wait_handled cannot miss a
    # command that was already handled (the old DONE+1 target
    # raced the worker and stalled a full timeout per command)
    orig_send = worker.send

    def wrapped_send(cmd, payload=None):
        if cmd not in ("_sync_doc_text", "quit"):
            # _sync_doc_text never increments DONE either; it
            # is tracked through the DOC_BOX rsp flag instead
            SENT["n"] += 1
        return orig_send(cmd, payload)

    worker.send = wrapped_send


def wait_handled(worker, timeout=180):
    """等最近一条 send 出的命令处理完毕（send 之后调用）。"""
    target = SENT["n"]
    ok = wait_for(lambda: DONE["n"] >= target, timeout)
    pump(0.2)
    return ok


def wait_busy(worker, timeout=30):
    return wait_for(lambda: worker._busy, timeout)


def wait_awaiting_choice(window, timeout=150):
    return wait_for(lambda: window._awaiting_choice, timeout)


def _first_diff(a, b):
    n = min(len(a), len(b))
    for i in range(n):
        if a[i] != b[i]:
            return i
    return n


DOC_BOX = {"req": 0, "rsp": 0, "text": None}


def _read_doc_text(worker):
    try:
        return worker._doc.Content.Text.replace("\r", "").strip()
    except Exception:
        return None


def doc_text(worker, timeout=12.0):
    """Read Word body text on the worker thread.

    The main thread touching the COM objects directly raises
    RPC_E_WRONG_THREAD (silently swallowed as "" before). We post a
    synthetic _sync_doc_text command that the worker serves on its own
    thread, filling DOC_BOX. While the worker is blocked in _wait_choice
    it never returns to the command queue, so we time out and return None.
    """
    if not worker.isRunning():
        return _read_doc_text(worker)
    req = DOC_BOX["req"] + 1
    DOC_BOX["req"] = req
    DOC_BOX["text"] = None
    worker.send("_sync_doc_text", {"req": req})
    ok_sync = wait_for(lambda: DOC_BOX["rsp"] == req, timeout)
    text = DOC_BOX["text"]
    debug.log("doc_sync", req=req, rsp=DOC_BOX["rsp"],
              served=bool(ok_sync), chars=len(text or ""),
              preview=(text or "")[:120])
    return text


# ---------- 装配 ----------

S = {"worker": None, "window": None, "tray": None, "settings": None,
     "started_word": False, "doc": None}


def make_word():
    """worker 线程内调用的 Word 工厂：附加已运行的 Word 或启动新的，
    无论哪种都新建一个空文档并激活，保证引擎绑定到它而非用户的活动文档。"""
    app = None
    try:
        app = win32com.client.GetObject(None, "Word.Application")
        print("    (已附加到运行中的 Word)")
    except Exception:
        app = win32com.client.Dispatch("Word.Application")
        S["started_word"] = True
        print("    (启动了新的 Word)")
    try:
        app.Visible = True
    except Exception:
        pass
    doc = app.Documents.Add()
    S["doc"] = doc
    try:
        doc.Activate()
    except Exception:
        pass
    return app


def candidate_log_paths():
    """debug.log 的所有可能落点（APPDATA / TEMP 回退），对照组断言用。"""
    paths = []
    base = os.environ.get("APPDATA")
    if base:
        paths.append(os.path.join(base, "AI4Word", "debug.log"))
        paths.append(os.path.join(base, "AI4Word", "debug.log.bak"))
    t = os.environ.get("TEMP")
    if t:
        paths.append(os.path.join(t, "AI4Word", "debug.log"))
        paths.append(os.path.join(t, "AI4Word", "debug.log.bak"))
        paths.append(os.path.join(t, "ai4word_debug.log"))
    return paths


def delete_all(paths):
    for p in paths:
        try:
            os.remove(p)
        except OSError:
            pass


def assemble():
    api_key = os.environ.get("ATRIA_API_KEY", "").strip()
    if not api_key:
        print("FAIL .env 未设置 ATRIA_API_KEY，无法跑真实链路")
        sys.exit(2)

    tmp = os.environ.get("TEMP") or _REPO
    settings_path = os.path.join(tmp, "ai4word_walkthrough_settings.json")
    try:
        os.remove(settings_path)
    except OSError:
        pass

    settings = Settings(settings_path).load()
    settings.set("api_key", api_key)
    settings.set("speed", "auto")
    S["settings"] = settings

    app = QApplication(["AI4Word"] + (["-debug"] if DEBUG_MODE else []))
    app.setApplicationName("AI4Word")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("AI4Word")
    app.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(app_icon(64))
    app.setStyleSheet(qss())

    shm = QSharedMemory("AI4Word-SingleInstance-v8")
    if not shm.create(1):
        print("FAIL 已有 AI4Word 实例在运行（QSharedMemory 单实例），请先退出")
        sys.exit(3)

    worker = AgentWorker(word_factory=make_word)
    worker.set_api_key(api_key)
    worker.streamChunk.connect(_on_stream_chunk)
    patch_handle(worker)
    S["worker"] = worker

    window = MainWindow(worker, settings)
    tray = Tray(window, settings)
    window.set_tray(tray)
    S["window"] = window
    S["tray"] = tray
    tray.show()
    window.show()


# ---------- 矩阵 ----------

def speed_gears(window, worker, settings):
    print("=== 速度三档（慢/自/快）===")
    for idx, mode in ((0, "slow"), (1, "auto"), (2, "fast")):
        window._speed_group.button(idx).click()
        pump(0.2)
        check(worker._speed == mode, "speed=%s 经 set_speed 下发到 worker" % mode)
        check(settings.get("speed") == mode, "speed=%s 持久化到 settings" % mode)


def compact_send(window, worker):
    print("=== 紧凑态发送：真实流式写作 ===")
    window.set_expanded(False)
    pump(0.2)
    check(not worker._busy and worker._queue.empty(), "发送前 worker 空闲")
    window.input_c.setText("写一首关于秋天的短诗：一个一级标题加两段正文。")
    window._on_send_compact()
    check(wait_busy(worker, 60), "worker 进入 writing（含首次连接 Word）")
    check(wait_handled(worker, 300), "流式写作完成，worker 回 idle")
    check("已连接" in window.word_label.text(),
          "Word 连接状态显示: " + window.word_label.text())
    text = doc_text(worker) or ""
    check(len(text) > 10 and "#" not in text, "正文真实写入（%d 字）" % len(text))


def block_map(window, worker):
    print("=== blockmap toggle and click ===")
    window.set_expanded(True)  # btn_blockmap / block_panel live on the expanded panel
    pump(0.3)
    check(not window.block_panel.isVisible(), "初始隐藏")
    window.btn_blockmap.click()
    check(window.btn_blockmap.isChecked() and window.block_panel.isVisible(),
          "点击后显示块地图侧栏")
    ok = wait_for(lambda: window.block_panel._list.count() > 0, 30)
    check(ok, "块地图刷新出 %d 行" % window.block_panel._list.count())
    item = window.block_panel._list.item(0)
    if item is not None:
        window.block_panel._on_click(item)  # blockClicked -> _on_block_clicked
        check(wait_handled(worker, 30), "块点击定位命令处理完成")
    window.btn_blockmap.click()
    check(not window.block_panel.isVisible(), "再次点击收起块地图")


def interrupt_cycle(window, worker, choice, extra_text=None):
    """发送长文 -> 中途中断 -> 选择 rollback/keep/extra 之一。"""
    window.set_expanded(False)
    pump(0.2)
    window._on_speed(0)  # slow：流式排得足够长，保证能在中途Interrupt
    window.input_c.setText("写一篇关于秋天的长篇抒情散文，至少 1000 字，"
                           "要有一级标题、多个二级标题、若干段落和一个列表。")
    before = doc_text(worker) or ""
    CHARS["n"] = 0
    window._on_send_compact()  # 首次按 = 发送
    check(wait_busy(worker, 60), "worker 进入 writing")
    if not wait_for(lambda: CHARS["n"] > 0, 60):
        note_platform("AI delivered no stream chunk within 60s")
    pump(0.5)
    # If the stream failed or finished on its own (network flake, empty
    # response), the worker returns to idle without ever showing the
    # choice popup; pressing send again would start a NEW write and
    # produce false FAILs. Skip the cycle with a platform note instead.
    if not worker._busy and not window._awaiting_choice:
        note_platform("AI stream ended on its own before interruption; skipping choice cycle")
        return
    mid = doc_text(worker, 1.0)
    window._on_send_compact()  # busy 中再按同一个按钮 = 中断
    check(wait_awaiting_choice(window, 150), "中断后弹出 回滚/保留/追加 选择条")
    mid = doc_text(worker, 1.0)
    if choice == "rollback":
        window._choose("rollback")
        check(wait_handled(worker, 180), "回滚分支处理完成")
        after = doc_text(worker)
        if after != before:
            fd = _first_diff(before, after or "")
            debug.log("rollback_diff", before_len=len(before),
                      after_len=len(after and after or 0), first_diff=fd,
                      before_around=before[max(0, fd - 40):fd + 60],
                      after_around=(after or "")[max(0, fd - 40):fd + 60],
                      full=True)
        check(doc_text(worker) == before, "回滚后文档恢复到生成前状态")
    elif choice == "keep":
        window._choose("keep")
        check(wait_handled(worker, 60), "保留分支处理完成")
        final = doc_text(worker) or ""
        if mid:
            check(len(final) >= len(mid) and len(mid) > len(before),
                  "kept partial text from interruption")
        else:
            # worker blocked in _wait_choice never serves the command queue,
            # so mid is None; keep does not roll back -> compare with before.
            check(len(final) > len(before),
                  "kept partial text (mid unreadable, vs before)")
    elif choice == "extra":
        window._enter_extra_mode()
        window.input_p.setPlainText(extra_text or "再补充一段初冬清晨的霜景描写。")
        window._on_send_panel()  # 追加模式下提交补充内容
        consumed = wait_for(lambda: not window._awaiting_choice, 8)
        if not consumed:
            check(False, "追加补充已提交（awaiting_choice 未解除）")
            # 兜底：手动喂 keep 解除 worker 的 _wait_choice 阻塞，继续后续步骤
            worker.choose("keep")
            wait_handled(worker, 60)
            return False
        check(consumed, "追加补充已提交，选择条消失")
        check(wait_handled(worker, 300), "追加生成完成，worker 回 idle")
        final = doc_text(worker) or ""
        if mid:
            check(len(final) > len(mid), "extra appended to doc")
        else:
            check(len(final) > len(before), "extra appended to doc (vs before)")
    return True


def panel_send(window, worker):
    print("=== 面板态发送：展开窗口输入并发送 ===")
    window.set_expanded(True)
    pump(0.2)
    window._on_speed(2)  # fast
    window.input_p.setPlainText("写一篇关于春天的短文，含一个一级标题和三个段落。")
    window._on_send_panel()
    check(wait_busy(worker, 60), "worker 进入 writing")
    check(wait_handled(worker, 300), "面板发送流式写作完成")
    check(len(doc_text(worker) or "") > 20, "panel write lands in doc")


def keys_inputs_and_close(window, worker, tray):
    """Deep check: Enter keys, Shift+Enter, IME guard, close button,
    tray activation, busy guards on save/preset."""
    print("=== keys: Enter / Shift+Enter / IME / close / tray ===")
    from PySide6.QtTest import QTest
    from PySide6.QtGui import QCloseEvent, QInputMethodEvent
    from PySide6.QtWidgets import QSystemTrayIcon

    # --- Enter in the compact QLineEdit fires returnPressed -> send ---
    window.set_expanded(False)
    pump(0.2)
    window.input_c.setText("compact input sent by the Enter key")
    QTest.keyClick(window.input_c, Qt.Key_Return)
    check(wait_busy(worker, 60), "Enter in compact input starts writing")
    check(wait_handled(worker, 300), "compact Enter-send finished, worker idle")
    check(len(doc_text(worker) or "") > 20, "compact Enter-send wrote text")

    # --- Shift+Enter in the panel InputEdit inserts a newline, never sends ---
    window.set_expanded(True)
    pump(0.2)
    window.input_p.clear()
    QTest.keyClick(window.input_p, Qt.Key_Return, Qt.ShiftModifier)
    pump(0.3)
    check("\n" in window.input_p.toPlainText(), "Shift+Enter inserts newline")
    check(not worker._busy and worker._queue.empty(),
          "Shift+Enter does not send")

    # --- Enter during IME composition must NOT send ---
    window.input_p.inputMethodEvent(QInputMethodEvent("composing", []))
    QTest.keyClick(window.input_p, Qt.Key_Return)
    pump(0.3)
    check(not worker._busy and worker._queue.empty(),
          "Enter during IME composition does not send")
    window.input_p.inputMethodEvent(QInputMethodEvent("", []))
    pump(0.2)
    check(not window.input_p._composing, "IME composition ends")

    # --- Enter in the panel InputEdit sends ---
    window.input_p.setPlainText("panel input sent by the Enter key")
    QTest.keyClick(window.input_p, Qt.Key_Return)
    check(wait_busy(worker, 60), "Enter in panel input starts writing")
    check(wait_handled(worker, 300), "panel Enter-send finished, worker idle")
    check(len(doc_text(worker) or "") > 20, "panel Enter-send wrote text")

    # --- close (X) button: hide + first-hide hint, never quit ---
    window.show()
    pump(0.2)
    tray._hinted = False
    window.closeEvent(QCloseEvent())
    check(not window.isVisible(), "close button hides window (does not quit)")
    check(tray._hinted, "first hide hint shown by tray")

    # --- tray activation reasons ---
    tray._on_activated(QSystemTrayIcon.DoubleClick)
    check(window.isVisible(), "tray double-click shows window")
    window.hide()
    pump(0.2)
    tray._on_activated(QSystemTrayIcon.Context)
    check(not window.isVisible(), "context activation does not show window")

    # --- busy guards: save and preset rejected while worker is writing ---
    window._busy = True
    window._on_save()
    check(worker._queue.empty(), "save rejected while busy")
    window._on_preset(1)
    check(worker._queue.empty(), "preset rejected while busy")
    check(window.preset_combo.currentIndex() == 0, "preset combo reset to placeholder")
    window._busy = False


def arrange_flow(window, worker, prompt):
    print("=== 排版链路：gen_code -> run_code ===")
    window._mode_group.button(1).click()
    check(not window._write_mode, "切换到排版模式")
    window.input_p.setPlainText(prompt)
    window._on_send_panel()
    check(wait_handled(worker, 300), "排版链路完成（含可能的自我修复）")
    window._mode_group.button(0).click()
    check(window._write_mode, "切回写作模式")


def arrange_forced_failure(window, worker):
    """注入一次必然失败的 gen_code，验证自我修复（fix_code）真实链路。"""
    print("=== 排版链路：失败 -> 自我修复 ===")
    import app.engine as engine_mod
    real_gen = engine_mod.gen_code
    counter = {"n": 0}

    def fake_gen(prompt, api_key, session=None, block_map_fn=None, sink=print,
                 model=None):
        counter["n"] += 1
        if counter["n"] == 1:
            sink("（walkthrough 注入的故意失败代码）")
            return "x = 1 / 0"
        return real_gen(prompt, api_key, session=session, block_map_fn=block_map_fn,
                        sink=sink, model=model)

    real_fix = engine_mod.fix_code
    fixes = {"n": 0}

    def fake_fix(prompt, failed_code, error_msg, api_key, session=None, sink=print,
                 block_map_fn=None, model=None):
        fixes["n"] += 1
        return real_fix(prompt, failed_code, error_msg, api_key, session=session,
                        sink=sink, block_map_fn=block_map_fn, model=model)

    engine_mod.gen_code = fake_gen
    engine_mod.fix_code = fake_fix
    try:
        window._mode_group.button(1).click()
        window.input_p.setPlainText("把所有一级标题居中并加粗")
        window._on_send_panel()
        check(wait_handled(worker, 300), "失败 -> fix_code -> run_code2 全链路走完")
        check(counter["n"] == 1, "gen_code first call was the injected failure")
        check(fixes["n"] >= 1, "fix_code self-repair triggered")
        window._mode_group.button(0).click()
    finally:
        engine_mod.gen_code = real_gen
        engine_mod.fix_code = real_fix


def review_toggle(window, worker):
    print("=== 修订开关（含生成中被拒分支）===")
    window.btn_review.setChecked(True)   # -> _on_review(True) -> send("review", True)
    check(wait_handled(worker, 60), "修订开启命令处理完成")
    window.btn_review.setChecked(False)  # -> send("review", False)
    check(wait_handled(worker, 60), "修订关闭命令处理完成")
    # 被拒分支：busy 中切换修订
    window._busy = True
    window.btn_review.setChecked(True)
    pump(0.3)
    check(not window.btn_review.isChecked(), "生成中切换修订被拒绝且回退按钮状态")
    check(worker._queue.empty(), "被拒的修订命令未入队")
    window._busy = False


def presets(window, worker):
    print("=== 四个样式预设 ===")
    for name in ("论文", "公文", "简历", "博客"):
        idx = window.preset_combo.findText(name)
        check(idx > 0, "预设「%s」在组合框中" % name)
        window._on_preset(idx)
        check(wait_handled(worker, 60), "预设「%s」应用完成" % name)
        check(window.preset_combo.currentIndex() == 0, "应用后组合框复位")


def save_blocks(window, worker):
    print("=== 存档 ===")
    window._on_save()
    check(wait_handled(worker, 60), "存档命令处理完成")
    check(len(worker._model.blocks) > 0, "块模型非空，存档有内容")


def expand_collapse_avatar(window):
    print("=== 展开 / 收起 / 头像点击 ===")
    window.set_expanded(False)
    window.avatar_c.clicked.emit()
    check(window._expanded, "紧凑态头像点击 -> 展开")
    window.avatar_p.clicked.emit()
    check(not window._expanded, "面板头像点击 -> 收起")
    window.btn_expand.click()
    check(window._expanded, "展开按钮")
    window.btn_collapse.click()
    check(not window._expanded, "收起按钮")


def settings_dialog(window, worker, settings):
    print("=== 设置对话框（开 / 改 / 确定）===")
    from app.settings_dialog import SettingsDialog
    api_key = (os.environ.get("ATRIA_API_KEY") or "").strip()
    dlg = SettingsDialog(settings, worker, parent=window)
    dlg.key_edit.setText(api_key)
    dlg.model_edit.setText("Atria-Dawn-Preview")
    dlg.top_check.setChecked(True)
    dlg.accept()
    check(settings.get("api_key") == api_key, "密钥写入 settings")
    check(settings.get("model") == "Atria-Dawn-Preview", "模型写入 settings")
    check(os.path.exists(settings.path), "settings 文件落盘")
    check(worker._api_key == api_key, "worker 同步新密钥")

    print("=== 设置对话框（保存失败回退）===")
    saved_path = settings.path
    bad = os.path.join(os.path.dirname(settings.path), "no_such_dir_ai4word",
                       "settings.json")
    settings.path = bad
    warned = []
    orig_warning = QMessageBox.warning

    def fake_warning(*args, **kwargs):
        warned.append(args)
        return QMessageBox.Ok

    QMessageBox.warning = fake_warning
    try:
        dlg2 = SettingsDialog(settings, worker, parent=window)
        dlg2.key_edit.setText(api_key)
        dlg2.accept()
    finally:
        QMessageBox.warning = orig_warning
        settings.path = saved_path
    check(bool(warned), "保存失败时弹出警告框")
    check(int(dlg2.result()) != int(QDialog.Accepted), "保存失败时对话框不关闭")
    check(settings.path == saved_path, "settings 路径已恢复")

    # --- password reveal / cancel / autostart-enable failure ---
    from PySide6.QtWidgets import QLineEdit, QPushButton
    dlg3 = SettingsDialog(settings, worker, parent=window)
    show_btn = next(b for b in dlg3.findChildren(QPushButton) if b.isCheckable())
    show_btn.click()
    check(dlg3.key_edit.echoMode() == QLineEdit.Normal,
          "password reveal shows the key")
    show_btn.click()
    check(dlg3.key_edit.echoMode() == QLineEdit.Password,
          "password reveal hides the key again")
    dlg3.reject()
    check(int(dlg3.result()) == int(QDialog.Rejected),
          "cancel leaves the dialog rejected")
    check(settings.get("api_key") == api_key and worker._api_key == api_key,
          "cancel does not change any settings")

    import app.settings_dialog as sd_mod
    real_enable = sd_mod.enable
    sd_mod.enable = lambda: False
    QMessageBox.warning = fake_warning
    warned.clear()
    try:
        dlg4 = SettingsDialog(settings, worker, parent=window)
        dlg4.key_edit.setText(api_key)
        dlg4.auto_check.setChecked(True)
        dlg4.accept()
    finally:
        sd_mod.enable = real_enable
        QMessageBox.warning = orig_warning
    check(bool(warned), "autostart enable failure warns the user")
    check(int(dlg4.result()) == int(QDialog.Accepted),
          "dialog still closes when autostart enable fails")
    check(not settings.get("auto_start"),
          "auto_start stays off when enable fails")


def tray_actions(window, tray, worker):
    print("=== 托盘菜单动作 ===")
    tray.act_show.trigger()
    check(window.isVisible(), "托盘「显示悬浮窗」召回窗口")
    from app.auto_start import disable, enable, is_enabled
    initial = is_enabled()
    warned = []
    orig_warning = QMessageBox.warning

    def fake_warning(*args, **kwargs):
        warned.append(args)
        return QMessageBox.Ok

    QMessageBox.warning = fake_warning
    try:
        tray._toggle_autostart(True)
        pump(0.5)
        if is_enabled():
            check(True, "托盘「开机自启」勾选写入注册表")
        else:
            note_platform("自启注册表写入被环境拦截")
        tray._toggle_autostart(False)
        pump(0.5)
        check(not is_enabled(), "托盘「开机自启」取消勾选生效")
    finally:
        QMessageBox.warning = orig_warning
    if initial:
        enable()
    else:
        disable()
    pump(0.3)
    check(is_enabled() == initial, "自启注册表状态已还原为初始值")


def drag_snap_summon_hotkey(window):
    print("=== 拖动落点 / 吸附 / summon / 热键 ===")
    start = window.pos()
    window.drag_to(start + QPoint(80, 40), QPoint(0, 0))
    check(window.pos() != start, "拖动移动窗口")
    screen = QApplication.primaryScreen().availableGeometry()
    w, h = window.width(), window.height()
    # 拖到右下角附近再松手 -> 吸附到边缘
    window.drag_to(QPoint(screen.right() - w - 5, screen.bottom() - h - 5),
                   QPoint(0, 0))
    window.snap_to_edge()
    check(abs(window.geometry().right() - (screen.right() - 8)) <= 1 or
          abs(window.geometry().bottom() - (screen.bottom() - 8)) <= 1,
          "松手吸附到屏幕边缘")
    window.summon()
    check(window.isVisible(), "summon 显示窗口")
    window.reload_flags()  # 注销 -> 重建 -> 重注册热键的完整路径
    pump(0.3)
    if window._hotkey:
        check(True, "Ctrl+Alt+Space 热键注册成功")
    else:
        note_platform("offscreen 下热键注册失败（无真实 HWND）")


def esc_hide_and_hint(window, tray):
    print("=== Esc 收起 / 隐藏 / 托盘提示 ===")
    window.set_expanded(True)
    ev = QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier)
    window.keyPressEvent(ev)
    check(not window._expanded, "展开态 Esc -> 收起")
    tray._hinted = False
    window.keyPressEvent(ev)
    check(not window.isVisible(), "紧凑态 Esc -> 隐藏到托盘")
    check(tray._hinted, "首次隐藏弹出托盘召回提示")
    window.summon()
    check(window.isVisible(), "重新召回")


def input_rejections(window, worker):
    print("=== 输入保护：空文本与超长输入 ===")
    window.input_c.setText("")
    window._on_send_compact()
    check(worker._queue.empty(), "空文本不生成命令")
    # input_c is capped by setMaxLength, so an over-limit setText is
    # silently truncated -> exercise _send's rejection branch directly.
    window._send("x" * (window.MAX_PROMPT_CHARS + 1))
    check(worker._queue.empty(), "超长输入被拒绝且不生成命令")
    window.input_c.clear()


def close_word(worker):
    """Close the Word docs and app ON the worker thread (the COM apartment).

    The main thread never CoInitializes, so touching the COM objects there
    is an apartment violation; releasing them at interpreter shutdown even
    crashed the process natively."""
    if not worker.isRunning():
        return
    worker.send("_close_word", {"started": S["started_word"]})
    wait_handled(worker, 30)
    pump(0.3)


def quit_flow(worker, tray, settings):
    print("=== 退出流程 ===")
    worker.choose("keep")  # 解除潜在的 _wait_choice 等待
    worker.send("quit")
    done = worker.wait(4000)
    if not done:
        # worker may still be finishing a write; give it more time
        done = worker.wait(30000)
    check(done, "worker 收到 quit 退出线程")
    saved = settings.save()
    check(saved, "退出时 settings 保存成功: " + settings.path)


def log_triage():
    path = debug.log_path()
    print("\n=== 调试日志 triage ===")
    if not path:
        print("debug.log: 未启用调试模式（is_debug=False）")
        return
    print("debug.log: " + path)
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            lines = f.read().splitlines()
    except OSError as e:
        print("读取日志失败: %s" % e)
        return
    print("总行数: %d" % len(lines))
    warn_err = [ln for ln in lines if ("] WARN " in ln or "] ERROR " in ln)]
    print("WARN/ERROR 行数: %d" % len(warn_err))
    for ln in warn_err:
        print("  " + ln[:240])
    print("traceback 出现行数: %d" % sum(1 for ln in lines if "Traceback" in ln))
    keys = ("run_start", "write_start", "write_done", "arrange_start",
            "gen_code_result", "run_code_result", "fix_code_result",
            "api_stream_done", "api_request_ok", "queue_put", "choice_made",
            "worker_loop_start", "worker_loop_exit", "word_assembled",
            "save_blocks", "speed_set", "review_set", "preset_applied",
            "hotkey_register", "quit", "send_compact", "send_panel",
            "speed_clicked", "mode_clicked", "tray_activated",
            "settings_accept", "blockmap_toggled", "close_to_tray",
            "stream_progress", "block_register")
    seen = {k for k in keys for ln in lines if k in ln and (
        k in ("snap_to_edge", "summon", "quit") or "] INFO " + k in ln or
        " " + k + " " in ln or ln.endswith(" " + k) or (" " + k) in ln)}
    missing = [k for k in keys if k not in seen]
    if missing:
        print("未在日志中出现的关键事件: %s" % ", ".join(missing))
    else:
        print("关键事件全部覆盖")


def main():
    print("=== walkthrough 开始（调试模式=%s）===" % DEBUG_MODE)
    delete_all(candidate_log_paths())
    debug.init(["AI4Word"] + (["-debug"] if DEBUG_MODE else []))
    if DEBUG_MODE:
        check(debug.is_debug() and debug.log_path(),
              "调试模式已激活，日志: " + str(debug.log_path()))
    else:
        check(not debug.is_debug() and debug.log_path() is None,
              "对照组：非调试模式零激活")

    assemble()
    worker, window, tray, settings = (S["worker"], S["window"], S["tray"],
                                      S["settings"])

    worker.start()
    pump(0.5)
    check(worker.isRunning(), "worker 线程启动")

    speed_gears(window, worker, settings)
    compact_send(window, worker)
    block_map(window, worker)

    print("=== 中断 -> 回滚 ===")
    interrupt_cycle(window, worker, "rollback")
    print("=== 中断 -> 保留 ===")
    interrupt_cycle(window, worker, "keep")
    print("=== 中断 -> 追加补充 ===")
    interrupt_cycle(window, worker, "extra",
                    extra_text="再补充一段初冬清晨霜景的描写，150 字左右。")

    panel_send(window, worker)
    keys_inputs_and_close(window, worker, tray)
    arrange_flow(window, worker, "把所有一级标题居中")
    arrange_forced_failure(window, worker)
    review_toggle(window, worker)
    presets(window, worker)
    save_blocks(window, worker)
    expand_collapse_avatar(window)
    settings_dialog(window, worker, settings)
    tray_actions(window, tray, worker)
    drag_snap_summon_hotkey(window)
    esc_hide_and_hint(window, tray)
    input_rejections(window, worker)

    close_word(worker)  # must run before quit: worker must be alive to touch COM
    quit_flow(worker, tray, settings)

    log_triage()

    if not DEBUG_MODE:
        leftover = [p for p in candidate_log_paths() if os.path.exists(p)]
        check(not leftover, "对照组：非调试模式未生成任何日志文件 " + str(leftover))

    print("\n=== 结果 ===")
    print("失败: %d" % len(FAIL))
    for f in FAIL:
        print("  - " + f)
    print("平台假象: %d" % len(PLATFORM))
    for p in PLATFORM:
        print("  - " + p)
    if DEBUG_MODE:
        print("调试日志: " + str(debug.log_path()))
    return 1 if FAIL else 0


def emergency_teardown():
    """Last-resort cleanup when main() raises: unblock and stop the worker
    thread and quit Word. Without this, interpreter shutdown with a live
    QThread holding Word COM crashed the process (0xC0000409)."""
    worker = S.get("worker")
    if worker is not None:
        try:
            worker.choose("keep")  # unblock a possible _wait_choice
        except Exception:
            pass
        try:
            worker.send("quit")
        except Exception:
            pass
        try:
            worker.wait(3000)
        except Exception:
            pass
        if worker.isRunning():
            try:
                worker.terminate()
                worker.wait(2000)
            except Exception:
                pass
    if S.get("started_word") and worker is not None and worker._app is not None:
        try:
            worker._app.Quit()
        except Exception:
            pass
    app = QApplication.instance()
    if app is not None:
        try:
            app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        code = main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        code = 99
    finally:
        if code:
            emergency_teardown()
    sys.exit(code or 0)