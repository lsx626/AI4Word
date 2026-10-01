# -*- coding: utf-8 -*-
"""GUI 后台引擎：一个 QThread 独占全部 Word COM 对象。

线程模型：所有 pywin32 调用（GetObject/Dispatch、读写 Selection、
StreamingWriter、DocModel、run_code）都在本线程内进行——run() 一开始
CoInitialize，结束时 CoUninitialize。GUI 线程只通过队列下发命令、
通过 Qt 信号接收状态，绝不直接触碰 COM 对象。

命令通过 send() 投递到队列；中断 / 回滚 / 追加补充三个动作要插队，
直接走专用 _choice_q，覆盖 _wait_choice 的等待。
"""
import os
import queue
import threading
import time

import pythoncom
from PySide6.QtCore import QThread, Signal

from app.agent import WRITER_SYSTEM, build_exec_globals, fix_code, gen_code
from doc_model import DocModel
from format_runner import run_code
from session import Session
from streaming_writer import StreamingWriter
from styles import apply_preset

try:
    import pywintypes
except ImportError:  # 非 Windows：仅用于异常归类
    pywintypes = None


class AgentWorker(QThread):
    # --- 状态信号 ---
    stateChanged = Signal(str)      # idle / connecting / writing / arranging
    wordStatus = Signal(str)        # Word 连接状态文本
    message = Signal(str, str)      # kind: user/assistant/info/code/error
    streamChunk = Signal(str)       # 阶段一流式原文增量
    streamStart = Signal()
    streamEnd = Signal()
    blockMap = Signal(str)          # 块地图文本（block_map() 的输出）
    interrupted = Signal()          # 流式被中断 → UI 给出 回滚 / 保留 / 追加
    progress = Signal(str)          # 实时进度文本（生成中/排版中状态条）

    def __init__(self, word_factory=None, parent=None):
        """word_factory：测试用注入的假 Word 构造器；默认是 app.agent.get_word。"""
        super().__init__(parent)
        if word_factory is None:
            from app.agent import get_word as _real_get_word
            word_factory = _real_get_word
        self._factory = word_factory
        self._queue = queue.Queue()
        self._choice_q = queue.Queue()
        self._interrupt = threading.Event()
        self._api_key = os.environ.get("ATRIA_API_KEY", "")
        self._speed = "auto"
        self._busy = False

        #_lazy：首次需要 Word 时才连接（GUI 启动不依赖 Word 是否已开）
        self._app = None
        self._doc = None
        self._sel = None
        self._writer = None
        self._model = None
        self._session = None
        self._exec_globals = None
        self._snap_before = None
        self._interrupted = False
        self._last_map = 0.0
        self._poll_acc = 0.0

    # ---------- 对外接口（GUI 线程调用） ----------

    def send(self, cmd, payload=None):
        self._queue.put((cmd, payload))

    def choose(self, value):
        """中断后的选择：'rollback' / 'keep' / ('extra', text)。"""
        self._choice_q.put(value)

    def set_api_key(self, key):
        self._api_key = (key or "").strip()

    # ---------- 线程主循环 ----------

    def run(self):
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            self.stateChanged.emit("idle")
            self.wordStatus.emit("未连接 Word")
            while True:
                try:
                    cmd, payload = self._queue.get(timeout=0.15)
                except queue.Empty:
                    self._idle_poll(0.15)
                    continue
                if cmd == "quit":
                    break
                self._handle(cmd, payload)
        finally:
            self._model = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        self.stateChanged.emit("idle")

    def _idle_poll(self, dt):
        """空闲时的看门狗轮询：检测用户手动改动并刷新块地图。"""
        if self._model is None:
            return
        self._poll_acc += dt
        if self._poll_acc < 0.5:
            return
        self._poll_acc = 0.0
        try:
            self._model.watch.poll()
            for ev in self._model.drain_events():
                self.message.emit("info", f"检测到你的手动改动：{ev}")
            now = time.time()
        except Exception:
            return
        try:
            if now - self._last_map > 1.5:
                self._last_map = now
                self.blockMap.emit(self._model.block_map())
        except Exception:
            pass

    def _set_progress(self, text):
        """向 UI 广播实时进度（状态条）；None 表示清空。"""
        if text is None:
            self.progress.emit("")
        else:
            self.progress.emit(str(text))

    def _handle(self, cmd, payload):
        try:
            if cmd in ("write", "arrange", "preset", "select_block", "save", "review", "refresh_map"):
                if not self._ensure_ready():
                    return
            if cmd == "write":
                self._cmd_write(payload)
            elif cmd == "arrange":
                self._cmd_arrange(payload)
            elif cmd == "speed":
                self._set_speed(payload)
            elif cmd == "review":
                self._cmd_review(payload)
            elif cmd == "preset":
                self._cmd_preset(payload)
            elif cmd == "select_block":
                self._cmd_select_block(payload)
            elif cmd == "save":
                self._cmd_save()
            elif cmd == "refresh_map":
                self._emit_map()
            elif cmd == "choice":
                self.message.emit("info", "当前没有待处理的中断选择。")
            else:
                self.message.emit("error", f"未知命令：{cmd}")
        except Exception as e:
            self.message.emit("error", f"执行失败：{e}")
            self._reset_if_dead(e)
            self._busy = False
            self._interrupted = False
            self.stateChanged.emit("idle")

    # ---------- Word 连接与环境装配 ----------

    def _ensure_ready(self):
        if self._model is not None:
            return True
        self.stateChanged.emit("connecting")
        self.wordStatus.emit("正在连接 Word…")
        self._set_progress("正在连接 Word…")
        try:
            app = self._factory()
        except Exception:
            app = None
        if app is None:
            self.message.emit("error", "无法连接 Microsoft Word，请先打开 Word 后再试。")
            self.wordStatus.emit("未连接 Word")
            self.stateChanged.emit("idle")
            return False
        try:
            doc = app.Documents.Add() if app.Documents.Count == 0 else app.ActiveDocument
            sel = app.Selection
            writer = StreamingWriter(app, doc, sel, model=None, char_delay=0.01)
            model = DocModel(app, doc, sel, writer)
            writer.model = model
            session = Session()
            try:
                restored = model.load_blocks()
            except Exception:
                restored = 0
            self._app, self._doc, self._sel = app, doc, sel
            self._writer, self._model, self._session = writer, model, session
            self._exec_globals = build_exec_globals(app, doc, sel, model, writer, session)
            name = "(未命名文档)"
            try:
                name = str(doc.Name) or name
            except Exception:
                pass
            self.wordStatus.emit(f"已连接 · {name}")
            if restored:
                self.message.emit("info", f"已从文档恢复 {restored} 个块索引。")
            else:
                # 首次接入没有任何存档的文档（例如用户新打开的非空文档）：
                # 把现有内容读进块模型，块地图与编辑原语即可作用于已有文字
                imported = self._import_if_nonempty(model)
                if imported:
                    self.message.emit("info", f"已读取现有文档 {imported} 个块。")
            self._emit_map()
        except Exception as e:
            self.message.emit("error", f"初始化 Word 环境失败：{e}")
            self.wordStatus.emit("未连接 Word")
            self.stateChanged.emit("idle")
            return False
        self.stateChanged.emit("idle")
        return True

    def _reset_if_dead(self, err):
        """COM 连接已断（Word 被关掉）时，下次命令重新连接。"""
        if pywintypes is None:
            return
        if isinstance(err, pywintypes.com_error):
            self._app = None
            self._model = None
            self.wordStatus.emit("未连接 Word")

    def _emit_map(self):
        if self._model is None:
            return
        try:
            self._last_map = time.time()
            self.blockMap.emit(self._model.block_map())
        except Exception:
            pass

    def _import_if_nonempty(self, model):
        """文档非空且无存档时，把现有段落登记为块；返回登记数。"""
        try:
            return model.import_document()
        except Exception as e:
            self.message.emit("error", f"读取现有文档失败：{e}")
            return 0

    # ---------- 阶段一：流式写入 ----------

    def _cmd_write(self, prompt):
        if self._busy:
            self.message.emit("info", "正在处理上一条指令，请先中断。")
            return
        if not (prompt or "").strip():
            return
        self._busy = True
        self._interrupt.clear()
        self._interrupted = False
        self.message.emit("user", prompt)
        self.stateChanged.emit("writing")
        self._writer.set_speed(self._speed)
        self._writer.reset_anchor()  # 新一波写入：锚点从当前光标重新捕获

        from ai_client import ai_stream
        system = WRITER_SYSTEM  # 只允许输出正文，禁止任何前言/讨论/提问
        try:
            self._snap_before = self._model.snapshot()
        except Exception:
            self._snap_before = None
        received = 0
        self._set_progress("正在生成…")
        self.streamStart.emit()
        try:
            for piece in ai_stream(prompt, self._api_key, system):
                if self._interrupt.is_set():
                    self._interrupted = True
                    break
                if not piece:
                    continue
                received += len(piece)
                self.streamChunk.emit(piece)
                self._writer.feed(piece)
                self._set_progress(f"正在生成… 已接收 {received} 字")
        except Exception as e:
            self.message.emit("error", f"流式生成出错：{e}")
        try:
            self._writer.flush()
        except Exception:
            pass
        self.streamEnd.emit()
        if received and not self._interrupted:
            self.message.emit("info", "已写入 Word。")
            self._set_progress(f"已写入 Word · 共 {received} 字")
        try:
            self._model.save_blocks()
        except Exception:
            pass
        self._emit_map()

        if self._interrupted:
            self._wait_choice()
        else:
            self._busy = False
            self.stateChanged.emit("idle")

    def _wait_choice(self):
        """中断后等用户选：回滚 / 保留 / 追加补充（超时 5 分钟按保留处理）。"""
        self.message.emit("info", "已中断。可回滚本次生成、保留现状，或追加补充内容。")
        self.interrupted.emit()
        try:
            choice = self._choice_q.get(timeout=300)
        except queue.Empty:
            choice = "keep"
        self._interrupted = False
        if choice == "rollback":
            if self._snap_before is not None:
                try:
                    self._model.restore_snapshot(self._snap_before)
                    self.message.emit("info", "已回滚到生成前的文档状态。")
                except Exception as e:
                    self.message.emit("error", f"回滚失败：{e}")
            else:
                self.message.emit("info", "没有可回滚的快照。")
            self._emit_map()
        elif isinstance(choice, tuple) and choice[0] == "extra":
            extra = (choice[1] or "").strip()
            if extra:
                self.message.emit("user", extra)
                self.stateChanged.emit("writing")
                from ai_client import ai_stream
                system = WRITER_SYSTEM
                self._writer.set_speed(self._speed)
                self._writer.reset_anchor()
                received = 0
                self._set_progress("正在追加生成…")
                self.streamStart.emit()
                try:
                    for piece in ai_stream(extra, self._api_key, system):
                        self.streamChunk.emit(piece)
                        self._writer.feed(piece)
                        received += len(piece)
                        self._set_progress(f"正在追加… 已接收 {received} 字")
                except Exception as e:
                    self.message.emit("error", f"追加生成出错：{e}")
                try:
                    self._writer.flush()
                except Exception:
                    pass
                self.streamEnd.emit()
                try:
                    self._model.save_blocks()
                except Exception:
                    pass
                self._emit_map()
        self._snap_before = None
        self._busy = False
        self.stateChanged.emit("idle")

    # ---------- 阶段二：AI 智能排版 ----------

    def _cmd_arrange(self, prompt):
        if self._busy:
            self.message.emit("info", "正在处理上一条指令，请先中断。")
            return
        if not (prompt or "").strip():
            return
        self._busy = True
        self._interrupt.clear()
        self.message.emit("user", prompt)
        self.stateChanged.emit("arranging")

        # 生成期间用户若动过文档，把漂移警告带给代码模型
        notices = []
        try:
            self._model.watch.poll()
            notices = self._model.drain_events()
        except Exception:
            pass
        for ev in notices:
            self.message.emit("info", f"检测到你的手动改动：{ev}")

        full_prompt = prompt
        if notices:
            full_prompt += ("\n\n注意：用户在我生成期间手动编辑了文档，块索引可能已偏移，"
                            "请先重新调用 block_map() 确认当前结构。")

        def sink(text):
            self.message.emit("code", str(text))

        code = None
        self._set_progress("正在请求 AI 生成排版代码…")
        try:
            code = gen_code(full_prompt, self._api_key, self._session,
                            block_map_fn=self._model.block_map, sink=sink)
        except Exception as e:
            self.message.emit("error", f"请求 AI 失败：{e}")
            code = None
        if code is None or self._interrupt.is_set():
            self.message.emit("info", "已取消本次排版指令。")
            self._busy = False
            self.stateChanged.emit("idle")
            return

        self._set_progress("正在执行排版…")
        try:
            ok, err = run_code(code, self._exec_globals, sink=sink)
        except Exception as e:
            ok, err = False, f"{type(e).__name__}: {e}"
        if ok:
            self.message.emit("info", "代码执行完毕，已应用。")
            self._session.record_turn(prompt, True)
            self._persist()
            self._emit_map()
            self._busy = False
            self.stateChanged.emit("idle")
            return

        if self._interrupt.is_set():
            self.message.emit("info", "已取消自我修复。")
            self._busy = False
            self.stateChanged.emit("idle")
            return
        corrected = None
        self._set_progress("排版失败，正在自我修复…")
        try:
            corrected = fix_code(full_prompt, code, err, self._api_key,
                                 self._session, sink=sink)
        except Exception as e:
            self.message.emit("error", f"自我修复请求失败：{e}")
        if corrected is None:
            self.message.emit("error", "AI 未能生成有效代码，请换个说法再试。")
            self._busy = False
            self.stateChanged.emit("idle")
            return
        try:
            ok2, err2 = run_code(corrected, self._exec_globals, sink=sink)
        except Exception as e:
            ok2, err2 = False, f"{type(e).__name__}: {e}"
        if ok2:
            self.message.emit("info", "修正代码执行完毕，已应用。")
            self._session.record_turn(prompt, True)
            self._persist()
            self._emit_map()
        else:
            self.message.emit("error", f"自我修复仍然失败（{err2}），试试更简单的指令。")
            self._session.record_turn(prompt, False)
        self._busy = False
        self.stateChanged.emit("idle")

    def _persist(self):
        try:
            self._model.save_blocks()
        except Exception:
            pass

    # ---------- 简单命令 ----------

    def _set_speed(self, mode):
        self._speed = mode if mode in ("auto", "slow", "fast") else "auto"
        if self._writer is not None:
            try:
                self._writer.set_speed(self._speed)
            except Exception:
                pass

    def _cmd_review(self, on):
        # 修订开关只翻转 TrackRevisions，不触碰文档内容，生成中切换无妨
        try:
            self._model.review_on() if on else self._model.review_off()
            self.message.emit("info", f"已{'开启' if on else '关闭'}修订模式。")
        except Exception as e:
            self.message.emit("error", f"修订模式切换失败：{e}")

    def _cmd_preset(self, name):
        if self._busy:
            self.message.emit("info", "正在处理上一条指令，请先中断再套用预设。")
            return
        try:
            result = apply_preset(self._doc, name)
            detail = "；".join(result.get("applied", [])) if isinstance(result, dict) else ""
            self.message.emit("info", f"已应用「{name}」预设。{detail}")
        except Exception as e:
            self.message.emit("error", f"应用预设失败：{e}")

    def _cmd_select_block(self, index):
        try:
            self._model.select_block(int(index))
        except Exception as e:
            self.message.emit("error", f"无法定位第 {index} 块：{e}")

    def _cmd_save(self):
        # 存档只是把当前块模型写进文档变量，生成中调用得到的是一致的中间快照
        try:
            self._model.save_blocks()
            self.message.emit("info", "块索引已随文档存档。")
        except Exception as e:
            self.message.emit("error", f"存档失败：{e}")

    def interrupt(self):
        """GUI 中断按钮：流式写入立即停；排版阶段等 AI 响应后取消。"""
        self._interrupt.set()
        if self._busy and not self._interrupted:
            # 正在等 AI（非写入中）时，不能立即停，标记后由 _cmd_arrange 检查
            pass
