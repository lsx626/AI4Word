# -*- coding: utf-8 -*-
"""悬浮窗主体：紧凑胶囊 ↔ 完整面板的双形态无边框窗口。

- 无边框 + 置顶（可设置）；不透明窗口（WA_TranslucentBackground 的分层
  窗口在部分显卡上会被输入光标闪烁 / 头像 30fps 局部更新闪烁穿透，露出桌面）。
  圆角由系统合成器原生绘制：Win11 走 DWM 圆角（抗锯齿），早于 Win11
  降级 QRegion mask——不要在 Win11 上用 mask，它会盖掉原生圆角留下锯齿。
- 紧凑态：发光头像 + 输入框 + 发送/展开按钮
- 展开态：消息流 / 工具条（档位·修订·预设·存档）/ 块地图侧栏 / 中断选择条
- 头像与标题栏的空白条可拖动窗口，拖动中靠近屏幕边缘磁性吸附，松手再吸附一次
- Ctrl+Alt+Space 召唤
"""
import ctypes
import ctypes.wintypes as wt
import sys

from PySide6.QtCore import (QEasingCurve, QPropertyAnimation, QRect, QSize, Qt,
                            Signal)
from PySide6.QtGui import (QColor, QLinearGradient, QPainter, QPalette, QPainterPath,
                           QPen, QRegion)
from PySide6.QtWidgets import (QApplication, QButtonGroup, QComboBox, QHBoxLayout,
                               QLabel, QLineEdit, QPushButton, QSplitter,
                               QTextEdit, QVBoxLayout, QWidget)

from app import debug
from app.avatar import Avatar
from app.icons import (app_icon, icon_collapse, icon_expand, icon_gear, icon_keep,
                       icon_list, icon_rollback, icon_save, icon_send, icon_stop)
from app.messages import BlockMapPanel, MessageList
from app.theme import AMBER, INK_0, INK_1, INK_4, TEXT_DIM

try:
    from styles import preset_list
except Exception:  # 打包/单文件降级
    def preset_list():
        return ["论文", "公文", "简历", "博客"]

COMPACT_SIZE = QSize(392, 66)
EXPANDED_SIZE = QSize(492, 600)
SNAP_MARGIN = 40          # 松手时离边缘多少像素内吸附
SNAP_DRAG_MARGIN = 24     # 拖动中的磁性吸附半径（比松手小，避免拖动时黏得太死）
SNAP_PAD = 8              # 吸附后与边缘保持的间距
CORNER_RADIUS = 8         # 描边半径：与 Win11 DWM 原生圆角一致

# DWM 窗口圆角（Windows 11 build 22000+ 支持，原生抗锯齿）：
# 33 = DWMWA_WINDOW_CORNER_PREFERENCE，2 = DWMWCP_ROUND
_DWMWA_WINDOW_CORNER_PREFERENCE = 33
_DWMWCP_ROUND = 2
_DWM_TRY_BUILD = 22000   # 支持 DWM 圆角的最低 Windows 11 内部版本


class DragStrip(QWidget):
    """标题栏里的空白拖动条：按下拖动窗口（磁性贴边），松手吸附。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.SizeAllCursor)
        self._offset = None

    def _window(self):
        return self.window()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._offset = event.globalPosition().toPoint() - self._window().pos()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._offset is not None:
            win = self._window()
            if hasattr(win, "drag_to"):
                win.drag_to(event.globalPosition().toPoint(), self._offset)
            else:
                win.move(event.globalPosition().toPoint() - self._offset)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._offset is not None:
            self._offset = None
            win = self._window()
            if hasattr(win, "snap_to_edge"):
                win.snap_to_edge()
        super().mouseReleaseEvent(event)


class InputEdit(QTextEdit):
    """展开态输入框：Enter 发送，Shift+Enter 换行；内容超高时出现滚动条。"""

    def __init__(self, on_send, parent=None):
        super().__init__(parent)
        self._on_send = on_send
        self.setAcceptRichText(False)
        self.setFixedHeight(56)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setPlaceholderText("写点什么… 例如：写一篇关于秋天的散文")
        self._composing = False

    def inputMethodEvent(self, event):
        # 记录 IME 组合态：组合未提交时按 Enter 是「确认候选词」，
        # 不能让它冒泡成发送（依赖具体输入法是否消费该 Enter，此处主动防御）
        try:
            self._composing = bool(event.preeditString())
        except Exception:
            self._composing = False
        super().inputMethodEvent(event)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter) and \
                not (event.modifiers() & Qt.ShiftModifier) and \
                not self._composing:
            self._on_send()
            return
        super().keyPressEvent(event)


class MainWindow(QWidget):
    arrived = Signal()  # 召唤动画结束（预留）

    def __init__(self, worker, settings, parent=None):
        super().__init__(parent)
        self.worker = worker
        self.settings = settings
        self._tray = None
        self._busy = False
        self._awaiting_choice = False
        self._extra_mode = False
        self._write_mode = True
        self._expanded = False
        self._anim = None
        self._hotkey = False
        self._err_timer = None

        self.set_flags()
        self._apply_palette()
        self.setWindowTitle("AI4Word 悬浮助手")
        self.setWindowIcon(app_icon(64))

        self._build_ui()
        self._wire_worker()
        self.worker.send("speed", settings.get("speed") or "auto")

        self._apply_expanded(bool(settings.get("expanded")), instant=True)
        self.resize(COMPACT_SIZE if not self._expanded else EXPANDED_SIZE)
        self._place_initial()

    # ---------- 窗口形态 ----------

    def set_flags(self):
        flags = Qt.FramelessWindowHint | Qt.Tool
        if self.settings.get("stay_on_top", True):
            flags |= Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)

    def _apply_palette(self):
        """不透明窗口：把 Window 角色设为主题深色，避免任何自动填充露出灰底。"""
        pal = self.palette()
        pal.setColor(QPalette.Window, QColor(INK_1))
        pal.setColor(QPalette.Base, QColor(INK_0))
        pal.setColor(QPalette.Text, QColor(TEXT_DIM))
        self.setPalette(pal)
        self.setFont(self.font())  # 触发一次刷新，保证 QSS 字体生效

    def reload_flags(self):
        was_visible = self.isVisible()
        # setWindowFlags 可能重建底层窗口（新 HWND）：先注销旧窗口的热键，
        # 再在新 HWND 上重新注册，否则 Ctrl+Alt+Space 失效
        self._unregister_hotkey()
        self.set_flags()
        self._apply_palette()
        if was_visible:
            self.show()
            # show() 不保证重发 showEvent：flags 未变时 Qt 对 setWindowFlags
            # 直接 early-return，已可见的窗口不隐藏不重建，随后的 show() 是
            # 空操作、不重发 showEvent（PySide6 实测 showEvent 计数保持 1）。
            # 此时「showEvent 里重注册」的契约静默失效——热键已被注销却没人
            # 注册回来，必须在这里显式补注册，不能只依赖 showEvent
            if not self._hotkey:
                self._register_hotkey()
        else:
            # 隐藏态下设完 flags 不走 showEvent，但窗口句柄已重建，
            # 必须立刻在新 HWND 上注册热键，否则最小化到托盘期间失联
            self._register_hotkey()

    def _place_initial(self):
        geo = self.settings.get("geometry")
        size = EXPANDED_SIZE if self._expanded else COMPACT_SIZE
        if isinstance(geo, (list, tuple)) and len(geo) == 2:
            rect = QRect(geo[0], geo[1], size.width(), size.height())
            self.setGeometry(self._clamp_to_screen(rect))
        else:
            screen = QApplication.primaryScreen()
            ag = screen.availableGeometry() if screen else QRect(0, 0, 800, 600)
            x = ag.right() - size.width() - 18
            y = ag.bottom() - size.height() - 14
            self.setGeometry(x, y, size.width(), size.height())

    def _clamp_to_screen(self, rect):
        screen = QApplication.screenAt(rect.center()) or QApplication.primaryScreen()
        if screen is None:
            return rect
        ag = screen.availableGeometry()
        if rect.width() > ag.width():
            rect.setWidth(ag.width())
        if rect.height() > ag.height():
            rect.setHeight(ag.height())
        if rect.left() < ag.left():
            rect.moveLeft(ag.left())
        if rect.right() > ag.right():
            rect.moveRight(ag.right())
        if rect.top() < ag.top():
            rect.moveTop(ag.top())
        if rect.bottom() > ag.bottom():
            rect.moveBottom(ag.bottom())
        return rect

    # ---------- 拖动与贴边吸附 ----------

    def _rounded_path(self, rect):
        path = QPainterPath()
        path.addRoundedRect(rect, CORNER_RADIUS, CORNER_RADIUS)
        return path

    def _stop_anim(self):
        """Settle a running expand/collapse animation onto its target.

        Dragging and snapping used to yield entirely to the animation,
        which swallowed any drag started within the 190ms collapse
        window; settle it immediately and follow the cursor instead.
        """
        anim = self._anim
        if anim is not None:
            if anim.state() == QPropertyAnimation.Running:
                target = anim.endValue()
                anim.stop()
                if isinstance(target, QRect):
                    self.setGeometry(target)
            self._anim = None

    def drag_to(self, global_pos, offset):
        """拖动中：跟随光标；离屏幕边缘很近时磁性吸附到边缘。"""
        rect = QRect(global_pos - offset, self.size())
        rect = self._clamp_to_screen(rect)
        screen = QApplication.screenAt(rect.center()) or QApplication.primaryScreen()
        if screen is not None:
            ag = screen.availableGeometry()
            x, y = rect.x(), rect.y()
            if abs(x - ag.left()) <= SNAP_DRAG_MARGIN:
                x = ag.left() + SNAP_PAD
            elif abs(rect.right() - ag.right()) <= SNAP_DRAG_MARGIN:
                x = ag.right() - rect.width() - SNAP_PAD
            if abs(y - ag.top()) <= SNAP_DRAG_MARGIN:
                y = ag.top() + SNAP_PAD
            elif abs(rect.bottom() - ag.bottom()) <= SNAP_DRAG_MARGIN:
                y = ag.bottom() - rect.height() - SNAP_PAD
            rect.moveTo(x, y)
            rect = self._clamp_to_screen(rect)
        # settle the animation first so the drag is never swallowed
        self._stop_anim()
        self.move(rect.topLeft())

    def snap_to_edge(self):
        """松手时若离屏幕边缘很近就吸附，并记住位置。"""
        rect = self.geometry()
        debug.log("snap_to_edge", pos=[rect.x(), rect.y()])
        screen = QApplication.screenAt(rect.center()) or QApplication.primaryScreen()
        if screen is None:
            return
        ag = screen.availableGeometry()
        x, y = rect.x(), rect.y()
        if abs(x - ag.left()) <= SNAP_MARGIN:
            x = ag.left() + SNAP_PAD
        elif abs(rect.right() - ag.right()) <= SNAP_MARGIN:
            x = ag.right() - rect.width() - SNAP_PAD
        if abs(y - ag.top()) <= SNAP_MARGIN:
            y = ag.top() + SNAP_PAD
        elif abs(rect.bottom() - ag.bottom()) <= SNAP_MARGIN:
            y = ag.bottom() - rect.height() - SNAP_PAD
        if x != rect.x() or y != rect.y():
            rect.moveTo(x, y)
            rect = self._clamp_to_screen(rect)
            self._stop_anim()
            self.move(rect.topLeft())
        self._remember_geometry()

    def _remember_geometry(self):
        self.settings.set("geometry", [self.x(), self.y()])

    def toggle_expand(self):
        debug.log("toggle_expand", to=not self._expanded)
        self._apply_expanded(not self._expanded)

    def set_expanded(self, val):
        self._apply_expanded(val)

    def _apply_expanded(self, val, instant=False):
        if val == self._expanded and not instant:
            return
        self._expanded = val
        cur = self.geometry()
        target_size = EXPANDED_SIZE if val else COMPACT_SIZE
        target = QRect(cur.topLeft(), target_size)
        target = self._clamp_to_screen(target)

        if val:
            self.input_p.setText(self.input_c.text())
        else:
            self.input_c.setText(self.input_p.toPlainText())

        self._compact.setVisible(not val)
        self._panel.setVisible(val)

        if instant:
            self.setGeometry(target)
            self._apply_window_shape()
        else:
            anim = QPropertyAnimation(self, b"geometry", self)
            anim.setDuration(190)
            anim.setEasingCurve(QEasingCurve.OutCubic)
            anim.setStartValue(cur)
            anim.setEndValue(target)
            anim.start()
            self._anim = anim
            anim.finished.connect(lambda: setattr(self, '_anim', None))

        if val:
            self.input_p.setFocus()
        else:
            self.input_c.setFocus()
        if not instant:
            # 构造期的 instant 调用只是恢复上次状态，不能把「默认 (0,0)」
            # 当成用户位置记住——那会覆盖 settings 里保存的窗口位置。
            self._remember_geometry()
            self.settings.set("expanded", val)
            self.settings.save()

    def summon(self):
        """从托盘召回：显示并闪一下。"""
        debug.log("summon", pos=[self.x(), self.y()], expanded=self._expanded)
        # 召回前校验几何：隐藏期间外接显示器拔掉后，Windows 不会迁移
        # 隐藏窗口，直接 show 会落在失效坐标上（窗口不可见且无法找回）
        self.setGeometry(self._clamp_to_screen(self.geometry()))
        if not self.isVisible():
            self.show()
        self.raise_()
        self.activateWindow()
        if self._expanded:
            self.input_p.setFocus()
        else:
            self.input_c.setFocus()

    # ---------- 界面搭建 ----------

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(11, 9, 11, 9)
        root.setSpacing(0)

        self._build_compact()
        self._build_panel()
        root.addWidget(self._compact)
        root.addWidget(self._panel)
        self._panel.setVisible(False)

    def _build_compact(self):
        self._compact = QWidget(self)
        lay = QHBoxLayout(self._compact)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        self.avatar_c = Avatar(40, self._compact)
        self.avatar_c.clicked.connect(self.toggle_expand)
        self.avatar_c.drag_window = self
        self.input_c = QLineEdit(self._compact)
        self.input_c.setPlaceholderText("写点什么，或告诉我如何排版…")
        self.input_c.setMaxLength(self.MAX_PROMPT_CHARS)
        self.input_c.returnPressed.connect(self._on_send_compact)
        self.btn_send_c = QPushButton(icon_send(), "", self._compact)
        self.btn_send_c.setObjectName("primary")
        self.btn_send_c.setFixedSize(38, 38)
        self.btn_send_c.clicked.connect(self._on_send_compact)
        self.btn_expand = QPushButton(icon_expand(), "", self._compact)
        self.btn_expand.setObjectName("tool")
        self.btn_expand.setFixedSize(32, 38)
        self.btn_expand.clicked.connect(lambda: self._apply_expanded(True))
        lay.addWidget(self.avatar_c)
        lay.addWidget(self.input_c, 1)
        lay.addWidget(self.btn_send_c)
        lay.addWidget(self.btn_expand)

    def _build_panel(self):
        self._panel = QWidget(self)
        lay = QVBoxLayout(self._panel)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        # 标题栏：头像 + 名称/状态 + 拖动条 + 工具按钮
        header = QWidget(self._panel)
        hlay = QHBoxLayout(header)
        hlay.setContentsMargins(2, 2, 2, 2)
        hlay.setSpacing(8)
        self.avatar_p = Avatar(34, header)
        self.avatar_p.clicked.connect(self.toggle_expand)
        self.avatar_p.drag_window = self
        title_box = QWidget(header)
        tlay = QVBoxLayout(title_box)
        tlay.setContentsMargins(0, 0, 0, 0)
        tlay.setSpacing(0)
        title = QLabel("AI4Word", title_box)
        title.setObjectName("title")
        self.word_label = QLabel("未连接 Word", title_box)
        self.word_label.setObjectName("wordStatus")
        tlay.addWidget(title)
        tlay.addWidget(self.word_label)
        strip = DragStrip(header)
        strip.setFixedHeight(34)
        self.btn_blockmap = QPushButton(icon_list(), "", header)
        self.btn_blockmap.setObjectName("tool")
        self.btn_blockmap.setCheckable(True)
        self.btn_blockmap.setFixedSize(32, 32)
        self.btn_blockmap.setToolTip("文档块地图")
        self.btn_blockmap.clicked.connect(self._toggle_blockmap)
        self.btn_settings = QPushButton(icon_gear(), "", header)
        self.btn_settings.setObjectName("tool")
        self.btn_settings.setFixedSize(32, 32)
        self.btn_settings.setToolTip("设置")
        self.btn_settings.clicked.connect(self._open_settings)
        self.btn_collapse = QPushButton(icon_collapse(), "", header)
        self.btn_collapse.setObjectName("tool")
        self.btn_collapse.setFixedSize(32, 32)
        self.btn_collapse.setToolTip("收起")
        self.btn_collapse.clicked.connect(lambda: self._apply_expanded(False))
        hlay.addWidget(self.avatar_p)
        hlay.addWidget(title_box)
        hlay.addWidget(strip, 1)
        hlay.addWidget(self.btn_blockmap)
        hlay.addWidget(self.btn_settings)
        hlay.addWidget(self.btn_collapse)

        # 实时状态条：生成 / 排版进度
        self._status_label = QLabel("", self._panel)
        self._status_label.setObjectName("statusBar")
        self._status_label.setAlignment(Qt.AlignCenter)
        self._status_label.setFixedHeight(18)
        self._status_label.setVisible(False)

        # 消息流 + 块地图侧栏
        splitter = QSplitter(Qt.Horizontal, self._panel)
        self.messages = MessageList(splitter)
        self.block_panel = BlockMapPanel(splitter)
        self.block_panel.setFixedWidth(212)
        self.block_panel.blockClicked.connect(self._on_block_clicked)
        splitter.addWidget(self.messages)
        splitter.addWidget(self.block_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        splitter.setHandleWidth(6)
        self.block_panel.setVisible(False)

        # 中断选择条
        self._choice_bar = QWidget(self._panel)
        clay = QHBoxLayout(self._choice_bar)
        clay.setContentsMargins(4, 0, 4, 0)
        clay.setSpacing(6)
        choice_label = QLabel("生成已中断", self._choice_bar)
        choice_label.setStyleSheet(f"color: {AMBER}; font-weight: 600;")
        btn_rollback = QPushButton(icon_rollback(), " 回滚", self._choice_bar)
        btn_keep = QPushButton(icon_keep(), " 保留", self._choice_bar)
        btn_extra = QPushButton("追加补充…", self._choice_bar)
        for b in (btn_rollback, btn_keep, btn_extra):
            b.setObjectName("tool")
        btn_rollback.clicked.connect(lambda: self._choose("rollback"))
        btn_keep.clicked.connect(lambda: self._choose("keep"))
        btn_extra.clicked.connect(self._enter_extra_mode)
        clay.addWidget(choice_label)
        clay.addStretch(1)
        clay.addWidget(btn_rollback)
        clay.addWidget(btn_keep)
        clay.addWidget(btn_extra)
        self._choice_bar.setVisible(False)

        # 工具条
        toolbar = QWidget(self._panel)
        blay = QHBoxLayout(toolbar)
        blay.setContentsMargins(2, 0, 2, 0)
        blay.setSpacing(5)

        self._speed_group = QButtonGroup(toolbar)
        for i, name in enumerate(("慢", "自", "快")):
            b = QPushButton(name, toolbar)
            b.setObjectName("tool")
            b.setCheckable(True)
            b.setFixedHeight(26)
            b.setFixedWidth(32)
            self._speed_group.addButton(b, i)
            blay.addWidget(b)
        self._speed_group.idClicked.connect(self._on_speed)
        cur_speed = self.settings.get("speed") or "auto"
        sid = {"slow": 0, "auto": 1, "fast": 2}.get(cur_speed, 1)
        self._speed_group.button(sid).setChecked(True)

        sep1 = self._vsep(toolbar)
        blay.addWidget(sep1)

        self._mode_group = QButtonGroup(toolbar)
        for i, name in enumerate(("写作", "排版")):
            b = QPushButton(name, toolbar)
            b.setObjectName("tool")
            b.setCheckable(True)
            b.setFixedHeight(26)
            b.setFixedWidth(44)
            self._mode_group.addButton(b, i)
            blay.addWidget(b)
        self._mode_group.idClicked.connect(self._on_mode)
        self._mode_group.button(0).setChecked(True)

        sep2 = self._vsep(toolbar)
        blay.addWidget(sep2)
        self.btn_review = QPushButton("修订", toolbar)
        self.btn_review.setObjectName("tool")
        self.btn_review.setCheckable(True)
        self.btn_review.setFixedHeight(26)
        self.btn_review.setFixedWidth(46)
        self.btn_review.toggled.connect(self._on_review)
        blay.addWidget(self.btn_review)
        self.preset_combo = QComboBox(toolbar)
        self.preset_combo.addItems(["套用预设…"] + list(preset_list()))
        self.preset_combo.setFixedHeight(26)
        self.preset_combo.setMaximumWidth(110)
        # 只按最短内容算闭合宽度（下拉列表本身仍可显示完整项名）
        self.preset_combo.setMinimumContentsLength(4)
        self.preset_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.preset_combo.setToolTip("一键套用成套排版样式")
        # 绑定 currentIndexChanged 而非 activated：activated 只在用户真实弹窗
        # 选择时触发，程序化 setCurrentIndex 改选永不触发，_on_preset 会被
        # 静默跳过（组合框不复位、preset 命令不投递）。currentIndexChanged
        # 覆盖程序化改选与键盘导航；本处复位回充 setCurrentIndex(0) 会再次
        # 触发，但被 _on_preset 的 idx<=0 守卫挡住，安全；且不与 activated
        # 并存，避免真实弹窗/键盘选择时两个信号各触发一次、双次投递命令。
        self.preset_combo.currentIndexChanged.connect(self._on_preset)
        blay.addWidget(self.preset_combo)
        blay.addStretch(1)
        self.btn_save = QPushButton(icon_save(), " 存档", toolbar)
        self.btn_save.setObjectName("tool")
        self.btn_save.setFixedHeight(26)
        self.btn_save.setFixedWidth(64)
        self.btn_save.clicked.connect(self._on_save)
        blay.addWidget(self.btn_save)

        # 输入行
        input_row = QWidget(self._panel)
        iray = QHBoxLayout(input_row)
        iray.setContentsMargins(0, 0, 0, 0)
        iray.setSpacing(8)
        self.input_p = InputEdit(self._on_send_panel, input_row)
        self.btn_send_p = QPushButton(icon_send(), "", input_row)
        self.btn_send_p.setObjectName("primary")
        self.btn_send_p.setFixedSize(44, 56)
        self.btn_send_p.clicked.connect(self._on_send_panel)
        iray.addWidget(self.input_p, 1)
        iray.addWidget(self.btn_send_p)

        lay.addWidget(header)
        lay.addWidget(self._status_label)
        lay.addWidget(splitter, 1)
        lay.addWidget(self._choice_bar)
        lay.addWidget(toolbar)
        lay.addWidget(input_row)

    def _vsep(self, parent):
        w = QWidget(parent)
        w.setFixedWidth(1)
        w.setStyleSheet(f"background: {INK_4};")
        return w

    # ---------- 引擎信号接线 ----------

    def _wire_worker(self):
        w = self.worker
        w.stateChanged.connect(self._on_state)
        w.wordStatus.connect(self._on_word_status)
        w.message.connect(self._on_message)
        w.progress.connect(self._on_progress)
        w.streamStart.connect(self.messages.begin_stream)
        w.streamChunk.connect(self.messages.stream_append)
        w.streamEnd.connect(self.messages.end_stream)
        w.blockMap.connect(self.block_panel.set_text)
        w.interrupted.connect(self._on_interrupted)

    def _on_state(self, state):
        debug.log("ui_state", state=state)
        busy = state in ("writing", "arranging", "connecting")
        self._busy = busy
        avatar_state = "working" if busy else "idle"
        self.avatar_c.set_state(avatar_state)
        self.avatar_p.set_state(avatar_state)
        self.btn_send_c.setIcon(icon_stop() if busy else icon_send())
        self.btn_send_p.setIcon(icon_stop() if busy else icon_send())
        self.btn_send_c.setToolTip("中断" if busy else "发送")
        self.btn_send_p.setToolTip("中断" if busy else "发送")
        if not busy:
            self._choice_bar.setVisible(False)
            self._awaiting_choice = False
            self._extra_mode = False
            self._set_input_placeholder()
            self._status_label.setVisible(False)
            self._status_label.setText("")

    def _on_word_status(self, text):
        self.word_label.setText(text)

    def _on_progress(self, text):
        """生成/排版进度：状态条实时显示（流式气泡同时展示原文）。"""
        self._status_label.setText(str(text))
        self._status_label.setVisible(bool(str(text)))

    def _on_message(self, kind, text):
        self.messages.add(kind, text)
        if kind == "error":
            self.avatar_c.set_state("error")
            self.avatar_p.set_state("error")
            from PySide6.QtCore import QTimer
            QTimer.singleShot(1800, self._restore_avatar_state)

    def _restore_avatar_state(self):
        if not self._busy:
            self.avatar_c.set_state("idle")
            self.avatar_p.set_state("idle")

    # ---------- 发送 / 中断 / 选择 ----------

    def _set_input_placeholder(self):
        if self._extra_mode:
            ph = "补充内容描述，Enter 追加…"
        elif self._write_mode:
            ph = "写点什么… 例如：写一篇关于秋天的散文"
        else:
            ph = "如何调整？例如：把所有一级标题居中并加粗"
        self.input_c.setPlaceholderText(ph)
        self.input_p.setPlaceholderText(ph)

    def _on_mode(self, idx):
        self._write_mode = (idx == 0)
        debug.log("mode_clicked", mode="write" if self._write_mode else "arrange")
        self._set_input_placeholder()

    def _on_review(self, on):
        """修订开关：生成中拒绝并回退按钮状态。

        放进命令队列的话会在本次生成结束后才执行、且执行时已不 busy——
        用户几分钟前点了一下、此时突然「已开启修订模式」毫无道理。
        """
        if self._busy:
            debug.log("review_rejected_busy", on=on)
            self.messages.add("info", "生成中不能切换修订模式，请先中断当前生成。")
            self.btn_review.blockSignals(True)
            self.btn_review.setChecked(not on)
            self.btn_review.blockSignals(False)
            return
        debug.log("review_toggled", on=bool(on))
        self.worker.send("review", bool(on))

    def _on_save(self):
        """存档按钮：生成中给出的只是中间快照且会被延迟，不如明确拒绝。"""
        if self._busy:
            self.messages.add("info", "正在生成，请先中断再存档。")
            debug.log("save_rejected_busy")
            return
        debug.log("save_clicked")
        self.worker.send("save")

    def _on_speed(self, idx):
        mode = ("slow", "auto", "fast")[idx]
        debug.log("speed_clicked", mode=mode)
        self.settings.set("speed", mode)
        self.settings.save()
        # 直接设置而非走命令队列：worker 处理 write 期间不消费队列，
        # 排队等生成结束才执行就完全没有实时效果
        self.worker.set_speed(mode)

    def _on_preset(self, idx):
        if idx <= 0:
            return
        if self._busy:
            # 生成中套预设会和写入器的样式设置互相打架；不放队列延迟执行
            # （延迟执行会让用户在 minutes 后看到莫名其妙的「已应用」）
            self.messages.add("info", "生成中不能套用预设，请先中断当前生成。")
            debug.log("preset_rejected_busy")
            self.preset_combo.setCurrentIndex(0)
            return
        name = self.preset_combo.itemText(idx)
        debug.log("preset_selected", name=name)
        self.worker.send("preset", name)
        self.preset_combo.setCurrentIndex(0)

    def _toggle_blockmap(self):
        on = self.btn_blockmap.isChecked()
        debug.log("blockmap_toggled", on=on)
        self.block_panel.setVisible(on)
        if on:
            self.worker.send("refresh_map")

    def _on_block_clicked(self, index):
        debug.log("block_clicked", index=index)
        self.worker.send("select_block", index)

    def _on_send_compact(self):
        debug.log("send_compact", busy=self._busy, extra=self._extra_mode)
        if self._busy and not self._extra_mode:
            self.worker.interrupt()
            return
        text = self.input_c.text().strip()
        self._send(text)

    def _on_send_panel(self):
        debug.log("send_panel", busy=self._busy, extra=self._extra_mode)
        if self._busy and not self._extra_mode:
            self.worker.interrupt()
            return
        text = self.input_p.toPlainText().strip()
        self._send(text)

    MAX_PROMPT_CHARS = 20000

    def _send(self, text):
        debug.log("send", chars=len(text),
                  mode="write" if self._write_mode else "arrange")
        if not text:
            return
        if len(text) > self.MAX_PROMPT_CHARS:
            debug.log("send_rejected_too_long", chars=len(text))
            self.messages.add("info", f"输入超过 {self.MAX_PROMPT_CHARS} 字上限，"
                              "请缩减后再发送。")
            return
        if self._awaiting_choice and not self._extra_mode:
            debug.log("send_rejected_awaiting_choice")
            if self._expanded:
                self.messages.add("info", "请先选择 回滚 / 保留 / 追加补充。")
            else:
                # 紧凑态看不到选择条：引导用户展开
                self.messages.add("info", "请先展开悬浮窗，选择 回滚 / 保留 / 追加补充。")
            return
        if self._extra_mode:
            self._extra_mode = False
            self._awaiting_choice = False
            self._choice_bar.setVisible(False)
            self._set_input_placeholder()
            debug.log("extra_submitted", chars=len(text))
            self.worker.choose(("extra", text))
        elif self._write_mode:
            debug.log("write_sent", chars=len(text), prompt=text, full=True)
            self.worker.send("write", text)
        else:
            debug.log("arrange_sent", chars=len(text), prompt=text, full=True)
            self.worker.send("arrange", text)
        # 乐观置 busy：stateChanged 信号跨线程回来有几十毫秒延迟，
        # 这期间第二次 Enter 会被 worker 拒收「正在处理」但输入已被清空
        # （文本丢失）；本地先锁住，第二条 Enter 就走「中断」逻辑且文本保留
        self._busy = True
        self.input_c.clear()
        self.input_p.clear()

    def _on_interrupted(self):
        debug.log("interrupted_ui")
        self._awaiting_choice = True
        self._extra_mode = False
        self._choice_bar.setVisible(True)
        self._set_input_placeholder()
        # 选项条只在展开面板里有：中断后自动展开，确保用户看得见选择入口
        if not self._expanded:
            self._apply_expanded(True)

    def _choose(self, value):
        debug.log("choose", value=str(value)[:50])
        self._awaiting_choice = False
        self._extra_mode = False
        self._choice_bar.setVisible(False)
        self._set_input_placeholder()
        self.worker.choose(value)

    def _enter_extra_mode(self):
        self._extra_mode = True
        debug.log("enter_extra_mode")
        # 不隐藏选择条：追加模式是「还能改主意」的状态，用户随时可以
        # 放弃输入直接点 回滚/保留（此前隐藏选择条 = 唯一退路是干等 5 分钟）
        self._set_input_placeholder()
        self.input_p.setFocus()

    def _exit_extra_mode(self):
        """追加模式打消：回到选择条等待。Esc 优先走这里而不是直接收起窗口。"""
        if self._extra_mode:
            self._extra_mode = False
            debug.log("exit_extra_mode")
            self._set_input_placeholder()

    # ---------- 窗口行为 ----------

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        rect = self.rect()
        grad = QLinearGradient(0, 0, 0, self.height())
        grad.setColorAt(0.0, QColor("#22242d"))
        grad.setColorAt(1.0, QColor("#16181f"))
        p.setBrush(grad)
        p.setPen(Qt.NoPen)
        p.drawRect(rect)  # 不透明：整幅铺满；圆角形状由系统/DWM 切出
        p.setBrush(Qt.NoBrush)
        border = AMBER if self._busy else INK_4
        p.setPen(QPen(QColor(border), 1.3))
        p.drawPath(self._rounded_path(rect.adjusted(1, 1, -1, -1)))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 无需重设形状：DWM 圆角跟窗口尺寸走；Win10 降级的 mask 跟着重建
        self._apply_window_shape()

    def _apply_window_shape(self):
        """窗口形状策略：

        Win11（build ≥ 22000）：交给 DWM 合成器画原生抗锯齿圆角，
        清掉 mask。QRegion mask 在 Win11 上会覆盖原生圆角，边缘是
        像素级锯齿——V9.0 的锯齿 complaints 就是它造成的。
        老系统：DwmSetWindowAttribute 不可用，降级 QRegion mask
        保住圆角形状（有锯齿但形状正确）。
        """
        hwnd = int(self.winId())
        if hwnd != 0 and self._try_dwm_round(hwnd):
            if not self.mask().isEmpty():
                self.clearMask()
            self._shape_winid = hwnd
            return
        r = self._rounded_path(self.rect().adjusted(1, 1, -1, -1))
        self.setMask(QRegion(r.toFillPolygon().toPolygon()))
        self._shape_winid = hwnd

    def _try_dwm_round(self, hwnd):
        """在 Win11 上把窗口圆角偏好设为 DWMWCP_ROUND；不支持则返回 False。"""
        try:
            build = sys.getwindowsversion().build
        except Exception:
            build = 0
        if build < _DWM_TRY_BUILD:
            return False
        try:
            pref = ctypes.c_int(_DWMWCP_ROUND)
            ok = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                ctypes.c_void_p(hwnd), _DWMWA_WINDOW_CORNER_PREFERENCE,
                ctypes.byref(pref), ctypes.sizeof(pref))
            return ok == 0  # S_OK
        except Exception:
            return False

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            if self._extra_mode:
                # 优先退出追加模式、回到选择条（否则它是死路：
                # 只能发字或干等 worker 5 分钟超时）
                self._exit_extra_mode()
            elif self._expanded:
                self._apply_expanded(False)
            else:
                debug.log("hide_via_esc")
                self.hide()
                if self._tray:
                    self._tray.first_hide_hint()
            return
        super().keyPressEvent(event)

    def closeEvent(self, event):
        if self._tray is not None:
            debug.log("close_to_tray")
            event.ignore()
            self.hide()
            self._tray.first_hide_hint()
            return
        super().closeEvent(event)

    def set_tray(self, tray):
        self._tray = tray

    def _open_settings(self):
        # settings_open / settings_closed 由 SettingsDialog 自己打（__init__
        # 与 done()）：本包装路径与「直接构造对话框」的调用方共用同一份接线，
        # 两个事件都恰好一次，不再各打一遍。
        from app.settings_dialog import SettingsDialog
        dlg = SettingsDialog(self.settings, self.worker, parent=self)
        dlg.exec()
        self.reload_flags()

    def open_settings(self):
        self._open_settings()

    # ---------- 全局热键 Ctrl+Alt+Space 召唤 ----------

    def nativeEvent(self, eventType, message):
        if eventType == b"windows_generic_MSG" and message:
            try:
                msg = wt.MSG.from_address(int(message))
                if msg.message == 0x0312:  # WM_HOTKEY
                    self.summon()
                    return True, 0
            except Exception:
                pass
        return super().nativeEvent(eventType, message)

    def _register_hotkey(self):
        try:
            user32 = ctypes.windll.user32
            ok = user32.RegisterHotKey(int(self.winId()), 1,
                                       0x0002 | 0x0001, 0x20)  # Ctrl+Alt, Space
            self._hotkey = bool(ok)
            debug.log("hotkey_register", ok=bool(ok))
        except Exception:
            self._hotkey = False
            debug.warn("hotkey_register_failed")
        if not self._hotkey and not getattr(self, "_hotkey_hinted", False):
            # 热键被别的程序占用（放大镜、其它工具）时不能静默：
            # 用户按 Ctrl+Alt+Space 毫无反应却不知道为什么
            self._hotkey_hinted = True
            try:
                self.messages.add("info", "Ctrl+Alt+Space 召唤热键注册失败："
                                  "可能被其它程序占用。可双击托盘图标召回。")
            except Exception:
                pass

    def _unregister_hotkey(self):
        if self._hotkey:
            try:
                ctypes.windll.user32.UnregisterHotKey(int(self.winId()), 1)
            except Exception:
                pass
            self._hotkey = False
            debug.log("hotkey_unregistered")

    def showEvent(self, event):
        super().showEvent(event)
        if not self._hotkey:
            self._register_hotkey()
        # HWND 可能重建（setWindowFlags / 首次 show），形状要重新应用
        hwnd = int(self.winId())
        if hwnd != 0 and hwnd != getattr(self, "_shape_winid", 0):
            self._apply_window_shape()

    def hideEvent(self, event):
        super().hideEvent(event)

    def __del__(self):  # noqa
        try:
            self._unregister_hotkey()
        except Exception:
            pass
