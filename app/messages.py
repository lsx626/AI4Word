# -*- coding: utf-8 -*-
"""消息流与块地图：聊天气泡（流式增量追加）+ 可点击的块索引面板。"""
import re

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QFrame, QLabel, QListWidget, QListWidgetItem,
                                QScrollArea, QVBoxLayout, QWidget)

from app.theme import CODE_FONT, INK_2, INK_3, RED, TEXT, TEXT_DIM, TEXT_FAINT

BUBBLE_QSS = f"""
QFrame#bubble_user {{
    background: rgba(242, 169, 59, 34);
    border: 1px solid rgba(242, 169, 59, 88);
    border-radius: 12px;
}}
QFrame#bubble_assistant {{
    background: {INK_2};
    border: 1px solid {INK_3};
    border-radius: 12px;
}}
QFrame#bubble_info {{ background: transparent; border: none; }}
QFrame#bubble_error {{
    background: rgba(224, 101, 90, 26);
    border: 1px solid rgba(224, 101, 90, 100);
    border-radius: 12px;
}}
QFrame#bubble_code {{
    background: #14151a;
    border: 1px solid {INK_3};
    border-radius: 10px;
}}
"""

_KIND_COLOR = {
    "user": "#ffd89a",
    "assistant": TEXT,
    "info": TEXT_DIM,
    "code": "#9ec3ff",
    "error": "#f0907f",
}


class MessageBubble(QFrame):
    """一条消息气泡。流式气泡用 append_text 增量更新（转义 + 换行渲染）。"""

    def __init__(self, kind, text="", parent=None):
        super().__init__(parent)
        self.kind = kind
        self.setObjectName(f"bubble_{kind}")
        self.setStyleSheet(BUBBLE_QSS)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 7, 10, 7)
        lay.setSpacing(0)
        self._raw = text
        self._label = QLabel(self._render(text), self)
        self._label.setWordWrap(True)
        self._label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self._label.setTextFormat(Qt.RichText)
        if kind == "code":
            self._label.setStyleSheet(f"font-family: {CODE_FONT}; font-size: 12px;")
        elif kind == "info":
            self._label.setStyleSheet("font-size: 12px;")
        lay.addWidget(self._label)
        if kind in ("info", "code"):
            self.setContentsMargins(0, 0, 0, 0)
        self.setMaximumWidth(430)

    def _render(self, text):
        import html
        t = html.escape(text or "")
        t = t.replace("\n", "<br>")
        if self.kind == "code":
            t = t.replace(" ", "&nbsp;")
        return t

    def append_text(self, piece):
        self._raw += piece
        self._label.setText(self._render(self._raw))

    def finalize(self):
        if not (self._raw or "").strip():
            self._label.setText('<span style="color:#6b7078">（无内容）</span>')


class MessageList(QScrollArea):
    """纵向消息流：自动滚动到底部，支持流式气泡的生命周期。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        self._container = QWidget()
        self._container.setStyleSheet("background: transparent;")
        self._layout = QVBoxLayout(self._container)
        self._layout.setContentsMargins(2, 6, 2, 6)
        self._layout.setSpacing(6)
        self._layout.addStretch(1)
        self.setWidget(self._container)
        self._streaming = None

    def add(self, kind, text):
        b = MessageBubble(kind, text)
        align = Qt.AlignRight if kind == "user" else Qt.AlignLeft
        self._layout.insertWidget(self._layout.count() - 1, b, 0, align)
        self._scroll_bottom()
        return b

    def begin_stream(self):
        self._streaming = self.add("assistant", "")
        return self._streaming

    def stream_append(self, piece):
        if self._streaming is not None:
            self._streaming.append_text(piece)
            self._scroll_bottom()

    def end_stream(self):
        if self._streaming is not None:
            self._streaming.finalize()
            self._streaming = None
            self._scroll_bottom()

    def _scroll_bottom(self):
        bar = self.verticalScrollBar()
        bar.setValue(bar.maximum())


class BlockMapPanel(QWidget):
    """文档块地图：点击某行 → 在 Word 里滚动定位并高亮该块。"""

    blockClicked = Signal(int)

    _LINE_RE = re.compile(r"^\s*\[(\d+)\]\s*(.*)$")

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        hint = QLabel("块地图 · 点击定位", self)
        hint.setStyleSheet(f"color: {TEXT_FAINT}; font-size: 11px;")
        lay.addWidget(hint)
        self._list = QListWidget(self)
        self._list.setObjectName("blockMap")
        lay.addWidget(self._list)
        self._list.itemClicked.connect(self._on_click)
        self._indices = []

    def set_text(self, text):
        self._list.clear()
        self._indices = []
        for line in str(text or "").splitlines():
            m = self._LINE_RE.match(line)
            if m:
                self._indices.append(int(m.group(1)))
                QListWidgetItem(m.group(2).strip() or "（空块）", self._list)
        if not self._indices:
            QListWidgetItem("（空文档）", self._list)

    def _on_click(self, item):
        i = self._list.row(item)
        if 0 <= i < len(self._indices):
            self.blockClicked.emit(self._indices[i])
