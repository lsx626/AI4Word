# -*- coding: utf-8 -*-
"""消息流与块地图：聊天气泡（流式增量追加）+ 可点击的块索引面板。

气泡高度用 QTextDocument 按 label 实际宽度确定性地计算——不能依赖
QLabel 的 heightForWidth：流式 setText 时它常被布局按错误的宽度求解，
导致多行消息只剩一行高（实时进度被截断）。
"""
import html
import re

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextDocument
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


class _WrapLabel(QLabel):
    """自动换行富文本标签：高度由实际宽度 + 文本确定，绝不截断。

    QLineEdit 式的实时追加（stream_append）依赖每次 setText 及宽度变化
    时重新计算 minimumHeight；QLabel 自带的 heightForWidth 在流式场景
    下求解宽度不可靠（常被布局以视口宽度求解），因此这里自行用
    QTextDocument 量高，布局以 minimumHeight 为准。
    """

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setWordWrap(True)
        self.setTextFormat(Qt.RichText)

    def _relayout_height(self):
        w = self.width()
        if w <= 0:
            # 尚未布局：按 sizeHint 的目标宽度先量一次（够准也避免 0 宽）
            w = max(1, min(self.sizeHint().width(), 400))
        try:
            doc = QTextDocument()
            doc.setDefaultFont(self.font())
            doc.setHtml(self.text())
            doc.setTextWidth(w)
            h = doc.size().height()
        except Exception:
            h = self.sizeHint().height()
        # +1 抹平取整误差，避免最后一行半像素被裁
        self.setMinimumHeight(max(1, int(h) + 1))

    def setText(self, text):
        super().setText(text)
        self._relayout_height()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._relayout_height()


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
        self._label = _WrapLabel(self._render(text), self)
        self._label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        if kind == "code":
            self._label.setStyleSheet(f"font-family: {CODE_FONT}; font-size: 12px;")
        elif kind == "info":
            self._label.setStyleSheet("font-size: 12px;")
        lay.addWidget(self._label)
        if kind in ("info", "code"):
            self.setContentsMargins(0, 0, 0, 0)

    def _render(self, text):
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
        # 全宽气泡（Slack 式）：流式长文本一行能容纳更多字，进度最清晰；
        # 发送方用颜色区分（用户=琥珀、助手=墨灰），不再依赖左右对齐
        self._layout.insertWidget(self._layout.count() - 1, b)
        self._trim()
        self._scroll_bottom()
        return b

    def _trim(self, keep=200):
        """长会话防爆：只保留最近 keep 条气泡，老的移除释放。"""
        total = self._layout.count() - 1  # 末尾是 stretch
        while total > keep:
            it = self._layout.takeAt(0)
            if it is None:
                break
            w = it.widget()
            if w is not None:
                w.deleteLater()
            total -= 1

    def begin_stream(self):
        self._streaming = self.add("assistant", "")
        return self._streaming

    def stream_append(self, piece):
        if self._streaming is not None:
            self._streaming.append_text(piece)
            # 用户上滚阅读历史时不要把视图拽回底部（每个 chunk 都拽一次
            # 根本没法看）；生成结束（end_stream）再回到底部
            if self._at_bottom():
                self._scroll_bottom()

    def end_stream(self):
        if self._streaming is not None:
            self._streaming.finalize()
            self._streaming = None
            self._scroll_bottom()

    def _at_bottom(self):
        """用户是否停在消息流底部（允许 16px 的抖动余量）。"""
        bar = self.verticalScrollBar()
        return bar.maximum() - bar.value() <= 16

    def _scroll_bottom(self):
        """滚到底部：立即滚一次（气泡高度同步可知时），布局生效后再滚一次。

        QLabel 的 minimumHeight 是同步设的，但布局重算在事件循环里完成，
        滚动条 maximum 同步取值时可能还是旧值——单拍一次 0ms 延迟兜底。
        """
        bar = self.verticalScrollBar()
        bar.setValue(bar.maximum())
        from PySide6.QtCore import QTimer
        QTimer.singleShot(0, lambda: bar.setValue(bar.maximum()))


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
