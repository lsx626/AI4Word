# -*- coding: utf-8 -*-
"""桌宠头像：发光琥珀球 + 呼吸 + 旋转光环。

画法全部 QPainter 现绘，任意分辨率清晰；三种状态：
- idle：缓慢呼吸、光环慢转
- working：光环快转、光芒更亮
- error：核心转红、慢闪

draw_orb 同时供给 app 图标与托盘图标，保证品牌一致。
"""
import math

from PySide6.QtCore import QPoint, QRectF, Signal, Qt, QTimer, QVariantAnimation
from PySide6.QtGui import QColor, QPainter, QRadialGradient
from PySide6.QtWidgets import QWidget

from app.theme import AMBER, AMBER_DEEP, AMBER_SOFT, INK_0, RED


def draw_orb(painter, rect, state="idle", angle=0.0, breath=0.0):
    """在 rect（逻辑像素）里画一个发光球。供头像 / 图标 / 托盘复用。"""
    painter.setRenderHint(QPainter.Antialiasing, True)
    cx = rect.center().x()
    cy = rect.center().y()
    r = min(rect.width(), rect.height()) / 2.0 - 1.0
    if r <= 0:
        return

    error = state == "error"
    working = state == "working"
    core_light = QColor("#ffcf7d") if not error else QColor("#f0907f")
    core_dark = QColor(AMBER_DEEP) if not error else QColor("#9a352c")
    halo = QColor(AMBER_SOFT) if not error else QColor(RED)

    # 外发光
    glow = QRadialGradient(cx, cy, r * 1.05)
    a1 = 150 if working else (90 + 40 * math.sin(2 * math.pi * breath))
    if error:
        a1 = 70 + 50 * math.sin(2 * math.pi * breath * 2)
    halo.setAlpha(int(max(0, min(255, a1))))
    glow.setColorAt(0.0, halo)
    halo2 = QColor(halo)
    halo2.setAlpha(0)
    glow.setColorAt(1.0, halo2)
    painter.setPen(Qt.NoPen)
    painter.setBrush(glow)
    painter.drawEllipse(QRectF(cx - r * 1.6, cy - r * 1.6, r * 3.2, r * 3.2))

    # 光环：三段旋转弧
    ring_r = r * 0.92
    painter.translate(cx, cy)
    painter.rotate(angle)
    pen_w = max(1.6, r * 0.11)
    from PySide6.QtGui import QPen
    pen = QPen(halo, pen_w, Qt.SolidLine, Qt.RoundCap)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)
    rr = QRectF(-ring_r, -ring_r, ring_r * 2, ring_r * 2)
    span = 70 * 16
    for k in range(3):
        start = int(k * 120 * 16)
        painter.drawArc(rr, start, span)
    painter.resetTransform()

    # 核心：带高光的球体
    s = 1.0 + 0.05 * math.sin(2 * math.pi * breath)
    core_r = r * 0.60 * s
    grad = QRadialGradient(cx - core_r * 0.35, cy - core_r * 0.38, core_r * 1.5)
    grad.setColorAt(0.0, core_light)
    grad.setColorAt(1.0, core_dark)
    painter.setPen(Qt.NoPen)
    painter.setBrush(grad)
    painter.drawEllipse(QRectF(cx - core_r, cy - core_r, core_r * 2, core_r * 2))

    # 字符标记：W
    from PySide6.QtGui import QFont
    font = QFont("Georgia, Times New Roman, serif")
    font.setBold(True)
    font.setPixelSize(max(8, int(core_r * 0.95)))
    painter.setFont(font)
    painter.setPen(QColor("#2a200d") if not error else QColor("#2a0d08"))
    painter.drawText(QRectF(cx - core_r, cy - core_r, core_r * 2, core_r * 2),
                     Qt.AlignCenter, "W")


class Avatar(QWidget):
    """发光球桌宠头像：呼吸 + 旋转，随引擎状态变色变速。"""

    clicked = Signal()

    def __init__(self, diameter=38, parent=None):
        super().__init__(parent)
        self._diameter = diameter
        self.setFixedSize(diameter, diameter)
        self.setCursor(Qt.PointingHandCursor)
        self.drag_window = None   # 拖动时移动的目标窗口
        self._drag_offset = None
        self._dragged = False
        self._state = "idle"
        self._angle = 0.0
        self._breath = 0.0

        self._breath_anim = QVariantAnimation(self)
        self._breath_anim.setStartValue(0.0)
        self._breath_anim.setEndValue(1.0)
        self._breath_anim.setDuration(2400)
        self._breath_anim.setLoopCount(-1)
        self._breath_anim.valueChanged.connect(self._on_breath)
        self._breath_anim.start()

        self._timer = QTimer(self)
        self._timer.setInterval(33)  # ~30fps
        self._timer.timeout.connect(self._tick)
        self._timer.start()

    def _on_breath(self, v):
        self._breath = float(v)

    def _tick(self):
        # 度/秒：working 快转、idle 慢转、error 慢闪
        win = self.window()
        if win is not None and not win.isVisible():
            return  # 窗口隐藏（最小化到托盘）时不必空转重绘
        speed = {"idle": 22.0, "working": 170.0, "error": 40.0}[self._state]
        self._angle = (self._angle + speed * 0.033) % 360.0
        self.update()

    def set_state(self, state):
        if state in ("idle", "working", "error") and state != self._state:
            self._state = state
            self.update()

    def paintEvent(self, _event):
        p = QPainter(self)
        draw_orb(p, self.rect(), self._state, self._angle, self._breath)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.drag_window is not None:
            self._drag_offset = event.globalPosition().toPoint() - self.drag_window.pos()
            self._dragged = False
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_offset is not None:
            self._dragged = True
            self.drag_window.move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._drag_offset is not None:
            offset, dragged = self._drag_offset, self._dragged
            self._drag_offset = None
            self._dragged = False
            event.accept()
            if dragged and hasattr(self.drag_window, "snap_to_edge"):
                self.drag_window.snap_to_edge()
            else:
                self.clicked.emit()
            return
        super().mouseReleaseEvent(event)
