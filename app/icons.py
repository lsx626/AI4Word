# -*- coding: utf-8 -*-
"""矢量图标：全部 QPainter 现绘，避免外部图标依赖。"""
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPen, QPixmap, QPolygonF
from PySide6.QtWidgets import QApplication

from app.avatar import draw_orb
from app.theme import AMBER, AMBER_SOFT, INK_3, INK_4, TEXT


def _canvas(size=20, color=TEXT):
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    return pix


def _icon(draw_fn, size=20, color=None):
    pix = _canvas(size)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing, True)
    if color is not None:
        p.setPen(QPen(QColor(color), 1.6, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
    draw_fn(p, size)
    p.end()
    return QIcon(pix)


def _arrow_up(p, s):
    # 发送：上箭头
    p.drawLine(s * 0.5, s * 0.82, s * 0.5, s * 0.22)
    p.drawPolyline(_poly(s, [(0.28, 0.44), (0.5, 0.2), (0.72, 0.44)]))


def _chevron_expand(p, s):
    # 展开：两条向外的箭头
    p.drawLine(s * 0.3, s * 0.62, s * 0.3, s * 0.4)
    p.drawLine(s * 0.3, s * 0.4, s * 0.12, s * 0.4)
    p.drawLine(s * 0.7, s * 0.38, s * 0.7, s * 0.6)
    p.drawLine(s * 0.7, s * 0.6, s * 0.88, s * 0.6)


def _chevron_collapse(p, s):
    # 收起：两条向内的箭头
    p.drawLine(s * 0.3, s * 0.62, s * 0.3, s * 0.4)
    p.drawLine(s * 0.3, s * 0.62, s * 0.12, s * 0.62)
    p.drawLine(s * 0.7, s * 0.38, s * 0.7, s * 0.6)
    p.drawLine(s * 0.7, s * 0.38, s * 0.88, s * 0.38)


def _stop(p, s):
    # 中断：实心圆角方块
    p.setBrush(QColor("#e0655a"))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(QRectF(s * 0.28, s * 0.28, s * 0.44, s * 0.44), 3, 3)


def _gear(p, s):
    # 设置：简化齿轮
    cx, cy, r = s * 0.5, s * 0.5, s * 0.2
    p.drawEllipse(QRectF(cx - r, cy - r, r * 2, r * 2))
    import math
    for k in range(8):
        a = math.radians(k * 45)
        x1 = cx + math.cos(a) * r
        y1 = cy + math.sin(a) * r
        x2 = cx + math.cos(a) * r * 1.7
        y2 = cy + math.sin(a) * r * 1.7
        p.drawLine(x1, y1, x2, y2)


def _save(p, s):
    # 保存：软盘
    p.drawRoundedRect(QRectF(s * 0.2, s * 0.2, s * 0.6, s * 0.6), 2, 2)
    p.fillRect(QRectF(s * 0.3, s * 0.2, s * 0.4, s * 0.18), QColor(INK_3))
    p.fillRect(QRectF(s * 0.32, s * 0.5, s * 0.36, s * 0.24), QColor(INK_3))


def _list(p, s):
    # 块地图：三条横线带圆点
    for k, y in enumerate((0.28, 0.5, 0.72)):
        y = s * y
        p.drawLine(s * 0.28, y, s * 0.82, y)
        p.drawEllipse(QRectF(s * 0.13 - 1.5, y - 1.5, 3, 3))


def _close(p, s):
    p.drawLine(s * 0.3, s * 0.3, s * 0.7, s * 0.7)
    p.drawLine(s * 0.7, s * 0.3, s * 0.3, s * 0.7)


def _rollback(p, s):
    # 回滚：逆时针箭头
    import math
    cx, cy, r = s * 0.5, s * 0.5, s * 0.26
    a0 = math.radians(-40)
    a1 = math.radians(200)
    from PySide6.QtGui import QPainterPath
    path = QPainterPath()
    path.arcMoveTo(cx - r, cy - r, r * 2, r * 2, a0)
    path.arcTo(cx - r, cy - r, r * 2, r * 2, a0, a1 - a0)
    p.drawPath(path)
    pts = []
    for dx, dy in ((-4, -2), (1, 5), (6, -4)):
        pts.append((cx + r * math.cos(a0) + dx, cy + r * math.sin(a0) + dy))
    p.drawPolyline(_poly_from(pts))


def _keep(p, s):
    # 保留：对勾
    p.drawPolyline(_poly(s, [(0.25, 0.52), (0.44, 0.72), (0.75, 0.3)]))


def _poly(size, ratios):
    """[(x_ratio, y_ratio), ...] -> QPolygonF。"""
    return QPolygonF([QPointF(size * rx, size * ry) for rx, ry in ratios])


def _poly_from(pts):
    return QPolygonF([QPointF(x, y) for x, y in pts])


def icon_send():
    return _icon(_arrow_up, 20, AMBER_SOFT)


def icon_expand():
    return _icon(_chevron_expand, 18, TEXT)


def icon_collapse():
    return _icon(_chevron_collapse, 18, TEXT)


def icon_stop():
    return _icon(_stop, 18)


def icon_gear():
    return _icon(_gear, 18, TEXT)


def icon_save():
    return _icon(_save, 18, TEXT)


def icon_list():
    return _icon(_list, 18, TEXT)


def icon_close():
    return _icon(_close, 16, TEXT)


def icon_rollback():
    return _icon(_rollback, 18, AMBER_SOFT)


def icon_keep():
    return _icon(_keep, 18, QColor("#62b97c"))


def app_icon(size=256):
    """应用 / 托盘图标：发光球本体。"""
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    draw_orb(p, pix.rect(), "idle", 30.0, 0.15)
    p.end()
    return QIcon(pix)
