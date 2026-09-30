# -*- coding: utf-8 -*-
"""离线测试：GUI 外观与行为回归（offscreen 平台）。

覆盖 V9.0 修复项：
- 预设下拉不再是逐字符（preset_list 返回列表）
- 消息气泡高度不再被截断（多行流式文本完整可见）
- 悬浮窗不透明（WA_TranslucentBackground 已移除，调色板为深色）
- 贴边吸附：松手吸附 + 拖动中磁性吸附
- 进度状态条接线
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QPoint, QRect, Qt  # noqa: E402
from PySide6.QtGui import QColor  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from app.engine import AgentWorker  # noqa: E402
from app.main_window import MainWindow  # noqa: E402
from app.messages import MessageList  # noqa: E402
from app.settings import Settings  # noqa: E402
from app.theme import qss  # noqa: E402
from styles import preset_list  # noqa: E402

_app = QApplication.instance() or QApplication(sys.argv)
_app.setStyleSheet(qss())


def _settings(tmpname="ai4w_ui.json"):
    s = Settings(os.path.join(os.environ["TEMP"], tmpname))
    s.load()
    s.set("geometry", [60, 60])
    s.set("expanded", False)
    return s


def _window():
    return MainWindow(AgentWorker(), _settings())


def test_preset_list_is_real_list():
    names = preset_list()
    assert isinstance(names, list), f"preset_list 必须返回 list，实际 {type(names)}"
    assert names == ["公文", "博客", "简历", "论文"], names
    from styles import preset_names
    assert isinstance(preset_names(), str)  # AI 提示词用字符串版本


def test_preset_combo_items():
    win = _window()
    items = [win.preset_combo.itemText(i) for i in range(win.preset_combo.count())]
    assert "套用预设…" in items[0]
    assert set(items[1:]) == {"论文", "公文", "简历", "博客"}, items
    assert len(items) == 5, f"预设下拉项数错误: {items}"


def test_message_label_full_height():
    """长文本消息的标签高度必须容纳全部行（旧 bug：只剩一行高）。"""
    import math
    sa = MessageList()
    sa.setFixedSize(560, 400)
    b = sa.begin_stream()
    long_text = "秋天来了，树叶黄了。这是一段比较长的流式文本内容，" \
                "用来验证消息气泡的自动换行高度是否完整显示。再加一句，确保多行。"
    sa.stream_append(long_text)
    sa.end_stream()
    sa.show()
    QTest.qWait(100)
    lbl = b._label
    # 独立测量：按实际宽度折行所需高度（QFontMetrics，不依赖 QLabel 自身布局）
    fm = lbl.fontMetrics()
    text_w = fm.horizontalAdvance(long_text)
    lines = max(2, math.ceil(text_w / max(1, lbl.width())))
    need = lines * fm.height()
    assert lbl.height() >= need - 2, \
        f"标签高度 {lbl.height()} 不足：文本约 {lines} 行需要 {need}" \
        f"（label.width={lbl.width()} text_w={text_w} line_h={fm.height()}）"
    assert lbl.height() >= fm.height() * 2, \
        f"标签高度被截断: {lbl.height()}（单行 {fm.height()}）"
    sa.close()


def test_message_text_is_painted():
    """消息文字必须真正画出来（白字像素 > 0），而不只是有几何。"""
    sa = MessageList()
    sa.setFixedSize(560, 300)
    sa.add("assistant", "一段足够长的消息文本，用于检查文字像素确实被绘制出来。")
    sa.show()
    QTest.qWait(100)
    pix = sa.grab()
    img = pix.toImage()
    from PySide6.QtCore import QByteArray
    buf = QByteArray()
    ok, _i = 0, 0
    for y in range(0, img.height(), 2):
        for x in range(0, img.width(), 2):
            c = img.pixelColor(x, y)
            _i += 1
            if c.red() > 200 and c.green() > 190 and c.blue() > 180:
                ok += 1
    sa.close()
    assert ok > 50, f"消息文字像素过少（{ok}），可能未被绘制"


def test_window_is_opaque():
    """V9.0：窗口不再用 WA_TranslucentBackground（输入时背景变透明的根因）。"""
    win = _window()
    assert not win.testAttribute(Qt.WA_TranslucentBackground)
    # 调色板底角色必须是深色，任何自动填充都不会露出灰白
    bg = win.palette().color(win.backgroundRole())
    assert bg.lightness() < 40, f"窗口背景色过亮: {bg.name()}"
    win.close()


def test_snap_on_release():
    win = _window()
    win.show()
    QTest.qWait(30)
    screen = QApplication.primaryScreen()
    ag = screen.availableGeometry()

    win.move(ag.left() + 20, ag.top() + 200)
    win.snap_to_edge()
    assert win.x() <= ag.left() + 12, f"左吸附失败: x={win.x()}"

    win.move(ag.right() - win.width() - 15, ag.top() + 200)
    win.snap_to_edge()
    assert win.x() >= ag.right() - win.width() - 12, f"右吸附失败: x={win.x()}"

    win.move(ag.left() + 300, ag.bottom() - win.height() - 12)
    win.snap_to_edge()
    assert win.y() >= ag.bottom() - win.height() - 12, f"底吸附失败: y={win.y()}"

    # 远离边缘不吸附
    win.move(ag.left() + 200, ag.top() + 200)
    x0, y0 = win.x(), win.y()
    win.snap_to_edge()
    assert (win.x(), win.y()) == (x0, y0), "远离边缘时不应吸附"
    win.close()


def test_drag_magnetism():
    """拖动中靠近边缘即吸附（drag_to），不必等松手。"""
    win = _window()
    win.show()
    QTest.qWait(30)
    screen = QApplication.primaryScreen()
    ag = screen.availableGeometry()
    origin = QPoint(ag.left() + 300, ag.top() + 200)
    win.move(origin)
    offset = QPoint(30, 30)  # 光标相对窗口左上角的偏移
    # 拖到左边缘附近
    win.drag_to(QPoint(ag.left() + 5, origin.y()), offset)
    assert win.x() <= ag.left() + 12, f"拖动磁吸失败: x={win.x()}"
    # 拖回中间：脱离磁吸
    win.drag_to(QPoint(ag.left() + 300, origin.y()), offset)
    assert win.x() > ag.left() + 12, "拖离边缘后应脱离磁吸"
    win.close()


def test_progress_bar_wiring():
    win = _window()
    win.set_expanded(True)  # 展开态下状态条才属于可见面板
    win._on_progress("正在生成… 已接收 12 字")
    assert win._status_label.isVisibleTo(win._panel)
    assert "12" in win._status_label.text()
    win._on_state("idle")
    assert not win._status_label.isVisibleTo(win._panel)
    assert not win._status_label.text()
    win.close()


def test_input_multiline_scroll():
    win = _window()
    assert win.input_p.verticalScrollBarPolicy() == Qt.ScrollBarAsNeeded
    win.close()


def test_rounded_mask_applied():
    offscreen = os.environ.get("QT_QPA_PLATFORM") == "offscreen"
    win = _window()
    win.show()
    QTest.qWait(30)
    mask = win.mask()
    if offscreen:
        # offscreen 平台不支持 mask（仅打印警告），至少不能报错
        win.close()
        return
    assert not mask.isEmpty(), "圆角 mask 未生效"
    # mask 的外包矩形应接近窗口大小（四个角被切掉）
    br = mask.boundingRect()
    assert br.width() >= win.width() - 6 and br.height() >= win.height() - 6
    # 角落应该在 mask 之外（露出桌面、不接收点击）
    assert not mask.contains(QPoint(2, 2)), "左上角应在圆角 mask 之外"
    win.close()


if __name__ == "__main__":
    for name, fn in sorted(list(globals().items())):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"ok: {name}")
    print("ALL UI TESTS DONE")
