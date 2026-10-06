# -*- coding: utf-8 -*-
"""离线测试：GUI 外观与行为回归（offscreen 平台）。

覆盖 V9.0 修复项：
- 预设下拉不再是逐字符（preset_list 返回列表）
- 消息气泡高度不再被截断（多行流式文本完整可见）
- 悬浮窗不透明（WA_TranslucentBackground 已移除，调色板为深色）
- 贴边吸附：松手吸附 + 拖动中磁性吸附
- 进度状态条接线
- reload_flags 可见 + flags 未变时显式补注册热键（showEvent 不重发）
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


def test_preset_combo_programmatic_selection_applies_preset():
    """回归：组合框「程序化改选」必须触发 _on_preset 并投递 preset 命令。

    旧 bug：preset_combo 只绑了 QComboBox.activated —— 该信号仅在用户真实
    弹窗选择时触发，setCurrentIndex 改选永不触发，_on_preset 被静默跳过
    （组合框不复位、preset 命令不投递），deep_test review_presets_save 的
    preset_selected/preset_applied 缺失与「应用后组合框复位 ×4 FAIL」即此根因。
    同时钉死三条约束：复位回充与占位项不重复投递；真实键盘改选只投递一次
    （activated 与 currentIndexChanged 不可并存，否则用户选择会双次投递）；
    生成中改选被拒绝且不投递。
    """
    win = _window()
    worker = win.worker
    combo = win.preset_combo
    win._busy = False
    q = worker._queue
    # 构造时 MainWindow 会投递首条 speed 命令（无 worker 线程消费），先清空
    while not q.empty():
        q.get_nowait()

    idx = combo.findText("简历")
    assert idx > 0

    # 程序化改选 -> 恰好一条 preset 命令 + 组合框复位
    combo.setCurrentIndex(idx)
    cmd, payload = q.get(timeout=2.0)
    assert cmd == "preset" and payload == "简历", (cmd, payload)
    assert combo.currentIndex() == 0, "应用后组合框应复位到占位项"
    assert q.empty(), "复位回充（currentIndexChanged(0)）不应再次投递命令"

    # 占位项（idx<=0）被选中不投递
    combo.setCurrentIndex(0)
    assert q.empty(), "选中占位项不应投递命令"

    # 真实用户选择路径（键盘导航；activated 与 currentIndexChanged 都会触发）
    # -> 也只投递一次
    win.show()
    QTest.qWait(30)
    QTest.keyClick(combo, Qt.Key_Down)   # 0 -> 1（公文）
    cmd, payload = q.get(timeout=2.0)
    assert cmd == "preset" and payload == "公文", (cmd, payload)
    assert combo.currentIndex() == 0
    assert q.empty(), "键盘改选不应双次投递命令"

    # 生成中改选应被拒绝：不投递命令且回充复位
    win._busy = True
    combo.setCurrentIndex(idx)
    assert q.empty(), "生成中不应投递 preset 命令"
    assert combo.currentIndex() == 0, "被拒绝后组合框应复位"

    win.close()


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


def test_reload_flags_reregisters_hotkey_when_visible():
    """回归：可见窗口 + flags 未变时，reload_flags 不能只靠 showEvent 重注册热键。

    机制（PySide6 实测）：setWindowFlags 传入相同 flags 时 Qt 直接
    early-return，已可见的窗口不隐藏不重建，随后的 show() 是空操作、不重发
    showEvent —— 旧实现先 _unregister_hotkey() 再把重注册完全托付给
    showEvent，这条路径静默失效：热键被注销后没人注册回来，Ctrl+Alt+Space
    就此失效（真机日志：summon 后 reload_flags 无 hotkey_register）。"""
    win = _window()
    calls = {"reg": 0, "unreg": 0}
    shows = {"n": 0}

    def fake_register():
        calls["reg"] += 1
        win._hotkey = True

    def fake_unregister():
        calls["unreg"] += 1
        win._hotkey = False

    # 打桩：离线环境 winId 不是真实 HWND，RegisterHotKey 结果不可控；
    # 这里只验证 reload_flags 的「注销 -> 重注册」接线，不碰 Windows API
    win._register_hotkey = fake_register
    win._unregister_hotkey = fake_unregister
    orig_show_event = win.showEvent

    def counting_show_event(event):
        shows["n"] += 1
        orig_show_event(event)

    win.showEvent = counting_show_event

    win.show()  # 首次 show -> showEvent -> 注册一次
    assert calls["reg"] == 1 and calls["unreg"] == 0 and win._hotkey, calls
    assert shows["n"] == 1, shows

    # 可见 + settings 未变：flags 与 __init__ 里 set_flags() 设的完全相同
    win.reload_flags()
    assert calls["unreg"] == 1, calls
    assert shows["n"] == 1, \
        f"flags 未变已可见时不应重发 showEvent（实际 {shows['n']} 次）"
    assert win._hotkey, "reload_flags 后热键未重新注册：Ctrl+Alt+Space 会失效"
    assert calls["reg"] == 2, \
        f"reload_flags 未显式补注册（reg={calls['reg']}，期望 2）"
    win.close()


def test_reload_flags_reregisters_hotkey_when_flags_changed():
    """可见窗口但 flags 真的变化（stay_on_top 切换）时 reload_flags 的契约：
    无论走 showEvent 重发还是显式补注册，结束时热键必须处于已注册状态。"""
    win = _window()
    calls = {"reg": 0}
    win._register_hotkey = lambda: (calls.__setitem__("reg", calls["reg"] + 1),
                                    setattr(win, "_hotkey", True))
    win._unregister_hotkey = lambda: setattr(win, "_hotkey", False)
    keep = bool(win.settings.get("stay_on_top", True))
    win.show()
    assert calls["reg"] == 1, calls

    win.settings.set("stay_on_top", not keep)
    win.reload_flags()
    assert win._hotkey, "flags 变化路径 reload_flags 后热键也未注册"
    assert calls["reg"] >= 2, calls
    win.settings.set("stay_on_top", keep)
    win.close()


def test_reload_flags_reregisters_hotkey_when_hidden():
    """隐藏态（最小化到托盘）reload_flags：隐藏分支已显式注册，保持不变。"""
    win = _window()
    calls = {"reg": 0}
    win._register_hotkey = lambda: (calls.__setitem__("reg", calls["reg"] + 1),
                                    setattr(win, "_hotkey", True))
    win._unregister_hotkey = lambda: setattr(win, "_hotkey", False)
    assert not win.isVisible()

    win.reload_flags()
    assert win._hotkey, "隐藏态 reload_flags 后热键未注册（托盘期间失联）"
    assert calls["reg"] == 1, calls
    win.close()


if __name__ == "__main__":
    for name, fn in sorted(list(globals().items())):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"ok: {name}")
    print("ALL UI TESTS DONE")
