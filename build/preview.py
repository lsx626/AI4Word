# -*- coding: utf-8 -*-
"""视觉 QA：离线渲染悬浮窗两种形态的截图（QT_QPA_PLATFORM=offscreen）。"""
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

app = QApplication(["AI4Word"])

from app.settings import Settings
from app.engine import AgentWorker
from app.main_window import MainWindow
from app.theme import qss

app.setStyleSheet(qss())

tmp = tempfile.mkdtemp()
settings = Settings(path=os.path.join(tmp, "settings.json")).load()
settings.set("geometry", [60, 60])

worker = AgentWorker()  # 不启动：只用于信号接线
win = MainWindow(worker, settings)

# 喂入假数据：消息流 / 流式气泡 / 块地图
win.messages.add("user", "写一篇关于秋天的散文，要有三个小标题。")
b = win.messages.begin_stream()
for piece in ["# 秋天的风\n\n", "第一片叶子落下的时候", "，风里已经有了桂花的味道。"]:
    b.append_text(piece)
win.messages.end_stream()
win.messages.add("info", "已写入 Word。")
win.messages.add("code", "[1/3] doc.Styles(WD_STYLE_HEADING_1).ParagraphFormat.Alignment = 1")
win.messages.add("code", "    [ok]")
win.messages.add("error", "第 2 块不存在，请检查索引。")
win.word_label.setText("已连接 · 秋天.docx")
win.block_panel.set_text(
    "[0] heading1: 秋天的风\n"
    "[1] paragraph: 第一片叶子落下的时候，风里已经有了桂花的味道。\n"
    "[2] heading2: 桂花\n"
    "[3] list: 银桂 / 金桂 / 丹桂\n"
    "[4] code: import this")
win.messages.add("user", "把所有一级标题居中并改为深蓝色")

win.show()
win.repaint()
app.processEvents()
out = os.environ.get("AI4WORD_PREVIEW_OUT") or os.path.join(ROOT, "build")
os.makedirs(out, exist_ok=True)

compact_path = os.path.join(out, "preview_compact.png")
win.grab().save(compact_path)
print("saved", compact_path)

win._apply_expanded(True, instant=True)
win.repaint()
app.processEvents()
expanded_path = os.path.join(out, "preview_expanded.png")
win.grab().save(expanded_path)
print("saved", expanded_path)

win.btn_blockmap.setChecked(True)
win._toggle_blockmap()
win.messages.add("assistant", "好的，正在把一级标题居中并改为深蓝色…")
win.repaint()
app.processEvents()
blocks_path = os.path.join(out, "preview_blocks.png")
win.grab().save(blocks_path)
print("saved", blocks_path)
