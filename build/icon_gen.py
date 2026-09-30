# -*- coding: utf-8 -*-
"""生成应用图标：QPainter 画 1024px 发光球，转多尺寸 ICO / PNG。

输出：assets/app_icon.ico（16/24/32/48/64/128/256）+ assets/app_icon.png。
无外部素材依赖，画法与 app/icons.app_icon 完全一致。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PIL import Image
from PySide6.QtCore import QRect, Qt
from PySide6.QtGui import QPainter, QPixmap
from PySide6.QtWidgets import QApplication

QApplication(sys.argv)  # QPixmap / 文本渲染需要 QGuiApplication

from app.avatar import draw_orb

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
SIZES = (16, 24, 32, 48, 64, 128, 256)


def render_png(size):
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    draw_orb(p, pix.rect(), "idle", 30.0, 0.15)
    p.end()
    return pix


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    big = render_png(1024).toImage()
    png_path = os.path.join(OUT_DIR, "app_icon.png")
    big.save(png_path, "PNG")
    pil_big = Image.open(png_path).convert("RGBA")
    frames = [pil_big.resize((s, s), Image.LANCZOS) for s in SIZES]
    ico_path = os.path.join(OUT_DIR, "app_icon.ico")
    pil_big.save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES])
    for s in (16, 32, 256):
        frames[SIZES.index(s)].save(os.path.join(OUT_DIR, f"app_icon_{s}.png"), "PNG")
    print(f"icon -> {ico_path} ({os.path.getsize(ico_path)} bytes), {png_path}")


if __name__ == "__main__":
    main()
