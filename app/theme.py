# -*- coding: utf-8 -*-
"""ink & amber 深色主题：颜色常量 + 全套 QSS。

刻意避开紫色主导的常见深色方案，用墨黑底 + 琥珀色强调点，配圆角与
半透明层次，让悬浮窗在任意壁纸上都不刺眼。
"""

INK_0 = "#131419"      # 最深：窗口背景渐变起点
INK_1 = "#1b1d24"      # 窗口背景
INK_2 = "#23262f"      # 面板 / 助手气泡
INK_3 = "#2e323d"      # 边框 / 输入框
INK_4 = "#3a3f4c"      # 悬停高亮

TEXT = "#ece9e1"       # 主文字
TEXT_DIM = "#9aa0ab"   # 次要文字
TEXT_FAINT = "#6b7078" # 更弱（时间戳等）

AMBER = "#f2a93b"      # 强调色
AMBER_SOFT = "#ffc46b" # 强调亮色
AMBER_DEEP = "#b97a1c" # 强调暗色

RED = "#e0655a"        # 错误
GREEN = "#62b97c"      # 成功

CODE_FONT = "Consolas, Cascadia Code, Menlo, monospace"
UI_FONT = "Microsoft YaHei UI, PingFang SC, Noto Sans CJK SC, sans-serif"


def qss():
    """整套样式表。窗口是无边框不透明的，背景由 paintEvent 自绘，圆角由 mask 切出。"""
    return f"""
    QWidget {{
        font-family: "{UI_FONT}";
        font-size: 13px;
        color: {TEXT};
    }}
    QLineEdit, QTextEdit, QPlainTextEdit {{
        background: rgba(19, 20, 25, 235);
        border: 1px solid {INK_3};
        border-radius: 10px;
        padding: 7px 10px;
        selection-background-color: {AMBER_DEEP};
    }}
    QTextEdit QScrollBar:vertical {{
        background: transparent;
        width: 8px;
        margin: 2px 0;
    }}
    QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
        border: 1px solid {AMBER_DEEP};
    }}
    QPushButton {{
        background: {INK_3};
        border: 1px solid {INK_4};
        border-radius: 9px;
        padding: 6px 12px;
        color: {TEXT};
    }}
    QPushButton:hover {{ background: {INK_4}; }}
    QPushButton:pressed {{ background: {INK_2}; }}
    QPushButton:disabled {{ color: {TEXT_FAINT}; border-color: {INK_3}; }}
    QPushButton#primary {{
        background: {AMBER_DEEP};
        border: 1px solid {AMBER};
        color: #241a08;
        font-weight: 600;
    }}
    QPushButton#primary:hover {{ background: {AMBER}; }}
    QPushButton#tool {{
        background: transparent;
        border: 1px solid transparent;
        border-radius: 8px;
        padding: 5px 8px;
        color: {TEXT_DIM};
    }}
    QPushButton#tool:hover {{ background: {INK_3}; color: {TEXT}; }}
    QPushButton#tool:checked {{
        background: rgba(242, 169, 59, 30);
        border-color: {AMBER_DEEP};
        color: {AMBER_SOFT};
    }}
    QLabel#wordStatus {{ color: {TEXT_DIM}; font-size: 12px; }}
    QLabel#title {{ font-size: 14px; font-weight: 600; color: {TEXT}; }}
    QLabel#statusBar {{ color: {AMBER}; font-size: 12px; }}
    QComboBox {{
        background: {INK_3};
        border: 1px solid {INK_4};
        border-radius: 8px;
        padding: 4px 10px;
        color: {TEXT};
    }}
    QComboBox:hover {{ border-color: {AMBER_DEEP}; }}
    QComboBox QAbstractItemView {{
        background: {INK_1};
        border: 1px solid {INK_4};
        selection-background-color: {AMBER_DEEP};
        outline: 0;
    }}
    QCheckBox {{ color: {TEXT_DIM}; spacing: 8px; }}
    QCheckBox::indicator {{
        width: 15px; height: 15px;
        border-radius: 4px;
        border: 1px solid {INK_4};
        background: {INK_2};
    }}
    QCheckBox::indicator:checked {{
        background: {AMBER_DEEP};
        border-color: {AMBER};
    }}
    QScrollBar:vertical {{
        background: transparent;
        width: 9px;
        margin: 4px 2px;
    }}
    QScrollBar::handle:vertical {{
        background: {INK_4};
        border-radius: 4px;
        min-height: 28px;
    }}
    QScrollBar::handle:vertical:hover {{ background: {AMBER_DEEP}; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
    QScrollBar:horizontal {{ height: 0; }}
    QListWidget#blockMap {{
        background: rgba(19, 20, 25, 180);
        border: 1px solid {INK_3};
        border-radius: 10px;
        outline: 0;
    }}
    QListWidget#blockMap::item {{ padding: 5px 8px; border-radius: 6px; }}
    QListWidget#blockMap::item:hover {{ background: {INK_3}; }}
    QListWidget#blockMap::item:selected {{
        background: rgba(242, 169, 59, 34);
        color: {AMBER_SOFT};
    }}
    QSplitter::handle {{ background: transparent; }}
    QDialog {{ background: {INK_1}; }}
    QLabel#link {{ color: {AMBER_SOFT}; }}
    """
