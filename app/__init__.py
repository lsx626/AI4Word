# -*- coding: utf-8 -*-
"""AI4Word GUI 层（PySide6 悬浮窗）。

本包只放界面与界面背后的线程胶水；核心排版能力全部复用根目录的
ai_client / doc_model / streaming_writer / format_runner / session / styles，
CLI（main.py）与 GUI（app 包）共用 app.agent 里的代码生成提示词与执行环境。
"""
__version__ = "9.3"
