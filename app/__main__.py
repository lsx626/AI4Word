# -*- coding: utf-8 -*-
"""GUI 启动入口：python -m app（开发态）或由打包后的 ai4word.pyw 调用。"""
import os
import sys

from dotenv import load_dotenv
from PySide6.QtCore import QSharedMemory, QTimer
from PySide6.QtWidgets import QApplication, QMessageBox

from app import __version__
from app.engine import AgentWorker
from app.icons import app_icon
from app.main_window import MainWindow
from app.settings import Settings
from app.theme import qss
from app.tray import Tray


def main(argv=None):
    argv = sys.argv if argv is None else list(argv)
    load_dotenv()  # 与 CLI 一致：.env 里的密钥作为默认来源

    app = QApplication(argv)
    app.setApplicationName("AI4Word")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("AI4Word")
    app.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(app_icon(64))
    app.setStyleSheet(qss())

    # 单实例：已运行则提示并退出
    shm = QSharedMemory("AI4Word-SingleInstance-v8")
    if not shm.create(1):
        QMessageBox.information(None, "AI4Word", "AI4Word 已经在运行了。")
        return 1

    settings = Settings().load()
    settings.apply_env()

    from app.auto_start import is_enabled
    try:
        settings.set("auto_start", bool(is_enabled()))
    except Exception:
        pass

    worker = AgentWorker()
    worker.set_api_key(settings.get("api_key") or os.environ.get("ATRIA_API_KEY") or "")
    worker.send("speed", settings.get("speed") or "auto")

    window = MainWindow(worker, settings)
    tray = Tray(window, settings)
    window.set_tray(tray)
    tray.show()
    window.show()  # 悬浮窗随启动出现（之后可最小化到托盘）

    worker.start()

    if not (settings.get("api_key") or "").strip():
        # 首次运行：引导填写 API 密钥
        QTimer.singleShot(600, window.open_settings)

    code = app.exec()
    try:
        # 退出时若 worker 正卡在「中断选择等待」上（_wait_choice 阻塞
        # 最多 5 分钟），quit 要等它结束才被处理 → 进程僵死。
        # 先喂一个 keep 选择解除等待，quit 随即被消费。
        worker.choose("keep")
        worker.send("quit")
        worker.wait(4000)
    except Exception:
        pass
    settings.save()
    return code


if __name__ == "__main__":
    sys.exit(main())
