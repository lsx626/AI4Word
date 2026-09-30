# -*- coding: utf-8 -*-
"""系统托盘：显示/隐藏、开机自启、设置、退出；关窗即最小化到托盘。"""
from PySide6.QtWidgets import QApplication, QMenu, QSystemTrayIcon

from app.auto_start import is_enabled
from app.icons import app_icon


class Tray(QSystemTrayIcon):
    def __init__(self, window, settings, parent=None):
        super().__init__(app_icon(64), parent)
        self.window = window
        self.settings = settings
        self._hinted = False
        self.setToolTip("AI4Word 悬浮助手（双击召回，或 Ctrl+Alt+Space）")

        menu = QMenu()
        self.act_show = menu.addAction("显示悬浮窗")
        self.act_show.triggered.connect(window.summon)
        menu.addSeparator()
        self.act_autostart = menu.addAction("开机自启动")
        self.act_autostart.setCheckable(True)
        self.act_autostart.setChecked(bool(is_enabled()))
        self.act_autostart.triggered.connect(self._toggle_autostart)
        self.act_settings = menu.addAction("设置…")
        self.act_settings.triggered.connect(window.open_settings)
        menu.addSeparator()
        self.act_quit = menu.addAction("退出")
        self.act_quit.triggered.connect(self._quit)
        self.setContextMenu(menu)

        self.activated.connect(self._on_activated)

    def _on_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.window.summon()

    def _toggle_autostart(self, on):
        from app.auto_start import disable, enable
        if on:
            enable()
        else:
            disable()
        self.settings.set("auto_start", bool(on))
        self.settings.save()

    def first_hide_hint(self):
        """第一次最小化到托盘时提示召回方式。"""
        if not self._hinted:
            self._hinted = True
            self.showMessage("AI4Word",
                             "已最小化到托盘。双击托盘图标，或按 Ctrl+Alt+Space 召回。",
                             QSystemTrayIcon.Information, 4000)

    def _quit(self):
        self.window._unregister_hotkey()
        QApplication.quit()
