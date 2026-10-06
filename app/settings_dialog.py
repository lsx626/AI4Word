# -*- coding: utf-8 -*-
"""设置对话框：API 密钥 / 服务地址 / 模型 / 开机自启 / 窗口置顶。"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QDialog, QDialogButtonBox, QFormLayout,
                                QHBoxLayout, QLabel, QLineEdit, QMessageBox,
                                QPushButton, QVBoxLayout)

from app import debug
from app.auto_start import disable, enable, is_enabled
from app.settings import DEFAULT_BASE_URL, DEFAULT_MODEL


class SettingsDialog(QDialog):
    def __init__(self, settings, worker, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.worker = worker
        self.setWindowTitle("AI4Word 设置")
        self.setModal(True)
        self.setMinimumWidth(440)

        form = QFormLayout()
        form.setSpacing(10)

        self.key_edit = QLineEdit(settings.get("api_key") or "", self)
        self.key_edit.setEchoMode(QLineEdit.Password)
        self.key_edit.setPlaceholderText("sk-…（discovery 平台密钥）")
        show_btn = QPushButton("显示", self)
        show_btn.setCheckable(True)
        show_btn.setFixedWidth(52)
        show_btn.toggled.connect(
            lambda on: self.key_edit.setEchoMode(
                QLineEdit.Normal if on else QLineEdit.Password))
        key_row = QHBoxLayout()
        key_row.setContentsMargins(0, 0, 0, 0)
        key_row.setSpacing(6)
        key_row.addWidget(self.key_edit, 1)
        key_row.addWidget(show_btn)
        form.addRow("API 密钥", key_row)

        self.url_edit = QLineEdit(settings.get("base_url") or DEFAULT_BASE_URL, self)
        self.url_edit.setPlaceholderText(DEFAULT_BASE_URL)
        form.addRow("服务地址", self.url_edit)

        self.model_edit = QLineEdit(settings.get("model") or DEFAULT_MODEL, self)
        form.addRow("模型", self.model_edit)

        hint = QLabel(
            '密钥到 <a href="https://discovery.intern-ai.org.cn/">discovery.intern-ai.org.cn</a>'
            ' 申请，模型默认 Atria-Dawn-Preview。', self)
        hint.setObjectName("link")
        hint.setOpenExternalLinks(True)
        hint.setWordWrap(True)

        self.auto_check = QCheckBox("开机自动启动 AI4Word", self)
        self.auto_check.setChecked(bool(is_enabled()))
        self.top_check = QCheckBox("悬浮窗始终置顶", self)
        self.top_check.setChecked(bool(settings.get("stay_on_top", True)))

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 14)
        lay.setSpacing(10)
        lay.addLayout(form)
        lay.addWidget(hint)
        lay.addWidget(self.auto_check)
        lay.addWidget(self.top_check)
        lay.addStretch(1)
        lay.addWidget(buttons)
        debug.log("settings_open",
                  auto_start=self.auto_check.isChecked(),
                  stay_on_top=self.top_check.isChecked())

    def accept(self):
        debug.log("settings_accept", api_key=self.key_edit.text().strip(),
                  model=self.model_edit.text().strip(),
                  base_url=self.url_edit.text().strip(),
                  auto_start=self.auto_check.isChecked(),
                  stay_on_top=self.top_check.isChecked())
        self.settings.set("api_key", self.key_edit.text().strip())
        self.settings.set("base_url", self.url_edit.text().strip() or DEFAULT_BASE_URL)
        self.settings.set("model", self.model_edit.text().strip() or DEFAULT_MODEL)
        self.settings.set("stay_on_top", self.top_check.isChecked())
        self.settings.apply_env()
        if self.worker is not None:
            self.worker.set_api_key(self.settings.get("api_key"))

        # 开机自启：注册表写入可能失败（组策略 / 杀软拦截），必须反馈，
        # 否则复选框勾着、settings 记着，下次开机却不自启
        want_auto = self.auto_check.isChecked()
        if want_auto:
            ok_auto = enable()
            if not ok_auto:
                QMessageBox.warning(self, "开机自启", "写入注册表失败（可能被杀毒软件或"
                                    "组策略拦截），开机自启未能生效。可在系统设置里"
                                    "手动添加启动项。")
                self.auto_check.setChecked(False)
                want_auto = False
        else:
            # 本机从未勾过自启时 Run 值不存在：先问 is_enabled()，避免无谓
            # 以 KEY_SET_VALUE 打开注册表（disable() 也已把「值不存在」
            # 当作 no-op 成功，不再产出 autostart_disable_failed 噪声告警）
            if is_enabled():
                disable()
        self.settings.set("auto_start", want_auto)

        # 保存失败不能静默：用户会以为密钥已经存好
        if not self.settings.save():
            debug.error("settings_save_failed", path=self.settings.path)
            QMessageBox.warning(self, "保存失败", f"设置未能写入磁盘（目录不可写或被占用）：\n"
                                f"{self.settings.path}\n本次会话仍按新设置使用，"
                                f"重启后失效。")
            return
        debug.log("settings_saved", path=self.settings.path)
        super().accept()

    def done(self, result):
        # settings_open 的对称收口：accept（确定）/ reject（取消）/ 按 Esc /
        # 点 X 最终都经 done()，所以这里发一次 settings_closed 覆盖所有入口
        # （含不经 MainWindow._open_settings 直接构造对话框的调用方）。
        # accept 在保存失败时提前 return、done 不被调用——对话框没关，
        # 就不能报「已关闭」。
        debug.log("settings_closed",
                  accepted=bool(result == QDialog.Accepted),
                  api_key=self.settings.get("api_key"))
        super().done(result)
