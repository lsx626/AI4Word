# -*- coding: utf-8 -*-
"""设置对话框：API 密钥 / 服务地址 / 模型 / 开机自启 / 窗口置顶。"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QDialog, QDialogButtonBox,
                                QFormLayout, QHBoxLayout, QLabel, QLineEdit,
                                QPushButton, QVBoxLayout)

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

    def accept(self):
        self.settings.set("api_key", self.key_edit.text().strip())
        self.settings.set("base_url", self.url_edit.text().strip() or DEFAULT_BASE_URL)
        self.settings.set("model", self.model_edit.text().strip() or DEFAULT_MODEL)
        self.settings.set("stay_on_top", self.top_check.isChecked())
        self.settings.save()
        self.settings.apply_env()
        if self.worker is not None:
            self.worker.set_api_key(self.settings.get("api_key"))
        want_auto = self.auto_check.isChecked()
        if want_auto:
            enable()
        else:
            disable()
        self.settings.set("auto_start", want_auto)
        self.settings.save()
        super().accept()
