"""
TelegramConfigDialog – PySide6 dialog to configure the Telegram bot.

Allows the user to enter the bot token and allowed chat IDs,
test the connection, and save/reload the bot.
"""
from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QPushButton, QLabel, QMessageBox, QGroupBox, QTextEdit,
)
from PySide6.QtCore import Qt


class TelegramConfigDialog(QDialog):
    """Configure Telegram bot token and allowed chat IDs."""

    def __init__(self, bot, parent=None):
        super().__init__(parent)
        self._bot = bot
        self.setWindowTitle("🤖 Telegram Bot Configuration")
        self.setMinimumWidth(500)
        self._build_ui()
        self._load_current()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        # Info
        info = QLabel(
            "<b>Hướng dẫn:</b><br>"
            "1. Mở Telegram, tìm <b>@BotFather</b>, gửi <code>/newbot</code> để tạo bot<br>"
            "2. Copy <b>Bot Token</b> vào ô bên dưới<br>"
            "3. Gửi tin nhắn cho bot của bạn, rồi truy cập:<br>"
            "&nbsp;&nbsp;&nbsp;<code>https://api.telegram.org/bot&lt;TOKEN&gt;/getUpdates</code><br>"
            "&nbsp;&nbsp;&nbsp;để tìm <b>chat_id</b> của bạn<br>"
            "4. Nhập chat_id vào ô (nhiều ID cách nhau bằng dấu phẩy)"
        )
        info.setWordWrap(True)
        info.setTextFormat(Qt.RichText)
        info.setStyleSheet("background: #EEF2FF; padding: 10px; border-radius: 6px; font-size: 12px;")
        layout.addWidget(info)

        # Form
        group = QGroupBox("Bot Settings")
        form = QFormLayout()

        self.le_token = QLineEdit()
        self.le_token.setPlaceholderText("123456789:ABCdef...")
        self.le_token.setEchoMode(QLineEdit.Password)
        form.addRow("Bot Token:", self.le_token)

        self.le_chat_ids = QLineEdit()
        self.le_chat_ids.setPlaceholderText("123456789, 987654321")
        form.addRow("Chat IDs:", self.le_chat_ids)

        group.setLayout(form)
        layout.addWidget(group)

        # Status
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("font-size: 12px; padding: 4px;")
        layout.addWidget(self.lbl_status)

        # Commands info
        cmd_group = QGroupBox("Lệnh có thể gửi từ Telegram")
        cmd_layout = QVBoxLayout()
        cmd_text = QLabel(
            "<code>/ping</code> – Kiểm tra bot sống<br>"
            "<code>/status</code> – Thông tin máy + app<br>"
            "<code>/changepw admin Pass123</code> – Đổi mật khẩu<br>"
            "<code>/lock</code> – Khóa/tắt app từ xa<br>"
            "<code>/help</code> – Hiển thị trợ giúp"
        )
        cmd_text.setTextFormat(Qt.RichText)
        cmd_text.setWordWrap(True)
        cmd_layout.addWidget(cmd_text)
        cmd_group.setLayout(cmd_layout)
        layout.addWidget(cmd_group)

        # Buttons
        btn_row = QHBoxLayout()
        self.btn_test = QPushButton("🔗 Test Connection")
        self.btn_test.clicked.connect(self._on_test)
        btn_row.addWidget(self.btn_test)

        self.btn_save = QPushButton("💾 Save & Restart Bot")
        self.btn_save.setStyleSheet(
            "background-color: #3b82f6; color: white; font-weight: bold; padding: 8px 16px;"
        )
        self.btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(self.btn_save)

        self.btn_close = QPushButton("Close")
        self.btn_close.clicked.connect(self.close)
        btn_row.addWidget(self.btn_close)

        layout.addLayout(btn_row)

    def _load_current(self):
        """Load current config into the form."""
        self.le_token.setText(self._bot._token)
        if self._bot._allowed_ids:
            self.le_chat_ids.setText(", ".join(str(x) for x in self._bot._allowed_ids))

        if self._bot.is_configured:
            self.lbl_status.setText("✅ Bot đang hoạt động")
            self.lbl_status.setStyleSheet("color: green; font-weight: bold;")
        else:
            self.lbl_status.setText("⚠️ Bot chưa được cấu hình")
            self.lbl_status.setStyleSheet("color: orange; font-weight: bold;")

    def _parse_chat_ids(self) -> list[int]:
        text = self.le_chat_ids.text().strip()
        if not text:
            return []
        ids = []
        for part in text.replace(";", ",").split(","):
            part = part.strip()
            if part:
                try:
                    ids.append(int(part))
                except ValueError:
                    pass
        return ids

    def _on_test(self):
        """Test bot token by calling getMe."""
        token = self.le_token.text().strip()
        if not token:
            QMessageBox.warning(self, "Error", "Vui lòng nhập Bot Token!")
            return

        import requests
        try:
            resp = requests.get(
                f"https://api.telegram.org/bot{token}/getMe", timeout=10
            )
            data = resp.json()
            if data.get("ok"):
                bot_info = data["result"]
                name = bot_info.get("first_name", "?")
                username = bot_info.get("username", "?")
                QMessageBox.information(
                    self, "Success",
                    f"✅ Kết nối thành công!\n\n"
                    f"Bot name: {name}\n"
                    f"Username: @{username}",
                )
                self.lbl_status.setText(f"✅ Connected: @{username}")
                self.lbl_status.setStyleSheet("color: green; font-weight: bold;")
            else:
                QMessageBox.critical(
                    self, "Error",
                    f"❌ Token không hợp lệ!\n{data.get('description', '')}",
                )
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"❌ Lỗi kết nối: {exc}")

    def _on_save(self):
        """Save config and restart bot."""
        token = self.le_token.text().strip()
        chat_ids = self._parse_chat_ids()

        if not token:
            QMessageBox.warning(self, "Error", "Vui lòng nhập Bot Token!")
            return
        if not chat_ids:
            QMessageBox.warning(self, "Error", "Vui lòng nhập ít nhất 1 Chat ID!")
            return

        # Stop old bot
        self._bot.stop()

        # Save config
        self._bot._save_config(token, chat_ids)

        # Reload
        self._bot._token = token
        self._bot._allowed_ids = chat_ids
        self._bot._offset = 0

        # Start new bot
        started = self._bot.start()
        if started:
            self.lbl_status.setText("✅ Bot đã lưu và khởi động lại!")
            self.lbl_status.setStyleSheet("color: green; font-weight: bold;")
            QMessageBox.information(
                self, "Success",
                "✅ Đã lưu cấu hình và khởi động bot!\n\n"
                "Gửi /ping cho bot trên Telegram để kiểm tra.",
            )
        else:
            self.lbl_status.setText("❌ Không thể khởi động bot")
            self.lbl_status.setStyleSheet("color: red; font-weight: bold;")
