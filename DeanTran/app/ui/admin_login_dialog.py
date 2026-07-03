"""
Admin Login Dialog – PySide6 modal dialog for entering admin password.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox,
)
from PySide6.QtCore import Qt

from app.core.auth_manager import AuthManager


class AdminLoginDialog(QDialog):
    """
    Modal login dialog (password only, username = 'admin').

    After a successful login the caller can inspect:
      * ``self.auth_manager.is_admin``
      * ``self.auth_manager.current_user``
    """

    def __init__(self, auth_manager: AuthManager, parent=None) -> None:
        super().__init__(parent)
        self.auth = auth_manager
        self.setWindowTitle("🔐 Admin Login")
        self.setFixedSize(320, 140)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)

        form = QFormLayout()
        self.txt_pass = QLineEdit()
        self.txt_pass.setPlaceholderText("Nhập mật khẩu admin")
        self.txt_pass.setEchoMode(QLineEdit.Password)
        form.addRow("Password:", self.txt_pass)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        self.btn_login = QPushButton("Login")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_login.setDefault(True)
        btn_row.addWidget(self.btn_login)
        btn_row.addWidget(self.btn_cancel)
        layout.addLayout(btn_row)

        self.btn_login.clicked.connect(self._on_login)
        self.btn_cancel.clicked.connect(self.reject)
        self.txt_pass.returnPressed.connect(self._on_login)

    def _on_login(self) -> None:
        pwd = self.txt_pass.text()
        if not pwd:
            QMessageBox.warning(self, "Error", "Vui lòng nhập mật khẩu.")
            return
        if self.auth.login("admin", pwd):
            self.accept()
        else:
            QMessageBox.warning(self, "Login Failed", "❌ Sai mật khẩu.")
