"""
Change Password Dialog – PySide6 modal dialog for changing user credentials.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QPushButton, QMessageBox,
)
from PySide6.QtCore import Qt

from app.core.auth_manager import AuthManager


class ChangePasswordDialog(QDialog):
    """
    Modal dialog to change password for the currently logged-in user.
    Returns ``QDialog.Accepted`` on success.
    """

    def __init__(self, auth_manager: AuthManager, parent=None) -> None:
        super().__init__(parent)
        self.auth = auth_manager
        self.setWindowTitle("Change Password")
        self.setFixedSize(350, 200)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)

        form = QFormLayout()
        
        self.txt_old_pass = QLineEdit()
        self.txt_old_pass.setPlaceholderText("Current Password")
        self.txt_old_pass.setEchoMode(QLineEdit.Password)
        
        self.txt_new_pass = QLineEdit()
        self.txt_new_pass.setPlaceholderText("New Password")
        self.txt_new_pass.setEchoMode(QLineEdit.Password)
        
        self.txt_confirm_pass = QLineEdit()
        self.txt_confirm_pass.setPlaceholderText("Confirm New Password")
        self.txt_confirm_pass.setEchoMode(QLineEdit.Password)

        form.addRow("Current:", self.txt_old_pass)
        form.addRow("New:", self.txt_new_pass)
        form.addRow("Confirm:", self.txt_confirm_pass)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        self.btn_change = QPushButton("Change Password")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_change.setDefault(True)
        btn_row.addWidget(self.btn_change)
        btn_row.addWidget(self.btn_cancel)
        layout.addLayout(btn_row)

        self.btn_change.clicked.connect(self._on_change)
        self.btn_cancel.clicked.connect(self.reject)
        self.txt_confirm_pass.returnPressed.connect(self._on_change)

    def _on_change(self) -> None:
        old_pwd = self.txt_old_pass.text()
        new_pwd = self.txt_new_pass.text()
        confirm_pwd = self.txt_confirm_pass.text()

        if not old_pwd or not new_pwd or not confirm_pwd:
            QMessageBox.warning(self, "Error", "Please fill in all fields.")
            return

        if new_pwd != confirm_pwd:
            QMessageBox.warning(self, "Error", "New password and confirmation do not match.")
            return

        current_user = self.auth.current_user
        if not current_user:
            QMessageBox.warning(self, "Error", "No user is currently logged in.")
            return

        # Attempt to verify current password by performing a theoretical login check
        # We can temporarily use the login method since they are already that user.
        if not self.auth.login(current_user, old_pwd):
            QMessageBox.warning(self, "Error", "Incorrect current password.")
            return

        # Change the password
        if self.auth.change_password(current_user, new_pwd):
            QMessageBox.information(self, "Success", "Password changed successfully!")
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "Failed to change password.")
