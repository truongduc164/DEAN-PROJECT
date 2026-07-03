"""
PromptEditDialog – Quick inline dialog to edit a prompt for a specific mode.
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
    QPushButton, QMessageBox,
)

from app.storage.prompt_store import PromptStore


class PromptEditDialog(QDialog):
    """Popup dialog for quick editing of a single prompt."""

    def __init__(self, mode: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"✏️ Chỉnh sửa Prompt — {mode}")
        self.setMinimumSize(550, 380)
        self.resize(600, 400)
        self._mode = mode
        self._store = PromptStore()
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        header = QLabel(f"📝 Prompt cho lĩnh vực: <b>{self._mode}</b>")
        header.setStyleSheet("font-size: 14px; color: #1e40af; padding: 4px 0;")
        layout.addWidget(header)

        hint = QLabel("Hướng dẫn AI dịch theo đúng ngữ cảnh lĩnh vực của bạn.")
        hint.setStyleSheet("color: #64748b; font-size: 12px;")
        layout.addWidget(hint)

        self.txt_prompt = QTextEdit()
        self.txt_prompt.setPlainText(self._store.get(self._mode))
        self.txt_prompt.setStyleSheet(
            "font-family: 'Consolas', 'Cascadia Code', monospace; font-size: 13px; "
            "padding: 10px; border: 2px solid #e2e8f0; border-radius: 8px; "
            "background: #f8fafc; line-height: 1.5;"
        )
        layout.addWidget(self.txt_prompt, stretch=1)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        btn_reset = QPushButton("↩ Khôi phục mặc định")
        btn_reset.setStyleSheet(
            "padding: 8px 16px; border: 2px solid #e2e8f0; border-radius: 6px; "
            "color: #64748b; font-size: 13px;"
        )
        btn_reset.clicked.connect(self._on_reset)
        btn_row.addWidget(btn_reset)

        btn_row.addStretch()

        btn_cancel = QPushButton("Hủy")
        btn_cancel.setStyleSheet(
            "padding: 8px 20px; border: 2px solid #e2e8f0; border-radius: 6px; "
            "color: #374151; font-size: 13px;"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)

        btn_save = QPushButton("💾 Lưu")
        btn_save.setStyleSheet(
            "padding: 8px 24px; background-color: #3b82f6; color: white; "
            "border-radius: 6px; font-weight: bold; font-size: 13px; border: none;"
        )
        btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(btn_save)

        layout.addLayout(btn_row)

    def _on_save(self) -> None:
        text = self.txt_prompt.toPlainText().strip()
        self._store.set(self._mode, text)
        self._store.save()
        self.accept()

    def _on_reset(self) -> None:
        self._store.delete(self._mode)
        self._store.save()
        self.txt_prompt.setPlainText(self._store.get(self._mode))
