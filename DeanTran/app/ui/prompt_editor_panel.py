"""
Prompt Editor Panel – simple UI for editing translation prompts per document type.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QGroupBox, QLabel, QTextEdit, QPushButton, QComboBox,
    QMessageBox,
)

from app.storage.prompt_store import PromptStore
from app.settings.settings_manager import settings


class PromptEditorPanel(QWidget):
    """Panel for viewing/editing prompts per document type."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._store = PromptStore()
        self._build_ui()
        self._load_prompt()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        group = QGroupBox("Prompt Editor")
        inner = QVBoxLayout()

        # Mode selector
        row = QHBoxLayout()
        row.addWidget(QLabel("Document Type:"))
        self.cb_mode = QComboBox()
        self.cb_mode.addItems(self._store.modes())
        self.cb_mode.setEditable(True)

        # Set to current doc type from settings
        current = settings.get("document_settings.document_type", "SOP")
        idx = self.cb_mode.findText(current)
        if idx >= 0:
            self.cb_mode.setCurrentIndex(idx)
        else:
            self.cb_mode.setCurrentText(current)

        self.cb_mode.currentTextChanged.connect(self._load_prompt)
        row.addWidget(self.cb_mode, stretch=1)
        inner.addLayout(row)

        # Prompt text editor
        self.txt_prompt = QTextEdit()
        self.txt_prompt.setStyleSheet(
            "font-family: Consolas; font-size: 13px; "
            "background-color: #fafafa; border: 1px solid #ccc;"
        )
        self.txt_prompt.setMinimumHeight(200)
        inner.addWidget(self.txt_prompt)

        # Save / Reset buttons
        btn_row = QHBoxLayout()
        btn_save = QPushButton("💾 Save Prompt")
        btn_save.setStyleSheet("padding: 8px; font-weight: bold;")
        btn_save.clicked.connect(self._on_save)

        btn_reset = QPushButton("↩ Reset to Default")
        btn_reset.clicked.connect(self._on_reset)

        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_reset)
        btn_row.addStretch()
        inner.addLayout(btn_row)

        group.setLayout(inner)
        layout.addWidget(group)

    def _load_prompt(self, mode: str | None = None) -> None:
        if mode is None:
            mode = self.cb_mode.currentText()
        prompt = self._store.get(mode)
        self.txt_prompt.setPlainText(prompt)

    def _on_save(self) -> None:
        mode = self.cb_mode.currentText().strip()
        if not mode:
            QMessageBox.warning(self, "Error", "Document type cannot be empty.")
            return
        text = self.txt_prompt.toPlainText().strip()
        self._store.set(mode, text)
        self._store.save()

        # Refresh mode list if new mode added
        current_items = [self.cb_mode.itemText(i) for i in range(self.cb_mode.count())]
        if mode not in current_items:
            self.cb_mode.addItem(mode)

        QMessageBox.information(self, "Saved", f"Prompt for '{mode}' saved.")

    def _on_reset(self) -> None:
        mode = self.cb_mode.currentText()
        self._store.delete(mode)
        self._store.save()
        self._load_prompt(mode)
