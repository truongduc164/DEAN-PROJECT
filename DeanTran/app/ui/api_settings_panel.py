"""
API Settings Panel – settings-driven, visible only when admin is logged in.

Model selector writes to SettingsManager.
API keys remain in SecureStorage (keyring) — never in JSON.
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QComboBox,
    QMessageBox, QSpinBox, QCheckBox, QDoubleSpinBox
)
from PySide6.QtCore import Signal

from app.core.secure_storage import SecureStorage
from app.settings.settings_manager import settings


_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-2.5-flash",
    "gemini-2.5-lite",
    "deepseek-v4-flash",
    "deepseek-v4-pro",
    "deepseek-chat",
    "deepseek-reasoner",
]


class ApiSettingsPanel(QWidget):
    """
    Admin-only panel for managing the Gemini API key and model.

    Model selection is persisted in SettingsManager (``selected_models.gemini``).
    API key remains in OS keyring via SecureStorage.
    """

    translator_changed = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.storage = SecureStorage()
        self._build_ui()
        self._refresh_display()

    def reload_from_settings(self) -> None:
        mode_saved = settings.get(
            "ocr_settings.ocr_textbox_placement_mode",
            settings.get("ocr_settings.ocr_textbox_mode", "whitespace"),
        )
        idx_mode = self.cb_ocr_textbox_mode.findData(mode_saved)
        self.cb_ocr_textbox_mode.setCurrentIndex(idx_mode if idx_mode >= 0 else 0)

    def _build_ui(self) -> None:
        group = QGroupBox("API Configuration (Admin)")
        layout = QVBoxLayout(self)
        form = QFormLayout()

        # API keys (masked display)
        from PySide6.QtWidgets import QListWidget
        self.list_keys = QListWidget()
        self.list_keys.setStyleSheet("font-family: Consolas;")
        
        btn_row = QHBoxLayout()
        btn_add = QPushButton("Add Key")
        btn_remove = QPushButton("Remove Selected")
        btn_add.clicked.connect(self._on_add_key)
        btn_remove.clicked.connect(self._on_remove_key)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_remove)
        btn_row.addStretch()
        
        self.lbl_key_stats = QLabel("Total keys: 0 | Active key: index 0")
        
        form.addRow("API Keys:", self.list_keys)
        form.addRow("", btn_row)
        form.addRow("", self.lbl_key_stats)

        # Model selector (reads/writes settings)
        self.cb_model = QComboBox()
        self.cb_model.addItems(_MODELS)
        self.cb_model.setEditable(True)

        # Load current model from settings
        tool = settings.get("translation_tool", "Gemini")
        if tool == "DeepSeek":
            current = settings.get("selected_models.deepseek", "deepseek-chat")
        else:
            current = settings.get("selected_models.gemini", "gemini-3.1-flash-lite")
            
        idx = self.cb_model.findText(current)
        if idx >= 0:
            self.cb_model.setCurrentIndex(idx)
        else:
            self.cb_model.setCurrentText(current)

        # Auto-save model changes to settings
        self.cb_model.currentTextChanged.connect(self._on_model_changed)
        form.addRow("Model:", self.cb_model)

        # Translation tool selector
        self.cb_tool = QComboBox()
        self.cb_tool.addItems(["Gemini", "DeepSeek"])
        self.cb_tool.setEditable(True)
        self.cb_tool.setCurrentText(settings.get("translation_tool", "Gemini"))
        self.cb_tool.currentTextChanged.connect(self._on_tool_changed)
        form.addRow("Tool:", self.cb_tool)

        # Dynamic Batch Configs
        self.sp_max_items = QSpinBox()
        self.sp_max_items.setRange(1, 1000)
        self.sp_max_items.setValue(settings.get("limits_settings.max_items_per_batch", 200))
        self.sp_max_items.valueChanged.connect(lambda v: settings.set("limits_settings.max_items_per_batch", v))
        form.addRow("Max Items per Batch:", self.sp_max_items)

        self.sp_max_chars = QSpinBox()
        self.sp_max_chars.setRange(1000, 1000000)
        self.sp_max_chars.setValue(settings.get("limits_settings.max_chars_per_batch", 30000))
        self.sp_max_chars.valueChanged.connect(lambda v: settings.set("limits_settings.max_chars_per_batch", v))
        form.addRow("Max Chars per Batch:", self.sp_max_chars)

        self.sp_retries = QSpinBox()
        self.sp_retries.setRange(0, 10)
        self.sp_retries.setValue(settings.get("limits_settings.retry_limit", 3))
        self.sp_retries.valueChanged.connect(lambda v: settings.set("limits_settings.retry_limit", v))
        form.addRow("Max Retries:", self.sp_retries)

        self.sp_timeout = QSpinBox()
        self.sp_timeout.setRange(10, 300)
        self.sp_timeout.setValue(settings.get("limits_settings.api_timeout", 60))
        self.sp_timeout.valueChanged.connect(lambda v: settings.set("limits_settings.api_timeout", v))
        form.addRow("API Timeout (s):", self.sp_timeout)

        self.sp_cb = QSpinBox()
        self.sp_cb.setRange(1, 100)
        self.sp_cb.setValue(settings.get("limits_settings.circuit_breaker_failures", 5))
        self.sp_cb.valueChanged.connect(lambda v: settings.set("limits_settings.circuit_breaker_failures", v))
        form.addRow("Circuit Breaker Failures:", self.sp_cb)

        # OCR Settings
        self.chk_ocr_enabled = QCheckBox("Enable OCR Translation")
        self.chk_ocr_enabled.setChecked(settings.get("ocr_settings.image_text_translation_enabled", False))
        self.chk_ocr_enabled.toggled.connect(lambda v: settings.set("ocr_settings.image_text_translation_enabled", v))
        form.addRow("", self.chk_ocr_enabled)

        self.chk_translate_ai = QCheckBox("Dịch AI (API) sau khi quét OCR")
        self.chk_translate_ai.setChecked(settings.get("ocr_settings.translate_with_api", True))
        self.chk_translate_ai.toggled.connect(lambda v: settings.set("ocr_settings.translate_with_api", v))
        form.addRow("", self.chk_translate_ai)

        self.cb_ocr_textbox_mode = QComboBox()
        self.cb_ocr_textbox_mode.addItem("Nearby whitespace (recommended)", "whitespace")
        self.cb_ocr_textbox_mode.addItem("Exact coordinates", "exact")
        self.cb_ocr_textbox_mode.addItem("Smart adjust", "smart_adjust")
        mode_saved = settings.get(
            "ocr_settings.ocr_textbox_placement_mode",
            settings.get("ocr_settings.ocr_textbox_mode", "whitespace"),
        )
        idx_mode = self.cb_ocr_textbox_mode.findData(mode_saved)
        self.cb_ocr_textbox_mode.setCurrentIndex(idx_mode if idx_mode >= 0 else 0)
        self.cb_ocr_textbox_mode.currentIndexChanged.connect(
            lambda _: (
                settings.set("ocr_settings.ocr_textbox_placement_mode", self.cb_ocr_textbox_mode.currentData()),
                settings.set("ocr_settings.ocr_textbox_mode", self.cb_ocr_textbox_mode.currentData()),
            )
        )
        form.addRow("OCR Placement Mode:", self.cb_ocr_textbox_mode)
        
        # We need a Key input for Google Vision
        self.le_vision_key = QLineEdit(settings.get("ocr_settings.google_vision_key", ""))
        self.le_vision_key.setEchoMode(QLineEdit.Password)
        self.le_vision_key.setPlaceholderText("API Key hoặc nguyên chuỗi JSON nội dung file Service Account")
        self.le_vision_key.editingFinished.connect(
            lambda: settings.set("ocr_settings.google_vision_key", self.le_vision_key.text().strip())
        )
        form.addRow("Google Vision Key:", self.le_vision_key)

        # DeepSeek API Key
        self.le_deepseek_key = QLineEdit(self.storage.load_deepseek_key())
        self.le_deepseek_key.setEchoMode(QLineEdit.Password)
        self.le_deepseek_key.setPlaceholderText("Nhập API Key của DeepSeek")
        self.le_deepseek_key.editingFinished.connect(
            lambda: self.storage.save_deepseek_key(self.le_deepseek_key.text().strip())
        )
        form.addRow("DeepSeek Key:", self.le_deepseek_key)

        self.chk_vision_batch = QCheckBox("Batch OCR Google Vision")
        self.chk_vision_batch.setChecked(settings.get("ocr_settings.google_vision_batch_enabled", True))
        self.chk_vision_batch.toggled.connect(
            lambda v: settings.set("ocr_settings.google_vision_batch_enabled", v)
        )
        form.addRow("", self.chk_vision_batch)

        self.sp_vision_batch_size = QSpinBox()
        self.sp_vision_batch_size.setRange(1, 16)
        self.sp_vision_batch_size.setValue(
            settings.get("ocr_settings.google_vision_max_images_per_request", 16)
        )
        self.sp_vision_batch_size.valueChanged.connect(
            lambda v: settings.set("ocr_settings.google_vision_max_images_per_request", v)
        )
        form.addRow("Vision images/request:", self.sp_vision_batch_size)

        self.chk_vision_dedupe = QCheckBox("Skip duplicate images before OCR")
        self.chk_vision_dedupe.setChecked(settings.get("ocr_settings.google_vision_dedupe_images", True))
        self.chk_vision_dedupe.toggled.connect(
            lambda v: settings.set("ocr_settings.google_vision_dedupe_images", v)
        )
        form.addRow("", self.chk_vision_dedupe)

        self.chk_vision_canvas = QCheckBox("Canvas OCR (merge many images into one request)")
        self.chk_vision_canvas.setChecked(settings.get("ocr_settings.google_vision_canvas_enabled", False))
        self.chk_vision_canvas.toggled.connect(
            lambda v: settings.set("ocr_settings.google_vision_canvas_enabled", v)
        )
        form.addRow("", self.chk_vision_canvas)

        self.sp_canvas_images = QSpinBox()
        self.sp_canvas_images.setRange(1, 16)
        self.sp_canvas_images.setValue(
            settings.get("ocr_settings.google_vision_canvas_images_per_canvas", 4)
        )
        self.sp_canvas_images.valueChanged.connect(
            lambda v: settings.set("ocr_settings.google_vision_canvas_images_per_canvas", v)
        )
        form.addRow("Images per canvas:", self.sp_canvas_images)

        self.sp_canvas_padding = QSpinBox()
        self.sp_canvas_padding.setRange(0, 200)
        self.sp_canvas_padding.setValue(settings.get("ocr_settings.google_vision_canvas_padding", 24))
        self.sp_canvas_padding.valueChanged.connect(
            lambda v: settings.set("ocr_settings.google_vision_canvas_padding", v)
        )
        form.addRow("Canvas padding(px):", self.sp_canvas_padding)

        from PySide6.QtWidgets import QDoubleSpinBox
        
        # OCR cleanup options
        self.chk_skip_noise = QCheckBox("Bỏ qua OCR Nhiễu (Skip noise blocks)")
        self.chk_skip_noise.setChecked(settings.get("ocr_settings.skip_noise_blocks", True))
        self.chk_skip_noise.toggled.connect(lambda v: settings.set("ocr_settings.skip_noise_blocks", v))
        form.addRow("", self.chk_skip_noise)
        
        self.chk_remove_overlays = QCheckBox("Xóa overlays chữ nhiễu (Remove noisy overlays)")
        self.chk_remove_overlays.setChecked(settings.get("ocr_settings.remove_noise_overlays_before_save", True))
        self.chk_remove_overlays.toggled.connect(lambda v: settings.set("ocr_settings.remove_noise_overlays_before_save", v))
        form.addRow("", self.chk_remove_overlays)

        self.sp_min_conf = QDoubleSpinBox()
        self.sp_min_conf.setRange(0.0, 100.0)
        self.sp_min_conf.setSingleStep(5.0)
        self.sp_min_conf.setDecimals(1)
        self.sp_min_conf.setValue(float(settings.get("ocr_settings.min_confidence", 0.0)))
        self.sp_min_conf.valueChanged.connect(lambda v: settings.set("ocr_settings.min_confidence", float(v)))
        form.addRow("Độ tin cậy OCR (Confidence %):", self.sp_min_conf)

        # Reinitialize button
        self.btn_reinit = QPushButton("Reinitialize Translator / Reload Settings")
        self.btn_reinit.setStyleSheet("padding: 8px; font-weight: bold;")
        self.btn_reinit.clicked.connect(self._on_reinit)

        inner = QVBoxLayout()
        inner.addLayout(form)
        inner.addWidget(self.btn_reinit)
        group.setLayout(inner)
        layout.addWidget(group)

    # ── slots ────────────────────────────────────────────────────────

    def _on_add_key(self) -> None:
        from PySide6.QtWidgets import QInputDialog
        keys_str, ok = QInputDialog.getMultiLineText(
            self, "Add API Key(s)",
            "Enter new Gemini API key(s) (one per line):",
        )
        if ok and keys_str.strip():
            current_keys = self.storage.load_api_keys()
            for line in keys_str.splitlines():
                k = line.strip()
                if k and k not in current_keys:
                    current_keys.append(k)
            self.storage.save_api_keys(current_keys)
            self._refresh_display()

    def _on_remove_key(self) -> None:
        row = self.list_keys.currentRow()
        if row >= 0:
            keys = self.storage.load_api_keys()
            if row < len(keys):
                keys.pop(row)
                self.storage.save_api_keys(keys)
                self._refresh_display()

    def _on_model_changed(self, text: str) -> None:
        tool = self.cb_tool.currentText()
        if tool == "DeepSeek" or text.startswith("deepseek"):
            settings.set("selected_models.deepseek", text)
        else:
            settings.set("selected_models.gemini", text)

    def _on_tool_changed(self, tool: str) -> None:
        settings.set("translation_tool", tool)
        if tool == "DeepSeek":
            current = settings.get("selected_models.deepseek", "deepseek-chat")
        else:
            current = settings.get("selected_models.gemini", "gemini-3.1-flash-lite")
        self.cb_model.blockSignals(True)
        idx = self.cb_model.findText(current)
        if idx >= 0:
            self.cb_model.setCurrentIndex(idx)
        else:
            self.cb_model.setCurrentText(current)
        self.cb_model.blockSignals(False)

    def _on_reinit(self) -> None:
        self.translator_changed.emit()
        QMessageBox.information(self, "Done", "Translator reinitialized.")

    def _refresh_display(self) -> None:
        self.list_keys.clear()
        keys = self.storage.load_api_keys()
        for k in keys:
            self.list_keys.addItem(SecureStorage.masked(k))
            
        # Display stats
        self.lbl_key_stats.setText(f"Total keys: {len(keys)} | Active key: index 0")
        
        # Refresh DeepSeek key field text
        if hasattr(self, "le_deepseek_key"):
            self.le_deepseek_key.setText(self.storage.load_deepseek_key())
