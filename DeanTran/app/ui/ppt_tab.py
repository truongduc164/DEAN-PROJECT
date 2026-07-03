"""
PptTab – Minimal log-centric UI for PowerPoint translation.

Same layout as ExcelTab: compact settings row, file list, progress, log panel.
"""
from __future__ import annotations

import logging
import re
import threading
import time
import traceback
from datetime import datetime
from functools import partial
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QProgressBar, QTextEdit, QComboBox, QCheckBox,
    QFileDialog, QMessageBox, QListWidget,
    QListWidgetItem, QAbstractItemView, QFrame, QDialog,
    QGroupBox, QFormLayout, QSpinBox, QDoubleSpinBox, QLineEdit, QRadioButton, QColorDialog,
)
from PySide6.QtCore import Qt, QThread, Signal, QObject, QTimer
from PySide6.QtGui import QColor

from app.core.signals import signals
from app.settings.settings_manager import settings

logger = logging.getLogger("DeanTran.ppt_tab")

_LANGUAGES = ["Chinese", "English", "Vietnamese"]
_LANG_DISPLAY = {
    "Chinese": "Chinese (中文)",
    "English": "English",
    "Vietnamese": "Vietnamese (Tiếng Việt)",
}
_LANG_REVERSE = {v: k for k, v in _LANG_DISPLAY.items()}

_OUTPUT_MODES = {
    "Ghi đè (Overwrite)": "overwrite",
    "Dịch trước nguồn (Prefix)": "prefix",
    "Dịch sau nguồn (Suffix)": "suffix",
}
_OUTPUT_REVERSE = {v: k for k, v in _OUTPUT_MODES.items()}


# ══════════════════════════════════════════════════════════════════════
# PPT Translation Worker
# ══════════════════════════════════════════════════════════════════════

class PptWorker(QObject):
    log = Signal(str, str)
    progress = Signal(int, int)
    finished = Signal(str)
    error = Signal(str)
    result_info = Signal(dict)
    file_started = Signal(str)  # filename being processed

    def __init__(self, file_paths: list[str], pause_event, cancel_event):
        super().__init__()
        self.file_paths = file_paths
        self._pause = pause_event
        self._cancel = cancel_event

    def run(self):
        import time
        t_worker_start = time.time()
        try:
            from app.core.translators.translator_service import create_translator
            from app.core.event_manager import EventManager
            from app.core.ppt_processor import PptProcessor
            from app.core.key_provider import KeyProvider

            model = settings.get("selected_models.gemini", "gemini-3.1-flash-lite")
            source = settings.get("language_settings.source_lang", "Chinese")
            target = settings.get("language_settings.target_lang", "Vietnamese")
            output_mode = settings.get("document_settings.output_mode", "overwrite")
            translate_tb = settings.get("processing_options.translate_textboxes", True)
            check_glossary = settings.get("processing_options.check_glossary_before_ai", True)

            self.log.emit("INFO", f"[config] model={model} {source}->{target} mode={output_mode} textboxes={translate_tb}")

            kp = KeyProvider()
            key = kp.get_key()
            key_prefix = key[:6] + "****" if key and len(key) >= 6 else "(none)"
            self.log.emit("INFO", f"[key] prefix={key_prefix}")

            translator = create_translator(model_name=model)
            self.log.emit("INFO", f"[translator] {type(translator).__name__}")

            aggregate = {
                "status": "SUCCESS", "translated": 0, "failed": 0,
                "elapsed": 0.0, "output_paths": [],
            }

            total_files = len(self.file_paths)
            for idx, fp in enumerate(self.file_paths, 1):
                if self._cancel.is_set():
                    self.log.emit("WARN", "Cancel requested – skipping remaining files.")
                    break

                fname = Path(fp).name
                self.file_started.emit(fname)
                self.log.emit("INFO", f"═══ File {idx}/{total_files}: {fname} ═══")

                em = EventManager()
                em.subscribe("log", lambda lvl, msg: self.log.emit(lvl, msg))
                em.subscribe("progress", lambda cur, tot: self.progress.emit(cur, tot))

                try:
                    from app.term_engine.glossary_loader import GlossaryLoader
                    glossary = None
                    if check_glossary:
                        glossary_path = settings.get("processing_options.glossary_file_path", "")
                        if glossary_path and Path(glossary_path).exists():
                            glossary = GlossaryLoader(glossary_path, source_lang=source, target_lang=target)

                    proc = PptProcessor(
                        glossary=glossary,
                        use_glossary=check_glossary,
                        translator=translator,
                        event_manager=em,
                        source_lang=source,
                        target_lang=target,
                        output_mode=output_mode,
                        translate_textboxes=translate_tb,
                        pause_event=self._pause,
                        cancel_event=self._cancel,
                    )

                    custom_dir = None
                    if settings.get("ppt_settings.use_custom_dir", False):
                        custom_dir_path = settings.get("ppt_settings.custom_dir_path", "")
                        if custom_dir_path:
                            custom_dir = Path(custom_dir_path)
                            custom_dir.mkdir(parents=True, exist_ok=True)

                    out = proc.process(fp, output_dir=custom_dir)
                    r = proc.last_result
                    aggregate["translated"] += r.get("translated", 0)
                    aggregate["failed"] += r.get("failed", 0)
                    aggregate["elapsed"] += r.get("elapsed", 0.0)
                    aggregate["output_paths"].append(str(out))

                    s = r.get("status", "SUCCESS")
                    if s == "FAILED":
                        aggregate["status"] = "FAILED"
                    elif s in ("PARTIAL", "CANCELLED") and aggregate["status"] == "SUCCESS":
                        aggregate["status"] = s

                    self.log.emit("INFO", f"File {idx}/{total_files} done: {Path(out).name}")

                except Exception as file_exc:
                    tb = traceback.format_exc()
                    self.log.emit("ERROR", f"File {fname} failed: {file_exc}")
                    logger.error("PPT file %s error:\n%s", fname, tb)
                    aggregate["failed"] += 1
                    if aggregate["status"] == "SUCCESS":
                        aggregate["status"] = "PARTIAL"

            if self._cancel.is_set() and aggregate["status"] == "SUCCESS":
                aggregate["status"] = "CANCELLED"

            aggregate["elapsed"] = time.time() - t_worker_start
            self.result_info.emit(aggregate)
            last = aggregate["output_paths"][-1] if aggregate["output_paths"] else ""
            self.finished.emit(last)

        except Exception as exc:
            tb = traceback.format_exc()
            self.log.emit("ERROR", f"PPT translation failed: {exc}")
            logger.error("PPT worker exception:\n%s", tb)
            self.result_info.emit({"status": "FAILED", "error": str(exc)})
            self.finished.emit("")


# ══════════════════════════════════════════════════════════════════════
# PptTab
# ══════════════════════════════════════════════════════════════════════

class PptTab(QWidget):
    def __init__(self):
        super().__init__()
        self._thread = None
        self._worker = None
        self._pause = threading.Event()
        self._pause.set()
        self._cancel = threading.Event()
        self._files: list[str] = []
        self._last_result = {}

        self._build_ui()
        self._load_settings()

        self._clock = QTimer(self)
        self._clock.timeout.connect(self._tick)
        self._clock.start(1000)
        self._tick()

    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(12)

        # ── Left Panel ─────────────────────────────────────────────
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)
        left_widget.setMaximumWidth(450)

        # 1. Nhập File
        left_layout.addWidget(QLabel("<b>1. Nhập File PPT:</b>"))
        row_files_btn = QHBoxLayout()
        self.btn_add = QPushButton("Add PPT")
        self.btn_add.clicked.connect(self._add_files)
        self.btn_rm = QPushButton("Remove")
        self.btn_rm.clicked.connect(self._remove)
        self.btn_clr = QPushButton("Clear")
        self.btn_clr.clicked.connect(self._clear)
        row_files_btn.addWidget(self.btn_add)
        row_files_btn.addWidget(self.btn_rm)
        row_files_btn.addWidget(self.btn_clr)
        row_files_btn.addStretch()
        left_layout.addLayout(row_files_btn)

        self.lst = QListWidget()
        self.lst.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.lst.setMaximumHeight(100)
        self.lst.setStyleSheet("font-size: 12px;")
        left_layout.addWidget(self.lst)

        # 2. Tính năng & Options
        options_group = QGroupBox("2. Chọn tính năng")
        options_form = QFormLayout()
        self.options_form = options_form
        options_form.setContentsMargins(8, 4, 8, 4)

        self.cb_src = QComboBox()
        self.cb_src.addItems([_LANG_DISPLAY[l] for l in _LANGUAGES])
        options_form.addRow("Source:", self.cb_src)

        self.cb_tgt = QComboBox()
        self.cb_tgt.addItems([_LANG_DISPLAY[l] for l in _LANGUAGES])
        options_form.addRow("Target:", self.cb_tgt)

        self.cb_mode = QComboBox()
        self.cb_mode.addItems(list(_OUTPUT_MODES.keys()))
        options_form.addRow("Output:", self.cb_mode)

        # Prompt selector
        prompt_row = QHBoxLayout()
        self.cb_prompt = QComboBox()
        self.cb_prompt.setEditable(True)
        self._refresh_prompt_list()
        self.cb_prompt.currentTextChanged.connect(
            lambda t: settings.set("document_settings.document_type", t)
        )
        self.btn_edit_prompt = QPushButton("✏️")
        self.btn_edit_prompt.setMaximumWidth(32)
        self.btn_edit_prompt.setToolTip("Chỉnh sửa prompt")
        self.btn_edit_prompt.clicked.connect(self._on_edit_prompt)
        prompt_row.addWidget(self.cb_prompt, stretch=1)
        prompt_row.addWidget(self.btn_edit_prompt)
        options_form.addRow("Prompt:", prompt_row)

        self.chk_textbox = QCheckBox("Dịch TextBox")
        self.chk_textbox.setChecked(True)
        options_form.addRow("", self.chk_textbox)

        self.chk_ocr = QCheckBox("Dịch Ảnh (OCR)")
        self.chk_ocr.setChecked(settings.get("ocr_settings.image_text_translation_enabled", False))
        self.chk_ocr.stateChanged.connect(self._on_ocr_toggled)
        options_form.addRow("", self.chk_ocr)

        self.cb_ocr_engine = QComboBox()
        self.cb_ocr_engine.addItems(["paddle", "google_vision"])
        self.cb_ocr_engine.setCurrentText(settings.get("ocr_settings.engine", "paddle"))
        self.cb_ocr_engine.currentTextChanged.connect(lambda t: settings.set("ocr_settings.engine", t))
        options_form.addRow("OCR Engine:", self.cb_ocr_engine)

        self.chk_auto_open = QCheckBox("Tự động mở file")
        self.chk_auto_open.setChecked(settings.get("processing_options.auto_open_file", False))
        self.chk_auto_open.stateChanged.connect(
            lambda s: settings.set("processing_options.auto_open_file", s == Qt.Checked.value)
        )
        options_form.addRow("", self.chk_auto_open)

        # Custom output directory
        self.chk_custom_dir = QCheckBox("Lưu vào thư mục khác (Custom Output)")
        self.chk_custom_dir.stateChanged.connect(self._on_custom_dir_toggled)

        custom_dir_row = QHBoxLayout()
        self.le_custom_dir = QLineEdit()
        self.le_custom_dir.setPlaceholderText("Mặc định: cùng thư mục file gốc")
        self.le_custom_dir.setReadOnly(True)
        self.btn_browse_custom_dir = QPushButton("Browse...")
        self.btn_browse_custom_dir.clicked.connect(self._browse_custom_dir)
        custom_dir_row.addWidget(self.le_custom_dir)
        custom_dir_row.addWidget(self.btn_browse_custom_dir)

        options_form.addRow(self.chk_custom_dir)
        options_form.addRow("Thư mục lưu:", custom_dir_row)

        self.chk_glossary = QCheckBox("Lọc từ điển trước (Glossary)")
        self.chk_glossary.setChecked(settings.get("processing_options.check_glossary_before_ai", True))
        self.chk_glossary.stateChanged.connect(self._on_glossary_toggled)
        options_form.addRow("", self.chk_glossary)

        glossary_row = QHBoxLayout()
        self.le_glossary_path = QLineEdit(settings.get("processing_options.glossary_file_path", ""))
        self.le_glossary_path.setPlaceholderText("File từ điển (.xlsx, .json)")
        self.le_glossary_path.editingFinished.connect(
            lambda: settings.set("processing_options.glossary_file_path", self.le_glossary_path.text().strip())
        )
        self.btn_browse_glossary = QPushButton("Browse...")
        self.btn_browse_glossary.clicked.connect(self._browse_glossary)
        self.btn_view_glossary = QPushButton("👁")
        self.btn_view_glossary.setMaximumWidth(32)
        self.btn_view_glossary.setToolTip("Xem từ điển")
        self.btn_view_glossary.clicked.connect(self._view_glossary)
        glossary_row.addWidget(self.le_glossary_path)
        glossary_row.addWidget(self.btn_browse_glossary)
        glossary_row.addWidget(self.btn_view_glossary)
        options_form.addRow("Từ điển:", glossary_row)

        self.chk_pinyin = QCheckBox("Show Pinyin below Chinese text")
        self.chk_pinyin.setChecked(settings.get("processing_options.add_chinese_pinyin", False))
        self.chk_pinyin.stateChanged.connect(
            lambda s: settings.set("processing_options.add_chinese_pinyin", s == Qt.Checked.value)
        )
        options_form.addRow("", self.chk_pinyin)

        pinyin_style_row = QHBoxLayout()
        pinyin_style_row.addWidget(QLabel("Pinyin:"))
        self.cb_pinyin_font = QComboBox()
        self.cb_pinyin_font.setEditable(True)
        self.cb_pinyin_font.addItems(["Arial", "Times New Roman", "Calibri", "Tahoma"])
        self.cb_pinyin_font.setCurrentText(settings.get("processing_options.pinyin_font_family", "Arial"))
        self.cb_pinyin_font.setMaximumWidth(120)
        self.cb_pinyin_font.setToolTip("Pinyin font family")
        pinyin_style_row.addWidget(self.cb_pinyin_font)
        self.sp_pinyin_size = QSpinBox()
        self.sp_pinyin_size.setRange(6, 48)
        self.sp_pinyin_size.setValue(settings.get("processing_options.pinyin_font_size", 10))
        self.sp_pinyin_size.setToolTip("Pinyin font size (pt)")
        self.sp_pinyin_size.setMaximumWidth(55)
        pinyin_style_row.addWidget(self.sp_pinyin_size)
        self.le_pinyin_color = QLineEdit(settings.get("processing_options.pinyin_font_color", "#888888"))
        self.le_pinyin_color.setMaximumWidth(75)
        self.le_pinyin_color.setToolTip("Pinyin font color (hex)")
        pinyin_style_row.addWidget(self.le_pinyin_color)
        self.btn_pinyin_color = QPushButton("🎨")
        self.btn_pinyin_color.setMaximumWidth(28)
        self.btn_pinyin_color.setToolTip("Pick pinyin color")
        self.btn_pinyin_color.clicked.connect(self._pick_pinyin_color)
        pinyin_style_row.addWidget(self.btn_pinyin_color)
        self.chk_pinyin_keep_format = QCheckBox("Keep original format")
        self.chk_pinyin_keep_format.setChecked(settings.get("processing_options.pinyin_format_mode", "custom") == "keep_original")
        self.chk_pinyin_keep_format.setToolTip("Keep original font/size/color for pinyin")
        pinyin_style_row.addWidget(self.chk_pinyin_keep_format)
        pinyin_style_row.addStretch()
        options_form.addRow("", pinyin_style_row)
        
        options_group.setLayout(options_form)
        left_layout.addWidget(options_group)

        # 3. Output Text Style
        self.grp_style = QGroupBox("3. Chọn Font & Format")
        style_layout = QVBoxLayout()
        style_layout.setContentsMargins(8, 4, 8, 4)

        self.lbl_place = QLabel("OCR Placement")
        self.lbl_place.setStyleSheet("font-weight: bold;")
        style_layout.addWidget(self.lbl_place)

        self.ocr_form = QFormLayout()
        self.ocr_form.setContentsMargins(0, 0, 0, 8)
        self.cb_ocr_display_mode = QComboBox()
        self.cb_ocr_display_mode.addItems(["Chỉ hiển thị chữ dịch", "Chữ dịch song song chữ gốc", "Chữ gốc song song chữ dịch"])
        # Map: Chỉ hiển thị chữ dịch -> overwrite, song song -> prefix/suffix
        ocr_mode_saved = settings.get("ocr_settings.ocr_display_mode", "overwrite")
        if ocr_mode_saved == "prefix":
            self.cb_ocr_display_mode.setCurrentIndex(1)
        elif ocr_mode_saved == "suffix":
            self.cb_ocr_display_mode.setCurrentIndex(2)
        else:
            self.cb_ocr_display_mode.setCurrentIndex(0)
            
        self.cb_ocr_display_container = QComboBox()
        self.cb_ocr_display_container.addItems(["Khung trong suốt (Text Box)", "Khối nền che chữ gốc (Shape)"])
        ocr_container_saved = settings.get("ocr_settings.ocr_display_container", "textbox")
        self.cb_ocr_display_container.setCurrentIndex(1 if ocr_container_saved == "shape" else 0)
        self.cb_ocr_placement_mode = QComboBox()
        self.cb_ocr_placement_mode.addItem("Whitespace", "whitespace")
        self.cb_ocr_placement_mode.addItem("Exact", "exact")
        self.cb_ocr_placement_mode.addItem("Smart adjust", "smart_adjust")
        placement_saved = settings.get(
            "ocr_settings.ocr_textbox_placement_mode",
            settings.get("ocr_settings.ocr_textbox_mode", "whitespace"),
        )
        placement_idx = self.cb_ocr_placement_mode.findData(placement_saved)
        self.cb_ocr_placement_mode.setCurrentIndex(placement_idx if placement_idx >= 0 else 0)
        self.cb_ocr_placement_mode.currentIndexChanged.connect(self._on_ocr_placement_changed)
        
        self.ocr_form.addRow("Mode dịch ảnh:", self.cb_ocr_display_mode)
        self.ocr_form.addRow("Đè ảnh bằng:", self.cb_ocr_display_container)
        self.ocr_form.addRow("OCR Placement:", self.cb_ocr_placement_mode)
        style_layout.addLayout(self.ocr_form)

        self.line_ocr_divider = QFrame()
        self.line_ocr_divider.setFrameShape(QFrame.HLine)
        self.line_ocr_divider.setFrameShadow(QFrame.Sunken)
        style_layout.addWidget(self.line_ocr_divider)

        lbl_fmt = QLabel("Translated Text Format")
        lbl_fmt.setStyleSheet("font-weight: bold;")
        style_layout.addWidget(lbl_fmt)

        self.rbtn_keep_format = QRadioButton("Giữ nguyên định dạng gốc")
        self.rbtn_keep_format.setChecked(settings.get("text_style_settings.keep_format", True))
        self.rbtn_keep_format.toggled.connect(self._toggle_custom_format)
        style_layout.addWidget(self.rbtn_keep_format)

        self.rbtn_custom_format = QRadioButton("Tùy chỉnh định dạng văn bản đích")
        self.rbtn_custom_format.setChecked(not self.rbtn_keep_format.isChecked())
        style_layout.addWidget(self.rbtn_custom_format)

        self.custom_format_widget = QWidget()
        custom_layout = QFormLayout(self.custom_format_widget)
        custom_layout.setContentsMargins(16, 0, 0, 0)

        self.cb_font = QComboBox()
        self.cb_font.setEditable(True)
        self.cb_font.addItems(["Mặc định", "Arial", "Times New Roman", "Calibri", "Tahoma"])
        self.cb_font.setCurrentText(settings.get("text_style_settings.font_family", "Arial"))
        custom_layout.addRow("Font:", self.cb_font)

        size_row = QHBoxLayout()
        self.chk_default_size = QCheckBox("Mặc định")
        self.sp_size = QSpinBox()
        self.sp_size.setRange(8, 72)
        self.sp_size.setValue(settings.get("text_style_settings.font_size", 14))
        self.chk_default_size.toggled.connect(self._on_default_size_toggled)
        size_row.addWidget(self.chk_default_size)
        size_row.addWidget(self.sp_size)
        size_row.addStretch()
        custom_layout.addRow("Size:", size_row)

        color_row = QHBoxLayout()
        self.chk_default_color = QCheckBox("Mặc định")
        self.le_color = QLineEdit(settings.get("text_style_settings.font_color", "#000000"))
        self.le_color.setMaximumWidth(90)
        self.le_color.editingFinished.connect(self._on_color_text_changed)

        self.btn_color_pick = QPushButton("...")
        self.btn_color_pick.setMaximumWidth(32)
        self.btn_color_pick.setToolTip("Pick color")
        self.btn_color_pick.clicked.connect(self._pick_color)

        self.btn_color_preview = QPushButton("")
        self.btn_color_preview.setMaximumWidth(32)
        self.btn_color_preview.setEnabled(False)

        self.chk_default_color.toggled.connect(self._on_default_color_toggled)

        color_row.addWidget(self.chk_default_color)
        color_row.addWidget(self.le_color)
        color_row.addWidget(self.btn_color_pick)
        color_row.addWidget(self.btn_color_preview)
        color_row.addStretch()
        custom_layout.addRow("Color:", color_row)

        self.chk_bold = QCheckBox("In đậm (Bold)")
        self.chk_bold.setTristate(True)
        custom_layout.addRow("", self.chk_bold)

        self.chk_italic = QCheckBox("In nghiêng (Italic)")
        self.chk_italic.setTristate(True)
        custom_layout.addRow("", self.chk_italic)

        self.chk_underline = QCheckBox("Gạch chân (Underline)")
        self.chk_underline.setTristate(True)
        custom_layout.addRow("", self.chk_underline)

        style_layout.addWidget(self.custom_format_widget)
        self._toggle_custom_format()

        self.grp_style.setLayout(style_layout)
        left_layout.addWidget(self.grp_style)

        left_layout.addStretch()
        root.addWidget(left_widget)

        # ── Right Panel ────────────────────────────────────────────
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        # Time & warn
        row_time = QHBoxLayout()
        self.lbl_time = QLabel("")
        self.lbl_time.setStyleSheet("color: #888; font-size: 12px;")
        self.lbl_warn = QLabel("")
        self.lbl_warn.setStyleSheet("color: red; font-weight: bold;")
        row_time.addWidget(self.lbl_time)
        row_time.addStretch()
        row_time.addWidget(self.lbl_warn)
        right_layout.addLayout(row_time)

        # Progress
        self.bar = QProgressBar()
        self.bar.setFormat("%p% (%v / %m)")
        right_layout.addWidget(self.bar)

        # Log panel
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet(
            "background-color: #1e1e2e; color: #cdd6f4; "
            "font-family: 'Cascadia Code', 'Consolas', monospace; "
            "font-size: 12px; padding: 6px; border-radius: 4px;"
        )
        right_layout.addWidget(self.txt_log, stretch=7)

        # Status
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px;")
        right_layout.addWidget(self.lbl_status)

        # Buttons
        row_btns = QHBoxLayout()
        self.btn_translate = QPushButton("Translate PPT")
        self.btn_translate.setStyleSheet(
            "padding: 10px 20px; font-weight: bold; font-size: 14px; "
            "background-color: #89b4fa; color: #1e1e2e; border-radius: 6px;"
        )
        self.btn_translate.clicked.connect(self._translate)

        self.btn_pause = QPushButton("Pause")
        self.btn_resume = QPushButton("Resume")
        self.btn_cancel = QPushButton("Cancel")
        for b in (self.btn_pause, self.btn_resume, self.btn_cancel):
            b.setStyleSheet("padding: 10px 14px; font-weight: bold; font-size: 13px;")
            b.setEnabled(False)

        self.btn_pause.clicked.connect(self._on_pause)
        self.btn_resume.clicked.connect(self._on_resume)
        self.btn_cancel.clicked.connect(self._on_cancel)

        row_btns.addWidget(self.btn_translate)
        row_btns.addWidget(self.btn_pause)
        row_btns.addWidget(self.btn_resume)
        row_btns.addWidget(self.btn_cancel)
        row_btns.addStretch()
        right_layout.addLayout(row_btns)

        root.addWidget(right_widget, stretch=1)

    def _toggle_custom_format(self):
        self.custom_format_widget.setEnabled(self.rbtn_custom_format.isChecked())

    @staticmethod
    def _set_form_row_visible(form_layout: QFormLayout, field_widget: QWidget, visible: bool):
        label = form_layout.labelForField(field_widget)
        if label is not None:
            label.setVisible(visible)
        field_widget.setVisible(visible)

    def _update_ocr_ui_visibility(self, enabled: bool | None = None):
        if enabled is None:
            enabled = self.chk_ocr.isChecked()
        self._set_form_row_visible(self.options_form, self.cb_ocr_engine, enabled)
        self.lbl_place.setVisible(enabled)
        self._set_form_row_visible(self.ocr_form, self.cb_ocr_display_mode, enabled)
        self._set_form_row_visible(self.ocr_form, self.cb_ocr_display_container, enabled)
        self._set_form_row_visible(self.ocr_form, self.cb_ocr_placement_mode, enabled)
        self.line_ocr_divider.setVisible(enabled)

    def _on_ocr_toggled(self, state):
        enabled = state == Qt.Checked.value
        settings.set("ocr_settings.image_text_translation_enabled", enabled)
        self._update_ocr_ui_visibility(enabled)

    def _update_glossary_ui_visibility(self, enabled: bool | None = None):
        if enabled is None:
            enabled = self.chk_glossary.isChecked()
        self.le_glossary_path.setEnabled(enabled)
        self.btn_browse_glossary.setEnabled(enabled)
        self.btn_view_glossary.setEnabled(enabled)

    def _on_glossary_toggled(self, state):
        enabled = state == Qt.Checked.value
        settings.set("processing_options.check_glossary_before_ai", enabled)
        self._update_glossary_ui_visibility(enabled)

    def _browse_glossary(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Select Glossary File", "", "Excel/JSON (*.xlsx *.xls *.json)"
        )
        if fname:
            self.le_glossary_path.setText(fname)
            settings.set("processing_options.glossary_file_path", fname)

    def _view_glossary(self):
        path = self.le_glossary_path.text().strip()
        if not path:
            QMessageBox.information(self, "Thông báo", "Chưa chọn file từ điển.")
            return
        from app.ui.dialogs.glossary_viewer_dialog import GlossaryViewerDialog
        dlg = GlossaryViewerDialog(path, parent=self)
        dlg.exec()

    def _on_custom_dir_toggled(self, state=None):
        enabled = self.chk_custom_dir.isChecked()
        settings.set("ppt_settings.use_custom_dir", enabled)
        self.le_custom_dir.setEnabled(enabled)
        self.btn_browse_custom_dir.setEnabled(enabled)

    def _browse_custom_dir(self):
        curr_dir = self.le_custom_dir.text().strip() or settings.get("file_picker.last_export_dir", "")
        dir_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu kết quả", curr_dir)
        if dir_path:
            self.le_custom_dir.setText(dir_path)
            settings.set("ppt_settings.custom_dir_path", dir_path)

    def _refresh_prompt_list(self):
        from app.storage.prompt_store import PromptStore
        store = PromptStore()
        modes = store.modes()
        current = settings.get("document_settings.document_type", "General")
        self.cb_prompt.blockSignals(True)
        self.cb_prompt.clear()
        self.cb_prompt.addItems(modes)
        idx = self.cb_prompt.findText(current)
        if idx >= 0:
            self.cb_prompt.setCurrentIndex(idx)
        else:
            self.cb_prompt.setCurrentText(current)
        self.cb_prompt.blockSignals(False)

    def _on_edit_prompt(self):
        mode = self.cb_prompt.currentText().strip()
        if not mode:
            return
        from app.ui.dialogs.prompt_edit_dialog import PromptEditDialog
        dlg = PromptEditDialog(mode, parent=self)
        if dlg.exec():
            self._refresh_prompt_list()

    def _on_ocr_placement_changed(self):
        placement_mode = self.cb_ocr_placement_mode.currentData()
        if not placement_mode:
            return
        settings.set("ocr_settings.ocr_textbox_placement_mode", placement_mode)
        settings.set("ocr_settings.ocr_textbox_mode", placement_mode)

    def _normalize_hex_color(self, value: str) -> str:
        text = (value or "").strip().upper()
        if not text:
            return "#000000"
        if not text.startswith("#"):
            text = f"#{text}"
        if re.fullmatch(r"#[0-9A-F]{6}", text):
            return text
        return "#000000"

    def _refresh_color_preview(self):
        color = self._normalize_hex_color(self.le_color.text())
        self.btn_color_preview.setStyleSheet(
            f"background-color: {color}; border: 1px solid #888; border-radius: 3px;"
        )

    def _on_color_text_changed(self):
        normalized = self._normalize_hex_color(self.le_color.text())
        self.le_color.setText(normalized)
        self._refresh_color_preview()

    def _on_default_color_toggled(self, checked):
        self.le_color.setEnabled(not checked)
        self.btn_color_pick.setEnabled(not checked)
        self.btn_color_preview.setVisible(not checked)

    def _on_default_size_toggled(self, checked):
        self.sp_size.setEnabled(not checked)

    def _pick_color(self):
        current = QColor(self._normalize_hex_color(self.le_color.text()))
        picked = QColorDialog.getColor(current, self, "Chọn màu chữ")
        if not picked.isValid():
            return
        self.le_color.setText(picked.name().upper())
        self._refresh_color_preview()

    def _pick_pinyin_color(self):
        current = QColor(self._normalize_hex_color(self.le_pinyin_color.text()))
        picked = QColorDialog.getColor(current, self, "Chọn màu pinyin")
        if not picked.isValid():
            return
        self.le_pinyin_color.setText(picked.name().upper())

    def _load_settings(self):
        src = settings.get("language_settings.source_lang", "Chinese")
        self.cb_src.setCurrentText(_LANG_DISPLAY.get(src, src))
        tgt = settings.get("language_settings.target_lang", "Vietnamese")
        self.cb_tgt.setCurrentText(_LANG_DISPLAY.get(tgt, tgt))
        om = settings.get("document_settings.output_mode", "overwrite")
        self.cb_mode.setCurrentText(_OUTPUT_REVERSE.get(om, "Ghi đè (Overwrite)"))
        tb = settings.get("processing_options.translate_textboxes", True)
        self.chk_textbox.setChecked(tb)
        ocr = settings.get("ocr_settings.image_text_translation_enabled", False)
        self.chk_ocr.setChecked(ocr)
        auto_open = settings.get("processing_options.auto_open_file", False)
        self.chk_auto_open.setChecked(auto_open)
        self.chk_pinyin.setChecked(settings.get("processing_options.add_chinese_pinyin", False))
        self.sp_pinyin_size.setValue(settings.get("processing_options.pinyin_font_size", 10))
        self.cb_pinyin_font.setCurrentText(settings.get("processing_options.pinyin_font_family", "Arial"))
        self.le_pinyin_color.setText(self._normalize_hex_color(settings.get("processing_options.pinyin_font_color", "#888888")))

        # Style settings
        
        ocr_mode_saved = settings.get("ocr_settings.ocr_display_mode", "overwrite")
        if ocr_mode_saved == "prefix":
            self.cb_ocr_display_mode.setCurrentIndex(1)
        elif ocr_mode_saved == "suffix":
            self.cb_ocr_display_mode.setCurrentIndex(2)
        else:
            self.cb_ocr_display_mode.setCurrentIndex(0)

        ocr_container_saved = settings.get("ocr_settings.ocr_display_container", "textbox")
        self.cb_ocr_display_container.setCurrentIndex(1 if ocr_container_saved == "shape" else 0)

        placement_saved = settings.get(
            "ocr_settings.ocr_textbox_placement_mode",
            settings.get("ocr_settings.ocr_textbox_mode", "whitespace"),
        )
        placement_idx = self.cb_ocr_placement_mode.findData(placement_saved)
        self.cb_ocr_placement_mode.setCurrentIndex(placement_idx if placement_idx >= 0 else 0)

        fmt_mode = settings.get(
            "text_style_settings.translated_text_format_mode",
            "keep_original_format" if settings.get("text_style_settings.keep_format", True) else "custom_format",
        )
        is_keep = fmt_mode == "keep_original_format"
        self.rbtn_keep_format.setChecked(is_keep)
        self.rbtn_custom_format.setChecked(not is_keep)
        self.cb_font.setCurrentText(settings.get("text_style_settings.font_family", "Arial"))
        font_size = settings.get("text_style_settings.font_size", 14)
        if font_size == 0:
            self.chk_default_size.setChecked(True)
            self.sp_size.setValue(14)
        else:
            self.chk_default_size.setChecked(False)
            self.sp_size.setValue(font_size)
        
        font_color = settings.get("text_style_settings.font_color", "#000000")
        if font_color == "default":
            self.chk_default_color.setChecked(True)
            self.le_color.setText("#000000")
        else:
            self.chk_default_color.setChecked(False)
            self.le_color.setText(self._normalize_hex_color(font_color))
        def _set_tri_state(chk, val):
            if val == "default":
                chk.setCheckState(Qt.PartiallyChecked)
            else:
                chk.setCheckState(Qt.Checked if val else Qt.Unchecked)

        _set_tri_state(self.chk_bold, settings.get("text_style_settings.bold", "default"))
        _set_tri_state(self.chk_italic, settings.get("text_style_settings.italic", "default"))
        _set_tri_state(self.chk_underline, settings.get("text_style_settings.underline", "default"))
        use_custom_dir = settings.get("ppt_settings.use_custom_dir", False)
        self.chk_custom_dir.setChecked(use_custom_dir)
        self.le_custom_dir.setText(settings.get("ppt_settings.custom_dir_path", ""))
        self._on_custom_dir_toggled(Qt.Checked if use_custom_dir else Qt.Unchecked)

        self._refresh_color_preview()
        self._toggle_custom_format()
        self._update_ocr_ui_visibility(self.chk_ocr.isChecked())
        self._update_glossary_ui_visibility(self.chk_glossary.isChecked())

    def reload_from_settings(self):
        self._load_settings()

    def _tick(self):
        self.lbl_time.setText(datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))

    # ── File management ──────────────────────────────────────────────

    def _add_files(self):
        fnames, _ = QFileDialog.getOpenFileNames(
            self, "Select PPT files", "", "PowerPoint (*.pptx)"
        )
        for f in fnames:
            if f not in self._files:
                self._files.append(f)
                self.lst.addItem(QListWidgetItem(Path(f).name))
                self._log("INFO", f"[file_loaded] {Path(f).name}")
        self._update_btn()

    def _remove(self):
        for item in self.lst.selectedItems():
            idx = self.lst.row(item)
            self._files.pop(idx)
            self.lst.takeItem(idx)
        self._update_btn()

    def _clear(self):
        self.lst.clear()
        self._files.clear()
        self._update_btn()

    def _update_btn(self):
        self.btn_translate.setEnabled(len(self._files) > 0)

    # ── Pause / Cancel ───────────────────────────────────────────────

    def _on_pause(self):
        self._pause.clear()
        self.btn_pause.setEnabled(False)
        self.btn_resume.setEnabled(True)
        self._log("INFO", "Paused")

    def _on_resume(self):
        self._pause.set()
        self.btn_pause.setEnabled(True)
        self.btn_resume.setEnabled(False)
        self._log("INFO", "Resumed")

    def _on_cancel(self):
        self._cancel.set()
        self.btn_pause.setEnabled(False)
        self.btn_resume.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self._log("WARN", "Cancel requested...")

    # ── Translate ────────────────────────────────────────────────────

    def _translate(self):
        if not self._files:
            return

        # Save settings
        settings.set("language_settings.source_lang", _LANG_REVERSE.get(self.cb_src.currentText(), "Chinese"))
        settings.set("language_settings.target_lang", _LANG_REVERSE.get(self.cb_tgt.currentText(), "Vietnamese"))
        settings.set("document_settings.output_mode", _OUTPUT_MODES.get(self.cb_mode.currentText(), "overwrite"))
        settings.set("processing_options.translate_textboxes", self.chk_textbox.isChecked())
        settings.set("ocr_settings.image_text_translation_enabled", self.chk_ocr.isChecked())
        settings.set("processing_options.auto_open_file", self.chk_auto_open.isChecked())
        settings.set("processing_options.add_chinese_pinyin", self.chk_pinyin.isChecked())
        settings.set("processing_options.pinyin_font_size", self.sp_pinyin_size.value())
        settings.set("processing_options.pinyin_font_family", self.cb_pinyin_font.currentText())
        settings.set("processing_options.pinyin_font_color", self._normalize_hex_color(self.le_pinyin_color.text()))
        settings.set("processing_options.pinyin_format_mode", "keep_original" if self.chk_pinyin_keep_format.isChecked() else "custom")

        idx_mode = self.cb_ocr_display_mode.currentIndex()
        ocr_mode_val = "overwrite"
        if idx_mode == 1: ocr_mode_val = "prefix"
        elif idx_mode == 2: ocr_mode_val = "suffix"
        settings.set("ocr_settings.ocr_display_mode", ocr_mode_val)
        
        idx_container = self.cb_ocr_display_container.currentIndex()
        settings.set("ocr_settings.ocr_display_container", "shape" if idx_container == 1 else "textbox")
        placement_mode = self.cb_ocr_placement_mode.currentData()
        settings.set("ocr_settings.ocr_textbox_placement_mode", placement_mode)
        settings.set("ocr_settings.ocr_textbox_mode", placement_mode)

        fmt_mode = "keep_original_format" if self.rbtn_keep_format.isChecked() else "custom_format"
        settings.set("text_style_settings.translated_text_format_mode", fmt_mode)
        settings.set("text_style_settings.keep_format", self.rbtn_keep_format.isChecked())
        settings.set("text_style_settings.font_family", self.cb_font.currentText())
        settings.set("text_style_settings.font_size", 0 if self.chk_default_size.isChecked() else self.sp_size.value())
        if self.chk_default_color.isChecked():
            settings.set("text_style_settings.font_color", "default")
        else:
            settings.set("text_style_settings.font_color", self._normalize_hex_color(self.le_color.text()))
        def _get_tri_state(chk):
            state = chk.checkState()
            if state == Qt.PartiallyChecked:
                return "default"
            return state == Qt.Checked

        settings.set("text_style_settings.bold", _get_tri_state(self.chk_bold))
        settings.set("text_style_settings.italic", _get_tri_state(self.chk_italic))
        settings.set("text_style_settings.underline", _get_tri_state(self.chk_underline))

        src = settings.get("language_settings.source_lang", "Chinese")
        tgt = settings.get("language_settings.target_lang", "Vietnamese")
        if src == tgt:
            QMessageBox.warning(self, "Error", "Source = Target!")
            return

        if self._thread and self._thread.isRunning():
            QMessageBox.warning(self, "Busy", "Translation already running.")
            return

        self._pause.set()
        self._cancel.clear()
        self.txt_log.clear()
        self.bar.setValue(0)
        self._last_result = {}

        self.btn_translate.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText("Running...")
        self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: #89b4fa; font-weight: bold;")

        self._thread = QThread()
        self._worker = PptWorker(list(self._files), self._pause, self._cancel)
        self._worker.moveToThread(self._thread)

        # Signal connections — proper Qt lifecycle:
        # 1. thread.started → worker.run (start work)
        # 2. worker signals → UI slots (progress, logs)
        # 3. worker.finished → thread.quit (stop event loop)
        # 4. thread.finished → deleteLater (safe cleanup)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._log)
        self._worker.progress.connect(self._on_progress)
        self._worker.file_started.connect(
            lambda name: self.lbl_status.setText(f"Translating: {name}")
        )
        self._worker.result_info.connect(self._store_result)
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(lambda e: self._log("ERROR", f"Error: {e}"))

        # Safe QThread shutdown chain
        self._worker.finished.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_thread)

        self._thread.start()

    def _store_result(self, result: dict):
        self._last_result = result

    def _on_progress(self, current: int, total: int):
        if total > 0:
            self.bar.setMaximum(total)
            self.bar.setValue(current)
            percent = (current / total) * 100
            self.lbl_status.setText(f"Đang dịch... {current}/{total} texts ({percent:.1f}%)")

    def _cleanup_thread(self):
        """Called when QThread.finished fires — safe to deleteLater."""
        if self._worker:
            self._worker.deleteLater()
            self._worker = None
        if self._thread:
            self._thread.deleteLater()
            self._thread = None

    def _on_done(self, out_path: str):
        self.btn_translate.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_resume.setEnabled(False)
        self.btn_cancel.setEnabled(False)

        r = self._last_result
        status = r.get("status", "FAILED") if r else "FAILED"
        t = r.get("translated", 0)
        f = r.get("failed", 0)
        out_paths = r.get("output_paths", [])
        out_str = "\n".join(out_paths) if out_paths else out_path

        # Summary log
        elapsed = r.get("elapsed", 0.0)
        self._log(
            "INFO",
            f"[summary] translated={t} failed={f} elapsed={elapsed:.1f}s",
        )

        if status == "SUCCESS" and out_path:
            self.lbl_status.setText("Done")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: #a6e3a1; font-weight: bold;")
            self._log("INFO", f"Done: {out_path}")
            report = f"Hoàn thành dịch thuật!\n- Tiến độ: 100%\n- Đã dịch: {t} text\n- Lỗi: {f} text\n- Thời gian hoàn thành: {elapsed:.1f} giây\n\nOutput:\n{out_str}"
            QMessageBox.information(self, "Báo cáo hoàn thành", report)
            if settings.get("processing_options.auto_open_file", False) and out_path:
                import os
                try:
                    os.startfile(out_path)
                except Exception as e:
                    self._log("WARN", f"Could not auto-open file: {e}")

        elif status == "PARTIAL":
            self.lbl_status.setText(f"Partial: {t} ok, {f} failed")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: orange; font-weight: bold;")
            QMessageBox.warning(self, "Partial", f"{t} translated, {f} failed.\n\n{out_str}")

        elif status == "CANCELLED":
            self.lbl_status.setText(f"Cancelled ({t} translated)")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: gray; font-weight: bold;")

        else:
            err = r.get("error", "Unknown") if r else "Unknown"
            self.lbl_status.setText(f"Failed: {err[:60]}")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: red; font-weight: bold;")
            QMessageBox.critical(self, "Failed", err)

    # ── Log ──────────────────────────────────────────────────────────

    def _log(self, level, msg):
        colors = {"INFO": "#89b4fa", "WARN": "#fab387", "ERROR": "#f38ba8"}
        color = colors.get(level, "#cdd6f4")
        self.txt_log.append(f'<span style="color:{color}">[{level}] {msg}</span>')
