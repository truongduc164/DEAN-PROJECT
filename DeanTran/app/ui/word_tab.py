"""
Word tab with PPT-style layout.
"""
from __future__ import annotations

import logging
import re
import threading
import traceback
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QColorDialog,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLineEdit,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QComboBox,
    QSpinBox,
    QVBoxLayout,
    QWidget,
    QTextEdit,
)

from app.settings.settings_manager import settings

logger = logging.getLogger("DeanTran.word_tab")

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


class WordWorker(QObject):
    log = Signal(str, str)
    progress = Signal(int, int)
    finished = Signal(str)
    error = Signal(str)
    result_info = Signal(dict)
    file_started = Signal(str)

    def __init__(self, file_paths: list[str], pause_event, cancel_event):
        super().__init__()
        self.file_paths = file_paths
        self._pause = pause_event
        self._cancel = cancel_event

    def run(self):
        import time
        t_worker_start = time.time()
        try:
            from app.core.event_manager import EventManager
            from app.core.key_provider import KeyProvider
            from app.core.translators.translator_service import create_translator
            from app.core.word_processor import WordProcessor

            model = settings.get("selected_models.gemini", "gemini-3.1-flash-lite")
            source = settings.get("language_settings.source_lang", "Chinese")
            target = settings.get("language_settings.target_lang", "Vietnamese")
            output_mode = settings.get("document_settings.output_mode", "overwrite")
            check_glossary = settings.get("processing_options.check_glossary_before_ai", True)
            doc_type = settings.get("document_settings.document_type", "General")

            kp = KeyProvider()
            key = kp.get_key()
            key_prefix = key[:6] + "****" if key and len(key) >= 6 else "(none)"
            self.log.emit("INFO", f"[config] model={model} {source}->{target} mode={output_mode}")
            self.log.emit("INFO", f"[key] prefix={key_prefix}")

            # Prompt
            from app.storage.prompt_store import PromptStore
            prompt = PromptStore().get(doc_type)
            if prompt:
                self.log.emit("INFO", f"[prompt] ({doc_type}): {prompt[:80]}…")

            translator = create_translator(model_name=model)
            aggregate = {
                "status": "SUCCESS",
                "translated": 0,
                "failed": 0,
                "elapsed": 0.0,
                "output_paths": [],
            }
            total_files = len(self.file_paths)

            for idx, fp in enumerate(self.file_paths, 1):
                if self._cancel.is_set():
                    self.log.emit("WARN", "Cancel requested, skipping remaining files.")
                    break

                fname = Path(fp).name
                self.file_started.emit(fname)
                self.log.emit("INFO", f"=== File {idx}/{total_files}: {fname} ===")

                em = EventManager()
                em.subscribe("log", lambda lvl, msg: self.log.emit(lvl, msg))
                em.subscribe("progress", lambda cur, tot: self.progress.emit(cur, tot))

                from app.term_engine.glossary_loader import GlossaryLoader
                glossary = None
                if check_glossary:
                    glossary_path = settings.get("processing_options.glossary_file_path", "")
                    if glossary_path and Path(glossary_path).exists():
                        glossary = GlossaryLoader(glossary_path, source_lang=source, target_lang=target)

                proc = WordProcessor(
                    translator=translator,
                    event_manager=em,
                    source_lang=source,
                    target_lang=target,
                    output_mode=output_mode,
                    glossary=glossary,
                    use_glossary=check_glossary,
                    pause_event=self._pause,
                    cancel_event=self._cancel,
                )

                custom_dir = None
                if settings.get("word_settings.use_custom_dir", False):
                    custom_dir_path = settings.get("word_settings.custom_dir_path", "")
                    if custom_dir_path:
                        custom_dir = Path(custom_dir_path)
                        custom_dir.mkdir(parents=True, exist_ok=True)

                out = proc.process(fp, output_dir=custom_dir)
                r = getattr(proc, "last_result", {})
                aggregate["translated"] += r.get("translated", 0)
                aggregate["failed"] += r.get("failed", 0)
                aggregate["elapsed"] += r.get("elapsed", 0.0)
                aggregate["output_paths"].append(str(out))

                s = r.get("status", "SUCCESS")
                if s == "FAILED":
                    aggregate["status"] = "FAILED"
                elif s in ("PARTIAL", "CANCELLED") and aggregate["status"] == "SUCCESS":
                    aggregate["status"] = s

            if self._cancel.is_set() and aggregate["status"] == "SUCCESS":
                aggregate["status"] = "CANCELLED"

            aggregate["elapsed"] = time.time() - t_worker_start
            self.result_info.emit(aggregate)
            last = aggregate["output_paths"][-1] if aggregate["output_paths"] else ""
            self.finished.emit(last)
        except Exception as exc:
            tb = traceback.format_exc()
            logger.error("Word worker exception:\n%s", tb)
            self.log.emit("ERROR", f"Word translation failed: {exc}")
            self.error.emit(str(exc))
            self.result_info.emit({"status": "FAILED", "error": str(exc)})
            self.finished.emit("")


class WordTab(QWidget):
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

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)
        left_widget.setMaximumWidth(450)

        left_layout.addWidget(QLabel("<b>1. Nhập File Word:</b>"))
        row_files_btn = QHBoxLayout()
        self.btn_add = QPushButton("+ Add DOCX")
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
        self.lst.setMaximumHeight(110)
        self.lst.setStyleSheet("font-size: 12px;")
        left_layout.addWidget(self.lst)

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
        self.chk_textbox.setChecked(settings.get("processing_options.translate_textboxes", True))
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

        self.chk_auto_open = QCheckBox("Auto open output file")
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

        self.chk_header_footer = QCheckBox("Translate headers / footers")
        self.chk_header_footer.setChecked(settings.get("processing_options.translate_headers_footers", True))
        options_form.addRow("", self.chk_header_footer)

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

        self.grp_style = QGroupBox("3. Translated Text Format")
        style_layout = QVBoxLayout()
        style_layout.setContentsMargins(8, 4, 8, 4)

        self.rbtn_keep_format = QRadioButton("Keep original format")
        self.rbtn_keep_format.toggled.connect(self._toggle_custom_format)
        style_layout.addWidget(self.rbtn_keep_format)

        self.rbtn_custom_format = QRadioButton("Custom translated format")
        style_layout.addWidget(self.rbtn_custom_format)

        self.custom_format_widget = QWidget()
        custom_layout = QFormLayout(self.custom_format_widget)
        custom_layout.setContentsMargins(16, 0, 0, 0)

        self.cb_font = QComboBox()
        self.cb_font.setEditable(True)
        self.cb_font.addItems(["Mặc định", "Arial", "Times New Roman", "Calibri", "Tahoma"])
        custom_layout.addRow("Font:", self.cb_font)

        size_row = QHBoxLayout()
        self.chk_default_size = QCheckBox("Mặc định")
        self.sp_size = QSpinBox()
        self.sp_size.setRange(8, 72)
        self.sp_size.setValue(14)
        self.chk_default_size.toggled.connect(self._on_default_size_toggled)
        size_row.addWidget(self.chk_default_size)
        size_row.addWidget(self.sp_size)
        size_row.addStretch()
        custom_layout.addRow("Size:", size_row)

        color_row = QHBoxLayout()
        self.chk_default_color = QCheckBox("Mặc định")
        self.le_color = QLineEdit("#000000")
        self.le_color.setMaximumWidth(90)
        self.le_color.editingFinished.connect(self._on_color_text_changed)
        self.btn_color_pick = QPushButton("🎨")
        self.btn_color_pick.setMaximumWidth(32)
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

        self.chk_bold = QCheckBox("Bold")
        self.chk_bold.setTristate(True)
        custom_layout.addRow("", self.chk_bold)
        self.chk_italic = QCheckBox("Italic")
        self.chk_italic.setTristate(True)
        custom_layout.addRow("", self.chk_italic)
        self.chk_underline = QCheckBox("Underline")
        self.chk_underline.setTristate(True)
        custom_layout.addRow("", self.chk_underline)

        style_layout.addWidget(self.custom_format_widget)
        self.grp_style.setLayout(style_layout)
        left_layout.addWidget(self.grp_style)
        left_layout.addStretch()
        root.addWidget(left_widget)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        row_time = QHBoxLayout()
        self.lbl_time = QLabel("")
        self.lbl_time.setStyleSheet("color: #888; font-size: 12px;")
        self.lbl_warn = QLabel("")
        self.lbl_warn.setStyleSheet("color: red; font-weight: bold;")
        row_time.addWidget(self.lbl_time)
        row_time.addStretch()
        row_time.addWidget(self.lbl_warn)
        right_layout.addLayout(row_time)

        self.lbl_current = QLabel("")
        self.lbl_current.setStyleSheet("font-size: 12px; color: #555;")
        right_layout.addWidget(self.lbl_current)

        self.bar = QProgressBar()
        self.bar.setFormat("%p% (%v / %m)")
        right_layout.addWidget(self.bar)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet(
            "background-color: #1e1e2e; color: #cdd6f4; "
            "font-family: 'Cascadia Code', 'Consolas', monospace; "
            "font-size: 12px; padding: 6px; border-radius: 4px;"
        )
        right_layout.addWidget(self.txt_log, stretch=7)

        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px;")
        right_layout.addWidget(self.lbl_status)

        row_btns = QHBoxLayout()
        self.btn_translate = QPushButton("Translate Word")
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

    def _load_settings(self):
        src = settings.get("language_settings.source_lang", "Chinese")
        self.cb_src.setCurrentText(_LANG_DISPLAY.get(src, src))
        tgt = settings.get("language_settings.target_lang", "Vietnamese")
        self.cb_tgt.setCurrentText(_LANG_DISPLAY.get(tgt, tgt))
        om = settings.get("document_settings.output_mode", "overwrite")
        self.cb_mode.setCurrentText(_OUTPUT_REVERSE.get(om, "Ghi đè (Overwrite)"))
        self.chk_textbox.setChecked(settings.get("processing_options.translate_textboxes", True))
        ocr = settings.get("ocr_settings.image_text_translation_enabled", False)
        self.chk_ocr.setChecked(ocr)
        self.cb_ocr_engine.setCurrentText(settings.get("ocr_settings.engine", "paddle"))
        glossary = settings.get("processing_options.check_glossary_before_ai", True)
        self.chk_glossary.setChecked(glossary)
        self.le_glossary_path.setText(settings.get("processing_options.glossary_file_path", ""))
        auto_open = settings.get("processing_options.auto_open_file", False)
        self.chk_auto_open.setChecked(auto_open)
        self.chk_pinyin.setChecked(settings.get("processing_options.add_chinese_pinyin", False))
        self.sp_pinyin_size.setValue(settings.get("processing_options.pinyin_font_size", 10))
        self.cb_pinyin_font.setCurrentText(settings.get("processing_options.pinyin_font_family", "Arial"))
        self.le_pinyin_color.setText(self._normalize_hex_color(settings.get("processing_options.pinyin_font_color", "#888888")))

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
        use_custom_dir = settings.get("word_settings.use_custom_dir", False)
        self.chk_custom_dir.setChecked(use_custom_dir)
        self.le_custom_dir.setText(settings.get("word_settings.custom_dir_path", ""))
        self._on_custom_dir_toggled(Qt.Checked if use_custom_dir else Qt.Unchecked)

        self._refresh_color_preview()
        self._toggle_custom_format()
        self._update_ocr_ui_visibility(ocr)
        self._update_glossary_ui_visibility(glossary)

    def reload_from_settings(self):
        self._load_settings()

    def _tick(self):
        self.lbl_time.setText(datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))

    def _toggle_custom_format(self):
        self.custom_format_widget.setEnabled(self.rbtn_custom_format.isChecked())

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
        picked = QColorDialog.getColor(current, self, "Chon mau chu")
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
        settings.set("word_settings.use_custom_dir", enabled)
        self.le_custom_dir.setEnabled(enabled)
        self.btn_browse_custom_dir.setEnabled(enabled)

    def _browse_custom_dir(self):
        curr_dir = self.le_custom_dir.text().strip() or settings.get("file_picker.last_export_dir", "")
        dir_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu kết quả", curr_dir)
        if dir_path:
            self.le_custom_dir.setText(dir_path)
            settings.set("word_settings.custom_dir_path", dir_path)

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

    def _add_files(self):
        fnames, _ = QFileDialog.getOpenFileNames(self, "Select Word files", "", "Word (*.docx)")
        for f in fnames:
            if f not in self._files:
                self._files.append(f)
                self.lst.addItem(QListWidgetItem(Path(f).name))
                self._log("INFO", f"[file_loaded] {Path(f).name}")
        self.btn_translate.setEnabled(len(self._files) > 0)

    def _remove(self):
        for item in self.lst.selectedItems():
            idx = self.lst.row(item)
            self._files.pop(idx)
            self.lst.takeItem(idx)
        self.btn_translate.setEnabled(len(self._files) > 0)

    def _clear(self):
        self.lst.clear()
        self._files.clear()
        self.btn_translate.setEnabled(False)

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
        self._log("WARN", "Cancel requested")

    def _translate(self):
        if not self._files:
            return

        settings.set("language_settings.source_lang", _LANG_REVERSE.get(self.cb_src.currentText(), "Chinese"))
        settings.set("language_settings.target_lang", _LANG_REVERSE.get(self.cb_tgt.currentText(), "Vietnamese"))
        settings.set("document_settings.output_mode", _OUTPUT_MODES.get(self.cb_mode.currentText(), "overwrite"))
        settings.set("processing_options.translate_textboxes", self.chk_textbox.isChecked())
        settings.set("ocr_settings.image_text_translation_enabled", self.chk_ocr.isChecked())
        settings.set("ocr_settings.engine", self.cb_ocr_engine.currentText())
        settings.set("processing_options.auto_open_file", self.chk_auto_open.isChecked())
        settings.set("processing_options.translate_headers_footers", self.chk_header_footer.isChecked())
        settings.set("processing_options.add_chinese_pinyin", self.chk_pinyin.isChecked())
        settings.set("processing_options.pinyin_font_size", self.sp_pinyin_size.value())
        settings.set("processing_options.pinyin_font_family", self.cb_pinyin_font.currentText())
        settings.set("processing_options.pinyin_font_color", self._normalize_hex_color(self.le_pinyin_color.text()))
        settings.set("processing_options.pinyin_format_mode", "keep_original" if self.chk_pinyin_keep_format.isChecked() else "custom")
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
        if self.chk_ocr.isChecked():
            self._log("INFO", "Word OCR image translation enabled")

        self.btn_translate.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText("Running...")
        self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: #89b4fa; font-weight: bold;")

        self._thread = QThread()
        self._worker = WordWorker(list(self._files), self._pause, self._cancel)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._log)
        self._worker.progress.connect(self._on_progress)
        self._worker.file_started.connect(lambda name: self.lbl_current.setText(name))
        self._worker.result_info.connect(self._store_result)
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(lambda e: self._log("ERROR", e))
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
            self.lbl_status.setText(f"Đang dịch... {current}/{total} đoạn ({percent:.1f}%)")

    def _cleanup_thread(self):
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
        translated = r.get("translated", 0)
        failed = r.get("failed", 0)
        elapsed = r.get("elapsed", 0.0)
        out_paths = r.get("output_paths", [])
        out_str = "\n".join(out_paths) if out_paths else out_path

        if status == "SUCCESS" and out_path:
            self.lbl_status.setText("Done")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: #a6e3a1; font-weight: bold;")
            report = f"Hoàn thành dịch thuật!\n- Tiến độ: 100%\n- Đã dịch: {translated} đoạn\n- Lỗi: {failed} đoạn\n- Thời gian hoàn thành: {elapsed:.1f} giây\n\nOutput:\n{out_str}"
            QMessageBox.information(self, "Báo cáo hoàn thành", report)
            if settings.get("processing_options.auto_open_file", False):
                import os
                try:
                    os.startfile(out_path)
                except Exception as exc:
                    self._log("WARN", f"Cannot auto-open output: {exc}")
        elif status == "PARTIAL":
            self.lbl_status.setText(f"Partial: {translated} ok, {failed} failed")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: orange; font-weight: bold;")
            QMessageBox.warning(self, "Partial", f"{translated} translated, {failed} failed.\n\n{out_str}")
        elif status == "CANCELLED":
            self.lbl_status.setText(f"Cancelled ({translated} translated)")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: gray; font-weight: bold;")
        else:
            err = r.get("error", "Unknown") if r else "Unknown"
            self.lbl_status.setText(f"Failed: {err[:60]}")
            self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px; color: red; font-weight: bold;")
            QMessageBox.critical(self, "Failed", err)

    def _log(self, level, msg):
        colors = {"INFO": "#89b4fa", "WARN": "#fab387", "ERROR": "#f38ba8"}
        color = colors.get(level, "#cdd6f4")
        self.txt_log.append(f'<span style="color:{color}">[{level}] {msg}</span>')
