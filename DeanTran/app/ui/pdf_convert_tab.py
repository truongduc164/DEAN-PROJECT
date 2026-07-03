"""
PDF Converter tab - convert PDF files to Word, Excel, and PowerPoint format.
"""
from __future__ import annotations

import logging
import time
import os
import traceback
import threading
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
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
    QVBoxLayout,
    QWidget,
    QTextEdit,
    QButtonGroup,
)

from app.settings.settings_manager import settings

logger = logging.getLogger("DeanTran.pdf_convert_tab")


class ConvertWorker(QObject):
    log = Signal(str, str)
    progress = Signal(int, int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(
        self,
        file_paths: list[str],
        target_format: str,
        use_custom_dir: bool,
        custom_dir_path: str,
        auto_open: bool,
        pause_event,
        cancel_event,
    ):
        super().__init__()
        self.file_paths = file_paths
        self.target_format = target_format
        self.use_custom_dir = use_custom_dir
        self.custom_dir_path = custom_dir_path
        self.auto_open = auto_open
        self._pause = pause_event
        self._cancel = cancel_event

    def run(self):
        try:
            from app.core.pdf_converter import pdf_to_word, pdf_to_excel, pdf_to_ppt

            total = len(self.file_paths)
            last_out = ""
            for idx, fp in enumerate(self.file_paths, 1):
                if self._cancel.is_set():
                    self.log.emit("WARN", "Tiến trình bị hủy bởi người dùng.")
                    break

                # Handle Pause
                while not self._pause.is_set():
                    if self._cancel.is_set():
                        break
                    time.sleep(0.2)

                if self._cancel.is_set():
                    self.log.emit("WARN", "Tiến trình bị hủy bởi người dùng.")
                    break

                fp_path = Path(fp)
                fname = fp_path.name
                self.log.emit("INFO", f"=== Tệp {idx}/{total}: {fname} ===")

                # Determine output path
                if self.use_custom_dir and self.custom_dir_path:
                    out_dir = Path(self.custom_dir_path)
                    out_dir.mkdir(parents=True, exist_ok=True)
                else:
                    out_dir = fp_path.parent

                out_name = fp_path.stem + f"_converted.{self.target_format}"
                out_path = out_dir / out_name

                # Avoid overwriting
                counter = 1
                while out_path.exists():
                    out_name = fp_path.stem + f"_converted_{counter}.{self.target_format}"
                    out_path = out_dir / out_name
                    counter += 1

                try:
                    if self.target_format == "docx":
                        pdf_to_word(fp_path, out_path, log_fn=lambda lvl, msg: self.log.emit(lvl, msg))
                    elif self.target_format == "xlsx":
                        pdf_to_excel(fp_path, out_path, log_fn=lambda lvl, msg: self.log.emit(lvl, msg))
                    elif self.target_format == "pptx":
                        pdf_to_ppt(fp_path, out_path, log_fn=lambda lvl, msg: self.log.emit(lvl, msg))
                    else:
                        raise ValueError(f"Định dạng không hỗ trợ: {self.target_format}")

                    last_out = str(out_path)

                    if self.auto_open:
                        try:
                            os.startfile(str(out_path))
                        except Exception as e:
                            self.log.emit("WARN", f"Không thể tự động mở file: {e}")

                except Exception as exc:
                    self.log.emit("ERROR", f"❌ Lỗi khi convert {fname}: {exc}")
                    logger.exception(exc)

                self.progress.emit(idx, total)

            self.finished.emit(last_out)
        except Exception as exc:
            self.log.emit("ERROR", f"Lỗi nghiêm trọng: {exc}")
            self.error.emit(str(exc))


class PdfConvertTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._files: list[str] = []
        self._thread: QThread | None = None
        self._worker: ConvertWorker | None = None
        self._pause = threading.Event()
        self._pause.set()
        self._cancel = threading.Event()

        self._build_ui()
        self._load_settings()

    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        # ── Left Column: Config & Control ──────────────────────────────
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(12)

        # Config Panel
        config_group = QGroupBox("1. Cài đặt chuyển đổi (PDF Converter)")
        config_form = QFormLayout(config_group)
        config_form.setContentsMargins(12, 16, 12, 12)
        config_form.setSpacing(8)

        # Radio Group for Format
        self.btn_group = QButtonGroup(self)
        self.rbtn_word = QRadioButton("Tài liệu Word (.docx)  📝")
        self.rbtn_excel = QRadioButton("Bảng tính Excel (.xlsx)  📊")
        self.rbtn_ppt = QRadioButton("Bản trình chiếu PowerPoint (.pptx)  📽️")

        self.btn_group.addButton(self.rbtn_word, 0)
        self.btn_group.addButton(self.rbtn_excel, 1)
        self.btn_group.addButton(self.rbtn_ppt, 2)

        format_layout = QVBoxLayout()
        format_layout.addWidget(self.rbtn_word)
        format_layout.addWidget(self.rbtn_excel)
        format_layout.addWidget(self.rbtn_ppt)
        config_form.addRow("Định dạng đầu ra:", format_layout)

        # Custom output directory
        self.chk_custom_dir = QCheckBox("Lưu vào thư mục khác (Custom Output)")
        self.chk_custom_dir.stateChanged.connect(self._on_custom_dir_toggled)
        config_form.addRow("", self.chk_custom_dir)

        custom_dir_row = QHBoxLayout()
        self.le_custom_dir = QLineEdit()
        self.le_custom_dir.setPlaceholderText("Mặc định: cùng thư mục file gốc")
        self.le_custom_dir.setReadOnly(True)
        self.btn_browse_custom_dir = QPushButton("Browse...")
        self.btn_browse_custom_dir.clicked.connect(self._browse_custom_dir)
        custom_dir_row.addWidget(self.le_custom_dir)
        custom_dir_row.addWidget(self.btn_browse_custom_dir)
        config_form.addRow("Thư mục lưu:", custom_dir_row)

        self.chk_auto_open = QCheckBox("Tự động mở file sau khi xuất")
        self.chk_auto_open.setChecked(settings.get("pdf_converter.auto_open", False))
        self.chk_auto_open.stateChanged.connect(
            lambda s: settings.set("pdf_converter.auto_open", s == Qt.Checked.value)
        )
        config_form.addRow("", self.chk_auto_open)

        left_layout.addWidget(config_group)

        # Controls Panel
        ctrl_group = QGroupBox("2. Điều khiển tiến trình")
        ctrl_layout = QVBoxLayout(ctrl_group)
        ctrl_layout.setContentsMargins(12, 16, 12, 12)
        ctrl_layout.setSpacing(8)

        self.btn_convert = QPushButton("🚀 Bắt đầu chuyển đổi")
        self.btn_convert.setStyleSheet("""
            QPushButton {
                background-color: #2563eb;
                color: white;
                font-size: 13px;
                font-weight: bold;
                padding: 10px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #1d4ed8;
            }
            QPushButton:disabled {
                background-color: #94a3b8;
                color: #cbd5e1;
            }
        """)
        self.btn_convert.clicked.connect(self._start_conversion)
        ctrl_layout.addWidget(self.btn_convert)

        row_btns = QHBoxLayout()
        self.btn_pause = QPushButton("⏸️ Tạm dừng")
        self.btn_pause.setEnabled(False)
        self.btn_pause.clicked.connect(self._on_pause)
        
        self.btn_resume = QPushButton("▶️ Tiếp tục")
        self.btn_resume.setEnabled(False)
        self.btn_resume.clicked.connect(self._on_resume)

        self.btn_cancel = QPushButton("⏹️ Hủy bỏ")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._on_cancel)

        row_btns.addWidget(self.btn_pause)
        row_btns.addWidget(self.btn_resume)
        row_btns.addWidget(self.btn_cancel)
        ctrl_layout.addLayout(row_btns)

        left_layout.addWidget(ctrl_group)
        left_layout.addStretch()

        root.addWidget(left_widget, stretch=1)

        # ── Right Column: File List & Logs ─────────────────────────────
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(12)

        # File List Panel
        lst_group = QGroupBox("3. Danh sách tệp PDF cần chuyển đổi")
        lst_layout = QVBoxLayout(lst_group)
        lst_layout.setContentsMargins(12, 16, 12, 12)
        lst_layout.setSpacing(8)

        self.lst = QListWidget()
        self.lst.setSelectionMode(QAbstractItemView.ExtendedSelection)
        lst_layout.addWidget(self.lst)

        row_file_actions = QHBoxLayout()
        self.btn_add = QPushButton("➕ Thêm file PDF")
        self.btn_add.clicked.connect(self._add_files)
        self.btn_remove = QPushButton("❌ Xóa")
        self.btn_remove.clicked.connect(self._remove)
        self.btn_clear = QPushButton("🗑️ Xóa hết")
        self.btn_clear.clicked.connect(self._clear)

        row_file_actions.addWidget(self.btn_add)
        row_file_actions.addWidget(self.btn_remove)
        row_file_actions.addWidget(self.btn_clear)
        lst_layout.addLayout(row_file_actions)

        right_layout.addWidget(lst_group, stretch=1)

        # Progress and logs
        log_group = QGroupBox("4. Nhật ký xử lý & Tiến độ")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(12, 16, 12, 12)
        log_layout.setSpacing(8)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        log_layout.addWidget(self.progress_bar)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("font-family: 'Consolas', monospace; font-size: 11px; background-color: #fafafa;")
        log_layout.addWidget(self.txt_log)

        # Bottom timestamp row
        self.lbl_time = QLabel()
        self.lbl_time.setStyleSheet("color: #64748b; font-size: 11px;")
        log_layout.addWidget(self.lbl_time)

        right_layout.addWidget(log_group, stretch=1)

        root.addWidget(right_widget, stretch=1)

        # Timer to refresh time label
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(1000)

    def _tick(self):
        self.lbl_time.setText(datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))

    def _load_settings(self):
        target_fmt = settings.get("pdf_converter.target_format", "docx")
        if target_fmt == "xlsx":
            self.rbtn_excel.setChecked(True)
        elif target_fmt == "pptx":
            self.rbtn_ppt.setChecked(True)
        else:
            self.rbtn_word.setChecked(True)

        self.btn_group.buttonClicked.connect(self._on_format_changed)

        use_custom_dir = settings.get("pdf_converter.use_custom_dir", False)
        self.chk_custom_dir.setChecked(use_custom_dir)
        self.le_custom_dir.setText(settings.get("pdf_converter.custom_dir_path", ""))
        self._on_custom_dir_toggled(Qt.Checked if use_custom_dir else Qt.Unchecked)

    def _on_format_changed(self, button):
        fmt_id = self.btn_group.id(button)
        fmt = "docx"
        if fmt_id == 1:
            fmt = "xlsx"
        elif fmt_id == 2:
            fmt = "pptx"
        settings.set("pdf_converter.target_format", fmt)

    def _on_custom_dir_toggled(self, state=None):
        enabled = self.chk_custom_dir.isChecked()
        settings.set("pdf_converter.use_custom_dir", enabled)
        self.le_custom_dir.setEnabled(enabled)
        self.btn_browse_custom_dir.setEnabled(enabled)

    def _browse_custom_dir(self):
        curr_dir = self.le_custom_dir.text().strip() or settings.get("file_picker.last_export_dir", "")
        dir_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu kết quả", curr_dir)
        if dir_path:
            self.le_custom_dir.setText(dir_path)
            settings.set("pdf_converter.custom_dir_path", dir_path)

    def _add_files(self):
        fnames, _ = QFileDialog.getOpenFileNames(self, "Select PDF files", "", "PDF (*.pdf)")
        for f in fnames:
            if f not in self._files:
                self._files.append(f)
                self.lst.addItem(QListWidgetItem(Path(f).name))
                self._log("INFO", f"[file_loaded] {Path(f).name}")
        self.btn_convert.setEnabled(len(self._files) > 0)

    def _remove(self):
        for item in self.lst.selectedItems():
            idx = self.lst.row(item)
            self._files.pop(idx)
            self.lst.takeItem(idx)
        self.btn_convert.setEnabled(len(self._files) > 0)

    def _clear(self):
        self.lst.clear()
        self._files.clear()
        self.btn_convert.setEnabled(False)

    def _log(self, level: str, msg: str):
        t = datetime.now().strftime("%H:%M:%S")
        self.txt_log.append(f"[{t}] [{level}] {msg}")

    # ── Thread control ───────────────────────────────────────────────

    def _start_conversion(self):
        if not self._files:
            QMessageBox.information(self, "Thông báo", "Chưa thêm file nào.")
            return

        fmt = settings.get("pdf_converter.target_format", "docx")
        use_custom = settings.get("pdf_converter.use_custom_dir", False)
        custom_path = settings.get("pdf_converter.custom_dir_path", "")
        auto_open = settings.get("pdf_converter.auto_open", False)

        self.txt_log.clear()
        self.progress_bar.setValue(0)
        self._log("INFO", "Khởi chạy tiến trình chuyển đổi PDF...")

        self._pause.set()
        self._cancel.clear()

        self._thread = QThread()
        self._worker = ConvertWorker(
            file_paths=self._files.copy(),
            target_format=fmt,
            use_custom_dir=use_custom,
            custom_dir_path=custom_path,
            auto_open=auto_open,
            pause_event=self._pause,
            cancel_event=self._cancel,
        )
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._log)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_thread)

        self._thread.start()

        # UI State
        self.btn_convert.setEnabled(False)
        self.btn_add.setEnabled(False)
        self.btn_remove.setEnabled(False)
        self.btn_clear.setEnabled(False)

        self.btn_pause.setEnabled(True)
        self.btn_resume.setEnabled(False)
        self.btn_cancel.setEnabled(True)

    def _on_progress(self, current: int, total: int):
        val = int((current / total) * 100)
        self.progress_bar.setValue(val)

    def _on_pause(self):
        self._pause.clear()
        self.btn_pause.setEnabled(False)
        self.btn_resume.setEnabled(True)
        self._log("WARN", "Tiến trình tạm dừng...")

    def _on_resume(self):
        self._pause.set()
        self.btn_pause.setEnabled(True)
        self.btn_resume.setEnabled(False)
        self._log("INFO", "Tiếp tục tiến trình...")

    def _on_cancel(self):
        self._cancel.set()
        self._pause.set()  # release if paused
        self._log("WARN", "Đang yêu cầu hủy tiến trình...")

    def _on_done(self, last_output: str):
        self._log("INFO", "=== Hoàn thành toàn bộ tệp ===")
        QMessageBox.information(self, "Thành công", "Tiến trình chuyển đổi hoàn tất!")
        self._cleanup_ui()

    def _on_error(self, err_msg: str):
        self._log("ERROR", f"Lỗi luồng xử lý: {err_msg}")
        QMessageBox.critical(self, "Lỗi", f"Tiến trình thất bại:\n{err_msg}")
        self._cleanup_ui()

    def _cleanup_thread(self):
        if self._worker:
            self._worker.deleteLater()
            self._worker = None
        if self._thread:
            self._thread.deleteLater()
            self._thread = None

    def _cleanup_ui(self):
        self.btn_convert.setEnabled(len(self._files) > 0)
        self.btn_add.setEnabled(True)
        self.btn_remove.setEnabled(True)
        self.btn_clear.setEnabled(True)

        self.btn_pause.setEnabled(False)
        self.btn_resume.setEnabled(False)
        self.btn_cancel.setEnabled(False)
