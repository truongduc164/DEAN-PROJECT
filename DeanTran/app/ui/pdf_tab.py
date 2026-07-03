"""
PDF Export tab – import translated Word/Excel/PPT files, configure page setup
(margins, paper size, orientation, centering, fit-to-page for Excel),
then export to PDF or keep original format with applied settings.
"""
from __future__ import annotations

import os
import traceback
import logging
import threading
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
    QTextEdit,
)
from app.settings.settings_manager import settings

logger = logging.getLogger("DeanTran.pdf_tab")

_PAPER_SIZES = ["A4", "A3", "A5", "Letter", "Legal", "B5"]
_ORIENTATIONS = ["Portrait", "Landscape", "Auto"]
_FILE_FILTER = "Documents (*.docx *.xlsx *.pptx);;Word (*.docx);;Excel (*.xlsx);;PowerPoint (*.pptx)"

# Scaling options matching Excel Print dialog
_SCALING_OPTIONS = [
    ("No Scaling", False, 0, 0),                  # no fit
    ("Fit Sheet on One Page", True, 1, 1),         # fit 1x1
    ("Fit All Columns on One Page", True, 1, 0),   # fit width=1, height=auto
    ("Fit All Rows on One Page", True, 0, 1),      # fit width=auto, height=1
    ("Custom Scaling…", True, 1, 1),               # user picks
]


# ---------------------------------------------------------------------------
#  Background worker
# ---------------------------------------------------------------------------
class _ExportWorker(QObject):
    log = Signal(str, str)
    progress = Signal(int, int)
    finished = Signal(str)    # final output path
    all_outputs = Signal(list)  # all output paths for merge
    error = Signal(str)

    def __init__(self, file_paths: list[str], setup_kwargs: dict, output_format: str, overwrite: bool = False, keep_original: bool = False):
        super().__init__()
        self.file_paths = file_paths
        self.setup_kwargs = setup_kwargs
        self.output_format = output_format
        self.overwrite = overwrite
        self.keep_original = keep_original

    def run(self):
        try:
            from app.core.pdf_exporter import PageSetup, export_document

            setup = PageSetup(**self.setup_kwargs)
            total = len(self.file_paths)
            last_output = ""
            output_paths: list[str] = []

            for idx, fp in enumerate(self.file_paths, 1):
                fname = Path(fp).name
                self.log.emit("INFO", f"=== File {idx}/{total}: {fname} ===")
                try:
                    custom_dir = None
                    if settings.get("pdf_settings.use_custom_dir", False):
                        custom_dir_path = settings.get("pdf_settings.custom_dir_path", "")
                        if custom_dir_path:
                            custom_dir = Path(custom_dir_path)
                            custom_dir.mkdir(parents=True, exist_ok=True)

                    out = export_document(
                        input_path=fp,
                        setup=setup,
                        output_format=self.output_format,
                        overwrite=self.overwrite,
                        output_dir=custom_dir,
                        log_fn=lambda lvl, msg: self.log.emit(lvl, msg),
                        keep_original=self.keep_original,
                    )
                    last_output = str(out)
                    output_paths.append(str(out))
                except Exception as exc:
                    tb = traceback.format_exc()
                    logger.error("Export error:\n%s", tb)
                    self.log.emit("ERROR", f"❌ {fname}: {exc}")

                self.progress.emit(idx, total)

            self.all_outputs.emit(output_paths)
            self.finished.emit(last_output)
        except Exception as exc:
            tb = traceback.format_exc()
            logger.error("Export worker exception:\n%s", tb)
            self.log.emit("ERROR", f"Export failed: {exc}")
            self.error.emit(str(exc))
            self.all_outputs.emit([])
            self.finished.emit("")


# ---------------------------------------------------------------------------
#  PDF Tab
# ---------------------------------------------------------------------------
class PdfTab(QWidget):
    def __init__(self):
        super().__init__()
        self._files: list[str] = []
        self._exported_paths: list[str] = []
        self._thread: QThread | None = None
        self._worker: _ExportWorker | None = None

        self._build_ui()
        self._load_settings()

        self._clock = QTimer(self)
        self._clock.timeout.connect(self._tick)
        self._clock.start(1000)
        self._tick()

    # ── UI construction ─────────────────────────────────────────────────
    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(12)

        # ── LEFT panel ──────────────────────────────────────────────────
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)
        left_widget.setMaximumWidth(480)

        # 1. File input
        left_layout.addWidget(QLabel("<b>1. Nhập file (Word / Excel / PPT):</b>"))
        row_btn = QHBoxLayout()
        self.btn_add = QPushButton("+ Add Files")
        self.btn_add.clicked.connect(self._add_files)
        self.btn_rm = QPushButton("Remove")
        self.btn_rm.clicked.connect(self._remove)
        self.btn_clr = QPushButton("Clear")
        self.btn_clr.clicked.connect(self._clear)
        row_btn.addWidget(self.btn_add)
        row_btn.addWidget(self.btn_rm)
        row_btn.addWidget(self.btn_clr)
        row_btn.addStretch()
        left_layout.addLayout(row_btn)

        self.lst = QListWidget()
        self.lst.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.lst.setMaximumHeight(90)
        self.lst.setStyleSheet("font-size: 12px;")
        left_layout.addWidget(self.lst)

        # 2. Page setup group
        page_group = QGroupBox("2. Cài đặt trang (Page Setup)")
        page_form = QFormLayout()
        page_form.setContentsMargins(8, 6, 8, 6)
        page_form.setVerticalSpacing(4)

        # Keep original checkbox
        self.chk_keep_original = QCheckBox("📋 Giữ nguyên định dạng gốc (Keep Original)")
        self.chk_keep_original.setToolTip("Giữ nguyên lề, khổ giấy, hướng trang gốc của file (bỏ qua các cấu hình bên dưới).")
        self.chk_keep_original.toggled.connect(self._on_keep_original_toggled)
        page_form.addRow(self.chk_keep_original)

        # Paper size & Orientation row
        paper_row = QHBoxLayout()
        paper_col = QVBoxLayout()
        paper_col.addWidget(QLabel("Khổ giấy:"))
        self.cb_paper = QComboBox()
        self.cb_paper.addItems(_PAPER_SIZES)
        paper_col.addWidget(self.cb_paper)
        paper_row.addLayout(paper_col)

        orient_col = QVBoxLayout()
        orient_col.addWidget(QLabel("Hướng trang:"))
        self.cb_orient = QComboBox()
        self.cb_orient.addItems(_ORIENTATIONS)
        orient_col.addWidget(self.cb_orient)
        paper_row.addLayout(orient_col)
        page_form.addRow(paper_row)

        # Margins label
        page_form.addRow(QLabel("<i>Lề (mm):</i>"))

        # Margins grid
        margin_grid = QHBoxLayout()
        tb_layout = QVBoxLayout()
        self.sp_margin_top = self._make_margin_spin(25.4)
        self.sp_margin_bottom = self._make_margin_spin(25.4)
        tb_layout.addLayout(self._labeled_spin("Trên:", self.sp_margin_top))
        tb_layout.addLayout(self._labeled_spin("Dưới:", self.sp_margin_bottom))

        lr_layout = QVBoxLayout()
        self.sp_margin_left = self._make_margin_spin(31.8)
        self.sp_margin_right = self._make_margin_spin(31.8)
        lr_layout.addLayout(self._labeled_spin("Trái:", self.sp_margin_left))
        lr_layout.addLayout(self._labeled_spin("Phải:", self.sp_margin_right))

        margin_grid.addLayout(tb_layout)
        margin_grid.addLayout(lr_layout)
        page_form.addRow(margin_grid)

        # Header / Footer distance
        hf_layout = QHBoxLayout()
        self.sp_header = self._make_margin_spin(12.7)
        self.sp_footer = self._make_margin_spin(12.7)
        hf_layout.addLayout(self._labeled_spin("Header:", self.sp_header))
        hf_layout.addLayout(self._labeled_spin("Footer:", self.sp_footer))
        page_form.addRow(hf_layout)

        # Centering
        center_row = QHBoxLayout()
        self.chk_center_h = QCheckBox("Căn giữa ngang")
        self.chk_center_v = QCheckBox("Căn giữa dọc")
        center_row.addWidget(self.chk_center_h)
        center_row.addWidget(self.chk_center_v)
        center_row.addStretch()
        page_form.addRow(center_row)

        # Fit to page options
        page_form.addRow(QLabel("<i>Thu vừa trang (Fit to Page):</i>"))
        self._fit_group = QButtonGroup(self)
        fit_layout = QVBoxLayout()
        fit_layout.setSpacing(2)

        self.rb_no_scaling = QRadioButton("Không thu (No Scaling)")
        self.rb_fit_one_page = QRadioButton("Thu vừa 1 trang (Fit Sheet on One Page)")
        self.rb_fit_columns = QRadioButton("Thu vừa cột trên 1 trang (Fit All Columns on One Page)")
        self.rb_fit_rows = QRadioButton("Thu vừa hàng trên 1 trang (Fit All Rows on One Page)")

        self._fit_group.addButton(self.rb_no_scaling, 0)
        self._fit_group.addButton(self.rb_fit_one_page, 1)
        self._fit_group.addButton(self.rb_fit_columns, 2)
        self._fit_group.addButton(self.rb_fit_rows, 3)
        self.rb_no_scaling.setChecked(True)

        fit_layout.addWidget(self.rb_no_scaling)
        fit_layout.addWidget(self.rb_fit_one_page)
        fit_layout.addWidget(self.rb_fit_columns)
        fit_layout.addWidget(self.rb_fit_rows)
        page_form.addRow(fit_layout)

        page_group.setLayout(page_form)
        left_layout.addWidget(page_group)

        # 2b. Excel-only options (Gridlines & Repeat rows)
        self.grp_excel_opts = QGroupBox("📊 Excel Options (Chỉ Excel)")
        excel_layout = QVBoxLayout()
        excel_layout.setContentsMargins(8, 6, 8, 6)

        excel_opts = QHBoxLayout()
        self.chk_gridlines = QCheckBox("In đường kẻ ô (Gridlines)")
        self.chk_gridlines.setToolTip("In kèm đường kẻ ô lưới khi in ra giấy")
        excel_opts.addWidget(self.chk_gridlines)
        excel_opts.addStretch()
        excel_layout.addLayout(excel_opts)

        repeat_row = QHBoxLayout()
        repeat_row.addWidget(QLabel("Lặp tiêu đề dòng 1 →"))
        self.sp_repeat_rows = QSpinBox()
        self.sp_repeat_rows.setRange(0, 50)
        self.sp_repeat_rows.setValue(0)
        self.sp_repeat_rows.setToolTip("Lặp dòng 1 đến dòng N ở đầu mỗi trang (0 = tắt)")
        self.sp_repeat_rows.setMinimumWidth(55)
        repeat_row.addWidget(self.sp_repeat_rows)
        repeat_row.addWidget(QLabel("(0 = tắt)"))
        repeat_row.addStretch()
        excel_layout.addLayout(repeat_row)

        self.grp_excel_opts.setLayout(excel_layout)
        left_layout.addWidget(self.grp_excel_opts)

        # 3. Output format
        output_group = QGroupBox("3. Định dạng xuất")
        output_layout = QVBoxLayout()
        output_layout.setContentsMargins(8, 6, 8, 6)

        self.rbtn_pdf = QRadioButton("Xuất PDF  📄")
        self.rbtn_pdf.setChecked(True)
        self.rbtn_apply = QRadioButton("Chỉnh file gốc (áp dụng cài đặt trang vào file)  📝")
        self.rbtn_apply.toggled.connect(self._on_apply_toggled)
        output_layout.addWidget(self.rbtn_pdf)
        output_layout.addWidget(self.rbtn_apply)

        # Overwrite option (only visible when "Chỉnh file gốc" is selected)
        self.overwrite_widget = QWidget()
        ow_layout = QVBoxLayout(self.overwrite_widget)
        ow_layout.setContentsMargins(20, 0, 0, 0)
        ow_layout.setSpacing(2)
        self.chk_overwrite = QCheckBox("⚠️ Ghi đè lên file gốc (không tạo bản sao)")
        self.chk_overwrite.setChecked(False)
        self.chk_overwrite.setToolTip("Nếu bỏ chọn: tạo file _formatted bên cạnh file gốc.\n"
                                      "Nếu chọn: ghi đè trực tiếp lên file gốc (KHÔNG THỂ HOÀN TÁC).")
        self.lbl_overwrite_hint = QLabel(
            '<span style="color: #888; font-size: 11px;">'
            'Mặc định: tạo bản sao _formatted (an toàn)</span>'
        )
        ow_layout.addWidget(self.chk_overwrite)
        ow_layout.addWidget(self.lbl_overwrite_hint)
        output_layout.addWidget(self.overwrite_widget)
        self.overwrite_widget.setVisible(False)

        # Custom output directory
        self.chk_custom_dir = QCheckBox("Lưu vào thư mục khác (Custom Output)")
        self.chk_custom_dir.stateChanged.connect(self._on_custom_dir_toggled)
        output_layout.addWidget(self.chk_custom_dir)

        self.custom_dir_widget = QWidget()
        custom_dir_lay = QHBoxLayout(self.custom_dir_widget)
        custom_dir_lay.setContentsMargins(20, 0, 0, 0)
        self.le_custom_dir = QLineEdit()
        self.le_custom_dir.setPlaceholderText("Mặc định: cùng thư mục file gốc")
        self.le_custom_dir.setReadOnly(True)
        self.btn_browse_custom_dir = QPushButton("Browse...")
        self.btn_browse_custom_dir.clicked.connect(self._browse_custom_dir)
        custom_dir_lay.addWidget(self.le_custom_dir)
        custom_dir_lay.addWidget(self.btn_browse_custom_dir)
        output_layout.addWidget(self.custom_dir_widget)

        self.chk_auto_open = QCheckBox("Tự động mở file sau khi xuất")
        output_layout.addWidget(self.chk_auto_open)

        self.chk_merge_pdf_after = QCheckBox("📑 Gộp tất cả PDF sau khi Export")
        self.chk_merge_pdf_after.setToolTip(
            "Sau khi export xong, tự động gộp tất cả file PDF\n"
            "thành 1 file PDF duy nhất."
        )
        output_layout.addWidget(self.chk_merge_pdf_after)

        output_group.setLayout(output_layout)
        left_layout.addWidget(output_group)

        # 4. Merge Tools
        merge_group = QGroupBox("4. Công cụ gộp file (Merge Tools)")
        merge_layout = QVBoxLayout()
        merge_layout.setContentsMargins(8, 6, 8, 6)
        merge_layout.setSpacing(6)

        self.btn_merge_excel = QPushButton("📊  Gộp Excel → Master BOM")
        self.btn_merge_excel.setToolTip(
            "Chọn nhiều file Excel → gộp thành 1 file Master BOM\n"
            "(mỗi file gốc = 1 hoặc nhiều sheet trong file master)"
        )
        self.btn_merge_excel.setStyleSheet(
            "padding: 8px 12px; font-weight: bold; font-size: 12px;"
        )
        self.btn_merge_excel.clicked.connect(self._merge_excel)
        merge_layout.addWidget(self.btn_merge_excel)

        self.btn_merge_pdf = QPushButton("📄  Gộp nhiều PDF → 1 PDF")
        self.btn_merge_pdf.setToolTip(
            "Chọn nhiều file PDF → gộp thành 1 file PDF duy nhất"
        )
        self.btn_merge_pdf.setStyleSheet(
            "padding: 8px 12px; font-weight: bold; font-size: 12px;"
        )
        self.btn_merge_pdf.clicked.connect(self._merge_pdf)
        merge_layout.addWidget(self.btn_merge_pdf)

        self.btn_merge_word = QPushButton("📝  Gộp nhiều Word → 1 Word")
        self.btn_merge_word.setToolTip(
            "Chọn nhiều file Word (.docx) → gộp thành 1 file Word duy nhất\n"
            "Mỗi file được tách bằng ngắt trang (page break)"
        )
        self.btn_merge_word.setStyleSheet(
            "padding: 8px 12px; font-weight: bold; font-size: 12px;"
        )
        self.btn_merge_word.clicked.connect(self._merge_word)
        merge_layout.addWidget(self.btn_merge_word)

        self.btn_merge_ppt = QPushButton("📽️  Gộp nhiều PPT → 1 PPT")
        self.btn_merge_ppt.setToolTip(
            "Chọn nhiều file PowerPoint (.pptx) → gộp thành 1 file.\n"
            "Giữ nguyên theme và slide master của file đầu tiên."
        )
        self.btn_merge_ppt.setStyleSheet(
            "padding: 8px 12px; font-weight: bold; font-size: 12px;"
        )
        self.btn_merge_ppt.clicked.connect(self._merge_ppt)
        merge_layout.addWidget(self.btn_merge_ppt)

        self.chk_custom_merge_path = QCheckBox("Chọn vị trí lưu khác")
        self.chk_custom_merge_path.setToolTip(
            "Mặc định: lưu tại thư mục gốc của file đầu tiên.\n"
            "Bật: cho phép chọn vị trí lưu tùy chỉnh."
        )
        merge_layout.addWidget(self.chk_custom_merge_path)

        merge_group.setLayout(merge_layout)
        left_layout.addWidget(merge_group)

        left_layout.addStretch()
        root.addWidget(left_widget)

        # ── RIGHT panel ─────────────────────────────────────────────────
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        # Time
        row_time = QHBoxLayout()
        self.lbl_time = QLabel("")
        self.lbl_time.setStyleSheet("color: #888; font-size: 12px;")
        row_time.addWidget(self.lbl_time)
        row_time.addStretch()
        right_layout.addLayout(row_time)

        # Progress
        self.bar = QProgressBar()
        self.bar.setFormat("%p% (%v / %m)")
        self.bar.setMaximum(1)
        self.bar.setValue(0)
        right_layout.addWidget(self.bar)

        # Log area
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet(
            "background-color: #1e1e2e; color: #cdd6f4; "
            "font-family: 'Cascadia Code', 'Consolas', monospace; "
            "font-size: 12px; padding: 6px; border-radius: 4px;"
        )
        self.txt_log.setText(
            "[INFO] PDF Export tab ready.\n"
            "[INFO] Supports: Word (.docx), Excel (.xlsx), PowerPoint (.pptx)\n"
            "[INFO] Add files → configure page setup → Export to PDF or original format."
        )
        right_layout.addWidget(self.txt_log, stretch=1)

        # Status label
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("font-size: 13px; padding: 2px 4px;")
        right_layout.addWidget(self.lbl_status)

        # Action buttons
        row_actions = QHBoxLayout()
        self.btn_export = QPushButton("📤  Export")
        self.btn_export.setStyleSheet(
            "padding: 10px 24px; font-weight: bold; font-size: 14px; "
            "background-color: #89b4fa; color: #1e1e2e; border-radius: 6px;"
        )
        self.btn_export.clicked.connect(self._export)
        self.btn_export.setEnabled(False)

        self.btn_reset = QPushButton("🔄  Reset mặc định")
        self.btn_reset.setStyleSheet(
            "padding: 10px 14px; font-weight: bold; font-size: 13px;"
        )
        self.btn_reset.clicked.connect(self._reset_defaults)

        row_actions.addWidget(self.btn_export)
        row_actions.addWidget(self.btn_reset)
        row_actions.addStretch()
        right_layout.addLayout(row_actions)

        root.addWidget(right_widget, stretch=1)

        # Initial visibility
        self._update_excel_opts_visibility()

    # ── Helpers ─────────────────────────────────────────────────────────
    @staticmethod
    def _make_margin_spin(default_mm: float) -> QDoubleSpinBox:
        sp = QDoubleSpinBox()
        sp.setRange(0.0, 200.0)
        sp.setDecimals(1)
        sp.setSingleStep(1.0)
        sp.setSuffix(" mm")
        sp.setValue(default_mm)
        sp.setMinimumWidth(100)
        return sp

    @staticmethod
    def _labeled_spin(label: str, spin) -> QHBoxLayout:
        row = QHBoxLayout()
        lbl = QLabel(label)
        lbl.setMinimumWidth(50)
        row.addWidget(lbl)
        row.addWidget(spin)
        return row

    def _on_apply_toggled(self, checked):
        """Show overwrite option only when 'Chỉnh file gốc' is selected."""
        self.overwrite_widget.setVisible(checked)

    def _on_custom_dir_toggled(self, state=None):
        enabled = self.chk_custom_dir.isChecked()
        settings.set("pdf_settings.use_custom_dir", enabled)
        self.le_custom_dir.setEnabled(enabled)
        self.btn_browse_custom_dir.setEnabled(enabled)

    def _browse_custom_dir(self):
        curr_dir = self.le_custom_dir.text().strip() or settings.get("file_picker.last_export_dir", "")
        dir_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu kết quả", curr_dir)
        if dir_path:
            self.le_custom_dir.setText(dir_path)
            settings.set("pdf_settings.custom_dir_path", dir_path)

    def _on_keep_original_toggled(self, checked):
        """Enable/disable page setup controls based on keep_original setting."""
        self.cb_paper.setEnabled(not checked)
        self.cb_orient.setEnabled(not checked)
        self.sp_margin_top.setEnabled(not checked)
        self.sp_margin_bottom.setEnabled(not checked)
        self.sp_margin_left.setEnabled(not checked)
        self.sp_margin_right.setEnabled(not checked)
        self.sp_header.setEnabled(not checked)
        self.sp_footer.setEnabled(not checked)
        self.chk_center_h.setEnabled(not checked)
        self.chk_center_v.setEnabled(not checked)
        self.rb_no_scaling.setEnabled(not checked)
        self.rb_fit_one_page.setEnabled(not checked)
        self.rb_fit_columns.setEnabled(not checked)
        self.rb_fit_rows.setEnabled(not checked)
        self.chk_gridlines.setEnabled(not checked)
        self.sp_repeat_rows.setEnabled(not checked)

    def _get_scaling_values(self) -> tuple[bool, int, int]:
        """Return (fit_to_page, fit_to_width, fit_to_height) from fit radio buttons."""
        fit_id = self._fit_group.checkedId()
        if fit_id == 0:  # No Scaling
            return False, 0, 0
        elif fit_id == 1:  # Fit Sheet on One Page
            return True, 1, 1
        elif fit_id == 2:  # Fit All Columns on One Page
            return True, 1, 0
        elif fit_id == 3:  # Fit All Rows on One Page
            return True, 0, 1
        return False, 0, 0

    def _has_excel_files(self) -> bool:
        return any(f.lower().endswith(".xlsx") for f in self._files)

    def _update_excel_opts_visibility(self):
        """Show Excel options group only when Excel files are in the list."""
        self.grp_excel_opts.setVisible(self._has_excel_files())

    # ── Settings persistence ────────────────────────────────────────────
    def _load_settings(self):
        self.chk_keep_original.setChecked(settings.get("pdf_export.keep_original", False))
        self.cb_paper.setCurrentText(settings.get("pdf_export.paper_size", "A4"))
        orient = settings.get("pdf_export.orientation", "portrait")
        self.cb_orient.setCurrentText(orient.capitalize())
        self.sp_margin_top.setValue(settings.get("pdf_export.margin_top_cm", 2.54) * 10)
        self.sp_margin_bottom.setValue(settings.get("pdf_export.margin_bottom_cm", 2.54) * 10)
        self.sp_margin_left.setValue(settings.get("pdf_export.margin_left_cm", 3.18) * 10)
        self.sp_margin_right.setValue(settings.get("pdf_export.margin_right_cm", 3.18) * 10)
        self.sp_header.setValue(settings.get("pdf_export.header_distance_cm", 1.27) * 10)
        self.sp_footer.setValue(settings.get("pdf_export.footer_distance_cm", 1.27) * 10)
        self.chk_center_h.setChecked(settings.get("pdf_export.center_horizontally", False))
        self.chk_center_v.setChecked(settings.get("pdf_export.center_vertically", False))
        fit_id = settings.get("pdf_export.fit_option", 0)
        btn = self._fit_group.button(fit_id)
        if btn:
            btn.setChecked(True)
        fmt = settings.get("pdf_export.output_format", "pdf")
        if fmt == "apply":
            self.rbtn_apply.setChecked(True)
        else:
            self.rbtn_pdf.setChecked(True)
        self.chk_auto_open.setChecked(settings.get("pdf_export.auto_open", False))
        self.chk_gridlines.setChecked(settings.get("pdf_export.print_gridlines", False))
        self.sp_repeat_rows.setValue(settings.get("pdf_export.repeat_rows", 0))

        use_custom_dir = settings.get("pdf_settings.use_custom_dir", False)
        self.chk_custom_dir.setChecked(use_custom_dir)
        self.le_custom_dir.setText(settings.get("pdf_settings.custom_dir_path", ""))
        self._on_custom_dir_toggled(Qt.Checked if use_custom_dir else Qt.Unchecked)

        self._update_excel_opts_visibility()
        self._on_keep_original_toggled(self.chk_keep_original.isChecked())

    def _save_settings(self):
        settings.set("pdf_export.keep_original", self.chk_keep_original.isChecked())
        settings.set("pdf_export.paper_size", self.cb_paper.currentText())
        settings.set("pdf_export.orientation", self.cb_orient.currentText().lower())
        settings.set("pdf_export.margin_top_cm", self.sp_margin_top.value() / 10)
        settings.set("pdf_export.margin_bottom_cm", self.sp_margin_bottom.value() / 10)
        settings.set("pdf_export.margin_left_cm", self.sp_margin_left.value() / 10)
        settings.set("pdf_export.margin_right_cm", self.sp_margin_right.value() / 10)
        settings.set("pdf_export.header_distance_cm", self.sp_header.value() / 10)
        settings.set("pdf_export.footer_distance_cm", self.sp_footer.value() / 10)
        settings.set("pdf_export.center_horizontally", self.chk_center_h.isChecked())
        settings.set("pdf_export.center_vertically", self.chk_center_v.isChecked())
        settings.set("pdf_export.fit_option", self._fit_group.checkedId())
        settings.set("pdf_export.output_format", "apply" if self.rbtn_apply.isChecked() else "pdf")
        settings.set("pdf_export.auto_open", self.chk_auto_open.isChecked())
        settings.set("pdf_export.print_gridlines", self.chk_gridlines.isChecked())
        settings.set("pdf_export.repeat_rows", self.sp_repeat_rows.value())
        settings.set("pdf_settings.use_custom_dir", self.chk_custom_dir.isChecked())
        settings.set("pdf_settings.custom_dir_path", self.le_custom_dir.text().strip())

    def _reset_defaults(self):
        self.chk_keep_original.setChecked(False)
        self.cb_paper.setCurrentText("A4")
        self.cb_orient.setCurrentText("Portrait")
        self.sp_margin_top.setValue(25.4)
        self.sp_margin_bottom.setValue(25.4)
        self.sp_margin_left.setValue(31.8)
        self.sp_margin_right.setValue(31.8)
        self.sp_header.setValue(12.7)
        self.sp_footer.setValue(12.7)
        self.chk_center_h.setChecked(False)
        self.chk_center_v.setChecked(False)
        self.rb_no_scaling.setChecked(True)
        self.chk_gridlines.setChecked(False)
        self.sp_repeat_rows.setValue(0)
        self.rbtn_pdf.setChecked(True)
        self._log("INFO", "Page setup reset to defaults (A4 / Portrait / standard margins).")

    # ── Clock ───────────────────────────────────────────────────────────
    def _tick(self):
        self.lbl_time.setText(datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))

    # ── File management ─────────────────────────────────────────────────
    def _add_files(self):
        last_dir = settings.get("file_picker.last_export_dir", "")
        fnames, _ = QFileDialog.getOpenFileNames(
            self, "Select files to export", last_dir, _FILE_FILTER
        )
        if fnames:
            settings.set("file_picker.last_export_dir", str(Path(fnames[0]).parent))
        for f in fnames:
            if f not in self._files:
                ext = Path(f).suffix.lower()
                icon = {"docx": "📝", "xlsx": "📊", "pptx": "📽️"}.get(ext.lstrip("."), "📄")
                self._files.append(f)
                self.lst.addItem(QListWidgetItem(f"{icon} {Path(f).name}"))
                self._log("INFO", f"[loaded] {Path(f).name}")
        self.btn_export.setEnabled(len(self._files) > 0)
        self._update_excel_opts_visibility()

    def _remove(self):
        for item in self.lst.selectedItems():
            idx = self.lst.row(item)
            self._files.pop(idx)
            self.lst.takeItem(idx)
        self.btn_export.setEnabled(len(self._files) > 0)
        self._update_fit_visibility()

    def _clear(self):
        self.lst.clear()
        self._files.clear()
        self.btn_export.setEnabled(False)
        self._update_excel_opts_visibility()

    # ── Export action ───────────────────────────────────────────────────
    def _export(self):
        if not self._files:
            return

        if self._thread and self._thread.isRunning():
            QMessageBox.warning(self, "Busy", "Export is already running.")
            return

        self._save_settings()

        output_format = "apply" if self.rbtn_apply.isChecked() else "pdf"
        overwrite = self.chk_overwrite.isChecked() and output_format == "apply"

        # Confirm overwrite
        if overwrite:
            reply = QMessageBox.warning(
                self, "⚠️ Xác nhận ghi đè",
                "Bạn đã chọn GHI ĐÈ lên file gốc.\n\n"
                "Dữ liệu gốc sẽ bị thay thế và KHÔNG THỂ HOÀN TÁC!\n\n"
                "Bạn có chắc chắn muốn tiếp tục?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        fit_to_page, fit_w, fit_h = self._get_scaling_values()
        setup_kwargs = {
            "paper_size": self.cb_paper.currentText(),
            "orientation": self.cb_orient.currentText().lower(),
            "margin_top_cm": self.sp_margin_top.value() / 10,
            "margin_bottom_cm": self.sp_margin_bottom.value() / 10,
            "margin_left_cm": self.sp_margin_left.value() / 10,
            "margin_right_cm": self.sp_margin_right.value() / 10,
            "header_distance_cm": self.sp_header.value() / 10,
            "footer_distance_cm": self.sp_footer.value() / 10,
            "center_horizontally": self.chk_center_h.isChecked(),
            "center_vertically": self.chk_center_v.isChecked(),
            "fit_to_page": fit_to_page,
            "fit_to_width": fit_w,
            "fit_to_height": fit_h,
            "print_gridlines": self.chk_gridlines.isChecked(),
            "repeat_rows": self.sp_repeat_rows.value(),
        }

        self.txt_log.clear()
        self.bar.setValue(0)
        self.bar.setMaximum(len(self._files))
        action_label = "Applying settings" if output_format == "apply" else "Exporting to PDF"
        self.lbl_status.setText(f"{action_label}…")
        self.lbl_status.setStyleSheet(
            "font-size: 13px; padding: 2px 4px; color: #89b4fa; font-weight: bold;"
        )
        self.btn_export.setEnabled(False)

        file_types = set(Path(f).suffix.lower() for f in self._files)
        types_str = ", ".join(sorted(t.upper().lstrip(".") for t in file_types))
        fmt_label = "PDF" if output_format == "pdf" else "Chỉnh file gốc"
        self._log("INFO", f"Starting: {len(self._files)} file(s) → {fmt_label}")
        self._log("INFO", f"File types: {types_str}")
        if self.chk_keep_original.isChecked():
            self._log("INFO", "Chế độ: Giữ nguyên định dạng gốc của từng file")
        else:
            self._log("INFO", f"Paper: {setup_kwargs['paper_size']} | {setup_kwargs['orientation'].capitalize()}")
            self._log("INFO", f"Margins (cm): T={setup_kwargs['margin_top_cm']} "
                      f"B={setup_kwargs['margin_bottom_cm']} "
                      f"L={setup_kwargs['margin_left_cm']} "
                      f"R={setup_kwargs['margin_right_cm']}")
            if setup_kwargs["center_horizontally"] or setup_kwargs["center_vertically"]:
                self._log("INFO", f"Centering: H={setup_kwargs['center_horizontally']} V={setup_kwargs['center_vertically']}")
            if fit_to_page:
                fit_names = {1: "Fit Sheet on One Page", 2: "Fit All Columns on One Page", 3: "Fit All Rows on One Page"}
                scaling_name = fit_names.get(self._fit_group.checkedId(), "Custom")
                self._log("INFO", f"Scaling: {scaling_name} ({fit_w}×{fit_h})")

        self._exported_paths = []
        self._thread = QThread()
        self._worker = _ExportWorker(
            file_paths=list(self._files),
            setup_kwargs=setup_kwargs,
            output_format=output_format,
            overwrite=overwrite,
            keep_original=self.chk_keep_original.isChecked(),
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._log)
        self._worker.progress.connect(self._on_progress)
        self._worker.all_outputs.connect(self._store_exported_paths)
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(lambda e: self._log("ERROR", e))
        self._worker.finished.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_thread)
        self._thread.start()

    def _on_progress(self, current: int, total: int):
        self.bar.setMaximum(total)
        self.bar.setValue(current)

    def _cleanup_thread(self):
        if self._worker:
            self._worker.deleteLater()
            self._worker = None
        if self._thread:
            self._thread.deleteLater()
            self._thread = None

    def _store_exported_paths(self, paths: list):
        self._exported_paths = list(paths)

    def _on_done(self, last_output: str):
        self.btn_export.setEnabled(len(self._files) > 0)

        if last_output:
            # Auto-merge PDFs if checkbox is checked
            if (self.chk_merge_pdf_after.isChecked()
                    and self.rbtn_pdf.isChecked()
                    and len(self._exported_paths) > 1):
                pdf_paths = [p for p in self._exported_paths if p.lower().endswith(".pdf")]
                if len(pdf_paths) > 1:
                    self._auto_merge_pdfs(pdf_paths)

            self.lbl_status.setText("✅ Export complete!")
            self.lbl_status.setStyleSheet(
                "font-size: 13px; padding: 2px 4px; color: #a6e3a1; font-weight: bold;"
            )
            output_dir = str(Path(last_output).parent)
            QMessageBox.information(
                self, "Export Complete",
                f"Files exported successfully!\n\nOutput folder:\n{output_dir}"
            )
            if self.chk_auto_open.isChecked():
                try:
                    os.startfile(last_output)
                except Exception as exc:
                    self._log("WARN", f"Cannot auto-open: {exc}")
        else:
            self.lbl_status.setText("❌ Export failed")
            self.lbl_status.setStyleSheet(
                "font-size: 13px; padding: 2px 4px; color: #f38ba8; font-weight: bold;"
            )

    def _auto_merge_pdfs(self, pdf_paths: list[str]):
        """Auto-merge exported PDFs into a single file."""
        from app.core.merge_utils import merge_pdfs
        output_dir = Path(pdf_paths[0]).parent
        merged_path = output_dir / "merged_output.pdf"
        # Avoid name collision
        counter = 1
        while merged_path.exists():
            merged_path = output_dir / f"merged_output_{counter}.pdf"
            counter += 1
        try:
            self._log("INFO", "── Auto-merging exported PDFs ──")
            merge_pdfs(pdf_paths, merged_path, log_fn=self._log)
        except Exception as exc:
            self._log("ERROR", f"Auto-merge failed: {exc}")

    # ── Merge Tools ─────────────────────────────────────────────────────
    def _collect_files_from_multiple_dirs(self, title: str, file_filter: str, last_dir_key: str) -> list[str]:
        """
        Helper method to collect files from multiple directories by prompting the user sequentially.
        """
        all_files = []
        current_dir = settings.get(last_dir_key, settings.get("file_picker.last_export_dir", ""))
        
        while True:
            fnames, _ = QFileDialog.getOpenFileNames(
                self, f"{title} (Đang chọn: {len(all_files)} file)", current_dir, file_filter
            )
            if fnames:
                all_files.extend(fnames)
                current_dir = str(Path(fnames[-1]).parent)
                settings.set(last_dir_key, current_dir)
            
            if not all_files:
                break
                
            reply = QMessageBox.question(
                self,
                "Chọn thêm file?",
                f"Đã chọn {len(all_files)} file.\nBạn có muốn chọn thêm file từ thư mục khác không?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                break
                
        return all_files

    def _merge_excel(self):
        """Pick multiple Excel files and merge into a Master BOM."""
        from app.core.merge_utils import merge_excel_to_master, build_merge_filename

        fnames = self._collect_files_from_multiple_dirs(
            "Chọn các file Excel để gộp", "Excel Files (*.xlsx *.xls)", "merge.last_dir_excel"
        )
        if not fnames or len(fnames) < 2:
            if fnames and len(fnames) == 1:
                QMessageBox.information(self, "Info", "Cần chọn ít nhất 2 file để gộp.")
            return

        output_dir = Path(fnames[0]).parent
        merge_name = build_merge_filename(fnames, ".xlsx")
        default_name = output_dir / merge_name

        # Always require choosing where to save the merged output
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file Master BOM kết quả tại", str(default_name), "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        self._log("INFO", "── Merge Excel → Master BOM ──")
        self._log("INFO", f"Input: {len(fnames)} files")
        for idx, f in enumerate(fnames, 1):
            self._log("INFO", f"  [{idx}] {f}")
            
        try:
            result = merge_excel_to_master(fnames, save_path, log_fn=self._log)
            QMessageBox.information(
                self, "Merge Complete",
                f"Master BOM saved successfully!\n\n{result}"
            )
        except Exception as exc:
            self._log("ERROR", f"Merge Excel failed: {exc}")
            QMessageBox.critical(self, "Error", f"Merge failed:\n{exc}")

    def _merge_pdf(self):
        """Pick multiple PDF files and merge into one."""
        from app.core.merge_utils import merge_pdfs, build_merge_filename

        fnames = self._collect_files_from_multiple_dirs(
            "Chọn các file PDF để gộp", "PDF Files (*.pdf)", "merge.last_dir_pdf"
        )
        if not fnames or len(fnames) < 2:
            if fnames and len(fnames) == 1:
                QMessageBox.information(self, "Info", "Cần chọn ít nhất 2 file để gộp.")
            return

        output_dir = Path(fnames[0]).parent
        merge_name = build_merge_filename(fnames, ".pdf")
        default_name = output_dir / merge_name

        # Always require choosing where to save the merged output
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file PDF kết quả tại", str(default_name), "PDF Files (*.pdf)"
        )
        if not save_path:
            return

        self._log("INFO", "── Merge PDFs ──")
        self._log("INFO", f"Input: {len(fnames)} files")
        for idx, f in enumerate(fnames, 1):
            self._log("INFO", f"  [{idx}] {f}")
            
        try:
            result = merge_pdfs(fnames, save_path, log_fn=self._log)
            QMessageBox.information(
                self, "Merge Complete",
                f"Merged PDF saved successfully!\n\n{result}"
            )
        except Exception as exc:
            self._log("ERROR", f"Merge PDF failed: {exc}")
            QMessageBox.critical(self, "Error", f"Merge failed:\n{exc}")

    def _merge_word(self):
        """Pick multiple Word files and merge into one."""
        from app.core.merge_utils import merge_word_documents, build_merge_filename

        fnames = self._collect_files_from_multiple_dirs(
            "Chọn các file Word để gộp", "Word Files (*.docx)", "merge.last_dir_word"
        )
        if not fnames or len(fnames) < 2:
            if fnames and len(fnames) == 1:
                QMessageBox.information(self, "Info", "Cần chọn ít nhất 2 file để gộp.")
            return

        output_dir = Path(fnames[0]).parent
        merge_name = build_merge_filename(fnames, ".docx")
        default_name = output_dir / merge_name

        # Always require choosing where to save the merged output
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file Word kết quả tại", str(default_name), "Word Files (*.docx)"
        )
        if not save_path:
            return

        self._log("INFO", "── Merge Word Documents ──")
        self._log("INFO", f"Input: {len(fnames)} files")
        for idx, f in enumerate(fnames, 1):
            self._log("INFO", f"  [{idx}] {f}")
            
        try:
            result = merge_word_documents(fnames, save_path, log_fn=self._log)
            QMessageBox.information(
                self, "Merge Complete",
                f"Merged Word saved successfully!\n\n{result}"
            )
        except Exception as exc:
            self._log("ERROR", f"Merge Word failed: {exc}")
            QMessageBox.critical(self, "Error", f"Merge failed:\n{exc}")

    def _merge_ppt(self):
        """Pick multiple PowerPoint files and merge into one."""
        from app.core.merge_utils import merge_pptx, build_merge_filename

        fnames = self._collect_files_from_multiple_dirs(
            "Chọn các file PowerPoint để gộp", "PowerPoint Files (*.pptx)", "merge.last_dir_ppt"
        )
        if not fnames or len(fnames) < 2:
            if fnames and len(fnames) == 1:
                QMessageBox.information(self, "Info", "Cần chọn ít nhất 2 file để gộp.")
            return

        output_dir = Path(fnames[0]).parent
        merge_name = build_merge_filename(fnames, ".pptx")
        default_name = output_dir / merge_name

        # Always require choosing where to save the merged output
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file PPT kết quả tại", str(default_name), "PowerPoint Files (*.pptx)"
        )
        if not save_path:
            return

        self._log("INFO", "── Merge PowerPoint ──")
        self._log("INFO", f"Input: {len(fnames)} files")
        for idx, f in enumerate(fnames, 1):
            self._log("INFO", f"  [{idx}] {f}")
            
        try:
            result = merge_pptx(fnames, save_path, log_fn=self._log)
            QMessageBox.information(
                self, "Merge Complete",
                f"Merged PPT saved successfully!\n\n{result}"
            )
        except Exception as exc:
            self._log("ERROR", f"Merge PPT failed: {exc}")
            QMessageBox.critical(self, "Error", f"Merge failed:\n{exc}")

    # ── Logging ─────────────────────────────────────────────────────────
    def _log(self, level: str, msg: str):
        colors = {"INFO": "#89b4fa", "WARN": "#fab387", "ERROR": "#f38ba8"}
        color = colors.get(level, "#cdd6f4")
        self.txt_log.append(f'<span style="color:{color}">[{level}] {msg}</span>')
