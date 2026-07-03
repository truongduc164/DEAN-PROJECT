"""
SelectSheetsDialog – Modern modal dialog to choose which sheets to translate.

Features:
  - Tree view with tri-state checkboxes: File → Sheets
  - Select All / Clear All buttons
  - Live summary "Đã chọn: X sheet"
  - 0-sheet warning
  - Keyboard: Enter=OK, Esc=Cancel
  - Persistent selection per file path
  - Modern dark-themed UI (Catppuccin Mocha palette)
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import openpyxl
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTreeWidget, QTreeWidgetItem, QMessageBox,
    QHeaderView, QFrame, QWidget,
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QIcon, QColor

from app.settings.settings_manager import settings


# ══════════════════════════════════════════════════════════════════════
# Helper: collect sheets from files
# ══════════════════════════════════════════════════════════════════════

def collect_all_sheets(file_paths: list[str]) -> dict[str, list[str]]:
    """Read sheet names from each Excel file.

    Returns dict: file_path → [sheet_name, ...]
    """
    result: dict[str, list[str]] = {}
    for fp in file_paths:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            result[fp] = list(wb.sheetnames)
            wb.close()
        except Exception:
            result[fp] = []
    return result


# ══════════════════════════════════════════════════════════════════════
# Persistence helpers
# ══════════════════════════════════════════════════════════════════════

_SETTINGS_KEY = "sheet_selection_cache"


def _load_saved_selections() -> dict[str, list[str]]:
    """Load saved sheet selections from settings."""
    return settings.get(_SETTINGS_KEY, {}) or {}


def _save_selections(selections: dict[str, list[str]]) -> None:
    """Persist sheet selections to settings."""
    settings.set(_SETTINGS_KEY, selections)


# ══════════════════════════════════════════════════════════════════════
# Catppuccin Mocha color palette
# ══════════════════════════════════════════════════════════════════════

_COLORS = {
    "base": "#1e1e2e",
    "mantle": "#181825",
    "crust": "#11111b",
    "surface0": "#313244",
    "surface1": "#45475a",
    "surface2": "#585b70",
    "overlay0": "#6c7086",
    "text": "#cdd6f4",
    "subtext0": "#a6adc8",
    "subtext1": "#bac2de",
    "blue": "#89b4fa",
    "green": "#a6e3a1",
    "red": "#f38ba8",
    "peach": "#fab387",
    "mauve": "#cba6f7",
    "teal": "#94e2d5",
    "yellow": "#f9e2af",
    "lavender": "#b4befe",
}

# ══════════════════════════════════════════════════════════════════════
# Master stylesheet
# ══════════════════════════════════════════════════════════════════════

_DIALOG_STYLESHEET = f"""
    QDialog {{
        background-color: {_COLORS['base']};
        color: {_COLORS['text']};
        font-family: 'Segoe UI', 'Arial', sans-serif;
        font-size: 13px;
    }}

    /* ── Header label ──────────────────────────────────────── */
    QLabel#dialogTitle {{
        color: {_COLORS['text']};
        font-size: 17px;
        font-weight: bold;
        padding: 2px 0;
    }}
    QLabel#dialogSubtitle {{
        color: {_COLORS['subtext0']};
        font-size: 12px;
        padding: 0 0 4px 0;
    }}
    QLabel#summaryLabel {{
        color: {_COLORS['green']};
        font-weight: bold;
        font-size: 13px;
        padding: 4px 10px;
        background-color: {_COLORS['surface0']};
        border-radius: 6px;
    }}

    /* ── Buttons ───────────────────────────────────────────── */
    QPushButton {{
        background-color: {_COLORS['surface0']};
        color: {_COLORS['text']};
        border: 1px solid {_COLORS['surface1']};
        border-radius: 6px;
        padding: 6px 14px;
        font-size: 12px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {_COLORS['surface1']};
        border-color: {_COLORS['overlay0']};
    }}
    QPushButton:pressed {{
        background-color: {_COLORS['surface2']};
    }}
    QPushButton#btnOk {{
        background-color: {_COLORS['blue']};
        color: {_COLORS['crust']};
        font-weight: bold;
        font-size: 13px;
        padding: 8px 28px;
        border: none;
        border-radius: 8px;
    }}
    QPushButton#btnOk:hover {{
        background-color: {_COLORS['lavender']};
    }}
    QPushButton#btnOk:pressed {{
        background-color: {_COLORS['mauve']};
    }}
    QPushButton#btnCancel {{
        background-color: transparent;
        color: {_COLORS['subtext0']};
        border: 1px solid {_COLORS['surface1']};
        padding: 8px 22px;
        font-size: 13px;
        border-radius: 8px;
    }}
    QPushButton#btnCancel:hover {{
        background-color: {_COLORS['surface0']};
        color: {_COLORS['text']};
    }}
    QPushButton#btnSelectAll {{
        color: {_COLORS['green']};
        border-color: {_COLORS['green']}40;
    }}
    QPushButton#btnSelectAll:hover {{
        background-color: {_COLORS['green']}18;
        border-color: {_COLORS['green']}80;
    }}
    QPushButton#btnClearAll {{
        color: {_COLORS['red']};
        border-color: {_COLORS['red']}40;
    }}
    QPushButton#btnClearAll:hover {{
        background-color: {_COLORS['red']}18;
        border-color: {_COLORS['red']}80;
    }}

    /* ── Tree widget ───────────────────────────────────────── */
    QTreeWidget {{
        background-color: {_COLORS['mantle']};
        color: {_COLORS['text']};
        border: 1px solid {_COLORS['surface0']};
        border-radius: 8px;
        padding: 4px;
        outline: none;
        font-size: 13px;
    }}
    QTreeWidget::item {{
        padding: 4px 2px;
        border-radius: 4px;
    }}
    QTreeWidget::item:hover {{
        background-color: {_COLORS['surface0']}80;
    }}
    QTreeWidget::item:selected {{
        background-color: {_COLORS['blue']}25;
        color: {_COLORS['text']};
    }}
    QTreeWidget::branch {{
        background-color: transparent;
    }}
    QTreeWidget QHeaderView::section {{
        background-color: {_COLORS['surface0']};
        color: {_COLORS['subtext1']};
        border: none;
        border-bottom: 2px solid {_COLORS['blue']}60;
        padding: 6px 8px;
        font-weight: bold;
        font-size: 12px;
    }}

    /* ── Separator line ────────────────────────────────────── */
    QFrame#separator {{
        background-color: {_COLORS['surface0']};
        max-height: 1px;
    }}
"""


# ══════════════════════════════════════════════════════════════════════
# SelectSheetsDialog
# ══════════════════════════════════════════════════════════════════════

class SelectSheetsDialog(QDialog):
    """Modal dialog for selecting which sheets to translate.

    Parameters
    ----------
    file_sheets : dict[str, list[str]]
        Mapping of file_path → list of sheet names.
    parent : QWidget | None
        Parent widget.

    After exec(), call ``get_selection()`` to retrieve the result.
    """

    def __init__(
        self,
        file_sheets: dict[str, list[str]],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Chọn Sheet Cần Dịch")
        self.setMinimumSize(560, 420)
        self.resize(640, 500)
        self.setModal(True)

        self._file_sheets = file_sheets
        self._result: dict[str, list[str]] | None = None

        self.setStyleSheet(_DIALOG_STYLESHEET)
        self._build_ui()
        self._populate_tree()
        self._update_summary()

    # ── Build UI ─────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # ── Header ───────────────────────────────────────────────
        title = QLabel("Chọn Sheet Cần Dịch")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)

        subtitle = QLabel("Chọn sheet cần dịch cho mỗi file")
        subtitle.setObjectName("dialogSubtitle")
        layout.addWidget(subtitle)

        # Separator
        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.HLine)
        layout.addWidget(sep)

        # ── Toolbar: Select All / Clear All + Summary ────────────
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        btn_all = QPushButton("✓  Chọn tất cả")
        btn_all.setObjectName("btnSelectAll")
        btn_all.setCursor(Qt.PointingHandCursor)
        btn_all.clicked.connect(self._select_all)

        btn_clear = QPushButton("✗  Bỏ chọn")
        btn_clear.setObjectName("btnClearAll")
        btn_clear.setCursor(Qt.PointingHandCursor)
        btn_clear.clicked.connect(self._clear_all)

        self.lbl_summary = QLabel("Đã chọn: 0 sheet")
        self.lbl_summary.setObjectName("summaryLabel")

        toolbar.addWidget(btn_all)
        toolbar.addWidget(btn_clear)
        toolbar.addStretch()
        toolbar.addWidget(self.lbl_summary)

        layout.addLayout(toolbar)

        # ── Tree ─────────────────────────────────────────────────
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["File / Sheet"])
        self.tree.setColumnCount(1)
        self.tree.setRootIsDecorated(True)
        self.tree.setAnimated(True)
        self.tree.setIndentation(24)
        self.tree.header().setStretchLastSection(True)
        self.tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tree.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.tree, stretch=1)

        # ── Bottom buttons ───────────────────────────────────────
        sep2 = QFrame()
        sep2.setObjectName("separator")
        sep2.setFrameShape(QFrame.HLine)
        layout.addWidget(sep2)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        self.btn_cancel = QPushButton("Hủy")
        self.btn_cancel.setObjectName("btnCancel")
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_ok = QPushButton("Bắt đầu dịch")
        self.btn_ok.setObjectName("btnOk")
        self.btn_ok.setCursor(Qt.PointingHandCursor)
        self.btn_ok.setDefault(True)
        self.btn_ok.clicked.connect(self._on_ok)

        btn_row.addStretch()
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_ok)

        layout.addLayout(btn_row)

    # ── Populate tree ────────────────────────────────────────────

    def _populate_tree(self):
        """Build the tree and restore previous selections."""
        saved = _load_saved_selections()
        self.tree.blockSignals(True)

        for fp, sheets in self._file_sheets.items():
            file_item = QTreeWidgetItem(self.tree)
            file_item.setText(0, f"📄  {Path(fp).name}")
            file_item.setData(0, Qt.UserRole, fp)
            file_item.setFlags(
                file_item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsAutoTristate
            )

            # Decide which sheets to pre-check
            saved_sheets = saved.get(fp, None)

            for sn in sheets:
                sheet_item = QTreeWidgetItem(file_item)
                sheet_item.setText(0, f"  {sn}")
                sheet_item.setData(0, Qt.UserRole + 1, sn)  # store clean name
                sheet_item.setFlags(
                    sheet_item.flags() | Qt.ItemIsUserCheckable
                )
                # Pre-check: if we have saved selections, use them; else check all
                if saved_sheets is None:
                    sheet_item.setCheckState(0, Qt.Checked)
                elif sn in saved_sheets:
                    sheet_item.setCheckState(0, Qt.Checked)
                else:
                    sheet_item.setCheckState(0, Qt.Unchecked)

            file_item.setExpanded(True)
            # Update parent check state
            self._update_parent_check(file_item)

        self.tree.blockSignals(False)
        self.tree.setFocus()

    # ── Event handlers ───────────────────────────────────────────

    def _on_item_changed(self, item: QTreeWidgetItem, column: int):
        """When a file item is toggled, propagate to child sheets."""
        if column != 0:
            return

        self.tree.blockSignals(True)

        if item.parent() is None:
            # File-level toggle → propagate to all children
            state = item.checkState(0)
            if state != Qt.PartiallyChecked:
                for i in range(item.childCount()):
                    item.child(i).setCheckState(0, state)
        else:
            # Sheet toggle → update parent
            self._update_parent_check(item.parent())

        self.tree.blockSignals(False)
        self._update_summary()

    def _update_parent_check(self, parent_item: QTreeWidgetItem):
        """Set parent check state based on children (tri-state)."""
        checked = 0
        total = parent_item.childCount()
        for i in range(total):
            if parent_item.child(i).checkState(0) == Qt.Checked:
                checked += 1

        if checked == 0:
            parent_item.setCheckState(0, Qt.Unchecked)
        elif checked == total:
            parent_item.setCheckState(0, Qt.Checked)
        else:
            parent_item.setCheckState(0, Qt.PartiallyChecked)

    def _update_summary(self):
        count = self._count_selected()
        self.lbl_summary.setText(f"Đã chọn: {count} sheet")

    def _count_selected(self) -> int:
        count = 0
        for i in range(self.tree.topLevelItemCount()):
            file_item = self.tree.topLevelItem(i)
            for j in range(file_item.childCount()):
                if file_item.child(j).checkState(0) == Qt.Checked:
                    count += 1
        return count

    # ── Select / Clear all ───────────────────────────────────────

    def _select_all(self):
        self.tree.blockSignals(True)
        for i in range(self.tree.topLevelItemCount()):
            fi = self.tree.topLevelItem(i)
            fi.setCheckState(0, Qt.Checked)
            for j in range(fi.childCount()):
                fi.child(j).setCheckState(0, Qt.Checked)
        self.tree.blockSignals(False)
        self._update_summary()

    def _clear_all(self):
        self.tree.blockSignals(True)
        for i in range(self.tree.topLevelItemCount()):
            fi = self.tree.topLevelItem(i)
            fi.setCheckState(0, Qt.Unchecked)
            for j in range(fi.childCount()):
                fi.child(j).setCheckState(0, Qt.Unchecked)
        self.tree.blockSignals(False)
        self._update_summary()

    # ── OK / Validate ────────────────────────────────────────────

    def _on_ok(self):
        """Validate and accept."""
        if self._count_selected() == 0:
            QMessageBox.warning(
                self, "Cảnh báo", "Bạn chưa chọn sheet nào."
            )
            return  # Keep dialog open

        # Build result
        self._result = {}
        save_data: dict[str, list[str]] = {}

        for i in range(self.tree.topLevelItemCount()):
            file_item = self.tree.topLevelItem(i)
            fp = file_item.data(0, Qt.UserRole)
            selected: list[str] = []

            for j in range(file_item.childCount()):
                child = file_item.child(j)
                sheet_name = child.data(0, Qt.UserRole + 1) or child.text(0).strip()

                if child.checkState(0) == Qt.Checked:
                    selected.append(sheet_name)

            self._result[fp] = selected if selected else None
            save_data[fp] = selected

        # Persist
        _save_selections(save_data)
        self.accept()

    # ── Public API ───────────────────────────────────────────────

    def get_selection(self) -> dict[str, list[str]] | None:
        """Get the selected sheets. None if dialog was cancelled."""
        return self._result
