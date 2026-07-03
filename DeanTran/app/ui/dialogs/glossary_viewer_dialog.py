"""
GlossaryViewerDialog – Preview glossary/dictionary contents with search.

Supports multi-language Excel files (3+ columns with language headers).
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QTableWidget, QTableWidgetItem, QPushButton, QHeaderView,
)


class GlossaryViewerDialog(QDialog):
    """Dialog to view glossary contents loaded from an Excel or JSON file."""

    def __init__(self, glossary_path: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("📖 Xem Từ Điển")
        self.setMinimumSize(650, 500)
        self.resize(750, 550)

        self._headers: list[str] = []
        self._rows: list[list[str]] = []
        self._load_glossary(glossary_path)
        self._build_ui()
        self._populate_table()

    def _load_glossary(self, path_str: str) -> None:
        path = Path(path_str)
        if not path.exists():
            return

        suffix = path.suffix.lower()
        if suffix == ".json":
            self._load_json(path)
        elif suffix in (".xlsx", ".xls"):
            self._load_xlsx(path)

    def _load_json(self, path: Path) -> None:
        import json
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                self._headers = ["Source", "Target"]
                self._rows = [[str(k), str(v)] for k, v in data.items()]
        except Exception:
            pass

    def _load_xlsx(self, path: Path) -> None:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(path), read_only=True)
            ws = wb.active
            all_rows = list(ws.iter_rows(values_only=True))
            wb.close()

            if not all_rows:
                return

            # Use row 1 as headers
            raw_headers = all_rows[0]
            self._headers = [str(h) if h else f"Col {i+1}" for i, h in enumerate(raw_headers)]

            # Data rows
            for row in all_rows[1:]:
                row_data = []
                for i in range(len(self._headers)):
                    val = row[i] if i < len(row) else None
                    row_data.append(str(val) if val else "")
                if any(cell.strip() for cell in row_data):  # skip fully empty rows
                    self._rows.append(row_data)
        except Exception:
            pass

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # Header
        n_entries = len(self._rows)
        n_langs = len(self._headers)
        header = QLabel(f"📖 Từ Điển — {n_entries} mục · {n_langs} ngôn ngữ")
        header.setStyleSheet(
            "font-size: 16px; font-weight: bold; color: #1e40af; padding: 4px 0;"
        )
        layout.addWidget(header)
        self._lbl_header = header

        # Language info
        if self._headers:
            lang_info = QLabel("🌐 " + " ↔ ".join(self._headers))
            lang_info.setStyleSheet("font-size: 13px; color: #3b82f6; padding: 2px 0;")
            layout.addWidget(lang_info)

        # Search bar
        search_row = QHBoxLayout()
        search_row.setSpacing(8)
        lbl_search = QLabel("🔍")
        lbl_search.setStyleSheet("font-size: 16px;")
        search_row.addWidget(lbl_search)

        self.le_search = QLineEdit()
        self.le_search.setPlaceholderText("Tìm kiếm từ...")
        self.le_search.setStyleSheet(
            "padding: 8px 12px; border: 2px solid #e2e8f0; border-radius: 8px; "
            "font-size: 13px; background: #f8fafc;"
        )
        self.le_search.textChanged.connect(self._on_search)
        search_row.addWidget(self.le_search, stretch=1)
        layout.addLayout(search_row)

        # Table
        col_count = len(self._headers) if self._headers else 2
        self.table = QTableWidget()
        self.table.setColumnCount(col_count)
        self.table.setHorizontalHeaderLabels(self._headers if self._headers else ["Source", "Target"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setStyleSheet(
            "QTableWidget { border: 1px solid #e2e8f0; border-radius: 6px; "
            "   gridline-color: #f1f5f9; font-size: 13px; }"
            "QTableWidget::item { padding: 6px 10px; }"
            "QTableWidget::item:alternate { background-color: #f8fafc; }"
            "QTableWidget::item:selected { background-color: #dbeafe; color: #1e40af; }"
            "QHeaderView::section { background-color: #3b82f6; color: white; "
            "   padding: 8px; font-weight: bold; font-size: 13px; border: none; }"
        )
        layout.addWidget(self.table, stretch=1)

        # Footer
        btn_row = QHBoxLayout()
        self.lbl_count = QLabel(f"Hiển thị: {len(self._rows)} / {len(self._rows)}")
        self.lbl_count.setStyleSheet("color: #64748b; font-size: 12px;")
        btn_row.addWidget(self.lbl_count)
        btn_row.addStretch()

        btn_close = QPushButton("Đóng")
        btn_close.setStyleSheet(
            "padding: 8px 24px; background-color: #3b82f6; color: white; "
            "border-radius: 6px; font-weight: bold; font-size: 13px;"
        )
        btn_close.clicked.connect(self.close)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

    def _populate_table(self, filter_text: str = "") -> None:
        filter_lower = filter_text.strip().lower()

        if filter_lower:
            filtered = [
                row for row in self._rows
                if any(filter_lower in cell.lower() for cell in row)
            ]
        else:
            filtered = self._rows

        self.table.setRowCount(len(filtered))
        for row_idx, row_data in enumerate(filtered):
            for col_idx, cell in enumerate(row_data):
                item = QTableWidgetItem(cell)
                self.table.setItem(row_idx, col_idx, item)

        self.lbl_count.setText(f"Hiển thị: {len(filtered)} / {len(self._rows)}")

    def _on_search(self, text: str) -> None:
        self._populate_table(text)
