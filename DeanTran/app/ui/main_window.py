"""
MainWindow – settings-driven application shell with tab navigation.
"""
from PySide6.QtCore import Qt, QEvent
from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QLabel, QVBoxLayout, QWidget,
    QMenuBar, QMessageBox, QApplication, QScrollArea,
)
from PySide6.QtGui import QAction, QShortcut, QKeySequence

from app.settings.settings_manager import settings

from app.core.auth_manager import AuthManager
from app.core.telegram_bot import TelegramBot
from app.ui.excel_tab import ExcelTab
from app.ui.ppt_tab import PptTab
from app.ui.word_tab import WordTab
from app.ui.pdf_tab import PdfTab
from app.ui.pdf_convert_tab import PdfConvertTab
from app.ui.api_settings_panel import ApiSettingsPanel
from app.ui.prompt_editor_panel import PromptEditorPanel
from app.ui.admin_login_dialog import AdminLoginDialog
from app.version import APP_TITLE


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1100, 750)

        # Auth
        self.auth = AuthManager()

        # Telegram bot (remote admin)
        self._telegram_bot = TelegramBot()
        self._telegram_bot.start()

        # Central tabs
        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.setCentralWidget(self.tabs)

        # Load Zoom level
        self.zoom_pct = settings.get("ui.zoom_level", 100)

        # Status Bar Zoom Label
        self.lbl_zoom = QLabel(f"Zoom: {self.zoom_pct}%")
        self.lbl_zoom.setStyleSheet("font-weight: bold; margin-right: 10px; color: #475569;")
        self.statusBar().addPermanentWidget(self.lbl_zoom)

        # Apply Zoom Stylesheet
        self._apply_zoom(self.zoom_pct)

        # Install global event filter to capture Ctrl+Wheel events application-wide
        QApplication.instance().installEventFilter(self)

        # Register Zoom Shortcuts
        self.shortcut_in = QShortcut(QKeySequence("Ctrl+="), self)
        self.shortcut_in.activated.connect(lambda: self._apply_zoom(self.zoom_pct + 10))
        self.shortcut_in_plus = QShortcut(QKeySequence("Ctrl++"), self)
        self.shortcut_in_plus.activated.connect(lambda: self._apply_zoom(self.zoom_pct + 10))
        self.shortcut_out = QShortcut(QKeySequence("Ctrl+-"), self)
        self.shortcut_out.activated.connect(lambda: self._apply_zoom(self.zoom_pct - 10))
        self.shortcut_reset = QShortcut(QKeySequence("Ctrl+0"), self)
        self.shortcut_reset.activated.connect(lambda: self._apply_zoom(100))

        # 1. Excel Tab
        self.excel_tab = ExcelTab()
        self.tabs.addTab(self._make_scrollable(self.excel_tab), "📊 Excel")

        # 2. PPT Tab
        self.ppt_tab = PptTab()
        self.tabs.addTab(self._make_scrollable(self.ppt_tab), "📽️ PowerPoint")

        # 3. Word Tab
        self.word_tab = WordTab()
        self.tabs.addTab(self._make_scrollable(self.word_tab), "📝 Word")

        # 4. PDF Tab
        self.pdf_tab = PdfTab()
        self.tabs.addTab(self._make_scrollable(self.pdf_tab), "📄 PDF Export")

        # 5. PDF Convert Tab
        self.pdf_convert_tab = PdfConvertTab()
        self.tabs.addTab(self._make_scrollable(self.pdf_convert_tab), "🔄 Convert PDF")

        # 6. Prompt Editor tab
        self.prompt_panel = PromptEditorPanel()
        self.tabs.addTab(self._make_scrollable(self.prompt_panel), "🧠 Prompt Editor")

        # 6. Admin API tab (hidden until login)
        self.api_panel = ApiSettingsPanel()
        self.api_tab_index = self.tabs.addTab(self._make_scrollable(self.api_panel), "⚙️ API Settings")
        self.tabs.setTabVisible(self.api_tab_index, False)

        # Menu bar
        self._build_menu()

    def _build_menu(self):
        menu = self.menuBar()
        admin_menu = menu.addMenu("Admin")

        self.act_login = QAction("Login as Admin", self)
        self.act_login.triggered.connect(self._on_admin_login)
        admin_menu.addAction(self.act_login)

        self.act_logout = QAction("Logout", self)
        self.act_logout.triggered.connect(self._on_admin_logout)
        self.act_logout.setEnabled(False)
        admin_menu.addAction(self.act_logout)

    def _on_admin_login(self):
        dlg = AdminLoginDialog(self.auth, parent=self)
        if dlg.exec():
            if self.auth.is_admin:
                self.tabs.setTabVisible(self.api_tab_index, True)
                self.act_login.setEnabled(False)
                self.act_logout.setEnabled(True)
                self.setWindowTitle(
                    f"DeanTran Translation App - Milestone 3  [Admin: {self.auth.current_user}]"
                )

    def _on_tab_changed(self, index: int):
        widget = self.tabs.widget(index)
        if isinstance(widget, QScrollArea):
            widget = widget.widget()
        ppt_tab = getattr(self, "ppt_tab", None)
        word_tab = getattr(self, "word_tab", None)
        pdf_tab = getattr(self, "pdf_tab", None)
        api_panel = getattr(self, "api_panel", None)
        if ppt_tab is not None and widget is ppt_tab:
            self.ppt_tab.reload_from_settings()
        elif word_tab is not None and widget is word_tab:
            self.word_tab.reload_from_settings()
        elif pdf_tab is not None and widget is pdf_tab:
            self.pdf_tab._load_settings()
        elif api_panel is not None and widget is api_panel:
            self.api_panel.reload_from_settings()

    def _on_admin_logout(self):
        self.auth.logout()
        self.tabs.setTabVisible(self.api_tab_index, False)
        self.act_login.setEnabled(True)
        self.act_change_pwd.setEnabled(False)
        self.act_logout.setEnabled(False)
        self.setWindowTitle("DeanTrans Translation App - Milestone 3")

    def closeEvent(self, event):
        """Stop Telegram bot on app close."""
        self._telegram_bot.stop()
        super().closeEvent(event)

    def _apply_zoom(self, zoom_pct: int):
        self.zoom_pct = min(150, max(70, zoom_pct))
        settings.set("ui.zoom_level", self.zoom_pct)
        self.setStyleSheet(self._build_stylesheet(self.zoom_pct))
        if hasattr(self, "lbl_zoom"):
            self.lbl_zoom.setText(f"Zoom: {self.zoom_pct}%")

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Wheel:
            if event.modifiers() == Qt.ControlModifier:
                delta = event.angleDelta().y()
                if delta > 0:
                    self._apply_zoom(self.zoom_pct + 10)
                elif delta < 0:
                    self._apply_zoom(self.zoom_pct - 10)
                return True
        return super().eventFilter(obj, event)

    def _build_stylesheet(self, zoom_pct: int) -> str:
        scale = zoom_pct / 100.0
        base_font_sz = max(8, int(12 * scale))
        tab_font_sz = max(9, int(13 * scale))
        group_font_sz = max(9, int(13 * scale))
        return f"""
            /* ── Base ────────────────────────────────────────── */
            QMainWindow {{
                background-color: #f0f4f8;
                font-family: 'Segoe UI', 'Arial', sans-serif;
            }}

            /* ── Tab Bar ─────────────────────────────────────── */
            QTabWidget::pane {{
                border: none;
                background: #f0f4f8;
            }}
            QTabBar::tab {{
                background: #e2e8f0;
                color: #475569;
                padding: {int(10 * scale)}px {int(22 * scale)}px;
                margin-right: 2px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: 600;
                font-size: {tab_font_sz}px;
                min-width: {int(100 * scale)}px;
            }}
            QTabBar::tab:selected {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3b82f6, stop:1 #2563eb);
                color: white;
            }}
            QTabBar::tab:hover:!selected {{
                background: #cbd5e1;
                color: #1e293b;
            }}

            /* ── GroupBox ─────────────────────────────────────── */
            QGroupBox {{
                font-weight: bold;
                font-size: {group_font_sz}px;
                color: #1e40af;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 18px;
                background: #ffffff;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 14px;
                padding: 2px 10px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3b82f6, stop:1 #60a5fa);
                color: white;
                border-radius: 4px;
            }}

            /* ── Buttons ─────────────────────────────────────── */
            QPushButton {{
                background-color: #e2e8f0;
                color: #334155;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 6px 14px;
                font-size: {base_font_sz}px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: #3b82f6;
                color: white;
                border-color: #2563eb;
            }}
            QPushButton:pressed {{
                background-color: #1d4ed8;
                color: white;
            }}
            QPushButton:disabled {{
                background-color: #f1f5f9;
                color: #94a3b8;
                border-color: #e2e8f0;
            }}

            /* ── ComboBox ────────────────────────────────────── */
            QComboBox {{
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 5px 10px;
                background: white;
                font-size: {base_font_sz}px;
                min-height: {int(22 * scale)}px;
            }}
            QComboBox:hover {{
                border-color: #3b82f6;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QComboBox QAbstractItemView {{
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                background: white;
                selection-background-color: #dbeafe;
                selection-color: #1e40af;
            }}

            /* ── SpinBox ─────────────────────────────────────── */
            QSpinBox, QDoubleSpinBox {{
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 5px 8px;
                background: white;
                font-size: {base_font_sz}px;
                min-height: {int(22 * scale)}px;
            }}
            QSpinBox:hover, QDoubleSpinBox:hover {{
                border-color: #3b82f6;
            }}

            /* ── LineEdit ────────────────────────────────────── */
            QLineEdit {{
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 5px 10px;
                background: white;
                font-size: {base_font_sz}px;
                min-height: {int(22 * scale)}px;
            }}
            QLineEdit:hover, QLineEdit:focus {{
                border-color: #3b82f6;
            }}

            /* ── CheckBox / RadioButton ──────────────────────── */
            QCheckBox, QRadioButton {{
                font-size: {base_font_sz}px;
                spacing: 6px;
                color: #334155;
            }}
            QCheckBox:hover, QRadioButton:hover {{
                color: #1e40af;
            }}

            /* ── ProgressBar ─────────────────────────────────── */
            QProgressBar {{
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                text-align: center;
                font-size: {base_font_sz}px;
                font-weight: bold;
                color: #334155;
                background: #f1f5f9;
                min-height: {int(20 * scale)}px;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3b82f6, stop:0.5 #60a5fa, stop:1 #3b82f6);
                border-radius: 5px;
            }}

            /* ── ListWidget ──────────────────────────────────── */
            QListWidget {{
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                background: white;
                font-size: {base_font_sz}px;
                padding: 4px;
            }}
            QListWidget::item {{
                padding: 4px 8px;
                border-radius: 4px;
            }}
            QListWidget::item:selected {{
                background: #dbeafe;
                color: #1e40af;
            }}
            QListWidget::item:hover {{
                background: #f1f5f9;
            }}

            /* ── Label ───────────────────────────────────────── */
            QLabel {{
                font-size: {base_font_sz}px;
                color: #334155;
            }}

            /* ── ScrollBar ───────────────────────────────────── */
            QScrollBar:vertical {{
                width: 10px;
                background: #f1f5f9;
                border-radius: 5px;
            }}
            QScrollBar::handle:vertical {{
                background: #94a3b8;
                border-radius: 5px;
                min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: #64748b;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}

            /* ── Menu ────────────────────────────────────────── */
            QMenuBar {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #1e3a5f, stop:1 #1e293b);
                color: white;
                font-size: {tab_font_sz}px;
                padding: 2px;
            }}
            QMenuBar::item {{
                padding: 6px 14px;
                border-radius: 4px;
            }}
            QMenuBar::item:selected {{
                background: #3b82f6;
            }}
            QMenu {{
                background: white;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 4px;
            }}
            QMenu::item {{
                padding: 8px 24px;
                border-radius: 4px;
            }}
            QMenu::item:selected {{
                background: #dbeafe;
                color: #1e40af;
            }}
        """

    def _make_scrollable(self, widget: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(widget)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        return scroll

    def _placeholder(self, text):
        widget = QWidget()
        layout = QVBoxLayout()
        label = QLabel(text)
        label.setStyleSheet("font-size: 20px; color: gray;")
        layout.addWidget(label)
        widget.setLayout(layout)
        return widget
