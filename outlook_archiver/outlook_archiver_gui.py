"""Outlook 附件自動分類歸檔工具 — PySide6 圖形介面。"""

from __future__ import annotations

import importlib.util
import logging
from pathlib import Path
import re
import sys

from PySide6.QtCore import QSettings, Qt, QThread, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QFont, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)


def _load_core_module():
    """安全載入核心模組 outlook_archiver（支援直譯與 PyInstaller 打包模式）。"""
    if getattr(sys, "frozen", False):
        import outlook_archiver as mod

        return mod

    module_path = Path(__file__).resolve().parent / "outlook_archiver.py"
    spec = importlib.util.spec_from_file_location("outlook_archiver_core", module_path)
    if not spec or not spec.loader:
        raise ImportError(f"無法載入核心模組：{module_path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


oa = _load_core_module()


class QtLogHandler(logging.Handler):
    """將 Python logging 訊息透過 Qt Signal 轉送至主視窗顯示。"""

    def __init__(self, emit_callback) -> None:
        super().__init__()
        self._emit_callback = emit_callback

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            self._emit_callback(record.levelno, msg)
        except Exception:
            pass


class ArchiveWorker(QThread):
    """在背景執行緒執行 Outlook COM 掃描與附件下載，避免凍結介面。"""

    log_received = Signal(int, str)
    progress_updated = Signal(dict, str)
    finished_run = Signal(dict, str)

    def __init__(
        self,
        *,
        dry_run: bool,
        days_back: int | None,
        account_email: str | None,
        output_root: Path,
        ignore_inline: bool,
        use_datetime_subfolder: bool,
    ) -> None:
        super().__init__()
        self.dry_run = dry_run
        self.days_back = days_back
        self.account_email = account_email
        self.output_root = output_root
        self.ignore_inline = ignore_inline
        self.use_datetime_subfolder = use_datetime_subfolder
        self._cancel_flag = False

    def request_cancel(self) -> None:
        self._cancel_flag = True

    def run(self) -> None:
        import pythoncom  # type: ignore

        pythoncom.CoInitialize()
        qt_handler = QtLogHandler(lambda lvl, msg: self.log_received.emit(lvl, msg))
        logger = oa.setup_logging(oa.LOG_FILE_PATH, verbose=False, extra_handlers=[qt_handler])

        stats = {
            "scanned_mails": 0,
            "skipped_checked": 0,
            "saved_attachments": 0,
            "skipped_attachments": 0,
            "errors": 0,
        }
        error_msg = ""

        mode_text = "模擬測試 (Dry-Run)" if self.dry_run else "正式歸檔"
        logger.info("=== [%s] 開始執行 (days_back=%s) ===", mode_text, self.days_back)

        state_store = oa.StateStore(oa.STATE_DB_PATH)
        try:
            rules = oa.parse_rules_markdown(oa.RULES_FILE)
            logger.info("已載入 %d 條分類規則：%s", len(rules), [r.folder_name for r in rules])

            _, namespace = oa.connect_outlook(logger)
            folder = oa.resolve_target_folder(namespace, self.account_email, oa.TARGET_FOLDER_NAME)
            logger.info("已定位目標資料夾：%s", getattr(folder, "FolderPath", getattr(folder, "Name", "")))

            stats = oa.process_mailbox(
                folder=folder,
                rules=rules,
                state_store=state_store,
                output_root=self.output_root,
                logger=logger,
                dry_run=self.dry_run,
                days_back=self.days_back,
                ignore_inline_images=self.ignore_inline,
                use_datetime_subfolder=self.use_datetime_subfolder,
                cancel_requested=lambda: self._cancel_flag,
                progress_callback=lambda s, subj: self.progress_updated.emit(s, subj),
            )
            logger.info(
                "=== 執行結束 | 掃描: %d | 略過已檢查: %d | 儲存附件: %d | 排除: %d | 錯誤: %d ===",
                stats["scanned_mails"],
                stats["skipped_checked"],
                stats["saved_attachments"],
                stats["skipped_attachments"],
                stats["errors"],
            )
        except Exception as exc:
            error_msg = str(exc)
            logger.exception("執行過程發生錯誤：%s", exc)
        finally:
            state_store.close()
            pythoncom.CoUninitialize()

        self.finished_run.emit(stats, error_msg)


_THEMES: dict[str, dict[str, str]] = {
    "light": {
        "window_bg": "#F8FAFC",
        "surface_bg": "#FFFFFF",
        "input_bg": "#FFFFFF",
        "border": "#E2E8F0",
        "input_border": "#CBD5E1",
        "text": "#1E293B",
        "title_text": "#0F172A",
        "muted_text": "#64748B",
        "btn_bg": "#F1F5F9",
        "btn_hover": "#E2E8F0",
        "btn_disabled_bg": "#F8FAFC",
        "btn_disabled_text": "#94A3B8",
        "primary": "#2563EB",
        "primary_border": "#1D4ED8",
        "primary_hover": "#1D4ED8",
        "danger": "#EF4444",
        "danger_border": "#DC2626",
        "danger_hover": "#DC2626",
        "tab_bg": "#E2E8F0",
        "tab_text": "#475569",
        "tab_selected_bg": "#FFFFFF",
        "tab_selected_text": "#2563EB",
        "table_header_bg": "#F1F5F9",
        "table_selection_bg": "#DBEAFE",
        "table_selection_text": "#0F172A",
        "hint_bg": "#EFF6FF",
        "hint_border": "#BFDBFE",
        "hint_text": "#475569",
        "log_info": "#1E293B",
        "log_success": "#15803D",
        "log_warning": "#D97706",
        "log_error": "#DC2626",
    },
    "dark": {
        "window_bg": "#0F172A",
        "surface_bg": "#1E293B",
        "input_bg": "#0F172A",
        "border": "#334155",
        "input_border": "#475569",
        "text": "#E2E8F0",
        "title_text": "#F8FAFC",
        "muted_text": "#94A3B8",
        "btn_bg": "#334155",
        "btn_hover": "#475569",
        "btn_disabled_bg": "#1E293B",
        "btn_disabled_text": "#64748B",
        "primary": "#3B82F6",
        "primary_border": "#2563EB",
        "primary_hover": "#2563EB",
        "danger": "#EF4444",
        "danger_border": "#DC2626",
        "danger_hover": "#DC2626",
        "tab_bg": "#1E293B",
        "tab_text": "#94A3B8",
        "tab_selected_bg": "#0F172A",
        "tab_selected_text": "#60A5FA",
        "table_header_bg": "#334155",
        "table_selection_bg": "#1D4ED8",
        "table_selection_text": "#F8FAFC",
        "hint_bg": "#172554",
        "hint_border": "#1E40AF",
        "hint_text": "#BFDBFE",
        "log_info": "#E2E8F0",
        "log_success": "#4ADE80",
        "log_warning": "#FBBF24",
        "log_error": "#F87171",
    },
}


def _detect_system_theme() -> str:
    """偵測 Windows / Qt 系統色彩模式，回傳 'dark' 或 'light'。"""
    app = QApplication.instance()
    if app is not None:
        scheme = app.styleHints().colorScheme()
        if scheme == Qt.ColorScheme.Dark:
            return "dark"
        if scheme == Qt.ColorScheme.Light:
            return "light"

    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if int(val) == 1 else "dark"
    except Exception:
        return "light"


def _build_stylesheet(t: dict[str, str]) -> str:
    """依主題色票生成主視窗 QSS 樣式表。"""
    return f"""
        QMainWindow, QDialog, QMessageBox {{
            background-color: {t["window_bg"]};
        }}
        QWidget {{
            font-family: "Microsoft JhengHei UI", "Segoe UI", sans-serif;
            font-size: 13px;
            color: {t["text"]};
        }}
        QGroupBox {{
            background-color: {t["surface_bg"]};
            border: 1px solid {t["border"]};
            border-radius: 8px;
            margin-top: 10px;
            padding-top: 12px;
            font-weight: bold;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 6px;
            color: {t["text"]};
        }}
        QLineEdit, QSpinBox, QComboBox, QListWidget, QTableWidget, QPlainTextEdit {{
            background-color: {t["input_bg"]};
            color: {t["text"]};
            border: 1px solid {t["input_border"]};
            border-radius: 6px;
            padding: 5px 8px;
            selection-background-color: {t["table_selection_bg"]};
            selection-color: {t["table_selection_text"]};
        }}
        QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QListWidget:focus {{
            border-color: {t["primary"]};
        }}
        QComboBox QAbstractItemView {{
            background-color: {t["surface_bg"]};
            color: {t["text"]};
            border: 1px solid {t["input_border"]};
            selection-background-color: {t["table_selection_bg"]};
            selection-color: {t["table_selection_text"]};
        }}
        QHeaderView::section {{
            background-color: {t["table_header_bg"]};
            color: {t["text"]};
            border: none;
            border-bottom: 1px solid {t["border"]};
            padding: 6px 8px;
            font-weight: 600;
        }}
        QTableWidget {{
            gridline-color: {t["border"]};
        }}
        QPushButton {{
            background-color: {t["btn_bg"]};
            color: {t["text"]};
            border: 1px solid {t["input_border"]};
            border-radius: 6px;
            padding: 6px 14px;
            font-weight: 600;
        }}
        QPushButton:hover {{
            background-color: {t["btn_hover"]};
        }}
        QPushButton:disabled {{
            color: {t["btn_disabled_text"]};
            background-color: {t["btn_disabled_bg"]};
        }}
        QPushButton#PrimaryBtn {{
            background-color: {t["primary"]};
            border-color: {t["primary_border"]};
            color: #FFFFFF;
        }}
        QPushButton#PrimaryBtn:hover {{
            background-color: {t["primary_hover"]};
        }}
        QPushButton#DangerBtn {{
            background-color: {t["danger"]};
            border-color: {t["danger_border"]};
            color: #FFFFFF;
        }}
        QPushButton#DangerBtn:hover {{
            background-color: {t["danger_hover"]};
        }}
        QTabWidget::pane {{
            border: 1px solid {t["border"]};
            border-radius: 8px;
            background: {t["window_bg"]};
        }}
        QTabBar::tab {{
            background: {t["tab_bg"]};
            color: {t["tab_text"]};
            padding: 8px 18px;
            margin-right: 4px;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
            font-weight: 600;
        }}
        QTabBar::tab:selected {{
            background: {t["tab_selected_bg"]};
            color: {t["tab_selected_text"]};
        }}
    """


class StatCard(QFrame):
    """簡潔統計卡片元件。"""

    def __init__(self, title: str, accent_color: str) -> None:
        super().__init__()
        self._accent_color = accent_color
        self.setObjectName("StatCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(2)

        self.title_label = QLabel(title)
        self.value_label = QLabel("0")
        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        self.apply_theme(_THEMES["light"])

    def apply_theme(self, theme: dict[str, str]) -> None:
        self.setStyleSheet(
            f"""
            QFrame#StatCard {{
                background-color: {theme["surface_bg"]};
                border: 1px solid {theme["border"]};
                border-left: 4px solid {self._accent_color};
                border-radius: 8px;
                padding: 8px 12px;
            }}
            """
        )
        self.title_label.setStyleSheet(f"color: {theme['muted_text']}; font-size: 12px;")
        self.value_label.setStyleSheet(
            f"color: {theme['title_text']}; font-size: 22px; font-weight: bold;"
        )

    def set_value(self, value: int) -> None:
        self.value_label.setText(str(value))


class MainWindow(QMainWindow):
    """Outlook 附件自動分類歸檔工具主視窗。"""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"Outlook 附件自動分類歸檔工具 {oa.APP_VERSION}")
        self.resize(1040, 720)
        self.worker: ArchiveWorker | None = None
        self.rules: list[oa.ClassificationRule] = []
        self._syncing_rule_form = False

        self.settings = QSettings("OutlookArchiver", "CADTool")
        saved_mode = str(self.settings.value("ui/theme_mode", "system"))
        self.theme_mode = saved_mode if saved_mode in ("system", "light", "dark") else "system"
        self._active_theme = _THEMES["light"]

        self._build_ui()
        self._apply_theme()

        app = QApplication.instance()
        if app is not None:
            app.styleHints().colorSchemeChanged.connect(self._on_system_color_scheme_changed)

        self._load_rules_into_editor()
        self._load_history_table()

    def _resolve_effective_theme_name(self) -> str:
        if self.theme_mode in ("light", "dark"):
            return self.theme_mode
        return _detect_system_theme()

    def _apply_theme(self) -> None:
        app = QApplication.instance()
        if app is not None:
            hints = app.styleHints()
            try:
                if self.theme_mode == "dark" and hasattr(hints, "setColorScheme"):
                    hints.setColorScheme(Qt.ColorScheme.Dark)
                elif self.theme_mode == "light" and hasattr(hints, "setColorScheme"):
                    hints.setColorScheme(Qt.ColorScheme.Light)
                elif self.theme_mode == "system" and hasattr(hints, "unsetColorScheme"):
                    hints.unsetColorScheme()
            except Exception:
                pass

        effective = self._resolve_effective_theme_name()
        theme = _THEMES[effective]
        self._active_theme = theme

        self.setStyleSheet(_build_stylesheet(theme))
        self.title_label.setStyleSheet(
            f"font-size: 19px; font-weight: bold; color: {theme['title_text']};"
        )
        self.status_label.setStyleSheet(f"color: {theme['muted_text']}; font-size: 12px;")
        self.hint_label.setStyleSheet(
            f"color: {theme['hint_text']}; background: {theme['hint_bg']}; "
            f"border: 1px solid {theme['hint_border']}; border-radius: 6px; padding: 8px;"
        )
        for card in self.stat_cards:
            card.apply_theme(theme)

    def _on_theme_changed(self, _index: int) -> None:
        selected_mode = str(self.theme_combo.currentData() or "system")
        self.theme_mode = selected_mode
        self.settings.setValue("ui/theme_mode", selected_mode)
        self._apply_theme()

    def _on_system_color_scheme_changed(self, _scheme) -> None:
        if self.theme_mode == "system":
            self._apply_theme()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(14, 12, 14, 12)
        root_layout.setSpacing(10)

        # 頂部標題列
        header_layout = QHBoxLayout()
        title_col = QVBoxLayout()
        self.title_label = QLabel(f"Outlook 附件自動分類歸檔工具 {oa.APP_VERSION}")
        self.status_label = QLabel("就緒（使用 Classic Outlook COM 介面）")
        title_col.addWidget(self.title_label)
        title_col.addWidget(self.status_label)

        theme_label = QLabel("主題：")
        self.theme_combo = QComboBox()
        self.theme_combo.addItem("系統設定", "system")
        self.theme_combo.addItem("淺色", "light")
        self.theme_combo.addItem("深色", "dark")
        idx = self.theme_combo.findData(self.theme_mode)
        self.theme_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)

        open_out_btn = QPushButton("開啟歸檔根目錄")
        open_out_btn.clicked.connect(self._open_output_dir)
        open_log_btn = QPushButton("開啟 Log 檔")
        open_log_btn.clicked.connect(self._open_log_file)

        header_layout.addLayout(title_col)
        header_layout.addStretch()
        header_layout.addWidget(theme_label)
        header_layout.addWidget(self.theme_combo)
        header_layout.addWidget(open_out_btn)
        header_layout.addWidget(open_log_btn)
        root_layout.addLayout(header_layout)

        # 分頁容器
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_runner_tab(), "執行與監控")
        self.tabs.addTab(self._build_rules_tab(), "分類規則設定 (rules.md)")
        self.tabs.addTab(self._build_history_tab(), "已歸檔紀錄")
        root_layout.addWidget(self.tabs)

    # ------------------------------------------------------------------
    # Tab 1: 執行與監控
    # ------------------------------------------------------------------
    def _build_runner_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 參數區
        cfg_box = QGroupBox("執行參數")
        cfg_grid = QGridLayout(cfg_box)
        cfg_grid.setSpacing(8)

        cfg_grid.addWidget(QLabel("輸出根目錄："), 0, 0)
        self.output_dir_edit = QLineEdit(str(oa.OUTPUT_ROOT_DIR))
        browse_btn = QPushButton("瀏覽...")
        browse_btn.clicked.connect(self._browse_output_dir)
        cfg_grid.addWidget(self.output_dir_edit, 0, 1, 1, 3)
        cfg_grid.addWidget(browse_btn, 0, 4)

        cfg_grid.addWidget(QLabel("掃描最近天數："), 1, 0)
        self.days_spin = QSpinBox()
        self.days_spin.setRange(0, 3650)
        self.days_spin.setValue(oa.DAYS_BACK if oa.DAYS_BACK is not None else 0)
        self.days_spin.setSpecialValueText("0（不限天數）")
        cfg_grid.addWidget(self.days_spin, 1, 1)

        cfg_grid.addWidget(QLabel("目標信箱（選填）："), 1, 2)
        self.account_edit = QLineEdit(oa.TARGET_ACCOUNT_EMAIL or "")
        self.account_edit.setPlaceholderText("留空使用 Outlook 預設帳號")
        cfg_grid.addWidget(self.account_edit, 1, 3, 1, 2)

        opt_row = QHBoxLayout()
        self.ignore_inline_chk = QCheckBox("排除簽名檔內嵌小圖")
        self.ignore_inline_chk.setChecked(oa.IGNORE_INLINE_IMAGES)
        self.datetime_subdir_chk = QCheckBox("依收件日期時間建子資料夾 (YYYY-MM-DD_AMHHMM)")
        self.datetime_subdir_chk.setChecked(oa.USE_DATETIME_SUBFOLDER)
        opt_row.addWidget(self.ignore_inline_chk)
        opt_row.addWidget(self.datetime_subdir_chk)
        opt_row.addStretch()
        cfg_grid.addLayout(opt_row, 2, 0, 1, 5)
        layout.addWidget(cfg_box)

        # 按鈕列
        btn_row = QHBoxLayout()
        self.run_btn = QPushButton("▶ 開始正式歸檔")
        self.run_btn.setObjectName("PrimaryBtn")
        self.run_btn.clicked.connect(lambda: self._start_run(dry_run=False))

        self.dry_run_btn = QPushButton("🧪 模擬測試 (Dry-Run)")
        self.dry_run_btn.clicked.connect(lambda: self._start_run(dry_run=True))

        self.stop_btn = QPushButton("⏹ 停止")
        self.stop_btn.setObjectName("DangerBtn")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop_run)

        clear_log_btn = QPushButton("清除畫面日誌")
        clear_log_btn.clicked.connect(lambda: self.log_view.clear())

        btn_row.addWidget(self.run_btn)
        btn_row.addWidget(self.dry_run_btn)
        btn_row.addWidget(self.stop_btn)
        btn_row.addStretch()
        btn_row.addWidget(clear_log_btn)
        layout.addLayout(btn_row)

        # 統計卡片
        cards_row = QHBoxLayout()
        self.card_scanned = StatCard("已掃描信件", "#2563EB")
        self.card_skipped_checked = StatCard("略過已檢查", "#64748B")
        self.card_saved = StatCard("儲存 / 預計儲存附件", "#16A34A")
        self.card_skipped_att = StatCard("排除簽名圖 / 非檔案", "#D97706")
        self.card_errors = StatCard("錯誤數", "#DC2626")
        self.stat_cards = (
            self.card_scanned,
            self.card_skipped_checked,
            self.card_saved,
            self.card_skipped_att,
            self.card_errors,
        )
        for card in self.stat_cards:
            cards_row.addWidget(card)
        layout.addLayout(cards_row)

        # 即時執行日誌
        log_box = QGroupBox("即時執行日誌")
        log_layout = QVBoxLayout(log_box)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFont("Consolas", 10))
        self.log_view.setMaximumBlockCount(3000)
        log_layout.addWidget(self.log_view)
        layout.addWidget(log_box, stretch=1)

        return tab

    # ------------------------------------------------------------------
    # Tab 2: 分類規則編輯器
    # ------------------------------------------------------------------
    def _build_rules_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        self.hint_label = QLabel(
            "說明：由上往下依序比對（兩階段：先比對主旨 / 寄件人 / Email / 附件檔名；全部未命中時再比對內文）。"
            "同一欄位多個關鍵字請用逗號「,」或頓號「、」分隔。"
        )
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左側規則清單與排序按鈕
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.rule_list = QListWidget()
        self.rule_list.currentRowChanged.connect(self._on_rule_selected)
        left_layout.addWidget(self.rule_list)

        list_btns = QHBoxLayout()
        add_btn = QPushButton("＋ 新增")
        add_btn.clicked.connect(self._add_rule)
        up_btn = QPushButton("↑ 上移")
        up_btn.clicked.connect(lambda: self._move_rule(-1))
        down_btn = QPushButton("↓ 下移")
        down_btn.clicked.connect(lambda: self._move_rule(1))
        del_btn = QPushButton("🗑 刪除")
        del_btn.clicked.connect(self._delete_rule)
        for btn in (add_btn, up_btn, down_btn, del_btn):
            list_btns.addWidget(btn)
        left_layout.addLayout(list_btns)
        splitter.addWidget(left_widget)

        # 右側欄位編輯表單
        right_box = QGroupBox("規則明細（即時暫存，按右下角「儲存至 rules.md」生效）")
        form = QFormLayout(right_box)
        form.setSpacing(10)

        self.rule_folder_edit = QLineEdit()
        self.rule_subjects_edit = QLineEdit()
        self.rule_senders_edit = QLineEdit()
        self.rule_emails_edit = QLineEdit()
        self.rule_attachments_edit = QLineEdit()
        self.rule_bodies_edit = QLineEdit()

        for widget in (
            self.rule_folder_edit,
            self.rule_subjects_edit,
            self.rule_senders_edit,
            self.rule_emails_edit,
            self.rule_attachments_edit,
            self.rule_bodies_edit,
        ):
            widget.textChanged.connect(self._on_rule_form_edited)

        form.addRow("目標資料夾名稱：", self.rule_folder_edit)
        form.addRow("主旨包含：", self.rule_subjects_edit)
        form.addRow("寄件人包含：", self.rule_senders_edit)
        form.addRow("Email 包含：", self.rule_emails_edit)
        form.addRow("附件包含（整封連動）：", self.rule_attachments_edit)
        form.addRow("內文包含（第二階段備援）：", self.rule_bodies_edit)

        splitter.addWidget(right_box)
        splitter.setSizes([300, 660])
        layout.addWidget(splitter, stretch=1)

        bottom_row = QHBoxLayout()
        reload_btn = QPushButton("↻ 從 rules.md 重新載入")
        reload_btn.clicked.connect(self._load_rules_into_editor)
        save_btn = QPushButton("💾 儲存規則至 rules.md")
        save_btn.setObjectName("PrimaryBtn")
        save_btn.clicked.connect(self._save_rules_from_editor)
        bottom_row.addStretch()
        bottom_row.addWidget(reload_btn)
        bottom_row.addWidget(save_btn)
        layout.addLayout(bottom_row)

        return tab

    # ------------------------------------------------------------------
    # Tab 3: 已歸檔紀錄
    # ------------------------------------------------------------------
    def _build_history_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        top_row = QHBoxLayout()
        self.history_search_edit = QLineEdit()
        self.history_search_edit.setPlaceholderText("搜尋分類資料夾或檔案路徑...")
        self.history_search_edit.textChanged.connect(self._filter_history_rows)

        refresh_btn = QPushButton("↻ 重新整理")
        refresh_btn.clicked.connect(self._load_history_table)

        reset_db_btn = QPushButton("🗑 清空防重複紀錄 (state.db)")
        reset_db_btn.setObjectName("DangerBtn")
        reset_db_btn.clicked.connect(self._clear_state_db)

        top_row.addWidget(self.history_search_edit, stretch=1)
        top_row.addWidget(refresh_btn)
        top_row.addWidget(reset_db_btn)
        layout.addLayout(top_row)

        self.history_table = QTableWidget(0, 3)
        self.history_table.setHorizontalHeaderLabels(["歸檔時間", "分類資料夾", "儲存檔案路徑（雙擊可開啟檔案）"])
        self.history_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.history_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.history_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.history_table.cellDoubleClicked.connect(self._open_history_item)
        layout.addWidget(self.history_table, stretch=1)

        return tab

    # ------------------------------------------------------------------
    # 執行與日誌處理
    # ------------------------------------------------------------------
    def _browse_output_dir(self) -> None:
        chosen = QFileDialog.getExistingDirectory(self, "選擇附件輸出根目錄", self.output_dir_edit.text())
        if chosen:
            self.output_dir_edit.setText(chosen)

    def _open_output_dir(self) -> None:
        target = Path(self.output_dir_edit.text().strip() or oa.OUTPUT_ROOT_DIR)
        target.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))

    def _open_log_file(self) -> None:
        if oa.LOG_FILE_PATH.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(oa.LOG_FILE_PATH)))
        else:
            QMessageBox.information(self, "提示", "尚未產生日誌檔案。")

    def _start_run(self, *, dry_run: bool) -> None:
        if self.worker and self.worker.isRunning():
            return

        output_root = Path(self.output_dir_edit.text().strip() or str(oa.OUTPUT_ROOT_DIR))
        days_val = self.days_spin.value()
        days_back = None if days_val == 0 else days_val
        account = self.account_edit.text().strip() or None

        for card in (
            self.card_scanned,
            self.card_skipped_checked,
            self.card_saved,
            self.card_skipped_att,
            self.card_errors,
        ):
            card.set_value(0)

        self.run_btn.setEnabled(False)
        self.dry_run_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("正在執行模擬測試..." if dry_run else "正在掃描並歸檔 Outlook 附件...")

        self.worker = ArchiveWorker(
            dry_run=dry_run,
            days_back=days_back,
            account_email=account,
            output_root=output_root,
            ignore_inline=self.ignore_inline_chk.isChecked(),
            use_datetime_subfolder=self.datetime_subdir_chk.isChecked(),
        )
        self.worker.log_received.connect(self._append_log_line)
        self.worker.progress_updated.connect(self._on_progress_updated)
        self.worker.finished_run.connect(self._on_run_finished)
        self.worker.start()

    def _stop_run(self) -> None:
        if self.worker and self.worker.isRunning():
            self.worker.request_cancel()
            self.stop_btn.setEnabled(False)
            self.status_label.setText("正在安全停止掃描...")

    def _append_log_line(self, levelno: int, message: str) -> None:
        fmt = QTextCharFormat()
        if levelno >= logging.ERROR:
            fmt.setForeground(QColor(self._active_theme["log_error"]))
        elif levelno >= logging.WARNING:
            fmt.setForeground(QColor(self._active_theme["log_warning"]))
        elif "[DRY-RUN]" in message or "已儲存附件" in message:
            fmt.setForeground(QColor(self._active_theme["log_success"]))
        else:
            fmt.setForeground(QColor(self._active_theme["log_info"]))

        cursor = self.log_view.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(message + "\n", fmt)
        self.log_view.setTextCursor(cursor)
        self.log_view.ensureCursorVisible()

    def _on_progress_updated(self, stats: dict, current_subject: str) -> None:
        self.card_scanned.set_value(stats.get("scanned_mails", 0))
        self.card_skipped_checked.set_value(stats.get("skipped_checked", 0))
        self.card_saved.set_value(stats.get("saved_attachments", 0))
        self.card_skipped_att.set_value(stats.get("skipped_attachments", 0))
        self.card_errors.set_value(stats.get("errors", 0))
        short_subj = (current_subject[:45] + "...") if len(current_subject) > 45 else current_subject
        self.status_label.setText(f"處理中：{short_subj}")

    def _on_run_finished(self, stats: dict, error_msg: str) -> None:
        self.card_scanned.set_value(stats.get("scanned_mails", 0))
        self.card_skipped_checked.set_value(stats.get("skipped_checked", 0))
        self.card_saved.set_value(stats.get("saved_attachments", 0))
        self.card_skipped_att.set_value(stats.get("skipped_attachments", 0))
        self.card_errors.set_value(stats.get("errors", 0))

        self.run_btn.setEnabled(True)
        self.dry_run_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self._load_history_table()

        if error_msg:
            self.status_label.setText(f"執行失敗：{error_msg}")
            QMessageBox.warning(self, "執行錯誤", error_msg)
        else:
            self.status_label.setText(
                f"完成：掃描 {stats.get('scanned_mails', 0)} 封，儲存 {stats.get('saved_attachments', 0)} 個附件"
            )

    # ------------------------------------------------------------------
    # 規則編輯器邏輯
    # ------------------------------------------------------------------
    def _load_rules_into_editor(self) -> None:
        try:
            self.rules = oa.parse_rules_markdown(oa.RULES_FILE)
        except Exception as exc:
            QMessageBox.warning(self, "讀取失敗", f"無法讀取 rules.md：{exc}")
            self.rules = []
        self._refresh_rule_list(selected_row=0 if self.rules else -1)

    def _refresh_rule_list(self, selected_row: int) -> None:
        self.rule_list.blockSignals(True)
        self.rule_list.clear()
        for rule in self.rules:
            self.rule_list.addItem(QListWidgetItem(rule.folder_name))
        self.rule_list.blockSignals(False)

        if 0 <= selected_row < len(self.rules):
            self.rule_list.setCurrentRow(selected_row)
            self._on_rule_selected(selected_row)
        else:
            self._clear_rule_form()

    def _clear_rule_form(self) -> None:
        self._syncing_rule_form = True
        for edit in (
            self.rule_folder_edit,
            self.rule_subjects_edit,
            self.rule_senders_edit,
            self.rule_emails_edit,
            self.rule_attachments_edit,
            self.rule_bodies_edit,
        ):
            edit.clear()
            edit.setEnabled(False)
        self._syncing_rule_form = False

    def _on_rule_selected(self, row: int) -> None:
        if row < 0 or row >= len(self.rules):
            self._clear_rule_form()
            return
        rule = self.rules[row]
        self._syncing_rule_form = True
        for edit in (
            self.rule_folder_edit,
            self.rule_subjects_edit,
            self.rule_senders_edit,
            self.rule_emails_edit,
            self.rule_attachments_edit,
            self.rule_bodies_edit,
        ):
            edit.setEnabled(True)
        self.rule_folder_edit.setText(rule.folder_name)
        self.rule_subjects_edit.setText(", ".join(rule.subjects))
        self.rule_senders_edit.setText(", ".join(rule.senders))
        self.rule_emails_edit.setText(", ".join(rule.emails))
        self.rule_attachments_edit.setText(", ".join(rule.attachments))
        self.rule_bodies_edit.setText(", ".join(rule.bodies))
        self._syncing_rule_form = False

    @staticmethod
    def _split_input(text: str) -> tuple[str, ...]:
        return tuple(p.strip() for p in re.split(r"[,、]", text) if p.strip())

    def _on_rule_form_edited(self) -> None:
        if self._syncing_rule_form:
            return
        row = self.rule_list.currentRow()
        if row < 0 or row >= len(self.rules):
            return

        folder_name = self.rule_folder_edit.text().strip() or "未命名資料夾"
        self.rules[row] = oa.ClassificationRule(
            folder_name=oa.sanitize_filename(folder_name),
            subjects=self._split_input(self.rule_subjects_edit.text()),
            senders=self._split_input(self.rule_senders_edit.text()),
            emails=self._split_input(self.rule_emails_edit.text()),
            attachments=self._split_input(self.rule_attachments_edit.text()),
            bodies=self._split_input(self.rule_bodies_edit.text()),
        )
        item = self.rule_list.item(row)
        if item is not None:
            item.setText(self.rules[row].folder_name)

    def _add_rule(self) -> None:
        new_rule = oa.ClassificationRule("新廠商", (), (), (), (), ())
        self.rules.append(new_rule)
        self._refresh_rule_list(len(self.rules) - 1)
        self.rule_folder_edit.setFocus()
        self.rule_folder_edit.selectAll()

    def _move_rule(self, delta: int) -> None:
        row = self.rule_list.currentRow()
        target = row + delta
        if row < 0 or target < 0 or target >= len(self.rules):
            return
        self.rules[row], self.rules[target] = self.rules[target], self.rules[row]
        self._refresh_rule_list(target)

    def _delete_rule(self) -> None:
        row = self.rule_list.currentRow()
        if row < 0 or row >= len(self.rules):
            return
        name = self.rules[row].folder_name
        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除規則「{name}」嗎？（需點擊儲存才會寫入 rules.md）",
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        del self.rules[row]
        next_row = min(row, len(self.rules) - 1)
        self._refresh_rule_list(next_row)

    def _save_rules_from_editor(self) -> None:
        try:
            oa.save_rules_markdown(oa.RULES_FILE, self.rules)
            self._load_rules_into_editor()
            QMessageBox.information(self, "儲存成功", f"已將 {len(self.rules)} 條規則寫入 rules.md。")
        except Exception as exc:
            QMessageBox.critical(self, "儲存失敗", f"寫入 rules.md 時發生錯誤：{exc}")

    # ------------------------------------------------------------------
    # 歸檔紀錄檢視器
    # ------------------------------------------------------------------
    def _load_history_table(self) -> None:
        rows: list[tuple[str, str, str]] = []
        if oa.STATE_DB_PATH.exists():
            store = oa.StateStore(oa.STATE_DB_PATH)
            try:
                rows = store.list_downloaded_attachments(limit=500)
            finally:
                store.close()

        self.history_table.setRowCount(len(rows))
        for r_idx, (ts, folder, path_str) in enumerate(rows):
            self.history_table.setItem(r_idx, 0, QTableWidgetItem(ts))
            self.history_table.setItem(r_idx, 1, QTableWidgetItem(folder))
            self.history_table.setItem(r_idx, 2, QTableWidgetItem(path_str))
        self._filter_history_rows(self.history_search_edit.text())

    def _filter_history_rows(self, keyword: str) -> None:
        kw = keyword.strip().lower()
        for r_idx in range(self.history_table.rowCount()):
            if not kw:
                self.history_table.setRowHidden(r_idx, False)
                continue
            folder_text = (self.history_table.item(r_idx, 1).text() or "").lower()
            path_text = (self.history_table.item(r_idx, 2).text() or "").lower()
            self.history_table.setRowHidden(r_idx, kw not in folder_text and kw not in path_text)

    def _open_history_item(self, row: int, _col: int) -> None:
        item = self.history_table.item(row, 2)
        if not item:
            return
        file_path = Path(item.text())
        if file_path.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(file_path)))
        elif file_path.parent.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(file_path.parent)))
        else:
            QMessageBox.warning(self, "檔案不存在", f"找不到檔案或目錄：\n{file_path}")

    def _clear_state_db(self) -> None:
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "請稍候", "執行中無法清空防重複紀錄。")
            return
        reply = QMessageBox.warning(
            self,
            "確認清空防重複紀錄",
            "清空 state.db 後，下次執行將會重新檢查並下載過去掃描範圍內的信件附件（已存在的檔案會加上 _1、_2 避免覆蓋）。\n\n確定要清空嗎？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        store = oa.StateStore(oa.STATE_DB_PATH)
        try:
            store.clear_all()
        finally:
            store.close()
        self._load_history_table()
        QMessageBox.information(self, "完成", "已清空防重複紀錄。")

    def closeEvent(self, event) -> None:
        if self.worker and self.worker.isRunning():
            self.worker.request_cancel()
            self.worker.wait(5000)
        super().closeEvent(event)


def main() -> int:
    if "--cli-dry-run" in sys.argv:
        return oa.main(["--dry-run", "--limit", "5"])

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
