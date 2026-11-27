"""Outlook 附件自動分類歸檔工具。

透過 Windows COM 連接桌面版 Classic Outlook，依據 rules.md 的規則將信件附件
自動分類儲存至目標資料夾，並使用 SQLite 記錄已下載附件以防止重複下載。
"""

from __future__ import annotations

import argparse
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import datetime, timedelta
import html
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import re
import shutil
import sqlite3
import subprocess
import sys
import time
from typing import Any
from urllib.parse import quote

# ==============================================================================
# 可調參數設定區（依實際需求調整）
# ==============================================================================
APP_VERSION = "V1.1"
BASE_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
RULES_FILE = BASE_DIR / "rules.md"
OUTPUT_ROOT_DIR = Path(r"d:\圖檔\CAD")
STATE_DB_PATH = BASE_DIR / "state.db"
LOG_FILE_PATH = BASE_DIR / "logs" / "archiver.log"

# 目標帳號與資料夾設定：
# - TARGET_ACCOUNT_EMAIL: 設為 None 代表使用 Outlook 預設帳號；
#   日後若要指定特定帳號（例如 Yahoo 信箱），請填入該帳號顯示名稱或 Email。
# - TARGET_FOLDER_NAME: 設為 None 代表讀取該帳號預設「收件匣」(olFolderInbox = 6)；
#   若要讀取自訂資料夾（例如 "收件匣/客戶圖檔"），請填入資料夾名稱或以 "/" 分隔的路徑。
TARGET_ACCOUNT_EMAIL: str | None = None
TARGET_FOLDER_NAME: str | None = None

# 未命中任何規則時的預設分類資料夾名稱
DEFAULT_FOLDER = "其他未分類"

# 掃描最近 N 天的信件（設為 None 代表不限天數，掃描全部信件）
DAYS_BACK: int | None = 7

# 若 Outlook 尚未啟動，自動啟動後的最長等待秒數與輪詢間隔
OUTLOOK_STARTUP_TIMEOUT = 45
OUTLOOK_POLL_INTERVAL = 2
OUTLOOK_MIN_INIT_WAIT = 10

# 內嵌圖片與附件過濾設定
IGNORE_INLINE_IMAGES = True
INLINE_IMAGE_MAX_BYTES = 15 * 1024  # image00* 開頭且小於 15KB 視為簽名檔小圖
INLINE_CID_MAX_BYTES = 30 * 1024    # 具 Content-ID 或隱藏屬性且小於 30KB 視為內嵌圖

# 副檔名過濾（設為 None 或空 tuple 表示不限制；小寫含點，例如 (".dwg", ".dxf", ".pdf")）
ALLOWED_EXTENSIONS: tuple[str, ...] | None = None
IGNORED_EXTENSIONS: tuple[str, ...] = ()

# 同名檔案處理模式："suffix_counter"（保留原名加 _1、_2）或 "prefix_date"
DUPLICATE_FILENAME_MODE = "suffix_counter"

# 是否依收件日期時間在分類資料夾底下建立子資料夾（例如 合宏展\2026-11-26_AM0938\）
USE_DATETIME_SUBFOLDER = True
DATETIME_SUBFOLDER_FORMAT = "%Y-%m-%d_%p%I%M"

# 是否在子資料夾內產生輕量信件紀錄檔 (.html)，記錄主旨/寄件人/時間/內文
SAVE_EMAIL_HTML = True
EMAIL_HTML_FILENAME = "_信件內容.html"

# ==============================================================================
# 常數與 MAPI Property Tags
# ==============================================================================
OL_FOLDER_INBOX = 6
OL_MAIL = 43
OL_BY_VALUE = 1

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

_INVALID_WIN_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_WIN_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_RULE_HEADER_RE = re.compile(r"^\s*##\s+(.+?)\s*$")
_RULE_FIELD_RE = re.compile(r"^\s*-\s*\*\*(.+?)\*\*\s*[:：]\s*(.*)$")


@dataclass(frozen=True)
class ClassificationRule:
    """單一分類資料夾的比對規則。"""

    folder_name: str
    subjects: tuple[str, ...]
    senders: tuple[str, ...]
    emails: tuple[str, ...]
    attachments: tuple[str, ...] = ()
    bodies: tuple[str, ...] = ()


_RULE_FIELD_KEYS = ("主旨包含", "寄件人包含", "Email包含", "附件包含", "內文包含")
_RULE_FIELD_ALIASES = {
    "主旨包含": "主旨包含",
    "寄件人包含": "寄件人包含",
    "email包含": "Email包含",
    "附件包含": "附件包含",
    "附件檔名包含": "附件包含",
    "內文包含": "內文包含",
    "內文關鍵字": "內文包含",
    "正文包含": "內文包含",
}


def setup_logging(
    log_path: Path = LOG_FILE_PATH,
    verbose: bool = False,
    extra_handlers: Sequence[logging.Handler] = (),
) -> logging.Logger:
    """設定檔案輪替與主控台雙向日誌（UTF-8）。"""
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("outlook_archiver")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=2 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)

    if sys.stdout is not None:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
        logger.addHandler(console_handler)

    for handler in extra_handlers:
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger


def sanitize_filename(name: str, fallback: str = "unnamed") -> str:
    """清除 Windows 不允許的檔名或資料夾字元並處理保留字。"""
    cleaned = _INVALID_WIN_CHARS_RE.sub("_", name).strip().rstrip(".")
    if not cleaned:
        return fallback
    stem = Path(cleaned).stem.upper()
    if stem in _RESERVED_WIN_NAMES:
        cleaned = f"_{cleaned}"
    return cleaned


def _split_keywords(raw_value: str) -> tuple[str, ...]:
    """依半形逗號或全形頓號切割關鍵字，移除空白與空字串。"""
    tokens = [part.strip() for part in re.split(r"[,、]", raw_value)]
    return tuple(token for token in tokens if token)


def parse_rules_markdown(rules_path: Path) -> list[ClassificationRule]:
    """解析 rules.md，回傳依序排列的分類規則列表。"""
    if not rules_path.exists():
        raise FileNotFoundError(f"找不到分類規則檔：{rules_path}")

    content = rules_path.read_text(encoding="utf-8")
    rules: list[ClassificationRule] = []

    current_folder: str | None = None
    fields: dict[str, list[str]] = {k: [] for k in _RULE_FIELD_KEYS}

    def flush_current() -> None:
        if current_folder is None:
            return
        safe_folder = sanitize_filename(current_folder)
        rules.append(
            ClassificationRule(
                folder_name=safe_folder,
                subjects=tuple(fields["主旨包含"]),
                senders=tuple(fields["寄件人包含"]),
                emails=tuple(fields["Email包含"]),
                attachments=tuple(fields["附件包含"]),
                bodies=tuple(fields["內文包含"]),
            )
        )

    for line in content.splitlines():
        header_match = _RULE_HEADER_RE.match(line)
        if header_match:
            flush_current()
            current_folder = header_match.group(1).strip()
            fields = {k: [] for k in _RULE_FIELD_KEYS}
            continue

        if current_folder is None:
            continue

        field_match = _RULE_FIELD_RE.match(line)
        if not field_match:
            continue

        key = field_match.group(1).strip().lower()
        value = field_match.group(2).strip()
        canonical_key = _RULE_FIELD_ALIASES.get(key)
        if canonical_key:
            fields[canonical_key].extend(_split_keywords(value))

    flush_current()
    return rules


def save_rules_markdown(rules_path: Path, rules: Sequence[ClassificationRule]) -> None:
    """將分類規則清單序列化並寫回 rules.md。"""
    lines = [
        "# Outlook 附件自動分類規則",
        "",
        "> **使用說明**：",
        "> - 程式會由上往下依序比對，採**兩階段比對**：",
        ">   1. 第一階段先比對 **主旨包含**、**寄件人包含**、**Email包含**、**附件包含**（命中第一條即歸類停止）。",
        ">   2. 若全部規則的前四項皆未命中，第二階段再比對 **內文包含**。",
        "> - 同一區塊內的條件為 **OR** 邏輯，任一符合即成立（整封信內只要有任一有效附件檔名命中，整封信所有附件都會歸入該資料夾），英文不分大小寫。",
        "> - 多個關鍵字請以半形逗號 `,` 或全形頓號 `、` 分隔，留空代表該欄位不比對。",
        "",
        "---",
    ]
    for rule in rules:
        safe_name = sanitize_filename(rule.folder_name)
        lines.extend(
            [
                "",
                f"## {safe_name}",
                f"- **主旨包含**: {', '.join(rule.subjects)}",
                f"- **寄件人包含**: {', '.join(rule.senders)}",
                f"- **Email包含**: {', '.join(rule.emails)}",
                f"- **附件包含**: {', '.join(rule.attachments)}",
                f"- **內文包含**: {', '.join(rule.bodies)}",
            ]
        )
    lines.append("")
    rules_path.parent.mkdir(parents=True, exist_ok=True)
    rules_path.write_text("\n".join(lines), encoding="utf-8")


def match_rule(
    rules: list[ClassificationRule],
    subject: str,
    sender_name: str,
    sender_email: str,
    attachment_names: Sequence[str] = (),
    body: str = "",
    body_getter: Callable[[], str] | None = None,
    default_folder: str = DEFAULT_FOLDER,
) -> tuple[str, str]:
    """兩階段由上往下比對分類規則（OR 邏輯、忽略英文大小寫）。

    第一階段先比對主旨、寄件人、Email、附件檔名；
    若全部規則均未命中，第二階段再讀取內文進行備援比對。
    """
    subject_l = (subject or "").lower()
    sender_l = (sender_name or "").lower()
    email_l = (sender_email or "").lower()
    attachments_l = [(name, (name or "").lower()) for name in attachment_names]

    # 第一階段：精準欄位比對
    for rule in rules:
        checks = (
            ("主旨包含", rule.subjects, subject_l),
            ("寄件人包含", rule.senders, sender_l),
            ("Email包含", rule.emails, email_l),
        )
        for label, keywords, haystack in checks:
            for kw in keywords:
                if kw.lower() in haystack:
                    return rule.folder_name, f"{label} '{kw}'"

        for kw in rule.attachments:
            kw_l = kw.lower()
            for orig_name, fname_l in attachments_l:
                if kw_l in fname_l:
                    return rule.folder_name, f"附件包含 '{kw}' ({orig_name})"

    # 第二階段：前四項未命中時，才延遲讀取內文並依序比對
    if any(rule.bodies for rule in rules):
        body_text = body_getter() if body_getter is not None else body
        body_l = (body_text or "").lower()
        if body_l:
            for rule in rules:
                for kw in rule.bodies:
                    if kw.lower() in body_l:
                        return rule.folder_name, f"內文包含 '{kw}'"

    return sanitize_filename(default_folder, fallback="其他未分類"), "未命中任何規則"


def resolve_unique_path(
    directory: Path,
    filename: str,
    received_dt: datetime | None = None,
    mode: str = DUPLICATE_FILENAME_MODE,
) -> Path:
    """產生不覆蓋既有檔案的安全儲存路徑。"""
    clean_name = sanitize_filename(filename, fallback="attachment")
    stem = Path(clean_name).stem
    suffix = Path(clean_name).suffix

    if mode == "prefix_date" and received_dt is not None:
        base_stem = f"{received_dt.strftime('%Y%m%d')}_{stem}"
    else:
        base_stem = stem

    candidate = directory / f"{base_stem}{suffix}"
    counter = 1
    while candidate.exists():
        candidate = directory / f"{base_stem}_{counter}{suffix}"
        counter += 1
    return candidate


def should_skip_attachment(
    filename: str,
    size_bytes: int,
    att_type: int,
    has_content_id: bool,
    is_hidden: bool,
    ignore_inline: bool = IGNORE_INLINE_IMAGES,
    allowed_extensions: tuple[str, ...] | None = ALLOWED_EXTENSIONS,
    ignored_extensions: tuple[str, ...] = IGNORED_EXTENSIONS,
) -> tuple[bool, str]:
    """判斷附件是否應該排除，回傳 (是否跳過, 原因)。"""
    if att_type != OL_BY_VALUE:
        return True, f"非檔案型附件 (Type={att_type})"

    ext = Path(filename).suffix.lower()
    if allowed_extensions:
        normalized_allowed = {e.lower() for e in allowed_extensions}
        if ext not in normalized_allowed:
            return True, f"副檔名不在白名單 ({ext or '無'})"

    if ignored_extensions:
        normalized_ignored = {e.lower() for e in ignored_extensions}
        if ext in normalized_ignored:
            return True, f"副檔名在排除清單 ({ext})"

    if ignore_inline:
        lower_name = filename.lower()
        if lower_name.startswith("image00") and size_bytes < INLINE_IMAGE_MAX_BYTES:
            return True, f"疑似簽名檔小圖 ({filename}, {size_bytes} bytes)"
        if (has_content_id or is_hidden) and size_bytes < INLINE_CID_MAX_BYTES:
            return True, f"內嵌或隱藏圖片 ({filename}, {size_bytes} bytes)"

    return False, ""


class StateStore:
    """使用 SQLite 儲存已檢查郵件與已下載附件紀錄，避免重複下載。"""

    def __init__(self, db_path: Path) -> None:
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._conn = sqlite3.connect(self.db_path)
        self._init_schema()

    def _init_schema(self) -> None:
        with self._conn:
            self._conn.execute(
                """
                CREATE TABLE IF NOT EXISTS processed_attachments (
                    message_key TEXT NOT NULL,
                    attachment_key TEXT NOT NULL,
                    saved_path TEXT NOT NULL,
                    target_folder TEXT NOT NULL,
                    processed_at TEXT NOT NULL,
                    PRIMARY KEY (message_key, attachment_key)
                )
                """
            )
            self._conn.execute(
                """
                CREATE TABLE IF NOT EXISTS checked_messages (
                    message_key TEXT PRIMARY KEY,
                    subject TEXT NOT NULL,
                    status TEXT NOT NULL,
                    checked_at TEXT NOT NULL
                )
                """
            )

    def is_message_checked(self, message_key: str) -> bool:
        row = self._conn.execute(
            "SELECT 1 FROM checked_messages WHERE message_key = ?",
            (message_key,),
        ).fetchone()
        return row is not None

    def mark_message_checked(self, message_key: str, subject: str, status: str) -> None:
        now_iso = datetime.now().isoformat(timespec="seconds")
        with self._conn:
            self._conn.execute(
                """
                INSERT OR REPLACE INTO checked_messages (message_key, subject, status, checked_at)
                VALUES (?, ?, ?, ?)
                """,
                (message_key, subject, status, now_iso),
            )

    def is_attachment_downloaded(self, message_key: str, attachment_key: str) -> bool:
        row = self._conn.execute(
            """
            SELECT 1 FROM processed_attachments
            WHERE message_key = ? AND attachment_key = ?
            """,
            (message_key, attachment_key),
        ).fetchone()
        return row is not None

    def mark_attachment_downloaded(
        self,
        message_key: str,
        attachment_key: str,
        saved_path: Path,
        target_folder: str,
    ) -> None:
        now_iso = datetime.now().isoformat(timespec="seconds")
        with self._conn:
            self._conn.execute(
                """
                INSERT OR REPLACE INTO processed_attachments
                    (message_key, attachment_key, saved_path, target_folder, processed_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (message_key, attachment_key, str(saved_path), target_folder, now_iso),
            )

    def list_downloaded_attachments(self, limit: int = 500) -> list[tuple[str, str, str]]:
        """依時間由新到舊回傳已下載附件清單 (processed_at, target_folder, saved_path)。"""
        rows = self._conn.execute(
            """
            SELECT processed_at, target_folder, saved_path
            FROM processed_attachments
            ORDER BY processed_at DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [(str(r[0]), str(r[1]), str(r[2])) for r in rows]

    def clear_all(self) -> None:
        """清空已檢查郵件與已下載附件紀錄。"""
        with self._conn:
            self._conn.execute("DELETE FROM processed_attachments")
            self._conn.execute("DELETE FROM checked_messages")

    def close(self) -> None:
        self._conn.close()


@dataclass(frozen=True)
class ArchiveReportRecord:
    """供報表檢視與篩選用的單筆已歸檔紀錄。"""

    vendor: str
    batch_label: str
    received_dt: datetime | None
    archived_at: str
    saved_path: str
    filename: str


_SUBFOLDER_DT_FORMATS = ("%Y-%m-%d_%p%I%M", "%Y-%m-%d_%H%M%S", "%Y-%m-%d")


def parse_received_dt_from_path(saved_path: str, archived_at: str) -> tuple[str, datetime | None]:
    """從儲存路徑的子資料夾名稱解析收件日期時間，失敗時退回歸檔時間。"""
    parent_name = Path(saved_path).parent.name
    for fmt in _SUBFOLDER_DT_FORMATS:
        try:
            parsed = datetime.strptime(parent_name, fmt)
            return parent_name, parsed
        except ValueError:
            continue

    try:
        fallback_dt = datetime.fromisoformat(archived_at)
        return fallback_dt.strftime("%Y-%m-%d"), fallback_dt.replace(tzinfo=None)
    except Exception:
        return parent_name or "未知日期", None


def build_archive_report_records(rows: Sequence[tuple[str, str, str]]) -> list[ArchiveReportRecord]:
    """將 StateStore 查詢列轉換為報表紀錄物件列表。"""
    records: list[ArchiveReportRecord] = []
    for archived_at, vendor, saved_path in rows:
        batch_label, received_dt = parse_received_dt_from_path(saved_path, archived_at)
        records.append(
            ArchiveReportRecord(
                vendor=vendor or DEFAULT_FOLDER,
                batch_label=batch_label,
                received_dt=received_dt,
                archived_at=archived_at,
                saved_path=saved_path,
                filename=Path(saved_path).name,
            )
        )
    return records


@dataclass(frozen=True)
class SmartDateFilterResult:
    """智慧數字日期／時間輸入解析結果。"""

    is_valid: bool
    windows: tuple[tuple[datetime, datetime], ...]
    status_text: str


def _build_am_pm_windows(
    year: int,
    month: int,
    day: int,
    hour: int,
    minute: int,
    tolerance_minutes: int,
) -> tuple[tuple[tuple[datetime, datetime], ...], str]:
    """依小時是否屬於 1~12 智慧展開 AM/PM 前後 N 分鐘搜尋區間。"""
    base_dt = datetime(year, month, day, hour, minute)
    delta = timedelta(minutes=tolerance_minutes)
    candidates = [base_dt]
    if 1 <= hour <= 11:
        candidates.append(base_dt.replace(hour=hour + 12))
    elif hour == 12:
        candidates.insert(0, base_dt.replace(hour=0))

    windows: list[tuple[datetime, datetime]] = []
    desc_parts: list[str] = []
    for dt in candidates:
        start = (dt - delta).replace(second=0, microsecond=0)
        end = (dt + delta).replace(second=59, microsecond=999999)
        windows.append((start, end))
        desc_parts.append(f"{start.strftime('%Y-%m-%d %H:%M')}~{end.strftime('%H:%M')}")
    return tuple(windows), " 或 ".join(desc_parts)


def parse_smart_datetime_windows(
    raw_text: str,
    tolerance_minutes: int = 5,
    reference_dt: datetime | None = None,
) -> SmartDateFilterResult:
    """解析數字型日期時間輸入並回傳前後 N 分鐘的篩選區間。

    支援格式：
    - 12 碼 `YYYYMMDDHHMM`（如 202610220240 / 202610221440）
    - 10 碼 `YYMMDDHHMM`（如 2610221440）
    - 8 碼 `YYYYMMDD`（若以 19/20 開頭，搜尋該日全天）或 `MMDDHHMM`（省略年份，如 10220240 / 10221440）
    - 6 碼 `YYMMDD`（搜尋該日全天）
    - 4 碼 `MMDD`（省略年份，搜尋該日全天）
    """
    stripped = raw_text.strip()
    if not stripped:
        return SmartDateFilterResult(is_valid=True, windows=(), status_text="")

    digits = re.sub(r"[-/:_\s]+", "", stripped)
    if not digits.isdigit():
        return SmartDateFilterResult(
            is_valid=False,
            windows=(),
            status_text="請輸入數字日期時間（例：202610220240、10221440、20261022）",
        )

    ref_year = (reference_dt or datetime.now()).year
    length = len(digits)

    try:
        if length == 12:
            y, m, d, hh, mm = (
                int(digits[0:4]),
                int(digits[4:6]),
                int(digits[6:8]),
                int(digits[8:10]),
                int(digits[10:12]),
            )
            windows, desc = _build_am_pm_windows(y, m, d, hh, mm, tolerance_minutes)
            return SmartDateFilterResult(is_valid=True, windows=windows, status_text=f"條件：{desc}")

        if length == 10:
            y, m, d, hh, mm = (
                2000 + int(digits[0:2]),
                int(digits[2:4]),
                int(digits[4:6]),
                int(digits[6:8]),
                int(digits[8:10]),
            )
            windows, desc = _build_am_pm_windows(y, m, d, hh, mm, tolerance_minutes)
            return SmartDateFilterResult(is_valid=True, windows=windows, status_text=f"條件：{desc}")

        if length == 8:
            if digits.startswith(("19", "20")):
                y, m, d = int(digits[0:4]), int(digits[4:6]), int(digits[6:8])
                day_start = datetime(y, m, d, 0, 0, 0)
                day_end = datetime(y, m, d, 23, 59, 59, 999999)
                return SmartDateFilterResult(
                    is_valid=True,
                    windows=((day_start, day_end),),
                    status_text=f"條件：{day_start.strftime('%Y-%m-%d')} 全日",
                )
            m, d, hh, mm = (
                int(digits[0:2]),
                int(digits[2:4]),
                int(digits[4:6]),
                int(digits[6:8]),
            )
            windows, desc = _build_am_pm_windows(ref_year, m, d, hh, mm, tolerance_minutes)
            return SmartDateFilterResult(is_valid=True, windows=windows, status_text=f"條件：{desc}")

        if length == 6:
            y, m, d = 2000 + int(digits[0:2]), int(digits[2:4]), int(digits[4:6])
            day_start = datetime(y, m, d, 0, 0, 0)
            day_end = datetime(y, m, d, 23, 59, 59, 999999)
            return SmartDateFilterResult(
                is_valid=True,
                windows=((day_start, day_end),),
                status_text=f"條件：{day_start.strftime('%Y-%m-%d')} 全日",
            )

        if length == 4:
            m, d = int(digits[0:2]), int(digits[2:4])
            day_start = datetime(ref_year, m, d, 0, 0, 0)
            day_end = datetime(ref_year, m, d, 23, 59, 59, 999999)
            return SmartDateFilterResult(
                is_valid=True,
                windows=((day_start, day_end),),
                status_text=f"條件：{day_start.strftime('%Y-%m-%d')} 全日",
            )
    except ValueError:
        return SmartDateFilterResult(
            is_valid=False,
            windows=(),
            status_text="日期或時間數值無效，請檢查月份、日期與小時分鐘",
        )

    return SmartDateFilterResult(
        is_valid=False,
        windows=(),
        status_text="請輸入 4/8/12 碼數字（例：1022、10220240、10221440、202610220240）",
    )


@dataclass(frozen=True)
class UnifiedReportQuery:
    """整合式搜尋框的智慧解析結果（同時支援關鍵字與數字日期時間）。"""

    keywords: tuple[str, ...]
    time_windows: tuple[tuple[datetime, datetime], ...]
    status_text: str
    has_datetime_filter: bool


_DATE_LIKE_TOKEN_RE = re.compile(r"^[\d\-/:_]+$")
_DATE_LIKE_FULL_RE = re.compile(r"^[\d\-/:_\s]+$")


def parse_unified_report_query(
    raw_text: str,
    tolerance_minutes: int = 5,
    reference_dt: datetime | None = None,
) -> UnifiedReportQuery:
    """同時解析混合關鍵字與數字日期／時間的搜尋輸入。

    例如：
    - '202610220240' / '10221440' / '2026-10-22 14:40' -> 純日期時間 ±5 分鐘篩選
    - '單井 11200345' -> 同時篩選廠商/檔名含 '單井' 且時間在 11/20 03:40±5分 或 15:40±5分
    - '4070' -> 無法構成有效日期時間，自動當作關鍵字搜尋（如 A050-4070.stp）
    """
    stripped = raw_text.strip()
    if not stripped:
        return UnifiedReportQuery(keywords=(), time_windows=(), status_text="", has_datetime_filter=False)

    # 若整體僅由數字與日期分隔符組成，優先嘗試整體作為日期時間解析（支援 '2026-10-22 14:40'）
    if _DATE_LIKE_FULL_RE.fullmatch(stripped):
        whole_res = parse_smart_datetime_windows(
            stripped,
            tolerance_minutes=tolerance_minutes,
            reference_dt=reference_dt,
        )
        if whole_res.is_valid and whole_res.windows:
            desc = whole_res.status_text.removeprefix("條件：")
            return UnifiedReportQuery(
                keywords=(),
                time_windows=whole_res.windows,
                status_text=f"智慧辨識日期時間：{desc}",
                has_datetime_filter=True,
            )

    keywords: list[str] = []
    windows: list[tuple[datetime, datetime]] = []
    date_descs: list[str] = []

    for token in stripped.split():
        if _DATE_LIKE_TOKEN_RE.fullmatch(token):
            digits = re.sub(r"[-/:_]+", "", token)
            if len(digits) in (4, 6, 8, 10, 12):
                res = parse_smart_datetime_windows(
                    token,
                    tolerance_minutes=tolerance_minutes,
                    reference_dt=reference_dt,
                )
                if res.is_valid and res.windows:
                    windows.extend(res.windows)
                    date_descs.append(res.status_text.removeprefix("條件："))
                    continue
        keywords.append(token)

    status_parts: list[str] = []
    if date_descs:
        status_parts.append(f"日期時間：{'；'.join(date_descs)}")
    if keywords:
        status_parts.append(f"關鍵字：{', '.join(keywords)}")

    prefix = "智慧辨識｜" if date_descs else ""
    return UnifiedReportQuery(
        keywords=tuple(keywords),
        time_windows=tuple(windows),
        status_text=f"{prefix}{' ｜ '.join(status_parts)}",
        has_datetime_filter=bool(windows),
    )


def filter_archive_report_records(
    records: Sequence[ArchiveReportRecord],
    *,
    vendor: str | None = None,
    keyword: str = "",
    keywords: Sequence[str] | None = None,
    start_dt: datetime | None = None,
    end_dt: datetime | None = None,
    time_windows: Sequence[tuple[datetime, datetime]] | None = None,
) -> list[ArchiveReportRecord]:
    """依廠商、關鍵字與指定收件日期時間區間篩選報表紀錄。"""
    kw_tokens = [k.strip().lower() for k in (keywords or ()) if k.strip()]
    if keyword.strip():
        kw_tokens.append(keyword.strip().lower())

    result: list[ArchiveReportRecord] = []
    for rec in records:
        if vendor and rec.vendor != vendor:
            continue
        if time_windows:
            if rec.received_dt is None or not any(
                w_start <= rec.received_dt <= w_end for w_start, w_end in time_windows
            ):
                continue
        if start_dt is not None and (rec.received_dt is None or rec.received_dt < start_dt):
            continue
        if end_dt is not None and (rec.received_dt is None or rec.received_dt > end_dt):
            continue
        if kw_tokens:
            haystack = f"{rec.vendor} {rec.batch_label} {rec.filename} {rec.saved_path}".lower()
            if not all(tok in haystack for tok in kw_tokens):
                continue
        result.append(rec)
    return result


def group_report_by_vendor_date(
    records: Sequence[ArchiveReportRecord],
) -> list[tuple[str, list[tuple[str, datetime | None, list[ArchiveReportRecord]]]]]:
    """依廠商與收件日期時間批次分組，並由新到舊排序。"""
    by_vendor: dict[str, dict[str, list[ArchiveReportRecord]]] = {}
    for rec in records:
        by_vendor.setdefault(rec.vendor, {}).setdefault(rec.batch_label, []).append(rec)

    grouped_vendors: list[
        tuple[datetime, str, list[tuple[str, datetime | None, list[ArchiveReportRecord]]]]
    ] = []
    for vendor, batches_map in by_vendor.items():
        sorted_batches: list[tuple[str, datetime | None, list[ArchiveReportRecord]]] = []
        latest_vendor_dt = datetime.min
        for batch_label, items in batches_map.items():
            batch_dt = max((item.received_dt for item in items if item.received_dt is not None), default=None)
            if batch_dt is not None and batch_dt > latest_vendor_dt:
                latest_vendor_dt = batch_dt
            sorted_items = sorted(items, key=lambda item: item.filename.lower())
            sorted_batches.append((batch_label, batch_dt, sorted_items))

        sorted_batches.sort(
            key=lambda b: (b[1] or datetime.min, b[0]),
            reverse=True,
        )
        grouped_vendors.append((latest_vendor_dt, vendor, sorted_batches))

    grouped_vendors.sort(key=lambda v: (v[0], v[1]), reverse=True)
    return [(vendor, batches) for _, vendor, batches in grouped_vendors]


# ==============================================================================
# Outlook COM 操作區
# ==============================================================================
def _is_outlook_process_running() -> bool:
    """檢查目前是否有 OUTLOOK.EXE 行程正在執行。"""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return "OUTLOOK.EXE" in result.stdout.upper()
    except Exception:
        return False


def _find_outlook_executable() -> str:
    """尋找本機 Classic Outlook (OUTLOOK.EXE) 執行檔路徑。"""
    in_path = shutil.which("OUTLOOK.EXE")
    if in_path:
        return in_path

    candidates = [
        Path(r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"),
        Path(r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"),
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return "OUTLOOK.EXE"


def connect_outlook(
    logger: logging.Logger,
    startup_timeout: int = OUTLOOK_STARTUP_TIMEOUT,
    poll_interval: int = OUTLOOK_POLL_INTERVAL,
    min_init_wait: int = OUTLOOK_MIN_INIT_WAIT,
) -> tuple[Any, Any]:
    """連接執行中的 Classic Outlook，若未開啟則自動啟動並輪詢等待初始化完成。"""
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    import win32timezone  # type: ignore  # noqa: F401  # COM 轉換 ReceivedTime 動態依賴此模組

    pythoncom.CoInitialize()

    def try_attach() -> tuple[Any, Any] | None:
        # 優先透過 ROT 取得已註冊實例；部分情況下 Outlook 未註冊 ROT 但為單實例 COM，
        # 因此確認 OUTLOOK.EXE 正在執行時可安全退回 Dispatch 連接既有實例。
        for getter in (
            lambda: win32com.client.GetActiveObject("Outlook.Application"),
            lambda: (
                win32com.client.Dispatch("Outlook.Application")
                if _is_outlook_process_running()
                else None
            ),
        ):
            try:
                app = getter()
                if app is None:
                    continue
                namespace = app.GetNamespace("MAPI")
                _ = namespace.Folders.Count
                return app, namespace
            except Exception:
                continue
        return None

    attached = try_attach()
    if attached is not None:
        logger.info("已成功連接執行中的 Classic Outlook。")
        return attached

    if not _is_outlook_process_running():
        exe_path = _find_outlook_executable()
        logger.info("未偵測到 OUTLOOK.EXE，正在背景啟動：%s", exe_path)
        creationflags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
            subprocess, "CREATE_NO_WINDOW", 0
        )
        subprocess.Popen([exe_path], creationflags=creationflags)
        logger.info("等待 Outlook 初始化（至少 %d 秒，最長 %d 秒）...", min_init_wait, startup_timeout)
        time.sleep(min_init_wait)
    else:
        logger.info("偵測到 OUTLOOK.EXE 正在啟動或初始化中，輪詢等待 COM 就緒...")

    deadline = time.monotonic() + max(startup_timeout - min_init_wait, 5)
    while time.monotonic() < deadline:
        attached = try_attach()
        if attached is not None:
            logger.info("Outlook COM 初始化完成並已連線。")
            return attached
        time.sleep(poll_interval)

    raise TimeoutError(
        f"等待 {startup_timeout} 秒後仍無法連接 Classic Outlook COM。"
        "請確認是使用傳統版 Outlook（非新版 Outlook），且腳本與 Outlook 權限層級一致。"
    )


def _match_store(namespace: Any, target_account: str) -> Any:
    """依帳號 Email 或 Store 名稱尋找目標信箱 Store。"""
    target_lower = target_account.strip().lower()

    # 先從 Accounts 比對 SmtpAddress / DisplayName
    try:
        for account in namespace.Accounts:
            smtp = str(getattr(account, "SmtpAddress", "") or "").lower()
            disp = str(getattr(account, "DisplayName", "") or "").lower()
            if target_lower in (smtp, disp) or target_lower in smtp:
                delivery_store = getattr(account, "DeliveryStore", None)
                if delivery_store is not None:
                    return delivery_store
    except Exception:
        pass

    # 備援：直接比對 namespace.Stores
    for store in namespace.Stores:
        store_name = str(getattr(store, "DisplayName", "") or "").lower()
        if target_lower == store_name or target_lower in store_name:
            return store

    available = [str(getattr(s, "DisplayName", "")) for s in namespace.Stores]
    raise LookupError(f"找不到目標信箱 '{target_account}'，目前可用信箱：{available}")


def resolve_target_folder(
    namespace: Any,
    account_email: str | None = TARGET_ACCOUNT_EMAIL,
    folder_name: str | None = TARGET_FOLDER_NAME,
) -> Any:
    """取得目標信箱資料夾（預設為預設帳號的收件匣）。"""
    if account_email:
        store = _match_store(namespace, account_email)
    else:
        store = namespace.DefaultStore

    if not folder_name:
        return store.GetDefaultFolder(OL_FOLDER_INBOX)

    parts = [part.strip() for part in folder_name.replace("\\", "/").split("/") if part.strip()]
    current = store.GetRootFolder()
    for part in parts:
        matched = None
        for sub in current.Folders:
            if str(sub.Name).strip().lower() == part.lower():
                matched = sub
                break
        if matched is None:
            # 若使用者直接寫 "收件匣/子資料夾" 而 root 下是 Inbox，嘗試從預設收件匣起算
            if current == store.GetRootFolder() and part.lower() in ("收件匣", "inbox"):
                matched = store.GetDefaultFolder(OL_FOLDER_INBOX)
            else:
                raise LookupError(f"在信箱 '{store.DisplayName}' 中找不到資料夾：'{folder_name}'")
        current = matched
    return current


def _safe_get_property(com_obj: Any, prop_tag: str, default: Any = None) -> Any:
    """安全讀取 MAPI PropertyAccessor 屬性。"""
    try:
        return com_obj.PropertyAccessor.GetProperty(prop_tag)
    except Exception:
        return default


def extract_sender_email(mail: Any) -> str:
    """解析寄件人真實 SMTP Email（支援 Exchange EX 內部地址解析）。"""
    email_type = str(getattr(mail, "SenderEmailType", "") or "").upper()
    raw_addr = str(getattr(mail, "SenderEmailAddress", "") or "").strip()

    if email_type == "EX":
        try:
            sender = mail.Sender
            if sender is not None:
                ex_user = sender.GetExchangeUser()
                if ex_user is not None and ex_user.PrimarySmtpAddress:
                    return str(ex_user.PrimarySmtpAddress).strip()
        except Exception:
            pass

        smtp_prop = _safe_get_property(mail, PR_SMTP_ADDRESS)
        if smtp_prop:
            return str(smtp_prop).strip()

    return raw_addr


def extract_message_key(mail: Any) -> str:
    """取得郵件防重複識別 Key（優先 InternetMessageID，備援 EntryID）。"""
    internet_id = _safe_get_property(mail, PR_INTERNET_MESSAGE_ID)
    if internet_id and str(internet_id).strip():
        return str(internet_id).strip()
    entry_id = str(getattr(mail, "EntryID", "") or "").strip()
    if entry_id:
        return f"entry:{entry_id}"
    subject = str(getattr(mail, "Subject", "") or "")
    received = str(getattr(mail, "ReceivedTime", "") or "")
    return f"fallback:{received}:{subject}"


def _to_naive_datetime(value: Any) -> datetime | None:
    """將 pywintypes.datetime 或 datetime 轉為本地 naive datetime 以便比較。"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    return None


def _get_attachment_flags(att: Any) -> tuple[bool, bool]:
    """檢查附件是否具備 Content-ID 或隱藏旗標（常為內嵌簽名檔圖片）。"""
    cid = _safe_get_property(att, PR_ATTACH_CONTENT_ID)
    hidden = _safe_get_property(att, PR_ATTACHMENT_HIDDEN, False)
    return bool(cid and str(cid).strip()), bool(hidden)


def _extract_mail_body(mail: Any, logger: logging.Logger) -> str:
    """安全讀取郵件純文字內文，僅於第一階段未命中時呼叫。"""
    try:
        return str(getattr(mail, "Body", "") or "")
    except Exception as exc:
        logger.warning("讀取郵件內文失敗（將略過內文比對）：%s", exc)
        return ""


def _extract_mail_html_body(mail: Any, logger: logging.Logger) -> str:
    """安全讀取郵件 HTML 內文。"""
    try:
        return str(getattr(mail, "HTMLBody", "") or "")
    except Exception as exc:
        logger.warning("讀取郵件 HTML 內文失敗（將改用純文字）：%s", exc)
        return ""


_SCRIPT_TAG_RE = re.compile(r"<script\b[^>]*>.*?</script>", re.IGNORECASE | re.DOTALL)
_STYLE_TAG_RE = re.compile(r"<style\b[^>]*>.*?</style>", re.IGNORECASE | re.DOTALL)
_BODY_TAG_RE = re.compile(r"<body\b[^>]*>(.*?)</body>", re.IGNORECASE | re.DOTALL)


def build_yahoo_search_query(
    *,
    sender_email: str,
    sender_name: str,
    subject: str,
    attachment_names: Sequence[str] = (),
) -> str:
    """組合 Yahoo Mail 搜尋字串（以寄件人為主，搭配主旨或第一個附件檔名）。"""
    parts: list[str] = []
    if sender_email:
        parts.append(f"from:{sender_email.strip()}")
    elif sender_name:
        parts.append(f"from:{sender_name.strip()}")

    clean_subj = (subject or "").strip()
    if clean_subj and clean_subj != "(無主旨)":
        parts.append(clean_subj)
    elif attachment_names:
        parts.append(attachment_names[0])
    return " ".join(parts)


def build_yahoo_search_url(query: str) -> str:
    """產生新版 Yahoo Mail (/n/) 網頁版搜尋連結。

    新版 Yahoo Mail SPA 路徑段對關鍵字採兩層 URL 編碼（例如 '@' -> '%2540'）。
    """
    cleaned = query.strip()
    if not cleaned:
        return "https://mail.yahoo.com/"
    encoded_twice = quote(quote(cleaned, safe=""), safe="")
    return f"https://mail.yahoo.com/n/search/keyword={encoded_twice}?.src=ym"


def build_email_html_document(
    *,
    subject: str,
    sender_name: str,
    sender_email: str,
    received_dt: datetime | None,
    matched_reason: str,
    attachment_names: Sequence[str],
    body_text: str = "",
    html_body: str = "",
    message_id: str = "",
) -> str:
    """組合含信件中繼資料摘要、Yahoo Mail 快速搜尋連結與安全內文的單一 UTF-8 HTML 檔案。"""
    received_str = received_dt.strftime("%Y-%m-%d %H:%M:%S") if received_dt else "未知"
    if sender_name and sender_email:
        sender_display = f"{html.escape(sender_name)} &lt;{html.escape(sender_email)}&gt;"
    elif sender_email:
        sender_display = html.escape(sender_email)
    else:
        sender_display = html.escape(sender_name or "未知")

    att_items = "".join(f"<li>{html.escape(name)}</li>" for name in attachment_names) or "<li>(無)</li>"

    search_query = build_yahoo_search_query(
        sender_email=sender_email,
        sender_name=sender_name,
        subject=subject,
        attachment_names=attachment_names,
    )
    yahoo_search_url = build_yahoo_search_url(search_query)
    safe_query = html.escape(search_query)
    safe_msg_id = html.escape(message_id)

    cleaned_html = _SCRIPT_TAG_RE.sub("", html_body or "").strip()
    extracted_styles = ""
    if cleaned_html:
        styles = _STYLE_TAG_RE.findall(cleaned_html)
        if styles:
            extracted_styles = "\n".join(styles)
        body_match = _BODY_TAG_RE.search(cleaned_html)
        body_fragment = body_match.group(1) if body_match else cleaned_html
    else:
        escaped_text = html.escape(body_text or "(無內文)")
        body_fragment = f'<pre class="plain-body">{escaped_text}</pre>'

    safe_subject = html.escape(subject or "(無主旨)")
    safe_reason = html.escape(matched_reason)
    msg_id_row = f"<tr><th>Message-ID</th><td><code>{safe_msg_id}</code></td></tr>" if safe_msg_id else ""

    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <title>{safe_subject}</title>
  <style>
    body {{
      font-family: "Microsoft JhengHei UI", "Segoe UI", sans-serif;
      background-color: #f8fafc;
      color: #1e293b;
      margin: 24px;
      line-height: 1.6;
    }}
    .meta-card {{
      background: #ffffff;
      border: 1px solid #cbd5e1;
      border-left: 5px solid #2563eb;
      border-radius: 8px;
      padding: 16px 20px;
      margin-bottom: 20px;
    }}
    .meta-card h1 {{
      font-size: 18px;
      margin: 0 0 12px 0;
      color: #0f172a;
    }}
    .meta-table {{
      border-collapse: collapse;
      width: 100%;
      font-size: 14px;
    }}
    .meta-table th {{
      text-align: left;
      color: #475569;
      width: 108px;
      padding: 4px 8px 4px 0;
      vertical-align: top;
    }}
    .meta-table td {{
      padding: 4px 0;
      word-break: break-word;
    }}
    .att-list {{
      margin: 0;
      padding-left: 18px;
    }}
    .action-bar {{
      margin-top: 14px;
      padding-top: 12px;
      border-top: 1px solid #e2e8f0;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 10px;
    }}
    .yahoo-btn {{
      display: inline-block;
      background-color: #6001d2;
      color: #ffffff;
      text-decoration: none;
      padding: 7px 14px;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 600;
    }}
    .yahoo-btn:hover {{
      background-color: #4c00a8;
    }}
    .copy-btn {{
      background-color: #f1f5f9;
      color: #1e293b;
      border: 1px solid #cbd5e1;
      padding: 6px 12px;
      border-radius: 6px;
      font-size: 13px;
      cursor: pointer;
    }}
    .copy-btn:hover {{
      background-color: #e2e8f0;
    }}
    .query-pill {{
      background-color: #f1f5f9;
      border: 1px solid #e2e8f0;
      border-radius: 4px;
      padding: 4px 8px;
      font-size: 12px;
      color: #334155;
    }}
    .mail-body {{
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 20px;
      overflow-x: auto;
    }}
    .plain-body {{
      white-space: pre-wrap;
      word-break: break-word;
      font-family: inherit;
      margin: 0;
    }}
  </style>
  {extracted_styles}
</head>
<body>
  <section class="meta-card">
    <h1>{safe_subject}</h1>
    <table class="meta-table">
      <tr><th>寄件人</th><td>{sender_display}</td></tr>
      <tr><th>收件時間</th><td>{html.escape(received_str)}</td></tr>
      <tr><th>命中規則</th><td>{safe_reason}</td></tr>
      {msg_id_row}
      <tr><th>歸檔附件</th><td><ul class="att-list">{att_items}</ul></td></tr>
    </table>
    <div class="action-bar">
      <a class="yahoo-btn" href="{html.escape(yahoo_search_url)}" target="_blank" rel="noopener noreferrer">在 Yahoo Mail 搜尋此信</a>
      <button type="button" class="copy-btn" id="copyQueryBtn">複製搜尋關鍵字</button>
      <code class="query-pill" id="yahooQueryText">{safe_query}</code>
    </div>
  </section>
  <section class="mail-body">
    {body_fragment}
  </section>
  <script>
    (function () {{
      var btn = document.getElementById("copyQueryBtn");
      var qEl = document.getElementById("yahooQueryText");
      if (!btn || !qEl) return;
      btn.addEventListener("click", function () {{
        var text = qEl.textContent || "";
        if (navigator.clipboard && navigator.clipboard.writeText) {{
          navigator.clipboard.writeText(text).then(function () {{
            btn.textContent = "已複製";
            setTimeout(function () {{ btn.textContent = "複製搜尋關鍵字"; }}, 1800);
          }});
        }}
      }});
    }})();
  </script>
</body>
</html>
"""


def build_target_dir(
    output_root: Path,
    target_folder: str,
    received_dt: datetime | None,
    use_datetime_subfolder: bool = USE_DATETIME_SUBFOLDER,
    datetime_format: str = DATETIME_SUBFOLDER_FORMAT,
) -> Path:
    """組合分類目標資料夾路徑，依設定附加收件日期時間子資料夾。"""
    base_dir = output_root / target_folder
    if not use_datetime_subfolder:
        return base_dir
    dt_for_folder = received_dt or datetime.now()
    sub_name = sanitize_filename(
        dt_for_folder.strftime(datetime_format),
        fallback="unknown_datetime",
    )
    return base_dir / sub_name


def process_mailbox(
    folder: Any,
    rules: list[ClassificationRule],
    state_store: StateStore,
    output_root: Path,
    logger: logging.Logger,
    *,
    dry_run: bool = False,
    days_back: int | None = DAYS_BACK,
    limit: int | None = None,
    default_folder: str = DEFAULT_FOLDER,
    ignore_inline_images: bool = IGNORE_INLINE_IMAGES,
    duplicate_mode: str = DUPLICATE_FILENAME_MODE,
    use_datetime_subfolder: bool = USE_DATETIME_SUBFOLDER,
    datetime_format: str = DATETIME_SUBFOLDER_FORMAT,
    save_email_html: bool = SAVE_EMAIL_HTML,
    email_html_filename: str = EMAIL_HTML_FILENAME,
    cancel_requested: Callable[[], bool] | None = None,
    progress_callback: Callable[[dict[str, int], str], None] | None = None,
) -> dict[str, int]:
    """掃描指定資料夾內的郵件並依規則分類下載附件。"""
    stats = {
        "scanned_mails": 0,
        "skipped_checked": 0,
        "saved_attachments": 0,
        "skipped_attachments": 0,
        "errors": 0,
    }

    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception as exc:
        logger.warning("無法排序郵件列表（將依預設順序掃描）：%s", exc)

    cutoff = datetime.now() - timedelta(days=days_back) if days_back is not None else None

    for item in items:
        if cancel_requested and cancel_requested():
            logger.warning("使用者已要求停止掃描。")
            break

        if limit is not None and stats["scanned_mails"] >= limit:
            logger.info("已達測試上限 (%d 封)，停止掃描。", limit)
            break

        try:
            if getattr(item, "Class", None) != OL_MAIL:
                continue

            received_dt = _to_naive_datetime(getattr(item, "ReceivedTime", None))
            if cutoff is not None and received_dt is not None and received_dt < cutoff:
                logger.info("已超過最近 %d 天掃描範圍，提前結束掃描。", days_back)
                break

            stats["scanned_mails"] += 1
            subject = str(getattr(item, "Subject", "") or "(無主旨)")
            sender_name = str(getattr(item, "SenderName", "") or "")
            sender_email = extract_sender_email(item)
            message_key = extract_message_key(item)

            if state_store.is_message_checked(message_key):
                stats["skipped_checked"] += 1
                logger.debug("略過已檢查信件：%s", subject)
                continue

            attachments = item.Attachments
            att_count = int(getattr(attachments, "Count", 0))
            if att_count == 0:
                if not dry_run:
                    state_store.mark_message_checked(message_key, subject, "no_attachments")
                continue

            valid_attachments: list[tuple[int, Any, str, int]] = []
            for idx in range(1, att_count + 1):
                att = attachments.Item(idx)
                raw_filename = str(getattr(att, "FileName", "") or f"attachment_{idx}")
                att_size = int(getattr(att, "Size", 0) or 0)
                att_type = int(getattr(att, "Type", 0) or 0)
                has_cid, is_hidden = _get_attachment_flags(att)

                skip, skip_reason = should_skip_attachment(
                    filename=raw_filename,
                    size_bytes=att_size,
                    att_type=att_type,
                    has_content_id=has_cid,
                    is_hidden=is_hidden,
                    ignore_inline=ignore_inline_images,
                )
                if skip:
                    stats["skipped_attachments"] += 1
                    logger.debug("跳過附件 [%s]：%s", raw_filename, skip_reason)
                    continue
                valid_attachments.append((idx, att, raw_filename, att_size))

            if not valid_attachments:
                if not dry_run:
                    state_store.mark_message_checked(message_key, subject, "no_valid_attachments")
                continue

            target_folder, reason = match_rule(
                rules,
                subject=subject,
                sender_name=sender_name,
                sender_email=sender_email,
                attachment_names=[item_info[2] for item_info in valid_attachments],
                body_getter=lambda m=item: _extract_mail_body(m, logger),
                default_folder=default_folder,
            )

            valid_saved_in_mail = 0
            all_already_downloaded = True
            mail_dest_dir: Path | None = None
            saved_filenames_in_mail: list[str] = []

            for idx, att, raw_filename, att_size in valid_attachments:
                attachment_key = f"{idx}:{raw_filename}:{att_size}"
                if state_store.is_attachment_downloaded(message_key, attachment_key):
                    logger.debug("附件已存在於下載紀錄，略過：%s", raw_filename)
                    continue

                all_already_downloaded = False
                dest_dir = build_target_dir(
                    output_root=output_root,
                    target_folder=target_folder,
                    received_dt=received_dt,
                    use_datetime_subfolder=use_datetime_subfolder,
                    datetime_format=datetime_format,
                )
                mail_dest_dir = dest_dir
                dest_path = resolve_unique_path(
                    dest_dir,
                    raw_filename,
                    received_dt=received_dt,
                    mode=duplicate_mode,
                )

                if dry_run:
                    logger.info(
                        "[DRY-RUN] 信件 '%s' (%s <%s>) 命中 [%s] -> 預計儲存至：%s",
                        subject,
                        sender_name,
                        sender_email,
                        reason,
                        dest_path,
                    )
                    valid_saved_in_mail += 1
                    stats["saved_attachments"] += 1
                    saved_filenames_in_mail.append(dest_path.name)
                else:
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    att.SaveAsFile(str(dest_path.resolve()))
                    state_store.mark_attachment_downloaded(
                        message_key,
                        attachment_key,
                        dest_path,
                        target_folder,
                    )
                    logger.info(
                        "已儲存附件：%s（信件：'%s' | 寄件人：%s <%s> | 命中：%s）",
                        dest_path,
                        subject,
                        sender_name,
                        sender_email,
                        reason,
                    )
                    valid_saved_in_mail += 1
                    stats["saved_attachments"] += 1
                    saved_filenames_in_mail.append(dest_path.name)

            if save_email_html and valid_saved_in_mail > 0 and mail_dest_dir is not None:
                html_dest = resolve_unique_path(
                    mail_dest_dir,
                    email_html_filename,
                    received_dt=received_dt,
                    mode=duplicate_mode,
                )
                if dry_run:
                    logger.info("[DRY-RUN] 預計產生信件紀錄：%s", html_dest)
                else:
                    display_msg_id = (
                        ""
                        if message_key.startswith(("entry:", "fallback:"))
                        else message_key
                    )
                    html_content = build_email_html_document(
                        subject=subject,
                        sender_name=sender_name,
                        sender_email=sender_email,
                        received_dt=received_dt,
                        matched_reason=reason,
                        attachment_names=saved_filenames_in_mail,
                        body_text=_extract_mail_body(item, logger),
                        html_body=_extract_mail_html_body(item, logger),
                        message_id=display_msg_id,
                    )
                    mail_dest_dir.mkdir(parents=True, exist_ok=True)
                    html_dest.write_text(html_content, encoding="utf-8")
                    logger.info("已產生信件紀錄：%s", html_dest)

            if not dry_run and (valid_saved_in_mail > 0 or all_already_downloaded):
                status = "downloaded" if valid_saved_in_mail > 0 else "no_valid_attachments"
                state_store.mark_message_checked(message_key, subject, status)

            if progress_callback:
                progress_callback(dict(stats), subject)

        except Exception as exc:
            stats["errors"] += 1
            logger.exception("處理單封信件時發生錯誤：%s", exc)
            if progress_callback:
                progress_callback(dict(stats), "處理錯誤")

    return stats


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Outlook 附件自動分類歸檔工具")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="模擬執行並輸出預期分類路徑，不實際建立資料夾、下載檔案或寫入 state.db",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=DAYS_BACK,
        help="覆寫掃描最近天數（設為 0 表示不限天數）",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="限制最多掃描幾封郵件（測試用）",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="顯示詳細除錯日誌",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {APP_VERSION}",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)
    logger = setup_logging(LOG_FILE_PATH, verbose=args.verbose)
    days_back = None if args.days == 0 else args.days

    logger.info(
        "=== 開始執行 Outlook 附件歸檔工具 %s (dry_run=%s, days_back=%s) ===",
        APP_VERSION,
        args.dry_run,
        days_back,
    )

    try:
        rules = parse_rules_markdown(RULES_FILE)
        logger.info("已載入 %d 條分類規則：%s", len(rules), [r.folder_name for r in rules])
    except Exception as exc:
        logger.error("讀取分類規則失敗：%s", exc)
        return 1

    state_store = StateStore(STATE_DB_PATH)
    try:
        _, namespace = connect_outlook(logger)
        folder = resolve_target_folder(namespace, TARGET_ACCOUNT_EMAIL, TARGET_FOLDER_NAME)
        logger.info("已定位目標資料夾：%s", getattr(folder, "FolderPath", getattr(folder, "Name", "")))

        stats = process_mailbox(
            folder=folder,
            rules=rules,
            state_store=state_store,
            output_root=OUTPUT_ROOT_DIR,
            logger=logger,
            dry_run=args.dry_run,
            days_back=days_back,
            limit=args.limit,
        )
        logger.info(
            "=== 執行完成 | 掃描信件: %d | 略過已檢查: %d | 儲存附件: %d | 排除附件: %d | 錯誤: %d ===",
            stats["scanned_mails"],
            stats["skipped_checked"],
            stats["saved_attachments"],
            stats["skipped_attachments"],
            stats["errors"],
        )
        return 0 if stats["errors"] == 0 else 2
    except Exception as exc:
        logger.exception("執行過程發生未預期錯誤：%s", exc)
        return 1
    finally:
        state_store.close()


if __name__ == "__main__":
    raise SystemExit(main())
