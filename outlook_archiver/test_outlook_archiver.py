"""outlook_archiver 單元測試與模擬整合測試。"""

from datetime import datetime, timedelta
import importlib.util
import logging
from pathlib import Path
import sys
import tempfile
import unittest

_MODULE_PATH = Path(__file__).resolve().parent / "outlook_archiver.py"
_SPEC = importlib.util.spec_from_file_location("outlook_archiver_mod", _MODULE_PATH)
assert _SPEC and _SPEC.loader
oa = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = oa
_SPEC.loader.exec_module(oa)


class DummyPropertyAccessor:
    def __init__(self, props: dict[str, object] | None = None) -> None:
        self._props = props or {}

    def GetProperty(self, tag: str) -> object:
        if tag not in self._props:
            raise KeyError(tag)
        return self._props[tag]


class DummyAttachment:
    def __init__(
        self,
        filename: str,
        size: int,
        att_type: int = oa.OL_BY_VALUE,
        content_id: str | None = None,
        hidden: bool = False,
        payload: bytes = b"dummy-cad-data",
    ) -> None:
        self.FileName = filename
        self.Size = size
        self.Type = att_type
        self._payload = payload
        props: dict[str, object] = {}
        if content_id is not None:
            props[oa.PR_ATTACH_CONTENT_ID] = content_id
        if hidden:
            props[oa.PR_ATTACHMENT_HIDDEN] = True
        self.PropertyAccessor = DummyPropertyAccessor(props)

    def SaveAsFile(self, path: str) -> None:
        Path(path).write_bytes(self._payload)


class DummyAttachments:
    def __init__(self, items: list[DummyAttachment]) -> None:
        self._items = items

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, idx: int) -> DummyAttachment:
        return self._items[idx - 1]


class DummyExchangeUser:
    def __init__(self, smtp: str) -> None:
        self.PrimarySmtpAddress = smtp


class DummySender:
    def __init__(self, smtp: str | None) -> None:
        self._smtp = smtp

    def GetExchangeUser(self):
        return DummyExchangeUser(self._smtp) if self._smtp else None


class DummyMailItem:
    def __init__(
        self,
        subject: str,
        sender_name: str,
        sender_email: str,
        attachments: list[DummyAttachment],
        *,
        received_time: datetime | None = None,
        email_type: str = "SMTP",
        exchange_smtp: str | None = None,
        internet_id: str | None = None,
        entry_id: str = "entry-1",
        item_class: int = oa.OL_MAIL,
        body: str = "",
    ) -> None:
        self.Class = item_class
        self.Subject = subject
        self.SenderName = sender_name
        self.SenderEmailAddress = sender_email
        self.SenderEmailType = email_type
        self.Sender = DummySender(exchange_smtp)
        self.ReceivedTime = received_time or datetime.now()
        self.Attachments = DummyAttachments(attachments)
        self.EntryID = entry_id
        self.Body = body
        props: dict[str, object] = {}
        if internet_id:
            props[oa.PR_INTERNET_MESSAGE_ID] = internet_id
        self.PropertyAccessor = DummyPropertyAccessor(props)


class DummyItems(list):
    def Sort(self, _field: str, _descending: bool = False) -> None:
        return None


class DummyFolder:
    def __init__(self, items: list[object]) -> None:
        self.Items = DummyItems(items)


class TestOutlookArchiver(unittest.TestCase):
    def test_parse_rules_markdown_handles_full_and_half_width_delimiters(self) -> None:
        md = (
            "# 規則\n"
            "## 合宏展\n"
            "- **主旨包含**：合宏展、報價\n"
            "- **寄件人包含**: 合宏展\n"
            "- **Email包含**: \n\n"
            "## 單井/測試*\n"
            "- **主旨包含**: \n"
            "- **寄件人包含**: 籃梅菁, 單井-籃梅菁、 單井\n"
            "- **Email包含**: singlewell \n"
            "- **附件包含**: SW- , sw_part\n"
            "- **內文包含**: 單井精密, SW_BODY\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            rules_path = Path(tmp) / "rules.md"
            rules_path.write_text(md, encoding="utf-8")
            rules = oa.parse_rules_markdown(rules_path)

        self.assertEqual(len(rules), 2)
        self.assertEqual(rules[0].folder_name, "合宏展")
        self.assertEqual(rules[0].subjects, ("合宏展", "報價"))
        self.assertEqual(rules[0].emails, ())
        self.assertEqual(rules[0].attachments, ())
        self.assertEqual(rules[0].bodies, ())
        self.assertEqual(rules[1].folder_name, "單井_測試_")
        self.assertEqual(rules[1].senders, ("籃梅菁", "單井-籃梅菁", "單井"))
        self.assertEqual(rules[1].emails, ("singlewell",))
        self.assertEqual(rules[1].attachments, ("SW-", "sw_part"))
        self.assertEqual(rules[1].bodies, ("單井精密", "SW_BODY"))

    def test_match_rule_order_case_insensitive_and_default(self) -> None:
        rules = [
            oa.ClassificationRule("合宏展", ("合宏展",), ("合宏展",), (), ("HHJ-",), ("合宏展內文",)),
            oa.ClassificationRule("單井", (), ("單井",), ("singlewell",), ("SW-",), ("單井內文",)),
        ]
        called = {"body": 0}

        def get_body_should_not_run() -> str:
            called["body"] += 1
            return "這裡面有合宏展內文"

        folder1, _ = oa.match_rule(
            rules, "Re: 合宏展 單井圖檔", "王小明", "test@example.com", body_getter=get_body_should_not_run
        )
        self.assertEqual(folder1, "合宏展")  # 由上往下第一條命中即停
        self.assertEqual(called["body"], 0)  # 第一階段命中時不應讀取內文

        folder2, _ = oa.match_rule(rules, "CAD 圖檔", "Alice", "Sales@SINGLEWELL.com.tw")
        self.assertEqual(folder2, "單井")  # 不分大小寫

        folder3, _ = oa.match_rule(
            rules, "(無主旨)", "Bob", "bob@other.com", attachment_names=["hhj-101.dwg", "說明.pdf"]
        )
        self.assertEqual(folder3, "合宏展")  # 附件檔名命中且不分大小寫

        # 兩階段測試：第二條的「主旨/寄件人/Email/附件」應優先於第一條的「內文包含」
        folder4, _ = oa.match_rule(
            rules,
            "請看圖檔",
            "單井業務",
            "user@other.com",
            attachment_names=["part.dwg"],
            body="備註：轉寄自合宏展內文",
        )
        self.assertEqual(folder4, "單井")

        # 第一階段全未命中時，第二階段才比對內文
        folder5, reason5 = oa.match_rule(
            rules,
            "(無主旨)",
            "Bob",
            "bob@other.com",
            attachment_names=["part.dwg"],
            body="您好，這是 單井內文 請查收",
        )
        self.assertEqual(folder5, "單井")
        self.assertIn("內文包含", reason5)

        folder6, _ = oa.match_rule(
            rules, "一般通知", "Bob", "bob@other.com", attachment_names=["other.dwg"], body="無關內容"
        )
        self.assertEqual(folder6, oa.DEFAULT_FOLDER)

    def test_resolve_unique_path_suffix_counter(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            target_dir = Path(tmp)
            (target_dir / "圖面.dwg").write_bytes(b"0")
            (target_dir / "圖面_1.dwg").write_bytes(b"1")
            resolved = oa.resolve_unique_path(target_dir, "圖面.dwg", mode="suffix_counter")
            self.assertEqual(resolved.name, "圖面_2.dwg")

    def test_should_skip_attachment_inline_and_type_filters(self) -> None:
        skip1, _ = oa.should_skip_attachment("image001.png", 5000, oa.OL_BY_VALUE, False, False, True)
        self.assertTrue(skip1)

        skip2, _ = oa.should_skip_attachment("image001.png", 25000, oa.OL_BY_VALUE, False, False, True)
        self.assertFalse(skip2)

        skip3, _ = oa.should_skip_attachment("logo.jpg", 12000, oa.OL_BY_VALUE, True, False, True)
        self.assertTrue(skip3)

        skip4, _ = oa.should_skip_attachment("圖檔.dwg", 1024, 5, False, False, True)
        self.assertTrue(skip4)

        skip5, _ = oa.should_skip_attachment("圖檔.dwg", 1024, oa.OL_BY_VALUE, False, False, True)
        self.assertFalse(skip5)

    def test_process_mailbox_end_to_end_and_dedup(self) -> None:
        rules = [
            oa.ClassificationRule("合宏展", ("合宏展",), ("合宏展",), ()),
            oa.ClassificationRule("單井", (), ("籃梅菁",), ("singlewell",)),
        ]
        logger = logging.getLogger("test_archiver")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        now = datetime.now().replace(hour=14, minute=30, second=45, microsecond=0)
        earlier = (now - timedelta(days=1)).replace(hour=9, minute=15, second=0, microsecond=0)
        expected_now_sub = now.strftime("%Y-%m-%d_PM0230")
        expected_earlier_sub = earlier.strftime("%Y-%m-%d_AM0915")
        mails = [
            DummyMailItem(
                subject="合宏展 零件圖面",
                sender_name="陳先生",
                sender_email="/O=ORG/OU=EXCHANGE/CN=RECIPIENTS/CN=USER1",
                email_type="EX",
                exchange_smtp="chen@herhong.com",
                internet_id="<msg-1@example.com>",
                attachments=[
                    DummyAttachment("image001.png", 4096),  # 應排除
                    DummyAttachment("A01.dwg", 40960, payload=b"dwg-1"),
                ],
                received_time=now,
            ),
            DummyMailItem(
                subject="新圖面請確認",
                sender_name="王小姐",
                sender_email="wang@SINGLEWELL.com.tw",
                internet_id="<msg-2@example.com>",
                attachments=[
                    DummyAttachment("A01.dwg", 20480, payload=b"dwg-2"),
                ],
                received_time=earlier,
            ),
            DummyMailItem(
                subject="很久以前的信",
                sender_name="合宏展",
                sender_email="old@herhong.com",
                internet_id="<msg-old@example.com>",
                attachments=[DummyAttachment("old.dwg", 10000)],
                received_time=now - timedelta(days=30),
            ),
        ]

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            db_path = root / "state.db"
            output_root = root / "out"
            store = oa.StateStore(db_path)
            try:
                # 先測 dry-run：不寫檔也不寫 state.db
                dry_stats = oa.process_mailbox(
                    DummyFolder(mails),
                    rules,
                    store,
                    output_root,
                    logger,
                    dry_run=True,
                    days_back=14,
                )
                self.assertEqual(dry_stats["saved_attachments"], 2)
                self.assertFalse((output_root / "合宏展").exists())

                # 正式執行：下載 2 個附件，並依收件日期時間分別建子資料夾
                real_stats = oa.process_mailbox(
                    DummyFolder(mails),
                    rules,
                    store,
                    output_root,
                    logger,
                    dry_run=False,
                    days_back=14,
                )
                self.assertEqual(real_stats["saved_attachments"], 2)
                self.assertEqual(real_stats["skipped_attachments"], 1)
                self.assertTrue((output_root / "合宏展" / expected_now_sub / "A01.dwg").exists())
                self.assertTrue((output_root / "單井" / expected_earlier_sub / "A01.dwg").exists())

                # 再次執行：應全部因防重複機制被略過
                second_stats = oa.process_mailbox(
                    DummyFolder(mails),
                    rules,
                    store,
                    output_root,
                    logger,
                    dry_run=False,
                    days_back=14,
                )
                self.assertEqual(second_stats["saved_attachments"], 0)
                self.assertEqual(second_stats["skipped_checked"], 2)
            finally:
                store.close()

    def test_process_mailbox_whole_mail_attachment_name_rule(self) -> None:
        rules = [
            oa.ClassificationRule("合宏展", ("合宏展",), ("合宏展",), (), ("HHJ-",)),
        ]
        logger = logging.getLogger("test_archiver_att_rule")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        now = datetime.now().replace(hour=12, minute=0, second=0, microsecond=0)
        expected_sub = now.strftime("%Y-%m-%d_%p%I%M")
        mails = [
            DummyMailItem(
                subject="(無主旨)",
                sender_name="外部寄件人",
                sender_email="someone@gmail.com",
                internet_id="<msg-att-1@example.com>",
                attachments=[
                    DummyAttachment("image001.png", 2048),  # 內嵌圖應被排除
                    DummyAttachment("HHJ-999.dwg", 30720),  # 命中附件包含
                    DummyAttachment("加工說明.pdf", 40960),  # 整封信連動一起進合宏展
                ],
                received_time=now,
            ),
        ]

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            store = oa.StateStore(root / "state.db")
            try:
                stats = oa.process_mailbox(
                    DummyFolder(mails),
                    rules,
                    store,
                    root / "out",
                    logger,
                    dry_run=False,
                    days_back=14,
                )
                self.assertEqual(stats["saved_attachments"], 2)
                self.assertEqual(stats["skipped_attachments"], 1)
                self.assertTrue((root / "out" / "合宏展" / expected_sub / "HHJ-999.dwg").exists())
                self.assertTrue((root / "out" / "合宏展" / expected_sub / "加工說明.pdf").exists())
            finally:
                store.close()

    def test_process_mailbox_body_fallback_rule(self) -> None:
        rules = [
            oa.ClassificationRule("合宏展", ("合宏展",), (), (), (), ("案號:HHJ",)),
        ]
        logger = logging.getLogger("test_archiver_body_rule")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        now = datetime.now().replace(hour=10, minute=5, second=0, microsecond=0)
        expected_sub = now.strftime("%Y-%m-%d_%p%I%M")
        mails = [
            DummyMailItem(
                subject="(無主旨)",
                sender_name="外部寄件人",
                sender_email="someone@gmail.com",
                internet_id="<msg-body-1@example.com>",
                attachments=[DummyAttachment("drawing.dwg", 20480)],
                received_time=now,
                body="請協助加工 案號:hhj 零件，謝謝。",
            ),
        ]

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            store = oa.StateStore(root / "state.db")
            try:
                stats = oa.process_mailbox(
                    DummyFolder(mails),
                    rules,
                    store,
                    root / "out",
                    logger,
                    dry_run=False,
                    days_back=14,
                )
                self.assertEqual(stats["saved_attachments"], 1)
                self.assertTrue((root / "out" / "合宏展" / expected_sub / "drawing.dwg").exists())
            finally:
                store.close()


if __name__ == "__main__":
    unittest.main()
