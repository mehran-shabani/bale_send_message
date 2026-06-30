from pathlib import Path
import tempfile
from unittest.mock import Mock, patch

from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase, override_settings
from openpyxl import Workbook

from .core import BaleSafirClient, ExcelRecipient, normalize_iran_mobile, read_excel_recipients, render_message, run_excel_batch
from .forms import UploadExcelForm
from .models import MessageRecipient


TEST_PHONE_LOCAL = "09120000000"
TEST_PHONE_STD = "989120000000"


class PhoneTests(TestCase):
    def test_normalize_iran_mobile(self):
        self.assertEqual(normalize_iran_mobile(TEST_PHONE_LOCAL), TEST_PHONE_STD)
        self.assertEqual(normalize_iran_mobile("+" + TEST_PHONE_STD), TEST_PHONE_STD)
        self.assertEqual(normalize_iran_mobile("۹۱۲۰۰۰۰۰۰۰"), TEST_PHONE_STD)
        self.assertIsNone(normalize_iran_mobile("123"))


class ExcelTests(TestCase):
    def make_excel(self):
        tmp = Path(tempfile.mkdtemp()) / "sample.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["ردیف", "نام", "نام خانوادگی", "موبایل"])
        ws.append([1, "علی", "رضایی", TEST_PHONE_LOCAL])
        ws.append([2, "زهرا", "احمدی", "123"])
        wb.save(tmp)
        return tmp

    def test_read_recipients(self):
        rows = read_excel_recipients(self.make_excel())
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0].full_name, "علی رضایی")
        self.assertEqual(rows[0].normalized_phone, TEST_PHONE_STD)
        self.assertIsNone(rows[1].normalized_phone)

    def test_read_recipients_invalid_sheet_message(self):
        with self.assertRaisesMessage(ValueError, "شیت «missing» در فایل اکسل پیدا نشد"):
            read_excel_recipients(self.make_excel(), sheet_name="missing")

    def test_read_recipients_missing_required_column_message(self):
        tmp = Path(tempfile.mkdtemp()) / "missing.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["نام", "موبایل"])
        wb.save(tmp)

        with self.assertRaisesMessage(ValueError, "ستون «نام خانوادگی»"):
            read_excel_recipients(tmp)

    def test_read_recipients_empty_header_message(self):
        tmp = Path(tempfile.mkdtemp()) / "empty_header.xlsx"
        wb = Workbook()
        wb.active.append([None, None, None])
        wb.save(tmp)

        with self.assertRaisesMessage(ValueError, "ردیف اول فایل اکسل باید header"):
            read_excel_recipients(tmp)


class MessageTemplateTests(TestCase):
    def test_render_message_allows_only_known_placeholders(self):
        recipient = ExcelRecipient(2, "علی", "رضایی", TEST_PHONE_LOCAL, TEST_PHONE_STD)
        self.assertEqual(render_message("سلام {full_name} {phone}", recipient), f"سلام علی رضایی {TEST_PHONE_STD}")

        with self.assertRaisesMessage(ValueError, "placeholderهای نامعتبر"):
            render_message("سلام {unknown}", recipient)

    def test_form_rejects_invalid_template_and_large_file(self):
        upload = SimpleUploadedFile("sample.xlsx", b"x" * 6)
        with override_settings(BALE_MAX_UPLOAD_SIZE_MB=0):
            form = UploadExcelForm(data={"message_template": "سلام {unknown}", "send_mode": "dry_run"}, files={"excel_file": upload})
            self.assertFalse(form.is_valid())
            self.assertIn("message_template", form.errors)
            self.assertIn("excel_file", form.errors)


class SafirClientTests(TestCase):
    @override_settings(BALE_API_ACCESS_KEY="test-key", BALE_BOT_ID=123456, BALE_SEND_URL="https://example.test/api", BALE_REQUEST_TIMEOUT=5)
    @patch("bale_sender.core.requests.post")
    def test_send_success(self, post_mock):
        resp = Mock()
        resp.status_code = 200
        resp.text = "{}"
        resp.json.return_value = {"ok": True}
        post_mock.return_value = resp
        result = BaleSafirClient().send(phone_number=TEST_PHONE_STD, text="سلام", request_id="r1")
        self.assertEqual(result["status"], MessageRecipient.Status.SENT)

    @override_settings(BALE_API_ACCESS_KEY="test-key", BALE_BOT_ID=123456, BALE_SEND_URL="https://example.test/api", BALE_REQUEST_TIMEOUT=5)
    @patch("bale_sender.core.requests.post")
    def test_not_bale_user(self, post_mock):
        resp = Mock()
        resp.status_code = 400
        resp.text = "{}"
        resp.json.return_value = {"code": "NotBaleUser", "message": "user not found"}
        post_mock.return_value = resp
        result = BaleSafirClient().send(phone_number=TEST_PHONE_STD, text="سلام", request_id="r1")
        self.assertEqual(result["status"], MessageRecipient.Status.NOT_BALE_USER)


class BatchTests(TestCase):
    def make_excel(self):
        tmp = Path(tempfile.mkdtemp()) / "sample.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["ردیف", "نام", "نام خانوادگی", "موبایل"])
        ws.append([1, "علی", "رضایی", TEST_PHONE_LOCAL])
        ws.append([2, "علی", "رضایی", TEST_PHONE_LOCAL])
        ws.append([3, "بد", "شماره", "123"])
        wb.save(tmp)
        return tmp

    def test_dry_run_batch_report(self):
        report = Path(tempfile.mkdtemp()) / "report.xlsx"
        batch = run_excel_batch(file_path=str(self.make_excel()), message_template="سلام {full_name}", dry_run=True, report_path=str(report))
        self.assertEqual(batch.total_rows, 3)
        self.assertEqual(batch.recipients.filter(status=MessageRecipient.Status.DRY_RUN).count(), 1)
        self.assertEqual(batch.recipients.filter(status=MessageRecipient.Status.DUPLICATE).count(), 1)
        self.assertEqual(batch.recipients.filter(status=MessageRecipient.Status.INVALID_PHONE).count(), 1)
        self.assertTrue(report.exists())
