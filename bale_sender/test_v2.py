import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

from django.test import TestCase, override_settings
from django.urls import reverse
from openpyxl import Workbook

from .core import (
    BaleSafirClient,
    ExcelRecipient,
    _build_message_data,
    read_excel_recipients,
    render_message,
)
from .models import MessageBatch, MessageRecipient, SafirConfiguration


class DynamicExcelTests(TestCase):
    def make_excel(self):
        path = Path(tempfile.mkdtemp()) / "dynamic.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["شماره تماس اصلی", "نام بیمار", "تاریخ مراجعه", "کد"])
        sheet.append(["09123456789", "علی رضایی", "۱۴۰۵/۰۷/۰۱", 1234])
        workbook.save(path)
        return path

    def test_arbitrary_columns_can_be_mapped_to_tokens(self):
        recipient = read_excel_recipients(
            self.make_excel(),
            phone_column="A",
            variable_mapping={"name": "B", "date": "تاریخ مراجعه", "code": "D"},
        )[0]
        text = render_message(
            "سلام {{name}}؛ تاریخ {{date}} و کد {{code}}",
            recipient,
        )
        self.assertEqual(text, "سلام علی رضایی؛ تاریخ ۱۴۰۵/۰۷/۰۱ و کد 1234")
        self.assertEqual(recipient.normalized_phone, "989123456789")

    def test_template_message_renders_each_text_field(self):
        batch = MessageBatch(
            message_type=MessageBatch.MessageType.TEMPLATE,
            template_id="visit-reminder",
            template_fields={"name": "{{name}}", "date": "{{date}}"},
            global_variables={},
            message_template="",
        )
        recipient = ExcelRecipient(
            2,
            "",
            "",
            "09123456789",
            "989123456789",
            {"name": "علی", "date": "فردا"},
        )
        final_text, data = _build_message_data(batch, recipient)
        self.assertIn("visit-reminder", final_text)
        self.assertEqual(
            data["template_message"]["text_fields"],
            {"name": "علی", "date": "فردا"},
        )

    def test_text_buttons_and_secure_payload_follow_safir_schema(self):
        batch = MessageBatch(
            message_type=MessageBatch.MessageType.TEXT,
            message_template="سلام {{name}}",
            global_variables={},
            is_secure=True,
            copy_text_template="{{code}}",
            buttons=[
                {
                    "text": "ورود",
                    "type": "url",
                    "value": "https://example.com/{{code}}",
                    "row": 1,
                },
                {"text": "کپی", "type": "copy_text", "value": "{{code}}", "row": 2},
            ],
        )
        recipient = ExcelRecipient(
            2,
            "",
            "",
            "09123456789",
            "989123456789",
            {"name": "زهرا", "code": "ABC123"},
        )
        _, data = _build_message_data(batch, recipient)
        self.assertTrue(data["is_secure"])
        self.assertEqual(data["message"]["copy_text"], "ABC123")
        keyboard = data["message"]["reply_markup"]["inline_keyboard"]
        self.assertEqual(keyboard[0][0]["url"], "https://example.com/ABC123")
        self.assertEqual(keyboard[1][0]["copy_text"], "ABC123")

    def test_otp_converts_persian_digits(self):
        batch = MessageBatch(
            message_type=MessageBatch.MessageType.OTP,
            otp_template="{{otp}}",
            global_variables={},
            message_template="",
        )
        recipient = ExcelRecipient(
            2, "", "", "09123456789", "989123456789", {"otp": "۱۲۳۴۵۶"}
        )
        _, data = _build_message_data(batch, recipient)
        self.assertEqual(data["otp_message"]["otp"], "123456")


@override_settings(
    BALE_API_ACCESS_KEY="test-key",
    BALE_BOT_ID=123,
    BALE_SEND_URL="https://example.test/send",
    BALE_REQUEST_TIMEOUT=2,
)
class SafirV3ResponseTests(TestCase):
    @patch("bale_sender.core.requests.post")
    def test_error_data_is_classified_even_with_http_200(self, post_mock):
        response = Mock(status_code=200, text="")
        response.json.return_value = {
            "message_id": "",
            "error_data": [
                {
                    "phone_number": "989123456789",
                    "code": 17,
                    "description": "کاربر بله نیست",
                }
            ],
        }
        post_mock.return_value = response
        result = BaleSafirClient(max_retries=0).send_data(
            phone_number="989123456789",
            request_id="stable-id",
            message_data={"message": {"text": "سلام"}},
        )
        self.assertEqual(result["status"], MessageRecipient.Status.NOT_BALE_USER)
        self.assertEqual(result["api_code"], "NotBaleUser")

    @patch("bale_sender.core.sleep")
    @patch("bale_sender.core.requests.post")
    def test_retry_reuses_same_request_id(self, post_mock, _sleep_mock):
        limited = Mock(status_code=429, text="")
        limited.json.return_value = {
            "error_data": [{"code": 3, "description": "rate limit"}]
        }
        success = Mock(status_code=200, text="")
        success.json.return_value = {"message_id": "message-1", "error_data": None}
        post_mock.side_effect = [limited, success]

        result = BaleSafirClient(max_retries=1).send_data(
            phone_number="989123456789",
            request_id="stable-id",
            message_data={"message": {"text": "سلام"}},
        )

        self.assertEqual(result["status"], MessageRecipient.Status.SENT)
        self.assertEqual(result["retry_count"], 1)
        self.assertEqual(
            post_mock.call_args_list[0].kwargs["json"]["request_id"],
            post_mock.call_args_list[1].kwargs["json"]["request_id"],
        )


class LocalSettingsViewTests(TestCase):
    def test_settings_page_has_no_login_redirect(self):
        response = self.client.get(reverse("bale_settings"))
        self.assertEqual(response.status_code, 200)

    def test_settings_can_store_key_and_global_variables(self):
        response = self.client.post(
            reverse("bale_settings"),
            {
                "api_access_key": "local-secret",
                "bot_id": 123456,
                "send_url": "https://safir.bale.ai/api/v3/send_message",
                "upload_url": "https://safir.bale.ai/api/v3/upload_file",
                "request_timeout": 20,
                "default_sleep_seconds": 0.4,
                "max_retries": 2,
                "message_price_rial": 100,
                "global_variables_text": '{"clinic":"ولیعصر"}',
            },
        )
        self.assertRedirects(response, reverse("bale_settings"))
        config = SafirConfiguration.load()
        self.assertEqual(config.api_access_key, "local-secret")
        self.assertEqual(config.global_variables, {"clinic": "ولیعصر"})
