from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from time import sleep
from uuid import uuid4
import json
import re

import requests
from django.conf import settings
from django.utils import timezone
from openpyxl import Workbook, load_workbook

from .models import MessageBatch, MessageRecipient

PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


@dataclass
class ExcelRecipient:
    row_number: int
    first_name: str
    last_name: str
    raw_phone: str
    normalized_phone: str | None

    @property
    def full_name(self) -> str:
        return " ".join(x for x in [self.first_name, self.last_name] if x).strip()


def normalize_iran_mobile(value: object) -> str | None:
    if value is None:
        return None
    s = str(value).strip().translate(PERSIAN_DIGITS)
    s = re.sub(r"[^0-9+]", "", s)
    if s.startswith("+98"):
        s = "98" + s[3:]
    elif s.startswith("0098"):
        s = "98" + s[4:]
    elif s.startswith("09") and len(s) == 11:
        s = "98" + s[1:]
    elif s.startswith("9") and len(s) == 10:
        s = "98" + s
    if re.fullmatch(r"989\d{9}", s):
        return s
    return None


def _clean_header(value: object) -> str:
    return str(value or "").strip().replace("ي", "ی").replace("ك", "ک")


def _find_header(headers: list[str], candidates: set[str]) -> int:
    for i, h in enumerate(headers):
        if h in candidates:
            return i
    raise ValueError(f"ستون مورد نیاز پیدا نشد. ستون‌های فایل: {headers}")


def read_excel_recipients(file_path: str | Path, sheet_name: str | None = None) -> list[ExcelRecipient]:
    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        headers = [_clean_header(x) for x in next(rows)]
    except StopIteration:
        return []

    first_idx = _find_header(headers, {"نام", "اسم", "first_name"})
    last_idx = _find_header(headers, {"نام خانوادگی", "فامیلی", "last_name"})
    phone_idx = _find_header(headers, {"موبایل", "شماره موبایل", "شماره همراه", "mobile", "phone"})

    result: list[ExcelRecipient] = []
    for row_number, row in enumerate(rows, start=2):
        first_name = str(row[first_idx] or "").strip()
        last_name = str(row[last_idx] or "").strip()
        raw_phone = str(row[phone_idx] or "").strip()
        if not any([first_name, last_name, raw_phone]):
            continue
        result.append(
            ExcelRecipient(
                row_number=row_number,
                first_name=first_name,
                last_name=last_name,
                raw_phone=raw_phone,
                normalized_phone=normalize_iran_mobile(raw_phone),
            )
        )
    return result


def render_message(template: str, recipient: ExcelRecipient) -> str:
    return template.format(
        first_name=recipient.first_name,
        last_name=recipient.last_name,
        full_name=recipient.full_name,
        phone=recipient.normalized_phone or recipient.raw_phone,
    )


class BaleSafirClient:
    def __init__(self) -> None:
        self.url = settings.BALE_SEND_URL
        self.api_key = settings.BALE_API_ACCESS_KEY
        self.bot_id = settings.BALE_BOT_ID
        self.timeout = settings.BALE_REQUEST_TIMEOUT

    def send(self, *, phone_number: str, text: str, request_id: str, button_text: str | None = None, button_url: str | None = None) -> dict:
        if not self.api_key or not self.bot_id:
            return {"status": MessageRecipient.Status.CONFIG_ERROR, "error": "BALE_API_ACCESS_KEY یا BALE_BOT_ID تنظیم نشده است."}

        message: dict = {"text": text}
        if button_text and button_url:
            message["reply_markup"] = {"inline_keyboard": [[{"text": button_text, "url": button_url}]]}

        payload = {
            "request_id": request_id,
            "bot_id": self.bot_id,
            "phone_number": phone_number,
            "message_data": {"message": message},
        }
        headers = {"api-access-key": self.api_key, "Content-Type": "application/json"}

        try:
            response = requests.post(self.url, headers=headers, json=payload, timeout=self.timeout)
            try:
                data = response.json()
            except ValueError:
                data = {"raw": response.text}
        except requests.RequestException as exc:
            return {"status": MessageRecipient.Status.FAILED, "error": str(exc)}

        code = str(data.get("code") or data.get("error") or "")
        message_text = str(data.get("message") or data.get("description") or "")
        raw = json.dumps(data, ensure_ascii=False)

        if 200 <= response.status_code < 300:
            status = MessageRecipient.Status.SENT
        elif code == "NotBaleUser":
            status = MessageRecipient.Status.NOT_BALE_USER
        elif response.status_code == 429 or code == "RateLimitExceeded":
            status = MessageRecipient.Status.RATE_LIMITED
        elif response.status_code == 402 or code == "PaymentRequired":
            status = MessageRecipient.Status.PAYMENT_REQUIRED
        elif code in {"InvalidPhone", "InvalidPhoneNumber"}:
            status = MessageRecipient.Status.INVALID_PHONE
        else:
            status = MessageRecipient.Status.FAILED

        return {"status": status, "http_status": response.status_code, "api_code": code, "api_message": message_text, "raw_response": raw}


def write_report(batch: MessageBatch, report_path: Path) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "گزارش ارسال"
    ws.append(["ردیف اکسل", "نام", "نام خانوادگی", "نام کامل", "شماره خام", "شماره استاندارد", "وضعیت", "HTTP", "کد API", "پیام API", "خطا", "متن نهایی"])
    for r in batch.recipients.all().order_by("row_number", "id"):
        ws.append([r.row_number, r.first_name, r.last_name, r.full_name, r.raw_phone, r.normalized_phone, r.status, r.http_status, r.api_code, r.api_message, r.error_message, r.final_text])
    wb.save(report_path)


def _refresh_totals(batch: MessageBatch) -> None:
    recipients = batch.recipients.all()
    batch.total_rows = recipients.count()
    batch.total_sent = recipients.filter(status=MessageRecipient.Status.SENT).count()
    batch.total_failed = recipients.filter(status=MessageRecipient.Status.FAILED).count()
    batch.total_invalid = recipients.filter(status=MessageRecipient.Status.INVALID_PHONE).count()
    batch.total_duplicate = recipients.filter(status=MessageRecipient.Status.DUPLICATE).count()
    batch.total_not_bale_user = recipients.filter(status=MessageRecipient.Status.NOT_BALE_USER).count()
    batch.total_rate_limited = recipients.filter(status=MessageRecipient.Status.RATE_LIMITED).count()
    batch.total_payment_required = recipients.filter(status=MessageRecipient.Status.PAYMENT_REQUIRED).count()
    batch.total_config_error = recipients.filter(status=MessageRecipient.Status.CONFIG_ERROR).count()
    batch.finished_at = timezone.now()
    batch.save()


def run_excel_batch(*, file_path: str, message_template: str, dry_run: bool = True, limit: int | None = None, sleep_seconds: float | None = None, sheet_name: str | None = None, button_text: str | None = None, button_url: str | None = None, skip_duplicates: bool = True, report_path: str | None = None) -> MessageBatch:
    recipients = read_excel_recipients(file_path, sheet_name=sheet_name)
    if limit:
        recipients = recipients[:limit]

    batch = MessageBatch.objects.create(
        source_file_name=Path(file_path).name,
        message_template=message_template,
        button_text=button_text or "",
        button_url=button_url or "",
        dry_run=dry_run,
        limit=limit,
    )

    seen: set[str] = set()
    client = BaleSafirClient()
    delay = settings.BALE_DEFAULT_SLEEP_SECONDS if sleep_seconds is None else sleep_seconds

    for item in recipients:
        final_text = render_message(message_template, item)
        obj = MessageRecipient.objects.create(
            batch=batch,
            row_number=item.row_number,
            first_name=item.first_name,
            last_name=item.last_name,
            full_name=item.full_name,
            raw_phone=item.raw_phone,
            normalized_phone=item.normalized_phone or "",
            final_text=final_text,
        )

        if not item.normalized_phone:
            obj.status = MessageRecipient.Status.INVALID_PHONE
            obj.error_message = "شماره موبایل معتبر نیست."
            obj.save()
            continue

        if skip_duplicates and item.normalized_phone in seen:
            obj.status = MessageRecipient.Status.DUPLICATE
            obj.error_message = "این شماره در همین فایل تکراری است."
            obj.save()
            continue
        seen.add(item.normalized_phone)

        request_id = f"bale-{batch.id}-{item.row_number}-{uuid4().hex[:8]}"
        obj.request_id = request_id

        if dry_run:
            obj.status = MessageRecipient.Status.DRY_RUN
            obj.api_message = "dry-run: ارسال واقعی انجام نشد."
            obj.save()
            continue

        result = client.send(
            phone_number=item.normalized_phone,
            text=final_text,
            request_id=request_id,
            button_text=button_text,
            button_url=button_url,
        )
        obj.status = result.get("status", MessageRecipient.Status.FAILED)
        obj.http_status = result.get("http_status")
        obj.api_code = result.get("api_code", "")
        obj.api_message = result.get("api_message", "")
        obj.raw_response = result.get("raw_response", "")
        obj.error_message = result.get("error", "")
        obj.sent_at = timezone.now() if obj.status == MessageRecipient.Status.SENT else None
        obj.save()
        if delay:
            sleep(delay)

    if not report_path:
        reports_dir = Path(settings.BASE_DIR) / "reports"
        report_path = str(reports_dir / f"bale_report_batch_{batch.id}.xlsx")
    batch.report_path = report_path
    batch.save(update_fields=["report_path"])
    write_report(batch, Path(report_path))
    _refresh_totals(batch)
    return batch
