from __future__ import annotations

import json
import re
import string
from dataclasses import dataclass, field
from hashlib import sha256
from itertools import chain
from pathlib import Path
from time import monotonic, sleep
from urllib.parse import urlparse
from zipfile import BadZipFile

import requests
from django.conf import settings
from django.db import OperationalError, ProgrammingError
from django.db.models import Count
from django.utils import timezone
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from .models import MessageBatch, MessageRecipient, SafirConfiguration

PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
PHONE_HEADERS = {
    "موبایل",
    "شماره",
    "شماره موبایل",
    "شماره همراه",
    "mobile",
    "phone",
    "phone_number",
}
FIRST_NAME_HEADERS = {"نام", "اسم", "first_name", "firstname"}
LAST_NAME_HEADERS = {"نام خانوادگی", "فامیلی", "last_name", "lastname"}
BUILTIN_TOKENS = frozenset(
    {"first_name", "last_name", "full_name", "phone", "row_number"}
)
TOKEN_RE = re.compile(r"{{\s*([^{}]+?)\s*}}")
LEGACY_TOKEN_RE = re.compile(
    r"(?<!{){(first_name|last_name|full_name|phone|row_number)}(?!})"
)
ERROR_CODES = {
    "2": "InternalServerError",
    "3": "RateLimitExceeded",
    "4": "InvalidInput",
    "8": "InvalidPhone",
    "17": "NotBaleUser",
    "20": "PaymentRequired",
    "21": "MaximumContactLimitReached",
}


@dataclass
class ExcelRecipient:
    row_number: int
    first_name: str
    last_name: str
    raw_phone: str
    normalized_phone: str | None
    values: dict[str, str] = field(default_factory=dict)

    @property
    def full_name(self) -> str:
        return " ".join(x for x in [self.first_name, self.last_name] if x).strip()

    @property
    def context(self) -> dict[str, str]:
        return {
            **self.values,
            "first_name": self.first_name,
            "last_name": self.last_name,
            "full_name": self.full_name,
            "phone": self.normalized_phone or self.raw_phone,
            "row_number": str(self.row_number),
        }


def normalize_iran_mobile(value: object) -> str | None:
    if value is None:
        return None
    number = str(value).strip().translate(PERSIAN_DIGITS)
    number = re.sub(r"[^0-9+]", "", number)
    if number.startswith("+98"):
        number = "98" + number[3:]
    elif number.startswith("0098"):
        number = "98" + number[4:]
    elif number.startswith("09") and len(number) == 11:
        number = "98" + number[1:]
    elif number.startswith("9") and len(number) == 10:
        number = "98" + number
    return number if re.fullmatch(r"989\d{9}", number) else None


def _clean_header(value: object) -> str:
    return str(value or "").strip().replace("ي", "ی").replace("ك", "ک")


def _safe_token(value: str, index: int) -> str:
    token = re.sub(r"[^\w]+", "_", _clean_header(value), flags=re.UNICODE).strip("_")
    if not token or token[0].isdigit():
        token = f"column_{get_column_letter(index + 1).lower()}"
    return token


def _load_workbook(file_path: str | Path):
    try:
        return load_workbook(file_path, read_only=True, data_only=True)
    except (InvalidFileException, BadZipFile, OSError) as exc:
        raise ValueError(
            "فایل اکسل قابل خواندن نیست. یک فایل سالم xlsx یا xlsm انتخاب کن."
        ) from exc


def inspect_excel(
    file_path: str | Path, sheet_name: str | None = None, header_row: int = 1
) -> dict:
    if header_row < 1:
        raise ValueError("شماره ردیف عنوان‌ها باید حداقل ۱ باشد.")
    workbook = _load_workbook(file_path)
    try:
        worksheet = workbook[sheet_name] if sheet_name else workbook.active
    except KeyError as exc:
        available = "، ".join(workbook.sheetnames)
        raise ValueError(
            f"شیت «{sheet_name}» پیدا نشد. شیت‌های موجود: {available}"
        ) from exc
    row = next(
        worksheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True),
        None,
    )
    if row is None or not any(str(value or "").strip() for value in row):
        raise ValueError(f"ردیف عنوان‌ها ({header_row}) خالی است.")
    columns = []
    used_tokens: set[str] = set()
    for index, value in enumerate(row):
        header = _clean_header(value)
        if not header:
            continue
        token = _safe_token(header, index)
        base = token
        suffix = 2
        while token in used_tokens or token in BUILTIN_TOKENS:
            token = f"{base}_{suffix}"
            suffix += 1
        used_tokens.add(token)
        columns.append(
            {
                "index": index,
                "key": get_column_letter(index + 1),
                "header": header,
                "token": token,
            }
        )
    return {
        "sheet_name": worksheet.title,
        "sheets": workbook.sheetnames,
        "header_row": header_row,
        "columns": columns,
    }


def _resolve_column(
    headers: list[str],
    selector: str | None,
    candidates: set[str] | None = None,
    *,
    required=False,
    title="ستون",
) -> int | None:
    cleaned = _clean_header(selector)
    if cleaned:
        if re.fullmatch(r"[A-Za-z]+", cleaned):
            number = 0
            for char in cleaned.upper():
                number = number * 26 + ord(char) - 64
            index = number - 1
            if 0 <= index < len(headers):
                return index
        if cleaned.isdigit() and 1 <= int(cleaned) <= len(headers):
            return int(cleaned) - 1
        matches = [i for i, header in enumerate(headers) if header == cleaned]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raise ValueError(
                f"عنوان «{cleaned}» تکراری است؛ برای {title} حرف ستون مثل A یا B را وارد کن."
            )
        raise ValueError(f"{title} «{cleaned}» در ردیف عنوان‌ها پیدا نشد.")
    if candidates:
        for index, header in enumerate(headers):
            if header.lower() in candidates or header in candidates:
                return index
    if required:
        visible = "، ".join(header for header in headers if header) or "بدون عنوان"
        raise ValueError(
            f"{title} پیدا نشد. نام یا حرف ستون را مشخص کن. ستون‌های موجود: {visible}"
        )
    return None


def _cell_text(row: tuple, index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    value = row[index]
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def extract_tokens(template: str) -> set[str]:
    normalized = LEGACY_TOKEN_RE.sub(
        lambda match: "{{" + match.group(1) + "}}", template or ""
    )
    if normalized.count("{{") != normalized.count("}}"):
        raise ValueError("آکولادهای متغیرها کامل نیستند. قالب درست: {{variable}}")
    tokens = {match.group(1).strip() for match in TOKEN_RE.finditer(normalized)}
    try:
        parsed = string.Formatter().parse(template or "")
        tokens.update(
            field_name
            for _literal, field_name, _format, _conversion in parsed
            if field_name
        )
    except ValueError as exc:
        raise ValueError("قالب پیام نامعتبر است؛ آکولادها را بررسی کن.") from exc
    if any(not token or any(ch.isspace() for ch in token) for token in tokens):
        raise ValueError(
            "نام متغیر نباید خالی یا دارای فاصله باشد؛ مانند {{first_name}}."
        )
    return tokens


def validate_message_template(
    template: str, available_tokens: set[str] | None = None
) -> None:
    tokens = extract_tokens(template)
    allowed = (
        set(BUILTIN_TOKENS)
        if available_tokens is None
        else set(available_tokens) | set(BUILTIN_TOKENS)
    )
    unknown = tokens - allowed
    if unknown:
        rendered = "، ".join(f"{{{{{name}}}}}" for name in sorted(unknown))
        raise ValueError(
            f"placeholderهای نامعتبر یا متغیرهای تعریف‌نشده در متن: {rendered}"
        )


def render_template(template: str, context: dict[str, object]) -> str:
    normalized = LEGACY_TOKEN_RE.sub(
        lambda match: "{{" + match.group(1) + "}}", template or ""
    )
    validate_message_template(normalized, set(context))
    return TOKEN_RE.sub(
        lambda match: str(context.get(match.group(1).strip(), "")), normalized
    )


def read_excel_recipients(
    file_path: str | Path,
    sheet_name: str | None = None,
    *,
    header_row: int = 1,
    phone_column: str | None = None,
    first_name_column: str | None = None,
    last_name_column: str | None = None,
    variable_mapping: dict[str, str] | None = None,
) -> list[ExcelRecipient]:
    workbook = _load_workbook(file_path)
    try:
        worksheet = workbook[sheet_name] if sheet_name else workbook.active
    except KeyError as exc:
        available = "، ".join(workbook.sheetnames)
        raise ValueError(
            f"شیت «{sheet_name}» در فایل اکسل پیدا نشد. شیت‌های موجود: {available}"
        ) from exc
    rows = worksheet.iter_rows(values_only=True)
    for _ in range(header_row - 1):
        next(rows, None)
    first_row = next(rows, None)
    if first_row is None or not any(str(value or "").strip() for value in first_row):
        raise ValueError(
            "ردیف اول فایل اکسل باید header یا شماره موبایل باشد، اما خالی است."
        )
    headers = [_clean_header(value) for value in first_row]
    phone_idx = _resolve_column(
        headers, phone_column, PHONE_HEADERS, title="ستون موبایل"
    )
    headerless = False
    if phone_idx is None and header_row == 1 and not phone_column:
        nonempty = [i for i, value in enumerate(first_row) if str(value or "").strip()]
        if len(nonempty) == 1 and normalize_iran_mobile(first_row[nonempty[0]]):
            phone_idx = nonempty[0]
            headerless = True
            data_rows = chain([(1, first_row)], enumerate(rows, start=2))
        else:
            visible = "، ".join(header for header in headers if header) or "بدون عنوان"
            raise ValueError(
                f"ستون «موبایل» در فایل اکسل پیدا نشد. ستون‌های موجود: {visible}"
            )
    else:
        data_rows = enumerate(rows, start=header_row + 1)
    if phone_idx is None:
        try:
            phone_idx = _resolve_column(
                headers,
                phone_column,
                PHONE_HEADERS,
                required=True,
                title="ستون موبایل",
            )
        except ValueError as exc:
            if not phone_column:
                visible = (
                    "، ".join(header for header in headers if header) or "بدون عنوان"
                )
                raise ValueError(
                    f"ستون «موبایل» در فایل اکسل پیدا نشد. ستون‌های موجود: {visible}"
                ) from exc
            raise
    first_idx = (
        None
        if headerless
        else _resolve_column(
            headers, first_name_column, FIRST_NAME_HEADERS, title="ستون نام"
        )
    )
    last_idx = (
        None
        if headerless
        else _resolve_column(
            headers, last_name_column, LAST_NAME_HEADERS, title="ستون نام خانوادگی"
        )
    )
    mapping = variable_mapping or {}
    resolved_mapping = {
        token: _resolve_column(
            headers, selector, required=True, title=f"ستون متغیر {token}"
        )
        for token, selector in mapping.items()
    }
    auto_columns: list[tuple[str, int]] = []
    used_tokens = set(BUILTIN_TOKENS) | set(mapping)
    for index, header in enumerate(headers):
        if not header:
            continue
        token = _safe_token(header, index)
        base = token
        suffix = 2
        while token in used_tokens:
            token = f"{base}_{suffix}"
            suffix += 1
        used_tokens.add(token)
        auto_columns.append((token, index))
    result = []
    for row_number, row in data_rows:
        if not any(str(value or "").strip() for value in row):
            continue
        raw_phone = _cell_text(row, phone_idx)
        first_name = _cell_text(row, first_idx)
        last_name = _cell_text(row, last_idx)
        values = {token: _cell_text(row, index) for token, index in auto_columns}
        values.update(
            {token: _cell_text(row, index) for token, index in resolved_mapping.items()}
        )
        result.append(
            ExcelRecipient(
                row_number,
                first_name,
                last_name,
                raw_phone,
                normalize_iran_mobile(raw_phone),
                values,
            )
        )
    return result


def render_message(
    template: str, recipient: ExcelRecipient, global_variables: dict | None = None
) -> str:
    return render_template(template, {**(global_variables or {}), **recipient.context})


def _apply_data_range(
    recipients: list[ExcelRecipient], *, range_start=None, range_end=None
) -> list[ExcelRecipient]:
    return recipients[max((range_start or 1) - 1, 0) : range_end if range_end else None]


def _recipient_preview_status(item: ExcelRecipient, seen: set[str]) -> tuple[str, str]:
    if not item.normalized_phone:
        return "invalid_phone", "شماره نامعتبر"
    if item.normalized_phone in seen:
        return "duplicate", "تکراری در فایل"
    seen.add(item.normalized_phone)
    return "ok", "آماده ارسال"


def _build_message_data(
    batch: MessageBatch, item: ExcelRecipient, attachment_file_id: str = ""
) -> tuple[str, dict]:
    context = {**(batch.global_variables or {}), **item.context}
    if batch.message_type == MessageBatch.MessageType.TEMPLATE:
        fields = {
            key: render_template(str(value), context)
            for key, value in (batch.template_fields or {}).items()
        }
        data = {
            "template_message": {
                "template_id": batch.template_id,
                "text_fields": fields,
            }
        }
        final_text = json.dumps(
            {"template_id": batch.template_id, "text_fields": fields},
            ensure_ascii=False,
        )
    elif batch.message_type == MessageBatch.MessageType.OTP:
        otp = (
            render_template(batch.otp_template, context)
            .translate(PERSIAN_DIGITS)
            .strip()
        )
        if not otp.isdigit():
            raise ValueError("مقدار OTP هر ردیف باید فقط عدد باشد.")
        data = {"otp_message": {"otp": otp}}
        final_text = f"OTP: {otp}"
    else:
        final_text = render_template(batch.message_template, context)
        message: dict = {}
        if final_text:
            message["text"] = final_text
        if attachment_file_id:
            message["file_id"] = attachment_file_id
        if batch.copy_text_template:
            message["copy_text"] = render_template(batch.copy_text_template, context)
        buttons_by_row: dict[int, list[dict]] = {}
        buttons = list(batch.buttons or [])
        if not buttons and batch.button_text and batch.button_url:
            buttons = [
                {
                    "text": batch.button_text,
                    "type": "url",
                    "value": batch.button_url,
                    "row": 1,
                }
            ]
        for button in buttons:
            text = render_template(str(button.get("text", "")), context)
            value = render_template(str(button.get("value", "")), context)
            kind = button.get("type", "url")
            if kind in {"url", "web_app"} and urlparse(value).scheme not in {
                "http",
                "https",
            }:
                raise ValueError(f"آدرس دکمه «{text}» معتبر نیست.")
            payload = {"text": text}
            payload["web_app" if kind == "web_app" else kind] = (
                {"url": value} if kind == "web_app" else value
            )
            buttons_by_row.setdefault(max(int(button.get("row", 1)), 1), []).append(
                payload
            )
        if buttons_by_row:
            message["reply_markup"] = {
                "inline_keyboard": [
                    buttons_by_row[row] for row in sorted(buttons_by_row)
                ]
            }
        if not message:
            raise ValueError("پیام متنی باید حداقل متن یا فایل داشته باشد.")
        data = {"message": message}
    if batch.is_secure and batch.message_type != MessageBatch.MessageType.OTP:
        data["is_secure"] = True
    return final_text, data


def build_excel_preview(
    file_path: str | Path,
    *,
    sheet_name: str | None = None,
    message_template: str,
    limit: int = 10,
    range_start=None,
    range_end=None,
    header_row: int = 1,
    phone_column: str | None = None,
    first_name_column: str | None = None,
    last_name_column: str | None = None,
    variable_mapping: dict | None = None,
    global_variables: dict | None = None,
    batch_options: dict | None = None,
) -> dict:
    recipients = read_excel_recipients(
        file_path,
        sheet_name,
        header_row=header_row,
        phone_column=phone_column,
        first_name_column=first_name_column,
        last_name_column=last_name_column,
        variable_mapping=variable_mapping,
    )
    total_file_rows = len(recipients)
    recipients = _apply_data_range(
        recipients, range_start=range_start, range_end=range_end
    )
    seen: set[str] = set()
    rows = []
    counts = {"ok": 0, "invalid_phone": 0, "duplicate": 0}
    for item in recipients:
        status, label = _recipient_preview_status(item, seen)
        counts[status] += 1
        if len(rows) < limit:
            if batch_options:
                preview_batch = MessageBatch(
                    **batch_options, global_variables=global_variables or {}
                )
                file_id = "preview-file-id" if preview_batch.attachment_path else ""
                final_text, _ = _build_message_data(preview_batch, item, file_id)
            else:
                final_text = render_message(message_template, item, global_variables)
            rows.append(
                {
                    "row_number": item.row_number,
                    "first_name": item.first_name,
                    "last_name": item.last_name,
                    "full_name": item.full_name,
                    "raw_phone": item.raw_phone,
                    "normalized_phone": item.normalized_phone or "",
                    "status": status,
                    "status_label": label,
                    "final_text": final_text,
                    "variables": item.context,
                }
            )
    inspection = None
    if not (header_row == 1 and recipients and recipients[0].row_number == 1):
        inspection = inspect_excel(file_path, sheet_name, header_row)
    return {
        "total_file_rows": total_file_rows,
        "total_rows": len(recipients),
        "range_start": range_start,
        "range_end": range_end,
        "valid_rows": counts["ok"],
        "invalid_rows": counts["invalid_phone"],
        "duplicate_rows": counts["duplicate"],
        "rows": rows,
        "shown_rows": len(rows),
        "is_sendable": counts["ok"] > 0,
        "inspection": inspection,
    }


def effective_configuration() -> dict:
    fallback = {
        "api_access_key": settings.BALE_API_ACCESS_KEY,
        "bot_id": settings.BALE_BOT_ID,
        "send_url": settings.BALE_SEND_URL,
        "upload_url": getattr(
            settings, "BALE_UPLOAD_URL", "https://safir.bale.ai/api/v3/upload_file"
        ),
        "request_timeout": settings.BALE_REQUEST_TIMEOUT,
        "default_sleep_seconds": settings.BALE_DEFAULT_SLEEP_SECONDS,
        "max_retries": getattr(settings, "BALE_MAX_RETRIES", 2),
        "message_price_rial": settings.BALE_MESSAGE_PRICE_RIAL,
        "global_variables": {},
    }
    try:
        config = SafirConfiguration.load()
    except (OperationalError, ProgrammingError):
        return fallback
    return {
        "api_access_key": config.api_access_key or fallback["api_access_key"],
        "bot_id": config.bot_id or fallback["bot_id"],
        "send_url": config.send_url or fallback["send_url"],
        "upload_url": config.upload_url or fallback["upload_url"],
        "request_timeout": config.request_timeout or fallback["request_timeout"],
        "default_sleep_seconds": config.default_sleep_seconds,
        "max_retries": config.max_retries,
        "message_price_rial": config.message_price_rial
        or fallback["message_price_rial"],
        "global_variables": config.global_variables or {},
    }


class BaleSafirClient:
    def __init__(self, *, max_retries: int | None = None) -> None:
        config = effective_configuration()
        self.url = config["send_url"]
        self.upload_url = config["upload_url"]
        self.api_key = config["api_access_key"]
        self.bot_id = config["bot_id"]
        self.timeout = config["request_timeout"]
        self.max_retries = config["max_retries"] if max_retries is None else max_retries

    @staticmethod
    def _error(data: dict) -> tuple[str, str]:
        errors = data.get("error_data")
        if isinstance(errors, list) and errors:
            error = errors[0] or {}
        elif isinstance(errors, dict):
            error = errors
        else:
            error = data
        raw_code = error.get("code") or error.get("error") or ""
        code = ERROR_CODES.get(str(raw_code), str(raw_code))
        description = str(error.get("description") or error.get("message") or "")
        return code, description

    @staticmethod
    def _status(http_status: int, code: str, has_message_id: bool) -> str:
        if 200 <= http_status < 300 and not code:
            return MessageRecipient.Status.SENT
        if code in {"NotBaleUser", "17"}:
            return MessageRecipient.Status.NOT_BALE_USER
        if http_status == 429 or code in {"RateLimitExceeded", "3"}:
            return MessageRecipient.Status.RATE_LIMITED
        if http_status == 402 or code in {"PaymentRequired", "20"}:
            return MessageRecipient.Status.PAYMENT_REQUIRED
        if code in {"InvalidPhone", "InvalidPhoneNumber", "8"}:
            return MessageRecipient.Status.INVALID_PHONE
        return MessageRecipient.Status.FAILED

    def send_data(
        self, *, phone_number: str, request_id: str, message_data: dict
    ) -> dict:
        if not self.api_key or not self.bot_id:
            return {
                "status": MessageRecipient.Status.CONFIG_ERROR,
                "error": "کلید API یا شناسه بازو تنظیم نشده است.",
                "retry_count": 0,
            }
        payload = {
            "request_id": request_id,
            "bot_id": int(self.bot_id),
            "phone_number": phone_number,
            "message_data": message_data,
        }
        headers = {"api-access-key": self.api_key, "Content-Type": "application/json"}
        last_result = None
        for attempt in range(self.max_retries + 1):
            try:
                response = requests.post(
                    self.url, headers=headers, json=payload, timeout=self.timeout
                )
                try:
                    data = response.json()
                except ValueError:
                    data = {"description": response.text}
                code, description = self._error(data)
                message_id = str(data.get("message_id") or "")
                status = self._status(response.status_code, code, bool(message_id))
                last_result = {
                    "status": status,
                    "http_status": response.status_code,
                    "api_code": code,
                    "api_message": description,
                    "message_id": message_id,
                    "raw_response": json.dumps(data, ensure_ascii=False),
                    "retry_count": attempt,
                }
                transient = response.status_code in {
                    429,
                    500,
                    502,
                    503,
                    504,
                } or code in {
                    "RateLimitExceeded",
                    "InternalServerError",
                    "2",
                    "3",
                }
                if not transient or attempt >= self.max_retries:
                    return last_result
            except requests.RequestException as exc:
                last_result = {
                    "status": MessageRecipient.Status.FAILED,
                    "error": str(exc),
                    "retry_count": attempt,
                }
                if attempt >= self.max_retries:
                    return last_result
            sleep(min(2**attempt, 8))
        return last_result or {
            "status": MessageRecipient.Status.FAILED,
            "error": "پاسخی از سفیر دریافت نشد.",
            "retry_count": 0,
        }

    def send(
        self,
        *,
        phone_number: str,
        text: str,
        request_id: str,
        button_text: str | None = None,
        button_url: str | None = None,
    ) -> dict:
        message: dict = {"text": text}
        if button_text and button_url:
            message["reply_markup"] = {
                "inline_keyboard": [[{"text": button_text, "url": button_url}]]
            }
        return self.send_data(
            phone_number=phone_number,
            request_id=request_id,
            message_data={"message": message},
        )

    def upload_file(self, file_path: str | Path) -> dict:
        if not self.api_key:
            return {"ok": False, "error": "کلید API سفیر تنظیم نشده است."}
        try:
            with open(file_path, "rb") as source:
                response = requests.post(
                    self.upload_url,
                    headers={"api-access-key": self.api_key},
                    files={"file": (Path(file_path).name, source)},
                    timeout=max(self.timeout, 60),
                )
            data = response.json()
        except (OSError, requests.RequestException, ValueError) as exc:
            return {"ok": False, "error": str(exc)}
        file_id = str(data.get("file_id") or "")
        if 200 <= response.status_code < 300 and file_id:
            return {"ok": True, "file_id": file_id}
        code, description = self._error(data)
        return {
            "ok": False,
            "error": description or code or f"HTTP {response.status_code}",
        }


def write_report(batch: MessageBatch, report_path: Path) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "گزارش ارسال"
    worksheet.freeze_panes = "A2"
    worksheet.append(
        [
            "ردیف اکسل",
            "نام کامل",
            "شماره خام",
            "شماره استاندارد",
            "وضعیت",
            "HTTP",
            "کد API",
            "پیام API",
            "خطا",
            "شناسه پیام",
            "تلاش مجدد",
            "متغیرها",
            "متن نهایی",
        ]
    )
    for item in (
        batch.recipients.all().order_by("row_number", "id").iterator(chunk_size=1000)
    ):
        worksheet.append(
            [
                item.row_number,
                item.full_name,
                item.raw_phone,
                item.normalized_phone,
                item.get_status_display(),
                item.http_status,
                item.api_code,
                item.api_message,
                item.error_message,
                item.message_id,
                item.retry_count,
                json.dumps(item.variables, ensure_ascii=False),
                item.final_text,
            ]
        )
    workbook.save(report_path)


def _refresh_totals(batch: MessageBatch, *, mark_finished: bool = False) -> None:
    counts = dict(
        batch.recipients.order_by()
        .values("status")
        .annotate(c=Count("id"))
        .values_list("status", "c")
    )
    fields = {
        "total_rows": sum(counts.values()),
        "total_sent": counts.get(MessageRecipient.Status.SENT, 0),
        "total_failed": counts.get(MessageRecipient.Status.FAILED, 0),
        "total_invalid": counts.get(MessageRecipient.Status.INVALID_PHONE, 0),
        "total_duplicate": counts.get(MessageRecipient.Status.DUPLICATE, 0),
        "total_not_bale_user": counts.get(MessageRecipient.Status.NOT_BALE_USER, 0),
        "total_rate_limited": counts.get(MessageRecipient.Status.RATE_LIMITED, 0),
        "total_payment_required": counts.get(
            MessageRecipient.Status.PAYMENT_REQUIRED, 0
        ),
        "total_config_error": counts.get(MessageRecipient.Status.CONFIG_ERROR, 0),
    }
    for name, value in fields.items():
        setattr(batch, name, value)
    update_fields = list(fields)
    if mark_finished:
        batch.finished_at = timezone.now()
        update_fields.append("finished_at")
    batch.save(update_fields=update_fields)


def _is_cancel_requested(batch: MessageBatch) -> bool:
    batch.refresh_from_db(fields=["cancel_requested"])
    return batch.cancel_requested


def _sleep_until_cancel_or_timeout(batch: MessageBatch, seconds: float) -> bool:
    if seconds <= 0:
        return _is_cancel_requested(batch)
    remaining = seconds
    last_check = 0.0
    while remaining > 0:
        now = monotonic()
        if now - last_check >= 1:
            if _is_cancel_requested(batch):
                return True
            last_check = now
        chunk = min(0.2, remaining)
        sleep(chunk)
        remaining -= chunk
    return _is_cancel_requested(batch)


def cancel_batch_immediately(batch: MessageBatch) -> MessageBatch:
    batch.cancel_requested = True
    batch.status = MessageBatch.Status.CANCELLED
    batch.finished_at = timezone.now()
    batch.error_message = (
        "پردازش با درخواست کاربر متوقف شد؛ ممکن است درخواست در حال اجرا ثبت شده باشد."
    )
    if not batch.report_path:
        batch.report_path = str(
            Path(settings.BASE_DIR) / "reports" / f"bale_report_batch_{batch.id}.xlsx"
        )
    batch.save(
        update_fields=[
            "cancel_requested",
            "status",
            "finished_at",
            "error_message",
            "report_path",
        ]
    )
    _refresh_totals(batch)
    write_report(batch, Path(batch.report_path))
    batch.refresh_from_db()
    return batch


def _stable_request_id(batch: MessageBatch, item: ExcelRecipient) -> str:
    digest = sha256(
        f"{batch.id}:{item.row_number}:{item.normalized_phone}".encode()
    ).hexdigest()[:20]
    return f"bale-{batch.id}-{item.row_number}-{digest}"


def process_excel_batch(
    *,
    batch: MessageBatch,
    file_path: str,
    sleep_seconds=None,
    sheet_name=None,
    skip_duplicates=None,
    report_path=None,
    range_start=None,
    range_end=None,
) -> MessageBatch:
    batch.refresh_from_db(fields=["cancel_requested"])
    if batch.cancel_requested:
        batch.status = MessageBatch.Status.CANCELLED
        batch.finished_at = timezone.now()
        batch.save(update_fields=["status", "finished_at"])
        return batch
    batch.status = MessageBatch.Status.RUNNING
    batch.error_message = ""
    batch.finished_at = None
    batch.save(update_fields=["status", "error_message", "finished_at"])
    try:
        recipients = read_excel_recipients(
            file_path,
            sheet_name or batch.sheet_name or None,
            header_row=batch.header_row,
            phone_column=batch.phone_column or None,
            first_name_column=batch.first_name_column or None,
            last_name_column=batch.last_name_column or None,
            variable_mapping=batch.variable_mapping,
        )
        recipients = _apply_data_range(
            recipients,
            range_start=range_start or batch.range_start,
            range_end=range_end or batch.range_end,
        )
        if batch.limit:
            recipients = recipients[: batch.limit]
        client = BaleSafirClient(max_retries=batch.max_retries)
        attachment_file_id = batch.attachment_file_id
        if batch.attachment_path and not attachment_file_id and not batch.dry_run:
            uploaded = client.upload_file(batch.attachment_path)
            if not uploaded["ok"]:
                raise ValueError(
                    f"بارگذاری فایل در سفیر ناموفق بود: {uploaded['error']}"
                )
            attachment_file_id = uploaded["file_id"]
            batch.attachment_file_id = attachment_file_id
            batch.save(update_fields=["attachment_file_id"])
        seen: set[str] = set()
        cancelled = False
        should_skip_duplicates = (
            batch.skip_duplicates if skip_duplicates is None else skip_duplicates
        )
        delay = batch.sleep_seconds if sleep_seconds is None else sleep_seconds
        for item in recipients:
            if _is_cancel_requested(batch):
                cancelled = True
                break
            try:
                final_text, message_data = _build_message_data(
                    batch,
                    item,
                    attachment_file_id
                    or ("dry-run-file" if batch.attachment_path else ""),
                )
            except ValueError as exc:
                final_text, message_data = "", {}
                render_error = str(exc)
            else:
                render_error = ""
            obj = MessageRecipient(
                batch=batch,
                row_number=item.row_number,
                first_name=item.first_name,
                last_name=item.last_name,
                full_name=item.full_name,
                raw_phone=item.raw_phone,
                normalized_phone=item.normalized_phone or "",
                final_text=final_text,
                variables=item.context,
                request_id=_stable_request_id(batch, item)
                if item.normalized_phone
                else "",
            )
            if render_error:
                obj.status = MessageRecipient.Status.FAILED
                obj.error_message = render_error
                obj.save()
                continue
            if not item.normalized_phone:
                obj.status = MessageRecipient.Status.INVALID_PHONE
                obj.error_message = "شماره موبایل معتبر نیست."
                obj.save()
                continue
            if should_skip_duplicates and item.normalized_phone in seen:
                obj.status = MessageRecipient.Status.DUPLICATE
                obj.error_message = "این شماره در همین فایل تکراری است."
                obj.save()
                continue
            seen.add(item.normalized_phone)
            if batch.dry_run:
                obj.status = MessageRecipient.Status.DRY_RUN
                obj.api_message = "dry-run: ارسال واقعی انجام نشد."
                obj.save()
                continue
            obj.save()
            is_legacy_text = (
                batch.message_type == MessageBatch.MessageType.TEXT
                and not batch.attachment_path
                and not batch.copy_text_template
                and not batch.buttons
                and not batch.is_secure
            )
            if is_legacy_text:
                result = client.send(
                    phone_number=item.normalized_phone,
                    text=final_text,
                    request_id=obj.request_id,
                    button_text=batch.button_text or None,
                    button_url=batch.button_url or None,
                )
            else:
                result = client.send_data(
                    phone_number=item.normalized_phone,
                    request_id=obj.request_id,
                    message_data=message_data,
                )
            for result_name, target_name in (
                ("status", "status"),
                ("http_status", "http_status"),
                ("api_code", "api_code"),
                ("api_message", "api_message"),
                ("raw_response", "raw_response"),
                ("error", "error_message"),
                ("message_id", "message_id"),
                ("retry_count", "retry_count"),
            ):
                if result_name in result:
                    setattr(obj, target_name, result[result_name])
            obj.sent_at = (
                timezone.now() if obj.status == MessageRecipient.Status.SENT else None
            )
            obj.save()
            if _sleep_until_cancel_or_timeout(batch, delay or 0):
                cancelled = True
                break
        if not report_path:
            report_path = str(
                Path(settings.BASE_DIR)
                / "reports"
                / f"bale_report_batch_{batch.id}.xlsx"
            )
        batch.report_path = report_path
        batch.save(update_fields=["report_path"])
        write_report(batch, Path(report_path))
        _refresh_totals(batch)
        batch.finished_at = timezone.now()
        if cancelled or batch.cancel_requested:
            batch.status = MessageBatch.Status.CANCELLED
            batch.error_message = "پردازش با درخواست کاربر متوقف شد."
        elif not batch.dry_run and (batch.total_failed or batch.total_config_error):
            batch.status = MessageBatch.Status.FAILED
            batch.error_message = "پردازش کامل شد، اما بعضی ردیف‌ها ناموفق بودند."
        else:
            batch.status = MessageBatch.Status.FINISHED
        batch.save(update_fields=["status", "error_message", "finished_at"])
    except Exception as exc:
        _refresh_totals(batch)
        batch.status = MessageBatch.Status.FAILED
        batch.error_message = f"خطا در پردازش: {exc}"
        batch.finished_at = timezone.now()
        batch.save(update_fields=["status", "error_message", "finished_at"])
        raise
    return batch


def send_single_recipient_test(
    *,
    first_name: str,
    last_name: str,
    phone: str,
    message_template: str,
    button_text=None,
    button_url=None,
) -> MessageBatch:
    item = ExcelRecipient(
        1,
        first_name.strip(),
        last_name.strip(),
        phone.strip(),
        normalize_iran_mobile(phone),
    )
    final_text = render_message(message_template, item)
    batch = MessageBatch.objects.create(
        source_file_name="ارسال تست تک‌نفره",
        message_template=message_template,
        button_text=button_text or "",
        button_url=button_url or "",
        dry_run=False,
        limit=1,
        status=MessageBatch.Status.RUNNING,
    )
    recipient = MessageRecipient.objects.create(
        batch=batch,
        row_number=1,
        first_name=item.first_name,
        last_name=item.last_name,
        full_name=item.full_name,
        raw_phone=item.raw_phone,
        normalized_phone=item.normalized_phone or "",
        final_text=final_text,
        variables=item.context,
    )
    if not item.normalized_phone:
        recipient.status = MessageRecipient.Status.INVALID_PHONE
        recipient.error_message = "شماره موبایل معتبر نیست."
    else:
        recipient.request_id = _stable_request_id(batch, item)
        result = BaleSafirClient().send(
            phone_number=item.normalized_phone,
            text=final_text,
            request_id=recipient.request_id,
            button_text=button_text,
            button_url=button_url,
        )
        recipient.status = result.get("status", MessageRecipient.Status.FAILED)
        recipient.http_status = result.get("http_status")
        recipient.api_code = result.get("api_code", "")
        recipient.api_message = result.get("api_message", "")
        recipient.raw_response = result.get("raw_response", "")
        recipient.error_message = result.get("error", "")
        recipient.message_id = result.get("message_id", "")
        recipient.retry_count = result.get("retry_count", 0)
        recipient.sent_at = (
            timezone.now() if recipient.status == MessageRecipient.Status.SENT else None
        )
    recipient.save()
    batch.report_path = str(
        Path(settings.BASE_DIR) / "reports" / f"bale_report_single_test_{batch.id}.xlsx"
    )
    batch.save(update_fields=["report_path"])
    write_report(batch, Path(batch.report_path))
    _refresh_totals(batch, mark_finished=True)
    batch.status = (
        MessageBatch.Status.FINISHED
        if recipient.status == MessageRecipient.Status.SENT
        else MessageBatch.Status.FAILED
    )
    batch.error_message = (
        "" if batch.status == MessageBatch.Status.FINISHED else "ارسال تست موفق نبود."
    )
    batch.save(update_fields=["status", "error_message"])
    return batch


def run_excel_batch(
    *,
    file_path: str,
    message_template: str,
    dry_run=True,
    limit=None,
    sleep_seconds=None,
    sheet_name=None,
    button_text=None,
    button_url=None,
    skip_duplicates=True,
    report_path=None,
    range_start=None,
    range_end=None,
) -> MessageBatch:
    batch = MessageBatch.objects.create(
        source_file_name=Path(file_path).name,
        source_file_path=str(file_path),
        message_template=message_template,
        button_text=button_text or "",
        button_url=button_url or "",
        dry_run=dry_run,
        limit=limit,
        range_start=range_start,
        range_end=range_end,
        sheet_name=sheet_name or "",
        skip_duplicates=skip_duplicates,
        sleep_seconds=settings.BALE_DEFAULT_SLEEP_SECONDS
        if sleep_seconds is None
        else sleep_seconds,
    )
    return process_excel_batch(
        batch=batch,
        file_path=file_path,
        sleep_seconds=sleep_seconds,
        sheet_name=sheet_name,
        skip_duplicates=skip_duplicates,
        report_path=report_path,
        range_start=range_start,
        range_end=range_end,
    )
