import logging
from io import BytesIO
from pathlib import Path
from threading import Thread
from uuid import uuid4

from django.conf import settings
from django.contrib import messages
from django.core.paginator import Paginator
from django.core import signing
from django.db import close_old_connections, transaction
from django.db.models import Count
from django.http import FileResponse, Http404, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from openpyxl import Workbook

from .forms import SingleMessageTestForm, UploadExcelForm
from .models import MessageBatch, MessageRecipient
from .core import build_excel_preview, cancel_batch_immediately, process_excel_batch, send_single_recipient_test


logger = logging.getLogger(__name__)
DEFAULT_REPORT_PAGE_SIZE = 100
MAX_REPORT_PAGE_SIZE = 500
DEFAULT_EXPORT_LIMIT = 1000
MAX_EXPORT_LIMIT = 10000

def _run_batch_in_background(batch_id: int, file_path: str, options: dict) -> None:
    close_old_connections()
    try:
        batch = MessageBatch.objects.get(pk=batch_id)
        process_excel_batch(batch=batch, file_path=file_path, **options)
    except Exception as exc:
        # process_excel_batch stores the readable failure message on the batch.
        logger.exception("Error processing batch %s in background: %s", batch_id, exc)
    finally:
        close_old_connections()

def _save_uploaded_file(uploaded) -> Path:
    upload_dir = Path(settings.BASE_DIR) / "uploads" / timezone.localtime().strftime("%Y%m%d")
    upload_dir.mkdir(parents=True, exist_ok=True)
    safe_name = Path(uploaded.name).name.replace(" ", "_")
    file_path = upload_dir / f"{uuid4().hex[:10]}_{safe_name}"
    with open(file_path, "wb+") as dest:
        for chunk in uploaded.chunks():
            dest.write(chunk)
    return file_path


def _make_upload_token(file_path: Path) -> str:
    return signing.dumps({"path": str(file_path.resolve()), "name": file_path.name}, salt="bale-upload-preview")


def _resolve_upload_token(token: str) -> Path:
    try:
        data = signing.loads(token, salt="bale-upload-preview", max_age=60 * 60 * 6)
    except signing.BadSignature as exc:
        raise ValueError("فایل پیش‌نمایش معتبر نیست یا زمان آن تمام شده است. لطفاً فایل اکسل را دوباره انتخاب کن.") from exc

    path = Path(data.get("path", "")).resolve()
    uploads_root = (Path(settings.BASE_DIR) / "uploads").resolve()
    if uploads_root not in path.parents or not path.exists():
        raise ValueError("فایل ذخیره‌شده برای پیش‌نمایش پیدا نشد. لطفاً فایل اکسل را دوباره انتخاب کن.")
    return path


def _batch_stats(batch: MessageBatch) -> dict[str, int]:
    counts = dict(batch.recipients.values("status").annotate(c=Count("id")).values_list("status", "c"))
    return {key: counts.get(key, 0) for key, _label in MessageRecipient.Status.choices}


def _bounded_int(value: object, *, default: int, minimum: int, maximum: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return default
    return max(minimum, min(number, maximum))


def _recipients_report_response(recipients, *, filename: str, title: str) -> HttpResponse:
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.freeze_panes = "A2"
    ws.append(
        [
            "شناسه گزارش",
            "فایل/منبع",
            "حالت",
            "زمان ثبت ارسال",
            "ردیف اکسل",
            "نام",
            "نام خانوادگی",
            "نام کامل",
            "شماره خام",
            "شماره استاندارد",
            "وضعیت",
            "HTTP",
            "کد API",
            "پیام API",
            "خطا",
            "متن نهایی",
        ]
    )
    for recipient in recipients:
        batch = recipient.batch
        ws.append(
            [
                batch.id,
                batch.source_file_name,
                "تست بدون ارسال" if batch.dry_run else "ارسال واقعی",
                timezone.localtime(recipient.created_at).strftime("%Y-%m-%d %H:%M:%S"),
                recipient.row_number,
                recipient.first_name,
                recipient.last_name,
                recipient.full_name,
                recipient.raw_phone,
                recipient.normalized_phone,
                recipient.get_status_display(),
                recipient.http_status,
                recipient.api_code,
                recipient.api_message,
                recipient.error_message,
                recipient.final_text,
            ]
        )

    widths = [12, 26, 16, 20, 12, 16, 18, 24, 18, 18, 18, 10, 16, 30, 30, 50]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=index).column_letter].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    response = HttpResponse(
        output.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def dashboard(request):
    preview = None
    upload_form = UploadExcelForm()

    if request.method == "POST":
        action = request.POST.get("action", "send")
        upload_form = UploadExcelForm(request.POST, request.FILES, validate_send_confirmation=(action == "send"))
        if upload_form.is_valid():
            try:
                uploaded = upload_form.cleaned_data.get("excel_file")
                if uploaded:
                    uploaded_path = _save_uploaded_file(uploaded)
                else:
                    uploaded_path = _resolve_upload_token(upload_form.cleaned_data["uploaded_file_token"])

                if action == "preview":
                    upload_token = _make_upload_token(uploaded_path)
                    preview = build_excel_preview(
                        uploaded_path,
                        sheet_name=upload_form.cleaned_data["sheet_name"] or None,
                        message_template=upload_form.cleaned_data["message_template"],
                        limit=10,
                        range_start=upload_form.cleaned_data["range_start"],
                        range_end=upload_form.cleaned_data["range_end"],
                    )
                    preview["file_name"] = uploaded_path.name
                    post_data = request.POST.copy()
                    post_data["uploaded_file_token"] = upload_token
                    upload_form = UploadExcelForm(post_data, validate_send_confirmation=False)
                    messages.success(request, "پیش‌نمایش آماده شد.")
                else:
                    button_text = upload_form.cleaned_data["button_text"] if upload_form.cleaned_data["button_enabled"] else None
                    button_url = upload_form.cleaned_data["button_url"] if upload_form.cleaned_data["button_enabled"] else None
                    dry_run = upload_form.cleaned_data["send_mode"] == "dry_run"
                    batch = MessageBatch.objects.create(
                        source_file_name=uploaded_path.name,
                        message_template=upload_form.cleaned_data["message_template"],
                        button_text=button_text or "",
                        button_url=button_url or "",
                        dry_run=dry_run,
                        limit=upload_form.cleaned_data["limit"],
                        range_start=upload_form.cleaned_data["range_start"],
                        range_end=upload_form.cleaned_data["range_end"],
                    )
                    options = {
                        "sleep_seconds": upload_form.cleaned_data["sleep_seconds"],
                        "sheet_name": upload_form.cleaned_data["sheet_name"] or None,
                        "skip_duplicates": upload_form.cleaned_data["skip_duplicates"],
                        "range_start": upload_form.cleaned_data["range_start"],
                        "range_end": upload_form.cleaned_data["range_end"],
                    }
                    transaction.on_commit(
                        lambda: Thread(
                            target=_run_batch_in_background,
                            args=(batch.id, str(uploaded_path), options),
                            daemon=True,
                        ).start()
                    )
                    messages.success(request, "پردازش شروع شد.")
                    return redirect("bale_batch_detail", batch_id=batch.id)
            except ValueError as exc:
                upload_form.add_error(None, str(exc))

    recent_batches = MessageBatch.objects.order_by("-started_at")[:10]
    return render(
        request,
        "bale_sender/dashboard.html",
        {
            "form": upload_form,
            "preview": preview,
            "recent_batches": recent_batches,
        },
    )


def single_test(request):
    form = SingleMessageTestForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        button_text = form.cleaned_data["button_text"] if form.cleaned_data["button_enabled"] else None
        button_url = form.cleaned_data["button_url"] if form.cleaned_data["button_enabled"] else None
        batch = send_single_recipient_test(
            first_name=form.cleaned_data["first_name"],
            last_name=form.cleaned_data["last_name"],
            phone=form.cleaned_data["phone"],
            message_template=form.cleaned_data["message_template"],
            button_text=button_text,
            button_url=button_url,
        )
        if batch.status == MessageBatch.Status.FINISHED:
            messages.success(request, "ارسال تست موفق بود.")
        else:
            messages.warning(request, "ارسال تست موفق نبود. جزئیات را بررسی کن.")
        return redirect("bale_batch_detail", batch_id=batch.id)

    return render(request, "bale_sender/single_test.html", {"form": form})


def batch_list(request):
    batches = MessageBatch.objects.order_by("-started_at")[:100]
    return render(request, "bale_sender/batch_list.html", {"batches": batches})


def batch_detail(request, batch_id):
    batch = get_object_or_404(MessageBatch, pk=batch_id)
    per_page = _bounded_int(
        request.GET.get("per_page"),
        default=DEFAULT_REPORT_PAGE_SIZE,
        minimum=25,
        maximum=MAX_REPORT_PAGE_SIZE,
    )
    paginator = Paginator(batch.recipients.order_by("-created_at", "-id"), per_page)
    page_obj = paginator.get_page(request.GET.get("page"))
    return render(
        request,
        "bale_sender/batch_detail.html",
        {
            "batch": batch,
            "recipients": page_obj.object_list,
            "page_obj": page_obj,
            "per_page": per_page,
            "total_recipients": paginator.count,
            "stats": _batch_stats(batch),
        },
    )


def batch_live_status(request, batch_id):
    batch = get_object_or_404(MessageBatch, pk=batch_id)
    recent_recipients = (
        batch.recipients.order_by("-created_at", "-id")[:10]
    )
    stats = _batch_stats(batch)
    return JsonResponse(
        {
            "batch": {
                "id": batch.id,
                "status": batch.status,
                "status_label": batch.get_status_display(),
                "cancel_requested": batch.cancel_requested,
                "is_active": batch.status in {MessageBatch.Status.PENDING, MessageBatch.Status.RUNNING} and not batch.cancel_requested,
                "total_rows": sum(stats.values()),
                "total_sent": stats[MessageRecipient.Status.SENT],
                "total_failed": stats[MessageRecipient.Status.FAILED],
                "total_invalid": stats[MessageRecipient.Status.INVALID_PHONE],
                "total_duplicate": stats[MessageRecipient.Status.DUPLICATE],
                "total_not_bale_user": stats[MessageRecipient.Status.NOT_BALE_USER],
                "total_rate_limited": stats[MessageRecipient.Status.RATE_LIMITED],
                "total_payment_required": stats[MessageRecipient.Status.PAYMENT_REQUIRED],
                "total_config_error": stats[MessageRecipient.Status.CONFIG_ERROR],
            },
            "recent_recipients": [
                {
                    "row_number": r.row_number,
                    "full_name": r.full_name,
                    "normalized_phone": r.normalized_phone,
                    "status": r.status,
                    "status_label": r.get_status_display(),
                    "http_status": r.http_status or "",
                    "api_code": r.api_code,
                    "message": " ".join(x for x in [r.api_message, r.error_message] if x),
                }
                for r in recent_recipients
            ],
        }
    )

def cancel_batch(request, batch_id):
    if request.method != "POST":
        raise Http404("مسیر توقف فقط با درخواست معتبر قابل استفاده است.")

    batch = get_object_or_404(MessageBatch, pk=batch_id)
    if batch.status in {MessageBatch.Status.PENDING, MessageBatch.Status.RUNNING}:
        try:
            cancel_batch_immediately(batch)
        except Exception as exc:
            logger.exception("Could not write cancellation report for batch %s: %s", batch.id, exc)
            messages.error(request, "درخواست توقف ثبت شد، اما ساخت گزارش توقف با خطا روبه‌رو شد.")
            return redirect("bale_batch_detail", batch_id=batch.id)
        messages.warning(request, "درخواست توقف ثبت شد.")
    else:
        messages.info(request, "این پردازش دیگر در حال اجرا نیست.")
    return redirect("bale_batch_detail", batch_id=batch.id)


def download_report(request, batch_id):
    batch = get_object_or_404(MessageBatch, pk=batch_id)
    if not batch.report_path:
        raise Http404("گزارش برای این batch وجود ندارد.")
    report_path = Path(batch.report_path)
    if not report_path.is_absolute():
        report_path = Path(settings.BASE_DIR) / report_path
    if not report_path.exists():
        raise Http404("فایل گزارش پیدا نشد.")
    return FileResponse(open(report_path, "rb"), as_attachment=True, filename=report_path.name)


def download_recent_recipients_report(request):
    limit = _bounded_int(
        request.GET.get("limit"),
        default=DEFAULT_EXPORT_LIMIT,
        minimum=1,
        maximum=MAX_EXPORT_LIMIT,
    )
    recipients = (
        MessageRecipient.objects.select_related("batch")
        .order_by("-created_at", "-id")[:limit]
    )
    return _recipients_report_response(
        recipients,
        filename=f"bale_recent_{limit}_recipients.xlsx",
        title="آخرین ارسال‌ها",
    )
