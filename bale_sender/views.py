from pathlib import Path
from threading import Thread
from uuid import uuid4

from django.conf import settings
from django.contrib import messages
from django.db import close_old_connections
from django.db.models import Count
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from .forms import UploadExcelForm
from .models import MessageBatch, MessageRecipient
from .core import process_excel_batch



def _run_batch_in_background(batch_id: int, file_path: str, options: dict) -> None:
    close_old_connections()
    try:
        batch = MessageBatch.objects.get(pk=batch_id)
        process_excel_batch(batch=batch, file_path=file_path, **options)
    except Exception:
        # process_excel_batch stores the readable failure message on the batch.
        pass
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


def _batch_stats(batch: MessageBatch) -> dict[str, int]:
    counts = dict(batch.recipients.values("status").annotate(c=Count("id")).values_list("status", "c"))
    return {key: counts.get(key, 0) for key, _label in MessageRecipient.Status.choices}


def dashboard(request):
    if request.method == "POST":
        form = UploadExcelForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_path = _save_uploaded_file(form.cleaned_data["excel_file"])
            button_text = form.cleaned_data["button_text"] if form.cleaned_data["button_enabled"] else None
            button_url = form.cleaned_data["button_url"] if form.cleaned_data["button_enabled"] else None
            dry_run = form.cleaned_data["send_mode"] == "dry_run"
            batch = MessageBatch.objects.create(
                source_file_name=uploaded_path.name,
                message_template=form.cleaned_data["message_template"],
                button_text=button_text or "",
                button_url=button_url or "",
                dry_run=dry_run,
                limit=form.cleaned_data["limit"],
            )
            options = {
                "sleep_seconds": form.cleaned_data["sleep_seconds"],
                "sheet_name": form.cleaned_data["sheet_name"] or None,
                "skip_duplicates": form.cleaned_data["skip_duplicates"],
            }
            Thread(
                target=_run_batch_in_background,
                args=(batch.id, str(uploaded_path), options),
                daemon=True,
            ).start()
            messages.success(request, "فایل ثبت شد و پردازش در پس‌زمینه شروع شد. وضعیت را در همین صفحه دنبال کن.")
            return redirect("bale_batch_detail", batch_id=batch.id)
    else:
        form = UploadExcelForm()

    recent_batches = MessageBatch.objects.order_by("-started_at")[:10]
    return render(request, "bale_sender/dashboard.html", {"form": form, "recent_batches": recent_batches})


def batch_list(request):
    batches = MessageBatch.objects.order_by("-started_at")[:100]
    return render(request, "bale_sender/batch_list.html", {"batches": batches})


def batch_detail(request, batch_id):
    batch = get_object_or_404(MessageBatch, pk=batch_id)
    recipients = batch.recipients.all()[:300]
    return render(request, "bale_sender/batch_detail.html", {"batch": batch, "recipients": recipients, "stats": _batch_stats(batch)})


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
