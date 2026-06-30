from pathlib import Path
from uuid import uuid4

from django.conf import settings
from django.contrib import messages
from django.db.models import Count
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from .forms import UploadExcelForm
from .models import MessageBatch, MessageRecipient
from .core import run_excel_batch


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
            try:
                batch = run_excel_batch(
                    file_path=str(uploaded_path),
                    message_template=form.cleaned_data["message_template"],
                    dry_run=dry_run,
                    limit=form.cleaned_data["limit"],
                    sleep_seconds=form.cleaned_data["sleep_seconds"],
                    sheet_name=form.cleaned_data["sheet_name"] or None,
                    button_text=button_text,
                    button_url=button_url,
                    skip_duplicates=form.cleaned_data["skip_duplicates"],
                )
            except Exception as exc:
                messages.error(request, f"پردازش انجام نشد: {exc}")
            else:
                if dry_run:
                    messages.success(request, "تست انجام شد و گزارش ساخته شد. اگر همه چیز درست بود، ارسال واقعی را اجرا کن.")
                else:
                    messages.success(request, "ارسال واقعی انجام شد و گزارش کامل آماده است.")
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
