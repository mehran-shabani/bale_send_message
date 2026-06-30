from django.db import models
from django.utils import timezone


class MessageBatch(models.Model):
    class Status(models.TextChoices):
        PENDING = "pending", "در انتظار"
        RUNNING = "running", "در حال پردازش"
        FINISHED = "finished", "پایان‌یافته"
        FAILED = "failed", "ناموفق"

    source_file_name = models.CharField(max_length=255)
    message_template = models.TextField()
    button_text = models.CharField(max_length=80, blank=True)
    button_url = models.URLField(blank=True)
    dry_run = models.BooleanField(default=True)
    limit = models.PositiveIntegerField(null=True, blank=True)
    total_rows = models.PositiveIntegerField(default=0)
    total_sent = models.PositiveIntegerField(default=0)
    total_failed = models.PositiveIntegerField(default=0)
    total_invalid = models.PositiveIntegerField(default=0)
    total_duplicate = models.PositiveIntegerField(default=0)
    total_not_bale_user = models.PositiveIntegerField(default=0)
    total_rate_limited = models.PositiveIntegerField(default=0)
    total_payment_required = models.PositiveIntegerField(default=0)
    total_config_error = models.PositiveIntegerField(default=0)
    status = models.CharField(max_length=20, choices=Status.choices, default=Status.PENDING, db_index=True)
    error_message = models.TextField(blank=True)
    report_path = models.CharField(max_length=500, blank=True)
    started_at = models.DateTimeField(default=timezone.now)
    finished_at = models.DateTimeField(null=True, blank=True)

    def __str__(self):
        mode = "dry-run" if self.dry_run else "send"
        return f"#{self.pk} - {self.source_file_name} - {mode}"


class MessageRecipient(models.Model):
    class Status(models.TextChoices):
        PENDING = "pending", "در انتظار"
        DRY_RUN = "dry_run", "تست بدون ارسال"
        SENT = "sent", "ارسال موفق API"
        FAILED = "failed", "خطا"
        INVALID_PHONE = "invalid_phone", "شماره نامعتبر"
        DUPLICATE = "duplicate", "تکراری"
        NOT_BALE_USER = "not_bale_user", "کاربر بله نیست"
        RATE_LIMITED = "rate_limited", "محدودیت سرعت"
        PAYMENT_REQUIRED = "payment_required", "نیاز به شارژ"
        CONFIG_ERROR = "config_error", "خطای تنظیمات"

    batch = models.ForeignKey(MessageBatch, on_delete=models.CASCADE, related_name="recipients")
    row_number = models.PositiveIntegerField()
    first_name = models.CharField(max_length=120, blank=True)
    last_name = models.CharField(max_length=120, blank=True)
    full_name = models.CharField(max_length=250, blank=True)
    raw_phone = models.CharField(max_length=50, blank=True)
    normalized_phone = models.CharField(max_length=20, blank=True, db_index=True)
    final_text = models.TextField(blank=True)
    request_id = models.CharField(max_length=120, blank=True, db_index=True)
    status = models.CharField(max_length=30, choices=Status.choices, default=Status.PENDING, db_index=True)
    http_status = models.PositiveIntegerField(null=True, blank=True)
    api_code = models.CharField(max_length=100, blank=True)
    api_message = models.TextField(blank=True)
    raw_response = models.TextField(blank=True)
    error_message = models.TextField(blank=True)
    created_at = models.DateTimeField(default=timezone.now)
    sent_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ["row_number", "id"]

    def __str__(self):
        return f"{self.full_name} - {self.normalized_phone} - {self.status}"
