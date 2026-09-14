from django.db import models
from django.utils import timezone


class SafirConfiguration(models.Model):
    """Single local configuration row for the Safir client.

    This project is deliberately a local dashboard without authentication.  The
    access key is therefore stored in the local SQLite database and is never
    included in reports or templates.
    """

    api_access_key = models.CharField(max_length=500, blank=True)
    bot_id = models.PositiveBigIntegerField(default=0)
    send_url = models.URLField(default="https://safir.bale.ai/api/v3/send_message")
    upload_url = models.URLField(default="https://safir.bale.ai/api/v3/upload_file")
    request_timeout = models.PositiveIntegerField(default=20)
    default_sleep_seconds = models.FloatField(default=0.4)
    max_retries = models.PositiveSmallIntegerField(default=2)
    message_price_rial = models.PositiveBigIntegerField(default=0)
    global_variables = models.JSONField(default=dict, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "تنظیمات سفیر"
        verbose_name_plural = "تنظیمات سفیر"

    @classmethod
    def load(cls) -> "SafirConfiguration":
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj

    def save(self, *args, **kwargs):
        self.pk = 1
        return super().save(*args, **kwargs)

    def __str__(self):
        return "تنظیمات محلی سفیر بله"


class MessageBatch(models.Model):
    class Status(models.TextChoices):
        PENDING = "pending", "در انتظار"
        RUNNING = "running", "در حال پردازش"
        FINISHED = "finished", "پایان‌یافته"
        FAILED = "failed", "ناموفق"
        CANCELLED = "cancelled", "متوقف‌شده"

    class MessageType(models.TextChoices):
        TEXT = "text", "متن / فایل"
        TEMPLATE = "template", "قالب تأییدشده سفیر"
        OTP = "otp", "رمز یک‌بارمصرف"

    source_file_name = models.CharField(max_length=255)
    source_file_path = models.CharField(max_length=500, blank=True)
    attachment_path = models.CharField(max_length=500, blank=True)
    attachment_file_id = models.CharField(max_length=500, blank=True)
    sheet_name = models.CharField(max_length=255, blank=True)
    header_row = models.PositiveIntegerField(default=1)
    phone_column = models.CharField(max_length=255, blank=True)
    first_name_column = models.CharField(max_length=255, blank=True)
    last_name_column = models.CharField(max_length=255, blank=True)
    variable_mapping = models.JSONField(default=dict, blank=True)
    global_variables = models.JSONField(default=dict, blank=True)
    message_type = models.CharField(
        max_length=20, choices=MessageType.choices, default=MessageType.TEXT
    )
    message_template = models.TextField()
    template_id = models.CharField(max_length=255, blank=True)
    template_fields = models.JSONField(default=dict, blank=True)
    otp_template = models.CharField(max_length=255, blank=True)
    is_secure = models.BooleanField(default=False)
    copy_text_template = models.TextField(blank=True)
    buttons = models.JSONField(default=list, blank=True)
    button_text = models.CharField(max_length=80, blank=True)
    button_url = models.URLField(blank=True)
    dry_run = models.BooleanField(default=True)
    sleep_seconds = models.FloatField(default=0.4)
    max_retries = models.PositiveSmallIntegerField(default=2)
    skip_duplicates = models.BooleanField(default=True)
    limit = models.PositiveIntegerField(null=True, blank=True)
    range_start = models.PositiveIntegerField(null=True, blank=True)
    range_end = models.PositiveIntegerField(null=True, blank=True)
    total_rows = models.PositiveIntegerField(default=0)
    total_sent = models.PositiveIntegerField(default=0)
    total_failed = models.PositiveIntegerField(default=0)
    total_invalid = models.PositiveIntegerField(default=0)
    total_duplicate = models.PositiveIntegerField(default=0)
    total_not_bale_user = models.PositiveIntegerField(default=0)
    total_rate_limited = models.PositiveIntegerField(default=0)
    total_payment_required = models.PositiveIntegerField(default=0)
    total_config_error = models.PositiveIntegerField(default=0)
    status = models.CharField(
        max_length=20, choices=Status.choices, default=Status.PENDING, db_index=True
    )
    cancel_requested = models.BooleanField(default=False)
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

    batch = models.ForeignKey(
        MessageBatch, on_delete=models.CASCADE, related_name="recipients"
    )
    row_number = models.PositiveIntegerField()
    first_name = models.CharField(max_length=120, blank=True)
    last_name = models.CharField(max_length=120, blank=True)
    full_name = models.CharField(max_length=250, blank=True)
    raw_phone = models.CharField(max_length=50, blank=True)
    normalized_phone = models.CharField(max_length=20, blank=True, db_index=True)
    final_text = models.TextField(blank=True)
    variables = models.JSONField(default=dict, blank=True)
    message_id = models.CharField(max_length=255, blank=True, db_index=True)
    retry_count = models.PositiveSmallIntegerField(default=0)
    request_id = models.CharField(max_length=120, blank=True, db_index=True)
    status = models.CharField(
        max_length=30, choices=Status.choices, default=Status.PENDING, db_index=True
    )
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
