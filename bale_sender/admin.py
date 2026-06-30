from django.contrib import admin
from .models import MessageBatch, MessageRecipient


@admin.register(MessageBatch)
class MessageBatchAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "source_file_name",
        "dry_run",
        "status",
        "cancel_requested",
        "total_rows",
        "total_sent",
        "total_failed",
        "total_invalid",
        "total_duplicate",
        "total_not_bale_user",
        "started_at",
        "finished_at",
    )
    list_filter = ("status", "dry_run", "cancel_requested", "started_at")
    readonly_fields = ("started_at", "finished_at")


@admin.register(MessageRecipient)
class MessageRecipientAdmin(admin.ModelAdmin):
    list_display = ("batch", "row_number", "full_name", "normalized_phone", "status", "http_status", "api_code", "sent_at")
    list_filter = ("status", "batch", "sent_at")
    search_fields = ("full_name", "raw_phone", "normalized_phone", "api_message", "error_message")
    readonly_fields = ("created_at", "sent_at")
