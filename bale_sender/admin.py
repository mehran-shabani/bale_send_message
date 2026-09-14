from django.contrib import admin

from .models import MessageBatch, MessageRecipient, SafirConfiguration


@admin.register(SafirConfiguration)
class SafirConfigurationAdmin(admin.ModelAdmin):
    exclude = ("api_access_key",)
    readonly_fields = ("updated_at",)

    def has_add_permission(self, request):
        return not SafirConfiguration.objects.exists()

    def has_delete_permission(self, request, obj=None):
        return False


@admin.register(MessageBatch)
class MessageBatchAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "source_file_name",
        "message_type",
        "dry_run",
        "status",
        "cancel_requested",
        "range_start",
        "range_end",
        "total_rows",
        "total_sent",
        "total_failed",
        "total_invalid",
        "total_duplicate",
        "total_not_bale_user",
        "started_at",
        "finished_at",
    )
    list_filter = (
        "status",
        "message_type",
        "dry_run",
        "cancel_requested",
        "started_at",
    )
    readonly_fields = ("started_at", "finished_at")


@admin.register(MessageRecipient)
class MessageRecipientAdmin(admin.ModelAdmin):
    list_display = (
        "batch",
        "row_number",
        "full_name",
        "normalized_phone",
        "status",
        "http_status",
        "api_code",
        "retry_count",
        "sent_at",
    )
    list_filter = ("status", "batch", "sent_at")
    search_fields = (
        "full_name",
        "raw_phone",
        "normalized_phone",
        "api_message",
        "error_message",
    )
    readonly_fields = ("created_at", "sent_at")
