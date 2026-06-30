# Generated manually for bale_sender
from django.db import migrations, models
import django.db.models.deletion
import django.utils.timezone


class Migration(migrations.Migration):
    initial = True

    dependencies = []

    operations = [
        migrations.CreateModel(
            name="MessageBatch",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("source_file_name", models.CharField(max_length=255)),
                ("message_template", models.TextField()),
                ("button_text", models.CharField(blank=True, max_length=80)),
                ("button_url", models.URLField(blank=True)),
                ("dry_run", models.BooleanField(default=True)),
                ("limit", models.PositiveIntegerField(blank=True, null=True)),
                ("total_rows", models.PositiveIntegerField(default=0)),
                ("total_sent", models.PositiveIntegerField(default=0)),
                ("total_failed", models.PositiveIntegerField(default=0)),
                ("total_invalid", models.PositiveIntegerField(default=0)),
                ("total_duplicate", models.PositiveIntegerField(default=0)),
                ("total_not_bale_user", models.PositiveIntegerField(default=0)),
                ("total_rate_limited", models.PositiveIntegerField(default=0)),
                ("total_payment_required", models.PositiveIntegerField(default=0)),
                ("total_config_error", models.PositiveIntegerField(default=0)),
                ("report_path", models.CharField(blank=True, max_length=500)),
                ("started_at", models.DateTimeField(default=django.utils.timezone.now)),
                ("finished_at", models.DateTimeField(blank=True, null=True)),
            ],
        ),
        migrations.CreateModel(
            name="MessageRecipient",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("row_number", models.PositiveIntegerField()),
                ("first_name", models.CharField(blank=True, max_length=120)),
                ("last_name", models.CharField(blank=True, max_length=120)),
                ("full_name", models.CharField(blank=True, max_length=250)),
                ("raw_phone", models.CharField(blank=True, max_length=50)),
                ("normalized_phone", models.CharField(blank=True, db_index=True, max_length=20)),
                ("final_text", models.TextField(blank=True)),
                ("request_id", models.CharField(blank=True, db_index=True, max_length=120)),
                ("status", models.CharField(choices=[("pending", "در انتظار"), ("dry_run", "تست بدون ارسال"), ("sent", "ارسال موفق API"), ("failed", "خطا"), ("invalid_phone", "شماره نامعتبر"), ("duplicate", "تکراری"), ("not_bale_user", "کاربر بله نیست"), ("rate_limited", "محدودیت سرعت"), ("payment_required", "نیاز به شارژ"), ("config_error", "خطای تنظیمات")], db_index=True, default="pending", max_length=30)),
                ("http_status", models.PositiveIntegerField(blank=True, null=True)),
                ("api_code", models.CharField(blank=True, max_length=100)),
                ("api_message", models.TextField(blank=True)),
                ("raw_response", models.TextField(blank=True)),
                ("error_message", models.TextField(blank=True)),
                ("created_at", models.DateTimeField(default=django.utils.timezone.now)),
                ("sent_at", models.DateTimeField(blank=True, null=True)),
                ("batch", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="recipients", to="bale_sender.messagebatch")),
            ],
            options={"ordering": ["row_number", "id"]},
        ),
    ]
