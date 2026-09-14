from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [("bale_sender", "0005_messagebatch_source_file_path")]

    operations = [
        migrations.CreateModel(
            name="SafirConfiguration",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                ("api_access_key", models.CharField(blank=True, max_length=500)),
                ("bot_id", models.PositiveBigIntegerField(default=0)),
                (
                    "send_url",
                    models.URLField(
                        default="https://safir.bale.ai/api/v3/send_message"
                    ),
                ),
                (
                    "upload_url",
                    models.URLField(default="https://safir.bale.ai/api/v3/upload_file"),
                ),
                ("request_timeout", models.PositiveIntegerField(default=20)),
                ("default_sleep_seconds", models.FloatField(default=0.4)),
                ("max_retries", models.PositiveSmallIntegerField(default=2)),
                ("message_price_rial", models.PositiveBigIntegerField(default=0)),
                ("global_variables", models.JSONField(blank=True, default=dict)),
                ("updated_at", models.DateTimeField(auto_now=True)),
            ],
            options={
                "verbose_name": "تنظیمات سفیر",
                "verbose_name_plural": "تنظیمات سفیر",
            },
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="attachment_file_id",
            field=models.CharField(blank=True, max_length=500),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="attachment_path",
            field=models.CharField(blank=True, max_length=500),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="buttons",
            field=models.JSONField(blank=True, default=list),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="copy_text_template",
            field=models.TextField(blank=True),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="first_name_column",
            field=models.CharField(blank=True, max_length=255),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="global_variables",
            field=models.JSONField(blank=True, default=dict),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="header_row",
            field=models.PositiveIntegerField(default=1),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="is_secure",
            field=models.BooleanField(default=False),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="last_name_column",
            field=models.CharField(blank=True, max_length=255),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="max_retries",
            field=models.PositiveSmallIntegerField(default=2),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="message_type",
            field=models.CharField(
                choices=[
                    ("text", "متن / فایل"),
                    ("template", "قالب تأییدشده سفیر"),
                    ("otp", "رمز یک‌بارمصرف"),
                ],
                default="text",
                max_length=20,
            ),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="otp_template",
            field=models.CharField(blank=True, max_length=255),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="phone_column",
            field=models.CharField(blank=True, max_length=255),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="sheet_name",
            field=models.CharField(blank=True, max_length=255),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="skip_duplicates",
            field=models.BooleanField(default=True),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="sleep_seconds",
            field=models.FloatField(default=0.4),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="template_fields",
            field=models.JSONField(blank=True, default=dict),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="template_id",
            field=models.CharField(blank=True, max_length=255),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="variable_mapping",
            field=models.JSONField(blank=True, default=dict),
        ),
        migrations.AddField(
            model_name="messagerecipient",
            name="message_id",
            field=models.CharField(blank=True, db_index=True, max_length=255),
        ),
        migrations.AddField(
            model_name="messagerecipient",
            name="retry_count",
            field=models.PositiveSmallIntegerField(default=0),
        ),
        migrations.AddField(
            model_name="messagerecipient",
            name="variables",
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
