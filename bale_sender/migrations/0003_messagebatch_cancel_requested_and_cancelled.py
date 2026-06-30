from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("bale_sender", "0002_messagebatch_error_message_messagebatch_status"),
    ]

    operations = [
        migrations.AddField(
            model_name="messagebatch",
            name="cancel_requested",
            field=models.BooleanField(default=False),
        ),
        migrations.AlterField(
            model_name="messagebatch",
            name="status",
            field=models.CharField(
                choices=[
                    ("pending", "در انتظار"),
                    ("running", "در حال پردازش"),
                    ("finished", "پایان‌یافته"),
                    ("failed", "ناموفق"),
                    ("cancelled", "متوقف‌شده"),
                ],
                db_index=True,
                default="pending",
                max_length=20,
            ),
        ),
    ]
