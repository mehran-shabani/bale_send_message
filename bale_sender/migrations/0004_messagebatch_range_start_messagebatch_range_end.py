from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("bale_sender", "0003_messagebatch_cancel_requested_and_cancelled"),
    ]

    operations = [
        migrations.AddField(
            model_name="messagebatch",
            name="range_start",
            field=models.PositiveIntegerField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="messagebatch",
            name="range_end",
            field=models.PositiveIntegerField(blank=True, null=True),
        ),
    ]
