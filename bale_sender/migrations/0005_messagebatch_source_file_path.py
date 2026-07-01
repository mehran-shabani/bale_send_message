from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("bale_sender", "0004_messagebatch_range_start_messagebatch_range_end"),
    ]

    operations = [
        migrations.AddField(
            model_name="messagebatch",
            name="source_file_path",
            field=models.CharField(blank=True, max_length=500),
        ),
    ]
