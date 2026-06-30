from django.core.management.base import BaseCommand, CommandError

from bale_sender.core import run_excel_batch


class Command(BaseCommand):
    help = "Send Bale Safir messages from an Excel file."

    def add_arguments(self, parser):
        parser.add_argument("--file", required=True, help="Path to Excel file")
        parser.add_argument("--message", required=True, help="Message template. Supports {full_name}, {first_name}, {last_name}, {phone}")
        parser.add_argument("--sheet", default=None, help="Sheet name. Empty means first sheet")
        parser.add_argument("--range-start", type=int, default=None, help="First data row to process. 1 means first row after header")
        parser.add_argument("--range-end", type=int, default=None, help="Last data row to process, inclusive")
        parser.add_argument("--limit", type=int, default=None, help="Limit rows for test")
        parser.add_argument("--sleep", type=float, default=None, help="Seconds between real sends")
        parser.add_argument("--button-text", default=None, help="Inline button text")
        parser.add_argument("--button-url", default=None, help="Inline button URL")
        parser.add_argument("--send", action="store_true", help="Actually send messages. Without this, dry-run is used.")
        parser.add_argument("--allow-duplicates", action="store_true", help="Do not skip duplicate phone numbers")

    def handle(self, *args, **options):
        dry_run = not options["send"]
        try:
            batch = run_excel_batch(
                file_path=options["file"],
                message_template=options["message"],
                dry_run=dry_run,
                limit=options["limit"],
                sleep_seconds=options["sleep"],
                sheet_name=options["sheet"],
                range_start=options["range_start"],
                range_end=options["range_end"],
                button_text=options["button_text"],
                button_url=options["button_url"],
                skip_duplicates=not options["allow_duplicates"],
            )
        except Exception as exc:
            raise CommandError(str(exc)) from exc

        self.stdout.write(self.style.SUCCESS(f"Batch #{batch.id} finished"))
        self.stdout.write(f"Total: {batch.total_rows}")
        self.stdout.write(f"Sent: {batch.total_sent}")
        self.stdout.write(f"Dry-run: {batch.recipients.filter(status='dry_run').count()}")
        self.stdout.write(f"Invalid: {batch.total_invalid}")
        self.stdout.write(f"Duplicate: {batch.total_duplicate}")
        self.stdout.write(f"Not Bale user: {batch.total_not_bale_user}")
        self.stdout.write(f"Failed: {batch.total_failed}")
        self.stdout.write(f"Report: {batch.report_path}")
