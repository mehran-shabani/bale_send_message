from django import forms
from django.conf import settings

from .core import validate_message_template


class UploadExcelForm(forms.Form):
    SEND_MODES = (
        ("dry_run", "فقط تست؛ ارسال واقعی انجام نشود"),
        ("send", "ارسال واقعی"),
    )

    excel_file = forms.FileField(label="فایل اکسل")
    sheet_name = forms.CharField(label="نام شیت", required=False, help_text="اگر خالی باشد، اولین شیت خوانده می‌شود.")
    message_template = forms.CharField(
        label="متن پیام",
        widget=forms.Textarea(attrs={"rows": 8}),
        initial="سلام {full_name}\nثبت‌نام و انتقال پزشک خانواده به درمانگاه ولیعصر صغاد از طریق سایت زیر انجام می‌شود:\nhttps://helssa.ir\nلغو11",
        help_text="متغیرها: {first_name}، {last_name}، {full_name}، {phone}",
    )
    send_mode = forms.ChoiceField(label="نوع اجرا", choices=SEND_MODES, initial="dry_run", widget=forms.RadioSelect)
    confirm_real_send = forms.BooleanField(label="تأیید می‌کنم ارسال واقعی انجام شود", required=False)
    limit = forms.IntegerField(label="محدودیت تعداد", required=False, min_value=1, help_text="برای تست، مثلاً ۵ بگذار.")
    sleep_seconds = forms.FloatField(label="فاصله بین ارسال‌ها / ثانیه", initial=settings.BALE_DEFAULT_SLEEP_SECONDS, min_value=0, required=False)
    skip_duplicates = forms.BooleanField(label="شماره‌های تکراری داخل فایل ارسال نشوند", initial=True, required=False)
    button_enabled = forms.BooleanField(label="دکمه لینک‌دار زیر پیام فعال باشد", initial=True, required=False)
    button_text = forms.CharField(label="متن دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_TEXT)
    button_url = forms.URLField(label="لینک دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_URL)

    def clean_excel_file(self):
        f = self.cleaned_data["excel_file"]
        name = f.name.lower()
        if not name.endswith((".xlsx", ".xlsm")):
            raise forms.ValidationError("فقط فایل Excel با پسوند xlsx یا xlsm پذیرفته می‌شود.")

        max_size_mb = getattr(settings, "BALE_MAX_UPLOAD_SIZE_MB", 10)
        max_size_bytes = max_size_mb * 1024 * 1024
        if f.size > max_size_bytes:
            raise forms.ValidationError(f"حجم فایل اکسل نباید بیشتر از {max_size_mb} مگابایت باشد.")
        return f

    def clean_message_template(self):
        template = self.cleaned_data["message_template"]
        try:
            validate_message_template(template)
        except ValueError as exc:
            raise forms.ValidationError(str(exc)) from exc
        return template

    def clean(self):
        data = super().clean()
        if data.get("send_mode") == "send" and not data.get("confirm_real_send"):
            raise forms.ValidationError("برای ارسال واقعی باید گزینه تأیید ارسال واقعی را فعال کنی.")
        if data.get("button_enabled") and not (data.get("button_text") and data.get("button_url")):
            raise forms.ValidationError("وقتی دکمه فعال است، متن دکمه و لینک دکمه باید پر باشند.")
        return data
