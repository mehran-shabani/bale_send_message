from django import forms
from django.conf import settings

from .core import validate_message_template


class UploadExcelForm(forms.Form):
    SEND_MODES = (
        ("dry_run", "فقط تست؛ ارسال واقعی انجام نشود"),
        ("send", "ارسال واقعی"),
    )

    excel_file = forms.FileField(label="فایل اکسل", required=False)
    uploaded_file_token = forms.CharField(required=False, widget=forms.HiddenInput)
    sheet_name = forms.CharField(label="نام شیت", required=False, help_text="اگر خالی باشد، اولین شیت خوانده می‌شود.")
    message_template = forms.CharField(
        label="متن پیام",
        widget=forms.Textarea(attrs={"rows": 8}),
        initial="سلام {full_name}\nثبت‌نام و انتقال پزشک خانواده به درمانگاه ولیعصر صغاد از طریق سایت زیر انجام می‌شود:\nhttps://helssa.ir\nلغو11",
        help_text="متغیرها: {first_name}، {last_name}، {full_name}، {phone}",
    )
    send_mode = forms.ChoiceField(label="نوع اجرا", choices=SEND_MODES, initial="dry_run", widget=forms.RadioSelect)
    confirm_real_send = forms.BooleanField(label="تأیید می‌کنم ارسال واقعی انجام شود", required=False)
    range_start = forms.IntegerField(
        label="از ردیف داده",
        required=False,
        min_value=1,
        help_text="ردیف ۱ یعنی اولین ردیف بعد از header اکسل.",
    )
    range_end = forms.IntegerField(
        label="تا ردیف داده",
        required=False,
        min_value=1,
        help_text="برای مثال ۱ تا ۱۰۰، سپس ۱۰۱ تا ۲۰۰.",
    )
    limit = forms.IntegerField(label="محدودیت تعداد", required=False, min_value=1, help_text="برای تست، مثلاً ۵ بگذار.")
    sleep_seconds = forms.FloatField(label="فاصله بین ارسال‌ها / ثانیه", initial=settings.BALE_DEFAULT_SLEEP_SECONDS, min_value=0, required=False)
    skip_duplicates = forms.BooleanField(label="شماره‌های تکراری داخل فایل ارسال نشوند", initial=True, required=False)
    button_enabled = forms.BooleanField(label="دکمه لینک‌دار زیر پیام فعال باشد", initial=True, required=False)
    button_text = forms.CharField(label="متن دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_TEXT)
    button_url = forms.URLField(label="لینک دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_URL)

    def __init__(self, *args, validate_send_confirmation: bool = True, **kwargs):
        self.validate_send_confirmation = validate_send_confirmation
        super().__init__(*args, **kwargs)

    def clean_excel_file(self):
        f = self.cleaned_data.get("excel_file")
        if not f:
            return f
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
        if not data.get("excel_file") and not data.get("uploaded_file_token"):
            self.add_error("excel_file", "برای پیش‌نمایش یا ارسال، فایل اکسل را انتخاب کن.")
        if self.validate_send_confirmation and data.get("send_mode") == "send" and not data.get("confirm_real_send"):
            raise forms.ValidationError("برای ارسال واقعی باید گزینه تأیید ارسال واقعی را فعال کنی.")
        if data.get("range_start") and data.get("range_end") and data["range_end"] < data["range_start"]:
            self.add_error("range_end", "عدد «تا ردیف داده» باید بزرگ‌تر یا مساوی «از ردیف داده» باشد.")
        if data.get("button_enabled") and not (data.get("button_text") and data.get("button_url")):
            raise forms.ValidationError("وقتی دکمه فعال است، متن دکمه و لینک دکمه باید پر باشند.")
        return data


class SingleMessageTestForm(forms.Form):
    first_name = forms.CharField(label="نام", max_length=120)
    last_name = forms.CharField(label="نام خانوادگی", max_length=120)
    phone = forms.CharField(label="شماره موبایل", max_length=50, help_text="مثلاً 09123456789")
    message_template = forms.CharField(
        label="متن پیام تست",
        widget=forms.Textarea(attrs={"rows": 6}),
        initial=UploadExcelForm.base_fields["message_template"].initial,
        help_text=UploadExcelForm.base_fields["message_template"].help_text,
    )
    button_enabled = forms.BooleanField(label="دکمه لینک‌دار زیر پیام فعال باشد", initial=True, required=False)
    button_text = forms.CharField(label="متن دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_TEXT)
    button_url = forms.URLField(label="لینک دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_URL)

    def clean_message_template(self):
        template = self.cleaned_data["message_template"]
        try:
            validate_message_template(template)
        except ValueError as exc:
            raise forms.ValidationError(str(exc)) from exc
        return template

    def clean_phone(self):
        from .core import normalize_iran_mobile

        phone = self.cleaned_data["phone"]
        normalized_phone = normalize_iran_mobile(phone)
        if not normalized_phone:
            raise forms.ValidationError("شماره موبایل معتبر نیست. نمونه درست: 09123456789")
        self.cleaned_data["normalized_phone"] = normalized_phone
        return phone

    def clean(self):
        data = super().clean()
        if data.get("button_enabled") and not (data.get("button_text") and data.get("button_url")):
            raise forms.ValidationError("وقتی دکمه فعال است، متن دکمه و لینک دکمه باید پر باشند.")
        return data
