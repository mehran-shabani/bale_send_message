import json
import re

from django import forms
from django.conf import settings

from .core import BUILTIN_TOKENS, validate_message_template
from .models import MessageBatch, SafirConfiguration


def _json_object(value: str, label: str) -> dict:
    if not (value or "").strip():
        return {}
    try:
        result = json.loads(value)
    except json.JSONDecodeError as exc:
        raise forms.ValidationError(
            f"{label} باید JSON معتبر باشد؛ خط {exc.lineno}، ستون {exc.colno}."
        ) from exc
    if not isinstance(result, dict):
        raise forms.ValidationError(f"{label} باید یک JSON object باشد.")
    return result


def _json_list(value: str, label: str) -> list:
    if not (value or "").strip():
        return []
    try:
        result = json.loads(value)
    except json.JSONDecodeError as exc:
        raise forms.ValidationError(
            f"{label} باید JSON معتبر باشد؛ خط {exc.lineno}، ستون {exc.colno}."
        ) from exc
    if not isinstance(result, list):
        raise forms.ValidationError(f"{label} باید یک JSON array باشد.")
    return result


class UploadExcelForm(forms.Form):
    SEND_MODES = (
        ("dry_run", "فقط پیش‌پردازش؛ هیچ پیامی ارسال نشود"),
        ("send", "ارسال واقعی از طریق سفیر"),
    )

    excel_file = forms.FileField(label="فایل اکسل", required=False)
    uploaded_file_token = forms.CharField(required=False, widget=forms.HiddenInput)
    attachment = forms.FileField(
        label="فایل پیوست پیام (اختیاری)",
        required=False,
        help_text="طبق مستندات سفیر تا ۵۰۰MB؛ سقف محلی از تنظیمات پروژه خوانده می‌شود.",
    )
    uploaded_attachment_token = forms.CharField(
        required=False, widget=forms.HiddenInput
    )
    sheet_name = forms.CharField(
        label="نام شیت",
        required=False,
        help_text="اگر خالی باشد، اولین شیت خوانده می‌شود.",
    )
    header_row = forms.IntegerField(
        label="ردیف عنوان ستون‌ها",
        initial=1,
        min_value=1,
        required=False,
        help_text="معمولاً ۱؛ اگر بالای جدول توضیح داری عدد درست را وارد کن.",
    )
    phone_column = forms.CharField(
        label="ستون موبایل",
        required=False,
        help_text="خالی = تشخیص خودکار؛ یا نام ستون/حرف Excel مثل C.",
    )
    first_name_column = forms.CharField(
        label="ستون نام",
        required=False,
        help_text="اختیاری؛ نام ستون یا حرف Excel.",
    )
    last_name_column = forms.CharField(
        label="ستون نام خانوادگی",
        required=False,
        help_text="اختیاری؛ نام ستون یا حرف Excel.",
    )
    variable_mapping = forms.CharField(
        label="نگاشت متغیرها به ستون‌های اکسل",
        required=False,
        widget=forms.Textarea(attrs={"rows": 5, "dir": "ltr"}),
        initial="{}",
        help_text='JSON؛ مثال: {"name":"A","visit_date":"تاریخ مراجعه","amount":"D"}',
    )
    message_type = forms.ChoiceField(
        label="نوع پیام سفیر",
        choices=MessageBatch.MessageType.choices,
        initial=MessageBatch.MessageType.TEXT,
        required=False,
    )
    message_template = forms.CharField(
        label="متن پیام",
        required=False,
        widget=forms.Textarea(attrs={"rows": 8}),
        initial=(
            "سلام {{full_name}}\n"
            "ثبت‌نام و انتقال پزشک خانواده به درمانگاه ولیعصر صغاد:\n"
            "https://helssa.ir\nلغو11"
        ),
        help_text="متغیر را با {{variable}} بنویس. متن ثابت هم مجاز است.",
    )
    template_id = forms.CharField(
        label="شناسه قالب تأییدشده سفیر",
        required=False,
        help_text="فقط برای نوع «قالب تأییدشده»؛ قالب باید قبلاً در پنل بله تأیید شده باشد.",
    )
    template_fields = forms.CharField(
        label="فیلدهای قالب سفیر",
        required=False,
        widget=forms.Textarea(attrs={"rows": 5, "dir": "ltr"}),
        initial="{}",
        help_text='JSON؛ مثال: {"name":"{{full_name}}","date":"{{visit_date}}"}',
    )
    otp_template = forms.CharField(
        label="مقدار OTP",
        required=False,
        help_text="عدد ثابت یا متغیر؛ مثلاً {{otp}}. خروجی هر ردیف باید فقط عدد باشد.",
    )
    is_secure = forms.BooleanField(
        label="پیام رمزدار/محرمانه باشد",
        required=False,
        help_text="برای متن، فایل یا قالب سفیر قابل استفاده است.",
    )
    copy_text_template = forms.CharField(
        label="متن دکمه رونوشت خود پیام",
        required=False,
        help_text="اختیاری؛ می‌تواند متغیر داشته باشد.",
    )
    buttons_json = forms.CharField(
        label="دکمه‌های شیشه‌ای",
        required=False,
        widget=forms.Textarea(attrs={"rows": 6, "dir": "ltr"}),
        initial="[]",
        help_text=(
            "JSON؛ type یکی از url، web_app، copy_text. مثال: "
            '[{"text":"ورود","type":"url","value":"https://helssa.ir","row":1}]'
        ),
    )
    send_mode = forms.ChoiceField(
        label="نوع اجرا",
        choices=SEND_MODES,
        initial="dry_run",
        widget=forms.RadioSelect,
    )
    confirm_real_send = forms.BooleanField(
        label="تأیید می‌کنم ارسال واقعی انجام شود",
        required=False,
    )
    range_start = forms.IntegerField(
        label="از ردیف داده",
        required=False,
        min_value=1,
        help_text="ردیف ۱ یعنی اولین ردیف بعد از عنوان‌ها.",
    )
    range_end = forms.IntegerField(
        label="تا ردیف داده",
        required=False,
        min_value=1,
    )
    limit = forms.IntegerField(
        label="محدودیت تعداد",
        required=False,
        min_value=1,
        help_text="برای تست مثلاً ۵.",
    )
    sleep_seconds = forms.FloatField(
        label="فاصله بین ارسال‌ها (ثانیه)",
        initial=settings.BALE_DEFAULT_SLEEP_SECONDS,
        min_value=0,
        max_value=60,
        required=False,
    )
    max_retries = forms.IntegerField(
        label="تعداد تلاش مجدد خطاهای موقت",
        initial=2,
        min_value=0,
        max_value=5,
        required=False,
    )
    skip_duplicates = forms.BooleanField(
        label="شماره تکراری داخل فایل ارسال نشود",
        initial=True,
        required=False,
    )

    def __init__(
        self,
        *args,
        validate_send_confirmation: bool = True,
        global_variables: dict | None = None,
        **kwargs,
    ):
        self.validate_send_confirmation = validate_send_confirmation
        self.global_variables = global_variables or {}
        super().__init__(*args, **kwargs)

    def clean_excel_file(self):
        file = self.cleaned_data.get("excel_file")
        if not file:
            return file
        if not file.name.lower().endswith((".xlsx", ".xlsm")):
            raise forms.ValidationError("فقط فایل xlsx یا xlsm پذیرفته می‌شود.")
        max_bytes = settings.BALE_MAX_UPLOAD_SIZE_MB * 1024 * 1024
        if file.size > max_bytes:
            raise forms.ValidationError(
                f"حجم اکسل نباید بیشتر از {settings.BALE_MAX_UPLOAD_SIZE_MB}MB باشد."
            )
        return file

    def clean_attachment(self):
        file = self.cleaned_data.get("attachment")
        if file:
            maximum = getattr(settings, "BALE_MAX_ATTACHMENT_SIZE_MB", 100)
            if file.size > maximum * 1024 * 1024:
                raise forms.ValidationError(
                    f"حجم فایل پیوست نباید بیشتر از {maximum}MB باشد."
                )
        return file

    def clean_variable_mapping(self):
        mapping = _json_object(
            self.cleaned_data.get("variable_mapping"), "نگاشت متغیرها"
        )
        for token, selector in mapping.items():
            if not isinstance(token, str) or not re.fullmatch(r"[^\s{}]+", token):
                raise forms.ValidationError(f"نام متغیر «{token}» معتبر نیست.")
            if token in BUILTIN_TOKENS:
                raise forms.ValidationError(
                    f"متغیر داخلی «{token}» قابل بازنویسی نیست."
                )
            if not isinstance(selector, (str, int)) or not str(selector).strip():
                raise forms.ValidationError(f"ستون متغیر «{token}» مشخص نشده است.")
        return {str(key): str(value) for key, value in mapping.items()}

    def clean_template_fields(self):
        fields = _json_object(self.cleaned_data.get("template_fields"), "فیلدهای قالب")
        if not all(
            isinstance(key, str) and isinstance(value, (str, int, float))
            for key, value in fields.items()
        ):
            raise forms.ValidationError(
                "نام و مقدار فیلدهای قالب باید ساده و متنی باشند."
            )
        return {str(key): str(value) for key, value in fields.items()}

    def clean_buttons_json(self):
        buttons = _json_list(self.cleaned_data.get("buttons_json"), "دکمه‌ها")
        for index, button in enumerate(buttons, start=1):
            if not isinstance(button, dict):
                raise forms.ValidationError(f"دکمه {index} باید object باشد.")
            if button.get("type", "url") not in {"url", "web_app", "copy_text"}:
                raise forms.ValidationError(f"type دکمه {index} معتبر نیست.")
            if (
                not str(button.get("text", "")).strip()
                or not str(button.get("value", "")).strip()
            ):
                raise forms.ValidationError(f"text و value دکمه {index} اجباری است.")
            try:
                row = int(button.get("row", 1))
            except (TypeError, ValueError) as exc:
                raise forms.ValidationError(f"row دکمه {index} باید عدد باشد.") from exc
            if row < 1 or row > 10:
                raise forms.ValidationError(f"row دکمه {index} باید بین ۱ تا ۱۰ باشد.")
            button["row"] = row
        return buttons

    def clean(self):
        data = super().clean()
        data["header_row"] = data.get("header_row") or 1
        data["message_type"] = data.get("message_type") or MessageBatch.MessageType.TEXT
        data["max_retries"] = (
            2 if data.get("max_retries") is None else data["max_retries"]
        )
        if not data.get("excel_file") and not data.get("uploaded_file_token"):
            self.add_error("excel_file", "فایل اکسل را انتخاب کن.")
        if (
            self.validate_send_confirmation
            and data.get("send_mode") == "send"
            and not data.get("confirm_real_send")
        ):
            raise forms.ValidationError("تأیید ارسال واقعی را فعال کن.")
        if (
            data.get("range_start")
            and data.get("range_end")
            and data["range_end"] < data["range_start"]
        ):
            self.add_error("range_end", "پایان بازه باید بزرگ‌تر یا مساوی شروع باشد.")

        message_type = data.get("message_type")
        if message_type == MessageBatch.MessageType.TEXT:
            if (
                not data.get("message_template")
                and not data.get("attachment")
                and not data.get("uploaded_attachment_token")
            ):
                self.add_error(
                    "message_template", "برای این نوع حداقل متن یا فایل لازم است."
                )
        elif message_type == MessageBatch.MessageType.TEMPLATE:
            if not data.get("template_id"):
                self.add_error("template_id", "شناسه قالب اجباری است.")
            if not data.get("template_fields"):
                self.add_error("template_fields", "حداقل یک فیلد قالب تعریف کن.")
            if data.get("attachment") or data.get("uploaded_attachment_token"):
                self.add_error("attachment", "فایل فقط در نوع متن/فایل پشتیبانی می‌شود.")
        elif message_type == MessageBatch.MessageType.OTP:
            if not data.get("otp_template"):
                self.add_error("otp_template", "مقدار OTP اجباری است.")
            if data.get("is_secure"):
                self.add_error(
                    "is_secure", "گزینه رمزدار برای پیام OTP استفاده نمی‌شود."
                )

        available = (
            set(BUILTIN_TOKENS)
            | set(self.global_variables)
            | set(data.get("variable_mapping") or {})
        )
        templates = []
        if data.get("message_template"):
            templates.append(("message_template", data["message_template"]))
        if data.get("otp_template"):
            templates.append(("otp_template", data["otp_template"]))
        if data.get("copy_text_template"):
            templates.append(("copy_text_template", data["copy_text_template"]))
        for value in (data.get("template_fields") or {}).values():
            templates.append(("template_fields", value))
        for button in data.get("buttons_json") or []:
            templates.extend(
                [
                    ("buttons_json", str(button.get("text", ""))),
                    ("buttons_json", str(button.get("value", ""))),
                ]
            )
        for field_name, template in templates:
            try:
                validate_message_template(template, available)
            except ValueError as exc:
                self.add_error(field_name, str(exc))
        return data


class SafirConfigurationForm(forms.ModelForm):
    api_access_key = forms.CharField(
        label="API Access Key",
        required=False,
        widget=forms.PasswordInput(
            render_value=False, attrs={"autocomplete": "new-password"}
        ),
        help_text="اگر خالی بگذاری، کلید فعلی بدون تغییر می‌ماند.",
    )
    global_variables_text = forms.CharField(
        label="متغیرهای ثابت سراسری",
        required=False,
        widget=forms.Textarea(attrs={"rows": 7, "dir": "ltr"}),
        help_text='JSON؛ مثال: {"clinic":"درمانگاه ولیعصر","site":"https://helssa.ir"}',
    )

    class Meta:
        model = SafirConfiguration
        fields = [
            "api_access_key",
            "bot_id",
            "send_url",
            "upload_url",
            "request_timeout",
            "default_sleep_seconds",
            "max_retries",
            "message_price_rial",
        ]
        labels = {
            "bot_id": "شناسه عددی بازو",
            "send_url": "آدرس API ارسال",
            "upload_url": "آدرس API آپلود فایل",
            "request_timeout": "مهلت درخواست (ثانیه)",
            "default_sleep_seconds": "فاصله پیش‌فرض ارسال‌ها",
            "max_retries": "تلاش مجدد پیش‌فرض",
            "message_price_rial": "تعرفه تخمینی هر پیام (ریال)",
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self.instance and self.instance.pk:
            self.fields["global_variables_text"].initial = json.dumps(
                self.instance.global_variables or {}, ensure_ascii=False, indent=2
            )

    def clean_global_variables_text(self):
        variables = _json_object(
            self.cleaned_data.get("global_variables_text"), "متغیرهای سراسری"
        )
        for key, value in variables.items():
            if not isinstance(key, str) or not re.fullmatch(r"[^\s{}]+", key):
                raise forms.ValidationError(f"نام متغیر «{key}» معتبر نیست.")
            if key in BUILTIN_TOKENS:
                raise forms.ValidationError(f"متغیر داخلی «{key}» قابل بازنویسی نیست.")
            if not isinstance(value, (str, int, float, bool)):
                raise forms.ValidationError(f"مقدار «{key}» باید ساده و متنی باشد.")
        return {str(key): str(value) for key, value in variables.items()}

    def clean_request_timeout(self):
        value = self.cleaned_data["request_timeout"]
        if not 1 <= value <= 300:
            raise forms.ValidationError("مهلت درخواست باید بین ۱ تا ۳۰۰ ثانیه باشد.")
        return value

    def clean_default_sleep_seconds(self):
        value = self.cleaned_data["default_sleep_seconds"]
        if not 0 <= value <= 60:
            raise forms.ValidationError("فاصله ارسال باید بین صفر تا ۶۰ ثانیه باشد.")
        return value

    def clean_max_retries(self):
        value = self.cleaned_data["max_retries"]
        if value > 5:
            raise forms.ValidationError("تعداد تلاش مجدد حداکثر ۵ است.")
        return value

    def save(self, commit=True):
        instance = super().save(commit=False)
        new_key = self.cleaned_data.get("api_access_key")
        if not new_key and self.instance.pk:
            instance.api_access_key = SafirConfiguration.objects.get(
                pk=self.instance.pk
            ).api_access_key
        instance.global_variables = self.cleaned_data["global_variables_text"]
        if commit:
            instance.save()
        return instance


class SingleMessageTestForm(forms.Form):
    first_name = forms.CharField(label="نام", max_length=120)
    last_name = forms.CharField(label="نام خانوادگی", max_length=120)
    phone = forms.CharField(label="شماره موبایل", max_length=50)
    message_template = forms.CharField(
        label="متن پیام تست",
        widget=forms.Textarea(attrs={"rows": 6}),
        initial=UploadExcelForm.base_fields["message_template"].initial,
        help_text=UploadExcelForm.base_fields["message_template"].help_text,
    )
    button_enabled = forms.BooleanField(
        label="دکمه لینک‌دار فعال باشد", initial=True, required=False
    )
    button_text = forms.CharField(
        label="متن دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_TEXT
    )
    button_url = forms.URLField(
        label="لینک دکمه", required=False, initial=settings.BALE_DEFAULT_BUTTON_URL
    )

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
        if not normalize_iran_mobile(phone):
            raise forms.ValidationError("شماره موبایل معتبر نیست؛ مثال: 09123456789")
        return phone

    def clean(self):
        data = super().clean()
        if data.get("button_enabled") and not (
            data.get("button_text") and data.get("button_url")
        ):
            raise forms.ValidationError("وقتی دکمه فعال است، متن و لینک آن را وارد کن.")
        return data
