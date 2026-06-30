from pathlib import Path
import os
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

SECRET_KEY = os.getenv("DJANGO_SECRET_KEY", "dev-only-secret-key")
DEBUG = os.getenv("DJANGO_DEBUG", "True").lower() == "true"
ALLOWED_HOSTS = [h.strip() for h in os.getenv("DJANGO_ALLOWED_HOSTS", "*").split(",") if h.strip()]

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "bale_sender",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "bale_safir_project.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    }
]

WSGI_APPLICATION = "bale_safir_project.wsgi.application"

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
    }
}

AUTH_PASSWORD_VALIDATORS = []
LANGUAGE_CODE = "fa-ir"
TIME_ZONE = "Asia/Tehran"
USE_I18N = True
USE_TZ = True
STATIC_URL = "static/"
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

BALE_SEND_URL = os.getenv("BALE_SEND_URL", "https://safir.bale.ai/api/v3/send_message")
BALE_API_ACCESS_KEY = os.getenv("BALE_API_ACCESS_KEY", "")
BALE_BOT_ID = int(os.getenv("BALE_BOT_ID", "0") or 0)
BALE_REQUEST_TIMEOUT = int(os.getenv("BALE_REQUEST_TIMEOUT", "20") or 20)
BALE_DEFAULT_SLEEP_SECONDS = float(os.getenv("BALE_DEFAULT_SLEEP_SECONDS", "0.4") or 0.4)
BALE_DEFAULT_BUTTON_TEXT = os.getenv("BALE_DEFAULT_BUTTON_TEXT", "ثبت‌نام در سایت")
BALE_DEFAULT_BUTTON_URL = os.getenv("BALE_DEFAULT_BUTTON_URL", "https://helssa.ir")

BALE_MAX_UPLOAD_SIZE_MB = int(os.getenv("BALE_MAX_UPLOAD_SIZE_MB", "10") or 10)
