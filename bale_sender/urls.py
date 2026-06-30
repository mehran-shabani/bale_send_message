from django.urls import path
from . import views

urlpatterns = [
    path("", views.dashboard, name="bale_dashboard"),
    path("single-test/", views.single_test, name="bale_single_test"),
    path("batches/", views.batch_list, name="bale_batch_list"),
    path("reports/recent-recipients.xlsx", views.download_recent_recipients_report, name="bale_recent_recipients_report"),
    path("batch/<int:batch_id>/", views.batch_detail, name="bale_batch_detail"),
    path("batch/<int:batch_id>/live-status/", views.batch_live_status, name="bale_batch_live_status"),
    path("batch/<int:batch_id>/cancel/", views.cancel_batch, name="bale_cancel_batch"),
    path("batch/<int:batch_id>/report/", views.download_report, name="bale_download_report"),
]
