"""Выбирает реальный backend (Google Sheets), если secrets настроены,
иначе локальный CSV-fallback для разработки/проверки без Google-аккаунта."""
from secrets_util import has_secret_section

if has_secret_section("gcp_service_account") and has_secret_section("sheet"):
    from sheets_backend import get_submission, upsert_submission, list_submissions
    BACKEND_NAME = "google_sheets"
else:
    from local_backend import get_submission, upsert_submission, list_submissions
    BACKEND_NAME = "local_csv"

__all__ = ["get_submission", "upsert_submission", "list_submissions", "BACKEND_NAME"]
