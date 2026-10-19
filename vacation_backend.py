"""Хранилище заявок на отпуск: Google Таблица (лист "vacations"), если secrets
настроены, иначе локальный CSV для разработки — как backend.py у формы
пожеланий. Одна строка таблицы = один период отпуска."""
import csv
import datetime
import functools
import os
import threading
import uuid

from secrets_util import has_secret_section
from vacation_logic import STATUS_LABELS, vacation_days

HEADER = ["ID", "ФИО", "Год", "Начало", "Конец", "Дней", "Статус",
          "Комментарий сотрудника", "Комментарий заведующего", "Подано", "Обновлено"]
# Старые названия колонок, которые ещё могут стоять в таблице.
HEADER_ALIASES = {"Комментарий руководителя": "Комментарий заведующего"}

# Все записи — под одной блокировкой: удаление и смена статуса ищут строку по
# номеру, и без блокировки одновременные действия двух людей могли бы попасть
# не в ту строку. Streamlit Cloud — один процесс, этого достаточно. RLock —
# чтобы приложение могло держать её вокруг «проверить и записать».
WRITE_LOCK = threading.RLock()


def _locked(method):
    @functools.wraps(method)
    def wrapper(*args, **kwargs):
        with WRITE_LOCK:
            return method(*args, **kwargs)
    return wrapper
WORKSHEET_NAME = "vacations"
LOCAL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "local_vacations.csv")

# Журнал действий (лист "vacation_log"): кто/когда подал, удалил, сменил статус.
# Входа по паролю нет, поэтому журнал — способ заметить, если кто-то подал или
# удалил заявку за другого.
LOG_HEADER = ["Время", "Год", "ФИО", "Действие", "Период", "Подробности", "Устройство"]
LOG_WORKSHEET_NAME = "vacation_log"
LOCAL_LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "local_vacation_log.csv")


def _parse_date(value):
    """2027-07-01 (как пишет приложение) или 01.07.2027 (если дату поправили
    в таблице руками)."""
    value = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    raise ValueError("непонятная дата «%s»" % value)


_STATUS_FROM_TEXT = dict({code: code for code in STATUS_LABELS},
                         **{label.lower(): code for code, label in STATUS_LABELS.items()})


def _to_record(row):
    start, end = _parse_date(row["Начало"]), _parse_date(row["Конец"])
    status = _STATUS_FROM_TEXT.get(str(row["Статус"]).strip()) or _STATUS_FROM_TEXT.get(str(row["Статус"]).strip().lower())
    if not status:
        raise ValueError("непонятный статус «%s»" % row["Статус"])
    if end < start:
        raise ValueError("конец раньше начала")
    return {
        "id": str(row["ID"]), "name": row["ФИО"], "year": start.year,
        "start": start, "end": end, "days": vacation_days(start, end),
        "status": status, "comment": row["Комментарий сотрудника"],
        "manager_comment": row["Комментарий заведующего"],
        "created": row["Подано"], "updated": row["Обновлено"],
    }


def _parse_rows(rows, year):
    """(заявки за год, [описания строк, которые не удалось прочитать]).
    Испорченная вручную строка не роняет приложение — её пропускаем, а
    заведующий видит предупреждение."""
    records, bad = [], []
    for i, row in enumerate(rows, start=2):  # строка 1 — заголовок
        try:
            rec = _to_record(row)
        except (ValueError, KeyError) as e:
            if str(row.get("Год", "")).strip() in ("", str(year)):
                bad.append("строка %d (%s): %s" % (i, row.get("ФИО", "?") or "?", e))
            continue
        if rec["year"] == year:
            records.append(rec)
    return records, bad


def _rows_from_values(values):
    """get_all_values() → список словарей по заголовкам (со старыми названиями колонок)."""
    if not values:
        return []
    header = [HEADER_ALIASES.get(h, h) for h in values[0]]
    return [dict(zip(header, row + [""] * (len(header) - len(row)))) for row in values[1:] if any(row)]


def _to_row(rec):
    return [rec["id"], rec["name"], rec["year"], rec["start"].isoformat(), rec["end"].isoformat(),
            rec["days"], rec["status"], rec["comment"], rec["manager_comment"],
            rec["created"], rec["updated"]]


def _to_log(row):
    return {"time": row["Время"], "year": int(row["Год"]), "name": row["ФИО"], "action": row["Действие"],
            "period": row["Период"], "details": row["Подробности"], "device": row["Устройство"]}


def new_log(timestamp, year, name, action, period, details="", device=""):
    return {"time": timestamp, "year": year, "name": name, "action": action,
            "period": period, "details": details, "device": device}


def _log_row(e):
    return [e["time"], e["year"], e["name"], e["action"], e["period"], e["details"], e["device"]]


def new_record(name, start, end, comment, status, timestamp):
    return {
        "id": uuid.uuid4().hex, "name": name, "year": start.year, "start": start, "end": end,
        "days": vacation_days(start, end), "status": status, "comment": comment,
        "manager_comment": "", "created": timestamp, "updated": timestamp,
    }


class _LocalBackend:
    def _read(self, path=LOCAL_FILE):
        if not os.path.exists(path):
            return []
        with open(path, newline="", encoding="utf-8") as f:
            return _rows_from_values(list(csv.reader(f)))

    def _write(self, rows, path=LOCAL_FILE, header=HEADER):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=header)
            w.writeheader()
            w.writerows(rows)

    @_locked
    def add_log(self, entries):
        rows = self._read(LOCAL_LOG_FILE)
        rows += [dict(zip(LOG_HEADER, _log_row(e))) for e in entries]
        self._write(rows, LOCAL_LOG_FILE, LOG_HEADER)

    def list_log(self, year):
        return [_to_log(r) for r in self._read(LOCAL_LOG_FILE) if str(r["Год"]) == str(year)]

    def list_vacations(self, year):
        return _parse_rows(self._read(), year)

    @_locked
    def add_vacation(self, rec):
        rows = self._read()
        rows.append(dict(zip(HEADER, _to_row(rec))))
        self._write(rows)

    @_locked
    def delete_vacation(self, rec_id):
        self._write([r for r in self._read() if r["ID"] != rec_id])

    @_locked
    def update_vacations(self, updates, timestamp):
        """updates: {id: {"status": .., "manager_comment": ..}}"""
        rows = self._read()
        for r in rows:
            u = updates.get(r["ID"])
            if u:
                r["Статус"] = u["status"]
                r["Комментарий заведующего"] = u["manager_comment"]
                r["Обновлено"] = timestamp
        self._write(rows)


class _SheetsBackend:
    def __init__(self):
        import gspread
        import streamlit as st
        from google.oauth2.service_account import Credentials

        @st.cache_resource
        def _client():
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=["https://www.googleapis.com/auth/spreadsheets"],
            )
            return gspread.authorize(creds)

        self._gspread = gspread
        self._client = _client
        self._spreadsheet_id = st.secrets["sheet"]["spreadsheet_id"]

    def _ws(self, title=WORKSHEET_NAME, header=HEADER):
        sh = self._client().open_by_key(self._spreadsheet_id)
        try:
            return sh.worksheet(title)
        except self._gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=title, rows=500, cols=len(header))
            ws.append_row(header, value_input_option="RAW")
            return ws

    @_locked
    def add_log(self, entries):
        if entries:
            self._ws(LOG_WORKSHEET_NAME, LOG_HEADER).append_rows(
                [_log_row(e) for e in entries], value_input_option="RAW")

    def list_log(self, year):
        ws = self._ws(LOG_WORKSHEET_NAME, LOG_HEADER)
        rows = ws.get_all_records(expected_headers=LOG_HEADER, numericise_ignore=["all"])
        return [_to_log(r) for r in rows if str(r["Год"]) == str(year)]

    def list_vacations(self, year):
        return _parse_rows(_rows_from_values(self._ws().get_all_values()), year)

    @_locked
    def add_vacation(self, rec):
        self._ws().append_row(_to_row(rec), value_input_option="RAW")

    @_locked
    def delete_vacation(self, rec_id):
        ws = self._ws()
        ids = ws.col_values(1)
        if rec_id in ids:
            ws.delete_rows(ids.index(rec_id) + 1)

    @_locked
    def update_vacations(self, updates, timestamp):
        ws = self._ws()
        ids = ws.col_values(1)
        a1 = self._gspread.utils.rowcol_to_a1
        cols = {f: HEADER.index(f) + 1 for f in ("Статус", "Комментарий заведующего", "Обновлено")}
        batch = []
        for rec_id, u in updates.items():
            if rec_id not in ids:
                continue
            row = ids.index(rec_id) + 1
            for field, value in (("Статус", u["status"]),
                                 ("Комментарий заведующего", u["manager_comment"]),
                                 ("Обновлено", timestamp)):
                batch.append({"range": a1(row, cols[field]), "values": [[value]]})
        if batch:
            ws.batch_update(batch, value_input_option="RAW")


# VACATION_LOCAL_BACKEND=1 — принудительно локальный CSV (для тестов, чтобы не
# писать в настоящую таблицу, даже если .streamlit/secrets.toml лежит рядом).
if (not os.environ.get("VACATION_LOCAL_BACKEND")
        and has_secret_section("gcp_service_account") and has_secret_section("sheet")):
    backend = _SheetsBackend()
    BACKEND_NAME = "google_sheets"
else:
    backend = _LocalBackend()
    BACKEND_NAME = "local_csv"
