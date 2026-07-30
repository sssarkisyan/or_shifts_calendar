"""Локальное CSV-хранилище — используется только когда Google-таблица не
настроена (secrets отсутствуют). Удобно для локальной проверки приложения
без создания сервис-аккаунта Google. В реальном деплое не используется."""
import csv
import json
import os

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "local_data.csv")
HEADER = ["ФИО", "Год", "Месяц", "Статусы", "Комментарий", "Отправлено"]


def _read_all():
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _write_all(rows):
    with open(DATA_FILE, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HEADER)
        w.writeheader()
        for r in rows:
            w.writerow(r)


def get_submission(name, year, month):
    for r in _read_all():
        if r["ФИО"] == name and str(r["Год"]) == str(year) and str(r["Месяц"]) == str(month):
            return {"statuses": json.loads(r["Статусы"]) if r["Статусы"] else {}, "comment": r["Комментарий"]}
    return None


def upsert_submission(name, year, month, statuses, comment, timestamp):
    rows = _read_all()
    found = False
    for r in rows:
        if r["ФИО"] == name and str(r["Год"]) == str(year) and str(r["Месяц"]) == str(month):
            r["Статусы"] = json.dumps(statuses, ensure_ascii=False)
            r["Комментарий"] = comment
            r["Отправлено"] = timestamp
            found = True
            break
    if not found:
        rows.append({
            "ФИО": name, "Год": year, "Месяц": month,
            "Статусы": json.dumps(statuses, ensure_ascii=False),
            "Комментарий": comment, "Отправлено": timestamp,
        })
    _write_all(rows)


def list_submissions(year, month):
    out = []
    for r in _read_all():
        if str(r["Год"]) == str(year) and str(r["Месяц"]) == str(month):
            out.append({
                "name": r["ФИО"],
                "statuses": json.loads(r["Статусы"]) if r["Статусы"] else {},
                "comment": r["Комментарий"],
                "submitted_at": r["Отправлено"],
            })
    return out
