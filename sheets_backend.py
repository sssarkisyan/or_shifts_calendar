"""Хранилище пожеланий в Google Sheets — реальный backend для деплоя.
Требует настроенных secrets: [gcp_service_account] и [sheet].spreadsheet_id
(см. secrets.toml.example и DEPLOY.md)."""
import json

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
HEADER = ["ФИО", "Год", "Месяц", "Статусы", "Комментарий", "Отправлено"]
WORKSHEET_NAME = "responses"


@st.cache_resource
def _client():
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return gspread.authorize(creds)


def _worksheet():
    sh = _client().open_by_key(st.secrets["sheet"]["spreadsheet_id"])
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=500, cols=len(HEADER))
        ws.append_row(HEADER)
    return ws


def _records(ws):
    return ws.get_all_records(expected_headers=HEADER)


def get_submission(name, year, month):
    ws = _worksheet()
    for r in _records(ws):
        if r["ФИО"] == name and str(r["Год"]) == str(year) and str(r["Месяц"]) == str(month):
            return {"statuses": json.loads(r["Статусы"]) if r["Статусы"] else {}, "comment": r["Комментарий"]}
    return None


def upsert_submission(name, year, month, statuses, comment, timestamp):
    ws = _worksheet()
    records = _records(ws)
    target_row = None
    for i, r in enumerate(records):
        if r["ФИО"] == name and str(r["Год"]) == str(year) and str(r["Месяц"]) == str(month):
            target_row = i + 2  # +1 header row, +1 for 1-indexing
            break
    row_values = [name, year, month, json.dumps(statuses, ensure_ascii=False), comment, timestamp]
    if target_row:
        ws.update(values=[row_values], range_name="A%d:F%d" % (target_row, target_row))
    else:
        ws.append_row(row_values)


def list_submissions(year, month):
    ws = _worksheet()
    out = []
    for r in _records(ws):
        if str(r["Год"]) == str(year) and str(r["Месяц"]) == str(month):
            out.append({
                "name": r["ФИО"],
                "statuses": json.loads(r["Статусы"]) if r["Статусы"] else {},
                "comment": r["Комментарий"],
                "submitted_at": r["Отправлено"],
            })
    return out
