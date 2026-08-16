import calendar as calmod
import datetime
import importlib
import io

import pandas as pd
import streamlit as st

from secrets_util import get_secret

# Один код, два деплоя: значение "department" в Secrets конкретного деплоя
# выбирает ростер/лимиты (config.py — врачи, config_nurses.py — медсёстры).
# Google-таблица и код доступа тоже свои у каждого деплоя (свои Secrets),
# этот выбор их не касается.
DEPARTMENT = get_secret("department", "doctors")
config = importlib.import_module("config_nurses" if DEPARTMENT == "nurses" else "config")
DEPARTMENT_TITLE = "Медсестры" if DEPARTMENT == "nurses" else "Врачи"

from backend import BACKEND_NAME, get_submission, list_submissions, upsert_submission

st.set_page_config(page_title=f"Пожелания по графику | {DEPARTMENT_TITLE}", page_icon="🩺", layout="centered")

if DEPARTMENT == "nurses":
    STATUSES = [
        {"code": "CAN", "label": "Могу"},
        {"code": "CAN_11", "label": "Могу только с 11:30"},
        {"code": "CANNOT", "label": "Не могу"},
        {"code": "OTHER_JOB", "label": "Работаю в другом месте"},
        {"code": "VACATION", "label": "В отпуске"},
    ]
else:
    STATUSES = [
        {"code": "CAN", "label": "Могу"},
        {"code": "CAN_8", "label": "Могу только 08:00-16:00"},
        {"code": "CAN_11", "label": "Могу только 11:00-21:00"},
        {"code": "CANNOT", "label": "Не могу"},
        {"code": "OTHER_JOB", "label": "Работаю в другом месте"},
        {"code": "VACATION", "label": "В отпуске"},
    ]
CODE_TO_LABEL = {s["code"]: s["label"] for s in STATUSES}
LABEL_TO_CODE = {s["label"]: s["code"] for s in STATUSES}
LABELS = [s["label"] for s in STATUSES]
AVAILABLE_CODES = {"CAN", "CAN_8", "CAN_11"}
MONTH_NAMES = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
               "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
DOW_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]


def days_in_month(year, month):
    return calmod.monthrange(year, month)[1]


def default_year_month():
    today = datetime.date.today()
    y, m = today.year, today.month + 1
    if m > 12:
        m -= 12
        y += 1
    return y, m


def load_person_month(name, year, month):
    st.session_state.current_name = name
    st.session_state.current_year = year
    st.session_state.current_month = month
    existing = get_submission(name, year, month)
    n = days_in_month(year, month)
    statuses = {"%04d-%02d-%02d" % (year, month, d): "CAN" for d in range(1, n + 1)}
    comment = ""
    if existing:
        for k, v in existing["statuses"].items():
            if k in statuses and v in CODE_TO_LABEL:
                statuses[k] = v
        comment = existing.get("comment", "")
    st.session_state.sub_statuses = statuses
    st.session_state.sub_comment = comment
    for iso, code in statuses.items():
        st.session_state["day_" + iso] = CODE_TO_LABEL[code]
    st.session_state.comment_box = comment


def build_availability_csv(rows_for_export, dates):
    """rows_for_export: list of {"name":.., "statuses": {iso: code}, "comment":..}"""
    header = ["ФИО"] + dates + ["Комментарий"]
    lines = [header]
    for r in rows_for_export:
        lines.append([r["name"]] + [r["statuses"].get(d, "CAN") for d in dates] + [r["comment"]])
    buf = io.StringIO()
    import csv as csv_mod
    w = csv_mod.writer(buf)
    w.writerows(lines)
    return ("﻿" + buf.getvalue()).encode("utf-8")


def build_staff_csv():
    buf = io.StringIO()
    import csv as csv_mod
    w = csv_mod.writer(buf)
    w.writerow(["ФИО", "Ставка", "Тип"])
    for s in config.STAFF:
        w.writerow([s["name"], s["rate"], s["kind"]])
    return ("﻿" + buf.getvalue()).encode("utf-8")


st.title(f"Пожелания по графику | {DEPARTMENT_TITLE}")
if BACKEND_NAME == "local_csv":
    st.error(
        "⚠️ ЛОКАЛЬНЫЙ РЕЖИМ: Google-secrets не обнаружены, данные сохраняются "
        "только во временный файл этого запуска и НЕ попадают в Google Таблицу. "
        "Если это задеплоенное приложение — проверьте Secrets в настройках "
        "приложения на share.streamlit.io и перезапустите его (Reboot app)."
    )
else:
    st.caption("✅ Подключено к Google Таблице.")

tab_me, tab_manager = st.tabs(["Мои пожелания", "Свод по отделению"])

with tab_me:
    names = [s["name"] for s in config.STAFF]
    y0, m0 = default_year_month()

    if "current_name" not in st.session_state:
        load_person_month(names[0], y0, m0)

    col1, col2, col3 = st.columns([2, 1, 1])
    name = col1.selectbox("ФИО", names,
                           index=names.index(st.session_state.current_name) if st.session_state.current_name in names else 0)
    month = col2.selectbox("Месяц", list(range(1, 13)), index=st.session_state.current_month - 1,
                            format_func=lambda i: MONTH_NAMES[i - 1])
    year_options = sorted(set([y0, y0 + 1, st.session_state.current_year]))
    year = col3.selectbox("Год", year_options,
                           index=year_options.index(st.session_state.current_year))

    if (name, year, month) != (st.session_state.current_name, st.session_state.current_year, st.session_state.current_month):
        load_person_month(name, year, month)
        st.rerun()

    n_days = days_in_month(year, month)

    st.markdown(
        " ".join(
            "`%s`" % s["label"] for s in STATUSES
        )
    )

    for d in range(1, n_days + 1):
        iso = "%04d-%02d-%02d" % (year, month, d)
        key = "day_" + iso
        prev_code = st.session_state.sub_statuses[iso]
        current_label = st.session_state.get(key, CODE_TO_LABEL[prev_code])
        current_code = LABEL_TO_CODE.get(current_label, prev_code)
        if current_code != prev_code:
            if current_code == "CANNOT":
                other_cannot = sum(
                    1 for k, v in st.session_state.sub_statuses.items() if k != iso and v == "CANNOT"
                )
                if other_cannot >= config.MAX_CANNOT_DAYS:
                    st.session_state[key] = CODE_TO_LABEL[prev_code]
                    st.toast("Лимит %d дней «Не могу» в месяц достигнут." % config.MAX_CANNOT_DAYS, icon="⚠️")
                else:
                    st.session_state.sub_statuses[iso] = current_code
            else:
                st.session_state.sub_statuses[iso] = current_code

        dow = datetime.date(year, month, d).weekday()
        c1, c2 = st.columns([1, 4])
        c1.markdown("**%d** %s" % (d, DOW_RU[dow]))
        c2.selectbox(" ", LABELS, key=key, label_visibility="collapsed")

    cannot_count = sum(1 for v in st.session_state.sub_statuses.values() if v == "CANNOT")
    if cannot_count >= config.MAX_CANNOT_DAYS:
        st.warning("Дней «Не могу»: %d / %d" % (cannot_count, config.MAX_CANNOT_DAYS))
    else:
        st.info("Дней «Не могу»: %d / %d" % (cannot_count, config.MAX_CANNOT_DAYS))

    comment = st.text_area(
        "Комментарий / пожелания (необязательно)",
        key="comment_box",
        placeholder="Например: предпочитаю эндоскопию, не больше 3 суточных в месяц...",
    )
    st.session_state.sub_comment = comment

    if st.button("Отправить пожелания", type="primary", width="stretch"):
        upsert_submission(
            name, year, month, st.session_state.sub_statuses, st.session_state.sub_comment,
            datetime.datetime.now().isoformat(timespec="seconds"),
        )
        st.success("Пожелания сохранены: %s, %s %d." % (name, MONTH_NAMES[month - 1], year))

with tab_manager:
    if "manager_unlocked" not in st.session_state:
        st.session_state.manager_unlocked = False

    if not st.session_state.manager_unlocked:
        st.write("Здесь собраны ответы всех сотрудников — доступ только для составителя графика.")
        pw = st.text_input("Код доступа", type="password", key="mgr_pw_input")
        if st.button("Войти", key="mgr_login_btn"):
            expected = get_secret("manager_passcode", config.MANAGER_PASSCODE_FALLBACK)
            if pw == expected:
                st.session_state.manager_unlocked = True
                st.rerun()
            else:
                st.error("Неверный код доступа.")
    else:
        y0, m0 = default_year_month()
        col1, col2, col3 = st.columns([1, 1, 1])
        m_month = col1.selectbox("Месяц", list(range(1, 13)), index=m0 - 1,
                                  format_func=lambda i: MONTH_NAMES[i - 1], key="mgr_month")
        m_year = col2.selectbox("Год", [y0 - 1, y0, y0 + 1], index=1, key="mgr_year")
        if col3.button("Выйти"):
            st.session_state.manager_unlocked = False
            st.rerun()

        n_days = days_in_month(m_year, m_month)
        dates = ["%04d-%02d-%02d" % (m_year, m_month, d) for d in range(1, n_days + 1)]
        rows = list_submissions(m_year, m_month)

        submitted_names = {r["name"] for r in rows}
        missing = [s["name"] for s in config.STAFF if s["name"] not in submitted_names]
        if missing:
            st.warning("Ещё не ответили (%d): %s" % (len(missing), ", ".join(missing)))

        if rows:
            table = {"ФИО": []}
            for d in dates:
                table[d[8:10]] = []
            table["Комментарий"] = []
            for r in rows:
                table["ФИО"].append(r["name"])
                for d in dates:
                    code = r["statuses"].get(d, "")
                    table[d[8:10]].append(CODE_TO_LABEL.get(code, "—"))
                table["Комментарий"].append(r["comment"])
            counts_row = {"ФИО": "Доступно на день"}
            for d in dates:
                n_avail = sum(1 for r in rows if r["statuses"].get(d) in AVAILABLE_CODES)
                counts_row[d[8:10]] = n_avail
            counts_row["Комментарий"] = ""

            df = pd.DataFrame(table)
            st.dataframe(df, width="stretch", hide_index=True)
            available_labels = " / ".join(CODE_TO_LABEL[c] for c in AVAILABLE_CODES if c in CODE_TO_LABEL)
            st.caption(
                "Доступно на день (%s): " % available_labels
                + ", ".join("%s=%s" % (d[8:10], counts_row[d[8:10]]) for d in dates)
            )

            csv_bytes = build_availability_csv(rows, dates)
            st.download_button(
                "Скачать свод пожеланий (CSV)", data=csv_bytes,
                file_name="orit_svod_%04d-%02d.csv" % (m_year, m_month), mime="text/csv",
            )
        else:
            st.info("Пока никто не отправил пожелания на этот месяц.")

        st.download_button(
            "Скачать список сотрудников (CSV)", data=build_staff_csv(),
            file_name="staff.csv", mime="text/csv",
        )
