"""Портал заявок на отпуск (врачи): сотрудник планирует периоды отпуска на год,
руководитель согласовывает их и выгружает годовой график в Excel.
Отдельный деплой из того же репозитория — главный файл vacation_app.py
(см. VACATION_DEPLOY.md)."""
import calendar as calmod
import datetime
import functools
import os

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import config
import vacation_config as vc
from secrets_util import get_secret
from vacation_backend import BACKEND_NAME, backend, new_log, new_record
from vacation_logic import (
    APPROVED, DOW_RU, MONTH_NAMES, PENDING, REJECTED, STATUS_LABELS, TeamRules,
    active, approved, approved_violations, bad_days, build_excel, collapse_dates, daterange, editor_updates,
    fmt_date, fmt_date_words, fmt_period, fmt_period_words, fmt_ranges, is_holiday, n_days, n_doctors, person_summary, record_conflicts, team_calendar,
    team_conflicts, vacation_days, validate_new_period,
)

st.set_page_config(page_title="График отпусков | Врачи", page_icon="🏖️", layout="centered")

STAFF = config.STAFF
NAMES = [s["name"] for s in STAFF]
KIND = {s["name"]: s["kind"] for s in STAFF}
# Правила пересечения отпусков основных сотрудников (совместители не ограничены).
RULES = TeamRules(
    main_names=frozenset(s["name"] for s in STAFF if s["kind"] in vc.OVERLAP_KINDS),
    max_overlap=vc.MAX_OVERLAP_DAYS,
    max_concurrent=vc.MAX_CONCURRENT_MAIN,
)
STATUS_ICON = {PENDING: "🕓", APPROVED: "✅", REJECTED: "❌"}


MSK = datetime.timezone(datetime.timedelta(hours=3))  # сервер Streamlit Cloud живёт в UTC


def now():
    return datetime.datetime.now(MSK).replace(tzinfo=None).isoformat(timespec="seconds")


# Чтение из Google Таблицы кэшируется на минуту: все пользователи ходят через
# один сервис-аккаунт, а у Google лимит ~60 чтений в минуту. После любой
# записи кэш сбрасывается, проверка перед подачей всегда читает свежие данные.
@st.cache_data(ttl=60, show_spinner=False)
def load_vacations(year):
    return backend.list_vacations(year)


@st.cache_data(ttl=60, show_spinner=False)
def load_log(year):
    return backend.list_log(year)


def write_log(*entries):
    backend.add_log(list(entries))
    load_vacations.clear()
    load_log.clear()


def device():
    """Короткое описание устройства по User-Agent — для журнала."""
    try:
        ua = st.context.headers.get("User-Agent", "") or ""
    except Exception:
        return ""
    os_name = next((label for key, label in (("iPhone", "iPhone"), ("iPad", "iPad"), ("Android", "Android"),
                                             ("Windows", "Windows"), ("Mac OS", "Mac"), ("Linux", "Linux"))
                    if key in ua), "")
    browser = next((label for key, label in (("YaBrowser", "Яндекс"), ("Edg/", "Edge"), ("OPR/", "Opera"),
                                             ("Firefox", "Firefox"), ("Chrome", "Chrome"), ("Safari", "Safari"))
                    if key in ua), "")
    return ", ".join(x for x in (os_name, browser) if x)


def fmt_time(ts):
    try:
        return datetime.datetime.fromisoformat(ts).strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return ts


def flash(kind, text):
    st.session_state.flash = (kind, text)


def show_flash():
    if "flash" in st.session_state:
        kind, text = st.session_state.pop("flash")
        getattr(st, kind)(text)


def submissions_locked():
    return vc.SUBMISSION_DEADLINE is not None and datetime.date.today() > vc.SUBMISSION_DEADLINE


# Годовой календарь — свой компонент Streamlit (vacation_calendar/index.html):
# у сотрудника по нему кликают, выбирая даты отпуска (первый клик — начало,
# второй — конец), у руководителя он только для просмотра.
vacation_calendar = components.declare_component(
    "vacation_calendar", path=os.path.join(os.path.dirname(os.path.abspath(__file__)), "vacation_calendar"))

# Бледно-розовый — подано, бледно-зелёный — одобрено (те же цвета в Excel).
C_PENDING, C_APPROVED, C_FULL, C_BAD = "#F9D5DC", "#CDEBD3", "#BDBDBD", "#FF7C80"


def visible_colleagues(me):
    """Чьи отпуска сотрудник видит в календаре — коллеги из его группы
    CALENDAR_GROUPS. Кто ни в одну группу не входит (совместители,
    договорники) — ничьих: ограничений по датам у них нет, чужие отпуска им
    не нужны."""
    groups = [g for g in vc.CALENDAR_GROUPS if KIND[me] in g]
    kinds = set().union(*groups) if groups else set()
    return {n for n, k in KIND.items() if k in kinds and n != me}


def calendar_data(year, recs, full=(), bad=(), mine=(), label=lambda n: n):
    """Данные для компонента календаря: (days, reqs).
    days — {iso: {"c": CSS-классы, "r": [id заявок в этот день]}}: заливка по
    плотности (vc-p1..3 — сколько коллег, если есть поданные заявки; vc-a1..3 —
    если все отпуска в этот день одобрены), full — дни, закрытые для основных
    (серые), bad — нарушения среди одобренных (красные), mine — свои отпуска.
    reqs — {id: {"s": начало, "e": конец, "t": текст}}: по наведению/касанию
    дня календарь обводит отпуска этих коллег целиком и показывает подсказку."""
    reqs, by_day = {}, {}
    for i, r in enumerate(sorted(recs, key=lambda r: (r["start"], r["name"]))):
        rid = str(i)
        reqs[rid] = {"s": r["start"].isoformat(), "e": r["end"].isoformat(),
                     "t": "%s — %s, %s" % (label(r["name"]), "одобрено" if r["status"] == APPROVED else "заявка подана",
                                          fmt_period(r["start"], r["end"]))}
        for d in daterange(max(r["start"], datetime.date(year, 1, 1)), min(r["end"], datetime.date(year, 12, 31))):
            by_day.setdefault(d, []).append((rid, r["status"]))
    days = {}
    for d in daterange(datetime.date(year, 1, 1), datetime.date(year, 12, 31)):
        day = by_day.get(d, [])
        cls = []
        if d in bad:
            cls.append("vc-bad")
        elif d in full:
            cls.append("vc-full")
        elif day:
            cls.append("vc-%s%d" % ("a" if all(st == APPROVED for _, st in day) else "p", min(len(day), 3)))
        if is_holiday(d):
            cls.append("vc-hol")
        elif d.weekday() >= 5:
            cls.append("vc-wk")
        if d in mine:
            cls.append("vc-mine")
        if cls or day:
            days[d.isoformat()] = {"c": " ".join(cls), "r": [rid for rid, _ in day]}
    return days, reqs


def to_n_doctors(n):
    """Дательный падеж: «1 основному врачу», «2 основным врачам»."""
    return "%d %s" % (n, "основному врачу" if n % 10 == 1 and n % 100 != 11 else "основным врачам")


# Заливка по плотности: 1 / 2 / 3+ коллег (розовая — есть поданные заявки,
# зелёная — все отпуска в этот день одобрены). Те же цвета — в index.html.
DENSITY_PENDING = ["#FBE3E8", "#F5B9C6", "#EC8CA2"]
DENSITY_APPROVED = ["#E0F2E5", "#B5DFC0", "#7FC492"]


def legend(full=False, bad=False, mine=False, colleagues=True, others=True):
    """[[стиль или список стилей образцов, текст], ...]; others=False — чужих
    отпусков в календаре нет, легенда только про свои."""
    who = "коллег" if colleagues else "сотрудников"
    items = []
    if mine:
        items.append(["box-shadow:inset 0 0 0 2px #2F6FDE", "мои отпуска"])
    if not others:
        return items
    items += [[["background:%s" % c for c in DENSITY_PENDING], "1 · 2 · 3+ %s, есть поданные заявки" % who],
              [["background:%s" % c for c in DENSITY_APPROVED], "1 · 2 · 3+ %s, всё одобрено" % who]]
    if full:
        items.append(["background:%s" % C_FULL, "недоступно (одобрено %s)" % to_n_doctors(vc.MAX_CONCURRENT_MAIN)])
    if bad:
        items.append(["background:%s" % C_BAD, "нарушение среди одобренных"])
    return items


def closed_days(records, exclude_name=None):
    """Дни, где отпуск уже одобрен MAX_CONCURRENT_MAIN основным (кроме exclude_name)."""
    main = team_calendar(records, RULES.main_names, exclude_name=exclude_name)
    return {d for d, info in main.items() if len(info["approved"]) >= RULES.max_concurrent}


def employee_calendar_args(year, records, me, selected):
    """Аргументы кликабельного календаря сотрудника: коллеги из его группы
    CALENDAR_GROUPS, свои отпуска, закрытые дни (для основных) — их нельзя
    выбрать началом или концом отпуска."""
    is_main = me in RULES.main_names
    full = closed_days(records, exclude_name=me) if is_main else set()
    colleagues = visible_colleagues(me)
    days, reqs = calendar_data(
        year, [r for r in active(records) if r["name"] in colleagues], full=full,
        mine={d for r in active(records) if r["name"] == me for d in daterange(r["start"], r["end"])})
    return dict(year=year, selectable=True, selected=selected, blocked=sorted(d.isoformat() for d in full),
                days=days, reqs=reqs, legend=legend(full=is_main, mine=True, others=bool(colleagues)))


def manager_calendar_args(year, records):
    """Календарь руководителя (только просмотр): все сотрудники, закрытые дни, нарушения."""
    days, reqs = calendar_data(
        year, active(records), full=closed_days(records), bad=bad_days(records, RULES),
        label=lambda n: n if n in RULES.main_names else "%s (%s)" % (n, KIND.get(n, "?")))
    return dict(year=year, selectable=False, selected=None, blocked=[], days=days, reqs=reqs,
                legend=legend(full=True, bad=True, colleagues=False))


st.title("График отпусков | Врачи")
if BACKEND_NAME == "local_csv":
    st.error(
        "⚠️ ЛОКАЛЬНЫЙ РЕЖИМ: Google-secrets не обнаружены, заявки сохраняются "
        "только во временный файл этого запуска и НЕ попадают в Google Таблицу. "
        "Если это задеплоенное приложение — проверьте Secrets в настройках "
        "приложения на share.streamlit.io и перезапустите его (Reboot app)."
    )
else:
    st.caption("✅ Подключено к Google Таблице.")

tab_me, tab_manager = st.tabs(["Мои отпуска", "Свод по отделению"])

with tab_me:
    col1, col2 = st.columns([3, 1])
    name = col1.selectbox("ФИО", NAMES, key="me_name")
    year = col2.selectbox("Год", [vc.PLAN_YEAR - 1, vc.PLAN_YEAR, vc.PLAN_YEAR + 1], index=1, key="me_year")

    records = load_vacations(year)
    mine = sorted((r for r in records if r["name"] == name), key=lambda r: r["start"])
    s = person_summary(name, records, vc.entitlement(name), vc.MIN_LONG_PART)
    locked = submissions_locked()

    show_flash()

    m1, m2, m3 = st.columns(3)
    m1.metric("Положено дней", s["entitlement"])
    m2.metric("Запланировано", s["planned"])
    m3.metric("Осталось", s["entitlement"] - s["planned"])
    if vc.carryover(name):
        st.caption("Положено: %s за год + %s остатка за прошлый год." % (n_days(vc.BASE_DAYS), n_days(vc.carryover(name))))
    if mine and not s["has_long"]:
        st.warning("Пока нет ни одной части отпуска от %s — она обязательна (ст. 125 ТК РФ)." % n_days(vc.MIN_LONG_PART))

    st.subheader("Мои отпуска")
    if not mine:
        st.info("На %d год отпусков пока нет." % year)
    for r in mine:
        with st.container(border=True):
            c1, c2 = st.columns([5, 1])
            c1.markdown("**%s** · %s · %s %s" % (
                fmt_period(r["start"], r["end"]), n_days(r["days"]), STATUS_ICON[r["status"]], STATUS_LABELS[r["status"]]))
            if r["comment"]:
                c1.caption("Ваш комментарий: " + r["comment"])
            if r["manager_comment"]:
                c1.caption("Комментарий руководителя: " + r["manager_comment"])
            if r["status"] == PENDING:
                others_appr = [o for o in approved(records) if o["name"] != name]
                hard = team_conflicts(name, r["start"], r["end"], others_appr, RULES)
                soft = [] if hard else record_conflicts(r, records, RULES)
                for c in hard:
                    c1.warning("Уже одобрено коллегам: %s. Руководитель, скорее всего, попросит перенести даты." % c)
                for c in soft:
                    c1.caption("⚠️ Совпадает с заявками коллег: %s — решит руководитель." % c)
            if r["status"] == PENDING and not locked:
                with c2.popover("Удалить"):
                    st.write("Удалить заявку **%s** (%s)?" % (fmt_period(r["start"], r["end"]), name))
                    if st.button("Да, удалить", key="del_" + r["id"], type="primary"):
                        backend.delete_vacation(r["id"])
                        write_log(new_log(now(), year, name, "Заявка удалена", fmt_period(r["start"], r["end"]),
                                          device=device()))
                        flash("success", "Отпуск %s удалён." % fmt_period(r["start"], r["end"]))
                        st.rerun()
    if any(r["status"] != PENDING for r in mine):
        st.caption("Одобренный или отклонённый отпуск изменить самостоятельно нельзя — обратитесь к руководителю.")

    my_log = sorted((e for e in load_log(year) if e["name"] == name), key=lambda e: e["time"], reverse=True)
    if my_log:
        with st.expander("История действий с моими заявками (%d)" % len(my_log)):
            st.caption("Если видите действие, которого не совершали, — сообщите руководителю.")
            for e in my_log:
                st.markdown("- %s — %s %s%s%s" % (
                    fmt_time(e["time"]), e["action"].lower(), e["period"],
                    (" — " + e["details"]) if e["details"] else "",
                    (" · " + e["device"]) if e["device"] else ""))

    st.subheader("Добавить отпуск")
    if locked:
        st.warning("Приём заявок закрыт %s. Изменения — через руководителя." % fmt_date(vc.SUBMISSION_DEADLINE))
    else:
        if name in RULES.main_names:
            st.caption(
                "Подать заявку можно на любые даты, кто подал раньше — значения не имеет, решает руководитель. "
                "Недоступны только дни, где отпуск уже одобрен %s; с каждым одобренным отпуском основного "
                "сотрудника можно пересекаться не больше чем на %s."
                % (to_n_doctors(vc.MAX_CONCURRENT_MAIN), n_days(vc.MAX_OVERLAP_DAYS)))
        else:
            st.caption("Подать заявку можно на любые даты. Ограничения на пересечение с отпусками коллег "
                       "на вас не распространяются.")

        # Выбор дат — кликом по календарю ниже. Значение компонента копируется в
        # sel_key (on_change), форма над календарём читает его оттуда.
        form_n = st.session_state.get("form_n", 0)
        pick_n = st.session_state.get("pick_n", 0)
        sel_key = "sel_%d_%d_%s_%d" % (form_n, pick_n, name, year)
        cal_key = "cal_" + sel_key
        st.session_state.cur_sel_key = sel_key  # для test_vacation_flow.py
        sel = st.session_state.get(sel_key) or {}
        start = datetime.date.fromisoformat(sel["start"]) if sel.get("start") else None
        end = datetime.date.fromisoformat(sel["end"]) if sel.get("end") else None

        st.caption("Кликните в календаре на первый, а затем на последний день отпуска."
                   + (" Наведение курсора на конкретный день — покажет, кто из коллег в отпуске "
                      "(на телефоне — коснитесь дня)." if visible_colleagues(name) else ""))
        vacation_calendar(
            **employee_calendar_args(year, records, name, {"start": sel.get("start"), "end": sel.get("end")}),
            key=cal_key, default=None,
            on_change=functools.partial(lambda sk, ck: st.session_state.__setitem__(sk, st.session_state.get(ck)),
                                        sel_key, cal_key),
        )

        errors = []
        if start and end:
            days = vacation_days(start, end)
            n_hol = (end - start).days + 1 - days
            st.markdown("**Выбрано:** %s — %s отпуска%s" % (
                fmt_period_words(start, end), n_days(days),
                " (праздничные дни не входят: %s)" % n_days(n_hol) if n_hol else ""))
            errors, warnings = validate_new_period(name, start, end, year, records, vc.entitlement(name),
                                                   RULES, vc.MIN_LONG_PART)
            for e in errors:
                st.error(e)
            for w in warnings:
                st.warning(w)
        elif start:
            st.info("Начало отпуска — %s. Теперь выберите в календаре последний день." % fmt_date_words(start))
        else:
            st.info("Выберите в календаре выше первый и последний день отпуска.")

        comment = st.text_input("Комментарий (необязательно)", key="comment_%d" % form_n)
        confirmed = st.checkbox("Подаю заявку от имени: **%s**" % name, key="confirm_%d_%s" % (form_n, name))

        b1, b2 = st.columns([3, 1])
        submit = b1.button("Подать заявку", type="primary", width="stretch",
                           disabled=not (start and end) or bool(errors) or not confirmed)
        if b2.button("Сбросить даты", width="stretch", disabled=not start):
            st.session_state.pick_n = pick_n + 1
            st.rerun()
        if submit:
            # перепроверка по свежим данным — вдруг коллеге только что одобрили отпуск
            errors, _ = validate_new_period(name, start, end, year, backend.list_vacations(year),
                                            vc.entitlement(name), RULES, vc.MIN_LONG_PART)
            if errors:
                for e in errors:
                    st.error(e)
            else:
                ts = now()
                backend.add_vacation(new_record(name, start, end, comment.strip(), PENDING, ts))
                write_log(new_log(ts, year, name, "Подана заявка", fmt_period(start, end),
                                  comment.strip(), device()))
                st.session_state.form_n = form_n + 1
                flash("success", "Заявка подана: %s (%s). Статус — «На рассмотрении»."
                      % (fmt_period_words(start, end), n_days(vacation_days(start, end))))
                st.rerun()

with tab_manager:
    if not st.session_state.get("manager_unlocked"):
        st.write("Здесь собраны заявки всех сотрудников — доступ только для руководителя.")
        pw = st.text_input("Код доступа", type="password", key="mgr_pw_input")
        if st.button("Войти", key="mgr_login_btn"):
            if pw == get_secret("manager_passcode", config.MANAGER_PASSCODE_FALLBACK):
                st.session_state.manager_unlocked = True
                st.rerun()
            else:
                st.error("Неверный код доступа.")
    else:
        col1, col2 = st.columns([3, 1])
        m_year = col1.selectbox("Год", [vc.PLAN_YEAR - 1, vc.PLAN_YEAR, vc.PLAN_YEAR + 1], index=1, key="mgr_year")
        if col2.button("Выйти"):
            st.session_state.manager_unlocked = False
            st.rerun()

        records = load_vacations(m_year)
        show_flash()

        # --- Контроль ---
        submitted = {r["name"] for r in records}
        summaries = {n: person_summary(n, records, vc.entitlement(n), vc.MIN_LONG_PART) for n in NAMES}
        n_pending = sum(1 for r in records if r["status"] == PENDING)
        over, crowded = approved_violations(records, RULES)
        crowded_ranges = collapse_dates(crowded)
        conflicts = {r["id"]: record_conflicts(r, records, RULES) for r in active(records)}
        pending_ok = [r for r in records if r["status"] == PENDING and not conflicts[r["id"]]]
        n_pending_conflict = n_pending - len(pending_ok)

        m1, m2, m3 = st.columns(3)
        m1.metric("Подали заявки", "%d / %d" % (len(submitted & set(NAMES)), len(NAMES)))
        m2.metric("На рассмотрении", n_pending)
        m3.metric("Из них с конфликтами", n_pending_conflict,
                  help="Заявки на рассмотрении, которые пересекаются с другими заявками или одобренными "
                       "отпусками больше, чем позволяют правила (подробности — в колонке «Конфликты»).")

        missing = [n for n in NAMES if n not in submitted]
        if missing:
            st.warning("Ещё не подали (%d): %s" % (len(missing), ", ".join(missing)))
        incomplete = ["%s (%d из %d)" % (n, summaries[n]["planned"], summaries[n]["entitlement"])
                      for n in NAMES if n in submitted and summaries[n]["planned"] < summaries[n]["entitlement"]]
        if incomplete:
            st.info("Запланировали не все дни: " + "; ".join(incomplete))
        no_long = [n for n in NAMES if n in submitted and not summaries[n]["has_long"]]
        if no_long:
            st.warning("Нет части от %d дней: %s" % (vc.MIN_LONG_PART, ", ".join(no_long)))

        def fmt_overlap(o):
            a, b, n = o
            return "%s (%s) и %s (%s) — %s" % (a["name"], fmt_period(a["start"], a["end"]),
                                                  b["name"], fmt_period(b["start"], b["end"]), n_days(n))

        if over:
            st.error("Среди ОДОБРЕННЫХ: пересечение основных сотрудников дольше %s:\n\n" % n_days(vc.MAX_OVERLAP_DAYS)
                     + "\n".join("- " + fmt_overlap(o) for o in over))
        if crowded:
            st.error("Среди ОДОБРЕННЫХ: больше %s в отпуске одновременно:\n\n" % n_doctors(vc.MAX_CONCURRENT_MAIN)
                     + "\n".join("- %s — %s" % (fmt_ranges([d for d in crowded if a <= d <= b]),
                                                ", ".join(sorted({n for d in crowded if a <= d <= b for n in crowded[d]})))
                                 for a, b in crowded_ranges))

        # --- Согласование ---
        st.subheader("Заявки")
        if not records:
            st.info("Пока никто не подал заявки на %d год." % m_year)
        else:
            show = st.multiselect("Показать", list(STATUS_LABELS.values()),
                                  default=list(STATUS_LABELS.values()), key="mgr_filter")
            order = {n: i for i, n in enumerate(NAMES)}
            rows = sorted((r for r in records if STATUS_LABELS[r["status"]] in show),
                          key=lambda r: (order.get(r["name"], 999), r["start"]))
            df = pd.DataFrame(
                [{"ФИО": r["name"], "Даты отпуска": fmt_period(r["start"], r["end"]), "Дней": r["days"],
                  "Статус": STATUS_LABELS[r["status"]], "Комментарий руководителя": r["manager_comment"],
                  "Конфликты": "; ".join(conflicts.get(r["id"], [])),
                  "Комментарий сотрудника": r["comment"]} for r in rows],
                index=[r["id"] for r in rows],
                columns=["ФИО", "Даты отпуска", "Дней", "Статус", "Комментарий руководителя", "Конфликты",
                         "Комментарий сотрудника"],
            )
            edited = st.data_editor(
                df, hide_index=True, width="stretch", key="mgr_editor_%d_%d" % (m_year, st.session_state.get("editor_n", 0)),
                disabled=["ФИО", "Даты отпуска", "Дней", "Конфликты", "Комментарий сотрудника"],
                column_config={
                    "Статус": st.column_config.SelectboxColumn(options=list(STATUS_LABELS.values()), required=True),
                    "Дней": st.column_config.NumberColumn(width="small"),
                },
            )
            c1, c2 = st.columns(2)
            if c1.button("Сохранить изменения", type="primary", width="stretch"):
                updates = editor_updates(df, edited)
                if updates:
                    ts = now()
                    backend.update_vacations(updates, ts)
                    by_id = {r["id"]: r for r in records}
                    write_log(*[new_log(ts, m_year, by_id[i]["name"],
                                        "Статус: " + STATUS_LABELS[u["status"]],
                                        fmt_period(by_id[i]["start"], by_id[i]["end"]),
                                        u["manager_comment"], "руководитель")
                                for i, u in updates.items()])
                    pairs, crowd = approved_violations(backend.list_vacations(m_year), RULES)
                    if pairs or crowd:
                        flash("warning", "Сохранено изменений: %d. Внимание: среди одобренных отпусков есть "
                                         "нарушения правил — см. ниже." % len(updates))
                    else:
                        flash("success", "Сохранено изменений: %d." % len(updates))
                else:
                    flash("info", "Изменений нет.")
                st.session_state.editor_n = st.session_state.get("editor_n", 0) + 1
                st.rerun()
            if pending_ok and c2.button("Одобрить без конфликтов (%d)" % len(pending_ok), width="stretch",
                                        help="Одобрить заявки на рассмотрении, которые ни с кем не конфликтуют. "
                                             "Конфликтующие нужно решить вручную в таблице."):
                ts = now()
                backend.update_vacations(
                    {r["id"]: {"status": APPROVED, "manager_comment": r["manager_comment"]} for r in pending_ok}, ts)
                write_log(*[new_log(ts, m_year, r["name"], "Статус: Одобрено", fmt_period(r["start"], r["end"]),
                                    "", "руководитель") for r in pending_ok])
                flash("success", "Одобрено заявок: %d." % len(pending_ok))
                st.session_state.editor_n = st.session_state.get("editor_n", 0) + 1
                st.rerun()

        log = sorted(load_log(m_year), key=lambda e: e["time"], reverse=True)
        if log:
            with st.expander("Журнал действий (%d)" % len(log)):
                st.dataframe(pd.DataFrame(
                    [{"Время": fmt_time(e["time"]), "ФИО": e["name"], "Действие": e["action"], "Даты отпуска": e["period"],
                      "Подробности": e["details"], "Устройство": e["device"]} for e in log]),
                    hide_index=True, width="stretch")

        # --- Календарь ---
        st.subheader("Календарь на год")
        vacation_calendar(**manager_calendar_args(m_year, records), key="mgr_cal_%d" % m_year, default=None)
        st.caption("Наведите на день — видно, кто в отпуске. Ниже — тот же месяц по сотрудникам.")

        st.subheader("По сотрудникам")
        cal_month = st.selectbox("Месяц", list(range(1, 13)), format_func=lambda i: MONTH_NAMES[i - 1], key="mgr_cal_month")
        n = calmod.monthrange(m_year, cal_month)[1]
        dates = [datetime.date(m_year, cal_month, d) for d in range(1, n + 1)]
        cols = ["%d %s" % (d.day, DOW_RU[d.weekday()]) for d in dates]
        by_day = {}
        for r in records:
            if r["status"] != REJECTED:
                for d in daterange(r["start"], r["end"]):
                    by_day[(r["name"], d)] = r["status"]
        cal = pd.DataFrame(
            [["От" if by_day.get((nm, d)) == APPROVED else "?" if by_day.get((nm, d)) == PENDING else ""
              for d in dates] for nm in NAMES],
            index=NAMES, columns=cols,
        )
        team = team_calendar(records, RULES.main_names)
        bad = bad_days(records, RULES)
        ROW_APPR, ROW_PEND = "Основных: одобрено", "Основных: подано"
        cal.loc[ROW_APPR] = [str(len(team.get(d, {}).get("approved", [])) or "") for d in dates]
        cal.loc[ROW_PEND] = [str(len(team.get(d, {}).get("pending", [])) or "") for d in dates]

        def cell_style(v):
            return {"От": "background-color:%s;color:#222" % C_APPROVED,
                    "?": "background-color:%s;color:#222" % C_PENDING}.get(v, "")

        def header_style(col):
            d = dates[cols.index(col.name)]
            base = "background-color:#F4B183" if is_holiday(d) else "background-color:#E7E6E6" if d.weekday() >= 5 else ""
            out = []
            for row, v in zip(col.index, col):
                if row == ROW_APPR:
                    out.append(("background-color:#FF7C80;" if d in bad else "") + "font-weight:bold")
                elif row == ROW_PEND:
                    out.append("font-weight:bold")
                else:
                    out.append(cell_style(v) or base)
            return out

        st.dataframe(cal.style.apply(header_style, axis=0), width="stretch",
                     column_config={c: st.column_config.TextColumn(width=36) for c in cols})
        st.caption("От (зелёный) — одобрено, ? (розовый) — подана заявка. Серым — выходные, оранжевым — "
                   "праздники. Красным в строке «%s» — нарушение среди одобренных: пересечение дольше %s "
                   "или больше %s в отпуске." % (ROW_APPR, n_days(vc.MAX_OVERLAP_DAYS), n_doctors(vc.MAX_CONCURRENT_MAIN)))

        st.download_button(
            "Скачать график отпусков (Excel)",
            data=build_excel(m_year, STAFF, records, vc.entitlement, vc.carryover,
                             RULES, vc.MIN_LONG_PART),
            file_name="grafik_otpuskov_%d.xlsx" % m_year,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
