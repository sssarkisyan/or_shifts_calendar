"""Подсчёт дней отпуска, проверки правил и выгрузка в Excel — без Streamlit,
чтобы это можно было проверять отдельно (test_vacation_flow.py)."""
import calendar as calmod
import dataclasses
import datetime
import functools
import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PENDING, APPROVED, REJECTED = "PENDING", "APPROVED", "REJECTED"
STATUS_LABELS = {PENDING: "На рассмотрении", APPROVED: "Одобрено", REJECTED: "Отклонено"}
ACTIVE_STATUSES = {PENDING, APPROVED}  # отклонённые заявки дни не занимают

MONTH_NAMES = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
               "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
DOW_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]


# Нерабочие праздничные дни по ст. 112 ТК РФ (месяц, день). Список задан явно,
# а не берётся из библиотеки holidays: разные её версии по-разному относятся
# к переносам, и на сервере портал начинал не считать в отпуск перенесённые
# выходные (3 и 10 мая, 14 июня, 5 ноября 2027). По ст. 120 ТК из отпуска
# исключаются только праздничные дни; перенесённый выходной — обычный выходной
# и в отпуск входит.
RU_HOLIDAYS_MD = (
    (1, 1), (1, 2), (1, 3), (1, 4), (1, 5), (1, 6), (1, 7), (1, 8),
    (2, 23), (3, 8), (5, 1), (5, 9), (6, 12), (11, 4),
)


@functools.lru_cache(maxsize=None)
def ru_holidays(year):
    """Праздничные нерабочие дни РФ (ст. 112 ТК) — без переносов выходных."""
    return frozenset(datetime.date(year, m, d) for m, d in RU_HOLIDAYS_MD)


def is_holiday(d):
    return d in ru_holidays(d.year)


# Производственный календарь (переносы выходных по постановлениям
# Правительства). Нужен только для отображения — какие дни в календарях и
# Excel показывать выходными; на подсчёт дней отпуска не влияет (там только
# праздники ст. 112, см. vacation_days). Год без записи — обычная неделя +
# праздники. Добавлять новый год по consultant.ru/law/ref/calendar/proizvodstvennye/.
PRODUCTION_CALENDAR = {
    # Постановление Правительства РФ от 17.09.2026 № 1187 «О переносе
    # выходных дней в 2027 году» (consultant.ru, производственный календарь 2027)
    2027: {
        "days_off": {(2, 22), (5, 3), (5, 10), (6, 14), (11, 5), (12, 31)},
        "workdays": {(2, 20)},  # рабочая суббота (предпраздничная, на час короче)
    },
}


def is_day_off(d):
    """Нерабочий день по производственному календарю: выходной, праздник или
    перенесённый выходной (рабочие субботы — рабочие)."""
    cal = PRODUCTION_CALENDAR.get(d.year, {})
    if (d.month, d.day) in cal.get("workdays", ()):
        return False
    if (d.month, d.day) in cal.get("days_off", ()) or is_holiday(d):
        return True
    return d.weekday() >= 5


def daterange(start, end):
    d = start
    while d <= end:
        yield d
        d += datetime.timedelta(days=1)


def vacation_days(start, end):
    """Дней отпуска в периоде: календарные дни минус праздничные (ст. 120 ТК)."""
    return sum(1 for d in daterange(start, end) if not is_holiday(d))


def fmt_date(d):
    return d.strftime("%d.%m.%Y")


def fmt_period(start, end):
    return "%s – %s" % (fmt_date(start), fmt_date(end))


MONTHS_GEN = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
              "августа", "сентября", "октября", "ноября", "декабря"]


def fmt_date_words(d):
    return "%d %s %d" % (d.day, MONTHS_GEN[d.month - 1], d.year)


def fmt_period_words(start, end):
    """«с 1 июля по 14 июля 2027» (год один раз, если совпадает)."""
    if start.year == end.year:
        return "с %d %s по %s" % (start.day, MONTHS_GEN[start.month - 1], fmt_date_words(end))
    return "с %s по %s" % (fmt_date_words(start), fmt_date_words(end))


def plural(n, one, few, many):
    """plural(5, "день", "дня", "дней") → "дней"."""
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def n_days(n):
    """5 → "5 дней", 1 → "1 день", 3 → "3 дня"."""
    return "%d %s" % (n, plural(n, "день", "дня", "дней"))


def n_doctors(n):
    return "%d %s" % (n, plural(n, "основной врач", "основных врача", "основных врачей"))


def n_doctors_gen(n):
    """Родительный падеж — после «больше», «у»: «2 основных врачей», «1 основного врача»."""
    return "%d %s" % (n, "основного врача" if n % 10 == 1 and n % 100 != 11 else "основных врачей")


def active(records):
    return [r for r in records if r["status"] in ACTIVE_STATUSES]


def occupancy(records, exclude_name=None):
    """{дата: [ФИО, ...]} — кто в отпуске в этот день (активные заявки)."""
    out = {}
    for r in active(records):
        if r["name"] == exclude_name:
            continue
        for d in daterange(r["start"], r["end"]):
            out.setdefault(d, []).append(r["name"])
    return out


def collapse_dates(dates):
    """[даты] → [(начало, конец)] — подряд идущие дни склеиваются в интервалы."""
    ranges = []
    for d in sorted(dates):
        if ranges and d - ranges[-1][1] == datetime.timedelta(days=1):
            ranges[-1][1] = d
        else:
            ranges.append([d, d])
    return [(a, b) for a, b in ranges]


def fmt_ranges(dates):
    return ", ".join(
        fmt_date(a) if a == b else "%s–%s" % (a.strftime("%d.%m"), fmt_date(b))
        for a, b in collapse_dates(dates)
    )


@dataclasses.dataclass(frozen=True)
class TeamRules:
    """Правила для основных сотрудников (main_names): отпуска двух из них
    пересекаются не больше max_overlap календарных дней, и в один день в
    отпуске не больше max_concurrent из них. Совместители не ограничены.
    Блокируют подачу только ОДОБРЕННЫЕ отпуска коллег — поданные, но не
    рассмотренные заявки лишь показываются (кто первым подал — не важно,
    решает руководитель)."""
    main_names: frozenset
    max_overlap: int
    max_concurrent: int


def overlap_days(a, b):
    """Сколько календарных дней пересекаются два периода (заявки)."""
    return max(0, (min(a["end"], b["end"]) - max(a["start"], b["start"])).days + 1)


def team_conflicts(name, start, end, others, rules, all_approved=True):
    """Нарушения командных правил, если сотрудник name уйдёт в отпуск
    start–end, а others — уже учитываемые отпуска коллег. Список строк.
    all_approved — others только одобренные (тогда «в отпуске уже … — это
    максимум»), иначе среди них есть поданные заявки («заявки уже у …»)."""
    if name not in rules.main_names:
        return []
    others = [r for r in others if r["name"] != name and r["name"] in rules.main_names]
    out = []
    new = {"start": start, "end": end}
    for r in sorted(others, key=lambda r: r["start"]):
        n = overlap_days(new, r)
        if n > rules.max_overlap:
            out.append("пересечение с %s (%s) — %s, допускается не больше %s"
                       % (r["name"], fmt_period(r["start"], r["end"]), n_days(n), n_days(rules.max_overlap)))
    occ = occupancy(others)
    full = [d for d in daterange(start, end) if len(occ.get(d, [])) >= rules.max_concurrent]
    if full:
        who = sorted({n for d in full for n in occ[d]})
        if all_approved:
            out.append("%s в отпуске уже %s (%s) — это максимум"
                       % (fmt_ranges(full), n_doctors(rules.max_concurrent), ", ".join(who)))
        else:
            out.append("%s заявки уже у %s (%s)"
                       % (fmt_ranges(full), n_doctors_gen(rules.max_concurrent), ", ".join(who)))
    return out


def approved(records):
    return [r for r in records if r["status"] == APPROVED]


def validate_new_period(name, start, end, year, all_records, entitlement, rules, min_long_part, today=None,
                        hard_limit=True):
    """(ошибки, предупреждения). Ошибки блокируют подачу: свои правила (лимит
    дней, часть ≥14, пересечение со своим периодом) и конфликт с ОДОБРЕННЫМИ
    отпусками коллег. Предупреждения — конфликт с поданными заявками коллег:
    подать можно, решение за заведующим. today — нельзя подать на прошедшие даты.
    hard_limit=False (совместители) — превышение лимита дней не блокирует, а
    только предупреждает; правило «часть ≥14» при подаче тогда не проверяется
    (без лимита длинную часть всегда можно добавить позже)."""
    if start > end:
        return ["Дата окончания раньше даты начала."], []
    if today is not None and start < today:
        return ["Нельзя подать отпуск на прошедшие даты."], []
    if start.year != year or end.year != year:
        return ["Отпуск должен целиком лежать в %d году." % year], []
    days = vacation_days(start, end)
    if days == 0:
        return ["В выбранных датах только праздничные дни — они в отпуск не входят."], []

    errors = []
    mine = [r for r in active(all_records) if r["name"] == name]
    for r in mine:
        if start <= r["end"] and r["start"] <= end:
            errors.append("Пересекается с вашим отпуском %s." % fmt_period(r["start"], r["end"]))

    used = sum(r["days"] for r in mine)
    soft = []
    if used + days > entitlement:
        if hard_limit:
            errors.append(
                "Превышен лимит: положено %s, уже запланировано %s, в этом отпуске %s (осталось %s)."
                % (n_days(entitlement), n_days(used), n_days(days), n_days(max(entitlement - used, 0)))
            )
        else:
            soft.append(
                "Больше положенного: положено %s, с этим отпуском будет %s. Подать можно — "
                "решение примет заведующий." % (n_days(entitlement), n_days(used + days))
            )
    elif hard_limit:
        has_long = any(r["days"] >= min_long_part for r in mine)
        remaining_after = entitlement - used - days
        if not has_long and days < min_long_part and remaining_after < min_long_part:
            errors.append(
                "Хотя бы одна часть отпуска должна быть не меньше %s (ст. 125 ТК РФ). "
                "После этого отпуска останется %s — длинную часть будет уже не набрать. "
                "Удлините отпуск или сначала запланируйте длинную часть."
                % (n_days(min_long_part), n_days(remaining_after))
            )

    errors += ["Даты недоступны: " + c + " (отпуска коллег уже одобрены)."
               for c in team_conflicts(name, start, end, approved(all_records), rules)]
    warnings = [] if errors else soft + [
        "На эти даты уже поданы заявки: " + c + ". Подать можно — решение примет заведующий."
        for c in team_conflicts(name, start, end, active(all_records), rules, all_approved=False)
    ]
    return errors, warnings


def record_conflicts(rec, records, rules):
    """Конфликты заявки rec с остальными активными заявками (для руководителя)."""
    others = [r for r in active(records) if r["id"] != rec["id"]]
    return team_conflicts(rec["name"], rec["start"], rec["end"], others, rules, all_approved=False)


def approved_violations(records, rules):
    """Нарушения среди одобренных отпусков: (пары с пересечением > max_overlap,
    {дата: [ФИО]} где одобренных основных больше max_concurrent)."""
    appr = [r for r in approved(records) if r["name"] in rules.main_names]
    pairs = []
    for i, a in enumerate(appr):
        for b in appr[i + 1:]:
            if a["name"] != b["name"] and overlap_days(a, b) > rules.max_overlap:
                pairs.append((a, b, overlap_days(a, b)))
    crowded = {d: who for d, who in occupancy(appr).items() if len(who) > rules.max_concurrent}
    return pairs, crowded


def bad_days(records, rules):
    """Дни, попадающие в нарушения среди одобренных отпусков."""
    pairs, crowded = approved_violations(records, rules)
    out = set(crowded)
    for a, b, _ in pairs:
        out.update(daterange(max(a["start"], b["start"]), min(a["end"], b["end"])))
    return out


def team_calendar(records, names, exclude_name=None):
    """{дата: {"approved": [ФИО], "pending": [ФИО]}} — отпуска сотрудников из
    names (кроме exclude_name) по дням."""
    out = {}
    for r in active(records):
        if r["name"] == exclude_name or r["name"] not in names:
            continue
        key = "approved" if r["status"] == APPROVED else "pending"
        for d in daterange(r["start"], r["end"]):
            out.setdefault(d, {"approved": [], "pending": []})[key].append(r["name"])
    return out


def person_summary(name, records, entitlement, min_long_part):
    mine = [r for r in records if r["name"] == name]
    act = active(mine)
    return {
        "entitlement": entitlement,
        "planned": sum(r["days"] for r in act),
        "approved": sum(r["days"] for r in act if r["status"] == APPROVED),
        "pending": sum(r["days"] for r in act if r["status"] == PENDING),
        "has_long": any(r["days"] >= min_long_part for r in act),
        "n_periods": len(mine),
    }


def editor_updates(original, edited):
    """Сравнить таблицу согласования до/после правки руководителем (DataFrame,
    индекс — ID заявки). Возвращает {id: {"status": код, "manager_comment": ..}}
    только для изменившихся строк."""
    label_to_status = {v: k for k, v in STATUS_LABELS.items()}
    updates = {}
    for rec_id in edited.index:
        label = edited.at[rec_id, "Статус"]
        comment = edited.at[rec_id, "Комментарий заведующего"] or ""
        if (label, comment) != (original.at[rec_id, "Статус"], original.at[rec_id, "Комментарий заведующего"]):
            updates[rec_id] = {"status": label_to_status[label], "manager_comment": comment}
    return updates


# ---------- Excel ----------

# Те же цвета, что в приложении: бледно-зелёный — одобрено, бледно-розовый — подано.
FILL_APPROVED = PatternFill("solid", fgColor="CDEBD3")
FILL_PENDING = PatternFill("solid", fgColor="F9D5DC")
FILL_TODO = PatternFill("solid", fgColor="FFE699")
FILL_WEEKEND = PatternFill("solid", fgColor="E7E6E6")
FILL_HOLIDAY = PatternFill("solid", fgColor="F4B183")
FILL_BAD = PatternFill("solid", fgColor="FF7C80")
BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center")


MONTH_SHORT = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
STAFF_GROUPS = [("Основные", {"основной"}), ("Совместители", {"совместитель"}),
                ("По договору", {"договор"}), ("Заведующий и профессор", {"заведующий", "профессор"})]
FILL_GROUP = PatternFill("solid", fgColor="D9E1F2")
FILL_HOLIDAY_LIGHT = PatternFill("solid", fgColor="FBE5D6")  # праздники в «Год по дням» — бледно, не забивая картину
FILL_WEEKEND_LIGHT = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
MONTH_EDGE = Side(style="medium", color="808080")


def short_name(full):
    """«Исхаков Руслан Ильдарович» → «Исхаков Р.И.»"""
    parts = full.split()
    return parts[0] + (" " + "".join(p[0] + "." for p in parts[1:3]) if len(parts) > 1 else "")


def _grouped_staff(staff):
    """[(название группы, [сотрудники]), ...] в порядке STAFF_GROUPS, без пустых групп."""
    out, seen = [], set()
    for title, kinds in STAFF_GROUPS:
        members = [s for s in staff if s["kind"] in kinds]
        seen.update(s["name"] for s in members)
        if members:
            out.append((title, members))
    rest = [s for s in staff if s["name"] not in seen]
    if rest:
        out.append(("Прочие", rest))
    return out


def _add_year_sheets(wb, year, staff, records, entitlement_fn, rules):
    """Два обзорных листа на весь год: «Год по месяцам» (даты отпусков в клетке
    месяца) и «Год по дням» (каждый день — клетка, отпуск закрашен)."""
    act = sorted(active(records), key=lambda r: r["start"])
    by_name = {}
    for r in act:
        by_name.setdefault(r["name"], []).append(r)
    groups = _grouped_staff(staff)
    small = Font(size=9)
    title_font = Font(bold=True, size=14)
    legend = "Зелёный — одобрено, розовый — на рассмотрении. Отклонённые заявки не показаны."

    # --- Год по месяцам ---
    ws = wb.create_sheet("Год по месяцам")
    ws["A1"] = "Отпуска врачей — %d год" % year
    ws["A1"].font = title_font
    ws["A2"] = legend + " В клетке — числа месяца."
    ws["A2"].font = Font(italic=True, size=9, color="595959")
    header = ["Сотрудник", "Ставка"] + MONTH_SHORT + ["Дней", "Положено"]
    ws.append([])
    ws.append(header)
    hdr_row = ws.max_row
    for c in ws[hdr_row]:
        c.font = BOLD
        c.alignment = CENTER
        c.border = Border(bottom=MONTH_EDGE)
    month_days = [0] * 12
    for title, members in groups:
        ws.append([title])
        r0 = ws.max_row
        ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=len(header))
        ws.cell(r0, 1).font = BOLD
        ws.cell(r0, 1).fill = FILL_GROUP
        for s in members:
            recs = by_name.get(s["name"], [])
            row = [short_name(s["name"]), s["rate"]]
            fills = []
            for m in range(1, 13):
                parts, statuses = [], set()
                for r in recs:
                    a = max(r["start"], datetime.date(year, m, 1))
                    b = min(r["end"], datetime.date(year, m, calmod.monthrange(year, m)[1]))
                    if a <= b:
                        parts.append(str(a.day) if a == b else "%d–%d" % (a.day, b.day))
                        statuses.add(r["status"])
                        month_days[m - 1] += (b - a).days + 1
                row.append(", ".join(parts))
                fills.append(None if not statuses else FILL_APPROVED if statuses == {APPROVED} else FILL_PENDING)
            total = sum(r["days"] for r in recs)
            ent = entitlement_fn(s["name"])
            row += [total or "", ent]
            ws.append(row)
            rr = ws.max_row
            for i, f in enumerate(fills):
                c = ws.cell(rr, 3 + i)
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
                if f:
                    c.fill = f
            ws.cell(rr, 2).alignment = CENTER
            dc = ws.cell(rr, 15)
            dc.alignment = CENTER
            dc.font = BOLD
            if total > ent:
                dc.fill = FILL_BAD
            ws.cell(rr, 16).alignment = CENTER
            ws.cell(rr, 16).font = Font(color="808080")
    ws.append(["Всего календарных дней отпуска", ""] + [d or "" for d in month_days])
    rr = ws.max_row
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
    for c in ws[rr]:
        c.font = BOLD
        c.alignment = CENTER
        c.border = Border(top=MONTH_EDGE)
    ws.cell(rr, 1).alignment = Alignment(horizontal="left")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 7
    for i in range(12):
        ws.column_dimensions[get_column_letter(3 + i)].width = 11
    ws.column_dimensions[get_column_letter(15)].width = 7
    ws.column_dimensions[get_column_letter(16)].width = 9
    ws.freeze_panes = "C%d" % (hdr_row + 1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # --- Год по дням ---
    ws = wb.create_sheet("Год по дням")
    days = list(daterange(datetime.date(year, 1, 1), datetime.date(year, 12, 31)))
    first_col = 3
    ws["A1"] = "Отпуска врачей — %d год, по дням" % year
    ws["A1"].font = title_font
    ws["A2"] = legend + " Серым — выходные, бледно-оранжевым — праздники."
    ws["A2"].font = Font(italic=True, size=9, color="595959")
    # строка 3 — месяцы (объединённые), 4 — числа (выходные/праздники красным)
    for m in range(1, 13):
        cols = [first_col + i for i, d in enumerate(days) if d.month == m]
        ws.merge_cells(start_row=3, start_column=cols[0], end_row=3, end_column=cols[-1])
        c = ws.cell(3, cols[0], MONTH_NAMES[m - 1])
        c.font = BOLD
        c.alignment = CENTER
    ws.cell(4, 1, "Сотрудник").font = BOLD
    ws.cell(4, 2, "Дней").font = BOLD
    for i, d in enumerate(days):
        col = first_col + i
        c = ws.cell(4, col, d.day)
        c.font = Font(size=7, bold=True, color="C0392B" if is_day_off(d) else None)
        c.alignment = CENTER
        c.border = Border(left=MONTH_EDGE if d.day == 1 else None, bottom=THIN)
        ws.column_dimensions[get_column_letter(col)].width = 2.6
    status_by = {}
    for r in act:
        for d in daterange(r["start"], r["end"]):
            status_by[(r["name"], d)] = r["status"]
    for title, members in groups:
        ws.append([title])
        c = ws.cell(ws.max_row, 1)
        c.font = BOLD
        c.fill = FILL_GROUP
        ws.cell(ws.max_row, 2).fill = FILL_GROUP
        for s in members:
            total = sum(r["days"] for r in by_name.get(s["name"], []))
            ws.append([short_name(s["name"]), total or ""])
            rr = ws.max_row
            ws.cell(rr, 2).alignment = CENTER
            ws.cell(rr, 2).font = small
            for i, d in enumerate(days):
                c = ws.cell(rr, first_col + i)
                st = status_by.get((s["name"], d))
                if st == APPROVED:
                    c.fill = FILL_APPROVED
                elif st == PENDING:
                    c.fill = FILL_PENDING
                elif is_holiday(d):
                    c.fill = FILL_HOLIDAY_LIGHT
                elif is_day_off(d):
                    c.fill = FILL_WEEKEND_LIGHT
                c.border = Border(left=MONTH_EDGE if d.day == 1 else None, bottom=THIN)
    # итоговая строка: сколько основных в отпуске в этот день
    occ = occupancy([r for r in act if r["name"] in rules.main_names])
    ws.append(["Основных в отпуске", ""])
    rr = ws.max_row
    ws.cell(rr, 1).font = BOLD
    for i, d in enumerate(days):
        n = len(occ.get(d, []))
        c = ws.cell(rr, first_col + i, n or None)
        c.font = Font(size=7, bold=True)
        c.alignment = CENTER
        c.border = Border(top=MONTH_EDGE, left=MONTH_EDGE if d.day == 1 else None)
        if n > rules.max_concurrent:
            c.fill = FILL_BAD
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 5
    ws.freeze_panes = "%s5" % get_column_letter(first_col)
    ws.page_setup.orientation = "landscape"


def build_excel(year, staff, records, entitlement_fn, rules, min_long_part):
    """Свод + год по месяцам + год по дням + 12 листов-календарей + список всех
    заявок. Возвращает bytes .xlsx."""
    wb = Workbook()
    names = [s["name"] for s in staff]

    # --- Свод ---
    ws = wb.active
    ws.title = "Свод"
    ws.append(["ФИО", "Положено", "Запланировано", "Одобрено", "На рассмотрении",
               "Осталось", "Есть часть от %s" % n_days(min_long_part), "Отпуска"])
    for c in ws[1]:
        c.font = BOLD
    for name in names:
        s = person_summary(name, records, entitlement_fn(name), min_long_part)
        periods = "; ".join(
            "%s (%s, %s)" % (fmt_period(r["start"], r["end"]), n_days(r["days"]), STATUS_LABELS[r["status"]].lower())
            for r in sorted(records, key=lambda r: r["start"]) if r["name"] == name
        )
        ws.append([name, s["entitlement"], s["planned"], s["approved"], s["pending"],
                   s["entitlement"] - s["planned"], "да" if s["has_long"] else "нет", periods])
        row = ws.max_row
        if s["planned"] < s["entitlement"]:
            ws.cell(row, 6).fill = FILL_TODO
        if not s["has_long"]:
            ws.cell(row, 7).fill = FILL_BAD
    ws.column_dimensions["A"].width = 36
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["H"].width = 90
    ws.append([])
    ws.append(["Обозначения в календарях:"])
    for label, fill in [("Одобрено", FILL_APPROVED), ("На рассмотрении", FILL_PENDING),
                        ("Праздник", FILL_HOLIDAY), ("Выходной", FILL_WEEKEND),
                        ("Нарушение среди одобренных: пересечение основных дольше %s "
                         "или больше %s в отпуске" % (n_days(rules.max_overlap), n_doctors_gen(rules.max_concurrent)),
                         FILL_BAD)]:
        ws.append([label])
        ws.cell(ws.max_row, 1).fill = fill

    _add_year_sheets(wb, year, staff, records, entitlement_fn, rules)

    # --- Календари по месяцам ---
    by_day = {}
    for r in active(records):
        for d in daterange(r["start"], r["end"]):
            by_day[(r["name"], d)] = r["status"]
    team = team_calendar(records, rules.main_names)
    bad = bad_days(records, rules)
    for m in range(1, 13):
        ws = wb.create_sheet(MONTH_NAMES[m - 1])
        n = calmod.monthrange(year, m)[1]
        dates = [datetime.date(year, m, d) for d in range(1, n + 1)]
        ws.append(["ФИО"] + [d.day for d in dates])
        ws.append([""] + [DOW_RU[d.weekday()] for d in dates])
        for col, d in enumerate(dates, start=2):
            for row in (1, 2):
                c = ws.cell(row, col)
                c.font = BOLD
                c.alignment = CENTER
                if is_holiday(d):
                    c.fill = FILL_HOLIDAY
                elif is_day_off(d):
                    c.fill = FILL_WEEKEND
        for name in names:
            ws.append([name] + ["От" if (name, d) in by_day else "" for d in dates])
            row = ws.max_row
            for col, d in enumerate(dates, start=2):
                st = by_day.get((name, d))
                c = ws.cell(row, col)
                c.alignment = CENTER
                if st == APPROVED:
                    c.fill = FILL_APPROVED
                elif st == PENDING:
                    c.fill = FILL_PENDING
        for label, key in (("Основных: одобрено", "approved"), ("Основных: подано", "pending")):
            ws.append([label] + [len(team.get(d, {}).get(key, [])) or "" for d in dates])
            row = ws.max_row
            ws.cell(row, 1).font = BOLD
            for col, d in enumerate(dates, start=2):
                c = ws.cell(row, col)
                c.alignment = CENTER
                c.font = BOLD
                if key == "approved" and d in bad:
                    c.fill = FILL_BAD
        ws.column_dimensions["A"].width = 36
        for col in range(2, n + 2):
            ws.column_dimensions[get_column_letter(col)].width = 4.5
        ws.freeze_panes = "B3"

    # --- Все заявки ---
    ws = wb.create_sheet("Заявки")
    ws.append(["ФИО", "Начало", "Конец", "Дней", "Статус", "Комментарий сотрудника",
               "Комментарий заведующего", "Подано", "Обновлено"])
    for c in ws[1]:
        c.font = BOLD
    for r in sorted(records, key=lambda r: (names.index(r["name"]) if r["name"] in names else 999, r["start"])):
        ws.append([r["name"], r["start"], r["end"], r["days"], STATUS_LABELS[r["status"]],
                   r["comment"], r["manager_comment"], r["created"], r["updated"]])
        ws.cell(ws.max_row, 2).number_format = "DD.MM.YYYY"
        ws.cell(ws.max_row, 3).number_format = "DD.MM.YYYY"
    for col, w in zip("ABCDEFGHI", [36, 12, 12, 7, 17, 40, 40, 20, 20]):
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
