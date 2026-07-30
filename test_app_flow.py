"""Скриптовая проверка через streamlit.testing.v1.AppTest — не часть основного
приложения, только для разработки. Прогоняет: лимит 5 дней "Не могу",
отправку пожеланий, вход руководителя, отображение свода."""
import json
import os

from streamlit.testing.v1 import AppTest

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "local_data.csv")
if os.path.exists(DATA_FILE):
    os.remove(DATA_FILE)

errors = []


def check(cond, msg):
    if not cond:
        errors.append(msg)
    print(("OK  " if cond else "FAIL"), msg)


at = AppTest.from_file("app.py")
at.run()
check(not at.exception, "initial load has no exception")

name = at.session_state["current_name"]
year = at.session_state["current_year"]
month = at.session_state["current_month"]
print("current_name:", name)

# Mark 6 days as CANNOT — the 6th should be blocked (revert to CAN).
import calendar as calmod
n_days = calmod.monthrange(year, month)[1]
target_days = ["%04d-%02d-%02d" % (year, month, d) for d in range(1, 7)]
check(n_days >= 7, "sanity: month has enough days for this test")
for i, iso in enumerate(target_days):
    at.selectbox(key="day_" + iso).select("Не могу").run()

cannot_count = sum(1 for v in at.session_state["sub_statuses"].values() if v == "CANNOT")
check(cannot_count == 5, "cannot_count is capped at 5 (got %d)" % cannot_count)
sixth_status = at.session_state["sub_statuses"][target_days[5]]
check(sixth_status == "CAN", "6th CANNOT attempt was reverted (got %s)" % sixth_status)
sixth_widget_value = at.selectbox(key="day_" + target_days[5]).value
check(sixth_widget_value == "Могу", "6th widget visually reverted to 'Могу' (got %s)" % sixth_widget_value)

# Submit.
at.button[0].click().run()
check(not at.exception, "no exception after submit")
check(any("сохранены" in s.value for s in at.success), "success message shown after submit")

# Verify persisted via local backend directly.
from local_backend import get_submission
year, month = at.session_state["current_year"], at.session_state["current_month"]
saved = get_submission(name, year, month)
check(saved is not None, "submission persisted in local_data.csv")
if saved:
    saved_cannot = sum(1 for v in saved["statuses"].values() if v == "CANNOT")
    check(saved_cannot == 5, "persisted data also has exactly 5 CANNOT days (got %d)" % saved_cannot)

# Manager tab: wrong passcode then correct fallback passcode.
at2 = AppTest.from_file("app.py")
at2.run()
at2.text_input(key="mgr_pw_input").set_value("wrong").run()
at2.button(key="mgr_login_btn").click().run()
check(any("Неверный" in e.value for e in at2.error), "wrong passcode rejected")

at2.text_input(key="mgr_pw_input").set_value("devlocal").run()
at2.button(key="mgr_login_btn").click().run()
check(at2.session_state["manager_unlocked"] is True, "correct fallback passcode unlocks manager tab")
check(not at2.exception, "no exception rendering manager summary")
check(len(at2.dataframe) >= 1, "summary dataframe rendered with the submitted response")

print()
if errors:
    print("FAILURES: %d" % len(errors))
    for e in errors:
        print(" -", e)
    raise SystemExit(1)
else:
    print("ALL CHECKS PASSED")
