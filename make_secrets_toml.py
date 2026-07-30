#!/usr/bin/env python3
"""Собирает корректный secrets.toml из скачанного JSON-ключа сервис-аккаунта.

Частая причина ошибки "Invalid format: please enter valid TOML" на Streamlit
Cloud — ручное копирование полей из JSON (запятые в конце строк, как в JSON,
не нужны в TOML; многострочный private_key ломает однострочную TOML-строку).
Этот скрипт читает JSON как есть и сам корректно формирует TOML — вручную
ничего переносить не нужно.

Использование:
    python3 make_secrets_toml.py путь/к/скачанному-ключу.json
Скрипт спросит код доступа руководителя и ID Google Таблицы, и запишет
результат в .streamlit/secrets.toml (для локальной проверки) и выведет
тот же текст на экран (чтобы скопировать в поле Secrets на share.streamlit.io).
"""
import json
import os
import sys

import tomli_w


def main():
    if len(sys.argv) != 2:
        print("Использование: python3 make_secrets_toml.py путь/к/скачанному-ключу.json")
        sys.exit(1)

    json_path = sys.argv[1]
    if not os.path.exists(json_path):
        print("Файл не найден: %s" % json_path)
        sys.exit(1)

    with open(json_path, encoding="utf-8") as f:
        service_account = json.load(f)

    required = {"type", "project_id", "private_key", "client_email"}
    missing = required - set(service_account.keys())
    if missing:
        print("Похоже, это не ключ сервис-аккаунта Google — не хватает полей: %s" % ", ".join(missing))
        sys.exit(1)

    passcode = input("Код доступа руководителя (например 241024): ").strip()
    sheet_id = input("ID Google Таблицы (кусок ссылки между /d/ и /edit): ").strip()

    secrets = {
        "manager_passcode": passcode,
        "sheet": {"spreadsheet_id": sheet_id},
        "gcp_service_account": service_account,
    }

    toml_text = tomli_w.dumps(secrets)

    # Проверяем сами себя — парсим то, что только что сформировали.
    import tomli
    tomli.loads(toml_text)

    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".streamlit")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "secrets.toml")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(toml_text)

    print("\nГотово и проверено (валидный TOML). Записано в:", out_path)
    print("Скопируйте текст ниже целиком в поле Secrets на share.streamlit.io:\n")
    print("=" * 60)
    print(toml_text)
    print("=" * 60)


if __name__ == "__main__":
    main()
