"""Безопасный доступ к st.secrets — не падает, если secrets.toml вообще не настроен
(например, при локальном запуске без Google-таблицы)."""
import streamlit as st


def get_secret(key, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return default


def has_secret_section(key):
    try:
        return key in st.secrets
    except Exception:
        return False
