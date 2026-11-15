"""Initial local application shell; transaction features follow in later phases."""

import streamlit as st

st.set_page_config(page_title="Обработка Excel и CSV", page_icon="📊", layout="wide")
st.title("Обработка Excel и CSV")
st.info(
    "Подготовлена основа приложения. Загрузка файлов, проверка данных и отчёты "
    "будут доступны после следующих этапов разработки."
)
st.caption("Локальное приложение. Данные не отправляются во внешние сервисы.")
