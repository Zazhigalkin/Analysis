import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.title("📈 Темп продаж по рейсам (упрощённый)")

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    # Переименование колонок
    df = df.rename(columns={
        'flt_date&num': 'flight',
        'Ind SS yesterday': 'yesterday',
        'Cap': 'total_seats'
    })

    required_columns = ['flight','yesterday','total_seats']
    if not all(col in df.columns for col in required_columns):
        st.error("⚠️ Файл не соответствует нужной структуре. Проверь названия колонок.")
        st.write("Найденные колонки:", df.columns.tolist())
    else:
        # Разбор flight на дату
        df[['flight_date','flight_number','route']] = df['flight'].str.split(" - ", expand=True)
        df['flight_date'] = pd.to_datetime(df['flight_date'], format="%Y.%m.%d", errors='coerce')
        df = df[df['flight_date'].notna()]

        today = datetime.today().date()
        df['days_to_flight'] = df['flight_date'].apply(lambda x: max((x.date() - today).days,1))

        # Продано до вчера = вчерашние продажи (можно адаптировать, если есть d_14_plus и т.д.)
        df['sold_so_far'] = df['yesterday']  # упрощённо, без исторических данных

        # Оставшиеся места
        df['remaining_seats'] = (df['total_seats'] - df['sold_so_far']).apply(lambda x: max(x,0))

        # Необходимый темп продаж на день
        df['daily_needed'] = df['remaining_seats'] / df['days_to_flight']

        # Сравнение вчера с планом
        df['yesterday_vs_needed'] = df['yesterday'] - df['daily_needed']

        def classify(row):
            if row['yesterday_vs_needed'] > 0:
                return "🟢 Опережаем"
            elif row['yesterday_vs_needed'] < 0:
                return "🔴 Отстаём"
            else:
                return "🟡 В графике"

        df['status'] = df.apply(classify, axis=1)

        # Итоговая таблица
        result = df[['flight','flight_date','flight_number','route','total_seats','sold_so_far',
                     'remaining_seats','days_to_flight','daily_needed','yesterday','yesterday_vs_needed','status']]

        st.subheader("📊 Результаты темпа продаж")
        st.dataframe(result, use_container_width=True)

        # Скачать Excel
        output=io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result.to_excel(writer, index=False, sheet_name='Sales Speed')
        st.download_button(label="💾 Скачать отчёт в Excel",
                           data=output.getvalue(),
                           file_name="sales_speed_simple.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("⬆️ Загрузите Excel файл, чтобы начать анализ.")
