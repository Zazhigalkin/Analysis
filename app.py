import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.title("📈 Анализ темпа продаж по рейсам (Cap vs вчера)")

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if uploaded_file:
    # Чтение Excel
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    # Переименование нужных колонок для удобства
    df = df.rename(columns={
        'flt_date&num': 'flight',
        'Ind SS': 'sold_total',
        'Ind SS yesterday': 'sold_yesterday',
        'Cap': 'total_seats'
    })

    required_columns = ['flight','sold_total','sold_yesterday','total_seats']
    if not all(col in df.columns for col in required_columns):
        st.error("⚠️ Файл не соответствует нужной структуре. Проверь названия колонок.")
        st.write("Найденные колонки:", df.columns.tolist())
    else:
        # Разделение flight на дату, номер рейса и маршрут
        df[['flight_date','flight_number','route']] = df['flight'].str.split(" - ", expand=True)
        df['flight_date'] = pd.to_datetime(df['flight_date'], format="%Y.%m.%d", errors='coerce')
        df = df[df['flight_date'].notna()]

        # Остаток мест
        df['remaining_seats'] = df['total_seats'] - df['sold_total']

        # Сравнение с продажами вчера
        df['diff_vs_yesterday'] = df['sold_yesterday'] - df['remaining_seats']

        # Классификация
        def classify(row):
            if row['diff_vs_yesterday'] > 0:
                return "🟢 Опережаем"
            elif row['diff_vs_yesterday'] < 0:
                return "🔴 Отстаём"
            else:
                return "🟡 В графике"

        df['status'] = df.apply(classify, axis=1)

        # Итоговая таблица
        result = df[['flight','flight_date','flight_number','route','total_seats','sold_total',
                     'sold_yesterday','remaining_seats','diff_vs_yesterday','status']]

        st.subheader("📊 Результаты анализа темпа продаж")
        st.dataframe(result, use_container_width=True)

        # Скачать Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result.to_excel(writer, index=False, sheet_name='Sales_Speed')
        st.download_button(label="💾 Скачать отчёт в Excel",
                           data=output.getvalue(),
                           file_name="sales_speed_vs_yesterday.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("⬆️ Загрузите Excel файл, чтобы начать анализ.")
