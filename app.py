import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.title("📈 Анализ темпа продаж по рейсам с учётом даты сегодня")

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    # Переименование колонок
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

        # Сегодняшняя дата
        today = datetime.today().date()

        # Дни до вылета
        df['days_to_flight'] = df['flight_date'].apply(lambda x: max((x.date() - today).days,1))

        # Остаток мест
        df['remaining_seats'] = df['total_seats'] - df['sold_total']

        # Необходимый дневной темп
        df['daily_needed'] = df['remaining_seats'] / df['days_to_flight']

        # Разница вчерашних продаж и плана на день
        df['diff_vs_plan'] = df['sold_yesterday'] - df['daily_needed']

        # Классификация: 🔵 перепродажа, 🟢 по плану, 🔴 отстаём, ⚪ до рейса ещё далеко
        def classify(row):
            # Если рейс далеко (>30 дней) и daily_needed <4 — мелкий/неактуальный рейс
            if row['days_to_flight'] > 30 and row['daily_needed'] < 4:
                return "⚪ До рейса ещё далеко"
            elif row['diff_vs_plan'] > 5:
                return "🔵 Перепродажа"
            elif -5 <= row['diff_vs_plan'] <= 5:
                return "🟢 По плану"
            else:
                return "🔴 Отстаём"

        df['status'] = df.apply(classify, axis=1)

        # Итоговая таблица
        result = df[['flight','flight_date','flight_number','route','total_seats','sold_total',
                     'sold_yesterday','remaining_seats','days_to_flight','daily_needed','diff_vs_plan','status']]

        st.subheader("📊 Результаты анализа темпа продаж")
        st.dataframe(result, use_container_width=True)

        # Вывод рейсов, на которые нужно обратить внимание (исключая "до рейса ещё далеко")
        attention_df = result[result['status'].isin(["🔴 Отстаём", "🔵 Перепродажа"])]
        if not attention_df.empty:
            st.subheader("⚠️ Рейсы, требующие внимания")
            st.dataframe(attention_df[['flight','flight_date','flight_number','route','sold_yesterday','daily_needed','diff_vs_plan','status']], use_container_width=True)

        # Скачать Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result.to_excel(writer, index=False, sheet_name='Sales_Speed')
        st.download_button(label="💾 Скачать отчёт в Excel",
                           data=output.getvalue(),
                           file_name="sales_speed_with_attention.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("⬆️ Загрузите Excel файл, чтобы начать анализ.")
