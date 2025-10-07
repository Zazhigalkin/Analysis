import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import plotly.express as px

st.set_page_config(page_title="Sales Speed Analyzer", layout="wide")
st.title("📈 Анализ скорости продаж по рейсам")

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if uploaded_file:
    # 1️⃣ Чтение Excel
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    # 2️⃣ Переименование колонок для удобства
    df = df.rename(columns={
        'flt_date&num': 'flight',
        'Ind SS today': 'today',
        'Ind SS yesterday': 'yesterday',
        'Ind SS 2-3 days before': 'd_2_3',
        'Ind SS 4-6 days before': 'd_4_6',
        'Ind SS 7-13 days before': 'd_7_13',
        'Ind SS last 14 days': 'd_14_plus',
        'Cap': 'total_seats'
    })

    required_columns = ['flight','today','yesterday','d_2_3','d_4_6','d_7_13','d_14_plus','total_seats']
    if not all(col in df.columns for col in required_columns):
        st.error("⚠️ Файл не соответствует нужной структуре. Проверь названия колонок.")
        st.write("Найденные колонки:", df.columns.tolist())
    else:
        # 3️⃣ Разбор flight на дату, номер рейса и маршрут
        def split_flight_info(flight_str):
            try:
                parts = str(flight_str).split(" - ")
                flight_date = pd.to_datetime(parts[0], format="%Y.%m.%d", errors='coerce')
                flight_number = parts[1] if len(parts) > 1 else None
                route = parts[2] if len(parts) > 2 else None
                return pd.Series([flight_date, flight_number, route])
            except:
                return pd.Series([pd.NaT, None, None])

        df[['flight_date','flight_number','route']] = df['flight'].apply(split_flight_info)
        df = df[df['flight_date'].notna()]

        # 4️⃣ Дни до вылета
        today = datetime.today().date()
        df['days_to_flight'] = df['flight_date'].apply(lambda x: max((x.date() - today).days, 1))

        # 5️⃣ Расчёт темпа продаж (старый метод)
        def calc_sales_speed(row):
            early = row['d_14_plus'] + row['d_7_13']
            recent = row['d_4_6'] + row['d_2_3'] + row['yesterday'] + row['today']
            return 0 if early==0 else recent/early

        df['sales_speed_ratio'] = df.apply(calc_sales_speed, axis=1)
        df['sales_speed_status'] = df['sales_speed_ratio'].apply(
            lambda r: "🟢 Быстро" if r>1.2 else ("🔴 Медленно" if r<0.8 else "🟡 Нормально")
        )

        # 6️⃣ Учёт лимита билетов
        df['sold_so_far'] = df['d_14_plus'] + df['d_7_13'] + df['d_4_6'] + df['d_2_3'] + df['yesterday'] + df['today']
        df['remaining_seats'] = (df['total_seats'] - df['sold_so_far']).apply(lambda x: max(x,0))
        df['daily_needed'] = df['remaining_seats']/df['days_to_flight']
        df['today_vs_needed'] = df['today'] - df['daily_needed']

        def daily_needed_status(row):
            if row['today_vs_needed'] > 0:
                return "🟢 Превышаем"
            elif row['today_vs_needed'] < 0:
                return "🔴 Отстаем"
            else:
                return "🟡 В графике"

        df['daily_needed_comparison'] = df.apply(daily_needed_status, axis=1)

        # 7️⃣ Аномалии
        def attention_flag(row):
            flags=[]
            notes=[]
            if row['today'] > max(row['d_14_plus'],1)*1.5:
                flags.append('Сегодня'); notes.append("Сегодня продажи резко выше исторических значений.")
            if row['yesterday'] > max(row['d_14_plus'],1)*1.5:
                flags.append('Вчера'); notes.append("Вчера продажи резко выше исторических значений.")
            if row['d_2_3'] > max(row['d_14_plus'],1)*1.5:
                flags.append('2-3 дня'); notes.append("Продажи 2-3 дня назад резко выше исторических значений.")
            if row['today'] < max(row['d_14_plus'],1)*0.5:
                flags.append('Сегодня'); notes.append("Сегодня продажи резко ниже исторических значений.")
            if row['yesterday'] < max(row['d_14_plus'],1)*0.5:
                flags.append('Вчера'); notes.append("Вчера продажи резко ниже исторических значений.")
            if row['d_2_3'] < max(row['d_14_plus'],1)*0.5:
                flags.append('2-3 дня'); notes.append("Продажи 2-3 дня назад резко ниже исторических значений.")
            return ', '.join(flags) if flags else None, '\n'.join(notes) if notes else None

        df[['attention_flag','notes']] = df.apply(attention_flag, axis=1, result_type='expand')

        # 8️⃣ Итоговый результат
        result = df[['flight','flight_date','flight_number','route','days_to_flight',
                     'sales_speed_ratio','sales_speed_status','total_seats','sold_so_far','remaining_seats',
                     'daily_needed','today_vs_needed','daily_needed_comparison','attention_flag','notes']]

        st.subheader("📊 Результаты")
        st.dataframe(result, use_container_width=True)

        # Рейсы, требующие внимания
        attention_df = result[result['attention_flag'].notna()]
        if not attention_df.empty:
            st.subheader("⚠️ Рейсы, требующие внимания")
            st.dataframe(attention_df[['flight','flight_date','flight_number','route','attention_flag','notes']], use_container_width=True)

        # 🔹 Графики
        fig_speed = px.bar(result,x='flight_number',y='sales_speed_ratio',color='sales_speed_status',
                           hover_data=['route','flight_date'],title="📊 Темп продаж по рейсам")
        st.plotly_chart(fig_speed,use_container_width=True)

        fig_needed = px.bar(result,x='flight_number',y='today_vs_needed',color='daily_needed_comparison',
                            hover_data=['route','flight_date','daily_needed','today'],
                            title="📈 Продажи vs необходимый план на день")
        st.plotly_chart(fig_needed,use_container_width=True)

        if not attention_df.empty:
            fig_attention = px.bar(attention_df,x='flight_number',y='today',color='attention_flag',
                                   hover_data=['route','flight_date','notes'],
                                   title="⚠️ Рейсы с аномалиями")
            st.plotly_chart(fig_attention,use_container_width=True)

        # Скачать Excel
        output=io.BytesIO()
        with pd.ExcelWriter(output,engine='xlsxwriter') as writer:
            result.to_excel(writer,index=False,sheet_name='Sales Speed')
        st.download_button(label="💾 Скачать отчёт в Excel",
                           data=output.getvalue(),
                           file_name="sales_speed_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("⬆️ Загрузите Excel файл, чтобы начать анализ.")
