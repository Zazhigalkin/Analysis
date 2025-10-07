import streamlit as st
import pandas as pd
import io


st.set_page_config(page_title="Sales Speed Analyzer", layout="wide")
st.title("📈 Анализ скорости продаж по рейсам")

# 1️⃣ Загрузка файла
uploaded_file = st.file_uploader("Загрузи Excel файл (структура как Test sales.xlsx)", type=["xlsx"])

if uploaded_file:
    # 2️⃣ Чтение Excel
    df = pd.read_excel(uploaded_file)

    # 3️⃣ Приведение колонок к рабочим названиям
    df = df.rename(columns={
        'flt_date&num': 'flight',
        'Ind SS today': 'today',
        'Ind SS yesterday': 'yesterday',
        'Ind SS 2-3 days before': 'd_2_3',
        'Ind SS 4-6 days before': 'd_4_6',
        'Ind SS 7-13 days before': 'd_7_13',
        'Ind SS last 14 days': 'd_14_plus'
    })

    if uploaded_file:
    # Показать предпросмотр данных для отладки
    raw_df = pd.read_excel(uploaded_file)
    st.write("Предпросмотр данных:", raw_df.head(3))
    
    # Позволить пользователю выбрать строку с заголовками
    header_row = st.number_input("Номер строки с заголовками", min_value=0, value=2)
    df = pd.read_excel(uploaded_file, header=header_row)
    
    # Показать доступные колонки для сопоставления
    st.write("Доступные колонки:", list(df.columns))

    required_columns = ['flight', 'today', 'yesterday', 'd_2_3', 'd_4_6', 'd_7_13', 'd_14_plus']
    if not all(col in df.columns for col in required_columns):
        st.error("⚠️ Файл не соответствует нужной структуре. Проверь названия колонок.")
    else:
        # 4️⃣ Расчёт темпа продаж
        def calc_sales_speed(row):
            early = row['d_14_plus'] + row['d_7_13']
            recent = row['d_4_6'] + row['d_2_3'] + row['yesterday'] + row['today']
            if early == 0:
                return 0
            return recent / early

        df['sales_speed_ratio'] = df.apply(calc_sales_speed, axis=1)

        # 5️⃣ Классификация
        def classify_speed(ratio):
            if ratio > 1.2:
                return "🟢 Быстро"
            elif ratio < 0.8:
                return "🔴 Медленно"
            else:
                return "🟡 Нормально"

        df['sales_speed_status'] = df['sales_speed_ratio'].apply(classify_speed)

        result = df[['flight', 'sales_speed_ratio', 'sales_speed_status']]

        st.subheader("📊 Результаты")
        st.dataframe(result, use_container_width=True)

        # 6️⃣ Возможность скачать Excel с результатами
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result.to_excel(writer, index=False, sheet_name='Sales Speed')
        st.download_button(
            label="💾 Скачать отчёт в Excel",
            data=output.getvalue(),
            file_name="sales_speed_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("⬆️ Загрузите Excel файл, чтобы начать анализ.")
