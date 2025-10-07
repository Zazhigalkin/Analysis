import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Sales Speed Analyzer", layout="wide")
st.title("📈 Анализ скорости продаж по рейсам")

# 1️⃣ Загрузка файла
uploaded_file = st.file_uploader("Загрузи Excel файл (структура как Test sales.xlsx)", type=["xlsx"])

if uploaded_file:
    # 2️⃣ Чтение Excel
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    
    # 3️⃣ Очистка названий колонок (убираем пробелы, переводим в lower case, заменяем спецсимволы)
    def clean_col(name):
        name = str(name).strip()                 # убрать пробелы в начале/конце
        name = name.lower()                      # перевести в нижний регистр
        name = re.sub(r'[^a-z0-9]', '_', name)  # заменить всё, кроме букв/цифр, на _
        return name

    df.columns = [clean_col(c) for c in df.columns]

    # 4️⃣ Словарь соответствий "чистое имя -> рабочее имя"
    rename_dict = {
        'flt_date_num': 'flight',
        'ind_ss_today': 'today',
        'ind_ss_yesterday': 'yesterday',
        'ind_ss_2_3_days_before': 'd_2_3',
        'ind_ss_4_6_days_before': 'd_4_6',
        'ind_ss_7_13_days_before': 'd_7_13',
        'ind_ss_last_14_days': 'd_14_plus'
    }

    df = df.rename(columns=rename_dict)

    required_columns = ['flight', 'today', 'yesterday', 'd_2_3', 'd_4_6', 'd_7_13', 'd_14_plus']
    if not all(col in df.columns for col in required_columns):
        st.error("⚠️ Файл не соответствует нужной структуре. Проверь названия колонок.")
        st.write("Найденные колонки:", df.columns.tolist())
    else:
        # 5️⃣ Расчёт темпа продаж
        def calc_sales_speed(row):
            early = row['d_14_plus'] + row['d_7_13']
            recent = row['d_4_6'] + row['d_2_3'] + row['yesterday'] + row['today']
            return 0 if early == 0 else recent / early

        df['sales_speed_ratio'] = df.apply(calc_sales_speed, axis=1)

        # 6️⃣ Классификация
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

        # 7️⃣ Возможность скачать Excel с результатами
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
