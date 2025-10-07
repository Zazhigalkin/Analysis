import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Sales Speed Analyzer", layout="wide")
st.title("📈 Анализ скорости продаж по рейсам")

uploaded_file = st.file_uploader("Загрузи Excel файл (структура как Test sales.xlsx)", type=["xlsx"])

if uploaded_file:
    # 1️⃣ Чтение без заголовков для поиска реальной строки с колонками
    df_temp = pd.read_excel(uploaded_file, engine="openpyxl", header=None)
    
    header_row_idx = None
    for i, row in df_temp.iterrows():
        if row.notna().sum() >= len(row)/2:
            header_row_idx = i
            break
    
    if header_row_idx is None:
        st.error("Не удалось найти заголовки в файле Excel")
    else:
        # 2️⃣ Чтение с найденными заголовками
        df = pd.read_excel(uploaded_file, engine="openpyxl", header=header_row_idx)
        
        # Очистка названий колонок
        def clean_col(name):
            name = str(name).strip()
            name = name.lower()
            name = re.sub(r'[^a-z0-9]', '_', name)
            return name
        
        df.columns = [clean_col(c) for c in df.columns]

        # Словарь соответствий
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
            # 3️⃣ Расчёт темпа продаж
            def calc_sales_speed(row):
                early = row['d_14_plus'] + row['d_7_13']
                recent = row['d_4_6'] + row['d_2_3'] + row['yesterday'] + row['today']
                return 0 if early == 0 else recent / early

            df['sales_speed_ratio'] = df.apply(calc_sales_speed, axis=1)

            # 4️⃣ Классификация
            def classify_speed(ratio):
                if ratio > 1.2:
                    return "🟢 Быстро"
                elif ratio < 0.8:
                    return "🔴 Медленно"
                else:
                    return "🟡 Нормально"

            df['sales_speed_status'] = df['sales_speed_ratio'].apply(classify_speed)

            # 5️⃣ Проверка на внимание (резкие изменения по дням)
            def attention_flag(row):
                flags = []
                if row['today'] > row['d_14_plus'] * 1.5:  # Сегодня резко больше чем раньше
                    flags.append('Сегодня')
                if row['yesterday'] > row['d_14_plus'] * 1.5:
                    flags.append('Вчера')
                if row['d_2_3'] > row['d_14_plus'] * 1.5:
                    flags.append('2-3 дня назад')
                return ', '.join(flags) if flags else None

            df['attention_flag'] = df.apply(attention_flag, axis=1)

            # 6️⃣ Итоговый результат
            result = df[['flight', 'sales_speed_ratio', 'sales_speed_status', 'attention_flag']]

            st.subheader("📊 Результаты")
            st.dataframe(result, use_container_width=True)

            # 7️⃣ Отдельно показываем рейсы, на которые нужно обратить внимание
            attention_df = result[result['attention_flag'].notna()]
            if not attention_df.empty:
                st.subheader("⚠️ Рейсы, требующие внимания")
                st.dataframe(attention_df[['flight', 'attention_flag']], use_container_width=True)

            # 8️⃣ Скачать Excel с результатами
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
