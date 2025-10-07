import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="Sales Speed Analyzer", layout="wide")
st.title("📈 Анализ скорости продаж по рейсам")

uploaded_file = st.file_uploader("Загрузи Excel файл (структура как Test sales.xlsx)", type=["xlsx"])

if uploaded_file:
    # 1️⃣ Чтение без заголовков для поиска реальной строки с колонками
    df_temp = pd.read_excel(uploaded_file, engine="openpyxl", header=None)
    
    # Находим первую строку с заголовками
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
        
        # 3️⃣ Очистка названий колонок
        def clean_col(name):
            name = str(name).strip()
            name = name.lower()
            name = re.sub(r'[^a-z0-9]', '_', name)
            return name
        
        df.columns = [clean_col(c) for c in df.columns]

        # 4️⃣ Словарь соответствий
        rename_dict = {
            'flt_date_num': 'flight',  # колонка с датой, номером рейса и маршрутом
            'ind_ss_today': 'today',
            'ind_ss_yesterday': 'yesterday',
            'ind_ss_2_3_days_before': 'd_2_3',
            'ind_ss_4_6_days_before': 'd_4_6',
            'ind_ss_7_13_days_before': 'd_7_13',
            'ind_ss_last_14_days': 'd_14_plus',
            'total_seats': 'total_seats'  # если есть
        }

        df = df.rename(columns=rename_dict)

        required_columns = ['flight', 'today', 'yesterday', 'd_2_3', 'd_4_6', 'd_7_13', 'd_14_plus']
        if not all(col in df.columns for col in required_columns):
            st.error("⚠️ Файл не соответствует нужной структуре. Проверь названия колонок.")
            st.write("Найденные колонки:", df.columns.tolist())
        else:
            # 5️⃣ Разбор flight на дату, номер рейса и маршрут
            def split_flight_info(flight_str):
                try:
                    parts = flight_str.split(" - ")
                    flight_date = pd.to_datetime(parts[0], format="%Y.%m.%d", errors='coerce')  # безопасно
                    flight_number = parts[1] if len(parts) > 1 else None
                    route = parts[2] if len(parts) > 2 else None
                    return pd.Series([flight_date, flight_number, route])
                except:
                    return pd.Series([pd.NaT, None, None])

            df[['flight_date', 'flight_number', 'route']] = df['flight'].apply(split_flight_info)

            # Удаляем строки, где дата рейса не распознана
            df = df[df['flight_date'].notna()]

            # 6️⃣ Определяем сегодняшнюю дату и дни до вылета
            today = datetime.today().date()
            df['days_to_flight'] = (df['flight_date'].dt.date - today).dt.days
            df.loc[df['days_to_flight'] < 1, 'days_to_flight'] = 1  # чтобы не делить на ноль

            # 7️⃣ Расчёт темпа продаж
            def calc_sales_speed(row):
                early = row['d_14_plus'] + row['d_7_13']
                recent = row['d_4_6'] + row['d_2_3'] + row['yesterday'] + row['today']
                return 0 if early == 0 else recent / early

            df['sales_speed_ratio'] = df.apply(calc_sales_speed, axis=1)

            # 8️⃣ Классификация скорости
            def classify_speed(ratio):
                if ratio > 1.2:
                    return "🟢 Быстро"
                elif ratio < 0.8:
                    return "🔴 Медленно"
                else:
                    return "🟡 Нормально"

            df['sales_speed_status'] = df['sales_speed_ratio'].apply(classify_speed)

            # 9️⃣ Определение аномалий
            def attention_flag(row):
                flags = []
                notes = []

                # резкий рост
                if row['today'] > max(row['d_14_plus'], 1) * 1.5:
                    flags.append('Сегодня')
                    notes.append("Сегодня продажи резко выше исторических значений.")
                if row['yesterday'] > max(row['d_14_plus'], 1) * 1.5:
                    flags.append('Вчера')
                    notes.append("Вчера продажи резко выше исторических значений.")
                if row['d_2_3'] > max(row['d_14_plus'], 1) * 1.5:
                    flags.append('2-3 дня назад')
                    notes.append("Продажи 2-3 дня назад резко выше исторических значений.")

                # резкое падение
                if row['today'] < max(row['d_14_plus'], 1) * 0.5:
                    flags.append('Сегодня')
                    notes.append("Сегодня продажи резко ниже исторических значений.")
                if row['yesterday'] < max(row['d_14_plus'], 1) * 0.5:
                    flags.append('Вчера')
                    notes.append("Вчера продажи резко ниже исторических значений.")
                if row['d_2_3'] < max(row['d_14_plus'], 1) * 0.5:
                    flags.append('2-3 дня назад')
                    notes.append("Продажи 2-3 дня назад резко ниже исторических значений.")

                return ', '.join(flags) if flags else None, '\n'.join(notes) if notes else None

            df[['attention_flag', 'notes']] = df.apply(attention_flag, axis=1, result_type='expand')

            # 🔟 Сравнение с планом продаж
            if 'total_seats' in df.columns:
                df['avg_daily_plan'] = df['total_seats'] / df['days_to_flight']
                df['today_vs_plan'] = df['today'] - df['avg_daily_plan']
                def daily_status(row):
                    if row['today_vs_plan'] > 0:
                        return "🟢 Опережаем"
                    elif row['today_vs_plan'] < 0:
                        return "🔴 Отстаем"
                    else:
                        return "🟡 В графике"
                df['daily_sales_comparison'] = df.apply(daily_status, axis=1)
            else:
                df['daily_sales_comparison'] = "Нет данных о total_seats"

            # 1️⃣1️⃣ Итоговый результат
            result = df[['flight', 'flight_date', 'flight_number', 'route', 'days_to_flight',
                         'sales_speed_ratio', 'sales_speed_status', 'daily_sales_comparison',
                         'attention_flag', 'notes']]

            st.subheader("📊 Результаты")
            st.dataframe(result, use_container_width=True)

            # 1️⃣2️⃣ Рейсы, требующие внимания
            attention_df = result[result['attention_flag'].notna()]
            if not attention_df.empty:
                st.subheader("⚠️ Рейсы, требующие внимания")
                st.dataframe(attention_df[['flight', 'flight_date', 'flight_number', 'route',
                                           'attention_flag', 'notes']], use_container_width=True)

            # 1️⃣3️⃣ Скачать Excel
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
