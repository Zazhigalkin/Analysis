import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import numpy as np

st.set_page_config(page_title="Анализ темпов продаж", page_icon="📈", layout="wide")
st.title("📈 Анализ темпа продаж по рейсам с учётом сегодняшней даты")

with st.expander("ℹ️ ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ И ЛОГИКЕ АНАЛИЗА"):
    st.markdown("""
    ### 📋 Требуемые колонки в файле:
    - `flt_date&num` - Дата и номер рейса (формат: "2024.01.15 - SU123 - MOSCOW-SOCHI")
    - `Ind SS` - Всего продано билетов (без жёстких блоков)
    - `Ind SS yesterday` - Продано вчера
    - `Cap` - Вместимость самолёта
    - `LF` - Load Factor (загрузка рейса)
    - `Av seats` - Доступные места

    ### 🎯 ЛОГИКА КЛАССИФИКАЦИИ СТАТУСОВ (в порядке приоритета):
    1. **🟢 По плану** …
    2. **⚪ До рейса ещё далеко** …
    3. **🔵 Перепродажа** …
    4. **🔴 Отстаём** …

    ### 📊 РАСЧЕТНЫЕ ПОКАЗАТЕЛИ:
    - **days_to_flight** = дни до вылета (минимум 1 день)
    - **remaining_seats** = доступные места (Av seats)
    - **sold_total** = Cap - Av seats  ← 🔧 теперь включает блоки
    - **daily_needed** = remaining_seats / days_to_flight
    - **diff_vs_plan** = sold_yesterday - daily_needed
    """)

with st.expander("📁 ФОРМАТ ЗАГРУЖАЕМОГО EXCEL-ФАЙЛА"):
    st.markdown("""
    ### Обязательные колонки в Excel файле:
    
    | Название колонки | Описание | Пример |
    |------------------|----------|---------|
    | `flt_date&num` | Дата и номер рейса | `2024.01.15 - SU123 - MOSCOW-SOCHI` |
    | `LF` | Load Factor (загрузка) | `98,7%` |
    | `Av fare` | Средний тариф | `4 039` |
    | `Cap` | Вместимость самолета | `227` |
    | `Av seats` | Доступные места | `3` |
    | `Ind SS` | Всего продано билетов (без блоков) | `214` |
    | `Ind SS today` | Продано сегодня | `0` |
    | `Ind SS yesterday` | **Продано вчера** | `4` |
    | … | … | … |

    ### 🔍 Для анализа используются:
    - **`flt_date&num`**, **`Cap`**, **`Av seats`**, **`Ind SS`**, **`Ind SS yesterday`**, **`LF`**
    """)

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

# 🔧 утилита для аккуратного парсинга числовых строк
def _to_num(x):
    return pd.to_numeric(
        str(x).replace('\u00a0','').replace(' ', '').replace(',', '.').replace('%',''),
        errors='coerce'
    )

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        if df.empty:
            st.error("❌ Файл пустой")
            st.stop()

        # Переименования
        df = df.rename(columns={
            'flt_date&num': 'flight',
            'Ind SS': 'sold_total_raw',          # 🔧 оригинальные продажи из файла (без блоков)
            'Ind SS yesterday': 'sold_yesterday',
            'Cap': 'total_seats',
            'LF': 'load_factor',
            # 'Av seats' оставим как есть, но ниже создадим 'av_seats'
        })

        # 🔧 проверяем наличие Av seats
        required_columns = ['flight','sold_total_raw','sold_yesterday','total_seats','load_factor','Av seats']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"❌ Отсутствуют обязательные колонки: {', '.join(missing_columns)}")
            st.write("📋 Найденные колонки:", df.columns.tolist())
            st.stop()

        st.write("📊 Первые 5 строк исходных данных:")
        st.dataframe(df.head())

        # Разбор flight
        flight_split = df['flight'].str.split(" - ", n=2, expand=True)
        if flight_split.shape[1] < 3:
            st.error("❌ Неверный формат колонки 'flight'. Ожидается: 'Дата - Номер рейса - Маршрут'")
            st.stop()
        df[['flight_date_str','flight_number','route']] = flight_split.iloc[:, :3]

        # Даты
        original_count = len(df)
        df['flight_date'] = pd.to_datetime(df['flight_date_str'], format="%Y.%m.%d", errors='coerce')
        invalid_dates = df['flight_date'].isna().sum()
        if invalid_dates > 0:
            st.warning(f"⚠️ Найдено {invalid_dates} строк с некорректной датой. Они будут исключены из анализа.")
            df = df[df['flight_date'].notna()].copy()
        if df.empty:
            st.error("❌ После обработки дат не осталось корректных записей")
            st.stop()
        st.success(f"✅ Обработано {len(df)} из {original_count} записей")

        # 🔧 Чистим числа
        df['total_seats'] = _to_num(df['total_seats'])
        df['sold_total_raw'] = _to_num(df['sold_total_raw']).fillna(0)
        df['sold_yesterday'] = _to_num(df['sold_yesterday']).fillna(0)
        df['av_seats'] = _to_num(df['Av seats']).fillna(np.nan)  # важно видеть пропуски
        df['load_factor_num'] = _to_num(df['load_factor']).fillna(0)

        # Контроль пропусков Av seats
        if df['av_seats'].isna().any():
            missing = int(df['av_seats'].isna().sum())
            st.warning(f"⚠️ У {missing} строк отсутствует значение 'Av seats' — они будут исключены, т.к. нужно корректно посчитать проданные с блоками.")
            df = df[df['av_seats'].notna()].copy()
        if df.empty:
            st.error("❌ После исключения строк без 'Av seats' не осталось данных")
            st.stop()

        # Дата анализа
        today = datetime.today().date()

        df['days_to_flight'] = df['flight_date'].apply(lambda x: max((x.date() - today).days, 1))
        past_flights = df[df['flight_date'].dt.date < today]
        if not past_flights.empty:
            st.warning(f"⚠️ Найдено {len(past_flights)} рейсов, которые уже вылетели. Они будут исключены.")
            df = df[df['flight_date'].dt.date >= today].copy()
        if df.empty:
            st.error("❌ После исключения вылетевших рейсов не осталось записей")
            st.stop()

        # 🔧 НОВЫЙ sold_total: считаем «продано с учётом блоков» = Cap - Av seats
        df['sold_total'] = (df['total_seats'] - df['av_seats']).clip(lower=0)

        # 🔧 Остаток мест = Av seats
        df['remaining_seats'] = df['av_seats'].clip(lower=0)

        # План на день
        df['daily_needed'] = np.where(
            (df['days_to_flight'] > 0) & (df['remaining_seats'] > 0),
            df['remaining_seats'] / df['days_to_flight'],
            0
        )

        # Отклонение vs план
        df['diff_vs_plan'] = df['sold_yesterday'] - df['daily_needed']

        def classify(row):
            days_to_flight = row['days_to_flight']
            daily_needed = row['daily_needed']
            diff = row['diff_vs_plan']
            load_factor = row['load_factor_num']
            sold_yesterday = row['sold_yesterday']

            if daily_needed < 3:
                return "🟢 По плану"
            if sold_yesterday == 0 and load_factor > 90:
                return "🟢 По плану"
            if days_to_flight > 30 and daily_needed < 4:
                if sold_yesterday > daily_needed:
                    return "🔵 Перепродажа"
                else:
                    return "⚪ До рейса ещё далеко"
            elif diff > max(5, daily_needed * 0.3):
                return "🔵 Перепродажа"
            elif abs(diff) <= max(5, daily_needed * 0.3):
                return "🟢 По плану"
            else:
                return "🔴 Отстаём"

        df['status'] = df.apply(classify, axis=1)

        # Итоговый набор колонок
        result_columns = [
            'flight', 'flight_date', 'flight_number', 'route',
            'total_seats', 'sold_total', 'sold_yesterday',
            'remaining_seats', 'days_to_flight', 'daily_needed',
            'diff_vs_plan', 'load_factor_num', 'status'
        ]
        result = df[result_columns].copy()

        # Форматирование
        result['daily_needed'] = result['daily_needed'].fillna(0).round(1)
        result['diff_vs_plan'] = result['diff_vs_plan'].fillna(0).round(1)
        result['sold_yesterday'] = result['sold_yesterday'].fillna(0).round(1)
        result['load_factor_num'] = result['load_factor_num'].fillna(0).round(1)
        result['sold_total'] = result['sold_total'].fillna(0).round(0).astype(int)
        result['remaining_seats'] = result['remaining_seats'].fillna(0).round(0).astype(int)

        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader("📊 Результаты анализа темпа продаж")
        with col2:
            status_counts = result['status'].value_counts()
            st.metric("Всего рейсов", len(result))

        status_colors = {
            "🔵 Перепродажа": "#1f77b4",
            "🟢 По плану": "#2ca02c",
            "🔴 Отстаём": "#d62728",
            "⚪ До рейса ещё далеко": "#7f7f7f"
        }
        cols = st.columns(len(status_counts))
        for i, (status, count) in enumerate(status_counts.items()):
            color = status_colors.get(status, "#7f7f7f")
            cols[i].markdown(
                f"<div style='background-color:{color}; padding:10px; border-radius:5px; color:white; text-align: center;'>"
                f"<b>{status}</b><br>{count} рейсов</div>", 
                unsafe_allow_html=True
            )

        st.subheader("🔍 Фильтры")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            selected_status = st.multiselect("Статус рейсов",
                options=result['status'].unique(),
                default=result['status'].unique())
        with col2:
            min_days = int(result['days_to_flight'].min())
            max_days = int(result['days_to_flight'].max())
            days_range = st.slider("Дни до вылета", min_days, max_days, (min_days, max_days))
        with col3:
            routes = st.multiselect("Маршруты",
                options=result['route'].unique(),
                default=result['route'].unique())
        with col4:
            min_load = float(result['load_factor_num'].min())
            max_load = float(result['load_factor_num'].max())
            load_factor_range = st.slider("Load Factor (%)", min_load, max_load, (min_load, max_load))

        filtered_result = result[
            (result['status'].isin(selected_status)) &
            (result['days_to_flight'].between(days_range[0], days_range[1])) &
            (result['route'].isin(routes)) &
            (result['load_factor_num'].between(load_factor_range[0], load_factor_range[1]))
        ]

        def highlight_rows(row):
            if row['status'] == '🔴 Отстаём':
                return ['background-color: #ffcccc'] * len(row)
            elif row['status'] == '🔵 Перепродажа':
                return ['background-color: #ccffcc'] * len(row)
            elif row['status'] == '⚪ До рейса ещё далеко':
                return ['background-color: #f0f0f0'] * len(row)
            else:
                return [''] * len(row)

        display_columns = [
            'flight', 'flight_date', 'flight_number', 'route',
            'total_seats', 'sold_total', 'sold_yesterday',
            'remaining_seats', 'days_to_flight', 'daily_needed',
            'diff_vs_plan', 'load_factor_num', 'status'
        ]
        formatted_result = filtered_result[display_columns].copy()

        display_df = formatted_result.copy()
        display_df['flight_date'] = display_df['flight_date'].dt.strftime('%Y-%m-%d')
        display_df['daily_needed'] = display_df['daily_needed'].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "0.0")
        display_df['diff_vs_plan'] = display_df['diff_vs_plan'].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "0.0")
        display_df['sold_yesterday'] = display_df['sold_yesterday'].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "0.0")
        display_df['load_factor_num'] = display_df['load_factor_num'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "0.0%")
        display_df = display_df.rename(columns={'load_factor_num': 'load_factor'})

        st.dataframe(
            display_df.style.apply(highlight_rows, axis=1),
            use_container_width=True,
            height=400
        )

        attention_df = filtered_result[filtered_result['status'].isin(["🔴 Отстаём", "🔵 Перепродажа"])]
        if not attention_df.empty:
            st.subheader("⚠️ Рейсы, требующие внимания")
            if 'checked_flights' not in st.session_state:
                st.session_state.checked_flights = {}

            for status in ["🔴 Отстаём", "🔵 Перепродажа"]:
                status_df = attention_df[attention_df['status'] == status]
                if not status_df.empty:
                    with st.expander(f"{status} ({len(status_df)} рейсов)"):

                        st.info("✅ Отмечайте рейсы, которые уже проверили")
                        for idx, row in status_df.iterrows():
                            flight_key = row['flight']
                            if flight_key not in st.session_state.checked_flights:
                                st.session_state.checked_flights[flight_key] = False

                            c1, c2 = st.columns([1, 10])
                            with c1:
                                is_checked = st.checkbox(
                                    "",
                                    value=st.session_state.checked_flights[flight_key],
                                    key=f"check_{flight_key}",
                                    help="Отметьте, если рейс уже проверен"
                                )
                                st.session_state.checked_flights[flight_key] = is_checked
                            with c2:
                                st.markdown(f"~~{row['flight']}~~ ✅" if is_checked else f"**{row['flight']}**")
                                st.markdown(f"""
                                - **Маршрут:** {row['route']}
                                - **Дата вылета:** {row['flight_date'].strftime('%Y-%m-%d')}
                                - **Продано вчера:** {row['sold_yesterday']:.1f} 
                                - **Необходимый темп:** {row['daily_needed']:.1f}
                                - **Отклонение:** {row['diff_vs_plan']:.1f}
                                - **Load Factor:** {row['load_factor_num']:.1f}%
                                - **Дней до вылета:** {row['days_to_flight']}
                                """)
                            st.markdown("---")

                        checked_count = sum(1 for key in st.session_state.checked_flights 
                                            if st.session_state.checked_flights[key] and 
                                            key in status_df['flight'].values)
                        total_count = len(status_df)
                        st.metric(
                            f"Проверено рейсов ({status})", 
                            f"{checked_count} из {total_count}",
                            delta=f"{checked_count/total_count*100:.1f}%" if total_count > 0 else "0%"
                        )

        # Экспорт
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_to_export = result.copy()
            result_to_export['flight_date'] = result_to_export['flight_date'].dt.strftime('%Y-%m-%d')
            result_to_export['daily_needed'] = result_to_export['daily_needed'].fillna(0).round(1)
            result_to_export['diff_vs_plan'] = result_to_export['diff_vs_plan'].fillna(0).round(1)
            result_to_export['sold_yesterday'] = result_to_export['sold_yesterday'].fillna(0).round(1)
            result_to_export['load_factor'] = result_to_export['load_factor_num'].fillna(0).round(1)
            result_to_export = result_to_export.drop('load_factor_num', axis=1)
            result_to_export.to_excel(writer, index=False, sheet_name='Sales_Speed')

        st.download_button(
            label="💾 Скачать полный отчёт в Excel",
            data=output.getvalue(),
            file_name=f"sales_speed_analysis_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Ошибка при обработке файла: {str(e)}")
        import traceback
        st.error(f"Детали ошибки: {traceback.format_exc()}")
        st.info("ℹ️ Пожалуйста, проверьте формат файла и попробуйте снова.")

else:
    st.info("⬆️ Загрузите Excel файл, чтобы начать анализ.")
