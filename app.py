import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import numpy as np

st.set_page_config(page_title="Анализ темпов продаж", page_icon="📈", layout="wide")

st.title("📈 Анализ темпа продаж по рейсам с учётом сегодняшней даты")

# Добавляем описание для пользователя
with st.expander("ℹ️ Инструкция по использованию"):
    st.markdown("""
    **Требуемые колонки в файле:**
    - `flt_date&num` - Дата и номер рейса (формат: "2024.01.15 - SU123 - MOSCOW-SOCHI")
    - `Ind SS` - Всего продано билетов
    - `Ind SS yesterday` - Продано вчера
    - `Cap` - Вместимость самолёта
    - `LF` - Load Factor (загрузка рейса)
    
    **Статусы рейсов:**
    - 🔵 Перепродажа - значительно превысили дневной план
    - 🟢 По плану - продажи в пределах нормы
    - 🔴 Отстаём - недовыполнение плана
    - ⚪ До рейса ещё далеко - рейс через >30 дней с малым дневным планом
    """)

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        
        # Валидация данных
        if df.empty:
            st.error("❌ Файл пустой")
            st.stop()
            
        # Переименование колонок
        df = df.rename(columns={
            'flt_date&num': 'flight',
            'Ind SS': 'sold_total',
            'Ind SS yesterday': 'sold_yesterday',
            'Cap': 'total_seats',
            'LF': 'load_factor'
        })

        required_columns = ['flight','sold_total','sold_yesterday','total_seats','load_factor']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"❌ Отсутствуют обязательные колонки: {', '.join(missing_columns)}")
            st.write("📋 Найденные колонки:", df.columns.tolist())
            st.stop()

        st.write("📊 Первые 5 строк данных:")
        st.dataframe(df.head())

        # Обработка данных
        # Разделение flight на дату, номер рейса и маршрут
        flight_split = df['flight'].str.split(" - ", n=2, expand=True)
        
        if flight_split.shape[1] < 3:
            st.error("❌ Неверный формат колонки 'flight'. Ожидается: 'Дата - Номер рейса - Маршрут'")
            st.stop()
            
        df[['flight_date_str','flight_number','route']] = flight_split.iloc[:, :3]
        
        # Преобразование даты
        original_count = len(df)
        df['flight_date'] = pd.to_datetime(df['flight_date_str'], format="%Y.%m.%d", errors='coerce')
        
        # Проверка корректности дат
        invalid_dates = df['flight_date'].isna().sum()
        if invalid_dates > 0:
            st.warning(f"⚠️ Найдено {invalid_dates} строк с некорректной датой. Они будут исключены из анализа.")
            df = df[df['flight_date'].notna()].copy()
        
        if df.empty:
            st.error("❌ После обработки дат не осталось корректных записей")
            st.stop()
            
        st.success(f"✅ Обработано {len(df)} из {original_count} записей")

        # Сегодняшняя дата
        today = datetime.today().date()
        
        # Дни до вылета - исправленная версия
        df['days_to_flight'] = df['flight_date'].apply(
            lambda x: max((x.date() - today).days, 1)
        )

        # Проверка на рейсы, которые уже вылетели
        past_flights = df[df['flight_date'].dt.date < today]
        if not past_flights.empty:
            st.warning(f"⚠️ Найдено {len(past_flights)} рейсов, которые уже вылетели. Они будут исключены.")
            df = df[df['flight_date'].dt.date >= today].copy()

        if df.empty:
            st.error("❌ После исключения вылетевших рейсов не осталось записей")
            st.stop()

        # Основные расчеты
        df['remaining_seats'] = df['total_seats'] - df['sold_total']
        
        # Защита от деления на ноль
        df['daily_needed'] = np.where(
            df['days_to_flight'] > 0,
            df['remaining_seats'] / df['days_to_flight'],
            0
        )
        
        df['diff_vs_plan'] = df['sold_yesterday'] - df['daily_needed']

        # ПРАВИЛЬНОЕ преобразование load_factor - убираем только символы, но не делим на 100
        df['load_factor_num'] = df['load_factor'].astype(str).str.replace(',', '.').str.rstrip('%').astype(float)
        # Load Factor уже в процентах (98.7), так что оставляем как есть

        # Улучшенная классификация с проверкой Load Factor
        def classify(row):
            days_to_flight = row['days_to_flight']
            daily_needed = row['daily_needed']
            diff = row['diff_vs_plan']
            load_factor = row['load_factor_num']  # Уже в процентах (98.7)
            sold_yesterday = row['sold_yesterday']
            
            # Проверка: если вчера не было продаж, но Load Factor > 90% - считаем по плану
            if sold_yesterday == 0 and load_factor > 90:
                return "🟢 По плану"
            
            # Долгосрочные рейсы с малым спросом
            if days_to_flight > 30 and daily_needed < 4:
                return "⚪ До рейса ещё далеко"
            # Значительное превышение плана
            elif diff > max(5, daily_needed * 0.3):  # 5 или 30% от дневного плана
                return "🔵 Перепродажа"
            # Небольшое отклонение от плана
            elif abs(diff) <= max(5, daily_needed * 0.3):
                return "🟢 По плану"
            # Значительное недовыполнение
            else:
                return "🔴 Отстаём"

        df['status'] = df.apply(classify, axis=1)

        # Итоговая таблица с правильным округлением
        result_columns = [
            'flight', 'flight_date', 'flight_number', 'route', 
            'total_seats', 'sold_total', 'sold_yesterday', 
            'remaining_seats', 'days_to_flight', 'daily_needed', 
            'diff_vs_plan', 'load_factor_num', 'status'
        ]
        
        result = df[result_columns].copy()
        
        # Округление до 1 знака после запятой для daily_needed и diff_vs_plan
        result['daily_needed'] = result['daily_needed'].round(1)
        result['diff_vs_plan'] = result['diff_vs_plan'].round(1)
        result['load_factor_num'] = result['load_factor_num'].round(1)

        # Визуализация
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.subheader("📊 Результаты анализа темпа продаж")
            
        with col2:
            status_counts = result['status'].value_counts()
            st.metric("Всего рейсов", len(result))

        # Статистика по статусам
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

        # Фильтрация данных
        st.subheader("🔍 Фильтры")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            selected_status = st.multiselect(
                "Статус рейсов",
                options=result['status'].unique(),
                default=result['status'].unique()
            )
        
        with col2:
            min_days = int(result['days_to_flight'].min())
            max_days = int(result['days_to_flight'].max())
            days_range = st.slider(
                "Дни до вылета",
                min_days, max_days, (min_days, max_days)
            )
            
        with col3:
            routes = st.multiselect(
                "Маршруты",
                options=result['route'].unique(),
                default=result['route'].unique()
            )
            
        with col4:
            min_load_factor = float(result['load_factor_num'].min())
            max_load_factor = float(result['load_factor_num'].max())
            load_factor_range = st.slider(
                "Load Factor (%)",
                min_load_factor, max_load_factor, (min_load_factor, max_load_factor)
            )

        # Применение фильтров
        filtered_result = result[
            (result['status'].isin(selected_status)) &
            (result['days_to_flight'].between(days_range[0], days_range[1])) &
            (result['route'].isin(routes)) &
            (result['load_factor_num'].between(load_factor_range[0], load_factor_range[1]))
        ]

        # Функция для подсветки строк
        def highlight_rows(row):
            if row['status'] == '🔴 Отстаём':
                return ['background-color: #ffcccc'] * len(row)
            elif row['status'] == '🔵 Перепродажа':
                return ['background-color: #ccffcc'] * len(row)
            elif row['status'] == '⚪ До рейса ещё далеко':
                return ['background-color: #f0f0f0'] * len(row)
            else:
                return [''] * len(row)

        # Форматирование числовых колонок для отображения
        display_columns = [
            'flight', 'flight_date', 'flight_number', 'route', 
            'total_seats', 'sold_total', 'sold_yesterday', 
            'remaining_seats', 'days_to_flight', 'daily_needed', 
            'diff_vs_plan', 'load_factor_num', 'status'
        ]
        
        formatted_result = filtered_result[display_columns].copy()
        
        # Форматируем отображение чисел
        display_df = formatted_result.copy()
        display_df['daily_needed'] = display_df['daily_needed'].apply(lambda x: f"{x:.1f}")
        display_df['diff_vs_plan'] = display_df['diff_vs_plan'].apply(lambda x: f"{x:.1f}")
        display_df['load_factor_num'] = display_df['load_factor_num'].apply(lambda x: f"{x:.1f}%")  # Теперь будет 68.0% вместо 0.68%
        
        # Переименовываем колонки для отображения
        display_df = display_df.rename(columns={'load_factor_num': 'load_factor'})
        
        st.dataframe(
            display_df.style.apply(highlight_rows, axis=1),
            use_container_width=True,
            height=400
        )

        # Рейсы, требующие внимания
        attention_df = filtered_result[filtered_result['status'].isin(["🔴 Отстаём", "🔵 Перепродажа"])]
        if not attention_df.empty:
            st.subheader("⚠️ Рейсы, требующие внимания")
            
            for status in ["🔴 Отстаём", "🔵 Перепродажа"]:
                status_df = attention_df[attention_df['status'] == status]
                if not status_df.empty:
                    with st.expander(f"{status} ({len(status_df)} рейсов)"):
                        display_cols = ['flight', 'flight_date', 'route', 'sold_yesterday', 'daily_needed', 'diff_vs_plan', 'days_to_flight', 'load_factor_num']
                        display_data = status_df[display_cols].copy()
                        display_data['daily_needed'] = display_data['daily_needed'].apply(lambda x: f"{x:.1f}")
                        display_data['diff_vs_plan'] = display_data['diff_vs_plan'].apply(lambda x: f"{x:.1f}")
                        display_data['load_factor_num'] = display_data['load_factor_num'].apply(lambda x: f"{x:.1f}%")
                        display_data = display_data.rename(columns={'load_factor_num': 'load_factor'})
                        st.dataframe(
                            display_data,
                            use_container_width=True
                        )

        # Дополнительная аналитика
        with st.expander("📈 Дополнительная аналитика"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Средний дневной темп", f"{filtered_result['daily_needed'].mean():.1f}")
                st.metric("Всего осталось мест", f"{filtered_result['remaining_seats'].sum():.0f}")
                st.metric("Средний Load Factor", f"{filtered_result['load_factor_num'].mean():.1f}%")
                
            with col2:
                st.metric("Медианное отклонение от плана", f"{filtered_result['diff_vs_plan'].median():.1f}")
                st.metric("Средние дни до вылета", f"{filtered_result['days_to_flight'].mean():.0f}")
                st.metric("Рейсов с LF > 90%", f"{len(filtered_result[filtered_result['load_factor_num'] > 90])}")
                
            with col3:
                st.metric("Рейсов отстаёт от плана", 
                         len(filtered_result[filtered_result['status'] == '🔴 Отстаём']))
                st.metric("Рейсов с перепродажей", 
                         len(filtered_result[filtered_result['status'] == '🔵 Перепродажа']))
                st.metric("Рейсов по плану", 
                         len(filtered_result[filtered_result['status'] == '🟢 По плану']))

        # Скачать Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Для Excel форматируем числа правильно
            result_to_export = result.copy()
            result_to_export['daily_needed'] = result_to_export['daily_needed'].round(1)
            result_to_export['diff_vs_plan'] = result_to_export['diff_vs_plan'].round(1)
            result_to_export['load_factor'] = result_to_export['load_factor_num'].round(1)
            result_to_export = result_to_export.drop('load_factor_num', axis=1)
            result_to_export.to_excel(writer, index=False, sheet_name='Sales_Speed')
            
            # Добавляем лист с аналитикой
            summary = pd.DataFrame({
                'Метрика': [
                    'Всего рейсов', 
                    'Отстают', 
                    'Перепродажа', 
                    'По плану', 
                    'Далёкие рейсы',
                    'Общий остаток мест',
                    'Средний Load Factor',
                    'Рейсов с LF > 90%'
                ],
                'Значение': [
                    len(result),
                    len(result[result['status'] == '🔴 Отстаём']),
                    len(result[result['status'] == '🔵 Перепродажа']),
                    len(result[result['status'] == '🟢 По плану']),
                    len(result[result['status'] == '⚪ До рейса ещё далеко']),
                    result['remaining_seats'].sum(),
                    f"{result['load_factor_num'].mean():.1f}%",
                    len(result[result['load_factor_num'] > 90])
                ]
            })
            summary.to_excel(writer, index=False, sheet_name='Аналитика')
            
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
