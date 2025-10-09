import streamlit as st
import pandas as pd
from datetime import datetime
import io
import numpy as np
import time

# ----------------------- REAL-TIME USER TRACKING -----------------------
def init_user_tracking():
    """Инициализация отслеживания пользователей в реальном времени"""
    if 'active_users' not in st.session_state:
        st.session_state.active_users = {}
    if 'user_last_activity' not in st.session_state:
        st.session_state.user_last_activity = {}

def get_user_session_id():
    """Генерируем ID сессии для пользователя"""
    if 'user_session_id' not in st.session_state:
        st.session_state.user_session_id = f"session_{int(time.time())}_{np.random.randint(1000, 9999)}"
    return st.session_state.user_session_id

def update_user_activity():
    """Обновляем активность пользователя"""
    init_user_tracking()
    session_id = get_user_session_id()
    current_time = time.time()
    
    # Обновляем время активности текущего пользователя
    st.session_state.user_last_activity[session_id] = current_time
    
    # Очищаем неактивных пользователей (неактивны более 5 минут)
    timeout = 300  # 5 минут в секундах
    active_users = {}
    for user_id, last_active in st.session_state.user_last_activity.items():
        if current_time - last_active < timeout:
            active_users[user_id] = last_active
    
    st.session_state.user_last_activity = active_users
    st.session_state.active_users_count = len(active_users)

# ----------------------- PAGE CONFIG -----------------------
st.set_page_config(page_title="Анализ темпов продаж", page_icon="📈", layout="wide")

# Обновляем активность пользователя при каждой загрузке
update_user_activity()

# ----------------------- HEADER WITH USER COUNT -----------------------
col1, col2, col3 = st.columns([3, 1, 1])
with col1:
    st.title("📈 Анализ темпа продаж по рейсам с учётом сегодняшней даты")
with col2:
    # Красивое отображение текущих пользователей
    active_users = st.session_state.get('active_users_count', 1)
    st.metric(
        "👥 Использует сейчас", 
        f"{active_users} чел",
        delta=None
    )
with col3:
    st.metric("🕐 Время обновления", datetime.now().strftime('%H:%M'))

# ----------------------- HELPERS ---------------------------
def clean_number(s: pd.Series) -> pd.Series:
    """Парсит числа: убирает пробелы/неразрывные пробелы, заменяет запятую на точку, возвращает float."""
    return pd.to_numeric(
        s.astype(str)
         .str.replace('\u00a0','', regex=False)
         .str.replace(' ','', regex=False)
         .str.replace(',','.', regex=False),
        errors='coerce'
    )

def clean_percent(s: pd.Series) -> pd.Series:
    """Парсит проценты: снимает %, чистит разделители и масштабирует при необходимости (0..1 -> *100)."""
    val = pd.to_numeric(
        s.astype(str)
         .str.replace('%','', regex=False)
         .str.replace('\u00a0','', regex=False)
         .str.replace(' ','', regex=False)
         .str.replace(',','.', regex=False),
        errors='coerce'
    )
    if val.mean(skipna=True) < 2:
        val = val * 100
    return val

# ----------------------- UI: INSTRUCTIONS ------------------
with st.expander("ℹ️ ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ И ЛОГИКЕ АНАЛИЗА"):
    st.markdown("""
    ### 📋 Требуемые колонки в файле:
    - `flt_date&num` — Дата и номер рейса (формат: `YYYY.MM.DD - XX123 - ORG-DST`)
    - `Ind SS` — Всего продано билетов (как в отчёте; **без жёстких блоков**)
    - `Ind SS yesterday` — Продано вчера
    - `Cap` — Вместимость самолёта
    - `LF` — Load Factor (загрузка)
    - `Av seats` — Доступные места

    ### 🎯 ЛОГИКА КЛАССИФИКАЦИИ СТАТУСОВ (в порядке приоритета):
    // ... остальная инструкция без изменений ...
    """)

# ----------------------- ОСТАЛЬНОЙ КОД БЕЗ ИЗМЕНЕНИЙ -----------------------
with st.expander("📁 ФОРМАТ ЗАГРУЖАЕМОГО EXCEL-ФАЙЛА"):
    st.markdown("""
    ### Обязательные колонки:
    // ... существующий контент ...
    """)

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if uploaded_file:
    try:
        # Обновляем активность при работе с файлом
        update_user_activity()
        
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        if df.empty:
            st.error("❌ Файл пустой")
            st.stop()

        # ---------- Rename ----------
        df = df.rename(columns={
            'flt_date&num': 'flight',
            'Ind SS': 'sold_total_raw',          
            'Ind SS yesterday': 'sold_yesterday',
            'Cap': 'total_seats',
            'LF': 'load_factor',
        })

        required_columns = ['flight','sold_total_raw','sold_yesterday','total_seats','load_factor','Av seats']
        missing_columns = [c for c in required_columns if c not in df.columns]
        if missing_columns:
            st.error(f"❌ Отсутствуют обязательные колонки: {', '.join(missing_columns)}")
            st.write("📋 Найденные колонки:", df.columns.tolist())
            st.stop()

        st.write("📊 Первые 5 строк исходных данных:")
        st.dataframe(df.head())

        # ---------- Split "flight" ----------
        flight_split = df['flight'].str.split(" - ", n=2, expand=True)
        if flight_split.shape[1] < 3:
            st.error("❌ Неверный формат 'flt_date&num'. Ожидается: 'YYYY.MM.DD - XX123 - ROUTE'")
            st.stop()
        df[['flight_date_str','flight_number','route']] = flight_split.iloc[:, :3]

        # ---------- Dates ----------
        original_count = len(df)
        df['flight_date'] = pd.to_datetime(df['flight_date_str'], format="%Y.%m.%d", errors='coerce')
        invalid_dates = df['flight_date'].isna().sum()
        if invalid_dates > 0:
            st.warning(f"⚠️ Найдено {invalid_dates} строк с некорректной датой. Они будут исключены.")
            df = df[df['flight_date'].notna()].copy()
        if df.empty:
            st.error("❌ После обработки дат не осталось корректных записей")
            st.stop()
        st.success(f"✅ Обработано {len(df)} из {original_count} записей")

        # ---------- Clean numbers (vectorized) ----------
        df['total_seats']     = clean_number(df['total_seats'])
        df['sold_total_raw']  = clean_number(df['sold_total_raw']).fillna(0)
        df['sold_yesterday']  = clean_number(df['sold_yesterday']).fillna(0)
        df['av_seats']        = clean_number(df['Av seats'])
        df['load_factor_num'] = clean_percent(df['load_factor']).fillna(0)

        # ---------- Require Av seats ----------
        if df['av_seats'].isna().any():
            missing = int(df['av_seats'].isna().sum())
            st.warning(f"⚠️ У {missing} строк нет 'Av seats' — они исключены (нужно для учёта блоков).")
            df = df[df['av_seats'].notna()].copy()
        if df.empty:
            st.error("❌ После исключения строк без 'Av seats' не осталось данных")
            st.stop()

        # ---------- Analysis date ----------
        today = datetime.today().date()
        df['days_to_flight'] = df['flight_date'].apply(lambda x: max((x.date() - today).days, 1))

        # Убираем уже вылетевшие
        past_flights = df[df['flight_date'].dt.date < today]
        if not past_flights.empty:
            st.warning(f"⚠️ Найдено {len(past_flights)} рейсов, которые уже вылетели. Они будут исключены.")
            df = df[df['flight_date'].dt.date >= today].copy()
        if df.empty:
            st.error("❌ После исключения вылетевших рейсов не осталось записей")
            st.stop()

        # ---------- NEW sold_total & remaining ----------
        df['sold_total']      = (df['total_seats'] - df['av_seats']).clip(lower=0)
        df['remaining_seats'] = df['av_seats'].clip(lower=0)

        # ---------- Daily plan & diffs ----------
        df['daily_needed'] = np.where(
            (df['days_to_flight'] > 0) & (df['remaining_seats'] > 0),
            df['remaining_seats'] / df['days_to_flight'],
            0
        )
        df['diff_vs_plan'] = df['sold_yesterday'] - df['daily_needed']

        # ---------- Classification (УПРОЩЕННАЯ ЛОГИКА - БЕЗ "ДАЛЕКО ДО РЕЙСА") ----------
        def classify(row):
            days_to_flight = row['days_to_flight']
            daily_needed = row['daily_needed']
            diff = row['diff_vs_plan']
            load_factor = row['load_factor_num']
            sold_yesterday = row['sold_yesterday']
            
            if daily_needed < 3 and diff > 10:
                return "🔵 Перепродажа"
            
            if sold_yesterday == 0 and load_factor > 90:
                return "🟢 По плану"
            
            if daily_needed < 3:
                return "🟢 По плану"
            
            if days_to_flight > 30 and daily_needed < 4:
                if sold_yesterday > daily_needed:
                    return "🔵 Перепродажа"
                else:
                    return "🟢 По плану"
            
            if diff > max(5, daily_needed * 0.3):
                return "🔵 Перепродажа"
            elif abs(diff) <= max(5, daily_needed * 0.3):
                return "🟢 По плану"
            else:
                return "🔴 Отстаём"

        df['status'] = df.apply(classify, axis=1)

        # ---------- Result set ----------
        result_columns = [
            'flight', 'flight_date', 'flight_number', 'route',
            'total_seats', 'sold_total', 'sold_yesterday',
            'remaining_seats', 'days_to_flight', 'daily_needed',
            'diff_vs_plan', 'load_factor_num', 'status'
        ]
        result = df[result_columns].copy()

        # ---------- Formatting ----------
        result['daily_needed']     = result['daily_needed'].fillna(0).round(1)
        result['diff_vs_plan']     = result['diff_vs_plan'].fillna(0).round(1)
        result['sold_yesterday']   = result['sold_yesterday'].fillna(0).round(1)
        result['load_factor_num']  = result['load_factor_num'].fillna(0).round(1)
        result['sold_total']       = result['sold_total'].fillna(0).round(0).astype(int)
        result['remaining_seats']  = result['remaining_seats'].fillna(0).round(0).astype(int)

        # ----------------------- SUMMARY HEADER -----------------------
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.subheader("📊 Результаты анализа темпа продаж")
        with col2:
            status_counts = result['status'].value_counts()
            st.metric("Всего рейсов", len(result))
        with col3:
            # Обновляем счетчик при отображении результатов
            update_user_activity()
            active_users = st.session_state.get('active_users_count', 1)
            st.metric("👥 Активных пользователей", active_users)

        status_colors = {
            "🔵 Перепродажа": "#1f77b4",
            "🟢 По плану": "#2ca02c",
            "🔴 Отстаём": "#d62728"
        }
        cols = st.columns(len(status_counts))
        for i, (status, count) in enumerate(status_counts.items()):
            color = status_colors.get(status, "#7f7f7f")
            cols[i].markdown(
                f"<div style='background-color:{color}; padding:10px; border-radius:5px; color:white; text-align: center;'>"
                f"<b>{status}</b><br>{count} рейсов</div>", 
                unsafe_allow_html=True
            )

        # ----------------------- ОСТАЛЬНОЙ КОД БЕЗ ИЗМЕНЕНИЙ -----------------------
        # ... (фильтры, таблица, блок внимания, экспорт) ...

    except Exception as e:
        st.error(f"❌ Ошибка при обработке файла: {str(e)}")
        import traceback
        st.error(f"Детали ошибки: {traceback.format_exc()}")
        st.info("ℹ️ Пожалуйста, проверьте формат файла и попробуйте снова.")

else:
    st.info("⬆️ Загрузите Excel файл, чтобы начать анализ.")
    
    # Показываем счетчик даже когда файл не загружен
    update_user_activity()
    active_users = st.session_state.get('active_users_count', 1)
    st.sidebar.markdown("---")
    st.sidebar.metric("👥 Использует сейчас", active_users)
