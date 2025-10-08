import streamlit as st
import pandas as pd
from datetime import datetime
import io
import numpy as np

# ----------------------- PAGE CONFIG -----------------------
st.set_page_config(page_title="Анализ темпов продаж", page_icon="📈", layout="wide")
st.title("📈 Анализ темпа продаж по рейсам с учётом сегодняшней даты")

# ----------------------- HELPERS ---------------------------
def clean_number(s: pd.Series) -> pd.Series:
    """Парсит числа: убирает пробелы, неразрывные пробелы, заменяет запятую на точку."""
    return pd.to_numeric(
        s.astype(str)
         .str.replace('\u00a0','', regex=False)
         .str.replace(' ','', regex=False)
         .str.replace(',','.', regex=False),
        errors='coerce'
    )

def clean_percent(s: pd.Series) -> pd.Series:
    """Парсит проценты: снимает %, чистит разделители и масштабирует при необходимости."""
    val = pd.to_numeric(
        s.astype(str)
         .str.replace('%','', regex=False)
         .str.replace('\u00a0','', regex=False)
         .str.replace(' ','', regex=False)
         .str.replace(',','.', regex=False),
        errors='coerce'
    )
    # если значения выглядят как 0.97 вместо 97 -> умножаем на 100
    if val.mean(skipna=True) < 2:
        val = val * 100
    return val

# ----------------------- UI: INSTRUCTIONS ------------------
with st.expander("ℹ️ ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ И ЛОГИКЕ АНАЛИЗА"):
    st.markdown("""
    ### 📋 Требуемые колонки:
    - `flt_date&num` — Дата и номер рейса (напр. `2024.01.15 - SU123 - MOSCOW-SOCHI`)
    - `Ind SS` — Всего продано билетов (без жёстких блоков)
    - `Ind SS yesterday` — Продано вчера
    - `Cap` — Вместимость самолёта
    - `LF` — Load Factor (%)
    - `Av seats` — Доступные места
    """)

uploaded_file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if uploaded_file:
    try:
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
            'LF': 'load_factor'
        })

        required_columns = ['flight','sold_total_raw','sold_yesterday','total_seats','load_factor','Av seats']
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            st.error(f"❌ Отсутствуют обязательные колонки: {', '.join(missing)}")
            st.stop()

        # ---------- Split flight ----------
        flight_split = df['flight'].str.split(" - ", n=2, expand=True)
        df[['flight_date_str','flight_number','route']] = flight_split.iloc[:, :3]
        df['flight_date'] = pd.to_datetime(df['flight_date_str'], format="%Y.%m.%d", errors='coerce')
        df = df[df['flight_date'].notna()].copy()

        # ---------- Clean numeric ----------
        df['total_seats']     = clean_number(df['total_seats'])
        df['sold_total_raw']  = clean_number(df['sold_total_raw']).fillna(0)
        df['sold_yesterday']  = clean_number(df['sold_yesterday']).fillna(0)
        df['av_seats']        = clean_number(df['Av seats'])
        df['load_factor_num'] = clean_percent(df['load_factor']).fillna(0)  # ✅ теперь в диапазоне 0–100

        # ---------- Base calculations ----------
        today = datetime.today().date()
        df['days_to_flight'] = df['flight_date'].apply(lambda x: max((x.date() - today).days, 1))
        df = df[df['flight_date'].dt.date >= today].copy()

        df['sold_total'] = (df['total_seats'] - df['av_seats']).clip(lower=0)
        df['remaining_seats'] = df['av_seats'].clip(lower=0)
        df['daily_needed'] = np.where(
            (df['days_to_flight'] > 0) & (df['remaining_seats'] > 0),
            df['remaining_seats'] / df['days_to_flight'],
            0
        )
        df['diff_vs_plan'] = df['sold_yesterday'] - df['daily_needed']

        # ---------- Classification ----------
        def classify(row):
            dn, diff, lf, sy, d2f = row['daily_needed'], row['diff_vs_plan'], row['load_factor_num'], row['sold_yesterday'], row['days_to_flight']
            if dn < 3 or (sy == 0 and lf > 90):
                return "🟢 По плану"
            if d2f > 30 and dn < 4:
                return "🔵 Перепродажа" if sy > dn else "⚪ До рейса ещё далеко"
            if diff > max(5, dn*0.3): return "🔵 Перепродажа"
            if abs(diff) <= max(5, dn*0.3): return "🟢 По плану"
            return "🔴 Отстаём"

        df['status'] = df.apply(classify, axis=1)

        # ---------- Prepare output ----------
        result = df[['flight','flight_date','flight_number','route','total_seats',
                     'sold_total','sold_yesterday','remaining_seats','days_to_flight',
                     'daily_needed','diff_vs_plan','load_factor_num','status']].copy()

        # ---------- Display ----------
        result['flight_date'] = result['flight_date'].dt.strftime('%Y-%m-%d')
        result['load_factor'] = result['load_factor_num'].round(0).astype(int).astype(str) + '%'  # ✅ формат 100%
        st.dataframe(
            result[['flight','route','flight_date','load_factor','sold_total','remaining_seats','status']],
            use_container_width=True
        )

        # ---------- Export ----------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export = result.copy()
            export.to_excel(writer, index=False, sheet_name='Sales_Speed')

        st.download_button(
            "💾 Скачать Excel",
            data=output.getvalue(),
            file_name=f"sales_speed_analysis_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        import traceback
        st.error(f"❌ Ошибка: {str(e)}")
        st.error(traceback.format_exc())
else:
    st.info("⬆️ Загрузите Excel файл для анализа.")
