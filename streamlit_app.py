from pathlib import Path
from tempfile import TemporaryDirectory
from datetime import date
import time
import uuid

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly import express as px

from transaction_uploader import (
    OUTPUT_COLUMNS,
    normalize_sheet,
    read_excel_sheets,
    upload_dataframe,
    get_database_url,
    get_db_connection,
    soft_delete_transactions,
    restore_transactions,
    update_transaction,
    purge_old_deleted_transactions,
)
from pdf_export import generate_monthly_pdf

st.set_page_config(page_title="Importador de transacciones", layout="wide")


@st.cache_data(ttl=300)
def load_db_data() -> pd.DataFrame:
    import psycopg2

    query = """
        select
          id as "ID",
          fecha as "Fecha",
          descripcion as "Descripción",
          monto as "Monto",
          tipo as "Tipo",
          origen as "Origen",
          categoria as "Categoria"
        from movimientos
        where activo is distinct from false
        order by fecha desc, id desc
    """

    with psycopg2.connect(get_database_url()) as conn:
        return pd.read_sql_query(query, conn)


@st.cache_data(ttl=300)
def load_deleted_data() -> pd.DataFrame:
    import psycopg2

    query = """
        select
          id as "ID",
          fecha as "Fecha",
          descripcion as "Descripción",
          monto as "Monto",
          tipo as "Tipo",
          origen as "Origen",
          categoria as "Categoria"
        from movimientos
        where activo = false
        order by fecha desc, id desc
    """

    with psycopg2.connect(get_database_url()) as conn:
        return pd.read_sql_query(query, conn)


@st.cache_data(ttl=3600)
def load_available_months() -> list[tuple[int, int]]:
    """Return (year, month) tuples for months present in the DB, newest first."""
    import psycopg2

    query = """
        select distinct
          extract(year from fecha)::int  as year,
          extract(month from fecha)::int as month
        from movimientos
        where activo is distinct from false
        order by year desc, month desc
    """
    with psycopg2.connect(get_database_url()) as conn:
        rows = pd.read_sql_query(query, conn)
    return [(int(r["year"]), int(r["month"])) for _, r in rows.iterrows()]


@st.cache_data(ttl=300)
def load_db_data_for_month(year: int, month: int) -> pd.DataFrame:
    """Load all active transactions for a specific year/month."""
    import psycopg2
    from calendar import monthrange

    last_day = monthrange(year, month)[1]
    query = """
        select
          id as "ID",
          fecha as "Fecha",
          descripcion as "Descripción",
          monto as "Monto",
          tipo as "Tipo",
          origen as "Origen",
          categoria as "Categoria"
        from movimientos
        where activo is distinct from false
          and fecha >= %(start)s
          and fecha <= %(end)s
        order by fecha desc, id desc
    """
    params = {
        "start": date(year, month, 1),
        "end": date(year, month, last_day),
    }
    with psycopg2.connect(get_database_url()) as conn:
        return pd.read_sql_query(query, conn, params=params)


def run_startup_tasks() -> None:
    """Run once per session: ensure schema and purge old deleted records."""
    if st.session_state.get("_startup_done"):
        return
    try:
        purged = purge_old_deleted_transactions(days=30)
        if purged:
            load_deleted_data.clear()
    except Exception:
        pass
    st.session_state["_startup_done"] = True


def format_date_column(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=str)

    if not isinstance(series, pd.Series):
        series = pd.Series(series)

    # 1. Identificamos qué filas son números (ej: 45431 de Excel/Sheets) y cuáles son texto ('2024-05-18')
    es_numero = pd.to_numeric(series, errors="coerce").notna()
    
    # Creamos una serie vacía para guardar las fechas correctas
    fechas_finales = pd.Series(index=series.index, dtype='datetime64[ns]')
    
    # 2. Convertimos SOLO los números usando el calendario de Excel
    if es_numero.any():
        fechas_finales[es_numero] = pd.to_datetime(
            pd.to_numeric(series[es_numero]), 
            unit="D", 
            origin="1899-12-30"
        )
    
    # 3. Convertimos SOLO los textos usando formato ISO explícito.
    # Los parsers de transaction_uploader ya emiten strings "YYYY-MM-DD".
    # NO usar dayfirst=True aquí: pandas interpreta "YYYY-MM-DD" con dayfirst
    # como "YYYY-DD-MM" e invierte mes y día (ej: 2026-05-04 → 2026-04-05).
    if (~es_numero).any():
        fechas_finales[~es_numero] = pd.to_datetime(
            series[~es_numero],
            format="%Y-%m-%d",
            errors="coerce",
        )

    # Finalmente, pasamos todo a texto limpio YYYY-MM-DD
    return fechas_finales.dt.strftime("%Y-%m-%d").fillna("")


def compute_totals(frame: pd.DataFrame) -> tuple[float, float, float]:
    ingresos = frame.loc[frame["Tipo"].astype(str).str.lower() == "ingreso", "Monto"].sum()
    egresos = frame.loc[frame["Tipo"].astype(str).str.lower() == "egreso", "Monto"].sum()
    return ingresos, egresos, ingresos - egresos


def format_local_number(value) -> str:
    if pd.isna(value):
        return "0"

    return f"{int(value):,}".replace(",", ".")


def format_amount_series(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=str)

    return pd.to_numeric(series, errors="coerce").fillna(0).apply(format_local_number)


def detect_currency_from_text(text: str) -> str:
    upper_text = str(text).upper()
    return "USD" if any(token in upper_text for token in ["USD", "U$S", "DOLARES", "DÓLARES"]) else "ARS"


def format_money_premium(amount, origin: str = "") -> str:
    currency = detect_currency_from_text(origin)
    symbol = "$" if currency == "ARS" else "U$S"
    numeric = pd.to_numeric(amount, errors="coerce")
    if pd.isna(numeric):
        numeric = 0.0
    formatted = f"{numeric:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{symbol} {formatted}"


def format_money_series_with_origin(amounts: pd.Series, origins: pd.Series) -> pd.Series:
    if amounts is None:
        return pd.Series(dtype=str)

    if origins is None:
        origins = pd.Series([""] * len(amounts), index=amounts.index)

    return pd.DataFrame({"amount": amounts, "origin": origins}).apply(
        lambda row: format_money_premium(row["amount"], row["origin"]),
        axis=1,
    )


def render_file_summary(ingresos: float, egresos: float, neto: float) -> None:
    color = "#28a745" if neto >= 0 else "#dc3545"
    st.markdown(
        f"<h2 style='text-align:center; color:{color}; margin-bottom:1rem;'>"
        f"Neto: {format_local_number(neto)}"
        f"</h2>",
        unsafe_allow_html=True,
    )

    left_col, right_col = st.columns(2)
    with left_col:
        st.markdown(
            f"<h3 style='color:#28a745; margin-top:0;'>Ingresos: {format_local_number(ingresos)}</h3>",
            unsafe_allow_html=True,
        )
    with right_col:
        st.markdown(
            f"<h3 style='color:#dc3545; margin-top:0;'>Egresos: {format_local_number(egresos)}</h3>",
            unsafe_allow_html=True,
        )


def render_upload_summary_cards(ingresos: float, egresos: float, neto: float, currency: str) -> None:
    symbol = "$" if currency == "ARS" else "U$S"
    net_color = "#00e676" if neto >= 0 else "#ff1744"
    ingresos_color = "#00e676"
    egresos_color = "#ff1744"

    st.markdown(
        """
        <style>
        .summary-wrap {
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 18px;
            padding: 1rem;
            margin-bottom: 1rem;
        }
        .summary-card {
            background: rgba(18, 18, 18, 0.72);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 16px;
            padding: 1rem 1.1rem;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.18);
        }
        .summary-card.center {
            text-align: center;
        }
        .summary-label {
            font-size: 0.85rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            opacity: 0.75;
            margin-bottom: 0.25rem;
        }
        .summary-value {
            font-weight: 800;
            line-height: 1.1;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="summary-wrap">
          <div class="summary-card center">
            <div class="summary-label">Saldo Neto</div>
            <div class="summary-value" style="font-size: 2rem; color: {net_color};">{symbol} {format_money_premium(neto, currency).split(' ', 1)[1]}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left_col, right_col = st.columns(2)
    with left_col:
        st.markdown(
            f"""
            <div class="summary-card">
              <div class="summary-label">Ingresos</div>
              <div class="summary-value" style="font-size: 1.5rem; color: {ingresos_color};">{format_money_premium(ingresos, currency)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right_col:
        st.markdown(
            f"""
            <div class="summary-card">
              <div class="summary-label">Egresos</div>
              <div class="summary-value" style="font-size: 1.5rem; color: {egresos_color};">{format_money_premium(egresos, currency)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_currency_summary_block(currency: str, ingresos: float, egresos: float, neto: float) -> None:
    label = "Pesos (ARS)" if currency == "ARS" else "Dólares (USD)"
    st.markdown(f"### {label}")
    with st.container():
        render_upload_summary_cards(ingresos, egresos, neto, currency)


def render_individual_file_card(file_name: str, ingresos: float, egresos: float, neto: float, currency: str) -> None:
    net_color = "#00e676" if neto >= 0 else "#ff1744"
    st.markdown(
        f"""
        <div style="background: rgba(18, 18, 18, 0.72); border: 1px solid rgba(255,255,255,0.08); border-radius: 14px; padding: 0.85rem 1rem; margin-bottom: 0.6rem;">
          <div style="font-weight: 700; margin-bottom: 0.35rem;">{file_name}</div>
          <div style="display:flex; gap: 1rem; flex-wrap: wrap; font-size: 0.95rem;">
            <span style="color:#00e676;">Ingresos: {format_money_premium(ingresos, currency)}</span>
            <span style="color:#ff1744;">Egresos: {format_money_premium(egresos, currency)}</span>
            <span style="color:{net_color};">Neto: {format_money_premium(neto, currency)}</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_account_summary_card(account_name: str, ingresos: float, egresos: float, neto: float, currency: str) -> None:
    label = f"Resumen de Cuenta: {account_name} ({currency})"
    render_individual_file_card(label, ingresos, egresos, neto, currency)


def resolve_file_account_label(frame: pd.DataFrame) -> str:
    if "Origen" not in frame.columns:
        return "Sin cuenta"

    values = [str(value).strip() for value in frame["Origen"].dropna().astype(str).tolist() if str(value).strip()]
    unique_values = list(dict.fromkeys(values))

    if not unique_values:
        return "Sin cuenta"
    if len(unique_values) == 1:
        return unique_values[0]
    return ", ".join(unique_values[:2]) + ("..." if len(unique_values) > 2 else "")


def style_tipo(val):
    if str(val).lower() == "ingreso":
        return "background-color: rgba(34, 197, 94, 0.14); color: #9ef0b2; font-weight: 700; border-radius: 999px; padding: 0.15rem 0.45rem;"
    if str(val).lower() == "egreso":
        return "background-color: rgba(239, 68, 68, 0.14); color: #ffb3b3; font-weight: 700; border-radius: 999px; padding: 0.15rem 0.45rem;"
    return ""


def parse_date_series(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype="datetime64[ns]")

    if not isinstance(series, pd.Series):
        series = pd.Series(series)

    numeric = pd.to_numeric(series, errors="coerce")
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    numeric_mask = numeric.notna()

    if numeric_mask.any():
        parsed.loc[numeric_mask] = pd.to_datetime(
            numeric.loc[numeric_mask],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )

    if (~numeric_mask).any():
        parsed.loc[~numeric_mask] = pd.to_datetime(series.loc[~numeric_mask], errors="coerce", dayfirst=True)

    return parsed


def normalize_dashboard_frame(df: pd.DataFrame) -> pd.DataFrame:
    dashboard_df = df.copy()

    if "Categoría" in dashboard_df.columns and "Categoria" not in dashboard_df.columns:
        dashboard_df.rename(columns={"Categoría": "Categoria"}, inplace=True)

    for col in ["Origen", "Tipo", "Categoria", "Descripción"]:
        if col not in dashboard_df.columns:
            dashboard_df[col] = ""

    dashboard_df["Fecha_dt"] = parse_date_series(dashboard_df.get("Fecha"))
    dashboard_df["Fecha_display"] = dashboard_df["Fecha_dt"].dt.date
    dashboard_df["Monto"] = pd.to_numeric(dashboard_df.get("Monto"), errors="coerce").fillna(0.0)
    dashboard_df["Origen"] = dashboard_df["Origen"].fillna("").astype(str)
    dashboard_df["Cuenta"] = dashboard_df["Origen"]
    dashboard_df["Tipo"] = dashboard_df["Tipo"].fillna("").astype(str)
    dashboard_df["Descripción"] = dashboard_df["Descripción"].fillna("").astype(str)
    dashboard_df = dashboard_df.drop(columns=["Categoria"], errors="ignore")
    dashboard_df["Moneda"] = dashboard_df["Origen"].str.upper().apply(
        lambda x: "USD" if any(token in x for token in ["USD", "U$S", "DOLARES", "DÓLARES"]) else "ARS"
    )

    return dashboard_df


def calculate_dashboard_totals(frame: pd.DataFrame) -> tuple[float, float, float]:
    tipo = frame.get("Tipo", pd.Series(dtype=str)).astype(str).str.lower()
    montos = pd.to_numeric(frame.get("Monto"), errors="coerce").fillna(0.0).abs()
    ingresos = montos[tipo == "ingreso"].sum()
    egresos = montos[tipo == "egreso"].sum()
    return ingresos, egresos, ingresos - egresos


def _currency_meta(currency: str) -> tuple[str, str]:
    currency = (currency or "ARS").upper()
    return ("Pesos (ARS)", "$") if currency == "ARS" else ("Dólares (USD)", "U$S")


def _format_currency_value(amount: float, symbol: str) -> str:
    numeric = pd.to_numeric(amount, errors="coerce")
    if pd.isna(numeric):
        numeric = 0.0
    formatted = f"{numeric:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{symbol} {formatted}"


def _dashboard_css() -> None:
    st.markdown(
        """
        <style>
        .dash-shell {
            background: rgba(10, 12, 18, 0.55);
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 18px;
            padding: 1rem 1rem 0.5rem 1rem;
            margin-bottom: 1rem;
        }
        .dash-filter-label {
            font-size: 0.78rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            opacity: 0.72;
            margin-bottom: 0.35rem;
        }
        .dash-kpi {
            background: linear-gradient(180deg, rgba(23, 27, 36, 0.95), rgba(16, 18, 25, 0.95));
            border: 1px solid rgba(255, 255, 255, 0.07);
            border-radius: 12px;
            padding: 1rem 1.05rem;
            box-shadow: 0 12px 30px rgba(0, 0, 0, 0.18);
            height: 100%;
        }
        .dash-kpi-label {
            font-size: 0.78rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            opacity: 0.72;
            margin-bottom: 0.35rem;
        }
        .dash-kpi-value {
            font-size: 1.45rem;
            font-weight: 800;
            line-height: 1.05;
        }
        .dash-kpi-value.neto {
            font-size: 2rem;
        }
        .dash-kpi-meta {
            font-size: 0.82rem;
            opacity: 0.75;
            margin-top: 0.35rem;
        }
        .dash-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.25rem 0.6rem;
            border-radius: 999px;
            font-size: 0.76rem;
            font-weight: 700;
            margin-top: 0.5rem;
        }
        .dash-badge.up { background: rgba(34, 197, 94, 0.12); color: #9ef0b2; }
        .dash-badge.down { background: rgba(239, 68, 68, 0.12); color: #ffb3b3; }
        .dash-badge.neutral { background: rgba(148, 163, 184, 0.12); color: #dbe4f0; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_card(title: str, value: str, delta: str | None = None, help_text: str | None = None) -> None:
    st.metric(label=title, value=value, delta=delta, help=help_text)


def _render_metric_cards_css() -> None:
    st.markdown(
        """
        <style>
        .metric-card {
            background: #1E1E1E;
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 12px;
            padding: 1.1rem 1.15rem;
            height: 100%;
            box-sizing: border-box;
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.18);
        }
        .metric-card.income {
            background: rgba(0, 255, 0, 0.03);
            border-left: 3px solid rgba(34, 197, 94, 0.65);
        }
        .metric-card.expense {
            background: rgba(255, 0, 0, 0.03);
            border-left: 3px solid rgba(239, 68, 68, 0.65);
        }
        .metric-card.neto {
            background: #1E1E1E;
        }
        .metric-label {
            font-size: 0.95rem;
            font-weight: 800;
            color: #F5F7FA;
            letter-spacing: 0.02em;
            margin-bottom: 0.55rem;
        }
        .metric-value {
            font-size: 2.1rem;
            font-weight: 900;
            line-height: 1;
            color: #FFFFFF;
        }
        .metric-trend {
            margin-top: 0.6rem;
            font-size: 0.88rem;
            font-weight: 700;
            opacity: 0.8;
        }
        .metric-trend.up { color: #7CFF9B; }
        .metric-trend.down { color: #FF8A8A; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard_kpis(frame: pd.DataFrame, currency: str, vertical: bool = False) -> None:
    ingresos, egresos, neto = calculate_dashboard_totals(frame)
    symbol = _currency_meta(currency)[1]
    net_delta = "Tendencia positiva" if neto >= 0 else "Tendencia negativa"

    _render_metric_cards_css()

    if vertical:
        st.markdown(
            f"""
            <div class="metric-card neto">
              <div class="metric-label">Saldo Neto</div>
              <div class="metric-value">{_format_currency_value(neto, symbol)}</div>
              <div class="metric-trend {'up' if neto >= 0 else 'down'}">{net_delta}</div>
            </div>
            <div class="metric-card income" style="margin-top:0.75rem;">
              <div class="metric-label">Ingresos</div>
              <div class="metric-value">{_format_currency_value(ingresos, symbol)}</div>
            </div>
            <div class="metric-card expense" style="margin-top:0.75rem;">
              <div class="metric-label">Egresos</div>
              <div class="metric-value">{_format_currency_value(egresos, symbol)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f"""
                <div class="metric-card neto">
                  <div class="metric-label">Saldo Neto</div>
                  <div class="metric-value">{_format_currency_value(neto, symbol)}</div>
                  <div class="metric-trend {'up' if neto >= 0 else 'down'}">{net_delta}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                f"""
                <div class="metric-card income">
                  <div class="metric-label">Ingresos</div>
                  <div class="metric-value">{_format_currency_value(ingresos, symbol)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f"""
                <div class="metric-card expense">
                  <div class="metric-label">Egresos</div>
                  <div class="metric-value">{_format_currency_value(egresos, symbol)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


_PERIOD_OPTIONS = ["Mes Actual", "Mes Anterior", "Últimos 3 Meses", "Año Actual", "Rango Personalizado"]


def _resolve_period_dates(period: str, today: date) -> tuple[date, date]:
    from calendar import monthrange

    if period == "Mes Actual":
        start = today.replace(day=1)
        end = today.replace(day=monthrange(today.year, today.month)[1])
    elif period == "Mes Anterior":
        first_of_current = today.replace(day=1)
        last_of_prev = first_of_current - pd.Timedelta(days=1)
        start = last_of_prev.replace(day=1)
        end = last_of_prev
    elif period == "Últimos 3 Meses":
        end = today.replace(day=monthrange(today.year, today.month)[1])
        # go back 3 months
        month = today.month - 2
        year = today.year
        if month <= 0:
            month += 12
            year -= 1
        start = date(year, month, 1)
    elif period == "Año Actual":
        start = date(today.year, 1, 1)
        end = date(today.year, 12, 31)
    else:
        # Rango Personalizado — caller handles it
        start = today.replace(day=1)
        end = today
    return start, end


def render_dashboard_filters(frame: pd.DataFrame, currency: str) -> tuple[pd.DataFrame, str]:
    if frame.empty or frame["Fecha_dt"].isna().all():
        return frame.iloc[0:0], ""

    today = date.today()
    currency_key = currency.lower()
    period_key = f"dashboard_period_{currency_key}"
    custom_range_key = f"dashboard_custom_range_{currency_key}"
    origin_key = f"dashboard_origin_{currency_key}"
    search_key = f"dashboard_search_{currency_key}"

    if period_key not in st.session_state:
        st.session_state[period_key] = "Mes Actual"
    if custom_range_key not in st.session_state:
        min_day = frame["Fecha_dt"].min().date()
        max_day = frame["Fecha_dt"].max().date()
        st.session_state[custom_range_key] = (min_day, max_day)
    if origin_key not in st.session_state:
        st.session_state[origin_key] = sorted([x for x in frame["Origen"].dropna().astype(str).unique().tolist() if x])
    if search_key not in st.session_state:
        st.session_state[search_key] = ""

    st.markdown("<div class='dash-shell'>", unsafe_allow_html=True)
    filter_cols = st.columns([1.6, 1.6, 2.0])

    with filter_cols[0]:
        st.markdown("<div class='dash-filter-label'>Período</div>", unsafe_allow_html=True)
        st.selectbox(
            "Período",
            options=_PERIOD_OPTIONS,
            key=period_key,
            label_visibility="collapsed",
        )
        if st.session_state[period_key] == "Rango Personalizado":
            st.date_input(
                "Rango personalizado",
                value=st.session_state[custom_range_key],
                key=custom_range_key,
                label_visibility="collapsed",
            )

    origin_options = sorted([x for x in frame["Origen"].dropna().astype(str).unique().tolist() if x])
    with filter_cols[1]:
        st.markdown("<div class='dash-filter-label'>Cuenta</div>", unsafe_allow_html=True)
        st.multiselect(
            "Cuenta",
            options=origin_options,
            default=st.session_state[origin_key],
            key=origin_key,
            label_visibility="collapsed",
        )

    with filter_cols[2]:
        st.markdown("<div class='dash-filter-label'>Buscar</div>", unsafe_allow_html=True)
        st.text_input(
            "Buscar transacción...",
            placeholder="Buscar transacción...",
            key=search_key,
            label_visibility="collapsed",
        )

    selected_period = st.session_state[period_key]
    selected_origins = st.session_state[origin_key]
    search_text = str(st.session_state[search_key]).strip().lower()

    if selected_period == "Rango Personalizado":
        raw_range = st.session_state[custom_range_key]
        if isinstance(raw_range, (list, tuple)) and len(raw_range) == 2:
            start_date, end_date = raw_range[0], raw_range[1]
        elif isinstance(raw_range, (list, tuple)) and len(raw_range) == 1:
            start_date = end_date = raw_range[0]
        else:
            start_date = end_date = raw_range if hasattr(raw_range, "year") else today
    else:
        start_date, end_date = _resolve_period_dates(selected_period, today)

    filtered = frame.copy()
    filtered = filtered[
        (filtered["Fecha_dt"].dt.date >= start_date)
        & (filtered["Fecha_dt"].dt.date <= end_date)
    ]

    if selected_origins:
        filtered = filtered[filtered["Cuenta"].isin(selected_origins)]
    else:
        filtered = filtered.iloc[0:0]

    if search_text:
        search_blob = (
            filtered[[c for c in ["Descripción", "Cuenta", "Tipo"] if c in filtered.columns]]
            .fillna("")
            .astype(str)
            .agg(" ".join, axis=1)
            .str.lower()
        )
        filtered = filtered[search_blob.str.contains(search_text, na=False)]

    st.markdown("</div>", unsafe_allow_html=True)
    return filtered, _currency_meta(currency)[1]


def build_cashflow_figure(frame: pd.DataFrame, currency: str):
    currency_label, symbol = _currency_meta(currency)
    if frame.empty or frame["Fecha_dt"].isna().all():
        fig = go.Figure()
        fig.update_layout(template="plotly_dark", height=360, margin=dict(l=10, r=10, t=40, b=10), title=f"Evolución temporal - {currency_label}")
        return fig

    working = frame.dropna(subset=["Fecha_dt"]).copy()
    tipo = working["Tipo"].astype(str).str.lower()
    monto = pd.to_numeric(working["Monto"], errors="coerce").fillna(0.0).abs()
    signed = monto.where(tipo != "egreso", -monto)
    working["Movimiento"] = signed
    daily = working.assign(Dia=working["Fecha_dt"].dt.normalize()).groupby("Dia", as_index=False)["Movimiento"].sum().sort_values("Dia")
    full_range = pd.date_range(daily["Dia"].min(), daily["Dia"].max(), freq="D")
    daily = daily.set_index("Dia").reindex(full_range, fill_value=0.0).rename_axis("Dia").reset_index()
    daily["Acumulado"] = daily["Movimiento"].cumsum()

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=daily["Dia"],
            y=daily["Acumulado"],
            mode="lines",
            line=dict(color="#7aa2ff", width=3),
            fill="tozeroy",
            fillcolor="rgba(122, 162, 255, 0.18)",
            hovertemplate=f"%{{x|%d/%m/%Y}}<br>Saldo acumulado: {symbol} %{{y:,.2f}}<extra></extra>",
        )
    )
    fig.update_layout(
        template="plotly_dark",
        height=360,
        margin=dict(l=10, r=10, t=48, b=10),
        title=f"Evolución temporal del cash flow - {currency_label}",
        xaxis_title="Fecha",
        yaxis_title="Saldo acumulado",
        hovermode="x unified",
        showlegend=False,
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.08)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.08)")
    return fig


def build_distribution_figure(frame: pd.DataFrame, currency: str):
    currency_label, symbol = _currency_meta(currency)
    if frame.empty:
        fig = go.Figure()
        fig.update_layout(template="plotly_dark", height=400, margin=dict(l=10, r=10, t=40, b=10), title=f"Distribución por cuenta - {currency_label}")
        return fig

    label_col = "Cuenta" if "Cuenta" in frame.columns else "Origen"
    base = frame.copy()
    base[label_col] = base[label_col].fillna("").astype(str).str.strip()
    base[label_col] = base[label_col].replace("", "Sin clasificar")
    values = pd.to_numeric(base["Monto"], errors="coerce").fillna(0.0).abs()
    grouped = base.assign(_valor=values).groupby(label_col, as_index=False)["_valor"].sum().sort_values("_valor", ascending=False)
    grouped = grouped[grouped["_valor"] > 0]

    if grouped.empty:
        fig = go.Figure()
        fig.update_layout(template="plotly_dark", height=400, margin=dict(l=10, r=10, t=40, b=10), title=f"Distribución por cuenta - {currency_label}")
        return fig

    colors = px.colors.sequential.Viridis[: max(len(grouped), 3)]
    fig = go.Figure(
        data=[
            go.Pie(
                labels=grouped[label_col],
                values=grouped["_valor"],
                hole=0.62,
                sort=False,
                textinfo="none",
                marker=dict(colors=colors[: len(grouped)], line=dict(color="rgba(255,255,255,0.12)", width=1)),
                hovertemplate="%{label}<br>%{percent} del total<br>Valor: " + symbol + " %{value:,.2f}<extra></extra>",
            )
        ]
    )
    fig.update_layout(
        template="plotly_dark",
        height=400,
        margin=dict(l=10, r=10, t=48, b=10),
        title=f"Distribución por cuenta - {currency_label}",
        showlegend=True,
        legend=dict(
            title_text="Cuenta",
            orientation="h",
            yanchor="top",
            y=-0.08,
            xanchor="center",
            x=0.5,
            font=dict(size=11),
        ),
    )
    return fig


def render_dashboard_detail_table(frame: pd.DataFrame, currency: str) -> None:
    table_df = frame.drop(columns=["Fecha_dt"], errors="ignore").copy()
    if "Fecha_display" in table_df.columns:
        table_df["Fecha"] = table_df["Fecha_display"]
        table_df.drop(columns=["Fecha_display"], inplace=True, errors="ignore")

    # Keep ID for tracking edits/deletes but don't expose Categoria/Origen duplicates
    has_id = "ID" in table_df.columns
    if "Cuenta" not in table_df.columns and "Origen" in frame.columns:
        table_df["Cuenta"] = frame["Origen"]
    table_df = table_df.drop(columns=["Categoria", "Origen", "Moneda"], errors="ignore")

    # Add delete checkbox column at the end
    table_df["Eliminar"] = False

    visible_columns = [col for col in ["Fecha", "Descripción", "Monto", "Tipo", "Cuenta", "ID", "Eliminar"] if col in table_df.columns]
    table_df = table_df[visible_columns]

    symbol = _currency_meta(currency)[1]
    money_format = f"{symbol} %,.2f"
    # column_order excludes ID so it stays hidden but available for tracking
    column_order = [col for col in ["Fecha", "Descripción", "Monto", "Tipo", "Cuenta", "Eliminar"] if col in table_df.columns]
    column_config: dict[str, object] = {
        "Eliminar": st.column_config.CheckboxColumn("Eliminar", help="Marcá para eliminar (soft delete)"),
    }

    if "Fecha" in table_df.columns:
        column_config["Fecha"] = st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", disabled=True)
    if "Monto" in table_df.columns:
        column_config["Monto"] = st.column_config.NumberColumn("Monto", format=money_format, min_value=0.0)
    if "Tipo" in table_df.columns:
        column_config["Tipo"] = st.column_config.SelectboxColumn("Tipo", options=["Ingreso", "Egreso"])
    if "Cuenta" in table_df.columns:
        column_config["Cuenta"] = st.column_config.TextColumn("Cuenta", disabled=True)
    if "Descripción" in table_df.columns:
        column_config["Descripción"] = st.column_config.TextColumn("Descripción")
    if "ID" in table_df.columns:
        column_config["ID"] = st.column_config.TextColumn("ID", disabled=True)

    # Unique key per currency so ARS/USD panels don't share state
    editor_key = f"detail_editor_{currency}"
    edited = st.data_editor(
        table_df,
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
        column_order=column_order,
        key=editor_key,
    )

    # Show save button only when there are actual changes in the editor
    editor_state = st.session_state.get(editor_key, {})
    has_changes = bool(
        editor_state.get("edited_rows")
        or editor_state.get("added_rows")
        or editor_state.get("deleted_rows")
    )

    if has_changes and st.button("💾 Guardar Cambios", key=f"save_changes_{currency}", type="primary"):
        ids_to_delete = edited.loc[edited["Eliminar"] == True, "ID"].tolist() if "ID" in edited.columns else []
        deleted_count = 0
        updated_count = 0

        if ids_to_delete:
            deleted_count = soft_delete_transactions(ids_to_delete)

        # Detect edited rows (compare with original, exclude delete-marked rows)
        if has_id and "ID" in edited.columns:
            original = table_df.set_index("ID")
            edited_indexed = edited[~edited["Eliminar"]].set_index("ID")
            for row_id, row in edited_indexed.iterrows():
                orig = original.loc[row_id] if row_id in original.index else None
                if orig is None:
                    continue
                kwargs: dict = {}
                if "Descripción" in row and row["Descripción"] != orig.get("Descripción"):
                    kwargs["descripcion"] = str(row["Descripción"])
                if "Monto" in row and row["Monto"] != orig.get("Monto"):
                    kwargs["monto"] = float(row["Monto"])
                if "Tipo" in row and row["Tipo"] != orig.get("Tipo"):
                    kwargs["tipo"] = str(row["Tipo"])
                if kwargs:
                    if update_transaction(str(row_id), **kwargs):
                        updated_count += 1

        if deleted_count or updated_count:
            load_db_data.clear()
            load_deleted_data.clear()
            msgs = []
            if deleted_count:
                msgs.append(f"{deleted_count} movimiento(s) eliminado(s)")
            if updated_count:
                msgs.append(f"{updated_count} movimiento(s) actualizado(s)")
            st.success(", ".join(msgs) + ".")
            st.rerun()
        else:
            st.info("No se detectaron cambios.")


def _render_pdf_section(currency: str) -> None:
    """Independent PDF generator: lets user pick a month regardless of dashboard filters."""
    month_names = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ]

    try:
        available_months = load_available_months()
    except Exception:
        available_months = []

    if not available_months:
        st.caption("Sin datos para generar PDF.")
        return

    month_labels = [f"{month_names[m - 1]} {y}" for y, m in available_months]
    pdf_month_key = f"pdf_month_select_{currency.lower()}"

    if pdf_month_key not in st.session_state:
        st.session_state[pdf_month_key] = month_labels[0]

    selected_label = st.selectbox(
        "Seleccionar mes para PDF",
        options=month_labels,
        key=pdf_month_key,
        label_visibility="visible",
    )

    selected_idx = month_labels.index(selected_label)
    sel_year, sel_month = available_months[selected_idx]
    file_name = f"informe_{currency}_{sel_year}_{sel_month:02d}.pdf"

    if st.button(f"📄 Generar PDF — {selected_label}", key=f"pdf_gen_btn_{currency.lower()}"):
        try:
            month_df_raw = load_db_data_for_month(sel_year, sel_month)
            if month_df_raw.empty:
                st.info("No hay movimientos para ese mes.")
                return
            month_df = normalize_dashboard_frame(month_df_raw)
            month_df = month_df[month_df["Moneda"] == currency]
            if month_df.empty:
                st.info(f"No hay movimientos en {currency} para ese mes.")
                return
            pdf_bytes = generate_monthly_pdf(month_df, sel_month, sel_year, currency)
        except Exception as exc:
            st.warning(f"No se pudo generar el PDF: {exc}")
            return

        if pdf_bytes is None:
            st.info("No hay movimientos para el período seleccionado.")
            return

        st.download_button(
            label=f"⬇️ Descargar {file_name}",
            data=pdf_bytes,
            file_name=file_name,
            mime="application/pdf",
            key=f"pdf_download_{currency.lower()}_{sel_year}_{sel_month:02d}",
        )


def render_dashboard_currency_panel(frame: pd.DataFrame, currency: str) -> None:
    label, _ = _currency_meta(currency)
    filtered, _ = render_dashboard_filters(frame, currency)

    # ── Title row ──────────────────────────────────────────────────────────
    title_col, badge_col = st.columns([5, 2])
    with title_col:
        st.markdown(f"### {label}")
    with badge_col:
        if not filtered.empty:
            st.markdown(
                f"<div style='margin-top:0.6rem; padding: 0.35rem 0.75rem; display:inline-flex; border-radius: 999px; background: rgba(255,255,255,0.08); font-weight:700;'>"
                f"{len(filtered)} movimientos visibles</div>",
                unsafe_allow_html=True,
            )

    if filtered.empty:
        st.info("No hay movimientos para los filtros seleccionados.")
    else:
        # ── Distribución en L: métricas apiladas | gráficos en paralelo ──────
        col_izquierda, col_derecha = st.columns([3, 9])

        with col_izquierda:
            render_dashboard_kpis(filtered, currency, vertical=True)

        with col_derecha:
            col_linea, col_dona = st.columns([6, 4])
            with col_linea:
                st.plotly_chart(build_cashflow_figure(filtered, currency), use_container_width=True)
            with col_dona:
                st.plotly_chart(build_distribution_figure(filtered, currency), use_container_width=True)

        st.markdown("### Detalle de transacciones")
        render_dashboard_detail_table(filtered, currency)

    st.markdown("---")
    st.markdown("#### 📄 Reporte PDF")
    _render_pdf_section(currency)


def render_upload_tab() -> None:
    st.subheader("Cargar movimientos")

    if "uploader_key_version" not in st.session_state:
        st.session_state["uploader_key_version"] = 0

    upload_tab, manual_tab = st.tabs(["📂 Subir Archivo Bancario", "💵 Carga Manual"])

    with upload_tab:
        st.write("Subí un archivo Excel para actualizar la base general.")
        uploaded_file = st.file_uploader(
            "Seleccioná uno o más archivos (.xls / .xlsx)",
            type=["xls", "xlsx"],
            accept_multiple_files=True,
            key=f"uploader_{st.session_state['uploader_key_version']}",
        )

    with manual_tab:
        # Show form only when there's no pending preview
        if "manual_preview" not in st.session_state:
            with st.form(key="manual_cash_form", clear_on_submit=True):
                col_fecha, col_desc = st.columns([1, 2])
                with col_fecha:
                    manual_fecha = st.date_input("Fecha", value=date.today())
                with col_desc:
                    manual_descripcion = st.text_input("Descripción")

                col_cuenta, col_moneda = st.columns(2)
                with col_cuenta:
                    manual_cuenta = st.text_input("Cuenta", value="Caja Efectivo")
                with col_moneda:
                    manual_moneda = st.selectbox("Moneda", ["ARS", "USD"])

                col_tipo, col_monto = st.columns([1, 1])
                with col_tipo:
                    manual_tipo = st.radio(
                        "Tipo de Movimiento",
                        ["Ingreso", "Egreso"],
                        horizontal=True,
                    )
                with col_monto:
                    manual_monto = st.number_input(
                        "Monto",
                        min_value=0.0,
                        step=1000.0,
                        value=None,
                        placeholder="0.00",
                    )

                submitted_manual = st.form_submit_button("Revisar y Registrar", type="primary")

            if submitted_manual:
                if not manual_descripcion.strip():
                    st.error("La descripción es obligatoria.")
                elif not manual_cuenta.strip():
                    st.error("La cuenta es obligatoria.")
                elif manual_monto is None or manual_monto <= 0:
                    st.error("El monto debe ser mayor a cero.")
                else:
                    signed_amount = float(manual_monto) * (-1 if manual_tipo == "Egreso" else 1)
                    st.session_state["manual_preview"] = {
                        "id": str(uuid.uuid4()),
                        "fecha": manual_fecha,
                        "descripcion": manual_descripcion.strip(),
                        "cuenta": manual_cuenta.strip(),
                        "monto": signed_amount,
                        "moneda": manual_moneda,
                        "tipo": manual_tipo,
                    }
                    st.rerun()

        else:
            # Preview step
            p = st.session_state["manual_preview"]
            st.markdown("#### Revisá el movimiento antes de confirmar")
            monto_fmt = f"{'🟢 +' if p['monto'] >= 0 else '🔴 -'}${abs(p['monto']):,.2f} {p['moneda']}"
            preview_df = pd.DataFrame([{
                "Fecha": p["fecha"].strftime("%d/%m/%Y"),
                "Descripción": p["descripcion"],
                "Cuenta": p["cuenta"],
                "Tipo": p["tipo"],
                "Monto": monto_fmt,
            }])
            st.dataframe(preview_df, use_container_width=True, hide_index=True)

            col_confirm, col_cancel = st.columns([1, 1])
            with col_confirm:
                if st.button("✅ Confirmar y guardar", type="primary", use_container_width=True):
                    try:
                        from transaction_uploader import TABLE_NAME, get_db_connection

                        with st.spinner("Registrando movimiento..."):
                            with get_db_connection() as conn:
                                with conn.cursor() as cur:
                                    cur.execute(
                                        f"insert into {TABLE_NAME} (id, fecha, descripcion, monto, tipo, origen, categoria) values (%s, %s, %s, %s, %s, %s, %s)",
                                        (
                                            p["id"],
                                            p["fecha"],
                                            p["descripcion"],
                                            p["monto"],
                                            p["tipo"],
                                            p["cuenta"],
                                            "",
                                        ),
                                    )
                                conn.commit()

                        del st.session_state["manual_preview"]
                        load_db_data.clear()
                        st.success("Movimiento registrado correctamente.")
                        time.sleep(1)
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Error al registrar el movimiento: {exc}")
            with col_cancel:
                if st.button("✏️ Corregir", use_container_width=True):
                    del st.session_state["manual_preview"]
                    st.rerun()

    if not uploaded_file:
        if st.session_state.get("subida_exitosa"):
            st.success("¡Planilla general actualizada correctamente!")
            return

        st.session_state.pop("upload_files_fingerprint", None)
        return

    uploaded_files = uploaded_file if isinstance(uploaded_file, list) else [uploaded_file]
    current_fingerprint = tuple((file.name, file.size) for file in uploaded_files)
    previous_fingerprint = st.session_state.get("upload_files_fingerprint")
    if previous_fingerprint != current_fingerprint:
        st.session_state["subida_exitosa"] = False
        st.session_state["upload_files_fingerprint"] = current_fingerprint

    if st.session_state.get("subida_exitosa"):
        st.success("¡Planilla general actualizada correctamente!")
        return

    with TemporaryDirectory() as tmpdir:
        frames = []
        file_summaries: dict[str, dict[str, float]] = {}

        for uploaded_file in uploaded_files:
            temp_path = Path(tmpdir) / uploaded_file.name
            temp_path.write_bytes(uploaded_file.getvalue())

            try:
                sheets = read_excel_sheets(temp_path, debug=False)
                file_frames = []

                for sheet_name, sheet_df in sheets.items():
                    normalized = normalize_sheet(sheet_df, bank_name=uploaded_file.name.split('.')[0], debug=False)
                    if normalized.empty:
                        continue

                    file_frames.append(normalized)
                    frames.append(normalized)

                file_df = pd.concat(file_frames, ignore_index=True) if file_frames else pd.DataFrame(columns=OUTPUT_COLUMNS)
                if file_df.empty:
                    continue

                ingresos, egresos, diferencia = compute_totals(file_df)
                currency = detect_currency_from_text(" ".join(file_df.get("Origen", pd.Series(dtype=str)).astype(str).tolist()))
                file_summaries[uploaded_file.name] = {
                    "Ingresos": ingresos,
                    "Egresos": egresos,
                    "Diferencia": diferencia,
                    "Moneda": currency,
                    "Cuenta": resolve_file_account_label(file_df),
                }
            except Exception as exc:
                st.error(f"Error al procesar {uploaded_file.name}: {exc}")

        df_para_sheets = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=OUTPUT_COLUMNS)
        if df_para_sheets.empty:
            st.warning("No se encontraron movimientos válidos en los archivos.")
            return

    summary_source = df_para_sheets.copy()
    summary_source["Moneda"] = summary_source["Origen"].astype(str).apply(detect_currency_from_text)

    currency_totals: dict[str, dict[str, float]] = {}
    for currency, group in summary_source.groupby("Moneda", dropna=False):
        ingresos, egresos, neto = compute_totals(group)
        currency_totals[currency] = {
            "Ingresos": ingresos,
            "Egresos": egresos,
            "Neto": neto,
        }

    st.subheader("Resumen general")
    for currency in ["ARS", "USD"]:
        totals = currency_totals.get(currency)
        if totals:
            render_currency_summary_block(currency, totals["Ingresos"], totals["Egresos"], totals["Neto"])

    if len(uploaded_files) >= 2:
        with st.expander("Ver desglose por archivo"):
            if not file_summaries:
                st.info("No hubo archivos válidos para desglosar.")
            else:
                for file_name, summary in file_summaries.items():
                    render_account_summary_card(
                        summary.get("Cuenta", file_name),
                        summary["Ingresos"],
                        summary["Egresos"],
                        summary["Diferencia"],
                        summary["Moneda"],
                    )

    st.subheader("Vista previa de los datos")
    df_vista_previa = df_para_sheets.drop(columns=["ID", "Categoria", "Categoría"], errors="ignore").copy()
    if "Origen" in df_vista_previa.columns:
        df_vista_previa.rename(columns={"Origen": "Cuenta"}, inplace=True)
    if "Fecha" in df_vista_previa.columns:
        df_vista_previa["Fecha"] = format_date_column(df_vista_previa["Fecha"])
    if "Monto" in df_vista_previa.columns:
        origins = df_vista_previa["Cuenta"] if "Cuenta" in df_vista_previa.columns else pd.Series([""] * len(df_vista_previa), index=df_vista_previa.index)
        df_vista_previa["Monto"] = format_money_series_with_origin(df_vista_previa["Monto"], origins)

    st.markdown(
        f"<div style='margin: 0.25rem 0 0.75rem 0; padding: 0.45rem 0.75rem; display:inline-block; border-radius: 999px; background: rgba(255,255,255,0.08); font-weight:700;'>Detectados {len(df_vista_previa)} movimientos</div>",
        unsafe_allow_html=True,
    )

    editor_df = df_vista_previa.copy()
    if "Descripción" not in editor_df.columns:
        editor_df["Descripción"] = ""

    editor_columns = {column: st.column_config.Column(column) for column in editor_df.columns if column != "Descripción"}
    editor_columns["Descripción"] = st.column_config.TextColumn("Descripción")

    edited_preview = st.data_editor(
        editor_df,
        use_container_width=True,
        hide_index=True,
        column_config=editor_columns,
        disabled=[col for col in editor_df.columns if col != "Descripción"],
        key="upload_preview_editor",
    )

    df_vista_previa = edited_preview.copy()
    if "Descripción" in df_vista_previa.columns and "Descripción" in df_para_sheets.columns:
        df_para_sheets.loc[df_vista_previa.index, "Descripción"] = df_vista_previa["Descripción"].astype(str).values

    if st.button("🚀 Confirmar y Subir", type="primary"):
        with st.spinner("Subiendo datos..."):
            try:
                upload_result = upload_dataframe(df_para_sheets)
                load_db_data.clear()
                st.session_state["uploader_key_version"] += 1
            except Exception as exc:
                st.error(f"Error al subir: {exc}")
                return

        inserted = upload_result.get("inserted", 0)

        if inserted > 0:
            st.success("¡Base general actualizada correctamente!")
            st.session_state["subida_exitosa"] = True
            st.rerun()
        else:
            st.info("Todas las transacciones de este archivo ya fueron subidas anteriormente.")
            return

    csv_bytes = df_vista_previa.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar CSV", csv_bytes, file_name="movimientos_normalizados.csv", mime="text/csv")


def render_trash_tab() -> None:
    st.markdown("### 🗑️ Papelera de reciclaje")
    st.caption("Acá se muestran los movimientos eliminados. Podés recuperarlos marcándolos y haciendo clic en Restaurar.")

    try:
        df = load_deleted_data()
    except Exception as exc:
        st.error(f"No se pudieron cargar los movimientos eliminados: {exc}")
        return

    if df.empty:
        st.info("No hay movimientos eliminados.")
        return

    dashboard_df = normalize_dashboard_frame(df)
    dashboard_df["Restaurar"] = False

    visible_cols = ["Restaurar", "Fecha", "Descripción", "Monto", "Tipo", "Origen", "ID"]
    table_df = dashboard_df.reindex(columns=[c for c in visible_cols if c in dashboard_df.columns or c == "Restaurar"])
    if "Fecha_display" in dashboard_df.columns:
        table_df["Fecha"] = dashboard_df["Fecha_display"]
    table_df = table_df[[c for c in visible_cols if c in table_df.columns]]

    display_cols = [c for c in table_df.columns if c != "ID"]
    column_config = {
        "Restaurar": st.column_config.CheckboxColumn("Restaurar", help="Marcá para recuperar este movimiento"),
        "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", disabled=True),
        "Monto": st.column_config.NumberColumn("Monto", format="$ %,.2f", disabled=True),
        "Tipo": st.column_config.TextColumn("Tipo", disabled=True),
        "Descripción": st.column_config.TextColumn("Descripción", disabled=True),
        "Origen": st.column_config.TextColumn("Cuenta", disabled=True),
    }

    edited = st.data_editor(
        table_df[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
        key="trash_editor",
    )

    selected_ids = table_df.loc[edited["Restaurar"] == True, "ID"].tolist() if "ID" in table_df.columns else []

    if st.button("♻️ Restaurar seleccionados", type="primary", disabled=len(selected_ids) == 0):
        count = restore_transactions(selected_ids)
        load_db_data.clear()
        load_deleted_data.clear()
        st.success(f"{count} movimiento(s) restaurado(s) con éxito.")
        st.rerun()


def render_dashboard_tab() -> None:
    _dashboard_css()

    header_left, header_right = st.columns([5, 1])
    with header_left:
        st.title("Home Banking")
        st.subheader("Dashboard Financiero")
    with header_right:
        st.markdown("<div style='height:0.45rem;'></div>", unsafe_allow_html=True)
        if st.button("Upload", use_container_width=True):
            st.session_state["view"] = "upload"
            st.rerun()

    try:
        df = load_db_data()
    except Exception as exc:
        st.error(f"No se pudieron cargar los datos de Neon: {exc}")
        return

    if df.empty:
        st.info("No hay datos para mostrar.")
        return

    dashboard_df = normalize_dashboard_frame(df)
    tabs = st.tabs(["Pesos (ARS)", "Dólares (USD)", "🗑️ Papelera"])

    with tabs[0]:
        render_dashboard_currency_panel(dashboard_df[dashboard_df["Moneda"] == "ARS"].copy(), "ARS")

    with tabs[1]:
        render_dashboard_currency_panel(dashboard_df[dashboard_df["Moneda"] == "USD"].copy(), "USD")

    with tabs[2]:
        render_trash_tab()
if "view" not in st.session_state:
    st.session_state["view"] = "dashboard"

run_startup_tasks()

st.markdown(
    "<div style='height:0.15rem;'></div>",
    unsafe_allow_html=True,
)

if st.session_state["view"] == "dashboard":
    render_dashboard_tab()
    if st.session_state.get("upload_success_message"):
        st.session_state["view"] = "dashboard"
else:
    if st.button("Volver al resumen"):
        st.session_state["view"] = "dashboard"
        st.rerun()
    render_upload_tab()
