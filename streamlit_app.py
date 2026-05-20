from pathlib import Path
from tempfile import TemporaryDirectory

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
)

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
        order by fecha desc, id desc
    """

    with psycopg2.connect(get_database_url()) as conn:
        return pd.read_sql_query(query, conn)


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
    
    # 3. Convertimos SOLO los textos usando el parseo normal de Pandas
    if (~es_numero).any():
        fechas_finales[~es_numero] = pd.to_datetime(
            series[~es_numero], 
            errors="coerce", 
            dayfirst=True
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


def render_kpi_card(title: str, value: str, meta: str, trend: str = "neutral", emphasize: bool = False) -> None:
    badge_text = {"up": "↗ tendencia positiva", "down": "↘ salida de fondos", "neutral": "• seguimiento"}.get(trend, "• seguimiento")
    trend_class = trend if trend in {"up", "down"} else "neutral"
    value_class = "dash-kpi-value neto" if emphasize else "dash-kpi-value"
    st.markdown(
        f"""
        <div class="dash-kpi">
          <div class="dash-kpi-label">{title}</div>
          <div class="{value_class}">{value}</div>
          <div class="dash-kpi-meta">{meta}</div>
          <div class="dash-badge {trend_class}">{badge_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard_kpis(frame: pd.DataFrame, currency: str) -> None:
    ingresos, egresos, neto = calculate_dashboard_totals(frame)
    symbol = _currency_meta(currency)[1]
    net_trend = "up" if neto >= 0 else "down"

    c1, c2, c3 = st.columns(3)
    with c1:
        render_kpi_card("Saldo Neto", _format_currency_value(neto, symbol), "Resultado acumulado del período", net_trend, emphasize=True)
    with c2:
        render_kpi_card("Ingresos", _format_currency_value(ingresos, symbol), "Entradas de dinero filtradas", "up")
    with c3:
        render_kpi_card("Egresos", _format_currency_value(egresos, symbol), "Salidas de dinero filtradas", "down")


def render_dashboard_filters(frame: pd.DataFrame, currency: str) -> tuple[pd.DataFrame, str]:
    min_date = frame["Fecha_dt"].min()
    max_date = frame["Fecha_dt"].max()
    if pd.isna(min_date) or pd.isna(max_date):
        return frame.iloc[0:0], ""

    min_day = min_date.date()
    max_day = max_date.date()
    currency_key = currency.lower()
    date_key = f"dashboard_date_{currency_key}"
    origin_key = f"dashboard_origin_{currency_key}"
    search_key = f"dashboard_search_{currency_key}"

    if date_key not in st.session_state:
        st.session_state[date_key] = (min_day, max_day)
    if origin_key not in st.session_state:
        st.session_state[origin_key] = sorted([x for x in frame["Origen"].dropna().astype(str).unique().tolist() if x])
    if search_key not in st.session_state:
        st.session_state[search_key] = ""

    st.markdown("<div class='dash-shell'>", unsafe_allow_html=True)
    filter_cols = st.columns([1.0, 1.5, 1.25, 1.25])

    with filter_cols[0]:
        st.markdown("<div class='dash-filter-label'>Moneda</div>", unsafe_allow_html=True)
        st.caption(_currency_meta(currency)[0])

    with filter_cols[1]:
        st.markdown("<div class='dash-filter-label'>Rango de fechas</div>", unsafe_allow_html=True)
        st.date_input(
            "Rango de fechas",
            value=st.session_state[date_key],
            min_value=min_day,
            max_value=max_day,
            key=date_key,
            label_visibility="collapsed",
        )

    origin_options = sorted([x for x in frame["Origen"].dropna().astype(str).unique().tolist() if x])
    with filter_cols[2]:
        st.markdown("<div class='dash-filter-label'>Cuenta</div>", unsafe_allow_html=True)
        st.multiselect(
            "Cuenta",
            options=origin_options,
            default=st.session_state[origin_key],
            key=origin_key,
            label_visibility="collapsed",
        )

    with filter_cols[3]:
        st.markdown("<div class='dash-filter-label'>Buscar</div>", unsafe_allow_html=True)
        st.text_input(
            "Buscar transacción...",
            placeholder="Buscar transacción...",
            key=search_key,
            label_visibility="collapsed",
        )

    selected_range = st.session_state[date_key]
    selected_origins = st.session_state[origin_key]
    search_text = str(st.session_state[search_key]).strip().lower()

    filtered = frame.copy()
    if isinstance(selected_range, (list, tuple)):
        if len(selected_range) == 2:
            start_date, end_date = selected_range
            filtered = filtered[(filtered["Fecha_dt"].dt.date >= start_date) & (filtered["Fecha_dt"].dt.date <= end_date)]
        elif len(selected_range) == 1:
            start_date = selected_range[0]
            if hasattr(start_date, "year"):
                filtered = filtered[filtered["Fecha_dt"].dt.date >= start_date]
    else:
        if hasattr(selected_range, "year"):
            filtered = filtered[filtered["Fecha_dt"].dt.date >= selected_range]
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
        fig.update_layout(template="plotly_dark", height=360, margin=dict(l=10, r=10, t=40, b=10), title=f"Distribución por cuenta - {currency_label}")
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
        fig.update_layout(template="plotly_dark", height=360, margin=dict(l=10, r=10, t=40, b=10), title=f"Distribución por cuenta - {currency_label}")
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
        height=360,
        margin=dict(l=10, r=10, t=48, b=10),
        title=f"Distribución por cuenta - {currency_label}",
        showlegend=True,
        legend_title_text="Cuenta",
    )
    return fig


def render_dashboard_detail_table(frame: pd.DataFrame, currency: str) -> None:
    table_df = frame.drop(columns=["Fecha_dt"], errors="ignore").copy()
    if "Fecha_display" in table_df.columns:
        table_df["Fecha"] = table_df["Fecha_display"]
        table_df.drop(columns=["Fecha_display"], inplace=True, errors="ignore")
    table_df = table_df.drop(columns=["ID", "Categoria", "Origen"], errors="ignore")
    if "Cuenta" not in table_df.columns and "Origen" in frame.columns:
        table_df["Cuenta"] = frame["Origen"]

    visible_columns = [col for col in ["Fecha", "Descripción", "Monto", "Tipo", "Cuenta"] if col in table_df.columns]
    other_columns = [col for col in table_df.columns if col not in visible_columns]
    table_df = table_df[visible_columns + other_columns]

    symbol = _currency_meta(currency)[1]
    money_format = f"{symbol} %,.2f"
    column_config: dict[str, object] = {}

    if "Fecha" in table_df.columns:
        column_config["Fecha"] = st.column_config.DateColumn("Fecha", format="DD/MM/YYYY")
    if "Monto" in table_df.columns:
        column_config["Monto"] = st.column_config.NumberColumn("Monto", format=money_format)
    if "Tipo" in table_df.columns:
        column_config["Tipo"] = st.column_config.TextColumn("Tipo")
    if "Cuenta" in table_df.columns:
        column_config["Cuenta"] = st.column_config.TextColumn("Cuenta")
    if "Descripción" in table_df.columns:
        column_config["Descripción"] = st.column_config.TextColumn("Descripción")

    st.dataframe(
        table_df.style.map(style_tipo, subset=["Tipo"]) if "Tipo" in table_df.columns else table_df,
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
    )


def render_dashboard_currency_panel(frame: pd.DataFrame, currency: str) -> None:
    label, _ = _currency_meta(currency)
    st.markdown(f"### {label}")
    filtered, _ = render_dashboard_filters(frame, currency)

    if filtered.empty:
        st.info("No hay movimientos para los filtros seleccionados.")
        return

    st.markdown(
        f"<div style='margin: 0.5rem 0 1rem 0; padding: 0.35rem 0.75rem; display:inline-flex; border-radius: 999px; background: rgba(255,255,255,0.08); font-weight:700;'>"
        f"{len(filtered)} movimientos visibles</div>",
        unsafe_allow_html=True,
    )

    render_dashboard_kpis(filtered, currency)

    chart_left, chart_right = st.columns(2)
    with chart_left:
        st.plotly_chart(build_cashflow_figure(filtered, currency), use_container_width=True)
    with chart_right:
        st.plotly_chart(build_distribution_figure(filtered, currency), use_container_width=True)

    st.markdown("### Detalle de transacciones")
    render_dashboard_detail_table(filtered, currency)


def render_upload_tab() -> None:
    st.subheader("Cargar movimientos")

    if "uploader_key_version" not in st.session_state:
        st.session_state["uploader_key_version"] = 0

    left_col, right_col = st.columns([4, 1])
    with left_col:
        st.write("Subí un archivo Excel para actualizar la base general.")

    with right_col:
        uploaded_file = st.file_uploader(
            "Upload",
            type=["xls", "xlsx"],
            accept_multiple_files=True,
            key=f"uploader_{st.session_state['uploader_key_version']}",
        )

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


def render_dashboard_tab() -> None:
    st.title("Tablero de Resumen")
    st.caption("Panel analítico financiero para seguimiento ejecutivo de caja, ingresos y egresos.")

    _dashboard_css()

    try:
        df = load_db_data()
    except Exception as exc:
        st.error(f"No se pudieron cargar los datos de Neon: {exc}")
        return

    if df.empty:
        st.info("No hay datos para mostrar.")
        return

    dashboard_df = normalize_dashboard_frame(df)
    tabs = st.tabs(["Pesos (ARS)", "Dólares (USD)"])

    with tabs[0]:
        render_dashboard_currency_panel(dashboard_df[dashboard_df["Moneda"] == "ARS"].copy(), "ARS")

    with tabs[1]:
        render_dashboard_currency_panel(dashboard_df[dashboard_df["Moneda"] == "USD"].copy(), "USD")
if "view" not in st.session_state:
    st.session_state["view"] = "dashboard"

header_left, header_right = st.columns([5, 1])
with header_left:
    st.title("Home-Banking")
with header_right:
    if st.button("Upload", use_container_width=True):
        st.session_state["view"] = "upload"

if st.session_state["view"] == "dashboard":
    render_dashboard_tab()
    if st.session_state.get("upload_success_message"):
        st.session_state["view"] = "dashboard"
else:
    if st.button("Volver al resumen"):
        st.session_state["view"] = "dashboard"
        st.rerun()
    render_upload_tab()
