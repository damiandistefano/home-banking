from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st

from transaction_uploader import (
    GOOGLE_SCOPES,
    OUTPUT_COLUMNS,
    normalize_sheet,
    read_excel_sheets,
    upload_dataframe,
    load_service_account_info,
)

# SPREADSHEET_ID = "15EEuMOCws2hPp6sw6Nfpd8lBu9AazCtctgmLENHsyh4" # Real
SPREADSHEET_ID = "1_AgKWi22JQxYCdrGmQn5aA_YT2zHs1fvVbZvCS9ARx8" # Prueba

st.set_page_config(page_title="Importador de transacciones", layout="wide")


@st.cache_data(ttl=300)
def load_sheet_data() -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials

    service_account_info = load_service_account_info()
    if not service_account_info:
        raise FileNotFoundError("Falta la credencial. Guardá el JSON en ./.datos_banco.txt")

    credentials = Credentials.from_service_account_info(service_account_info, scopes=GOOGLE_SCOPES)
    client = gspread.authorize(credentials)
    
    worksheet = client.open_by_key(SPREADSHEET_ID).worksheet("Movimientos")
    rows = worksheet.get_all_records(value_render_option="UNFORMATTED_VALUE")
    return pd.DataFrame(rows)


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


def style_tipo(val):
    if str(val).lower() == "ingreso":
        return "background-color: #2e7d32; color: white; font-weight: bold; border-radius: 4px;"
    if str(val).lower() == "egreso":
        return "background-color: #c62828; color: white; font-weight: bold; border-radius: 4px;"
    return ""


def render_upload_tab() -> None:
    st.subheader("Cargar movimientos")

    if "uploader_key_version" not in st.session_state:
        st.session_state["uploader_key_version"] = 0

    left_col, right_col = st.columns([4, 1])
    with left_col:
        st.write("Subí un archivo Excel para actualizar la planilla general.")

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

    with st.expander("Ver desglose por archivo"):
        if not file_summaries:
            st.info("No hubo archivos válidos para desglosar.")
        else:
            for file_name, summary in file_summaries.items():
                render_individual_file_card(
                    file_name,
                    summary["Ingresos"],
                    summary["Egresos"],
                    summary["Diferencia"],
                    summary["Moneda"],
                )

    st.subheader("Vista previa de los datos")
    df_vista_previa = df_para_sheets.drop(columns=["ID", "Referencia"], errors="ignore").copy()
    if "Fecha" in df_vista_previa.columns:
        df_vista_previa["Fecha"] = format_date_column(df_vista_previa["Fecha"])
    if "Monto" in df_vista_previa.columns:
        origins = df_vista_previa["Origen"] if "Origen" in df_vista_previa.columns else pd.Series([""] * len(df_vista_previa), index=df_vista_previa.index)
        df_vista_previa["Monto"] = format_money_series_with_origin(df_vista_previa["Monto"], origins)

    st.markdown(
        f"<div style='margin: 0.25rem 0 0.75rem 0; padding: 0.45rem 0.75rem; display:inline-block; border-radius: 999px; background: rgba(255,255,255,0.08); font-weight:700;'>Detectados {len(df_vista_previa)} movimientos</div>",
        unsafe_allow_html=True,
    )

    st.dataframe(df_vista_previa.style.map(style_tipo, subset=["Tipo"]), use_container_width=True, hide_index=True)

    if st.button("🚀 Confirmar y Subir a Sheets", type="primary"):
        with st.spinner("Subiendo datos a Google Sheets..."):
            try:
                upload_result = upload_dataframe(df_para_sheets)
                load_sheet_data.clear()
                st.session_state["uploader_key_version"] += 1
            except Exception as exc:
                st.error(f"Error al subir: {exc}")
                return

        inserted = upload_result.get("inserted", 0)

        if inserted > 0:
            st.success("¡Planilla general actualizada correctamente!")
            st.balloons()
            st.session_state["subida_exitosa"] = True
            st.rerun()
        else:
            st.info("Todas las transacciones de este archivo ya fueron subidas anteriormente.")
            return

    csv_bytes = df_vista_previa.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar CSV", csv_bytes, file_name="movimientos_normalizados.csv", mime="text/csv")


def render_dashboard_tab() -> None:
    st.title("Tablero de Resumen")

    try:
        df = load_sheet_data()
    except Exception as exc:
        st.error(f"No se pudieron cargar los datos de Google Sheets: {exc}")
        return

    if df.empty:
        st.info("No hay datos para mostrar.")
        return

    df = df.copy()

    df["Fecha"] = format_date_column(df.get("Fecha"))

    if "Categoría" in df.columns and "Categoria" not in df.columns:
        df.rename(columns={"Categoría": "Categoria"}, inplace=True)

    for col in ["Origen", "Tipo", "Categoria", "Descripción"]:
        if col not in df.columns:
            df[col] = ""

    df["Moneda"] = df["Origen"].astype(str).str.upper().apply(
        lambda x: "USD" if any(token in x for token in ["USD", "U$S", "DOLARES", "DÓLARES"]) else "ARS"
    )

    moneda_filter = st.selectbox("Moneda", options=["Todas", "ARS", "USD"], index=0)

    origenes = sorted([x for x in df["Origen"].dropna().astype(str).unique().tolist() if x])
    origen_filter = st.multiselect("Filtrar por Origen", options=origenes, default=origenes)

    filtered = df.copy()
    if origen_filter:
        filtered = filtered[filtered["Origen"].isin(origen_filter)]
    if moneda_filter != "Todas":
        filtered = filtered[filtered["Moneda"] == moneda_filter]

    def render_metrics(frame: pd.DataFrame, currency: str) -> None:
        if frame.empty:
            st.info(f"No hay movimientos en {currency}.")
            return

        ingresos = frame.loc[frame["Tipo"].astype(str).str.lower() == "ingreso", "Monto"].sum()
        egresos = frame.loc[frame["Tipo"].astype(str).str.lower() == "egreso", "Monto"].sum()
        saldo = ingresos - egresos

        k1, k2, k3 = st.columns(3)
        k1.metric("Total de Ingresos", format_money_premium(ingresos, currency))
        k2.metric("Total de Egresos", format_money_premium(egresos, currency))
        k3.metric("Saldo Neto", format_money_premium(saldo, currency))

    if moneda_filter == "Todas":
        st.subheader("Pesos (ARS)")
        ars_frame = filtered[filtered["Moneda"] == "ARS"]
        render_metrics(ars_frame, "ARS")
        st.write("---")
        st.subheader("Dólares (USD)")
        usd_frame = filtered[filtered["Moneda"] == "USD"]
        render_metrics(usd_frame, "USD")
    else:
        render_metrics(filtered, moneda_filter)

    st.subheader("Resumen por origen")
    filtered = filtered.drop(columns=["Referencia"], errors="ignore")

    summary_source = filtered.copy()
    if "Origen" not in summary_source.columns:
        summary_source["Origen"] = ""
    summary_source["Origen"] = summary_source["Origen"].astype(str)
    grouped = summary_source.groupby("Origen", dropna=False)
    summary_rows = []
    for origen, group in grouped:
        ingresos, egresos, diferencia = compute_totals(group)
        summary_rows.append({
            "Origen": origen or "Sin origen",
            "Ingresos": format_money_premium(ingresos, origen),
            "Egresos": format_money_premium(egresos, origen),
            "Diferencia": format_money_premium(diferencia, origen),
        })
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    st.write("### Detalle de Transacciones")
    display_df = filtered.drop(columns=["ID", "Referencia"], errors="ignore").copy()
    display_df = display_df.dropna(axis=1, how="all")
    if "Monto" in display_df.columns:
        display_df["Monto"] = format_amount_series(display_df["Monto"])
    if "Moneda" not in display_df.columns:
        origen_source = display_df["Origen"] if "Origen" in display_df.columns else pd.Series([""] * len(display_df), index=display_df.index)
        display_df["Moneda"] = origen_source.astype(str).str.upper().apply(
            lambda x: "USD" if any(token in x for token in ["USD", "U$S", "DOLARES", "DÓLARES"]) else "ARS"
        )

    st.dataframe(
        display_df,
        use_container_width=True,
        column_config={"Moneda": st.column_config.TextColumn("Moneda")},
        hide_index=True
    )
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
