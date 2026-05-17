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
SPREADSHEET_ID = "15EEuMOCws2hPp6sw6Nfpd8lBu9AazCtctgmLENHsyh4"

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


def render_upload_tab() -> None:
    st.title("Importador de transacciones")
    st.caption("Usa `./.datos_banco.txt` y comparte la hoja con el `client_email` de esa cuenta.")

    success_message = st.session_state.pop("upload_success_message", None)
    if success_message:
        st.success(success_message)

    uploaded_files = st.file_uploader("Subí tus archivos .xls o .xlsx", type=["xls", "xlsx"], accept_multiple_files=True)
    debug = st.checkbox("Modo debug", value=True)

    if not uploaded_files:
        return

    frames = []
    for uploaded_file in uploaded_files:
        with TemporaryDirectory() as tmpdir:
            temp_path = Path(tmpdir) / uploaded_file.name
            temp_path.write_bytes(uploaded_file.getvalue())

            try:
                sheets = read_excel_sheets(temp_path, debug=debug)
            except Exception as exc:
                st.error(f"No se pudo leer {uploaded_file.name}: {exc}")
                continue

            for sheet_name, sheet_df in sheets.items():
                normalized = normalize_sheet(sheet_df, bank_name=uploaded_file.name.split('.')[0], debug=debug)
                if not normalized.empty:
                    frames.append(normalized)

    consolidated = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=OUTPUT_COLUMNS)
    display_df = consolidated.drop(columns=["ID"], errors="ignore")
    st.dataframe(display_df, use_container_width=True)

    csv_bytes = display_df.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar CSV", csv_bytes, file_name="movimientos_normalizados.csv", mime="text/csv")

    if st.button("Subir a Google Sheets", type="primary", disabled=consolidated.empty):
        try:
            total = upload_dataframe(consolidated)
            st.session_state["upload_success_message"] = f"Se subieron {total} transacciones con éxito."
            load_sheet_data.clear()
            st.rerun()
        except Exception as exc:
            st.error(f"Error al subir: {exc}")


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
    
    df["Fecha"] = pd.to_datetime(df.get("Fecha"), errors="coerce")

    # 2. LIMPIEZA DE COLUMNAS
    # Unificamos "Categoria" y nos aseguramos de que existan las mínimas requeridas
    if "Categoría" in df.columns and "Categoria" not in df.columns:
        df.rename(columns={"Categoría": "Categoria"}, inplace=True)
        
    for col in ["Origen", "Tipo", "Categoria", "Descripción"]:
        if col not in df.columns:
            df[col] = ""

    df["Moneda"] = df["Origen"].astype(str).str.upper().apply(
        lambda x: "USD" if any(token in x for token in ["USD", "U$S", "DOLARES", "DÓLARES"]) else "ARS"
    )

    # 4. FILTROS EN LA INTERFAZ
    moneda_filter = st.selectbox("Moneda", options=["Todas", "ARS", "USD"], index=0)

    # Como en tu base no hay "Banco" pero sí "Origen" (Ej: Galicia - USD), filtramos por ahí
    origenes = sorted([x for x in df["Origen"].dropna().astype(str).unique().tolist() if x])
    origen_filter = st.multiselect("Filtrar por Origen", options=origenes, default=origenes)

    filtered = df.copy()
    if origen_filter:
        filtered = filtered[filtered["Origen"].isin(origen_filter)]
    if moneda_filter != "Todas":
        filtered = filtered[filtered["Moneda"] == moneda_filter]

    def format_money(amount: float, currency: str) -> str:
        symbol = "$" if currency == "ARS" else "U$S"
        parts = f"{amount:.2f}".split(".")
        entero_con_puntos = f"{int(parts[0]):,}".replace(",", ".")
        centavos = parts[1]
        return f"{symbol} {entero_con_puntos},{centavos}"

    def render_metrics(frame: pd.DataFrame, currency: str) -> None:
        if frame.empty:
            st.info(f"No hay movimientos en {currency}.")
            return

        ingresos = frame.loc[frame["Tipo"].astype(str).str.lower() == "ingreso", "Monto"].sum()
        egresos = frame.loc[frame["Tipo"].astype(str).str.lower() == "egreso", "Monto"].sum()
        saldo = ingresos - egresos

        k1, k2, k3 = st.columns(3)
        k1.metric("Total de Ingresos", format_money(ingresos, currency))
        k2.metric("Total de Egresos", format_money(egresos, currency))
        k3.metric("Saldo Neto", format_money(saldo, currency))

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

    st.write("### Detalle de Transacciones")
    
    # Ocultamos ID y eliminamos columnas que estén 100% vacías para que quede limpio
    display_df = filtered.drop(columns=["ID"], errors="ignore").copy()
    display_df = display_df.dropna(axis=1, how='all')
    if "Moneda" not in display_df.columns:
        origen_source = display_df["Origen"] if "Origen" in display_df.columns else pd.Series([""] * len(display_df), index=display_df.index)
        display_df["Moneda"] = origen_source.astype(str).str.upper().apply(
            lambda x: "USD" if any(token in x for token in ["USD", "U$S", "DOLARES", "DÓLARES"]) else "ARS"
        )
    display_df["Moneda"] = display_df["Moneda"].replace({"ARS": "$", "USD": "U$S"})
    
    st.dataframe(
        display_df,
        use_container_width=True,
        column_config={"Moneda": st.column_config.TextColumn("Moneda")},
        hide_index=True
    )

menu = st.sidebar.radio(
    "Navegación",
    ["📥 Cargar Movimientos", "📊 Tablero de Resumen"],
)

if menu == "📥 Cargar Movimientos":
    render_upload_tab()
else:
    render_dashboard_tab()
