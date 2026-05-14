from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st

from transaction_uploader import OUTPUT_COLUMNS, normalize_sheet, read_excel_sheets, upload_dataframe


st.set_page_config(page_title="Importador de transacciones", layout="wide")
st.title("Importador de transacciones")
st.caption("Usa `service_account.json` y comparte la hoja con el `client_email` de esa cuenta.")

uploaded_files = st.file_uploader("Subí tus archivos .xls o .xlsx", type=["xls", "xlsx"], accept_multiple_files=True)
debug = st.checkbox("Modo debug", value=True)

if uploaded_files:
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
    st.dataframe(consolidated, use_container_width=True)

    csv_bytes = consolidated.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar CSV", csv_bytes, file_name="movimientos_normalizados.csv", mime="text/csv")

    if st.button("Subir a Google Sheets", type="primary", disabled=consolidated.empty):
        try:
            total = upload_dataframe(consolidated)
            st.success(f"Se subieron {total} transacciones con éxito.")
        except Exception as exc:
            st.error(f"Error al subir: {exc}")
