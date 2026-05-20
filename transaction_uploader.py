from __future__ import annotations

import argparse
import hashlib
import re
import os
from pathlib import Path

import pandas as pd
import psycopg2
from psycopg2.extras import execute_values

try:
    import streamlit as st
except ImportError:
    st = None

try:
    import openpyxl  # noqa: F401

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import xlrd  # noqa: F401

    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


INPUT_DIR = Path("./input")
OUTPUT_COLUMNS = ["Fecha", "Descripción", "Monto", "Tipo", "Origen", "Categoria", "ID"]
TABLE_NAME = "movimientos"


def get_database_url() -> str:
    database_url = os.environ.get("DATABASE_URL", "").strip()

    if not database_url and st is not None:
        try:
            database_url = str(st.secrets.get("DATABASE_URL", "")).strip()
        except Exception:
            database_url = ""

    if not database_url:
        raise EnvironmentError("Falta DATABASE_URL. Configurala en Streamlit Secrets o como variable de entorno.")
    return database_url


def get_db_connection():
    return psycopg2.connect(get_database_url())


def ensure_table_exists() -> None:
    with get_db_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                f"""
                create table if not exists {TABLE_NAME} (
                  id text primary key,
                  fecha date not null,
                  descripcion text not null,
                  monto numeric(14,2) not null,
                  tipo text not null,
                  origen text not null,
                  categoria text default ''
                )
                """
            )
        conn.commit()


def file_signature(file_path: Path) -> str:
    head = file_path.read_bytes()[:4]
    if head.startswith(b"PK"):
        return "xlsx"
    if head == b"\xd0\xcf\x11\xe0":
        return "xls"
    return "unknown"


def read_excel_sheets(file_path: Path, debug: bool = False) -> dict[str, pd.DataFrame]:
    signature = file_signature(file_path)
    engines: list[str] = []

    if signature == "xlsx":
        if HAS_OPENPYXL:
            engines.append("openpyxl")
        if HAS_XLRD:
            engines.append("xlrd")
    elif signature == "xls":
        if HAS_XLRD:
            engines.append("xlrd")
        if HAS_OPENPYXL:
            engines.append("openpyxl")
    else:
        if HAS_OPENPYXL:
            engines.append("openpyxl")
        if HAS_XLRD:
            engines.append("xlrd")

    if not engines:
        raise ImportError("Faltan dependencias para leer Excel: instalá openpyxl y/o xlrd")

    last_error: Exception | None = None
    for engine in engines:
        try:
            if debug:
                print(f"Leyendo {file_path.name} con engine={engine} (firma={signature})")
            return pd.read_excel(file_path, sheet_name=None, header=None, engine=engine)
        except Exception as exc:
            last_error = exc
            if debug:
                print(f"Fallo engine={engine} en {file_path.name}: {exc}")

    raise last_error or ValueError(f"No se pudo leer {file_path.name}")


def clean_amount(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().upper()
    for char in ["$", "U$S", "USD", " "]:
        text = text.replace(char, "")

    if not text:
        return 0.0

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    if text.startswith("-"):
        negative = True
        text = text[1:]

    dots = text.count('.')
    commas = text.count(',')

    if dots > 0 and commas > 0:
        if text.rfind(',') > text.rfind('.'):
            text = text.replace('.', '').replace(',', '.')
        else:
            text = text.replace(',', '')
    elif commas > 0:
        digits_after = len(text) - text.rfind(',') - 1
        if commas > 1 or digits_after == 3:
            text = text.replace(',', '')
        else:
            text = text.replace(',', '.')
    elif dots > 0:
        digits_after = len(text) - text.rfind('.') - 1
        if dots > 1 or digits_after == 3:
            text = text.replace('.', '')

    text = re.sub(r"[^0-9\.-]", "", text)
    if text in {"", ".", "-"}:
        return 0.0

    amount = float(text)
    return -amount if negative else amount


def header_score(values: list[str]) -> int:
    normalized = [str(v).strip().lower() for v in values]
    score = 0
    if any(v in normalized for v in ["fecha", "date", "f. mov.", "f mov", "operacion", "operación"]):
        score += 1
    if any(v in normalized for v in ["descripción", "descripcion", "concepto", "detalle", "glosa", "movimiento"]):
        score += 1
    if any(v in normalized for v in ["monto", "importe", "amount", "valor", "saldo", "débito", "debito", "crédito", "credito"]):
        score += 1
    return score


def detect_header_row(df_raw: pd.DataFrame, debug: bool = False) -> int:
    best_row = -1
    best_score = 0

    for index, row in df_raw.iterrows():
        values = ["" if pd.isna(cell) else str(cell).strip() for cell in row.tolist()]
        score = header_score(values)
        if debug:
            print(f"Fila {index} score={score} valores={values[:8]}")
        if score > best_score:
            best_score = score
            best_row = index

    return best_row if best_score >= 2 else -1


def detect_macro_origin(df_raw: pd.DataFrame, bank_name: str = "") -> str:
    for _, row in df_raw.head(10).iterrows():
        text = " ".join("" if pd.isna(cell) else str(cell) for cell in row.tolist()).lower()
        if "caja de ahorros en pesos" in text or "caja de ahorro en pesos" in text:
            return "Macro - Caja Ahorro Pesos"
        if "cuenta corriente" in text:
            return "Macro - Cuenta Corriente"
    return bank_name or "Macro"


def detect_galicia_origin(df_raw: pd.DataFrame, bank_name: str = "") -> str:
    for _, row in df_raw.head(10).iterrows():
        text = " ".join("" if pd.isna(cell) else str(cell) for cell in row.tolist()).lower()
        if "banco galicia" in text:
            if "caja ahorro pesos" in text or "caja de ahorro pesos" in text:
                return "Galicia - Caja Ahorro Pesos"
            if "cuenta corriente" in text:
                return "Galicia - Cuenta Corriente"
            return "Galicia"
    return bank_name or "Galicia"


def detect_currency_origin(df_raw: pd.DataFrame, origin: str) -> str:
    text = " ".join(
        "" if pd.isna(cell) else str(cell)
        for _, row in df_raw.head(10).iterrows()
        for cell in row.tolist()
    ).lower()
    if any(token in text for token in ["dolares", "dólares", "usd", "u$s", "usd$", "moneda: dolares"]):
        return f"{origin} - USD"
    return origin


def build_transaction_id(row: pd.Series) -> str:
    base = "|".join(
        [
            str(row.get("Fecha", "")),
            str(row.get("Descripción", "")),
            str(row.get("Monto", "")),
            str(row.get("Origen", "")),
        ]
    )
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:16]


def detect_columns(df: pd.DataFrame) -> tuple[str, str, str] | None:
    lowered = {str(col).strip().lower(): col for col in df.columns}

    date_candidates = ["fecha", "date", "f. mov.", "f mov", "operacion", "operación"]
    desc_candidates = ["descripción", "descripcion", "concepto", "detalle", "description", "glosa", "movimiento"]
    amount_candidates = ["monto", "importe", "amount", "valor", "saldo", "debito", "débito", "credito", "crédito"]

    date_col = next((lowered[c] for c in date_candidates if c in lowered), None)
    desc_col = next((lowered[c] for c in desc_candidates if c in lowered), None)
    amount_col = next((lowered[c] for c in amount_candidates if c in lowered), None)

    if date_col and desc_col and amount_col:
        return date_col, desc_col, amount_col
    return None


def detect_optional_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lowered = {str(col).strip().lower(): col for col in df.columns}
    return next((lowered[c] for c in candidates if c in lowered), None)


def clean_reference(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def parse_excel_date(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=str)

    if not isinstance(series, pd.Series):
        series = pd.Series(series)

    numeric = pd.to_numeric(series, errors="coerce")
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)

    if numeric.notna().any():
        numeric_dates = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
        parsed = parsed.where(numeric.isna(), numeric_dates)

    return parsed.dt.strftime("%Y-%m-%d")


def normalize_sheet(df_raw: pd.DataFrame, bank_name: str = "", debug: bool = False) -> pd.DataFrame:
    header_row = detect_header_row(df_raw, debug=debug)
    if header_row == -1:
        if debug:
            print(f"No se detectó encabezado. Columnas crudas: {list(df_raw.columns)}")
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = [str(c).strip() for c in df_raw.iloc[header_row].tolist()]
    df = df.loc[:, ~df.columns.duplicated()]

    cols = detect_columns(df)
    if cols is None:
        if debug:
            print(f"No se detectaron columnas útiles. Encabezados: {list(df.columns)}")
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    date_col, desc_col, amount_col = cols
    santander_savings_col = detect_optional_column(df, ["caja de ahorro", "caja ahorro"])
    santander_current_col = detect_optional_column(df, ["cuenta corriente", "cta cte", "cta. cte."])
    galicia_debit_col = detect_optional_column(df, ["débito", "debito"])
    galicia_credit_col = detect_optional_column(df, ["crédito", "credito"])
    macro_debit_col = detect_optional_column(df, ["débito", "debito"])
    macro_credit_col = detect_optional_column(df, ["crédito", "credito"])

    is_santander_layout = bool(santander_savings_col or santander_current_col)
    is_galicia_layout = bool(galicia_debit_col and galicia_credit_col)
    is_macro_layout = "macro" in str(bank_name).lower()
    is_usd_sheet = detect_currency_origin(df_raw, "").endswith("USD")

    if is_santander_layout:
        savings_series = df[santander_savings_col].apply(clean_amount) if santander_savings_col else pd.Series([0.0] * len(df), index=df.index)
        current_series = df[santander_current_col].apply(clean_amount) if santander_current_col else pd.Series([0.0] * len(df), index=df.index)

        def row_amount(row_index: int) -> float:
            savings_amount = float(savings_series.loc[row_index])
            current_amount = float(current_series.loc[row_index])
            if savings_amount != 0:
                return savings_amount
            if current_amount != 0:
                return current_amount
            return 0.0

        amount_series = pd.Series([row_amount(idx) for idx in df.index], index=df.index)
        origin_series = pd.Series(
            [
                "Santander - Caja de Ahorro Pesos" if (santander_savings_col and float(savings_series.loc[idx]) != 0)
                else "Santander - Cuenta Corriente" if (santander_current_col and float(current_series.loc[idx]) != 0)
                else "Santander"
                for idx in df.index
            ],
            index=df.index,
        )
        if is_usd_sheet:
            origin_series = pd.Series([f"{origin} - USD" for origin in origin_series], index=df.index)
            
    elif is_galicia_layout:
        # ARREGLO AQUÍ: Forzamos .abs() para evitar el doble negativo
        debit_series = df[galicia_debit_col].apply(clean_amount).abs()
        credit_series = df[galicia_credit_col].apply(clean_amount).abs()
        amount_series = credit_series - debit_series
        galicia_origin = detect_currency_origin(df_raw, detect_galicia_origin(df_raw, bank_name=bank_name))
        origin_series = pd.Series([galicia_origin for _ in df.index], index=df.index)
        
    elif is_macro_layout and macro_debit_col and macro_credit_col:
        # ARREGLO AQUÍ: También protegemos a Macro por las dudas
        debit_series = df[macro_debit_col].apply(clean_amount).abs()
        credit_series = df[macro_credit_col].apply(clean_amount).abs()
        amount_series = credit_series - debit_series
        macro_origin = detect_currency_origin(df_raw, detect_macro_origin(df_raw, bank_name=bank_name))
        origin_series = pd.Series([macro_origin for _ in df.index], index=df.index)
        
    else:
        amount_series = df[amount_col].apply(clean_amount)
        if amount_col.strip().lower() == "importe":
            macro_origin = detect_currency_origin(df_raw, detect_macro_origin(df_raw, bank_name=bank_name))
            origin_series = pd.Series([macro_origin for _ in df.index], index=df.index)
        else:
            origin_series = pd.Series([detect_currency_origin(df_raw, bank_name or "") for _ in df.index], index=df.index)

    out = pd.DataFrame()
    out["ID"] = ""
    out["Fecha"] = parse_excel_date(df[date_col])
    out["Descripción"] = df[desc_col].astype(str).str.strip()
    signed_amounts = amount_series.copy()
    out["Tipo"] = signed_amounts.apply(lambda amount: "Ingreso" if amount >= 0 else "Egreso")
    out["Monto"] = signed_amounts.abs()
    out["Origen"] = origin_series
    out["Categoria"] = ""
    out = out.dropna(subset=["Fecha"])
    out = out[out["Descripción"].ne("") & out["Descripción"].ne("nan")]
    out["ID"] = out.apply(build_transaction_id, axis=1)

    if debug:
        print(f"Detectado: fecha={date_col}, descripcion={desc_col}, monto={amount_col}, filas={len(out)}")

    return out[OUTPUT_COLUMNS]


def read_excel_files(folder: Path, debug: bool = False) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    for file_path in sorted(folder.glob("*.xls*")):
        try:
            sheets = read_excel_sheets(file_path, debug=debug)
        except Exception as exc:
            print(f"Error en {file_path.name}: {exc}")
            continue

        if debug:
            print(f"Hojas en {file_path.name}: {list(sheets.keys())}")

        for sheet_name, sheet_df in sheets.items():
            normalized = normalize_sheet(sheet_df, bank_name=file_path.stem, debug=debug)
            if not normalized.empty:
                frames.append(normalized)
            elif debug:
                print(f"Hoja no normalizable: {sheet_name}")

    if not frames:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    consolidated = pd.concat(frames, ignore_index=True)
    consolidated["Fecha"] = consolidated["Fecha"].dt.strftime("%Y-%m-%d")
    return consolidated[OUTPUT_COLUMNS]


def export_csv(rows: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    rows.to_csv(output_path, index=False, encoding="utf-8-sig")


def append_to_database(rows: pd.DataFrame) -> dict[str, int]:
    if rows.empty:
        return {"inserted": 0, "duplicates": 0, "original": 0}

    ensure_table_exists()

    sheet_columns = ["ID", "Fecha", "Descripción", "Monto", "Tipo", "Origen", "Categoria"]
    payload = rows.reindex(columns=sheet_columns).copy()
    payload["Fecha"] = pd.to_datetime(payload["Fecha"], errors="coerce").dt.date
    payload["ID"] = payload["ID"].astype(str).str.strip()
    original_count = len(payload)
    payload = payload[payload["Fecha"].notna() & payload["ID"].ne("")].drop_duplicates(subset=["ID"], keep="first")
    records = [
        (
            row["ID"],
            row["Fecha"],
            row["Descripción"],
            row["Monto"],
            row["Tipo"],
            row["Origen"],
            row["Categoria"],
        )
        for _, row in payload.fillna("").iterrows()
    ]

    if not records:
        return {"inserted": 0, "duplicates": original_count, "original": original_count}

    with get_db_connection() as conn:
        with conn.cursor() as cur:
            existing_ids = set()
            cur.execute(f"select id from {TABLE_NAME} where id = any(%s)", ([record[0] for record in records],))
            existing_ids.update(row[0] for row in cur.fetchall())

            filtered_records = [record for record in records if record[0] not in existing_ids]
            inserted_count = len(filtered_records)
            duplicates_count = original_count - inserted_count

            if filtered_records:
                execute_values(
                    cur,
                    f"insert into {TABLE_NAME} (id, fecha, descripcion, monto, tipo, origen, categoria) values %s",
                    filtered_records,
                )
        conn.commit()

    return {"inserted": inserted_count, "duplicates": duplicates_count, "original": original_count}


def process_folder(folder: Path = INPUT_DIR, debug: bool = False) -> pd.DataFrame:
    if not folder.exists():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {folder}")
    return read_excel_files(folder, debug=debug)


def upload_dataframe(rows: pd.DataFrame) -> dict[str, int]:
    return append_to_database(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Importa Excel bancarios a Neon o exporta un CSV local.")
    parser.add_argument("--input-dir", default=str(INPUT_DIR), help="Carpeta con archivos Excel")
    parser.add_argument("--export-csv", default="", help="Ruta para exportar un CSV local")
    parser.add_argument("--debug", action="store_true", help="Muestra diagnóstico de lectura y columnas")
    parser.add_argument("--skip-upload", action="store_true", help="No sube nada a Neon")
    args = parser.parse_args()

    consolidated = process_folder(Path(args.input_dir), debug=args.debug)

    if args.export_csv:
        export_csv(consolidated, Path(args.export_csv))
        print(f"CSV exportado en: {args.export_csv}")

    uploaded = 0
    if not args.skip_upload:
        uploaded = upload_dataframe(consolidated)

    print(f"Transacciones consolidadas: {len(consolidated)}")
    print(f"Transacciones subidas con éxito: {uploaded}")


if __name__ == "__main__":
    main()
