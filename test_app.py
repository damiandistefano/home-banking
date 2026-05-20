from __future__ import annotations

import datetime as dt
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

from transaction_uploader import build_transaction_id, normalize_sheet
from streamlit_app import normalize_dashboard_frame


APP_PATH = Path(__file__).resolve().parent / "streamlit_app.py"


@pytest.fixture(autouse=True)
def _database_url_env(monkeypatch):
    monkeypatch.setenv("DATABASE_URL", "postgresql://fake:fake@localhost:5432/fake")


def _make_excel_bytes(frame: pd.DataFrame, sheet_name: str = "Hoja1") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


def _write_excel_file(tmp_path: Path, filename: str, frame: pd.DataFrame) -> Path:
    path = tmp_path / filename
    path.write_bytes(_make_excel_bytes(frame))
    return path


def _galicia_frame(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=["Fecha", "Descripción", "Débito", "Crédito"])


def _rendered_text(app: AppTest) -> str:
    chunks: list[str] = []
    for attr in ("error", "warning", "info", "success", "markdown", "caption", "title", "subheader"):
        collection = getattr(app, attr, None)
        if not collection:
            continue
        for element in collection:
            for value_attr in ("value", "text", "body"):
                value = getattr(element, value_attr, None)
                if isinstance(value, str) and value:
                    chunks.append(value)
                    break
    return "\n".join(chunks)


def _click_button(app: AppTest, label: str) -> None:
    for button in getattr(app, "button", []):
        if getattr(button, "label", None) == label:
            button.click()
            return
    raise AssertionError(f"No se encontró el botón: {label}")


def _enter_upload_view(app: AppTest) -> None:
    app.session_state["view"] = "upload"
    app.run()


def _upload_files(widget, files) -> None:
    if hasattr(widget, "set_value"):
        widget.set_value(files)
        return
    raise AssertionError("El widget file_uploader no expone set_value")


def _clear_file_uploader(widget) -> None:
    if hasattr(widget, "clear"):
        widget.clear()
        try:
            widget.root.session_state[widget.id] = None
        except Exception:
            pass
        return
    for candidate in (None, []):
        try:
            if hasattr(widget, "set_value"):
                widget.set_value(candidate)
                try:
                    widget.root.session_state[widget.id] = None
                except Exception:
                    pass
                return
        except Exception:
            continue
    raise AssertionError("No se pudo limpiar el file_uploader")


def _set_date_input(widget, value: dt.date) -> None:
    candidates = ([value], (value,), value)
    for candidate in candidates:
        try:
            if hasattr(widget, "set_value"):
                widget.set_value(candidate)
                return
        except Exception:
            continue
    raise AssertionError("No se pudo forzar el date_input a un rango parcial")


class _FakeCursor:
    def __init__(self, existing_ids: set[str]):
        self.existing_ids = existing_ids
        self.fetchall_result: list[tuple[str]] = []
        self.executed: list[tuple[str, object | None]] = []

    def execute(self, query, params=None):
        self.executed.append((query, params))
        if "select id from movimientos where id = any(%s)" in str(query).lower():
            ids = list(params[0]) if params else []
            self.fetchall_result = [(row_id,) for row_id in ids if row_id in self.existing_ids]

    def fetchall(self):
        return list(self.fetchall_result)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class _FakeConnection:
    def __init__(self, cursor: _FakeCursor):
        self._cursor = cursor
        self.committed = False

    def cursor(self):
        return self._cursor

    def commit(self):
        self.committed = True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


def _dashboard_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "ID": "id-1",
                "Fecha": "2026-03-01",
                "Descripción": "Sueldo",
                "Monto": 1000,
                "Tipo": "Ingreso",
                "Origen": "Galicia - Caja Ahorro Pesos",
                "Categoria": "",
            },
            {
                "ID": "id-2",
                "Fecha": "2026-03-02",
                "Descripción": "Supermercado",
                "Monto": 250,
                "Tipo": "Egreso",
                "Origen": "Galicia - Caja Ahorro Pesos",
                "Categoria": "",
            },
            {
                "ID": "id-3",
                "Fecha": "2026-03-03",
                "Descripción": "Freelance",
                "Monto": 500,
                "Tipo": "Ingreso",
                "Origen": "Galicia - Caja Ahorro Pesos",
                "Categoria": "",
            },
        ]
    )


def _valid_galicia_rows_one() -> pd.DataFrame:
    return _galicia_frame(
        [
            {"Fecha": "2026-03-01", "Descripción": "Sueldo", "Débito": 0, "Crédito": 1000},
            {"Fecha": "2026-03-02", "Descripción": "Supermercado", "Débito": 250, "Crédito": 0},
        ]
    )


def _valid_galicia_rows_two() -> pd.DataFrame:
    return _galicia_frame(
        [
            {"Fecha": "2026-03-03", "Descripción": "Freelance", "Débito": 0, "Crédito": 500},
            {"Fecha": "2026-03-04", "Descripción": "Servicios", "Débito": 100, "Crédito": 0},
        ]
    )


def test_invalid_file_upload_shows_friendly_error_and_no_traceback(tmp_path):
    bad_file = tmp_path / "notas.xlsx"
    bad_file.write_bytes(b"esto no es un xlsx real")

    with patch("streamlit_app.read_excel_sheets", side_effect=ValueError("Formato de archivo no soportado")):
        app = AppTest.from_file(APP_PATH).run()
        _enter_upload_view(app)

        uploader = app.file_uploader[0]
        _upload_files(uploader, [(bad_file.name, bad_file.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
        app.run()

        error_text = _rendered_text(app)
        assert len(getattr(app, "exception", [])) == 0
        assert "Formato de archivo no soportado" in error_text or "Error al procesar notas.xlsx" in error_text
        assert "Traceback" not in error_text


def test_upload_lifecycle_reset_and_consolidation(tmp_path):
    file_one = _write_excel_file(tmp_path, "galicia_pesos.xlsx", _valid_galicia_rows_one())
    file_two = _write_excel_file(tmp_path, "galicia_pesos_2.xlsx", _valid_galicia_rows_two())

    app = AppTest.from_file(APP_PATH).run()
    _enter_upload_view(app)

    uploader = app.file_uploader[0]

    _upload_files(
        uploader,
        [(file_one.name, file_one.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")],
    )
    app.run()
    one_file_text = _rendered_text(app)
    assert "Detectados 2 movimientos" in one_file_text
    assert "Ingresos" in one_file_text and "$ 1.000,00" in one_file_text
    assert "Egresos" in one_file_text and "$ 250,00" in one_file_text
    assert "Saldo Neto" in one_file_text and "$ 750,00" in one_file_text

    _clear_file_uploader(uploader)
    try:
        app.session_state["uploader_0"] = None
    except Exception:
        pass
    app = app.run()
    uploader = app.file_uploader[0]
    assert "upload_files_fingerprint" not in app.session_state
    assert uploader.value in (None, [])
    assert len(getattr(app, "exception", [])) == 0

    _upload_files(
        uploader,
        [
            (file_one.name, file_one.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            (file_two.name, file_two.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ],
    )
    app.run()
    two_file_text = _rendered_text(app)
    assert "Detectados 4 movimientos" in two_file_text
    assert "$ 1.500,00" in two_file_text
    assert "$ 350,00" in two_file_text
    assert "$ 1.150,00" in two_file_text
    assert len(getattr(app, "exception", [])) == 0


def test_calendar_filter_handles_single_click_range_without_indexerror():
    frame = _dashboard_df()
    script = f'''
import datetime as dt
import pandas as pd
import streamlit as st
from streamlit_app import normalize_dashboard_frame, render_dashboard_filters

frame = normalize_dashboard_frame(pd.DataFrame({frame.to_dict(orient="records")!r}))
try:
    filtered, symbol = render_dashboard_filters(frame, "ARS")
    st.write(len(filtered))
    st.write(symbol)
except Exception as exc:
    st.error(str(exc))
'''

    app = AppTest.from_string(script).run()
    date_widget = app.date_input[0]
    _set_date_input(date_widget, dt.date(2026, 3, 1))
    app.run()

    assert len(getattr(app, "exception", [])) == 0
    assert not app.error


def test_dedup_skips_existing_transaction_ids(tmp_path):
    file_frame = _galicia_frame(
        [
            {"Fecha": "2026-03-01", "Descripción": "Duplicada", "Débito": 0, "Crédito": 150},
            {"Fecha": "2026-03-02", "Descripción": "Nueva", "Débito": 0, "Crédito": 300},
        ]
    )
    excel_path = _write_excel_file(tmp_path, "galicia_duplicados.xlsx", file_frame)
    normalized = normalize_sheet(pd.read_excel(excel_path, header=None), bank_name="galicia_duplicados", debug=False)
    duplicate_id = normalized.iloc[0]["ID"]

    fake_cursor = _FakeCursor(existing_ids={duplicate_id})
    fake_connection = _FakeConnection(fake_cursor)
    execute_values_mock = MagicMock()

    with patch("streamlit_app.load_db_data", return_value=_dashboard_df().copy()), patch(
        "transaction_uploader.ensure_table_exists", return_value=None
    ), patch("transaction_uploader.get_db_connection", return_value=fake_connection), patch(
        "transaction_uploader.execute_values", execute_values_mock
    ):
        app = AppTest.from_file(APP_PATH).run()
        _enter_upload_view(app)

        uploader = app.file_uploader[0]
        _upload_files(
            uploader,
            [(excel_path.name, excel_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")],
        )
        app.run()

        _click_button(app, "🚀 Confirmar y Subir")
        app.run()

    assert execute_values_mock.call_count == 1
    inserted_records = execute_values_mock.call_args.args[2]
    assert len(inserted_records) == 1
    assert inserted_records[0][0] != duplicate_id
    assert fake_connection.committed is True
    assert len(getattr(app, "exception", [])) == 0
