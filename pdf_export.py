from __future__ import annotations

import io
import calendar
from typing import Optional

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from fpdf import FPDF


# ── colour palette ────────────────────────────────────────────────────────────
_BG         = (15,  20,  35)     # dark navy
_PANEL      = (25,  32,  50)     # card background
_GREEN      = (34, 197,  94)
_RED        = (239, 68,  68)
_BLUE       = (122, 162, 255)
_TEXT_MAIN  = (240, 244, 255)
_TEXT_DIM   = (148, 163, 184)
_BORDER     = (50,  62,  90)


def _month_name_es(month: int) -> str:
    names = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ]
    return names[month - 1]


def _fmt(amount: float, symbol: str = "$") -> str:
    return f"{symbol} {amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _week_of_month(dt: pd.Timestamp) -> int:
    """Return 1-based week number within the month."""
    return (dt.day - 1) // 7 + 1


def _build_weekly_chart(frame: pd.DataFrame, mes: int, año: int, moneda: str) -> bytes:
    """Return a PNG bar chart (bytes) of weekly expenses for the given month."""
    symbol = "$" if moneda == "ARS" else "U$S"
    month_days = calendar.monthrange(año, mes)[1]
    week_labels = []
    week_ends = []
    d = 1
    while d <= month_days:
        end = min(d + 6, month_days)
        week_labels.append(f"{d}-{end}")
        week_ends.append(end)
        d += 7

    n_weeks = len(week_labels)
    weekly_totals = [0.0] * n_weeks

    for _, row in frame.iterrows():
        if str(row.get("Tipo", "")).lower() == "egreso":
            try:
                dt = pd.to_datetime(row.get("Fecha_dt") or row.get("Fecha"), errors="coerce")
                if pd.isna(dt):
                    continue
                wk = min(_week_of_month(dt) - 1, n_weeks - 1)
                weekly_totals[wk] += abs(float(row.get("Monto", 0) or 0))
            except Exception:
                continue

    fig, ax = plt.subplots(figsize=(7.2, 2.8))
    fig.patch.set_facecolor(tuple(c / 255 for c in _BG))
    ax.set_facecolor(tuple(c / 255 for c in _PANEL))

    bar_color = tuple(c / 255 for c in _RED)
    bars = ax.bar(week_labels, weekly_totals, color=bar_color, width=0.55, zorder=3)

    for bar, val in zip(bars, weekly_totals):
        if val > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + max(weekly_totals) * 0.02,
                f"{symbol} {val:,.0f}".replace(",", "."),
                ha="center", va="bottom", fontsize=7,
                color=tuple(c / 255 for c in _TEXT_DIM),
            )

    ax.set_xlabel("Semana del mes", fontsize=8, color=tuple(c / 255 for c in _TEXT_DIM))
    ax.set_ylabel(f"Egresos ({moneda})", fontsize=8, color=tuple(c / 255 for c in _TEXT_DIM))
    ax.set_title(
        f"Gastos por semana - {_month_name_es(mes)} {año}",
        fontsize=10, color=tuple(c / 255 for c in _TEXT_MAIN), pad=10,
    )
    ax.tick_params(colors=tuple(c / 255 for c in _TEXT_DIM), labelsize=7)
    for spine in ax.spines.values():
        spine.set_edgecolor(tuple(c / 255 for c in _BORDER))
    ax.yaxis.grid(True, color=tuple(c / 255 for c in _BORDER), linewidth=0.5, zorder=0)
    ax.set_axisbelow(True)

    plt.tight_layout(pad=0.5)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf.read()


class _PDF(FPDF):
    def __init__(self, mes: int, año: int, moneda: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self._mes  = mes
        self._año  = año
        self._mon  = moneda
        self.set_margins(14, 14, 14)
        self.set_auto_page_break(auto=True, margin=16)

    # ── header ────────────────────────────────────────────────────────────────
    def header(self):
        # gradient-like banner
        self.set_fill_color(*_BG)
        self.rect(0, 0, 210, 28, style="F")

        self.set_fill_color(*_BLUE)
        self.rect(0, 26, 210, 2, style="F")

        self.set_y(7)
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(*_TEXT_MAIN)
        label = f"Resumen Mensual - {_month_name_es(self._mes)} {self._año}  |  {self._mon}"
        self.cell(0, 10, label, align="C")
        self.ln(14)

    # ── footer ────────────────────────────────────────────────────────────────
    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*_TEXT_DIM)
        self.cell(0, 8, f"Pagina {self.page_no()} - Home Banking", align="C")

    # ── helper: filled rounded rectangle ──────────────────────────────────────
    def _kpi_card(self, x: float, y: float, w: float, h: float,
                  label: str, value: str, color: tuple) -> None:
        self.set_fill_color(*_PANEL)
        self.set_draw_color(*_BORDER)
        self.rect(x, y, w, h, style="FD")

        # accent bar on top
        self.set_fill_color(*color)
        self.rect(x, y, w, 1.5, style="F")

        self.set_xy(x + 3, y + 4)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*_TEXT_DIM)
        self.cell(w - 6, 5, label.upper())

        self.set_xy(x + 3, y + 10)
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(*color)
        self.cell(w - 6, 7, value)


def generate_monthly_pdf(
    df: pd.DataFrame,
    mes: int,
    año: int,
    moneda: str,
) -> Optional[bytes]:
    """
    Generate a PDF summary for the given DataFrame (already filtered by the caller).

    `mes` and `año` are used only for the PDF header/title.
    Returns None if the dataframe is empty.
    """
    work = df.copy()

    if "Fecha_dt" not in work.columns:
        work["Fecha_dt"] = pd.to_datetime(work.get("Fecha"), errors="coerce", dayfirst=True)

    work = work.dropna(subset=["Fecha_dt"])

    if "Moneda" in work.columns:
        work = work[work["Moneda"].str.upper() == moneda.upper()]

    if work.empty:
        return None

    # ── KPI totals ─────────────────────────────────────────────────────────
    tipo   = work.get("Tipo", pd.Series(dtype=str)).astype(str).str.lower()
    montos = pd.to_numeric(work.get("Monto"), errors="coerce").fillna(0.0).abs()
    ingresos = montos[tipo == "ingreso"].sum()
    egresos  = montos[tipo == "egreso"].sum()
    neto     = ingresos - egresos
    symbol   = "$" if moneda == "ARS" else "U$S"

    # ── weekly chart ───────────────────────────────────────────────────────
    chart_png = _build_weekly_chart(work, mes, año, moneda)

    # ── build PDF ──────────────────────────────────────────────────────────
    pdf = _PDF(mes, año, moneda)
    pdf.set_fill_color(*_BG)
    pdf.add_page()

    # background fill for whole page
    pdf.set_fill_color(*_BG)
    pdf.rect(0, 0, 210, 297, style="F")
    # re-draw header over background
    pdf.set_fill_color(*_BG)
    pdf.rect(0, 0, 210, 28, style="F")
    pdf.set_fill_color(*_BLUE)
    pdf.rect(0, 26, 210, 2, style="F")

    y_cursor = pdf.get_y()

    # ── KPI cards (3 across) ───────────────────────────────────────────────
    card_w  = 56
    card_h  = 22
    gap     = 5
    x_start = 14
    for i, (label, value, color) in enumerate([
        ("Ingresos Totales", _fmt(ingresos, symbol), _GREEN),
        ("Egresos Totales",  _fmt(egresos,  symbol), _RED),
        ("Saldo Neto",       _fmt(neto,     symbol), _BLUE if neto >= 0 else _RED),
    ]):
        pdf._kpi_card(x_start + i * (card_w + gap), y_cursor, card_w, card_h, label, value, color)

    y_cursor += card_h + 8

    # ── weekly bar chart ───────────────────────────────────────────────────
    chart_buf = io.BytesIO(chart_png)
    # fpdf2 accepts BytesIO directly
    pdf.set_xy(14, y_cursor)
    chart_h = 52
    pdf.image(chart_buf, x=14, y=y_cursor, w=182, h=chart_h)
    y_cursor += chart_h + 8

    # ── transactions table ─────────────────────────────────────────────────
    pdf.set_xy(14, y_cursor)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*_TEXT_MAIN)
    pdf.cell(0, 6, "Detalle de Movimientos", ln=True)
    y_cursor += 7

    # table header
    col_widths = [28, 88, 34, 32]
    headers    = ["Fecha", "Descripción", "Monto", "Tipo"]
    pdf.set_xy(14, y_cursor)
    pdf.set_fill_color(*_PANEL)
    pdf.set_draw_color(*_BORDER)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(*_BLUE)
    for header, w in zip(headers, col_widths):
        pdf.cell(w, 7, header, border=1, fill=True)
    pdf.ln()
    y_cursor += 7

    # rows
    pdf.set_font("Helvetica", "", 7.5)
    fill_even = (*_PANEL,)
    fill_odd  = (*_BG,)

    table_rows = work.sort_values("Fecha_dt").head(200)  # cap at 200 rows
    for idx, (_, row) in enumerate(table_rows.iterrows()):
        fecha_str = row["Fecha_dt"].strftime("%d/%m/%Y")
        desc      = str(row.get("Descripción") or row.get("Descripcion") or "")[:52]
        monto_val = abs(float(pd.to_numeric(row.get("Monto"), errors="coerce") or 0))
        monto_str = _fmt(monto_val, symbol)
        tipo_val  = str(row.get("Tipo") or "").capitalize()
        tipo_color = _GREEN if tipo_val.lower() == "ingreso" else _RED

        fill = fill_even if idx % 2 == 0 else fill_odd
        pdf.set_fill_color(*fill)
        pdf.set_draw_color(*_BORDER)

        x_row = 14
        row_h = 6

        # Fecha
        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*_TEXT_DIM)
        pdf.cell(col_widths[0], row_h, fecha_str, border=1, fill=True)
        x_row += col_widths[0]

        # Descripción
        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*_TEXT_MAIN)
        pdf.cell(col_widths[1], row_h, desc, border=1, fill=True)
        x_row += col_widths[1]

        # Monto
        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*_TEXT_MAIN)
        pdf.cell(col_widths[2], row_h, monto_str, border=1, fill=True, align="R")
        x_row += col_widths[2]

        # Tipo (coloured)
        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*tipo_color)
        pdf.cell(col_widths[3], row_h, tipo_val, border=1, fill=True, align="C")
        pdf.ln()

    # ── return bytes ───────────────────────────────────────────────────────
    return bytes(pdf.output())
