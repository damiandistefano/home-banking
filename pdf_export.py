from __future__ import annotations

import io
import calendar
from typing import Optional

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from fpdf import FPDF


# ── colour palette ────────────────────────────────────────────────────────────
_BG        = (15,  20,  35)
_PANEL     = (25,  32,  50)
_GREEN     = (34, 197,  94)
_RED       = (239,  68,  68)
_BLUE      = (122, 162, 255)
_TEXT_MAIN = (240, 244, 255)
_TEXT_DIM  = (148, 163, 184)
_BORDER    = (50,  62,  90)

# Donut palette (FinTech dark)
_DONUT_COLORS = [
    "#7aa2ff", "#f97316", "#a78bfa", "#34d399", "#fb7185",
    "#fbbf24", "#38bdf8", "#c084fc", "#4ade80", "#f472b6",
]


def _month_name_es(month: int) -> str:
    names = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ]
    return names[month - 1]


def _fmt(amount: float, symbol: str = "$") -> str:
    return f"{symbol} {amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _week_of_month(dt: pd.Timestamp) -> int:
    return (dt.day - 1) // 7 + 1


def _fmt_axis_amount(x: float, _) -> str:
    """Human-readable axis tick: avoids scientific notation."""
    if x >= 1_000_000:
        return f"{x / 1_000_000:,.1f}M".replace(",", ".")
    if x >= 1_000:
        return f"{x / 1_000:,.0f}K".replace(",", ".")
    return f"{x:,.0f}".replace(",", ".")


def _build_weekly_chart(frame: pd.DataFrame, mes: int, año: int, moneda: str) -> bytes:
    """Weekly bar chart of expenses for the given month."""
    import matplotlib.ticker as mticker

    symbol = "$" if moneda == "ARS" else "U$S"
    month_days = calendar.monthrange(año, mes)[1]
    week_labels: list[str] = []
    d = 1
    while d <= month_days:
        end = min(d + 6, month_days)
        week_labels.append(f"{d}-{end}")
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

    fig, ax = plt.subplots(figsize=(4.5, 3.0))
    fig.patch.set_facecolor(tuple(c / 255 for c in _BG))
    ax.set_facecolor(tuple(c / 255 for c in _PANEL))

    bar_color = tuple(c / 255 for c in _RED)
    bars = ax.bar(week_labels, weekly_totals, color=bar_color, width=0.55, zorder=3)

    max_val = max(weekly_totals) if any(weekly_totals) else 1
    for bar, val in zip(bars, weekly_totals):
        if val > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + max_val * 0.02,
                f"{symbol} {_fmt_axis_amount(val, None)}",
                ha="center", va="bottom", fontsize=6.5,
                color=tuple(c / 255 for c in _TEXT_DIM),
            )

    ax.set_xlabel("Semana del mes", fontsize=7.5, color=tuple(c / 255 for c in _TEXT_DIM))
    ax.set_ylabel(f"Egresos ({moneda})", fontsize=7.5, color=tuple(c / 255 for c in _TEXT_DIM))
    ax.set_title(
        f"Gastos por semana — {_month_name_es(mes)} {año}",
        fontsize=9, color=tuple(c / 255 for c in _TEXT_MAIN), pad=8,
    )
    ax.tick_params(colors=tuple(c / 255 for c in _TEXT_DIM), labelsize=6.5)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_axis_amount))
    for spine in ax.spines.values():
        spine.set_edgecolor(tuple(c / 255 for c in _BORDER))
    ax.yaxis.grid(True, color=tuple(c / 255 for c in _BORDER), linewidth=0.5, zorder=0)
    ax.set_axisbelow(True)

    plt.tight_layout(pad=0.4)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _build_account_chart(frame: pd.DataFrame, moneda: str) -> bytes:
    """Donut chart of expense distribution by bank / account (Cuenta / Origen)."""
    symbol = "$" if moneda == "ARS" else "U$S"

    # Detect which column names the account
    account_col = next(
        (c for c in ["Cuenta", "Origen", "Banco"] if c in frame.columns),
        None,
    )
    egresos = frame[frame["Tipo"].astype(str).str.lower() == "egreso"].copy()
    montos = pd.to_numeric(egresos.get("Monto"), errors="coerce").fillna(0.0).abs()

    if account_col:
        labels_raw = egresos[account_col].fillna("Sin cuenta").astype(str).str.strip()
        labels_raw = labels_raw.replace("", "Sin cuenta")
        grouped = (
            pd.DataFrame({"label": labels_raw, "monto": montos})
            .groupby("label", as_index=False)["monto"]
            .sum()
            .sort_values("monto", ascending=False)
        )
        grouped = grouped[grouped["monto"] > 0]
    else:
        grouped = pd.DataFrame(columns=["label", "monto"])

    fig, ax = plt.subplots(figsize=(5.5, 3.4))
    fig.patch.set_facecolor(tuple(c / 255 for c in _BG))
    ax.set_facecolor(tuple(c / 255 for c in _BG))

    if grouped.empty:
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center",
                color=tuple(c / 255 for c in _TEXT_DIM), fontsize=10,
                transform=ax.transAxes)
        ax.axis("off")
    else:
        values = grouped["monto"].tolist()
        labels = grouped["label"].tolist()
        colors = [_DONUT_COLORS[i % len(_DONUT_COLORS)] for i in range(len(labels))]
        total = sum(values)

        wedges, texts = ax.pie(
            values,
            labels=None,
            colors=colors,
            startangle=90,
            wedgeprops=dict(width=0.55, edgecolor=tuple(c / 255 for c in _BG), linewidth=1.5),
            pctdistance=0.75,
        )

        import numpy as np
        for wedge, val in zip(wedges, values):
            pct = val / total * 100
            if pct >= 5:
                angle = (wedge.theta1 + wedge.theta2) / 2
                r = 0.68
                x = r * np.cos(np.radians(angle))
                y = r * np.sin(np.radians(angle))
                ax.text(x, y, f"{pct:.1f}%", ha="center", va="center",
                        fontsize=7, color="white", fontweight="bold")

        # Legend outside — full label (up to 28 chars) + rounded amount
        legend_labels = [
            f"{lbl[:28]}  {symbol} {_fmt_axis_amount(val, None)}"
            for lbl, val in zip(labels, values)
        ]
        ax.legend(
            wedges, legend_labels,
            loc="center left",
            bbox_to_anchor=(1.02, 0.5),
            fontsize=7.5,
            frameon=False,
            labelcolor=tuple(c / 255 for c in _TEXT_DIM),
        )

    ax.set_title(
        f"Egresos por cuenta — {moneda}",
        fontsize=9, color=tuple(c / 255 for c in _TEXT_MAIN), pad=8,
    )

    plt.tight_layout(pad=0.4)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, facecolor=fig.get_facecolor(),
                bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


class _PDF(FPDF):
    def __init__(self, mes: int, año: int, moneda: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self._mes = mes
        self._año = año
        self._mon = moneda
        self.set_margins(14, 14, 14)
        self.set_auto_page_break(auto=True, margin=16)

    def header(self):
        # 1. Full-page background — drawn first so nothing gets covered
        self.set_fill_color(*_BG)
        self.rect(0, 0, 210, 297, style="F")

        # 2. Top banner
        self.set_fill_color(*_BG)
        self.rect(0, 0, 210, 28, style="F")

        # 3. Accent line under banner
        self.set_fill_color(*_BLUE)
        self.rect(0, 26, 210, 2, style="F")

        # 4. Title text
        self.set_y(7)
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(*_TEXT_MAIN)
        label = f"Resumen Mensual - {_month_name_es(self._mes)} {self._año}  |  {self._mon}"
        self.cell(0, 10, label, align="C")
        self.ln(14)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*_TEXT_DIM)
        self.cell(0, 8, f"Pagina {self.page_no()} - Home Banking", align="C")

    def _kpi_card(self, x: float, y: float, w: float, h: float,
                  label: str, value: str, color: tuple) -> None:
        self.set_fill_color(*_PANEL)
        self.set_draw_color(*_BORDER)
        self.rect(x, y, w, h, style="FD")

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

    def _draw_table_header(self, y: float, col_widths: list[int], headers: list[str]) -> float:
        self.set_xy(14, y)
        self.set_fill_color(*_PANEL)
        self.set_draw_color(*_BORDER)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*_BLUE)
        for header, w in zip(headers, col_widths):
            self.cell(w, 7, header, border=1, fill=True)
        self.ln()
        return y + 7


def generate_monthly_pdf(
    df: pd.DataFrame,
    mes: int,
    año: int,
    moneda: str,
) -> Optional[bytes]:
    """Generate a monthly financial summary PDF and return it as bytes."""
    work = df.copy()

    if "Fecha_dt" not in work.columns:
        work["Fecha_dt"] = pd.to_datetime(work.get("Fecha"), errors="coerce", dayfirst=True)

    work = work.dropna(subset=["Fecha_dt"])

    if "Moneda" in work.columns:
        work = work[work["Moneda"].str.upper() == moneda.upper()]

    if work.empty:
        return None

    # ── KPIs ──────────────────────────────────────────────────────────────────
    tipo     = work.get("Tipo", pd.Series(dtype=str)).astype(str).str.lower()
    montos   = pd.to_numeric(work.get("Monto"), errors="coerce").fillna(0.0).abs()
    ingresos = montos[tipo == "ingreso"].sum()
    egresos  = montos[tipo == "egreso"].sum()
    neto     = ingresos - egresos
    symbol   = "$" if moneda == "ARS" else "U$S"

    # ── Charts (generated before PDF so they're ready to embed) ───────────────
    weekly_png  = _build_weekly_chart(work, mes, año, moneda)
    account_png = _build_account_chart(work, moneda)

    # ── Build PDF ─────────────────────────────────────────────────────────────
    pdf = _PDF(mes, año, moneda)
    pdf.add_page()  # header() fires here → background + banner drawn correctly

    y_cursor = pdf.get_y()

    # ── KPI cards ─────────────────────────────────────────────────────────────
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

    # ── Side-by-side charts (each ~89 mm wide, 2 mm gap) ──────────────────────
    chart_h   = 58
    chart_w   = 89
    chart_gap = 4

    pdf.image(io.BytesIO(weekly_png),  x=14,                      y=y_cursor, w=chart_w, h=chart_h)
    pdf.image(io.BytesIO(account_png), x=14 + chart_w + chart_gap, y=y_cursor, w=chart_w, h=chart_h)
    y_cursor += chart_h + 8

    # ── Transactions table ─────────────────────────────────────────────────────
    # Fecha | Descripción | Cuenta | Monto | Tipo  — total 182 mm usable
    col_widths = [25, 57, 42, 32, 26]
    headers    = ["Fecha", "Descripción", "Cuenta", "Monto", "Tipo"]

    pdf.set_xy(14, y_cursor)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*_TEXT_MAIN)
    pdf.cell(0, 6, "Detalle de Movimientos", ln=True)
    y_cursor += 7

    y_cursor = pdf._draw_table_header(y_cursor, col_widths, headers)

    pdf.set_font("Helvetica", "", 7.5)
    fill_even = (*_PANEL,)
    fill_odd  = (*_BG,)

    table_rows = work.sort_values("Fecha_dt").head(200)
    for idx, (_, row) in enumerate(table_rows.iterrows()):
        # ── Page overflow: new page + redraw header ─────────────────────────
        if pdf.get_y() > 270:
            pdf.add_page()
            y_cursor = pdf.get_y()
            y_cursor = pdf._draw_table_header(y_cursor, col_widths, headers)
            pdf.set_font("Helvetica", "", 7.5)

        fecha_str  = row["Fecha_dt"].strftime("%d/%m/%Y")
        desc       = str(row.get("Descripción") or row.get("Descripcion") or "")[:34]
        account_col = next((c for c in ["Cuenta", "Origen", "Banco"] if c in row.index), None)
        cuenta_str = str(row[account_col] if account_col else "").strip()[:26]
        monto_val  = abs(float(pd.to_numeric(row.get("Monto"), errors="coerce") or 0))
        monto_str  = _fmt(monto_val, symbol)
        tipo_val   = str(row.get("Tipo") or "").capitalize()
        tipo_color = _GREEN if tipo_val.lower() == "ingreso" else _RED

        fill  = fill_even if idx % 2 == 0 else fill_odd
        row_h = 6

        pdf.set_fill_color(*fill)
        pdf.set_draw_color(*_BORDER)

        x_row = 14
        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*_TEXT_DIM)
        pdf.cell(col_widths[0], row_h, fecha_str, border=1, fill=True)
        x_row += col_widths[0]

        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*_TEXT_MAIN)
        pdf.cell(col_widths[1], row_h, desc, border=1, fill=True)
        x_row += col_widths[1]

        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*_TEXT_DIM)
        pdf.cell(col_widths[2], row_h, cuenta_str, border=1, fill=True)
        x_row += col_widths[2]

        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*_TEXT_MAIN)
        pdf.cell(col_widths[3], row_h, monto_str, border=1, fill=True, align="R")
        x_row += col_widths[3]

        pdf.set_xy(x_row, pdf.get_y())
        pdf.set_text_color(*tipo_color)
        pdf.cell(col_widths[4], row_h, tipo_val, border=1, fill=True, align="C")
        pdf.ln()

    return bytes(pdf.output())
