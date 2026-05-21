# -*- coding: utf-8 -*-
"""
Generacion de los graficos del reporte con matplotlib.

Replica dos graficos que hoy se hacen en Excel y se copian a mano:

  - Grafico de barras "Deliveries"  -> Figuras 1 a 4
        Sum of Volume vs Sum of Docket Volume por cada ticket de entrega.

  - Grafico de dona "TRANSACTIONS DISTRIBUTION"  -> Figuras 5 a 8
        To equipment / Other Dispenses / Transfers.

El software los construye y los inserta directamente en el Word; ya no hace
falta copiar y pegar desde el Excel.
"""
from __future__ import annotations

import matplotlib
matplotlib.use("Agg")          # backend sin ventana (solo guarda archivos)
import matplotlib.dates as mdates
import matplotlib.pyplot as plt

# Colores del tema de Excel (para que se vea igual al original).
EXCEL_BLUE = "#4472C4"
EXCEL_ORANGE = "#ED7D31"
EXCEL_GRAY = "#A5A5A5"


def _num(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _save(fig, path: str) -> str:
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def delivery_bar_chart(rows: list, path: str) -> str:
    """Grafico de barras de entregas (Figuras 1-4).

    rows: lista de dicts {date, volume, docket}. Si esta vacia se genera un
    grafico que indica que no hubo entregas confirmadas.
    """
    # figsize ajustado para que al insertar a 105 mm de ancho la altura
    # resulte ~63 mm, igual que las imagenes del reporte hecho a mano
    # (compensa los margenes que matplotlib agrega con bbox_inches='tight').
    fig, ax = plt.subplots(figsize=(6.0, 3.1))
    ax.set_title("Deliveries", fontsize=12, fontweight="bold")

    if not rows:
        ax.text(0.5, 0.5, "No confirmed deliveries in this period",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)

    labels = [str(r.get("date", "")) for r in rows]
    volume = [_num(r.get("volume")) for r in rows]
    docket = [_num(r.get("docket")) for r in rows]
    x = range(len(rows))
    width = 0.38

    ax.bar([i - width / 2 for i in x], volume, width,
           label="Sum of Volume", color=EXCEL_BLUE)
    ax.bar([i + width / 2 for i in x], docket, width,
           label="Sum of Docket Volume", color=EXCEL_ORANGE)
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, rotation=0, fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12),
              ncol=2, frameon=False, fontsize=9)
    return _save(fig, path)


def transaction_pie_chart(equipment, other, transfers, path: str) -> str:
    """Grafico de dona "TRANSACTIONS DISTRIBUTION" (Figuras 5-8)."""
    # figsize ajustado para que al insertar a 95 mm de ancho la altura sea
    # ~70 mm, igual que las donas del reporte hecho a mano.
    fig, ax = plt.subplots(figsize=(5.0, 3.3))
    ax.set_title("TRANSACTIONS DISTRIBUTION", fontsize=12,
                 fontweight="bold", pad=24)

    values = [_num(equipment), _num(other), _num(transfers)]
    labels = ["To equipment", "Other Dispenses", "Transfers"]
    colors = [EXCEL_BLUE, EXCEL_ORANGE, EXCEL_GRAY]

    if sum(values) <= 0:
        ax.text(0.5, 0.5, "No transactions in this period",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)

    wedges, _texts, autotexts = ax.pie(
        values, colors=colors, startangle=90, counterclock=False,
        autopct=lambda p: ("%.0f%%" % round(p)) if p > 0 else "",
        pctdistance=0.78, wedgeprops=dict(width=0.42, edgecolor="white"))
    for at in autotexts:
        at.set_color("white")
        at.set_fontsize(9)
    ax.legend(wedges, labels, loc="upper center",
              bbox_to_anchor=(0.5, 1.10), ncol=3, frameon=False, fontsize=9)
    ax.set(aspect="equal")
    return _save(fig, path)


def _fmt_thousands(value) -> str:
    """30000 -> '30,000.00' (formato anglosajon del grafico de Excel)."""
    return "{:,.2f}".format(_num(value))


def tank_log_chart(points: list, title: str, path: str,
                   safe_fill=None) -> str:
    """Grafico de linea del nivel del tanque (Figuras 9, 11, 13, 15, 17).

    Replica el estilo del grafico exportado por Excel: titulo gris arriba a
    la izquierda, linea roja de Safe Fill Level, linea azul del volumen.
    """
    # figsize ajustado para que al insertar a 158 mm de ancho la altura
    # sea ~65 mm, igual que los Tank Log de los reportes hechos a mano.
    fig, ax = plt.subplots(figsize=(8.0, 2.6))
    ax.set_title(title, loc="left", fontsize=11, color="#7F7F7F",
                 pad=10, fontweight="normal")

    if not points:
        ax.text(0.5, 0.5, "No tank-level data available",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)

    xs = [p[0] for p in points]
    ys = [_num(p[1]) for p in points]
    ax.plot(xs, ys, color="#4472C4", linewidth=1.6)

    if safe_fill:
        sf = _num(safe_fill)
        ax.axhline(sf, color="#C00000", linewidth=1.4)
        ax.text(xs[0], sf, " Safe Fill Level (%s L)" % _fmt_thousands(sf),
                color="#C00000", fontsize=9, fontweight="bold",
                va="bottom", ha="left")

    ax.set_ylabel("Volume", color="#7F7F7F", fontsize=10)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m/%Y %H:%M"))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(maxticks=14))
    fig.autofmt_xdate(rotation=45)

    ax.grid(axis="y", alpha=0.5, color="#D9D9D9", linewidth=0.7)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color("#D9D9D9")
    ax.tick_params(colors="#7F7F7F", labelsize=8)
    return _save(fig, path)


def recon_trend_chart(points: list, title: str, path: str) -> str:
    """Grafico de tendencia semanal de variance % (Figuras 10,12,...).

    points: lista de (date, recon_pct). Replica el grafico
    "Weekly variance change <producto>" de las hojas WeeklyVariance.
    """
    # figsize ajustado para que al insertar a 150 mm de ancho la altura sea
    # ~80 mm, igual que los graficos 'Reconciliation' hechos a mano.
    fig, ax = plt.subplots(figsize=(7.5, 3.5))
    ax.set_title(title, fontsize=12, fontweight="bold")

    if not points:
        ax.text(0.5, 0.5, "No reconciliation history available",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)

    labels = [p[0].strftime("%d/%m/%Y") for p in points]
    ys = [_num(p[1]) for p in points]
    x = list(range(len(points)))
    ax.plot(x, ys, color=EXCEL_ORANGE, linewidth=1.8, marker="o",
            markersize=4, label="Recon %")
    for i, v in zip(x, ys):
        ax.annotate("%.2f" % v, (i, v), textcoords="offset points",
                    xytext=(0, 7), ha="center", fontsize=7)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
    ax.set_ylabel("Variance (%)", fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    ax.legend(loc="upper left", fontsize=8, frameon=False)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    return _save(fig, path)
