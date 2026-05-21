# -*- coding: utf-8 -*-
"""
Lectura del CSV "stock_trend_..." exportado del FMS.

Ese archivo contiene la serie de tiempo del nivel de cada tanque (una
medicion cada ~1,5 horas). Se usa para construir los graficos "Tank Log"
del reporte (Figuras 9, 11, 13, 15, 17).
"""
from __future__ import annotations

import csv
import datetime


def _num(value) -> float:
    if value is None or value == "":
        return 0.0
    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _match_tank_product(column_name: str):
    """Asocia una columna 'Tank WS - ...' con la clave del producto."""
    up = str(column_name).upper().replace(" ", "")
    if "S4CX10W" in up:
        return "S4CX10W"
    if "S4CX30" in up:
        return "S4CX30"
    if "S5CFDM60" in up:
        return "S5CFDM60"
    if "TELLUS" in up or "S3M46" in up:
        return "Tellus S3M46"
    if "15W40" in up or "RIMULA" in up:
        return "15W40"
    return None


def _parse_timestamp(text):
    """Convierte '2026-05-04 00:00:00 -0300' a datetime sin zona horaria."""
    if not text:
        return None
    raw = str(text).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S %z", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M"):
        try:
            dt = datetime.datetime.strptime(raw, fmt)
            return dt.replace(tzinfo=None)
        except ValueError:
            continue
    return None


def load_tank_trends(csv_path: str) -> tuple:
    """Lee el CSV de tendencia de tanques.

    Devuelve (trends, safe_fill):
      trends    -> {product_key: [(datetime, volumen), ...]}
      safe_fill -> {product_key: capacidad_de_llenado_seguro}
    """
    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        return {}, {}

    header = rows[0]
    # Solo interesan las columnas de tanque ("Tank WS - ...").
    col_product = {}
    for idx, name in enumerate(header):
        if name and str(name).strip().lower().startswith("tank "):
            key = _match_tank_product(name)
            if key:
                col_product[idx] = key

    trends = {key: [] for key in col_product.values()}
    safe_fill = {}

    for row in rows[1:]:
        if not row:
            continue
        label = str(row[0]).strip().lower()
        # Fila de capacidad de llenado seguro.
        if "safe fill" in label:
            for idx, key in col_product.items():
                if idx < len(row):
                    safe_fill[key] = _num(row[idx])
            continue
        ts = _parse_timestamp(row[0])
        if ts is None:
            continue
        for idx, key in col_product.items():
            if idx < len(row):
                trends[key].append((ts, _num(row[idx])))

    return trends, safe_fill
