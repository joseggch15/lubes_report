# -*- coding: utf-8 -*-
"""
Escribe los datos extraidos de los PDFs Veridapt directamente en el Excel
historico "Reconciliation Lubs Weekly Ver2.xlsx".

Para cada producto se inserta una fila nueva (o se actualiza la existente)
en la hoja "Recon <producto>". Se respetan:

  - Valores literales: A (Date), B (Site), C (Opening), D (Deliveries),
                       H (Closing), L (To equipment), M (Other Dispenses),
                       N (Transfers).
  - Formulas calculadas: E (=SUM(L:N)), F (Calculated stock),
                         G (Net Stock change), I (Variance), J (%).
    Se traducen las referencias relativas del row anterior al nuevo row.
  - Estilos (fuente, fondo, bordes, formato numerico) clonados del row
    anterior, para que la fila nueva se vea igual al resto.
"""
from __future__ import annotations

import copy
import datetime

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

import history


# Columnas (1-indexadas para openpyxl) de la hoja "Recon <producto>".
_COL_DATE = 1          # A
_COL_SITE = 2          # B
_COL_OPENING = 3       # C
_COL_DELIVERIES = 4    # D
_COL_TRANSACTIONS = 5  # E (formula =SUM(L:N))
_COL_CALC_STOCK = 6    # F (formula)
_COL_NET_CHANGE = 7    # G (formula)
_COL_CLOSING = 8       # H
_COL_VARIANCE = 9      # I (formula)
_COL_PCT = 10          # J (formula)
_COL_TO_EQUIPMENT = 12 # L
_COL_OTHER = 13        # M
_COL_TRANSFERS = 14    # N


def _last_data_row(ws) -> int:
    """Numero de la ultima fila con fecha valida en la columna A."""
    last = 2  # filas 1-2 son encabezados
    # Recorremos solo las primeras 16 columnas (ws.max_column suele estar
    # inflado por estilos en columnas vacias).
    for row_idx in range(3, ws.max_row + 1):
        if isinstance(ws.cell(row=row_idx, column=_COL_DATE).value,
                      datetime.datetime):
            last = row_idx
    return last


def _find_existing_row(ws, target_date: datetime.date,
                       tank_name: str | None) -> int | None:
    """Devuelve la fila si ya existe una entrada para esa fecha+tanque."""
    for row_idx in range(3, ws.max_row + 1):
        cell_date = ws.cell(row=row_idx, column=_COL_DATE).value
        if not isinstance(cell_date, datetime.datetime):
            continue
        if cell_date.date() != target_date:
            continue
        if not tank_name:
            return row_idx
        cell_site = ws.cell(row=row_idx, column=_COL_SITE).value
        if cell_site and str(cell_site).strip().lower() == tank_name.lower():
            return row_idx
    return None


def _last_row_for_tank(ws, tank_name: str | None) -> int | None:
    """Ultima fila con datos para el tanque dado (None = cualquier tanque)."""
    last = None
    for row_idx in range(3, ws.max_row + 1):
        cell_date = ws.cell(row=row_idx, column=_COL_DATE).value
        if not isinstance(cell_date, datetime.datetime):
            continue
        if tank_name is not None:
            cell_site = ws.cell(row=row_idx, column=_COL_SITE).value
            if (not cell_site or
                    str(cell_site).strip().lower() != tank_name.lower()):
                continue
        last = row_idx
    return last


def _resolve_tank_name(ws, product_key: str) -> str:
    """Decide el valor de la columna 'Site' (Tank) para una nueva fila.

    Prioriza el filtro PRODUCT_TANK_FILTER (para S4CX30 = 'Tank 21').
    Si no, hereda el tanque de la fila mas reciente de la hoja.
    """
    forced = history.PRODUCT_TANK_FILTER.get(product_key)
    if forced:
        return forced
    last = _last_row_for_tank(ws, None)
    if last is not None:
        site = ws.cell(row=last, column=_COL_SITE).value
        if site:
            return str(site).strip()
    return ""


def _compute_derived(info: dict) -> dict:
    """Calcula en Python los valores que normalmente serian formulas en el
    Excel (columnas E, F, G, I, J). Asi escribimos numeros en vez de
    formulas, y openpyxl con data_only=True puede leerlos.

    Si en cambio dejaramos la formula '=SUM(L:N)' sin valor cacheado, al
    re-leer el archivo openpyxl devuelve None y la app interpreta
    Transactions = 0, perdiendo informacion. Eso es exactamente lo que
    estaba pasando con S5CFDM60 / Tellus S3M46 al recargar.
    """
    opening = float(info.get("opening", 0) or 0)
    inflow = float(info.get("inflow", 0) or 0)
    closing = float(info.get("closing", 0) or 0)
    to_eq = float(info.get("to_equipment", 0) or 0)
    other = float(info.get("other_dispenses", 0) or 0)
    trans_out = float(info.get("transfers_out", 0) or 0)

    transactions = to_eq + other + trans_out             # E = SUM(L:N)
    calc_stock = opening + inflow - transactions          # F = C+D-E
    net_change = closing - opening                        # G = H-C
    variance = closing - calc_stock                       # I = H-F
    pct = (variance / transactions) if transactions else 0.0  # J = I/E
    return {
        "transactions": transactions,
        "calc_stock": calc_stock,
        "net_change": net_change,
        "variance": variance,
        "pct": pct,
    }


def _copy_row_template(ws, source_row: int, target_row: int,
                       max_col: int = 16) -> None:
    """Clona formulas (traducidas) y estilos del row fuente al row destino.

    max_col limita el barrido a las primeras N columnas para evitar
    procesar miles de columnas vacias.
    """
    for col in range(1, max_col + 1):
        src = ws.cell(row=source_row, column=col)
        tgt = ws.cell(row=target_row, column=col)

        # Traducir formulas relativas.
        val = src.value
        if isinstance(val, str) and val.startswith("="):
            col_letter = get_column_letter(col)
            try:
                translator = Translator(
                    val, origin="%s%d" % (col_letter, source_row))
                tgt.value = translator.translate_formula(
                    "%s%d" % (col_letter, target_row))
            except Exception:
                tgt.value = val   # si falla, dejar la formula literal

        # Clonar estilos (la fila nueva se ve igual a las anteriores).
        if src.has_style:
            tgt.font = copy.copy(src.font)
            tgt.fill = copy.copy(src.fill)
            tgt.border = copy.copy(src.border)
            tgt.alignment = copy.copy(src.alignment)
            tgt.number_format = src.number_format
            tgt.protection = copy.copy(src.protection)


def write_pdf_data(excel_path: str, product_key: str,
                   period_end: datetime.date, info: dict) -> dict:
    """Inserta o actualiza una fila en la hoja Recon del producto.

    Parametros
    ----------
    excel_path  : ruta al Excel historico.
    product_key : clave del reporte (ej. '15W40').
    period_end  : fecha que se usara como la fecha de la fila.
    info        : dict con los datos parseados del PDF (de pdf_import).

    Retorna
    -------
    dict {
        'sheet':   nombre de la hoja escrita (o None si no existe),
        'row':     numero de fila escrita,
        'action':  'inserted' | 'updated' | 'no_sheet',
        'tank':    valor escrito en columna Site,
    }
    """
    wb = openpyxl.load_workbook(excel_path)
    sheet_name = history._match_sheet(product_key, wb.sheetnames)
    if sheet_name is None:
        return {"sheet": None, "row": None, "action": "no_sheet", "tank": ""}

    ws = wb[sheet_name]
    tank_name = _resolve_tank_name(ws, product_key)

    existing = _find_existing_row(ws, period_end, tank_name or None)
    if existing is not None:
        target_row = existing
        action = "updated"
    else:
        # Determinar el row fuente para copiar formulas y estilos.
        source_row = (_last_row_for_tank(ws, tank_name or None)
                      or _last_row_for_tank(ws, None))
        target_row = _last_data_row(ws) + 1
        if source_row is not None:
            _copy_row_template(ws, source_row, target_row)
        action = "inserted"

    # Escribir TODOS los valores como numeros (incluyendo los que en el
    # Excel original eran formulas). Asi data_only=True puede leerlos al
    # recargar — antes las formulas devolvian None y la app mostraba
    # Transactions = 0 incorrectamente.
    derived = _compute_derived(info)
    ws.cell(row=target_row, column=_COL_DATE,
            value=datetime.datetime.combine(period_end, datetime.time()))
    ws.cell(row=target_row, column=_COL_SITE, value=tank_name)
    ws.cell(row=target_row, column=_COL_OPENING, value=info["opening"])
    ws.cell(row=target_row, column=_COL_DELIVERIES, value=info["inflow"])
    ws.cell(row=target_row, column=_COL_TRANSACTIONS,
            value=derived["transactions"])
    ws.cell(row=target_row, column=_COL_CALC_STOCK,
            value=derived["calc_stock"])
    ws.cell(row=target_row, column=_COL_NET_CHANGE,
            value=derived["net_change"])
    ws.cell(row=target_row, column=_COL_CLOSING, value=info["closing"])
    ws.cell(row=target_row, column=_COL_VARIANCE, value=derived["variance"])
    ws.cell(row=target_row, column=_COL_PCT, value=derived["pct"])
    ws.cell(row=target_row, column=_COL_TO_EQUIPMENT,
            value=info["to_equipment"])
    ws.cell(row=target_row, column=_COL_OTHER, value=info["other_dispenses"])
    ws.cell(row=target_row, column=_COL_TRANSFERS,
            value=info["transfers_out"])

    wb.save(excel_path)
    return {"sheet": sheet_name, "row": target_row, "action": action,
            "tank": tank_name}


def write_multiple(excel_path: str, items: list) -> tuple:
    """Escribe varios PDFs al Excel en una sola pasada.

    items: lista de (product_key, period_end:date, info_dict)
    Retorna (results, errors):
        results: lista de (product_key, dict resultado)
        errors:  lista de (product_key, str error)
    """
    results = []
    errors = []
    # Abrir el workbook una sola vez para todos los productos
    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as exc:
        return [], [("(workbook)", str(exc))]

    for product_key, period_end, info in items:
        try:
            sheet_name = history._match_sheet(product_key, wb.sheetnames)
            if sheet_name is None:
                results.append((product_key, {
                    "sheet": None, "row": None,
                    "action": "no_sheet", "tank": ""}))
                continue

            ws = wb[sheet_name]
            tank_name = _resolve_tank_name(ws, product_key)
            existing = _find_existing_row(ws, period_end,
                                          tank_name or None)
            if existing is not None:
                target_row = existing
                action = "updated"
            else:
                source_row = (_last_row_for_tank(ws, tank_name or None)
                              or _last_row_for_tank(ws, None))
                target_row = _last_data_row(ws) + 1
                if source_row is not None:
                    _copy_row_template(ws, source_row, target_row)
                action = "inserted"

            derived = _compute_derived(info)
            ws.cell(row=target_row, column=_COL_DATE,
                    value=datetime.datetime.combine(
                        period_end, datetime.time()))
            ws.cell(row=target_row, column=_COL_SITE, value=tank_name)
            ws.cell(row=target_row, column=_COL_OPENING,
                    value=info["opening"])
            ws.cell(row=target_row, column=_COL_DELIVERIES,
                    value=info["inflow"])
            ws.cell(row=target_row, column=_COL_TRANSACTIONS,
                    value=derived["transactions"])
            ws.cell(row=target_row, column=_COL_CALC_STOCK,
                    value=derived["calc_stock"])
            ws.cell(row=target_row, column=_COL_NET_CHANGE,
                    value=derived["net_change"])
            ws.cell(row=target_row, column=_COL_CLOSING,
                    value=info["closing"])
            ws.cell(row=target_row, column=_COL_VARIANCE,
                    value=derived["variance"])
            ws.cell(row=target_row, column=_COL_PCT,
                    value=derived["pct"])
            ws.cell(row=target_row, column=_COL_TO_EQUIPMENT,
                    value=info["to_equipment"])
            ws.cell(row=target_row, column=_COL_OTHER,
                    value=info["other_dispenses"])
            ws.cell(row=target_row, column=_COL_TRANSFERS,
                    value=info["transfers_out"])

            results.append((product_key, {
                "sheet": sheet_name, "row": target_row,
                "action": action, "tank": tank_name}))
        except Exception as exc:
            errors.append((product_key, str(exc)))

    # Guardar todos los cambios al final.
    try:
        wb.save(excel_path)
    except Exception as exc:
        errors.append(("(save)", str(exc)))

    return results, errors
