# -*- coding: utf-8 -*-
"""
Extraccion de datos desde PDFs de reconciliacion detallada de Veridapt.

El sistema Veridapt genera un PDF por producto con la reconciliacion
semanal. Este modulo extrae los datos relevantes para alimentar
automaticamente el Excel de entrada del reporte.

Datos extraidos de cada PDF:
  - Producto (nombre y clave del reporte).
  - Periodo (fecha inicio y fin).
  - Opening Stock, Closing Stock, Inflow (deliveries), Outflow (transactions).
  - Dispensing: To Equipment, Other Dispenses, Transfers out.
"""
from __future__ import annotations

import os
import re

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


# -----------------------------------------------------------------------
# Mapeo de nombres de producto del PDF a las claves del reporte
# -----------------------------------------------------------------------

# Cada clave es un product_key del reporte; los valores son fragmentos que,
# al encontrarse en el nombre del producto (en mayusculas, sin espacios),
# identifican al producto.
_PRODUCT_PATTERNS: dict[str, list[str]] = {
    "S4CX30":       ["S4CX30"],
    "15W40":        ["15W40", "RIMULA"],
    "S5CFDM60":     ["S5CFDM60"],
    "Tellus S3M46": ["TELLUS", "S3M46"],
}


def is_available() -> bool:
    """Devuelve True si pdfplumber esta instalado."""
    return pdfplumber is not None


def _parse_number(text: str) -> float:
    """Convierte texto tipo '23,843.16 L' a float."""
    text = str(text).strip()
    # Quitar unidad 'L' y '%' al final.
    text = re.sub(r'\s*[L%]$', '', text)
    text = text.replace(',', '')
    try:
        return float(text)
    except ValueError:
        return 0.0


def _match_product(name: str) -> str | None:
    """Asocia un nombre de producto del PDF a la clave del reporte.

    Devuelve la clave (ej. '15W40') o None si no se reconoce.
    No debe confundir S4CX10W con S4CX30: S4CX10W NO contiene 'S4CX30'.
    """
    if not name:
        return None
    up = name.upper().replace(" ", "")
    # S4CX10W no es un producto del reporte consolidado; evitar falso
    # positivo si alguien busca 'S4CX30' dentro de 'S4CX10W'.
    if "S4CX10W" in up:
        return None
    for key, fragments in _PRODUCT_PATTERNS.items():
        for frag in fragments:
            if frag in up:
                return key
    return None


# -----------------------------------------------------------------------
# Regex para extraer datos de la pagina 1 (Product Summary)
# -----------------------------------------------------------------------

# Fila de la tabla Product Summary.  Formato esperado:
#   <error%>  <product>  <opening> L  <closing> L  <inflow> L  <outflow> L  <error> L
_RE_SUMMARY_ROW = re.compile(
    r'[\d.]+%\s+'              # Error As % Of Outflow
    r'(.+?)\s+'                # Product name (non-greedy)
    r'([\d,]+\.?\d*)\s*L\s+'   # Opening Stock
    r'([\d,]+\.?\d*)\s*L\s+'   # Closing Stock
    r'([\d,]+\.?\d*)\s*L\s+'   # Inflow
    r'([\d,]+\.?\d*)\s*L\s+'   # Outflow
    r'([\d,]+\.?\d*)\s*L'      # Error
)

# Periodo: dd/mm/yyyy HH:MM - dd/mm/yyyy HH:MM
_RE_PERIOD = re.compile(
    r'(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}\s*-\s*'
    r'(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}'
)

# Titulo con nombre de producto: "Detailed Reconciliation <producto>"
_RE_TITLE = re.compile(r'Detailed Reconciliation\s+(.+)')

# Campos de dispensing en las secciones de tanque (pagina 2+).
_RE_TO_EQUIPMENT   = re.compile(r'To Equipment\s+([\d,]+\.?\d*)\s*L')
_RE_OTHER_DISPENSES = re.compile(r'Other Dispenses\s+([\d,]+\.?\d*)\s*L')
_RE_TRANSFERS_OUT  = re.compile(r'Transfers out\s+([\d,]+\.?\d*)\s*L')


# -----------------------------------------------------------------------
# Funcion principal de parsing
# -----------------------------------------------------------------------

def parse_veridapt_pdf(path: str) -> dict:
    """Extrae los datos de un PDF de reconciliacion detallada de Veridapt.

    Parametros
    ----------
    path : str
        Ruta al archivo PDF.

    Retorna
    -------
    dict con claves:
        product_key    : str | None  - clave del reporte ('15W40', etc.)
        product_name   : str         - nombre tal como aparece en el PDF
        opening        : float       - Opening Stock (litros)
        closing        : float       - Closing Stock (litros)
        inflow         : float       - Inflow total (= deliveries)
        outflow        : float       - Outflow total (= transactions)
        to_equipment   : float       - Dispensing to Equipment (litros)
        other_dispenses: float       - Other Dispenses (litros)
        transfers_out  : float       - Transfers out / Service Trucks (litros)
        period_start   : str         - Fecha inicio (dd/mm/yyyy)
        period_end     : str         - Fecha fin (dd/mm/yyyy)

    Raises
    ------
    ImportError  si pdfplumber no esta instalado.
    Exception    si el archivo no se puede leer.
    """
    if pdfplumber is None:
        raise ImportError(
            "Se requiere la libreria 'pdfplumber' para leer PDFs.\n"
            "Instale con:  pip install pdfplumber")

    with pdfplumber.open(path) as pdf:
        pages_text = [page.extract_text() or "" for page in pdf.pages]

    page1 = pages_text[0] if pages_text else ""
    full_text = "\n".join(pages_text)

    result = {
        "product_key": None,
        "product_name": "",
        "opening": 0.0,
        "closing": 0.0,
        "inflow": 0.0,
        "outflow": 0.0,
        "to_equipment": 0.0,
        "other_dispenses": 0.0,
        "transfers_out": 0.0,
        "period_start": "",
        "period_end": "",
    }

    # 1) Producto: primero del titulo, luego de la tabla Product Summary.
    for line in page1.split("\n"):
        mt = _RE_TITLE.match(line.strip())
        if mt:
            candidate = mt.group(1).strip()
            # Descartar la linea subtitulo "Detailed Reconciliation"
            if candidate.lower() not in ("detailed reconciliation", ""):
                result["product_name"] = candidate
                result["product_key"] = _match_product(candidate)
                break

    # 2) Periodo.
    mp = _RE_PERIOD.search(page1)
    if mp:
        result["period_start"] = mp.group(1)
        result["period_end"] = mp.group(2)

    # 3) Product Summary: Opening, Closing, Inflow, Outflow.
    ms = _RE_SUMMARY_ROW.search(page1)
    if ms:
        result["opening"] = _parse_number(ms.group(2))
        result["closing"] = _parse_number(ms.group(3))
        result["inflow"] = _parse_number(ms.group(4))
        result["outflow"] = _parse_number(ms.group(5))
        # Si no se pudo determinar el producto del titulo, intentar con
        # el nombre de la fila del Product Summary.
        if result["product_key"] is None:
            summary_name = ms.group(1).strip()
            result["product_name"] = summary_name
            result["product_key"] = _match_product(summary_name)

    # 4) Dispensing: sumar los valores de todas las secciones de tanque.
    #    (un PDF puede tener varios tanques para un mismo producto)
    for val in _RE_TO_EQUIPMENT.findall(full_text):
        result["to_equipment"] += _parse_number(val)
    for val in _RE_OTHER_DISPENSES.findall(full_text):
        result["other_dispenses"] += _parse_number(val)
    for val in _RE_TRANSFERS_OUT.findall(full_text):
        result["transfers_out"] += _parse_number(val)

    return result


def parse_multiple_pdfs(paths: list[str]) -> tuple[list, list]:
    """Parsea varios PDFs de Veridapt.

    Retorna
    -------
    (loaded, skipped)
        loaded  : lista de (product_key, info_dict, filename)
        skipped : lista de (filename, razon)
    """
    loaded = []
    skipped = []
    for path in paths:
        fname = os.path.basename(path)
        try:
            info = parse_veridapt_pdf(path)
        except Exception as exc:
            skipped.append((fname, str(exc)))
            continue
        if info["product_key"] is None:
            name = info["product_name"] or "(desconocido)"
            skipped.append((fname, "Producto no reconocido: %s" % name))
            continue
        loaded.append((info["product_key"], info, fname))
    return loaded, skipped
