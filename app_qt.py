# -*- coding: utf-8 -*-
"""
Interfaz grafica moderna (PySide6) del generador del
"Weekly Lubes tank reconciliation report".

Funcionalidades:
  - Carga del Excel historico "Reconciliation Lubs Weekly Ver2.xlsx".
  - Selector de fecha amigable: se eligen semanas de una lista desplegable.
  - Extraccion de datos por producto y por fecha; si una fecha no tiene
    informacion registrada, se avisa explicitamente.
  - Edicion manual de todos los datos en tablas y formularios.
  - Adjuntado de las 18 figuras + la tabla de tareas.
  - Generacion del reporte Word con un clic.

Ejecutar:  python run.py
"""
from __future__ import annotations

import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFrame, QGroupBox, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit,
    QProgressDialog, QPushButton, QScrollArea, QStyleFactory, QTabWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

import report_model as m
import history
import excel_writer
import pdf_import
import stock_trend

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, "plantilla_reporte.docx")

# Paleta de la interfaz.
PRIMARY = "#1F4E78"
ACCENT = "#2E7D32"
DANGER = "#C62828"
BG = "#F4F6F9"
TEXT = "#202124"        # texto principal (gris muy oscuro, casi negro)
TEXT_MUTED = "#5F6368"  # texto secundario

# IMPORTANTE: En Windows 11 con tema oscuro del sistema, Qt heredaria texto
# blanco/claro sobre fondo blanco, dejando los campos ilegibles. Por eso a
# cada control que muestra texto se le fuerza un 'color' explicito.
STYLESHEET = f"""
QMainWindow, QWidget {{ background: {BG}; color: {TEXT}; }}
QGroupBox {{
    font-weight: bold; color: {PRIMARY};
    border: 1px solid #C9D3DF; border-radius: 8px;
    margin-top: 14px; padding: 10px;
}}
QGroupBox::title {{
    subcontrol-origin: margin; left: 12px; padding: 0 4px;
    color: {PRIMARY};
}}
QLabel {{ color: {TEXT}; background: transparent; }}
QStatusBar {{ color: {TEXT}; background: {BG}; }}
QStatusBar QLabel {{ color: {TEXT}; }}
QPushButton {{
    background: {PRIMARY}; color: white; border: none;
    border-radius: 6px; padding: 7px 14px; font-weight: bold;
}}
QPushButton:hover {{ background: #2A5F92; color: white; }}
QPushButton:disabled {{ background: #9AA8B8; color: #ECECEC; }}
QPushButton#accent {{ background: {ACCENT}; color: white; }}
QPushButton#accent:hover {{ background: #388E3C; color: white; }}
QTabWidget::pane {{
    border: 1px solid #C9D3DF; border-radius: 6px; background: white;
}}
QTabBar::tab {{
    background: #E3E9F0; color: {TEXT};
    padding: 8px 16px; margin-right: 2px;
    border-top-left-radius: 6px; border-top-right-radius: 6px;
}}
QTabBar::tab:selected {{ background: white; color: {PRIMARY}; font-weight: bold; }}
QTabBar::tab:hover {{ background: #D6DEE8; color: {TEXT}; }}
QTableWidget {{
    background: white; color: {TEXT}; gridline-color: #DCE3EB;
    selection-background-color: #CFE3F5; selection-color: {TEXT};
}}
QTableWidget::item {{ color: {TEXT}; }}
QHeaderView::section {{
    background: {PRIMARY}; color: white; padding: 6px; border: none;
    font-weight: bold;
}}
QComboBox, QLineEdit, QPlainTextEdit {{
    background: white; color: {TEXT};
    border: 1px solid #C9D3DF; border-radius: 5px; padding: 4px;
    selection-background-color: #CFE3F5; selection-color: {TEXT};
}}
QComboBox:disabled, QLineEdit:disabled, QPlainTextEdit:disabled {{
    background: #ECEEF1; color: {TEXT_MUTED};
}}
QComboBox QAbstractItemView {{
    background: white; color: {TEXT};
    selection-background-color: #CFE3F5; selection-color: {TEXT};
    border: 1px solid #C9D3DF;
}}
QComboBox::drop-down {{ border: none; width: 18px; }}
QScrollArea {{ background: {BG}; border: none; }}
QMessageBox {{ background: {BG}; color: {TEXT}; }}
QMessageBox QLabel {{ color: {TEXT}; }}
QToolTip {{ background: #FFFFE1; color: {TEXT}; border: 1px solid #C9D3DF; }}
QLabel#title {{ font-size: 16px; font-weight: bold; color: {PRIMARY}; }}
QProgressDialog {{ background: white; }}
QProgressDialog QLabel {{ color: {PRIMARY}; font-weight: bold; }}
QProgressBar {{
    background: #E3E9F0; border: 1px solid #C9D3DF;
    border-radius: 4px; text-align: center; min-height: 16px;
}}
QProgressBar::chunk {{ background: {PRIMARY}; border-radius: 3px; }}
"""


def _apply_light_palette(app: QApplication) -> None:
    """Fuerza el estilo Fusion y una paleta clara, sin importar el tema del
    sistema operativo (clave en Windows 11 con modo oscuro: si no se hace,
    Qt hereda texto blanco para QLabel/QLineEdit dejando los campos
    ilegibles)."""
    if "Fusion" in QStyleFactory.keys():
        app.setStyle("Fusion")

    pal = QPalette()
    bg = QColor(BG)
    base = QColor("white")
    text = QColor(TEXT)
    disabled = QColor(TEXT_MUTED)

    pal.setColor(QPalette.Window, bg)
    pal.setColor(QPalette.WindowText, text)
    pal.setColor(QPalette.Base, base)
    pal.setColor(QPalette.AlternateBase, QColor("#F0F3F7"))
    pal.setColor(QPalette.Text, text)
    pal.setColor(QPalette.Button, QColor(PRIMARY))
    pal.setColor(QPalette.ButtonText, QColor("white"))
    pal.setColor(QPalette.BrightText, QColor("white"))
    pal.setColor(QPalette.ToolTipBase, QColor("#FFFFE1"))
    pal.setColor(QPalette.ToolTipText, text)
    pal.setColor(QPalette.Highlight, QColor("#CFE3F5"))
    pal.setColor(QPalette.HighlightedText, text)
    pal.setColor(QPalette.PlaceholderText, disabled)

    # Estado deshabilitado: que no se vea blanco-sobre-blanco.
    pal.setColor(QPalette.Disabled, QPalette.WindowText, disabled)
    pal.setColor(QPalette.Disabled, QPalette.Text, disabled)
    pal.setColor(QPalette.Disabled, QPalette.ButtonText, QColor("#ECECEC"))

    app.setPalette(pal)


def kpi_card(title: str, value: str, color: str) -> QLabel:
    """Tarjeta KPI con titulo y valor, coloreada segun severidad."""
    lbl = QLabel(f"<b>{title}</b><br><span style='font-size:14px'>{value}</span>")
    lbl.setTextFormat(Qt.RichText)
    lbl.setAlignment(Qt.AlignCenter)
    lbl.setStyleSheet(
        f"QLabel {{ background: white; border: 2px solid {color}; "
        f"border-radius: 8px; padding: 8px 16px; color: {color}; }}")
    return lbl


# ----------------------------------------------------------------------
# Utilidad: ejecutar una operacion lenta con barra de progreso
# ----------------------------------------------------------------------

class _BackgroundWorker(QThread):
    """QThread que ejecuta una funcion arbitraria y guarda su resultado /
    excepcion en un contenedor compartido.

    Se usa para mover operaciones lentas (lectura de Excel/CSV grandes,
    generacion del reporte) fuera del hilo de la GUI, asi la ventana
    sigue respondiendo y se puede mostrar una barra de progreso.
    """

    def __init__(self, func, args=(), kwargs=None, holder=None):
        super().__init__()
        self._func = func
        self._args = args
        self._kwargs = kwargs or {}
        self._holder = holder if holder is not None else {}

    def run(self):
        try:
            self._holder["result"] = self._func(*self._args, **self._kwargs)
        except BaseException as exc:           # noqa: BLE001
            self._holder["error"] = exc


def _run_with_progress(parent, title: str, message: str,
                        func, *args, **kwargs):
    """Ejecuta `func(*args, **kwargs)` en un hilo aparte y mientras tanto
    muestra una barra de progreso indeterminada (animada) modal.

    - El usuario ve que la app esta trabajando (no parece colgada).
    - La ventana principal queda bloqueada hasta que termine la tarea.
    - Si `func` lanza una excepcion, se relanza desde aqui.
    """
    holder: dict = {"result": None, "error": None}
    worker = _BackgroundWorker(func, args, kwargs, holder)

    dlg = QProgressDialog(message, None, 0, 0, parent)
    dlg.setWindowTitle(title)
    dlg.setWindowModality(Qt.ApplicationModal)
    dlg.setMinimumDuration(0)           # mostrar de inmediato
    dlg.setAutoClose(False)
    dlg.setAutoReset(False)
    dlg.setCancelButton(None)           # no permitir cancelar
    # Quitar el boton de cierre de la barra de titulo para que el usuario
    # no la pueda cerrar a la fuerza durante la carga.
    dlg.setWindowFlags(
        (dlg.windowFlags() | Qt.CustomizeWindowHint)
        & ~Qt.WindowCloseButtonHint
        & ~Qt.WindowContextHelpButtonHint)
    dlg.setMinimumWidth(380)

    # Cuando el worker termina, cerramos el dialogo (dlg.exec() retorna).
    worker.finished.connect(dlg.close)

    worker.start()
    dlg.exec()
    worker.wait()                       # asegurar limpieza del hilo

    if holder["error"] is not None:
        raise holder["error"]
    return holder["result"]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generador de Reporte Semanal de Lubricantes  -  "
                            "Newmont Merian FMS")
        self.resize(1180, 820)
        self.setStyleSheet(STYLESHEET)

        self.data = m.default_data()
        self.history = history.HistoryStore()
        self.image_edits: dict = {}
        self.template_path = DEFAULT_TEMPLATE
        self._current_product = m.PRODUCT_KEYS[0]

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.addWidget(self._build_controls())
        layout.addWidget(self._build_kpis())
        layout.addWidget(self._build_tabs(), stretch=1)
        layout.addWidget(self._build_footer())

        self.statusBar().showMessage(
            "Cargue el Excel historico para extraer datos por fecha, "
            "o edite los datos manualmente.")
        self._refresh_all_from_data()

    # ====================================================================
    # Construccion de la interfaz
    # ====================================================================

    def _build_load_button(self, text: str, slot, key):
        """Crea un boton de carga con un pequeño indicador de estado debajo.
        Si `key` es None, no se crea indicador (boton de accion pura, p.ej.
        'Crear Excel en blanco')."""
        col = QVBoxLayout()
        col.setSpacing(2)
        btn = QPushButton(text)
        btn.clicked.connect(slot)
        col.addWidget(btn)
        if key is not None:
            lbl = QLabel("(sin cargar)")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet(
                "color: %s; font-size: 11px; font-style: italic; "
                "background: transparent;" % TEXT_MUTED)
            self.status_labels[key] = lbl
            col.addWidget(lbl)
        else:
            # Placeholder para alinear visualmente con los botones que si
            # tienen indicador (mantiene la altura de las columnas).
            spacer = QLabel(" ")
            spacer.setStyleSheet("font-size: 11px; background: transparent;")
            col.addWidget(spacer)
        return col

    def _set_loaded_status(self, key: str, text: str | None) -> None:
        """Actualiza el indicador del boton `key`. `text=None` -> resetea a
        '(sin cargar)' en gris; cualquier otro texto se muestra en verde
        prefijado con un check."""
        lbl = self.status_labels.get(key)
        if lbl is None:
            return
        if text is None:
            lbl.setText("(sin cargar)")
            lbl.setStyleSheet(
                "color: %s; font-size: 11px; font-style: italic; "
                "background: transparent;" % TEXT_MUTED)
            lbl.setToolTip("")
        else:
            lbl.setText("✓ " + text)
            lbl.setStyleSheet(
                "color: %s; font-size: 11px; font-weight: bold; "
                "background: transparent;" % ACCENT)
            lbl.setToolTip(text)

    def _build_controls(self) -> QWidget:
        box = QGroupBox("Datos y seleccion de fecha")
        outer = QVBoxLayout(box)

        # Cada boton de carga lleva debajo un indicador con "(sin cargar)"
        # en gris o "OK <archivo>" en verde, para que de un vistazo se vea
        # que fuentes ya estan cargadas y cuales faltan.
        self.status_labels = {}

        row1 = QHBoxLayout()
        row1.addLayout(self._build_load_button(
            "Cargar Excel historico...", self._on_load_history, "historico"))
        row1.addLayout(self._build_load_button(
            "Cargar tendencia de tanques (CSV)...",
            self._on_load_stock_trend, "trend"))
        row1.addLayout(self._build_load_button(
            "Cargar PDF Veridapt...", self._on_load_pdf, "pdf"))
        row1.addLayout(self._build_load_button(
            "Cargar Excel de entrada...", self._on_load_input, "input"))
        row1.addLayout(self._build_load_button(
            "Crear Excel en blanco...", self._on_create_blank, None))
        row1.addStretch(1)
        outer.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Semana disponible:"))
        self.date_combo = QComboBox()
        self.date_combo.setMinimumWidth(170)
        row2.addWidget(self.date_combo)
        self.btn_fetch = QPushButton("Traer datos de esta fecha")
        self.btn_fetch.setObjectName("accent")
        self.btn_fetch.clicked.connect(self._on_fetch_week)
        row2.addWidget(self.btn_fetch)
        row2.addSpacing(24)
        row2.addWidget(QLabel("Periodo (texto del reporte):"))
        self.period_edit = QLineEdit(self.data["period_full"])
        self.period_edit.setMinimumWidth(210)
        row2.addWidget(self.period_edit)
        row2.addSpacing(14)
        row2.addWidget(QLabel("Fecha de portada:"))
        self.cover_edit = QLineEdit(self.data["cover_date"])
        self.cover_edit.setMinimumWidth(140)
        row2.addWidget(self.cover_edit)
        row2.addStretch(1)
        outer.addLayout(row2)
        return box

    def _build_kpis(self) -> QWidget:
        frame = QFrame()
        frame.setFrameShape(QFrame.StyledPanel)
        frame.setStyleSheet("QFrame { background: transparent; }")
        self.kpi_layout = QHBoxLayout(frame)
        self.kpi_layout.setContentsMargins(0, 0, 0, 0)
        return frame

    def _build_tabs(self) -> QWidget:
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_consolidated_tab(), "Tabla Consolidada")
        self.tabs.addTab(self._build_products_tab(), "Productos")
        self.tabs.addTab(self._build_deliveries_tab(), "Entregas (Tickets)")
        self.tabs.addTab(self._build_images_tab(), "Imagenes / Figuras")
        return self.tabs

    def _build_consolidated_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel("Ingrese Opening / Deliveries / Transactions / "
                             "Closing. El resto de columnas se calcula solo."))
        self.tbl_cons = QTableWidget(len(m.PRODUCT_KEYS), 5)
        self.tbl_cons.setHorizontalHeaderLabels(
            ["Producto", "Opening Stock", "Deliveries",
             "Transactions", "Closing Stock"])
        self.tbl_cons.verticalHeader().setVisible(False)
        self.tbl_cons.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        self.tbl_cons.itemChanged.connect(self._on_cons_changed)
        lay.addWidget(self.tbl_cons)
        return w

    def _build_products_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)

        top = QHBoxLayout()
        top.addWidget(QLabel("Producto:"))
        self.product_combo = QComboBox()
        self.product_combo.addItems(m.PRODUCT_KEYS)
        self.product_combo.currentTextChanged.connect(
            self._on_product_selected)
        top.addWidget(self.product_combo)
        top.addSpacing(20)
        top.addWidget(QLabel("Traer este producto de la fecha:"))
        self.product_date_combo = QComboBox()
        self.product_date_combo.setMinimumWidth(150)
        top.addWidget(self.product_date_combo)
        btn_pfetch = QPushButton("Traer producto")
        btn_pfetch.clicked.connect(self._on_fetch_product)
        top.addWidget(btn_pfetch)
        top.addStretch(1)
        lay.addLayout(top)

        lay.addWidget(QLabel("Frase de variance (si se deja vacia se genera "
                             "automaticamente):"))
        self.ed_recon = QPlainTextEdit()
        self.ed_recon.setFixedHeight(56)
        lay.addWidget(self.ed_recon)

        lay.addWidget(QLabel("Consideraciones (una por linea):"))
        self.ed_cons = QPlainTextEdit()
        self.ed_cons.setFixedHeight(110)
        lay.addWidget(self.ed_cons)

        lay.addWidget(QLabel("Narrativa de la figura de entregas (vacia = "
                             "automatica):"))
        self.ed_narr = QPlainTextEdit()
        self.ed_narr.setFixedHeight(56)
        lay.addWidget(self.ed_narr)

        grid = QHBoxLayout()
        grid.addWidget(QLabel("Issued:"))
        self.ed_issued = QLineEdit()
        grid.addWidget(self.ed_issued, stretch=1)
        lay.addLayout(grid)

        disp = QHBoxLayout()
        disp.addWidget(QLabel("Dispensing to Equipment (L):"))
        self.ed_eq = QLineEdit()
        disp.addWidget(self.ed_eq)
        disp.addWidget(QLabel("Other equipment (L):"))
        self.ed_ot = QLineEdit()
        disp.addWidget(self.ed_ot)
        disp.addWidget(QLabel("Transfers / Service Trucks (L):"))
        self.ed_tr = QLineEdit()
        disp.addWidget(self.ed_tr)
        lay.addLayout(disp)

        btn_apply = QPushButton("Aplicar cambios a este producto")
        btn_apply.clicked.connect(self._apply_product_form)
        lay.addWidget(btn_apply, alignment=Qt.AlignLeft)
        lay.addStretch(1)
        return w

    def _build_deliveries_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel("Una fila por ticket de entrega. Variance y % se "
                             "calculan solos. (No estan en el Excel historico: "
                             "se cargan a mano.)"))
        self.tbl_del = QTableWidget(0, 4)
        self.tbl_del.setHorizontalHeaderLabels(
            ["Producto", "Fecha", "Volume", "Docket Volume"])
        self.tbl_del.verticalHeader().setVisible(False)
        self.tbl_del.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        lay.addWidget(self.tbl_del)
        btns = QHBoxLayout()
        btn_add = QPushButton("Agregar fila")
        btn_add.clicked.connect(self._add_delivery_row)
        btn_del = QPushButton("Eliminar fila seleccionada")
        btn_del.clicked.connect(self._del_delivery_row)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        btns.addStretch(1)
        lay.addLayout(btns)
        return w

    def _build_images_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        lay = QVBoxLayout(inner)
        nota = QLabel(
            "Las Figuras 1 a 18 las construye el software automaticamente:\n"
            " - Fig. 1-8: con los datos del Excel (entregas y transacciones).\n"
            " - Fig. 9, 11, 13, 15, 17 (Tank Logs): requieren cargar el CSV "
            "de tendencia de tanques.\n"
            " - Fig. 10, 12, 14, 16, 18 (Reconciliation): con las hojas "
            "WeeklyVariance del Excel historico.\n"
            "Solo adjunte una imagen si quiere reemplazar la generada. "
            "La tabla de tareas (fig_tasks) si se adjunta manualmente.")
        nota.setWordWrap(True)
        nota.setStyleSheet("color:#1F4E78; font-weight:bold;")
        lay.addWidget(nota)
        auto = {"fig%d" % i for i in range(1, 19)}
        for fig in m.FIGURE_KEYS:
            row = QHBoxLayout()
            suffix = "   [AUTOMATICA]" if fig in auto else ""
            label = QLabel(m.FIGURE_LABELS[fig] + suffix)
            label.setFixedWidth(330)
            edit = QLineEdit()
            self.image_edits[fig] = edit
            btn = QPushButton("Examinar")
            btn.clicked.connect(lambda _=False, f=fig: self._pick_image(f))
            row.addWidget(label)
            row.addWidget(edit, stretch=1)
            row.addWidget(btn)
            lay.addLayout(row)
        lay.addStretch(1)
        scroll.setWidget(inner)
        return scroll

    def _build_footer(self) -> QWidget:
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        btn_save = QPushButton("Guardar Excel de entrada...")
        btn_save.clicked.connect(self._on_save_input)
        lay.addWidget(btn_save)
        lay.addStretch(1)
        btn_gen = QPushButton("GENERAR REPORTE")
        btn_gen.setObjectName("accent")
        btn_gen.setMinimumWidth(220)
        f = QFont()
        f.setBold(True)
        btn_gen.setFont(f)
        btn_gen.clicked.connect(self._on_generate)
        lay.addWidget(btn_gen)
        return w

    # ====================================================================
    # Sincronizacion datos <-> widgets
    # ====================================================================

    def _refresh_all_from_data(self):
        self.period_edit.setText(self.data["period_full"])
        self.cover_edit.setText(self.data.get("cover_date", ""))
        self._refresh_consolidated()
        self._refresh_deliveries()
        self._on_product_selected(self.product_combo.currentText())
        self._refresh_kpis()

    def _refresh_consolidated(self):
        self.tbl_cons.blockSignals(True)
        for i, row in enumerate(self.data["consolidated"]):
            item0 = QTableWidgetItem(str(row["site"]))
            item0.setFlags(item0.flags() & ~Qt.ItemIsEditable)
            item0.setBackground(QColor("#E3E9F0"))
            item0.setForeground(QColor(TEXT))
            self.tbl_cons.setItem(i, 0, item0)
            for col, key in enumerate(
                    ["opening", "deliveries", "transactions", "closing"], 1):
                it = QTableWidgetItem(self._fmt_num(row[key]))
                it.setForeground(QColor(TEXT))
                self.tbl_cons.setItem(i, col, it)
        self.tbl_cons.blockSignals(False)

    def _refresh_deliveries(self):
        self.tbl_del.setRowCount(0)
        for key in m.PRODUCT_KEYS:
            for d in self.data["deliveries"].get(key, []):
                self._insert_delivery_row(key, d["date"], d["volume"],
                                          d["docket"])

    def _insert_delivery_row(self, product, date, volume, docket):
        r = self.tbl_del.rowCount()
        self.tbl_del.insertRow(r)
        combo = QComboBox()
        combo.addItems(m.PRODUCT_KEYS)
        if product in m.PRODUCT_KEYS:
            combo.setCurrentText(product)
        self.tbl_del.setCellWidget(r, 0, combo)
        self.tbl_del.setItem(r, 1, QTableWidgetItem(str(date)))
        self.tbl_del.setItem(r, 2, QTableWidgetItem(self._fmt_num(volume)))
        self.tbl_del.setItem(r, 3, QTableWidgetItem(self._fmt_num(docket)))

    @staticmethod
    def _fmt_num(value) -> str:
        n = m._num(value)
        return str(int(n)) if n == int(n) else str(n)

    def _on_cons_changed(self, _item):
        self._collect_consolidated()
        self._refresh_kpis()

    def _on_product_selected(self, key):
        self._current_product = key
        p = self.data["products"].get(key, {})
        self.ed_recon.setPlainText(p.get("recon_sentence", ""))
        self.ed_cons.setPlainText("\n".join(p.get("considerations", [])))
        self.ed_narr.setPlainText(p.get("delivery_narrative", ""))
        self.ed_issued.setText(p.get("issued", ""))
        self.ed_eq.setText(self._fmt_num(p.get("disp_equipment", 0)))
        self.ed_ot.setText(self._fmt_num(p.get("disp_other", 0)))
        self.ed_tr.setText(self._fmt_num(p.get("disp_transfers", 0)))

    def _apply_product_form(self):
        key = self._current_product
        cons = [c.strip() for c in
                self.ed_cons.toPlainText().split("\n") if c.strip()]
        self.data["products"][key] = {
            "recon_sentence": self.ed_recon.toPlainText().strip(),
            "considerations": cons,
            "delivery_narrative": self.ed_narr.toPlainText().strip(),
            "issued": self.ed_issued.text().strip(),
            "disp_equipment": m._num(self.ed_eq.text()),
            "disp_other": m._num(self.ed_ot.text()),
            "disp_transfers": m._num(self.ed_tr.text()),
        }
        self.statusBar().showMessage("Cambios aplicados al producto %s." % key)

    def _collect_consolidated(self):
        rows = []
        for i in range(self.tbl_cons.rowCount()):
            def cell(c):
                it = self.tbl_cons.item(i, c)
                return it.text() if it else ""
            rows.append({
                "site": cell(0),
                "opening": m._num(cell(1)), "deliveries": m._num(cell(2)),
                "transactions": m._num(cell(3)), "closing": m._num(cell(4)),
            })
        self.data["consolidated"] = rows

    def _collect_deliveries(self):
        deliveries = {k: [] for k in m.PRODUCT_KEYS}
        for r in range(self.tbl_del.rowCount()):
            combo = self.tbl_del.cellWidget(r, 0)
            key = combo.currentText() if combo else m.PRODUCT_KEYS[0]

            def cell(c):
                it = self.tbl_del.item(r, c)
                return it.text() if it else ""
            deliveries.setdefault(key, []).append({
                "date": cell(1), "volume": m._num(cell(2)),
                "docket": m._num(cell(3)),
            })
        self.data["deliveries"] = deliveries

    def _collect_all(self) -> dict:
        self._apply_product_form()
        self._collect_consolidated()
        self._collect_deliveries()
        self.data["period_full"] = self.period_edit.text().strip()
        self.data["cover_date"] = self.cover_edit.text().strip()
        return self.data

    # ====================================================================
    # KPIs
    # ====================================================================

    def _refresh_kpis(self):
        while self.kpi_layout.count():
            item = self.kpi_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for raw in self.data["consolidated"]:
            try:
                computed = m.compute_consolidated_row(raw)
            except Exception:
                continue
            pct = computed.get("_pct_value", 0.0)
            mag = abs(pct)
            color = ACCENT if mag < 2 else "#E0A000" if mag < 6 else DANGER
            self.kpi_layout.addWidget(kpi_card(
                str(raw["site"]),
                "Variance %s" % computed["pct"], color))
        self.kpi_layout.addStretch(1)

    # ====================================================================
    # Acciones - Excel historico y extraccion por fecha
    # ====================================================================

    def _populate_date_combos(self):
        dates = self.history.available_dates()
        for combo in (self.date_combo, self.product_date_combo):
            combo.clear()
            for day in dates:
                combo.addItem(day.strftime("%d/%m/%Y"), day)

    def _on_load_history(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar Excel historico de reconciliacion", "",
            "Excel (*.xlsx)")
        if not path:
            return
        try:
            _run_with_progress(
                self,
                "Cargando Excel historico",
                "Leyendo  %s ...\n\nPor favor espere, este archivo puede ser "
                "grande." % os.path.basename(path),
                self.history.load, path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                 "No se pudo leer el Excel historico:\n%s" % exc)
            return
        dates = self.history.available_dates()
        if not dates:
            QMessageBox.warning(self, "Sin datos",
                                "El Excel no contiene hojas 'Recon ...' con "
                                "datos reconocibles.")
            return
        self._populate_date_combos()
        self._set_loaded_status(
            "historico",
            "%s (%d semanas)" % (os.path.basename(path), len(dates)))
        self.statusBar().showMessage(
            "Historico cargado. Seleccione una semana y pulse "
            "'Traer datos de esta fecha'.")

    def _on_load_stock_trend(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar CSV de tendencia de tanques (stock_trend)", "",
            "CSV (*.csv)")
        if not path:
            return
        try:
            trends, safe = _run_with_progress(
                self,
                "Cargando tendencia de tanques",
                "Leyendo CSV %s ..." % os.path.basename(path),
                stock_trend.load_tank_trends, path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                 "No se pudo leer el CSV:\n%s" % exc)
            return
        if not trends:
            QMessageBox.warning(
                self, "Sin datos",
                "El CSV no contiene columnas de tanque reconocibles "
                "('Tank WS - ...').")
            return
        self.data["tank_trends"] = trends
        self.data["tank_safe_fill"] = safe
        total = sum(len(v) for v in trends.values())
        self._set_loaded_status(
            "trend",
            "%s (%d tanques)" % (os.path.basename(path), len(trends)))
        QMessageBox.information(
            self, "Tendencia cargada",
            "Tendencia de tanques cargada (%d mediciones).\nLos Tank Logs "
            "(Figuras 9, 11, 13, 15, 17) se generaran automaticamente."
            % total)
        self.statusBar().showMessage(
            "Tendencia de tanques cargada: %s" % os.path.basename(path))

    def _on_load_pdf(self):
        """Carga uno o varios PDFs de reconciliacion Veridapt y extrae
        Opening, Closing, Inflow (deliveries), Outflow (transactions) y el
        desglose de dispensing (To Equipment, Other Dispenses, Transfers out)
        para cada producto reconocido.

        Ademas, si el Excel historico esta cargado, escribe una fila nueva
        (o actualiza la existente) en la hoja 'Recon <producto>' para que
        la informacion del PDF quede registrada en el Excel.
        """
        if not pdf_import.is_available():
            QMessageBox.critical(
                self, "Libreria faltante",
                "Se requiere la libreria 'pdfplumber' para leer PDFs.\n\n"
                "Instale con:\n  pip install pdfplumber")
            return

        paths, _ = QFileDialog.getOpenFileNames(
            self, "Cargar PDF de reconciliacion Veridapt", "",
            "PDF (*.pdf)")
        if not paths:
            return

        self._collect_all()
        try:
            loaded, skipped = _run_with_progress(
                self,
                "Cargando PDFs Veridapt",
                "Extrayendo datos de %d archivo(s) PDF..." % len(paths),
                pdf_import.parse_multiple_pdfs, paths)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                 "No se pudieron leer los PDFs:\n%s" % exc)
            return

        # ---------------------------------------------------------------
        # 1) Aplicar los datos extraidos sobre la estructura en memoria
        # ---------------------------------------------------------------
        for key, info, _fname in loaded:
            # Tabla consolidada: opening, deliveries, transactions, closing.
            for row in self.data["consolidated"]:
                if row["site"] == key:
                    row["opening"] = round(info["opening"], 2)
                    row["deliveries"] = round(info["inflow"], 2)
                    row["transactions"] = round(info["outflow"], 2)
                    row["closing"] = round(info["closing"], 2)
                    break

            # Producto: dispensing.
            prod = self.data["products"].setdefault(key, {})
            prod["disp_equipment"] = round(info["to_equipment"], 2)
            prod["disp_other"] = round(info["other_dispenses"], 2)
            prod["disp_transfers"] = round(info["transfers_out"], 2)

        # Periodo: tomar del primer PDF cargado.
        period_end_date = None
        if loaded:
            info0 = loaded[0][1]
            if info0["period_start"] and info0["period_end"]:
                try:
                    from datetime import datetime as _dt
                    start = _dt.strptime(info0["period_start"],
                                         "%d/%m/%Y").date()
                    end = _dt.strptime(info0["period_end"],
                                       "%d/%m/%Y").date()
                    self.data["period_full"] = m.format_period(start, end)
                    self.data["cover_date"] = m.format_cover_date(end)
                    period_end_date = end
                except ValueError:
                    pass

        # ---------------------------------------------------------------
        # 2) Escribir los datos al Excel historico (si esta cargado)
        # ---------------------------------------------------------------
        excel_results = []
        excel_errors = []
        excel_path = (self.history.source_path
                      if self.history.is_loaded() else None)

        if loaded and excel_path and period_end_date is not None:
            items = [(k, period_end_date, i) for (k, i, _fn) in loaded]
            try:
                excel_results, excel_errors = _run_with_progress(
                    self,
                    "Escribiendo Excel historico",
                    "Insertando filas en el Excel historico...",
                    excel_writer.write_multiple, excel_path, items)
            except Exception as exc:
                excel_errors = [("(write)", str(exc))]

            # Re-cargar el historico para que la fecha nueva aparezca en el
            # combo "Semana disponible" y para que las tendencias se
            # actualicen.
            try:
                _run_with_progress(
                    self,
                    "Recargando Excel historico",
                    "Refrescando datos...",
                    self.history.load, excel_path)
                self._populate_date_combos()
                idx = self.date_combo.findText(
                    period_end_date.strftime("%d/%m/%Y"))
                if idx >= 0:
                    self.date_combo.setCurrentIndex(idx)
                    pidx = self.product_date_combo.findText(
                        period_end_date.strftime("%d/%m/%Y"))
                    if pidx >= 0:
                        self.product_date_combo.setCurrentIndex(pidx)

                # Auto-aplicar los datos historicos sobre la estructura.
                # Importante para S4CX30: el PDF aporta SOLO Tank 2; en
                # el Excel ahora tambien existe Tank1 (S4CX10W).  Al
                # re-aplicar, la tabla consolidada se rellena con la
                # AGREGACION (Tank 2 + Tank 1) que es lo que el reporte
                # final necesita.
                try:
                    history.apply_week_to_data(
                        self.data, self.history, period_end_date)
                except Exception:
                    pass
            except Exception:
                pass

        self._refresh_all_from_data()

        # ---------------------------------------------------------------
        # 3) Mostrar resumen al usuario
        # ---------------------------------------------------------------
        msg = []
        if loaded:
            msg.append("Productos actualizados desde PDF:\n")
            for key, info, fname in loaded:
                msg.append(
                    "  • %s  (Opening: %s,  Closing: %s,  "
                    "Deliveries: %s,  Transactions: %s)\n"
                    "       Dispensing → Equipment: %s,  Other: %s,  "
                    "Transfers: %s\n"
                    "       Archivo: %s" % (
                        key,
                        m.fmt_int(info["opening"]),
                        m.fmt_int(info["closing"]),
                        m.fmt_int(info["inflow"]),
                        m.fmt_int(info["outflow"]),
                        m.fmt_int(info["to_equipment"]),
                        m.fmt_int(info["other_dispenses"]),
                        m.fmt_int(info["transfers_out"]),
                        fname))

        if excel_results:
            msg.append("\nFilas escritas en el Excel historico (%s):"
                       % os.path.basename(excel_path))
            for key, res in excel_results:
                action = res["action"]
                if action == "inserted":
                    txt = "fila %d insertada" % res["row"]
                elif action == "updated":
                    txt = "fila %d actualizada" % res["row"]
                elif action == "no_sheet":
                    txt = "sin hoja 'Recon ...' para este producto (omitido)"
                else:
                    txt = action
                msg.append("  • %s → hoja '%s', %s, Tank='%s'"
                           % (key, res["sheet"] or "-", txt,
                              res["tank"]))
        elif loaded and not excel_path:
            msg.append("\nNota: el Excel historico no esta cargado. "
                       "Los datos solo se aplicaron en pantalla. "
                       "Cargue el Excel historico antes de cargar los "
                       "PDFs para que tambien queden guardados.")

        if excel_errors:
            msg.append("\nErrores al escribir el Excel:")
            for key, reason in excel_errors:
                msg.append("  • %s: %s" % (key, reason))

        if skipped:
            msg.append("\nArchivos ignorados:")
            for fname, reason in skipped:
                msg.append("  • %s: %s" % (fname, reason))

        if loaded:
            # Indicador junto al boton: nombre del PDF si fue uno solo,
            # cantidad si fueron varios.
            if len(paths) == 1:
                status_txt = os.path.basename(paths[0])
            else:
                status_txt = "%d PDFs cargados" % len(paths)
            self._set_loaded_status("pdf", status_txt)
            QMessageBox.information(
                self, "PDF cargado", "\n".join(msg))
        else:
            QMessageBox.warning(
                self, "Sin datos de lubricantes",
                "\n".join(msg) if msg else
                "No se encontraron productos de lubricantes en los PDFs "
                "seleccionados.")

        if excel_errors:
            self.statusBar().showMessage(
                "PDF Veridapt: %d producto(s) actualizado(s) en pantalla; "
                "%d error(es) al escribir Excel." %
                (len(loaded), len(excel_errors)))
        elif excel_results:
            self.statusBar().showMessage(
                "PDF Veridapt: %d producto(s) actualizado(s) y escrito(s) "
                "al Excel historico." % len(loaded))
        else:
            self.statusBar().showMessage(
                "PDF Veridapt: %d producto(s) actualizado(s) (solo en "
                "pantalla)." % len(loaded))

    def _on_fetch_week(self):
        if not self.history.is_loaded():
            QMessageBox.information(
                self, "Sin historico",
                "Primero cargue el Excel historico de reconciliacion.")
            return
        day = self.date_combo.currentData()
        if day is None:
            return
        self._collect_all()
        found, missing = history.apply_week_to_data(self.data, self.history, day)
        # El periodo y la fecha de portada se derivan de la fecha elegida,
        # para que TODO el reporte quede consistente.
        period, cover = self.history.period_for(day)
        self.data["period_full"] = period
        self.data["cover_date"] = cover
        self._refresh_all_from_data()
        if missing:
            names = ", ".join(missing)
            QMessageBox.warning(
                self, "Informacion no registrada",
                "La fecha %s NO tiene informacion registrada para:\n\n%s\n\n"
                "Esos productos quedaron sin actualizar; revise o ingrese sus "
                "datos manualmente." % (day.strftime("%d/%m/%Y"), names))
        self.statusBar().showMessage(
            "Datos del %s aplicados. Encontrados: %d  -  No registrados: %d"
            % (day.strftime("%d/%m/%Y"), len(found), len(missing)))

    def _on_fetch_product(self):
        if not self.history.is_loaded():
            QMessageBox.information(
                self, "Sin historico",
                "Primero cargue el Excel historico de reconciliacion.")
            return
        day = self.product_date_combo.currentData()
        key = self.product_combo.currentText()
        if day is None:
            return
        rec = self.history.lookup(key, day)
        if rec is None:
            QMessageBox.warning(
                self, "Informacion no registrada",
                "El producto '%s' NO tiene informacion registrada para la "
                "fecha %s." % (key, day.strftime("%d/%m/%Y")))
            return
        self._collect_all()
        for row in self.data["consolidated"]:
            if row["site"] == key:
                row["opening"] = round(rec.opening, 2)
                row["deliveries"] = round(rec.deliveries, 2)
                row["transactions"] = round(rec.transactions, 2)
                row["closing"] = round(rec.closing, 2)
        prod = self.data["products"].setdefault(key, {})
        prod["disp_equipment"] = round(rec.to_equipment, 2)
        prod["disp_other"] = round(rec.other, 2)
        prod["disp_transfers"] = round(rec.transfers, 2)
        self.data["deliveries"][key] = \
            self.history.deliveries_for_week(day).get(key, [])
        self._refresh_all_from_data()
        self.statusBar().showMessage(
            "Producto '%s' actualizado con datos del %s (tanques: %s)."
            % (key, day.strftime("%d/%m/%Y"), ", ".join(rec.tanks) or "-"))

    # ====================================================================
    # Acciones - Excel de entrada / generacion
    # ====================================================================

    def _on_load_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar Excel de entrada", "", "Excel (*.xlsx)")
        if not path:
            return
        try:
            self.data = _run_with_progress(
                self,
                "Cargando Excel de entrada",
                "Leyendo  %s ..." % os.path.basename(path),
                m.load_excel, path)
            self._refresh_all_from_data()
            self._set_loaded_status("input", os.path.basename(path))
            self.statusBar().showMessage("Datos cargados desde %s"
                                         % os.path.basename(path))
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                 "No se pudo cargar el Excel:\n%s" % exc)

    def _on_create_blank(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Crear Excel de entrada en blanco",
            "datos_reporte_lubricantes.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            m.save_excel(m.default_data(), path)
            QMessageBox.information(self, "Excel creado",
                                    "Excel de entrada creado:\n%s" % path)
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def _on_save_input(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Excel de entrada",
            "datos_reporte_lubricantes.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            m.save_excel(self._collect_all(), path)
            QMessageBox.information(self, "Guardado",
                                    "Datos guardados en:\n%s" % path)
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def _pick_image(self, fig):
        path, _ = QFileDialog.getOpenFileName(
            self, m.FIGURE_LABELS[fig], "",
            "Imagenes (*.png *.jpg *.jpeg);;Todos (*.*)")
        if path:
            self.image_edits[fig].setText(path)

    def _add_delivery_row(self):
        key = self.product_combo.currentText()
        self._insert_delivery_row(key, "", 0, 0)

    def _del_delivery_row(self):
        rows = sorted({i.row() for i in self.tbl_del.selectedIndexes()},
                      reverse=True)
        for r in rows:
            self.tbl_del.removeRow(r)

    def _on_generate(self):
        if not os.path.isfile(self.template_path):
            QMessageBox.critical(
                self, "Falta la plantilla",
                "No se encuentra la plantilla Word:\n%s\n\nEjecute primero "
                "prepare_template.py." % self.template_path)
            return
        # Avisar si faltan datos para las Figuras 9-18.
        missing = []
        if not self.data.get("tank_trends"):
            missing.append("Tank Logs (Figuras 9, 11, 13, 15, 17) - "
                           "falta cargar el CSV 'stock_trend'")
        if not self.data.get("recon_trends"):
            missing.append("Reconciliation trend (Figuras 10, 12, 14, 16, 18) "
                           "- falta cargar el Excel historico y traer una fecha")
        if missing:
            reply = QMessageBox.question(
                self, "Faltan datos para algunas figuras",
                "Las siguientes figuras quedaran en blanco en el reporte:\n\n"
                + "\n".join("  - " + m_ for m_ in missing)
                + "\n\nGenerar igual?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        out, _ = QFileDialog.getSaveFileName(
            self, "Guardar reporte generado",
            "Weekly Lubes tank reconciliation report.docx", "Word (*.docx)")
        if not out:
            return
        try:
            data = self._collect_all()
            images = {fig: self.image_edits[fig].text().strip()
                      for fig in m.FIGURE_KEYS}
            _run_with_progress(
                self,
                "Generando reporte",
                "Generando graficos y armando el documento Word ...\n\n"
                "Esto puede tardar varios segundos.",
                m.generate_report, data, images, self.template_path, out)
            QMessageBox.information(self, "Reporte generado",
                                    "Reporte creado correctamente:\n%s" % out)
            self.statusBar().showMessage("Reporte generado: %s" % out)
        except Exception as exc:
            QMessageBox.critical(self, "Error al generar",
                                 "%s\n\n%s" % (exc, traceback.format_exc()))


def launch() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    # Forzar paleta clara: en Windows 11 con tema oscuro, sin esto, las
    # etiquetas y los campos se renderizan con texto blanco sobre blanco.
    _apply_light_palette(app)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(launch())
