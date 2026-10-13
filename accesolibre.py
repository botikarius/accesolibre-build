#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AccesoLibre 2007
================
Suite de bases de datos relacional de escritorio, de código abierto,
inspirada funcionalmente en Microsoft Access 2007: cinta de opciones,
panel de navegación y objetos Tabla / Consulta / Formulario / Informe / Macro.

Motor interno: SQLite (archivo .oal).
Compatibilidad nativa con ACCDB/MDB mediante Jackcess (LGPL).
Licencia: GPL-3.0-or-later.
"""

import csv
import html
import json
import os
import re
import sys
from datetime import datetime

from PySide6.QtCore import QSize, Qt
from PySide6.QtPrintSupport import QPrintDialog
from PySide6.QtSql import QSqlDatabase, QSqlQuery, QSqlQueryModel, QSqlTableModel
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QDoubleSpinBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout, QHeaderView,
    QInputDialog, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMenu,
    QMessageBox, QPlainTextEdit, QPushButton, QSpinBox, QSplitter, QStyle, QTabWidget,
    QTableView, QTableWidget, QTableWidgetItem, QTextBrowser, QToolButton, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

# ---------------------------------------------------------------------------
# Constantes y utilidades
# ---------------------------------------------------------------------------
TIPOS = ["TEXTO", "MEMO", "ENTERO", "DECIMAL", "FECHA", "BOOLEANO", "AUTONUMERICO"]
TIPO_SQLITE = {
    "TEXTO": "TEXT", "MEMO": "TEXT", "ENTERO": "INTEGER", "DECIMAL": "REAL",
    "FECHA": "TEXT", "BOOLEANO": "INTEGER", "AUTONUMERICO": "INTEGER",
}
ACCIONES_MACRO = ["Abrir tabla", "Abrir consulta", "Abrir formulario",
                  "Ejecutar SQL", "Mostrar mensaje"]

CSS_INFORME = """
body{font-family:'Segoe UI',Arial,sans-serif;margin:24px;color:#1f2a3d}
h1{font-size:20px;color:#2b579a;border-bottom:2px solid #2b579a;padding-bottom:6px}
p.meta{font-size:11px;color:#6b7a90}
table{border-collapse:collapse;width:100%;font-size:12px;margin-bottom:14px}
th{background:#2b579a;color:#fff;padding:4px 8px;text-align:left}
td{border:1px solid #c9d2e0;padding:3px 8px}
tr:nth-child(even) td{background:#f2f6fc}
tr.tot td{background:#dde7f5;font-weight:bold}
div.grupo{background:#e8eef8;border-left:4px solid #2b579a;padding:5px 8px;
          font-weight:bold;margin:12px 0 4px}
"""


def esc(v):
    return str(v).replace("'", "''")


def _valor_defecto(texto, tipo):
    v = (texto or "").strip()
    if not v:
        return None
    if tipo in ("ENTERO", "BOOLEANO"):
        try:
            return str(int(v))
        except ValueError:
            return None
    if tipo == "DECIMAL":
        try:
            return str(float(v.replace(",", ".")))
        except ValueError:
            return None
    return "'" + esc(v) + "'"


def _inferir_tipo(valores):
    entero = decimal = True
    hay = False
    for v in valores:
        if v is None or v == "":
            continue
        hay = True
        try:
            int(v)
        except ValueError:
            entero = False
            try:
                float(str(v).replace(",", "."))
            except ValueError:
                decimal = False
    if not hay:
        return "TEXT"
    if entero:
        return "INTEGER"
    if decimal:
        return "REAL"
    return "TEXT"


def _convertir(v, tipo):
    if v is None or v == "":
        return None
    if tipo == "INTEGER":
        try:
            return int(v)
        except ValueError:
            return None
    if tipo == "REAL":
        try:
            return float(str(v).replace(",", "."))
        except ValueError:
            return None
    return v


def _col_numerica(filas, i):
    tiene = False
    for f in filas:
        v = f[i]
        if v is None:
            continue
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            return False
        tiene = True
    return tiene


def campos_de_sql(sql):
    q = QSqlQuery()
    q.prepare(f"SELECT * FROM ({sql}) AS _sub LIMIT 0")
    if not q.exec():
        return []
    rec = q.record()
    return [rec.fieldName(i) for i in range(rec.count())]


def html_informe(titulo, cab, filas, agrupar=None):
    idx = cab.index(agrupar) if agrupar and agrupar in cab else None
    num_cols = [i for i in range(len(cab)) if i != idx and _col_numerica(filas, i)]
    p = [f"<html><head><meta charset='utf-8'><style>{CSS_INFORME}</style></head><body>"]
    p.append(f"<h1>{html.escape(titulo)}</h1>")
    p.append(f"<p class='meta'>Generado el {datetime.now():%d/%m/%Y %H:%M} · "
             f"{len(filas)} registros</p>")

    def tabla_html(grupo):
        t = ["<table><tr>" +
             "".join(f"<th>{html.escape(str(c))}</th>" for c in cab) + "</tr>"]
        for f in grupo:
            t.append("<tr>" + "".join(
                f"<td>{html.escape('' if v is None else str(v))}</td>" for v in f)
                + "</tr>")
        if num_cols:
            celdas, etiqueta = [], False
            for i in range(len(cab)):
                if i in num_cols:
                    total = sum(f[i] for f in grupo if isinstance(f[i], (int, float)))
                    celdas.append(f"<td>{total:,.2f}</td>")
                else:
                    celdas.append("<td><b>TOTALES</b></td>" if not etiqueta else "<td></td>")
                    etiqueta = True
            t.append("<tr class='tot'>" + "".join(celdas) + "</tr>")
        t.append("</table>")
        return "".join(t)

    if idx is None:
        p.append(tabla_html(filas))
    else:
        grupos = {}
        for f in filas:
            grupos.setdefault("" if f[idx] is None else str(f[idx]), []).append(f)
        for clave in sorted(grupos):
            p.append(f"<div class='grupo'>{html.escape(clave)} — "
                     f"{len(grupos[clave])} registros</div>")
            p.append(tabla_html(grupos[clave]))
    p.append("<p class='meta'>AccesoLibre 2007 · software libre (GPL-3.0)</p></body></html>")
    return "".join(p)


# ---------------------------------------------------------------------------
# Motor de datos (SQLite)
# ---------------------------------------------------------------------------
class GestorBD:
    def __init__(self, ruta):
        self.ruta = ruta
# Parche de compatibilidad PySide6 / PyQt
        try:
            conn_name = getattr(QSqlDatabase, 'defaultConnection', 'qt_sql_default_connection')
            if QSqlDatabase.contains(conn_name):
                QSqlDatabase.removeDatabase(conn_name)
        except Exception:
            pass
    self.db = QSqlDatabase.addDatabase("QSQLITE")
        self.db.setDatabaseName(ruta)
        if not self.db.open():
            raise RuntimeError(self.db.lastError().text())
        self.ejecutar("CREATE TABLE IF NOT EXISTS _oa_objetos "
                      "(nombre TEXT PRIMARY KEY, tipo TEXT, definicion TEXT)")

    def ejecutar(self, sql, params=None):
        q = QSqlQuery()
        q.prepare(sql)
        if params:
            for p in params:
                q.addBindValue(p)
        if not q.exec():
            raise RuntimeError(q.lastError().text())
        return q

    def tablas(self):
        q = self.ejecutar(
            "SELECT name FROM sqlite_master WHERE type='table' "
            "AND name NOT LIKE '\\_oa%' ESCAPE '\\' "
            "AND name NOT LIKE 'sqlite\\_%' ESCAPE '\\' ORDER BY name")
        out = []
        while q.next():
            out.append(q.value(0))
        return out

    def columnas(self, tabla):
        q = self.ejecutar(f"PRAGMA table_info([{tabla}])")
        out = []
        while q.next():
            out.append({"nombre": q.value(1), "tipo": q.value(2) or "TEXT",
                        "pk": bool(q.value(5))})
        return out

    def guardar_objeto(self, nombre, tipo, definicion):
        self.ejecutar("INSERT OR REPLACE INTO _oa_objetos VALUES (?,?,?)",
                      [nombre, tipo, json.dumps(definicion, ensure_ascii=False)])

    def objetos(self, tipo):
        q = self.ejecutar("SELECT nombre FROM _oa_objetos WHERE tipo=? ORDER BY nombre",
                          [tipo])
        out = []
        while q.next():
            out.append(q.value(0))
        return out

    def obtener_objeto(self, nombre):
        q = self.ejecutar("SELECT tipo, definicion FROM _oa_objetos WHERE nombre=?",
                          [nombre])
        if q.next():
            return q.value(0), json.loads(q.value(1))
        return None, None

    def eliminar_objeto(self, nombre):
        self.ejecutar("DELETE FROM _oa_objetos WHERE nombre=?", [nombre])


# ---------------------------------------------------------------------------
# Diseñadores
# ---------------------------------------------------------------------------
class DisenadorTabla(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Diseño de tabla")
        self.resize(760, 480)
        lay = QVBoxLayout(self)
        fila = QHBoxLayout()
        fila.addWidget(QLabel("Nombre de la tabla:"))
        self.txtNombre = QLineEdit()
        fila.addWidget(self.txtNombre, 1)
        lay.addLayout(fila)
        self.grid = QTableWidget(0, 5)
        self.grid.setHorizontalHeaderLabels(
            ["Nombre del campo", "Tipo de datos", "Clave principal",
             "Requerido", "Valor predeterminado"])
        self.grid.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.grid.verticalHeader().setVisible(False)
        lay.addWidget(self.grid, 1)
        barra = QHBoxLayout()
        b_add = QPushButton("Añadir campo")
        b_add.clicked.connect(self._nueva_fila)
        b_del = QPushButton("Quitar campo")
        b_del.clicked.connect(self._quitar_fila)
        barra.addWidget(b_add)
        barra.addWidget(b_del)
        barra.addStretch(1)
        lay.addLayout(barra)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                              QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Crear tabla")
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self._nueva_fila()

    def _nueva_fila(self):
        r = self.grid.rowCount()
        self.grid.insertRow(r)
        self.grid.setItem(r, 0, QTableWidgetItem(""))
        cmb = QComboBox()
        cmb.addItems(TIPOS)
        self.grid.setCellWidget(r, 1, cmb)
        for c in (2, 3):
            it = QTableWidgetItem()
            it.setCheckState(Qt.CheckState.Unchecked)
            self.grid.setItem(r, c, it)
        self.grid.setItem(r, 4, QTableWidgetItem(""))

    def _quitar_fila(self):
        r = self.grid.currentRow()
        if r >= 0:
            self.grid.removeRow(r)

    def nombre_tabla(self):
        return re.sub(r"\W", "_", self.txtNombre.text().strip())

    def construir_sql(self):
        tabla = self.nombre_tabla()
        defs = []
        for r in range(self.grid.rowCount()):
            it = self.grid.item(r, 0)
            nom = it.text().strip() if it else ""
            if not nom:
                continue
            cmb = self.grid.cellWidget(r, 1)
            tipo = cmb.currentText() if cmb else "TEXTO"
            trozos = [f"[{nom}]", TIPO_SQLITE[tipo]]
            if tipo == "AUTONUMERICO":
                trozos.append("PRIMARY KEY AUTOINCREMENT")
            else:
                if self.grid.item(r, 2).checkState() == Qt.CheckState.Checked:
                    trozos.append("PRIMARY KEY")
                if self.grid.item(r, 3).checkState() == Qt.CheckState.Checked:
                    trozos.append("NOT NULL")
                it_def = self.grid.item(r, 4)
                d = _valor_defecto(it_def.text() if it_def else "", tipo)
                if d:
                    trozos.append(f"DEFAULT {d}")
            defs.append(" ".join(trozos))
        if not tabla or not defs:
            return None
        return f"CREATE TABLE [{tabla}] (\n  " + ",\n  ".join(defs) + "\n)"


class AgregarCampo(QDialog):
    def __init__(self, parent, tabla):
        super().__init__(parent)
        self.setWindowTitle(f"Agregar campo a '{tabla}'")
        lay = QFormLayout(self)
        lay.addRow(QLabel(f"La tabla '{tabla}' ya existe. Solo es posible AGREGAR campos."))
        self.txtNombre = QLineEdit()
        self.cmbTipo = QComboBox()
        self.cmbTipo.addItems(["TEXTO", "MEMO", "ENTERO", "DECIMAL", "FECHA", "BOOLEANO"])
        lay.addRow("Nombre del campo:", self.txtNombre)
        lay.addRow("Tipo de datos:", self.cmbTipo)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                              QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Agregar")
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addRow(bb)

    def resultado(self):
        nom = re.sub(r"\W", "_", self.txtNombre.text().strip())
        return nom, TIPO_SQLITE[self.cmbTipo.currentText()]


class DisenadorConsultas(QDialog):
    def __init__(self, parent, gestor, defn=None):
        super().__init__(parent)
        self.gestor = gestor
        self.setWindowTitle("Diseño de consulta")
        self.resize(820, 620)
        lay = QVBoxLayout(self)
        fila1 = QHBoxLayout()
        fila1.addWidget(QLabel("Nombre:"))
        self.txtNombre = QLineEdit()
        fila1.addWidget(self.txtNombre, 1)
        fila1.addWidget(QLabel("Tabla:"))
        self.cmbTabla = QComboBox()
        self.cmbTabla.addItems(gestor.tablas())
        fila1.addWidget(self.cmbTabla, 1)
        lay.addLayout(fila1)
        fila2 = QHBoxLayout()
        self.lstCampos = QListWidget()
        self.lstCampos.itemChanged.connect(self._actualizar_sql)
        fila2.addWidget(self.lstCampos, 1)
        centro = QVBoxLayout()
        centro.addWidget(QLabel("Criterios de selección (se combinan con AND):"))
        self.tblCriterios = QTableWidget(4, 3)
        self.tblCriterios.setHorizontalHeaderLabels(["Campo", "Condición", "Valor"])
        self.tblCriterios.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tblCriterios.verticalHeader().setVisible(False)
        for r in range(4):
            self.tblCriterios.setCellWidget(r, 0, QComboBox())
            cmb_op = QComboBox()
            cmb_op.addItems(["=", "<>", ">", "<", ">=", "<=", "COMODÍN (LIKE)"])
            cmb_op.currentTextChanged.connect(self._actualizar_sql)
            self.tblCriterios.setCellWidget(r, 1, cmb_op)
            self.tblCriterios.setItem(r, 2, QTableWidgetItem(""))
        self.tblCriterios.itemChanged.connect(self._actualizar_sql)
        centro.addWidget(self.tblCriterios, 1)
        fila_ord = QHBoxLayout()
        fila_ord.addWidget(QLabel("Ordenar por:"))
        self.cmbOrden = QComboBox()
        fila_ord.addWidget(self.cmbOrden, 1)
        self.chkDesc = QCheckBox("Descendente")
        self.chkDesc.stateChanged.connect(self._actualizar_sql)
        fila_ord.addWidget(self.chkDesc)
        centro.addLayout(fila_ord)
        fila2.addLayout(centro, 2)
        lay.addLayout(fila2, 1)
        lay.addWidget(QLabel("SQL generado:"))
        self.txtSQL = QPlainTextEdit()
        self.txtSQL.setReadOnly(True)
        self.txtSQL.setMaximumHeight(64)
        lay.addWidget(self.txtSQL)
        self._modelo_previa = QSqlQueryModel()
        self.tblPrevia = QTableView()
        self.tblPrevia.setModel(self._modelo_previa)
        lay.addWidget(self.tblPrevia, 1)
        barra = QHBoxLayout()
        b_prev = QPushButton("Vista previa")
        b_prev.clicked.connect(self._previsualizar)
        barra.addWidget(b_prev)
        barra.addStretch(1)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                              QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Guardar")
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        bb.accepted.connect(self._aceptar)
        bb.rejected.connect(self.reject)
        barra.addWidget(bb)
        lay.addLayout(barra)
        self.cmbTabla.currentTextChanged.connect(self._cargar_campos)
        if self.cmbTabla.count():
            self._cargar_campos(self.cmbTabla.currentText())
        if defn:
            self.txtNombre.setText(defn.get("nombre", ""))
            if defn.get("tabla"):
                self.cmbTabla.setCurrentText(defn["tabla"])
            for i in range(self.lstCampos.count()):
                self.lstCampos.item(i).setCheckState(Qt.CheckState.Unchecked)
            for c in defn.get("campos", []):
                for i in range(self.lstCampos.count()):
                    if self.lstCampos.item(i).text() == c:
                        self.lstCampos.item(i).setCheckState(Qt.CheckState.Checked)
            for r, cr in enumerate(defn.get("criterios", [])[:4]):
                self.tblCriterios.cellWidget(r, 0).setCurrentText(cr[0] or "")
                self.tblCriterios.cellWidget(r, 1).setCurrentText(cr[1] or "=")
                self.tblCriterios.item(r, 2).setText(cr[2] or "")
            self.cmbOrden.setCurrentText(defn.get("orden") or "")
            self.chkDesc.setChecked(bool(defn.get("desc")))
        self._actualizar_sql()

    def _cargar_campos(self, tabla):
        campos = [c["nombre"] for c in self.gestor.columnas(tabla)] if tabla else []
        self.lstCampos.blockSignals(True)
        self.lstCampos.clear()
        for c in campos:
            it = QListWidgetItem(c)
            it.setCheckState(Qt.CheckState.Checked)
            self.lstCampos.addItem(it)
        self.lstCampos.blockSignals(False)
        for r in range(self.tblCriterios.rowCount()):
            combo = self.tblCriterios.cellWidget(r, 0)
            actual = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems([""] + campos)
            if actual in campos:
                combo.setCurrentText(actual)
            combo.blockSignals(False)
        actual_o = self.cmbOrden.currentText()
        self.cmbOrden.blockSignals(True)
        self.cmbOrden.clear()
        self.cmbOrden.addItems([""] + campos)
        if actual_o in campos:
            self.cmbOrden.setCurrentText(actual_o)
        self.cmbOrden.blockSignals(False)
        self._actualizar_sql()

    def _campos_marcados(self):
        return [self.lstCampos.item(i).text() for i in range(self.lstCampos.count())
                if self.lstCampos.item(i).checkState() == Qt.CheckState.Checked]

    def _criterios(self):
        out = []
        for r in range(self.tblCriterios.rowCount()):
            campo = self.tblCriterios.cellWidget(r, 0).currentText()
            op = self.tblCriterios.cellWidget(r, 1).currentText()
            it = self.tblCriterios.item(r, 2)
            out.append([campo, op, it.text() if it else ""])
        return out

    def generar_sql(self):
        tabla = self.cmbTabla.currentText()
        if not tabla:
            return ""
        marcados = self._campos_marcados()
        campos = ", ".join(f"[{c}]" for c in marcados) if marcados else "*"
        sql = f"SELECT {campos} FROM [{tabla}]"
        conds = []
        for campo, op, valor in self._criterios():
            if not campo or valor == "":
                continue
            if op == "COMODÍN (LIKE)":
                patron = valor if ("%" in valor or "_" in valor) else f"%{valor}%"
                conds.append(f"[{campo}] LIKE '{esc(patron)}'")
            else:
                conds.append(f"[{campo}] {op} '{esc(valor)}'")
        if conds:
            sql += " WHERE " + " AND ".join(conds)
        orden = self.cmbOrden.currentText()
        if orden:
            sql += f" ORDER BY [{orden}] {'DESC' if self.chkDesc.isChecked() else 'ASC'}"
        return sql

    def _actualizar_sql(self, *_):
        self.txtSQL.setPlainText(self.generar_sql())

    def _previsualizar(self):
        sql = self.generar_sql()
        if not sql:
            return
        self._modelo_previa.setQuery(sql)
        if self._modelo_previa.lastError().isValid():
            QMessageBox.warning(self, "Error en la consulta",
                                self._modelo_previa.lastError().text())

    def _aceptar(self):
        if not self.txtNombre.text().strip() or not self.cmbTabla.currentText():
            QMessageBox.warning(self, "Diseño de consulta",
                                "Indique un nombre y seleccione una tabla de origen.")
            return
        self.accept()

    def resultado(self):
        defn = {"nombre": self.txtNombre.text().strip(),
                "tabla": self.cmbTabla.currentText(),
                "campos": self._campos_marcados(),
                "criterios": self._criterios(),
                "orden": self.cmbOrden.currentText(),
                "desc": self.chkDesc.isChecked()}
        defn["sql"] = self.generar_sql()
        return defn


class DisenadorFormulario(QDialog):
    def __init__(self, parent, gestor, defn=None):
        super().__init__(parent)
        self.gestor = gestor
        self.setWindowTitle("Diseño de formulario")
        self.resize(420, 460)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Nombre del formulario:"))
        self.txtNombre = QLineEdit()
        lay.addWidget(self.txtNombre)
        lay.addWidget(QLabel("Tabla de origen:"))
        self.cmbTabla = QComboBox()
        self.cmbTabla.addItems(gestor.tablas())
        self.cmbTabla.currentTextChanged.connect(self._cargar)
        lay.addWidget(self.cmbTabla)
        lay.addWidget(QLabel("Campos incluidos:"))
        self.lstCampos = QListWidget()
        lay.addWidget(self.lstCampos, 1)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                              QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Guardar")
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        bb.accepted.connect(self._aceptar)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        if self.cmbTabla.count():
            self._cargar(self.cmbTabla.currentText())
        if defn:
            self.txtNombre.setText(defn.get("nombre", ""))
            if defn.get("tabla"):
                self.cmbTabla.setCurrentText(defn["tabla"])
            for i in range(self.lstCampos.count()):
                it = self.lstCampos.item(i)
                it.setCheckState(Qt.CheckState.Checked
                                 if it.text() in (defn.get("campos") or [])
                                 else Qt.CheckState.Unchecked)

    def _cargar(self, tabla):
        self.lstCampos.clear()
        for c in self.gestor.columnas(tabla):
            it = QListWidgetItem(c["nombre"])
            it.setCheckState(Qt.CheckState.Checked)
            self.lstCampos.addItem(it)

    def _aceptar(self):
        if not self.txtNombre.text().strip() or not self.cmbTabla.currentText():
            QMessageBox.warning(self, "Formulario", "Indique un nombre y una tabla.")
            return
        self.accept()

    def resultado(self):
        campos = [self.lstCampos.item(i).text() for i in range(self.lstCampos.count())
                  if self.lstCampos.item(i).checkState() == Qt.CheckState.Checked]
        return (self.txtNombre.text().strip(),
                {"tabla": self.cmbTabla.currentText(), "campos": campos})


class DisenadorInforme(QDialog):
    def __init__(self, parent, gestor, defn=None):
        super().__init__(parent)
        self.gestor = gestor
        self.setWindowTitle("Diseño de informe")
        self.resize(420, 260)
        lay = QFormLayout(self)
        self.txtNombre = QLineEdit()
        self.origenes = [("tabla", t) for t in gestor.tablas()] + \
                        [("consulta", c) for c in gestor.objetos("consulta")]
        self.cmbOrigen = QComboBox()
        for tipo, nom in self.origenes:
            self.cmbOrigen.addItem(f"{'Tabla' if tipo == 'tabla' else 'Consulta'}: {nom}")
        self.cmbOrigen.currentIndexChanged.connect(self._cargar_campos)
        self.cmbAgrupar = QComboBox()
        self.txtTitulo = QLineEdit()
        lay.addRow("Nombre del informe:", self.txtNombre)
        lay.addRow("Origen de datos:", self.cmbOrigen)
        lay.addRow("Agrupar por:", self.cmbAgrupar)
        lay.addRow("Título:", self.txtTitulo)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                              QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Guardar")
        bb.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        bb.accepted.connect(self._aceptar)
        bb.rejected.connect(self.reject)
        lay.addRow(bb)
        if self.origenes:
            self._cargar_campos(0)
        if defn:
            self.txtNombre.setText(defn.get("nombre", ""))
            self.txtTitulo.setText(defn.get("titulo", ""))
            for i, (tipo, nom) in enumerate(self.origenes):
                if tipo == defn.get("origen_tipo") and nom == defn.get("origen"):
                    self.cmbOrigen.setCurrentIndex(i)
                    break
            self._cargar_campos(self.cmbOrigen.currentIndex())
            if defn.get("agrupar"):
                self.cmbAgrupar.setCurrentText(defn["agrupar"])

    def _cargar_campos(self, idx):
        self.cmbAgrupar.clear()
        self.cmbAgrupar.addItem("(sin agrupar)")
        if idx < 0 or idx >= len(self.origenes):
            return
        tipo, nombre = self.origenes[idx]
        if tipo == "tabla":
            campos = [c["nombre"] for c in self.gestor.columnas(nombre)]
        else:
            _, d = self.gestor.obtener_objeto(nombre)
            campos = campos_de_sql(d["sql"]) if d else []
        self.cmbAgrupar.addItems(campos)

    def _aceptar(self):
        if not self.txtNombre.text().strip() or not self.origenes:
            QMessageBox.warning(self, "Informe",
                                "Indique un nombre y asegúrese de tener un origen de datos.")
            return
        self.accept()

    def resultado(self):
        tipo, origen = self.origenes[self.cmbOrigen.currentIndex()]
        agrupar = self.cmbAgrupar.currentText()
        return (self.txtNombre.text().strip(),
                {"origen_tipo": tipo, "origen": origen,
                 "agrupar": None if agrupar == "(sin agrupar)" else agrupar,
                 "titulo": self.txtTitulo.text().strip() or self.txtNombre.text().strip()})


# ---------------------------------------------------------------------------
# Vistas
# ---------------------------------------------------------------------------
class VistaDatos(QWidget):
    def __init__(self, gestor, tabla=None, sql=None):
        super().__init__()
        lay = QVBoxLayout(self)
        barra = QHBoxLayout()
        if sql is not None:
            self.modelo = QSqlQueryModel()
            self.modelo.setQuery(sql)
        else:
            self.modelo = QSqlTableModel()
            self.modelo.setTable(tabla)
            self.modelo.setEditStrategy(QSqlTableModel.EditStrategy.OnManualSubmit)
            self.modelo.select()
            btn = QPushButton("Guardar cambios")
            btn.clicked.connect(self.guardar)
            barra.addWidget(btn)
            btn = QPushButton("Deshacer")
            btn.clicked.connect(self.modelo.revertAll)
            barra.addWidget(btn)
            btn = QPushButton("Nuevo registro")
            btn.clicked.connect(self.nuevo_registro)
            barra.addWidget(btn)
            btn = QPushButton("Eliminar fila")
            btn.clicked.connect(self.eliminar_fila)
            barra.addWidget(btn)
            self.txtFiltro = QLineEdit()
            self.txtFiltro.setPlaceholderText("Filtro, p. ej.: Ciudad='Madrid'")
            barra.addWidget(self.txtFiltro, 1)
            btn = QPushButton("Filtrar")
            btn.clicked.connect(self.filtrar)
            barra.addWidget(btn)
            btn = QPushButton("Quitar filtro")
            btn.clicked.connect(self.quitar_filtro)
            barra.addWidget(btn)
        self.vista = QTableView()
        self.vista.setModel(self.modelo)
        self.vista.setAlternatingRowColors(True)
        self.vista.resizeColumnsToContents()
        lay.addLayout(barra)
        lay.addWidget(self.vista, 1)

    def guardar(self):
        if self.modelo.submitAll():
            QMessageBox.information(self, "Hoja de datos", "Cambios guardados.")
        else:
            QMessageBox.warning(self, "Error al guardar", self.modelo.lastError().text())

    def nuevo_registro(self):
        self.modelo.insertRow(self.modelo.rowCount())
        self.vista.scrollToBottom()

    def eliminar_fila(self):
        idx = self.vista.currentIndex()
        if idx.isValid():
            self.modelo.removeRow(idx.row())

    def filtrar(self):
        self.modelo.setFilter(self.txtFiltro.text().strip())
        self.modelo.select()

    def quitar_filtro(self):
        self.txtFiltro.clear()
        self.modelo.setFilter("")
        self.modelo.select()


class Formulario(QWidget):
    def __init__(self, gestor, defn):
        super().__init__()
        from PySide6.QtSql import QSqlTableModel, QDataWidgetMapper
        tabla = defn["tabla"]
        sel = set(defn.get("campos") or [])
        self.modelo = QSqlTableModel(self)
        self.modelo.setTable(tabla)
        self.modelo.setEditStrategy(QSqlTableModel.EditStrategy.OnManualSubmit)
        self.modelo.select()
        self.mapper = QDataWidgetMapper(self)
        self.mapper.setModel(self.modelo)
        self.mapper.setSubmitPolicy(QDataWidgetMapper.SubmitPolicy.ManualSubmit)
        lay = QVBoxLayout(self)
        cab = QLabel(f"Formulario: {tabla}")
        cab.setStyleSheet("font-weight:bold;font-size:14px;color:#2b579a")
        lay.addWidget(cab)
        form = QFormLayout()
        for i, col in enumerate(gestor.columnas(tabla)):
            if sel and col["nombre"] not in sel:
                continue
            if col["pk"] and col["tipo"].upper().startswith("INT"):
                form.addRow(col["nombre"], QLabel("(autonumérico)"))
                continue
            if col["tipo"].upper().startswith("INT"):
                w = QSpinBox()
                w.setRange(-2 ** 31, 2 ** 31 - 1)
            elif col["tipo"].upper().startswith("REAL"):
                w = QDoubleSpinBox()
                w.setRange(-1e12, 1e12)
                w.setDecimals(4)
            else:
                w = QLineEdit()
            form.addRow(col["nombre"], w)
            self.mapper.addMapping(w, i)
        lay.addLayout(form)
        nav = QHBoxLayout()
        for texto, slot in (("|<", self.mapper.toFirst), ("<", self.mapper.toPrevious),
                            (">", self.mapper.toNext), (">|", self.mapper.toLast)):
            b = QPushButton(texto)
            b.clicked.connect(slot)
            nav.addWidget(b)
        b = QPushButton("Nuevo")
        b.clicked.connect(self.nuevo)
        nav.addWidget(b)
        b = QPushButton("Guardar")
        b.clicked.connect(self.guardar)
        nav.addWidget(b)
        b = QPushButton("Eliminar")
        b.clicked.connect(self.eliminar)
        nav.addWidget(b)
        nav.addStretch(1)
        self.lblEstado = QLabel()
        nav.addWidget(self.lblEstado)
        lay.addLayout(nav)
        lay.addStretch(1)
        self.mapper.currentChanged.connect(self._actualizar_estado)
        self.mapper.toFirst()
        self._actualizar_estado()

    def _actualizar_estado(self, *_):
        total = self.modelo.rowCount()
        actual = self.mapper.currentIndex() + 1
        self.lblEstado.setText(f"Registro {actual} de {total}" if total else "Sin registros")

    def nuevo(self):
        fila = self.modelo.rowCount()
        self.modelo.insertRow(fila)
        self.mapper.setCurrentIndex(fila)

    def guardar(self):
        if self.mapper.submit():
            self.modelo.select()
            self._actualizar_estado()
        else:
            QMessageBox.warning(self, "Error al guardar", self.modelo.lastError().text())

    def eliminar(self):
        if self.mapper.currentIndex() < 0:
            return
        if QMessageBox.question(self, "Eliminar registro",
                                "¿Eliminar el registro actual?") != QMessageBox.StandardButton.Yes:
            return
        self.modelo.removeRow(self.mapper.currentIndex())
        self.modelo.submitAll()
        self.modelo.select()
        self.mapper.toFirst()


class VisorInforme(QWidget):
    def __init__(self, documento_html):
        super().__init__()
        lay = QVBoxLayout(self)
        barra = QHBoxLayout()
        b = QPushButton("Imprimir…")
        b.clicked.connect(self.imprimir)
        barra.addWidget(b)
        barra.addStretch(1)
        self.visor = QTextBrowser()
        self.visor.setHtml(documento_html)
        lay.addLayout(barra)
        lay.addWidget(self.visor, 1)

    def imprimir(self):
        dlg = QPrintDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.visor.print_(dlg.printer())


class EditorMacro(QWidget):
    def __init__(self, principal, nombre=None, pasos=None):
        super().__init__()
        self.principal = principal
        lay = QVBoxLayout(self)
        fila = QHBoxLayout()
        fila.addWidget(QLabel("Nombre de la macro:"))
        self.txtNombre = QLineEdit(nombre or "")
        fila.addWidget(self.txtNombre, 1)
        lay.addLayout(fila)
        self.grid = QTableWidget(0, 2)
        self.grid.setHorizontalHeaderLabels(["Acción", "Parámetro"])
        self.grid.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.grid.verticalHeader().setVisible(False)
        lay.addWidget(self.grid, 1)
        for p in (pasos or []):
            self._agregar(p)
        barra = QHBoxLayout()
        b = QPushButton("Añadir acción")
        b.clicked.connect(self._agregar)
        barra.addWidget(b)
        b = QPushButton("Quitar seleccionada")
        b.clicked.connect(self._quitar)
        barra.addWidget(b)
        barra.addStretch(1)
        b = QPushButton("Ejecutar ahora")
        b.clicked.connect(self.ejecutar)
        barra.addWidget(b)
        b = QPushButton("Guardar macro")
        b.clicked.connect(self.guardar)
        barra.addWidget(b)
        lay.addLayout(barra)

    def _agregar(self, paso=None):
        r = self.grid.rowCount()
        self.grid.insertRow(r)
        cmb = QComboBox()
        cmb.addItems(ACCIONES_MACRO)
        if paso and paso[0] in ACCIONES_MACRO:
            cmb.setCurrentText(paso[0])
        self.grid.setCellWidget(r, 0, cmb)
        self.grid.setItem(r, 1, QTableWidgetItem(paso[1] if paso else ""))

    def _quitar(self):
        r = self.grid.currentRow()
        if r >= 0:
            self.grid.removeRow(r)

    def pasos(self):
        out = []
        for r in range(self.grid.rowCount()):
            cmb = self.grid.cellWidget(r, 0)
            it = self.grid.item(r, 1)
            if cmb:
                out.append([cmb.currentText(), it.text() if it else ""])
        return out

    def guardar(self):
        nombre = self.txtNombre.text().strip()
        if not nombre:
            QMessageBox.warning(self, "Macro", "Indique un nombre para la macro.")
            return
        self.principal.gestor.guardar_objeto(nombre, "macro", {"pasos": self.pasos()})
        self.principal.refrescar_navegacion()
        self.principal.statusBar().showMessage(f"Macro '{nombre}' guardada.")

    def ejecutar(self):
        self.principal.ejecutar_pasos(self.pasos())


class ConsolaSQL(QDialog):
    def __init__(self, parent, gestor):
        super().__init__(parent)
        self.gestor = gestor
        self.setWindowTitle("Ejecutar SQL directo")
        self.resize(720, 480)
        lay = QVBoxLayout(self)
        self.editor = QPlainTextEdit()
        self.editor.setPlaceholderText("SELECT * FROM Clientes WHERE Ciudad='Madrid';")
        lay.addWidget(self.editor, 1)
        barra = QHBoxLayout()
        b = QPushButton("Ejecutar  (también INSERT/UPDATE/DELETE/DDL)")
        b.clicked.connect(self.ejecutar)
        barra.addWidget(b)
        self.lblEstado = QLabel("")
        barra.addWidget(self.lblEstado, 1)
        lay.addLayout(barra)
        self.modelo = QSqlQueryModel()
        self.vista = QTableView()
        self.vista.setModel(self.modelo)
        lay.addWidget(self.vista, 2)

    def ejecutar(self):
        sql = self.editor.toPlainText().strip()
        if not sql:
            return
        self.modelo.setQuery(sql)
        if self.modelo.lastError().isValid():
            self.lblEstado.setText("Error: " + self.modelo.lastError().text())
        else:
            self.lblEstado.setText(
                f"Correcto — {self.modelo.query().numRowsAffected()} filas afectadas, "
                f"{self.modelo.rowCount()} filas devueltas.")


# ---------------------------------------------------------------------------
# Ventana principal
# ---------------------------------------------------------------------------
class VentanaPrincipal(QMainWindow):
    SECCIONES = [("Tablas", "tabla"), ("Consultas", "consulta"),
                 ("Formularios", "formulario"), ("Informes", "informe"),
                 ("Macros", "macro")]

    def __init__(self):
        super().__init__()
        self.gestor = None
        self.setWindowTitle("AccesoLibre 2007 — suite de bases de datos de código abierto")
        self.resize(1180, 740)
        central = QWidget()
        self.setCentralWidget(central)
        raiz = QVBoxLayout(central)
        raiz.setContentsMargins(4, 4, 4, 4)
        raiz.setSpacing(3)
        raiz.addWidget(self._construir_cinta())
        divisor = QSplitter()
        self.arbol = QTreeWidget()
        self.arbol.setHeaderHidden(True)
        self.arbol.setMinimumWidth(220)
        self.arbol.itemDoubleClicked.connect(self._abrir_seleccion)
        self.arbol.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.arbol.customContextMenuRequested.connect(self._menu_contextual)
        self.espacio = QTabWidget()
        self.espacio.setTabsClosable(True)
        self.espacio.tabCloseRequested.connect(self._cerrar_pestana)
        divisor.addWidget(self.arbol)
        divisor.addWidget(self.espacio)
        divisor.setSizes([240, 940])
        raiz.addWidget(divisor, 1)
        self.statusBar().showMessage("Cree o abra una base de datos para comenzar.")

    def _boton(self, texto, icono, slot):
        b = QToolButton()
        b.setText(texto)
        b.setToolTip(texto)
        b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        b.setIcon(self.style().standardIcon(icono))
        b.setIconSize(QSize(22, 22))
        b.clicked.connect(slot)
        return b

    def _grupo(self, titulo, botones):
        g = QGroupBox(titulo)
        h = QHBoxLayout(g)
        for b in botones:
            h.addWidget(b)
        h.addStretch(1)
        g.setFixedHeight(88)
        return g

    def _construir_cinta(self):
        SP = QStyle.StandardPixmap
        cinta = QTabWidget()
        cinta.setFixedHeight(124)
        ini = QWidget()
        h = QHBoxLayout(ini)
        h.addWidget(self._grupo("Archivo", [
            self._boton("Nueva BD", SP.SP_FileDialogNewFolder, self.nueva_bd),
            self._boton("Abrir BD", SP.SP_DialogOpenButton, self.abrir_bd),
            self._boton("Ejemplo", SP.SP_DialogHelpButton, self.abrir_ejemplo)]))
        h.addWidget(self._grupo("Vistas", [
            self._boton("Abrir objeto", SP.SP_FileDialogContentsView, self.abrir_seleccion),
            self._boton("Vista diseño", SP.SP_FileDialogDetailedView, self.diseno_seleccion)]))
        h.addWidget(self._grupo("Objetos", [
            self._boton("Eliminar", SP.SP_TrashIcon, self.eliminar_seleccion),
            self._boton("Actualizar", SP.SP_BrowserReload, self.refrescar_navegacion)]))
        h.addStretch(1)
        cinta.addTab(ini, "Inicio")
        crear = QWidget()
        h = QHBoxLayout(crear)
        h.addWidget(self._grupo("Tablas",
                                [self._boton("Tabla", SP.SP_FileDialogListView, self.crear_tabla)]))
        h.addWidget(self._grupo("Consultas",
                                [self._boton("Diseño de consulta", SP.SP_FileDialogDetailedView,
                                             self.crear_consulta)]))
        h.addWidget(self._grupo("Formularios",
                                [self._boton("Formulario", SP.SP_FileDialogStart,
                                             self.crear_formulario)]))
        h.addWidget(self._grupo("Informes",
                                [self._boton("Informe", SP.SP_FileDialogInfoView,
                                             self.crear_informe)]))
        h.addWidget(self._grupo("Macros",
                                [self._boton("Macro", SP.SP_MediaPlay, self.crear_macro)]))
        h.addStretch(1)
        cinta.addTab(crear, "Crear")
        ext = QWidget()
        h = QHBoxLayout(ext)
        h.addWidget(self._grupo("Acceso NATIVO", [
            self._boton("Abrir ACCDB directo", SP.SP_DirOpenIcon,
                        self.abrir_accdb_directo)]))
        h.addWidget(self._grupo("Importar", [
            self._boton("Importar CSV", SP.SP_DirOpenIcon, self.importar_csv),
            self._boton("Importar ACCDB/MDB", SP.SP_DialogOpenButton, self.importar_accdb)]))
        h.addWidget(self._grupo("Exportar", [
            self._boton("Exportar CSV", SP.SP_DirIcon, self.exportar_csv),
            self._boton("Exportar a ACCDB", SP.SP_DialogSaveButton, self.exportar_accdb)]))
        h.addStretch(1)
        cinta.addTab(ext, "Datos externos")
        her = QWidget()
        h = QHBoxLayout(her)
        h.addWidget(self._grupo("SQL",
                                [self._boton("Ejecutar SQL", SP.SP_MediaPlay,
                                             self.consola_sql)]))
        h.addWidget(self._grupo("Mantenimiento", [
            self._boton("Compactar", SP.SP_DialogResetButton, self.compactar),
            self._boton("Copia de seguridad", SP.SP_DialogSaveButton,
                        self.copia_seguridad)]))
        h.addStretch(1)
        cinta.addTab(her, "Herramientas de BD")
        return cinta

    def _conectar(self, ruta):
        self.espacio.clear()
        try:
            self.gestor = GestorBD(ruta)
        except Exception as e:
            self.gestor = None
            QMessageBox.critical(self, "No se pudo abrir la base de datos", str(e))
            return
        self.setWindowTitle(f"AccesoLibre 2007 — {os.path.basename(ruta)}")
        self.refrescar_navegacion()
        self.statusBar().showMessage(f"Conectado a {ruta}")

    def nueva_bd(self):
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Nueva base de datos", "basedatos.oal",
            "Base de datos AccesoLibre (*.oal);;SQLite (*.db *.sqlite)")
        if ruta:
            self._conectar(ruta)

    def abrir_bd(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Abrir base de datos", "",
            "Bases de datos (*.oal *.db *.sqlite *.sqlite3);;Todos los archivos (*)")
        if ruta:
            self._conectar(ruta)

    def _sin_bd(self):
        QMessageBox.information(self, "AccesoLibre",
                                "Primero cree o abra una base de datos (Inicio → Archivo).")

    def refrescar_navegacion(self):
        self.arbol.clear()
        if not self.gestor:
            return
        for titulo, tipo in self.SECCIONES:
            raiz = QTreeWidgetItem([titulo])
            raiz.setExpanded(True)
            nombres = (self.gestor.tablas() if tipo == "tabla"
                       else self.gestor.objetos(tipo))
            for n in nombres:
                hijo = QTreeWidgetItem([n])
                hijo.setData(0, Qt.ItemDataRole.UserRole, (tipo, n))
                raiz.addChild(hijo)
            self.arbol.addTopLevelItem(raiz)

    def _seleccion(self):
        if not self.gestor:
            self._sin_bd()
            return None
        item = self.arbol.currentItem()
        datos = item.data(0, Qt.ItemDataRole.UserRole) if item else None
        if not datos:
            self.statusBar().showMessage("Seleccione un objeto en el panel de navegación.")
            return None
        return datos

    def _cerrar_pestana(self, indice):
        w = self.espacio.widget(indice)
        if hasattr(w, "cerrar") and w.cerrar() is False:
            return
        self.espacio.removeTab(indice)
        w.deleteLater()

    def _abrir_seleccion(self, item, *_):
        datos = item.data(0, Qt.ItemDataRole.UserRole) if item else None
        if not datos:
            return
        tipo, nombre = datos
        try:
            if tipo == "tabla":
                self.abrir_tabla(nombre)
            elif tipo == "consulta":
                self.ejecutar_consulta(nombre)
            elif tipo == "formulario":
                self.abrir_formulario(nombre)
            elif tipo == "informe":
                self.abrir_informe(nombre)
            elif tipo == "macro":
                self.editar_macro(nombre)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def abrir_seleccion(self):
        item = self.arbol.currentItem()
        if item:
            self._abrir_seleccion(item)

    def _menu_contextual(self, pos):
        item = self.arbol.itemAt(pos)
        if not item or not item.data(0, Qt.ItemDataRole.UserRole):
            return
        menu = QMenu(self)
        menu.addAction("Abrir", lambda: self._abrir_seleccion(item))
        menu.addAction("Vista diseño", self.diseno_seleccion)
        menu.addSeparator()
        menu.addAction("Eliminar", self.eliminar_seleccion)
        menu.exec(self.arbol.mapToGlobal(pos))

    def _pestana(self, widget, titulo):
        self.espacio.addTab(widget, titulo)
        self.espacio.setCurrentWidget(widget)

    def abrir_tabla(self, nombre):
        self._pestana(VistaDatos(self.gestor, tabla=nombre), f"Tabla: {nombre}")

    def crear_tabla(self):
        if not self.gestor:
            return self._sin_bd()
        dlg = DisenadorTabla(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        sql = dlg.construir_sql()
        if not sql:
            QMessageBox.warning(self, "Diseño incompleto",
                                "Indique un nombre de tabla y al menos un campo.")
            return
        try:
            self.gestor.ejecutar(sql)
        except Exception as e:
            QMessageBox.critical(self, "Error SQL", str(e))
            return
        nombre = dlg.nombre_tabla()
        self.refrescar_navegacion()
        self.statusBar().showMessage(f"Tabla '{nombre}' creada.")
        self.abrir_tabla(nombre)

    def crear_consulta(self):
        self._editar_consulta(None)

    def _editar_consulta(self, defn):
        if not self.gestor:
            return self._sin_bd()
        dlg = DisenadorConsultas(self, self.gestor, defn)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        d = dlg.resultado()
        self.gestor.guardar_objeto(d["nombre"], "consulta", d)
        self.refrescar_navegacion()
        self.statusBar().showMessage(f"Consulta '{d['nombre']}' guardada.")

    def ejecutar_consulta(self, nombre):
        _, defn = self.gestor.obtener_objeto(nombre)
        self._pestana(VistaDatos(self.gestor, sql=defn["sql"]), f"Consulta: {nombre}")

    def crear_formulario(self):
        self._editar_formulario(None)

    def _editar_formulario(self, defn):
        if not self.gestor:
            return self._sin_bd()
        dlg = DisenadorFormulario(self, self.gestor, defn)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        nombre, d = dlg.resultado()
        self.gestor.guardar_objeto(nombre, "formulario", d)
        self.refrescar_navegacion()
        self.abrir_formulario(nombre)

    def abrir_formulario(self, nombre):
        _, defn = self.gestor.obtener_objeto(nombre)
        self._pestana(Formulario(self.gestor, defn), f"Formulario: {nombre}")

    def crear_informe(self):
        self._editar_informe(None)

    def _editar_informe(self, defn):
        if not self.gestor:
            return self._sin_bd()
        dlg = DisenadorInforme(self, self.gestor, defn)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        nombre, d = dlg.resultado()
        self.gestor.guardar_objeto(nombre, "informe", d)
        self.refrescar_navegacion()
        self.abrir_informe(nombre)

    def abrir_informe(self, nombre):
        _, d = self.gestor.obtener_objeto(nombre)
        if d["origen_tipo"] == "tabla":
            sql = f"SELECT * FROM [{d['origen']}]"
        else:
            _, dq = self.gestor.obtener_objeto(d["origen"])
            sql = dq["sql"]
        q = self.gestor.ejecutar(sql)
        rec = q.record()
        cab = [rec.fieldName(i) for i in range(rec.count())]
        filas = []
        while q.next():
            filas.append([q.value(i) for i in range(rec.count())])
        doc = html_informe(d.get("titulo") or nombre, cab, filas, d.get("agrupar"))
        self._pestana(VisorInforme(doc), f"Informe: {nombre}")

    def crear_macro(self):
        if not self.gestor:
            return self._sin_bd()
        self._pestana(EditorMacro(self), "Macro nueva")

    def editar_macro(self, nombre):
        _, defn = self.gestor.obtener_objeto(nombre)
        self._pestana(EditorMacro(self, nombre, defn.get("pasos", [])),
                      f"Macro: {nombre}")

    def ejecutar_pasos(self, pasos):
        try:
            for accion, p1 in pasos:
                if accion == "Abrir tabla":
                    self.abrir_tabla(p1)
                elif accion == "Abrir consulta":
                    self.ejecutar_consulta(p1)
                elif accion == "Abrir formulario":
                    self.abrir_formulario(p1)
                elif accion == "Ejecutar SQL":
                    self.gestor.ejecutar(p1)
                elif accion == "Mostrar mensaje":
                    QMessageBox.information(self, "Macro", p1)
        except Exception as e:
            QMessageBox.critical(self, "Error en la macro", str(e))

    def diseno_seleccion(self):
        datos = self._seleccion()
        if not datos:
            return
        tipo, nombre = datos
        try:
            if tipo == "tabla":
                self._agregar_campo(nombre)
            elif tipo == "consulta":
                _, defn = self.gestor.obtener_objeto(nombre)
                self._editar_consulta(defn)
            elif tipo == "formulario":
                _, defn = self.gestor.obtener_objeto(nombre)
                defn["nombre"] = nombre
                self._editar_formulario(defn)
            elif tipo == "informe":
                _, defn = self.gestor.obtener_objeto(nombre)
                defn["nombre"] = nombre
                self._editar_informe(defn)
            elif tipo == "macro":
                self.editar_macro(nombre)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _agregar_campo(self, tabla):
        dlg = AgregarCampo(self, tabla)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        nom, tipo = dlg.resultado()
        if not nom:
            return
        self.gestor.ejecutar(f"ALTER TABLE [{tabla}] ADD COLUMN [{nom}] {tipo}")
        self.statusBar().showMessage(f"Campo '{nom}' agregado a '{tabla}'.")

    def eliminar_seleccion(self):
        datos = self._seleccion()
        if not datos:
            return
        tipo, nombre = datos
        aviso = (f"la tabla '{nombre}' y TODOS sus datos" if tipo == "tabla"
                 else f"el objeto '{nombre}'")
        if QMessageBox.question(self, "Eliminar",
                                f"¿Está seguro de eliminar {aviso}?") != QMessageBox.StandardButton.Yes:
            return
        try:
            if tipo == "tabla":
                self.gestor.ejecutar(f"DROP TABLE [{nombre}]")
            else:
                self.gestor.eliminar_objeto(nombre)
            self.refrescar_navegacion()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def importar_csv(self):
        if not self.gestor:
            return self._sin_bd()
        ruta, _ = QFileDialog.getOpenFileName(self, "Importar CSV", "",
                                              "Archivos CSV (*.csv *.txt)")
        if not ruta:
            return
        try:
            with open(ruta, newline="", encoding="utf-8-sig") as f:
                filas = [fila for fila in csv.reader(f) if fila]
            if len(filas) < 2:
                raise ValueError("El CSV necesita cabecera y al menos una fila de datos.")
            cab = [re.sub(r"\W", "_", c.strip() or f"col{i}")
                   for i, c in enumerate(filas[0])]
            datos = filas[1:]
            tipos = [_inferir_tipo([f[i] for f in datos if i < len(f)])
                     for i in range(len(cab))]
            nombre = re.sub(r"\W", "_",
                            os.path.splitext(os.path.basename(ruta))[0]) or "importado"
            existentes = set(self.gestor.tablas())
            base, k = nombre, 1
            while nombre in existentes:
                nombre = f"{base}_{k}"
                k += 1
            cols_sql = ", ".join(f"[{c}] {t}" for c, t in zip(cab, tipos))
            self.gestor.ejecutar(f"CREATE TABLE [{nombre}] ({cols_sql})")
            marcas = ",".join("?" * len(cab))
            for f in datos:
                f = (f + [""] * len(cab))[:len(cab)]
                self.gestor.ejecutar(f"INSERT INTO [{nombre}] VALUES ({marcas})",
                                     [_convertir(v, t) for v, t in zip(f, tipos)])
            self.refrescar_navegacion()
            QMessageBox.information(self, "Importar",
                                    f"Se importaron {len(datos)} filas a la tabla '{nombre}'.")
        except Exception as e:
            QMessageBox.critical(self, "Error al importar", str(e))

    def exportar_csv(self):
        if not self.gestor:
            return self._sin_bd()
        tablas = self.gestor.tablas()
        if not tablas:
            return
        tabla, ok = QInputDialog.getItem(self, "Exportar a CSV", "Tabla:", tablas, 0, False)
        if not ok:
            return
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar CSV", f"{tabla}.csv",
                                              "CSV (*.csv)")
        if not ruta:
            return
        q = self.gestor.ejecutar(f"SELECT * FROM [{tabla}]")
        cols = [c["nombre"] for c in self.gestor.columnas(tabla)]
        with open(ruta, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(cols)
            while q.next():
                w.writerow([q.value(i) for i in range(len(cols))])
        QMessageBox.information(self, "Exportar", f"Tabla '{tabla}' exportada a {ruta}.")

    def importar_accdb(self):
        if not self.gestor:
            return self._sin_bd()
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Importar base de datos Access", "",
            "Bases de datos Access (*.accdb *.mdb)")
        if not ruta:
            return
        try:
            from puente_access import importar_access
        except ImportError:
            QMessageBox.critical(
                self, "Falta el puente Access",
                "Instale:  pip install JPype1\n"
                "y un JRE:  sudo apt install default-jre-headless")
            return
        self.statusBar().showMessage("Importando ACCDB/MDB (puede tardar)…")
        try:
            resumen = importar_access(ruta, self.gestor)
            self.refrescar_navegacion()
            QMessageBox.information(self, "Importación Access", resumen)
        except Exception as e:
            QMessageBox.critical(self, "Error al importar Access", str(e))

    def exportar_accdb(self):
        if not self.gestor:
            return self._sin_bd()
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Exportar a ACCDB", "datos.accdb",
            "Base de datos Access 2007+ (*.accdb)")
        if not ruta:
            return
        if not ruta.lower().endswith(".accdb"):
            ruta += ".accdb"
        try:
            from puente_access import exportar_access
            resumen = exportar_access(ruta, self.gestor)
            QMessageBox.information(self, "Exportación Access", resumen)
        except Exception as e:
            QMessageBox.critical(self, "Error al exportar Access", str(e))

    def abrir_accdb_directo(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Abrir y editar NATIVamente un archivo Access", "",
            "Bases de datos Access (*.accdb *.mdb)")
        if not ruta:
            return
        try:
            from modo_directo import VistaDirecto
            ventana = VistaDirecto(ruta)
        except Exception as e:
            QMessageBox.critical(self, "No se pudo abrir en modo nativo", str(e))
            return
        self._pestana(ventana, f"ACCDB: {os.path.basename(ruta)}")
        self.statusBar().showMessage(
            "Modo nativo: los cambios se guardan directamente en el .accdb.")

    def consola_sql(self):
        if not self.gestor:
            return self._sin_bd()
        ConsolaSQL(self, self.gestor).exec()

    def compactar(self):
        if not self.gestor:
            return self._sin_bd()
        self.gestor.ejecutar("VACUUM")
        QMessageBox.information(self, "Compactar", "Base de datos compactada correctamente.")

    def copia_seguridad(self):
        if not self.gestor:
            return self._sin_bd()
        ruta, _ = QFileDialog.getSaveFileName(self, "Copia de seguridad", "copia.oal",
                                              "Base de datos (*.oal *.db)")
        if not ruta:
            return
        self.gestor.ejecutar(f"VACUUM INTO '{esc(ruta)}'")
        QMessageBox.information(self, "Copia de seguridad", f"Copia creada en {ruta}.")

    def abrir_ejemplo(self):
        ruta = os.path.join(os.path.expanduser("~"), "AccesoLibre_Ejemplo.oal")
        self._conectar(ruta)
        if not self.gestor or "Clientes" in self.gestor.tablas():
            return
        try:
            self.gestor.ejecutar(
                "CREATE TABLE Clientes ("
                "IdCliente INTEGER PRIMARY KEY AUTOINCREMENT, "
                "Nombre TEXT NOT NULL, Ciudad TEXT, "
                "LimiteCredito REAL DEFAULT 0, FechaAlta TEXT)")
            for fila in [("Ana Torres", "Madrid", 1500.0, "2026-01-12"),
                         ("Luis Pérez", "Sevilla", 2300.5, "2026-02-03"),
                         ("Marta Díaz", "Madrid", 900.0, "2026-02-19"),
                         ("Carlos Ruiz", "Valencia", 3100.0, "2026-03-08"),
                         ("Elena Gil", "Sevilla", 1750.25, "2026-04-22"),
                         ("Jorge Navarro", "Valencia", 640.0, "2026-05-30"),
                         ("Sofía Ramos", "Madrid", 4200.0, "2026-07-11"),
                         ("Pablo Sanz", "Bilbao", 1200.0, "2026-09-01")]:
                self.gestor.ejecutar(
                    "INSERT INTO Clientes (Nombre, Ciudad, LimiteCredito, FechaAlta) "
                    "VALUES (?,?,?,?)", list(fila))
            self.gestor.guardar_objeto("Clientes por ciudad", "consulta", {
                "nombre": "Clientes por ciudad", "tabla": "Clientes",
                "campos": ["Nombre", "Ciudad", "LimiteCredito"], "criterios": [],
                "orden": "Ciudad", "desc": False,
                "sql": "SELECT [Nombre], [Ciudad], [LimiteCredito] FROM [Clientes] "
                       "ORDER BY [Ciudad] ASC"})
            self.gestor.guardar_objeto("Ficha de cliente", "formulario", {
                "tabla": "Clientes",
                "campos": ["Nombre", "Ciudad", "LimiteCredito", "FechaAlta"]})
            self.gestor.guardar_objeto("Listado de clientes", "informe", {
                "origen_tipo": "tabla", "origen": "Clientes", "agrupar": "Ciudad",
                "titulo": "Listado de clientes por ciudad"})
            self.gestor.guardar_objeto("Macro de bienvenida", "macro", {
                "pasos": [["Mostrar mensaje", "¡Bienvenido a la base de datos de ejemplo!"],
                          ["Abrir tabla", "Clientes"]]})
            self.refrescar_navegacion()
            self.statusBar().showMessage("Base de datos de ejemplo creada.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def closeEvent(self, evento):
        for i in reversed(range(self.espacio.count())):
            w = self.espacio.widget(i)
            if hasattr(w, "cerrar") and w.cerrar() is False:
                evento.ignore()
                return
        evento.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("AccesoLibre 2007")
    app.setStyleSheet(
        "QGroupBox {font-size:11px; border:1px solid #c5d3e8; border-radius:4px;"
        " margin-top:10px}"
        "QGroupBox::title {subcontrol-origin: margin; left:6px; padding:0 3px;"
        " color:#2b579a}")
    v = VentanaPrincipal()
    v.show()
    mb = QMessageBox(v)
    mb.setWindowTitle("Bienvenido a AccesoLibre 2007")
    mb.setText("AccesoLibre 2007 — suite de bases de datos relacional de código abierto.\n"
               "Incluye tablas, consultas, formularios, informes y macros sobre SQLite.\n"
               "Compatible con archivos ACCDB/MDB de Microsoft Access.")
    b_nueva = mb.addButton("Crear base de datos nueva…", QMessageBox.ButtonRole.AcceptRole)
    b_abrir = mb.addButton("Abrir base de datos existente…", QMessageBox.ButtonRole.ActionRole)
    b_ejemplo = mb.addButton("Abrir base de ejemplo", QMessageBox.ButtonRole.ActionRole)
    mb.addButton("Salir", QMessageBox.ButtonRole.RejectRole)
    mb.exec()
    pulsado = mb.clickedButton()
    if pulsado == b_nueva:
        v.nueva_bd()
    elif pulsado == b_abrir:
        v.abrir_bd()
    elif pulsado == b_ejemplo:
        v.abrir_ejemplo()
    else:
        v.close()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
