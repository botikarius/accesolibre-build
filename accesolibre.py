#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AccesoLibre 2007
================
Suite de bases de datos relacional de escritorio, de código abierto.
Motor: SQLite + Compatibilidad nativa ACCDB/MDB vía Jackcess.
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
    QApplication, QCheckBox, QComboBox, QDataWidgetMapper, QDialog, QDialogButtonBox,
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
# Motor de datos
# ---------------------------------------------------------------------------
class GestorBD:
    def __init__(self, ruta):
        self.ruta = ruta
        
        # CORRECCIÓN DEFINITIVA DE COMPATIBILIDAD PySide6/PyQt
        DEFAULT_CONN_NAME = getattr(QSqlDatabase, 'defaultConnection', 'qt_sql_default_connection')
        
        if QSqlDatabase.contains(DEFAULT_CONN_NAME):
            QSqlDatabase.removeDatabase(DEFAULT_CONN_NAME)
            
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
# Diseñador de tablas
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


# ---------------------------------------------------------------------------
# Diseñador de consultas
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Diseñadores de formulario e informe
# ---------------------------------------------------------------------------
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
# Objetos abiertos en el área de trabajo
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

        self.statusBar().showMessage("Cree o
