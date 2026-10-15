#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
modo_directo.py — Lectura/escritura NATIVA sobre .accdb y .mdb
=============================================================
Edita las filas del archivo Access original mediante Jackcess (LGPL).
Soporta archivos cifrados con contraseña (pide la clave al abrir).
"""

import os
import sys
import glob
import shutil
from datetime import datetime

import jpype
import jpype.imports

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMessageBox, QStyledItemDelegate, QTableView,
    QVBoxLayout, QWidget, QInputDialog, QPushButton,
)


def recurso(relativo):
    bases = []
    if getattr(sys, "frozen", False):
        bases.append(os.path.dirname(os.path.abspath(sys.executable)))
        if hasattr(sys, "_MEIPASS"):
            bases.append(sys._MEIPASS)
    bases.append(os.path.dirname(os.path.abspath(__file__)))
    for b in bases:
        p = os.path.join(b, relativo)
        if os.path.exists(p):
            return p
    return os.path.join(bases[0], relativo)


def _buscar_libjvm():
    import platform
    bases = []
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            bases.append(os.path.join(sys._MEIPASS, "jre"))
        bases.append(os.path.join(os.path.dirname(sys.executable), "..", "Resources", "jre"))
        bases.append(os.path.join(os.path.dirname(sys.executable), "jre"))
    bases.append(recurso("jre"))
    bases.append("/Library/Java/JavaVirtualMachines")
    bases.append(os.path.expanduser("~/Library/Java/JavaVirtualMachines"))

    sistema = platform.system()
    if sistema == "Darwin":
        nombre_lib = "libjvm.dylib"
        rutas = ["lib/server/libjvm.dylib",
                 "jre/lib/server/libjvm.dylib",
                 "Contents/Home/lib/server/libjvm.dylib"]
    else:
        nombre_lib = "libjvm.so"
        rutas = ["lib/server/libjvm.so",
                 "jre/lib/server/libjvm.so"]

    for base in bases:
        if not os.path.isdir(base):
            continue
        for rel in rutas:
            cand = os.path.join(base, rel)
            if os.path.isfile(cand):
                return cand
        for arch in glob.glob(os.path.join(base, "**", nombre_lib), recursive=True):
            if os.path.isfile(arch):
                return arch
    return None


def iniciar_java():
    if jpype.isJVMStarted():
        return
    bases = []
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            bases.append(os.path.join(sys._MEIPASS, "jar"))
        bases.append(os.path.join(os.path.dirname(sys.executable), "..", "Resources", "jar"))
        bases.append(os.path.join(os.path.dirname(sys.executable), "jar"))
    bases.append(recurso("jar"))

    jars = []
    for base in bases:
        if os.path.isdir(base):
            for f in sorted(os.listdir(base)):
                if f.endswith(".jar"):
                    jars.append(os.path.join(base, f))
    if not jars:
        from puente_access import asegurar_jars
        jars = asegurar_jars()

    opts = ["-Djava.awt.headless=true", "-Dfile.encoding=UTF-8"]
    jvm = _buscar_libjvm()
    try:
        if jvm:
            jpype.startJVM(*opts, jvmPath=jvm, classpath=jars, convertStrings=True)
        else:
            jpype.startJVM(*opts, classpath=jars, convertStrings=True)
    except TypeError:
        if not jvm:
            raise
        raiz = os.path.dirname(os.path.dirname(os.path.dirname(jvm)))
        os.environ["JAVA_HOME"] = raiz
        os.environ["PATH"] = (os.path.join(raiz, "bin") + os.pathsep
                              + os.environ.get("PATH", ""))
        jpype.startJVM(*opts, classpath=jars, convertStrings=True)
    except Exception as e:
        raise RuntimeError(
            "No se pudo iniciar Java.\n"
            "Instale: sudo apt install default-jre-headless\n\n" + str(e)) from e


def _es_binario(tipo):
    t = tipo.upper()
    return any(x in t for x in ("OLE", "BINARY", "ATTACH", "COMPLEX"))


def _a_py(v, tipo):
    if v is None or _es_binario(tipo):
        return None
    t = tipo.upper()
    if "DATE" in t or "TIME" in t:
        return str(v)
    if "MONEY" in t or t == "NUMERIC":
        return str(v)
    if "BOOLEAN" in t:
        return str(v).strip().lower() in ("true", "1", "-1")
    if isinstance(v, (bool, int, float, str)):
        return v
    n = type(v).__name__
    if n in ("JBoolean", "Boolean"):
        return str(v).lower() in ("true", "1", "-1")
    if n in ("JLong", "JInt", "JShort", "JByte", "Long", "Integer",
             "Short", "Byte", "BigInteger"):
        try:
            return int(str(v))
        except ValueError:
            return str(v)
    if n in ("JDouble", "JFloat", "Double", "Float"):
        try:
            return float(str(v))
        except ValueError:
            return str(v)
    return str(v)


def normalizar(texto, tipo):
    t = tipo.upper()
    s = str(texto).strip()
    if s == "":
        return None
    if "BOOLEAN" in t:
        return s.lower() in ("1", "true", "sí", "si", "s", "yes", "verdadero")
    if "DATE" in t or "TIME" in t:
        s2 = s.replace("T", " ")
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(s2, fmt).strftime("%Y-%m-%d %H:%M:%S")
            except ValueError:
                pass
        raise ValueError("Fecha inválida. Use AAAA-MM-DD o AAAA-MM-DD HH:MM:SS.")
    if "MONEY" in t or t == "NUMERIC":
        try:
            float(s.replace(",", "."))
        except ValueError:
            raise ValueError("Se requiere un número.")
        return s.replace(",", ".")
    if any(x in t for x in ("BYTE", "SHORT", "INT", "LONG")):
        try:
            return int(s)
        except ValueError:
            try:
                f = float(s.replace(",", "."))
            except ValueError:
                raise ValueError("Se requiere un número entero.")
            if f != int(f):
                raise ValueError("Se requiere un número entero.")
            return int(f)
    if any(x in t for x in ("FLOAT", "DOUBLE", "REAL")):
        try:
            return float(s.replace(",", "."))
        except ValueError:
            raise ValueError("Se requiere un número.")
    return s


def _a_java(v, tipo):
    if v is None:
        return None
    t = tipo.upper()
    from java.lang import Integer, Long, Double, Boolean, String
    if "DATE" in t or "TIME" in t:
        from java.sql import Timestamp
        s = str(v)
        if len(s) == 10:
            s += " 00:00:00"
        if "." not in s:
            s += ".0"
        return Timestamp.valueOf(s)
    if "BOOLEAN" in t:
        return Boolean(bool(v))
    if "MONEY" in t or t == "NUMERIC":
        from java.math import BigDecimal
        return BigDecimal(str(v))
    if any(x in t for x in ("BYTE", "SHORT", "INT", "LONG")):
        n = int(v)
        return Long(n) if ("LONG" in t or "BIGINT" in t) else Integer(n)
    if any(x in t for x in ("FLOAT", "DOUBLE", "REAL")):
        return Double(float(v))
    return String(str(v))


class AccesoDirecto:
    def __init__(self, ruta, password=None):
        self.ruta = os.path.abspath(ruta)
        if not os.path.isfile(self.ruta):
            raise FileNotFoundError(f"No existe el archivo: {self.ruta}")
        iniciar_java()
        from java.io import File
        from com.healthmarketscience.jackcess import DatabaseBuilder

        self.solo_lectura = not os.access(self.ruta, os.W_OK)
        builder = DatabaseBuilder(File(self.ruta))
        builder.setReadOnly(self.solo_lectura)

        if password:
            try:
                from com.healthmarketscience.jackcess.crypt import CryptCodecProvider
                builder.setCodecProvider(CryptCodecProvider(password))
            except Exception as e:
                raise RuntimeError(
                    f"Error al configurar el descifrado: {e}\n"
                    "Faltan las librerías jackcess-encrypt / Bouncy Castle en 'jar'.")

        try:
            self.db = builder.open()
        except Exception as e:
            msg = str(e)
            if ("UnsupportedCodecException" in msg or "Decoding not supported" in msg
                    or "encrypted" in msg.lower() or "password" in msg.lower()):
                raise RuntimeError(
                    "El archivo está cifrado con contraseña.\n"
                    "Por favor, introduzca la contraseña correcta.") from e
            raise RuntimeError(
                "No se pudo abrir el archivo Access.\n"
                "Compruebe que:\n"
                "  • no está abierto en Microsoft Access;\n"
                "  • no está dañado o es una versión muy antigua.\n\n"
                f"Detalle técnico: {msg}") from e

        self.aviso_lock = os.path.exists(os.path.splitext(self.ruta)[0] + ".laccdb")

    def tablas(self):
        out = []
        for t in self.db.getTables():
            n = str(t.getName())
            if not n.upper().startswith(("MSYS", "USYS")):
                out.append(n)
        return sorted(out)

    def metadata(self, nombre):
        t = self.db.getTable(nombre)
        cols = []
        for c in t.getColumns():
            tipo = str(c.getType()).split(".")[-1].upper()
            complejo = bool(getattr(c, "isComplex", lambda: False)())
            cols.append({"nombre": str(c.getName()), "tipo": tipo,
                         "auto": str(c.isAutoNumber()).lower() == "true",
                         "blob": _es_binario(tipo) or complejo, "pk": False})
        idx = t.getPrimaryKeyIndex()
        pk = [str(ic.getName()) for ic in idx.getColumns()] if idx is not None else []
        for c in cols:
            c["pk"] = c["nombre"] in pk
        return cols, pk

    def leer(self, nombre):
        cols, pk = self.metadata(nombre)
        t = self.db.getTable(nombre)
        filas = [{c["nombre"]: _a_py(row.get(c["nombre"]), c["tipo"]) for c in cols}
                 for row in t]
        return cols, pk, filas

    def copia_seguridad(self):
        base, ext = os.path.splitext(self.ruta)
        destino = f"{base}.copia-{datetime.now():%Y%m%d-%H%M%S}{ext}"
        shutil.copy2(self.ruta, destino)
        return destino

    def insertar(self, nombre, cols, fila):
        from java.util import HashMap
        t = self.db.getTable(nombre)
        m = HashMap()
        for c in cols:
            if c["auto"] or c["blob"]:
                continue
            m.put(c["nombre"], _a_java(fila.get(c["nombre"]), c["tipo"]))
        t.addRow(m)

    def actualizar(self, nombre, cols, fila):
        from java.util import HashMap
        t = self.db.getTable(nombre)
        m = HashMap()
        for c in cols:
            if c["blob"]:
                continue
            m.put(c["nombre"], _a_java(fila.get(c["nombre"]), c["tipo"]))
        if t.updateRow(m) is False:
            raise RuntimeError("No se encontró la fila original.")

    def borrar(self, nombre, cols, fila):
        from java.util import HashMap
        t = self.db.getTable(nombre)
        m = HashMap()
        for c in cols:
            if c["pk"]:
                m.put(c["nombre"], _a_java(fila.get(c["nombre"]), c["tipo"]))
        if t.deleteRow(m) is False:
            raise RuntimeError("No se encontró la fila que se quería eliminar.")

    def cerrar(self):
        if getattr(self, "db", None) is not None:
            try:
                self.db.close()
            finally:
                self.db = None


class ModeloDirecto(QAbstractTableModel):
    def __init__(self, acceso, nombre):
        super().__init__()
        self.acceso = acceso
        self.nombre = nombre
        self.error = ""
        self._recargar()

    def _recargar(self):
        self.beginResetModel()
        self.cols, self.pk, self.filas = self.acceso.leer(self.nombre)
        self.orig = [dict(f) for f in self.filas]
        self.nuevos, self.sucios, self.bajas = set(), set(), set()
        self.endResetModel()

    @property
    def lectura(self):
        return self.acceso.solo_lectura or not self.pk

    def rowCount(self, parent=None):
        return len(self.filas)

    def columnCount(self, parent=None):
        return len(self.cols)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation != Qt.Orientation.Horizontal:
            return section + 1
        if role == Qt.ItemDataRole.DisplayRole:
            c = self.cols[section]
            return c["nombre"] + (" *" if c["pk"] else "")
        if role == Qt.ItemDataRole.ToolTipRole:
            c = self.cols[section]
            partes = [c["tipo"]]
            if c["pk"]:
                partes.append("clave principal (no editable)")
            if c["auto"]:
                partes.append("autonumérico (lo genera Access)")
            if c["blob"]:
                partes.append("binario (solo lectura)")
            return " · ".join(partes)
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        base = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        if self.lectura or index.row() in self.bajas:
            return base
        c = self.cols[index.column()]
        if c["auto"] or c["blob"] or c["pk"]:
            return base
        return base | Qt.ItemFlag.ItemIsEditable

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        r, col = index.row(), index.column()
        c = self.cols[col]
        v = self.filas[r].get(c["nombre"])
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            if v is None:
                return ""
            if "BOOLEAN" in c["tipo"].upper():
                return "Sí" if v else "No"
            return str(v)
        if role == Qt.ItemDataRole.BackgroundRole:
            if r in self.bajas:
                return QBrush(QColor("#ffe0e0"))
            if r in self.nuevos:
                return QBrush(QColor("#e2f5e2"))
            if r in self.sucios:
                return QBrush(QColor("#fff3cf"))
        if role == Qt.ItemDataRole.ToolTipRole:
            if r in self.bajas:
                return "Marcada para eliminar al guardar"
            if r in self.sucios:
                return "Modificada · sin guardar"
            if r in self.nuevos:
                return "Registro nuevo · sin guardar"
        return None

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role != Qt.ItemDataRole.EditRole or not (
                self.flags(index) & Qt.ItemFlag.ItemIsEditable):
            return False
        r = index.row()
        c = self.cols[index.column()]
        texto = "" if value is None else str(value)
        try:
            nuevo = normalizar(texto, c["tipo"]) if texto.strip() else None
        except ValueError as e:
            self.error = f"{c['nombre']}: {e}"
            return False
        if self.filas[r].get(c["nombre"]) == nuevo:
            return False
        self.filas[r][c["nombre"]] = nuevo
        if r not in self.nuevos:
            self.sucios.add(r)
        self.dataChanged.emit(index, index)
        return True

    def pendientes(self):
        return len(self.nuevos) + len(self.sucios) + len(self.bajas)

    def nuevo_registro(self):
        if self.lectura:
            return False
        self.beginInsertRows(QModelIndex(), len(self.filas), len(self.filas))
        self.filas.append({c["nombre"]: None for c in self.cols})
        self.orig.append({})
        self.nuevos.add(len(self.filas) - 1)
        self.endInsertRows()
        return True

    def marcar_baja(self, r):
        if self.lectura or not (0 <= r < len(self.filas)):
            return
        self.bajas.add(r)
        self.dataChanged.emit(self.index(r, 0), self.index(r, len(self.cols) - 1))

    def descartar(self):
        self._recargar()

    def guardar(self, hacer_backup=True):
        if self.lectura:
            raise RuntimeError(
                "Tabla de solo lectura: se requiere clave principal "
                "y permisos de escritura sobre el archivo.")
        backup = self.acceso.copia_seguridad() if hacer_backup else None
        aplicados, errores = 0, []
        for r in sorted(self.nuevos):
            if r in self.bajas:
                continue
            try:
                self.acceso.insertar(self.nombre, self.cols, self.filas[r])
                aplicados += 1
            except Exception as e:
                errores.append(f"Alta (fila {r + 1}): {e}")
        for r in sorted(self.sucios):
            if r in self.bajas or r in self.nuevos:
                continue
            try:
                self.acceso.actualizar(self.nombre, self.cols, self.filas[r])
                aplicados += 1
            except Exception as e:
                errores.append(f"Modificación (fila {r + 1}): {e}")
        for r in sorted(self.bajas):
            if r in self.nuevos:
                continue
            try:
                self.acceso.borrar(self.nombre, self.cols, self.orig[r])
                aplicados += 1
            except Exception as e:
                errores.append(f"Baja (fila {r + 1}): {e}")
        self._recargar()
        return aplicados, errores, backup


class DelegadoTipo(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        c = index.model().cols[index.column()]
        if "BOOLEAN" in c["tipo"].upper():
            cmb = QComboBox(parent)
            cmb.addItems(["No", "Sí"])
            return cmb
        return QLineEdit(parent)

    def setEditorData(self, editor, index):
        v = str(index.data(Qt.ItemDataRole.EditRole) or "")
        if isinstance(editor, QComboBox):
            editor.setCurrentText("Sí" if v == "Sí" else "No")
        else:
            editor.setText(v)

    def setModelData(self, editor, model, index):
        v = (editor.currentText() if isinstance(editor, QComboBox)
             else editor.text())
        model.setData(index, v, Qt.ItemDataRole.EditRole)


class VistaDirecto(QWidget):
    def __init__(self, ruta, password=None):
        super().__init__()
        self.ruta = ruta
        self.password = password
        self.acceso = None
        self.modelo = None
        self._cerrado = False
        self._abrir_archivo()

    def _abrir_archivo(self):
        try:
            self.acceso = AccesoDirecto(self.ruta, self.password)
        except RuntimeError as e:
            msg = str(e)
            if ("contraseña" in msg.lower() or "password" in msg.lower()
                    or "cifrado" in msg.lower()):
                password, ok = QInputDialog.getText(
                    self, "Archivo cifrado",
                    "Este archivo Access está protegido con contraseña.\n"
                    "Introduzca la contraseña para abrirlo:",
                    QLineEdit.EchoMode.Password)
                if not (ok and password):
                    return
                try:
                    self.acceso = AccesoDirecto(self.ruta, password)
                    self.password = password
                except Exception as e2:
                    QMessageBox.critical(
                        self, "Error",
                        f"No se pudo abrir el archivo con esa contraseña:\n{e2}")
                    return
            else:
                QMessageBox.critical(self, "Error", msg)
                return
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return
        if self.acceso is None:
            return
        self._construir_interfaz()

    def _construir_interfaz(self):
        lay = QVBoxLayout(self)
        titulo = QLabel(f"EDICIÓN NATIVA · {self.acceso.ruta}")
        titulo.setStyleSheet("font-weight:bold;color:#2b579a;font-size:13px")
        lay.addWidget(titulo)
        if self.acceso.aviso_lock:
            aviso = QLabel("⚠ Existe un archivo .laccdb. "
                           "Asegúrese de que MS Access NO tiene esta base abierta.")
            aviso.setStyleSheet("color:#a33;font-weight:bold")
            lay.addWidget(aviso)
        fila1 = QHBoxLayout()
        fila1.addWidget(QLabel("Tabla:"))
        self.cmb = QComboBox()
        self.cmb.addItems(self.acceso.tablas())
        self.cmb.setMinimumWidth(220)
        fila1.addWidget(self.cmb, 1)
        self._boton(fila1, "Abrir tabla", self.abrir_tabla)
        self._boton(fila1, "Recargar", self.recargar)
        lay.addLayout(fila1)
        fila2 = QHBoxLayout()
        self.btnGuardar = self._boton(fila2, "GUARDAR en el .accdb", self.guardar)
        self._boton(fila2, "Descartar cambios", self.descartar)
        self._boton(fila2, "Nuevo registro", self.nuevo)
        self._boton(fila2, "Marcar para eliminar", self.marcar_baja)
        self.chkBackup = QCheckBox("Copia de seguridad antes de guardar")
        self.chkBackup.setChecked(True)
        fila2.addWidget(self.chkBackup)
        fila2.addStretch(1)
        lay.addLayout(fila2)
        self.vista = QTableView()
        self.vista.setAlternatingRowColors(True)
        self.vista.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.vista.setEditTrigger(QAbstractItemView.EditTrigger.DoubleClicked |
                                  QAbstractItemView.EditTrigger.EditKeyPressed)
        self.vista.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(self.vista, 1)
        self.estado = QLabel(
            "Doble clic en una celda para editar.  "
            "Amarillo = modificado · Verde = nuevo · Rojo = marcado para eliminar.")
        self.estado.setWordWrap(True)
        lay.addWidget(self.estado)
        self._actualizar_botones()

    def _boton(self, layout, texto, slot):
        b = QPushButton(texto)
        b.clicked.connect(slot)
        layout.addWidget(b)
        return b

    def _actualizar_botones(self):
        hay = self.modelo is not None
        self.btnGuardar.setEnabled(hay and not self.modelo.lectura
                                   and self.modelo.pendientes() > 0)

    def _preguntar_descarte(self):
        if self.modelo and self.modelo.pendientes():
            r = QMessageBox.question(
                self, "Cambios sin guardar",
                f"Hay {self.modelo.pendientes()} cambio(s) sin guardar. ¿Descartarlos?")
            return r == QMessageBox.StandardButton.Yes
        return True

    def abrir_tabla(self):
        if not self.cmb.currentText():
            return
        if not self._preguntar_descarte():
            return
        try:
            self.modelo = ModeloDirecto(self.acceso, self.cmb.currentText())
        except Exception as e:
            QMessageBox.critical(self, "Error al leer la tabla", str(e))
            return
        self.vista.setModel(self.modelo)
        self.vista.setItemDelegate(DelegadoTipo(self.vista))
        self.vista.resizeColumnsToContents()
        if self.modelo.lectura:
            self.estado.setText("SOLO LECTURA: esta tabla no tiene clave principal "
                                "o el archivo no permite escritura.")
        else:
            self.estado.setText(f"{self.modelo.rowCount()} filas · "
                                "Doble clic para editar · Guardar escribe en el .accdb original.")
        self._actualizar_botones()

    def recargar(self):
        if self.modelo and not self._preguntar_descarte():
            return
        if self.modelo:
            self.abrir_tabla()

    def nuevo(self):
        if not self.modelo:
            return
        if self.modelo.nuevo_registro():
            self.vista.scrollToBottom()
            self._actualizar_botones()
        else:
            QMessageBox.information(self, "Solo lectura", "Esta tabla no admite edición.")

    def marcar_baja(self):
        if not self.modelo:
            return
        idx = self.vista.currentIndex()
        if idx.isValid():
            self.modelo.marcar_baja(idx.row())
            self._actualizar_botones()
        else:
            QMessageBox.information(self, "Seleccionar fila",
                                    "Pulse primero una celda de la fila a eliminar.")

    def descartar(self):
        if self.modelo:
            self.modelo.descartar()
            self._actualizar_botones()

    def guardar(self):
        if not self.modelo or not self.modelo.pendientes():
            return
        n = self.modelo.pendientes()
        r = QMessageBox.question(
            self, "Guardar en el archivo Access",
            f"Se van a aplicar {n} cambio(s) DIRECTAMENTE sobre:\n{self.acceso.ruta}\n\n"
            "Los guardados se aplican fila a fila (sin transacción). ¿Continuar?")
        if r != QMessageBox.StandardButton.Yes:
            return
        try:
            aplicados, errores, backup = self.modelo.guardar(self.chkBackup.isChecked())
        except Exception as e:
            QMessageBox.critical(self, "No se pudo guardar", str(e))
            return
        self._actualizar_botones()
        if errores:
            texto = (f"Aplicados {aplicados} cambio(s).\n"
                     f"Fallaron {len(errores)} operación(es):\n\n"
                     + "\n".join(errores[:8]))
            if backup:
                texto += f"\n\nCopia de seguridad: {backup}"
            QMessageBox.warning(self, "Guardado parcial", texto)
        else:
            msg = f"✔ {aplicados} cambio(s) guardados en el .accdb original."
            if backup:
                msg += f"\n\nCopia creada:\n{backup}"
            QMessageBox.information(self, "Guardado nativo", msg)
        self.estado.setText(msg)

    def cerrar(self):
        if self._cerrado:
            return True
        if self.modelo and self.modelo.pendientes():
            mb = QMessageBox(self)
            mb.setWindowTitle("Cambios sin guardar")
            mb.setText("Hay cambios sin guardar en el archivo Access.")
            b_guardar = mb.addButton("Guardar", QMessageBox.ButtonRole.AcceptRole)
            mb.addButton("Descartar", QMessageBox.ButtonRole.DestructiveRole)
            b_cancelar = mb.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)
            mb.exec()
            if mb.clickedButton() == b_cancelar:
                return False
            if mb.clickedButton() == b_guardar:
                self.guardar()
                if self.modelo.pendientes():
                    return False
        if self.acceso:
            self.acceso.cerrar()
        self._cerrado = True
        return True

    def closeEvent(self, evento):
        if self.cerrar():
            evento.accept()
        else:
            evento.ignore()
