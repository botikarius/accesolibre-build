#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
modo_directo.py — Lectura/escritura NATIVA sobre .accdb y .mdb
=============================================================
Edita las filas del archivo Access original mediante Jackcess (LGPL).
Soporta archivos cifrados con contraseña.

Requiere: JPype1 + JRE (embebido si se empaqueta el programa con jlink).
"""

import os
import sys
import glob
import shutil
from datetime import datetime

import jpype
import jpype.imports

from PySide6.QtCore import Qt, QAbstractTableModel
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMessageBox, QStyledItemDelegate, QTableView,
    QVBoxLayout, QWidget, QInputDialog, QPushButton,
)


# ---------------------------------------------------------------------------
# Localización de recursos y arranque de la JVM
# ---------------------------------------------------------------------------
def recurso(relativo):
    """Funciona en ejecución desde fuente y dentro del binario PyInstaller."""
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
    """Busca libjvm.dylib en las rutas típicas de macOS dentro del bundle .app"""
    import platform
    import glob as _glob
    
    bases = []
    
    # Ruta dentro del bundle PyInstaller en macOS
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            bases.append(os.path.join(sys._MEIPASS, "jre"))
        bases.append(os.path.join(os.path.dirname(sys.executable), "..", "Resources", "jre"))
        bases.append(os.path.join(os.path.dirname(sys.executable), "jre"))
    
    # Ruta de desarrollo (cuando ejecutas desde fuente)
    bases.append(recurso("jre"))
    
    # Rutas típicas de Java instalado en macOS
    bases.append("/Library/Java/JavaVirtualMachines")
    bases.append(os.path.expanduser("~/Library/Java/JavaVirtualMachines"))
    
    # Buscar libjvm.dylib (macOS) o libjvm.so (Linux)
    sistema = platform.system()
    if sistema == "Darwin":  # macOS
        nombre_lib = "libjvm.dylib"
        rutas_busqueda = [
            "lib/server/libjvm.dylib",
            "jre/lib/server/libjvm.dylib",
            "Contents/Home/lib/server/libjvm.dylib",
        ]
    else:  # Linux
        nombre_lib = "libjvm.so"
        rutas_busqueda = [
            "lib/server/libjvm.so",
            "jre/lib/server/libjvm.so",
        ]
    
    for base in bases:
        if not os.path.isdir(base):
            continue
        for ruta_rel in rutas_busqueda:
            candidato = os.path.join(base, ruta_rel)
            if os.path.isfile(candidato):
                return candidato
        
        # Búsqueda recursiva por si la estructura es diferente
        for archivo in _glob.glob(os.path.join(base, "**", nombre_lib), recursive=True):
            if os.path.isfile(archivo):
                return archivo
    
    return None


def iniciar_java():
    """Arranca la JVM una sola vez, usando el JRE embebido si existe."""
    if jpype.isJVMStarted():
        return

    # Buscar JARs en la carpeta 'jar' (tanto en desarrollo como en bundle)
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
        # Si no hay JARs locales, descargar los básicos (sin cifrado)
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
        os.environ["PATH"] = (os.path.join(raiz, "bin") + os.pathsep + os.environ.get("PATH", ""))
        jpype.startJVM(*opts, classpath=jars, convertStrings=True)
    except Exception as e:
        raise RuntimeError(
            "No se pudo iniciar Java.\n"
            "Instale: sudo apt install default-jre-headless\n\n" + str(e)) from e


# ---------------------------------------------------------------------------
# Conversión de valores Java <-> Python
# ---------------------------------------------------------------------------
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
    """Valida y normaliza lo que el usuario escribe en una celda."""
    t = tipo.upper()
    s = str(texto).strip()
    if s == "":
        return None

    if "BOOLEAN" in t:
        return s.lower() in ("1", "true", "sí", "si", "s", "yes", "verdadero")

    if "DATE" in t or "TIME" in t:
        s2 = s.replace("T", " ")
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
                    "%Y-%m-%d", "%d/%m/%Y"):
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
    """Construye objetos Java EXPLÍCITOS (evita ambigüedades de sobrecarga)."""
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


# ---------------------------------------------------------------------------
# Acceso nativo al archivo
# ---------------------------------------------------------------------------
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
        
        # Si hay contraseña, configurar el codec de cifrado
        if password:
            try:
                from com.healthmarketscience.jackcess.crypt import CryptCodecProvider
                codec = CryptCodecProvider(password)
                builder.setCodecProvider(codec)
            except Exception as e:
                raise RuntimeError(
                    f"Error al configurar el descifrado: {e}\n"
                    "Asegúrese de que las librerías jackcess-encrypt y Bouncy Castle "
                    "están disponibles en la carpeta 'jar'.")
        
        try:
            self.db = builder.open()
        except Exception as e:
            error_msg = str(e)
            if "UnsupportedCodecException" in error_msg or "Decoding not supported" in error_msg:
                raise RuntimeError(
                    "El archivo está cifrado con contraseña.\n"
                    "Por favor, introduzca la contraseña correcta.") from e
            elif "encrypted" in error_msg.lower() or "password" in error_msg.lower():
                raise RuntimeError(
                    "El archivo está protegido con contraseña.\n"
                    "Por favor, introduzca la contraseña correcta.") from e
            else:
                raise RuntimeError(
                    "No se pudo abrir el archivo Access.\n"
                    "Compruebe que:\n"
                    "  • no está abierto en Microsoft Access;\n"
                    "  • no está dañado o es una versión muy antigua.\n\n"
                    f"Detalle técnico: {error_msg}") from e

        self.aviso_lock = os.path.exists(
            os.path.splitext(self.ruta)[0] + ".laccdb")

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
            cols.append({
                "nombre": str(c.getName()),
                "tipo": tipo,
                "auto": bool(str(c.isAutoNumber()).lower() == "true"),
                "blob": _es_binario(tipo) or complejo,
                "pk": False,
            })
        idx = t.getPrimaryKeyIndex()
        pk
