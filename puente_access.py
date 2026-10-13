#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
puente_access.py — compatibilidad ACCDB/MDB para AccesoLibre
=============================================================
Puente de interoperabilidad con archivos de Microsoft Access
(.accdb desde 2007 hasta 365, y .mdb de Jet 3/4), basado en la librería
libre Jackcess (LGPL), que implementa el formato en disco de Jet/ACE.
"""

import os
import urllib.request

import jpype
import jpype.imports

LIB_DIR = os.path.join(os.path.expanduser("~"), ".accesolibre", "java")

JARS = {
    "jackcess-4.0.5.jar":
        "https://repo1.maven.org/maven2/com/healthmarketscience/jackcess/"
        "jackcess/4.0.5/jackcess-4.0.5.jar",
    "commons-lang3-3.12.0.jar":
        "https://repo1.maven.org/maven2/org/apache/commons/commons-lang3/"
        "3.12.0/commons-lang3-3.12.0.jar",
    "commons-logging-1.2.jar":
        "https://repo1.maven.org/maven2/commons-logging/commons-logging/"
        "1.2/commons-logging-1.2.jar",
}

MAPA_TIPOS_IMPORT = {
    "TEXT": "TEXT", "MEMO": "TEXT", "GUID": "TEXT", "REP_ID": "TEXT",
    "BYTE": "INTEGER", "INT": "INTEGER", "LONG": "INTEGER", "BOOLEAN": "INTEGER",
    "FLOAT": "REAL", "MONEY": "REAL", "NUMERIC": "REAL",
    "SHORT_DATE_TIME": "TEXT",
    "BINARY": "BLOB", "OLE": "BLOB",
}


def asegurar_jars():
    os.makedirs(LIB_DIR, exist_ok=True)
    rutas = []
    for nombre, url in JARS.items():
        destino = os.path.join(LIB_DIR, nombre)
        if not os.path.exists(destino):
            try:
                urllib.request.urlretrieve(url, destino)
            except Exception as e:
                raise RuntimeError(
                    f"No se pudo descargar {nombre} ({e}).\n"
                    f"Descárguelo manualmente de {url}\n"
                    f"y cópielo en {LIB_DIR}")
        rutas.append(destino)
    return rutas


def iniciar_jvm():
    from modo_directo import iniciar_java
    iniciar_java()


def _bool(v):
    return str(v).lower() == "true"


def _py(v):
    if v is None:
        return None
    if isinstance(v, (str, int, float, bytes)):
        return v
    nombre = type(v).__name__
    if nombre in ("JBoolean", "Boolean"):
        return 1 if _bool(v) else 0
    if nombre in ("JLong", "JInt", "JShort", "JByte", "Long", "Integer",
                  "Short", "Byte", "BigInteger"):
        return int(str(v))
    if nombre in ("JDouble", "JFloat", "Double", "Float", "BigDecimal"):
        return float(str(v))
    if nombre in ("Timestamp", "Date", "Time", "Calendar"):
        return str(v)
    if nombre.startswith("JArray") or nombre.endswith("[]"):
        return None
    return str(v)


def importar_access(ruta_access, gestor):
    iniciar_jvm()
    from com.healthmarketscience.jackcess import DatabaseBuilder
    from java.io import File
    db = DatabaseBuilder.open(File(ruta_access))
    try:
        tablas = []
        for t in db.getTables():
            nombre = str(t.getName())
            if nombre.upper().startswith(("MSYS", "USYS")):
                continue
            tablas.append((nombre, t))
        resumen, existentes = [], set(gestor.tablas())
        for nombre, t in tablas:
            destino, k = nombre, 1
            while destino in existentes:
                destino = f"{nombre}_{k}"
                k += 1
            existentes.add(destino)
            cols = [{"nombre": str(c.getName()),
                     "tipo": MAPA_TIPOS_IMPORT.get(str(c.getType()), "TEXT"),
                     "auto": _bool(c.isAutoNumber())}
                    for c in t.getColumns()]
            pk = []
            idx = t.getPrimaryKeyIndex()
            if idx is not None:
                pk = [str(ic.getName()) for ic in idx.getColumns()]
            autos = [c["nombre"] for c in cols if c["auto"]]
            if len(pk) == 1 and autos and autos[0] == pk[0]:
                defs = [f"[{c['nombre']}] INTEGER PRIMARY KEY AUTOINCREMENT"
                        if c["nombre"] in autos else f"[{c['nombre']}] {c['tipo']}"
                        for c in cols]
            else:
                defs = [f"[{c['nombre']}] {c['tipo']}" for c in cols]
                if pk:
                    defs.append("PRIMARY KEY (" + ", ".join(f"[{p}]" for p in pk) + ")")
            gestor.ejecutar(f"CREATE TABLE [{destino}] (" + ", ".join(defs) + ")")
            n, marcas = 0, ",".join("?" * len(cols))
            nombres_cols = [c["nombre"] for c in cols]
            for fila in t:
                valores = [_py(fila.get(nc)) for nc in nombres_cols]
                gestor.ejecutar(f"INSERT INTO [{destino}] VALUES ({marcas})", valores)
                n += 1
            resumen.append(f"• {destino}: {n} filas")
        origen = os.path.basename(ruta_access)
        cuerpo = "\n".join(resumen) if resumen else "El archivo no contiene tablas de usuario."
        return f"Importación completada desde:\n{origen}\n\n{cuerpo}"
    finally:
        db.close()


def _tipo_jackcess(tipo_sqlite):
    from com.healthmarketscience.jackcess import DataType
    t = (tipo_sqlite or "").upper()
    if "INT" in t:
        return DataType.LONG
    if any(x in t for x in ("REAL", "DOUB", "FLOA", "DEC", "NUM")):
        return DataType.FLOAT
    return DataType.TEXT


def _java(v, tipo_sqlite):
    if v is None:
        return None
    t = (tipo_sqlite or "").upper()
    if "INT" in t:
        return int(v)
    if any(x in t for x in ("REAL", "DOUB", "FLOA", "DEC", "NUM")):
        return float(v)
    return str(v)


def exportar_access(ruta_salida, gestor):
    iniciar_jvm()
    from com.healthmarketscience.jackcess import DatabaseBuilder, TableBuilder
    from java.io import File
    if os.path.exists(ruta_salida):
        raise RuntimeError("El archivo de destino ya existe.")
    FF = Database.FileFormat
    formato = getattr(FF, "V2010", None) or getattr(FF, "V2007", None)
    db = DatabaseBuilder.create(formato, File(ruta_salida))
    try:
        resumen = []
        for tabla in gestor.tablas():
            cols = gestor.columnas(tabla)
            tb = TableBuilder(tabla)
            for c in cols:
                tb.addColumn(c["nombre"], _tipo_jackcess(c["tipo"]))
            pks = [c["nombre"] for c in cols if c["pk"]]
            if pks:
                tb.setPrimaryKey(*pks)
            t = tb.toTable(db)
            q = gestor.ejecutar(f"SELECT * FROM [{tabla}]")
            nombres = [c["nombre"] for c in cols]
            n = 0
            while q.next():
                fila = {c["nombre"]: _java(q.value(i), c["tipo"])
                        for i, c in enumerate(cols)}
                try:
                    t.addRow(fila)
                except Exception:
                    t.addRow(*[fila[x] for x in nombres])
                n += 1
            resumen.append(f"• {tabla}: {n} filas")
        return f"Exportación completada a:\n{ruta_salida}\n\n" + "\n".join(resumen)
    finally:
        db.close()