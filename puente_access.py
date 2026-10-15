#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
puente_access.py — compatibilidad ACCDB/MDB para AccesoLibre
=============================================================
Puente de interoperabilidad con archivos de Microsoft Access
(.accdb desde 2007 hasta 365, y .mdb de Jet 3/4), basado en la librería
libre Jackcess (LGPL), que implementa el formato en disco de Jet/ACE.

Requisitos:
  • JRE 11 o superior:   sudo apt install default-jre-headless
  • Paquete Python:      pip install JPype1
  • Las JAR se descargan solas la primera vez (o cópielas manualmente
    en ~/.accesolibre/java/).
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
    "jackcess-encrypt-3.0.0.jar":
        "https://repo1.maven.org/maven2/com/healthmarketscience/jackcess/"
        "jackcess-encrypt/3.0.0/jackcess-encrypt-3.0.0.jar",
    "bcpkix-jdk15on-1.70.jar":
        "https://repo1.maven.org/maven2/org/bouncycastle/bcpkix-jdk15on/"
        "1.70/bcpkix-jdk15on-1.70.jar",
    "bcprov-jdk15on-1.70.jar":
        "https://repo1.maven.org/maven2/org/bouncycastle/bcprov-jdk15on/"
        "1.70/bcprov-jdk15on-1.70.jar",
}

MAPA_TIPOS_IMPORT = {
    "TEXT": "TEXT", "MEMO": "TEXT", "GUID": "TEXT", "REP_ID": "TEXT",
    "BYTE": "INTEGER", "INT": "INTEGER", "LONG": "INTEGER", "BOOLEAN": "INTEGER",
    "FLOAT": "REAL", "MONEY": "REAL", "NUMERIC": "REAL",
    "SHORT_DATE_TIME": "TEXT",
    "BINARY": "BLOB", "OLE": "BLOB",
}


def asegurar_jars():
    """Descarga las librerías Java (Jackcess y dependencias) si faltan."""
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


def
