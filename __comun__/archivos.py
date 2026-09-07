"""Reglas compartidas para descartar temporales y copias de Windows.

No contiene patrones de planillas ni recorre por sí solo el árbol del caso: cada
programa conserva su caché y su búsqueda de dominio. Este módulo decide qué
nombres nunca deben ganar por fecha y elige el más reciente entre entradas ya
obtenidas por el consumidor.
"""

import re
from pathlib import Path


PATRON_COPIA = re.compile(
    r"(-\s*cop(?:ia|y)(?:\s*\(\d+\))?|\(\d+\))\s*$", re.IGNORECASE
)


def es_temporal(nombre):
    """Detecta temporales de Excel y entradas ocultas usadas como auxiliares."""
    nombre = Path(str(nombre)).name
    return nombre.startswith("~$") or nombre.startswith(".")


def es_copia(nombre):
    """Detecta ``- copia``, ``- Copy``, sus numeradas y el sufijo ``(N)``."""
    return bool(PATRON_COPIA.search(Path(str(nombre)).stem))


def preferir_originales(entradas, nombre=lambda entrada: entrada.name):
    """Descarta copias si existe al menos un original; si no, conserva todas."""
    entradas = list(entradas)
    originales = [entrada for entrada in entradas if not es_copia(nombre(entrada))]
    return originales or entradas


def mas_reciente(entradas, mtime=lambda entrada: entrada.stat().st_mtime):
    """Devuelve la entrada más reciente o ``None`` sin volver a recorrer carpetas."""
    entradas = list(entradas)
    return max(entradas, key=mtime) if entradas else None
