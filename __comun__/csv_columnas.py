"""Lectura de un .csv por letra de columna y fila, como si fuera una hoja.

Nace con la versión 15 minutos: el CO de la hoja "SC y CO" de la planilla 5
sale de ``Reporte_CRA_15min_AAMM.csv`` y no de un Excel. Lo usan
``Revisor_Relq_15min/actualizadores/Actualiza_SC_CO.py`` (para pegarlo) y el
Revisor 15 min (V10, para compararlo), así que vive acá y no copiado en los dos.

Devuelve lo mismo que ``excel_xml.leer_columnas_rapido``: ``{COL: {fila: valor}}``,
con la fila contada como en Excel (el encabezado es la fila 1). Las celdas
vacías no aparecen. Ante un archivo que no se puede leer devuelve ``None``.

Formato del Reporte_CRA real (ver Balance_SSCC, ``docs/Estructura_Archivos_
Reales.md`` §D.1): UTF-8, separador ``;``, punto decimal. Igual se toleran
``,`` o tabulador como separador, coma decimal cuando el separador es ``;``
y archivos guardados en cp1252.

Solo biblioteca estándar.
"""

import csv
import re
from pathlib import Path

_RE_ENTERO = re.compile(r"[-+]?\d+")
_RE_DECIMAL = re.compile(r"[-+]?(\d+\.\d*|\.\d+|\d+)([eE][-+]?\d+)?")


def letra_columna(n):
    """1 -> A, 27 -> AA."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _separador(primera_linea):
    cuentas = {sep: primera_linea.count(sep) for sep in (";", ",", "\t")}
    sep = max(cuentas, key=cuentas.get)
    return sep if cuentas[sep] else ";"


def valor_csv(texto, separador=";"):
    """Texto de una celda -> int, float, str o None (vacía).

    Con separador ``;`` se acepta la coma como decimal ("3,5"). Un número con
    ceros a la izquierda se deja como número igual que haría Excel ("0038" -> 38).
    """
    t = (texto or "").strip()
    if not t:
        return None
    if _RE_ENTERO.fullmatch(t):
        return int(t)
    if separador == ";" and "," in t and "." not in t:
        t2 = t.replace(",", ".")
    else:
        t2 = t
    if _RE_DECIMAL.fullmatch(t2):
        return float(t2)
    return t


def leer_filas_csv(ruta):
    """(filas, separador): todas las filas como listas de texto, o (None, None)."""
    ruta = Path(ruta)
    for cod in ("utf-8-sig", "cp1252"):
        try:
            with open(ruta, "r", encoding=cod, newline="") as f:
                contenido = f.read()
            break
        except UnicodeDecodeError:
            continue
        except OSError:
            return None, None
    else:
        return None, None
    primera = contenido.split("\n", 1)[0]
    sep = _separador(primera)
    filas = list(csv.reader(contenido.splitlines(), delimiter=sep))
    return filas, sep


def leer_columnas_csv(ruta, columnas=None, fila_inicio=1, log=None):
    """``{COL: {fila: valor}}`` de las columnas pedidas (todas si ``None``)."""
    filas, sep = leer_filas_csv(ruta)
    if filas is None:
        if log:
            log(f"    ! No se pudo leer {Path(ruta).name}")
        return None
    objetivo = None if columnas is None else [str(c).upper() for c in columnas]
    datos = {c: {} for c in (objetivo or [])}
    fila_inicio = int(fila_inicio)
    for nfila, fila in enumerate(filas, start=1):
        if nfila < fila_inicio:
            continue
        for i, texto in enumerate(fila, start=1):
            col = letra_columna(i)
            if objetivo is not None and col not in datos:
                continue
            v = valor_csv(texto, sep)
            if v is None:
                continue
            datos.setdefault(col, {})[nfila] = v
    return datos
