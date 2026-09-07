"""Búsqueda de carpetas y archivos con caché acotada a una relectura."""

from datetime import datetime
import os
from pathlib import Path
import re

from __comun__ import archivos as _archivos
from __comun__ import texto as _texto


TOL_MTIME = 2
normalizar = _texto.suave_textual
es_temporal = _archivos.es_temporal
RE_COPIA = _archivos.PATRON_COPIA
es_copia = _archivos.es_copia

# Apagado por omisión: fuera del context manager las fechas siempre vienen del disco.
_DIR_CACHE = {"on": False, "datos": {}, "hits": 0, "scans": 0}


def _escanear(carpeta):
    """Devuelve subcarpetas y archivos en un único recorrido de ``carpeta``."""
    subs, archs = {}, {}
    _DIR_CACHE["scans"] += 1
    try:
        with os.scandir(str(carpeta)) as it:
            for entrada in it:
                try:
                    if entrada.is_dir():
                        subs[normalizar(entrada.name)] = Path(entrada.path)
                    elif entrada.is_file():
                        stat = entrada.stat()
                        archs[entrada.name] = (
                            Path(entrada.path), stat.st_mtime, stat.st_size
                        )
                except OSError:
                    continue
    except OSError:
        pass
    return subs, archs


def leer_dir(carpeta):
    """Obtiene el listado de una carpeta, usando el caché si está encendido."""
    if carpeta is None:
        return {}, {}
    if _DIR_CACHE["on"]:
        clave = str(carpeta)
        if clave in _DIR_CACHE["datos"]:
            _DIR_CACHE["hits"] += 1
            return _DIR_CACHE["datos"][clave]
        datos = _escanear(carpeta)
        _DIR_CACHE["datos"][clave] = datos
        return datos
    return _escanear(carpeta)


class cache_directorios:
    """Enciende el caché durante un bloque y publica ``(scans, hits)`` al salir."""

    def __enter__(self):
        _DIR_CACHE.update(on=True, datos={}, hits=0, scans=0)
        return self

    def __exit__(self, *_):
        self.stats = (_DIR_CACHE["scans"], _DIR_CACHE["hits"])
        _DIR_CACHE.update(on=False, datos={}, hits=0, scans=0)
        return False


def buscar_carpeta(base, nombre):
    """Busca una subcarpeta tolerando tildes, mayúsculas y espacios extra."""
    if not base:
        return None
    objetivo = normalizar(nombre)
    mapa, _ = leer_dir(base)
    if objetivo in mapa:
        return mapa[objetivo]
    subs = list(mapa.values())
    candidatas = [d for d in subs if normalizar(d.name).startswith(objetivo)]
    if candidatas:
        return sorted(candidatas, key=lambda d: len(d.name))[0]
    candidatas = [d for d in subs if objetivo in normalizar(d.name)]
    if candidatas:
        return sorted(candidatas, key=lambda d: len(d.name))[0]
    return None


def resolver_carpeta(base, partes):
    """Resuelve partes sucesivas; cada una puede ofrecer nombres alternativos."""
    actual = base
    for parte in partes:
        opciones = parte if isinstance(parte, (list, tuple)) else [parte]
        siguiente = None
        for opcion in opciones:
            siguiente = buscar_carpeta(actual, opcion)
            if siguiente is not None:
                break
        if siguiente is None:
            return None
        actual = siguiente
    return actual


def buscar_archivo(carpeta, patron_regex, extensiones):
    """Devuelve el archivo original más reciente que coincide con el patrón."""
    if not carpeta:
        return None
    patron = re.compile(patron_regex)
    _, archivos = leer_dir(carpeta)
    candidatas, fechas = [], {}
    for nombre, (ruta, fecha, _tamano) in archivos.items():
        if es_temporal(nombre) or ruta.suffix.lower() not in extensiones:
            continue
        if patron.search(normalizar(ruta.stem)):
            candidatas.append(ruta)
            fechas[ruta] = fecha
    if not candidatas:
        return None
    originales = [ruta for ruta in candidatas if not es_copia(ruta.stem)]
    if originales:
        candidatas = originales
    candidatas.sort(key=lambda ruta: fechas.get(ruta) or 0, reverse=True)
    return candidatas[0]


def listar_diarios(carpeta, patron_regex, extensiones):
    """Devuelve ``{fecha_AAAAMMDD: Path}`` para las planillas diarias."""
    resultado = {}
    if not carpeta:
        return resultado
    patron = re.compile(patron_regex)
    _, archivos = leer_dir(carpeta)
    for nombre, (ruta, _fecha, _tamano) in archivos.items():
        if es_temporal(nombre) or ruta.suffix.lower() not in extensiones:
            continue
        coincidencia = patron.search(normalizar(ruta.stem))
        if coincidencia:
            fecha = coincidencia.group(1)
            if es_copia(ruta.stem) and fecha in resultado:
                continue
            if fecha in resultado and es_copia(resultado[fecha].stem):
                resultado[fecha] = ruta
            else:
                resultado.setdefault(fecha, ruta)
    return resultado


def mtime(ruta):
    """Fecha de modificación, tomada del recorrido cuando el caché está activo."""
    if ruta is None:
        return None
    if _DIR_CACHE["on"]:
        _, archivos = leer_dir(Path(ruta).parent)
        dato = archivos.get(Path(ruta).name)
        if dato is not None:
            return dato[1]
    try:
        return ruta.stat().st_mtime
    except Exception:
        return None


def tamano(ruta):
    """Tamaño, tomado del recorrido cuando el caché está activo."""
    if ruta is None:
        return None
    if _DIR_CACHE["on"]:
        _, archivos = leer_dir(Path(ruta).parent)
        dato = archivos.get(Path(ruta).name)
        if dato is not None:
            return dato[2]
    try:
        return ruta.stat().st_size
    except Exception:
        return None


def fmt_fecha(ts):
    if ts is None:
        return "—"
    return datetime.fromtimestamp(ts).strftime("%d-%m-%Y %H:%M:%S")


def iguales_mtime(a, b):
    return a is not None and b is not None and abs(a - b) <= TOL_MTIME


def fmt_monto(valor):
    if valor is None:
        return "—"
    try:
        return f"{float(valor):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    except Exception:
        return str(valor)
