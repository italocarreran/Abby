"""
ActualizaRemplazos.py

Entradas (por ahora, en la carpeta desde donde se ejecuta el script):
"Reemplazos forzados.xlsx"                                      Copiar del ultimo mes reliquidado y confirmar
"datos_reuc_reemplazos_*.xlsx"                                  descargar de la pag REUC
"Reporte_SEC_*.xlsx"                                            descargar de la pag REUC (antes datos_reuc_*)
"Cuadros de Pago_Balances_SEN_*_Simplificado_def.xlsb"          de PLABACOM, el más actualizado
"1_CUADROS_PAGO_SSCC_*.xlsm"                                    del disco T, el más actualizado

Archivo destino (se selecciona en la ventana):
"0_CUADROS_RELIQUIDACIÓN SSCC_*.xlsm"  -> el del mes que estamos trabajando.

El programa:
  1. Genera el archivo de salida "Reemplazos_AAAAMMDD_SSCC.xlsx" (igual que antes).
  2. Escribe directamente en la hoja EMPRESAS del archivo destino:
        - Columnas B:C  -> EMPRESA / RUT   (desde la hoja "EMPRESAS")
        - Columnas H:I  -> Reemplazada / Reemplazante
                           (desde "Reemplazos validos" + "Reemplazos forzados")
     Ambos bloques se limpian antes de pegar, conservando el formato de las celdas.
"""

import os
import re
import json

import shutil
import socket
import subprocess
import sys

import time
import traceback
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

import xlwings as xw


# =========================================================
# CONFIG POR PC/USUARIO
# =========================================================
# La carpeta de auxiliares vive en __config__/, junto al resto de los datos
# intermedios del sistema, y es compartida por todos los usuarios. Ahi van los
# Reporte_SEC_* / datos_reuc_reemplazos_* descargados y el archivo "Reemplazos forzados". Como el .py
# suele estar en un disco compartido, el config guarda las rutas SEPARADAS POR
# PC+USUARIO (ver get_usuario), asi que cada persona conserva las suyas sin
# pisar las de los demas.
DIR_CONFIG = Path(__file__).resolve().parents[2] / "__config__"
NOMBRE_AUXILIARES = "Auxiliares REUC"
CARPETA_AUXILIARES = DIR_CONFIG / NOMBRE_AUXILIARES
# Donde estaba antes: al lado del .py. Ya no se lee ni se escribe ahi; solo se
# mira para avisar que quedaron archivos por mover (el codigo no los mueve
# solo, igual que con el resto de __config__).
CARPETA_AUXILIARES_LEGADO = Path(__file__).parent / "Auxiliares"
CONFIG_PATH = DIR_CONFIG / "reemplazos_reuc.json"

def _morir(titulo, mensaje):
    """Aborta mostrando el motivo en una ventana. Sin esto, lanzado desde el
    Revisor con pythonw (sin consola) el script moriria callado y solo se veria
    que la ventana no aparece. No puede vivir en __comun__/: es justamente lo
    que avisa cuando __comun__/ es lo que falta."""
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(titulo, mensaje)
        r.destroy()
    except Exception:
        pass
    print(f"{titulo}\n\n{mensaje}", file=sys.stderr)
    raise SystemExit(1)

# Implementaciones compartidas; los envoltorios conservan la interfaz historica.
# Va detras de _morir a proposito: lanzado por el Revisor con pythonw no hay
# consola, asi que un ImportError suelto moriria callado.
_RAIZ_COMUN = Path(__file__).resolve().parents[2]
if str(_RAIZ_COMUN) not in sys.path:
    sys.path.insert(0, str(_RAIZ_COMUN))
try:
    from __comun__ import config as _cfg
    from __comun__ import traspaso as _traspaso
    from __comun__ import texto as _texto
except ImportError as e:
    _morir("Falta la carpeta __comun__/",
           "No se pudieron cargar los modulos de __comun__/.\n\n"
           "Tiene que estar la carpeta '__comun__' hermana de Revisor_Relq.\n"
           "Baja el repositorio completo, no los .py sueltos.\n\n"
           f"Carpeta actual: {_RAIZ_COMUN}\n\nDetalle: {e}")

# La ventana chica comun de esta version (actualizadores/ventana_chica.py).
try:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "actualizadores"))
    from ventana_chica import VentanaChica
except ImportError as e:
    _morir("Falta ventana_chica.py",
           "No se pudo cargar actualizadores/ventana_chica.py.\n\n"
           "Copia la carpeta Revisor_Relq_15min completa, no los .py sueltos.\n\n"
           f"Detalle: {e}")



get_usuario = _cfg.clave_equipo


# Modo 15 minutos: reemplazos_reuc.json se comparte con la version horaria, que vive en
# la misma raiz con el mismo __comun__/ y el mismo __config__/. Para que una no
# le pise a la otra la carpeta, el ultimo mes ni las rutas recordadas, todo lo
# de esta version va dentro de la clave SECCION_CONFIG, con la MISMA forma que
# el archivo entero ({equipo_usuario: {...}, "_valores": {...}}). Lo que hay
# afuera es de la horaria y no se toca. Se resuelve aca y no en __comun__/ a
# proposito: asi sirve cualquiera de las dos copias de esa carpeta.
# El valor tiene que ser el mismo en todos los scripts de esta version (lo
# comprueba revisor/test_seccion_config.py).
SECCION_CONFIG = "modo_15min"


def _seccion_config(todo):
    """La seccion de esta version dentro del archivo completo. Si existe pero
    no es un dict no se pisa: modificar() toma la excepcion como fallo y no
    escribe."""
    sec = todo.setdefault(SECCION_CONFIG, {})
    if not isinstance(sec, dict):
        raise TypeError(f"{SECCION_CONFIG} no es un dict")
    return sec


def leer_config():
    sec = _cfg.leer_todo(CONFIG_PATH).get(SECCION_CONFIG)
    data = sec.get(get_usuario(), {}) if isinstance(sec, dict) else {}
    return data if isinstance(data, dict) else {}


def _modificar_config(mutador):
    return _cfg.modificar(CONFIG_PATH, lambda todo: mutador(_seccion_config(todo)))


def guardar_config(data):
    return _modificar_config(
        lambda todo: todo.setdefault(get_usuario(), {}).update(data))


# =========================================================
# UTILIDADES
# =========================================================
# ---------------------------------------------------------------------------
# TRASPASO DESDE EL REVISOR
# ---------------------------------------------------------------------------
# El Revisor escribe un JSON en __config__/AAAA/MM Mes/ y pasa su ruta como argv[1].
# Sin argumento este script funciona como siempre: rutas a mano y su propio
# reemplazos_reuc.json vive en __config__ y NO es el config compartido.
TRASPASO_ORIGEN = _traspaso.ORIGEN
TRASPASO_VERSION_MAX = _traspaso.VERSION_ACTUAL


def leer_traspaso(argv):
    return _traspaso.leer_argumento(argv)


def abrir_en_explorador(ruta, es_archivo=False):
    p = Path(ruta)
    if not p.exists():
        return
    if sys.platform == "win32":
        if es_archivo:
            # Abre la carpeta y deja el archivo resaltado, SIN ejecutarlo.
            subprocess.Popen(["explorer", "/select,", str(p)])
        else:
            subprocess.Popen(["explorer", str(p)])
        return
    carpeta = p.parent if es_archivo else p
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(carpeta)])
    else:
        subprocess.Popen(["xdg-open", str(carpeta)])


normalizar = _texto.suave_requerido


def buscar_archivo(ruta, patron):
    """Busca por glob y devuelve el más reciente (ignora temporales ~$)."""
    archivos = [
        p for p in Path(ruta).glob(patron)
        if p.is_file() and not p.name.startswith("~$")
    ]
    if not archivos:
        raise FileNotFoundError(f"No se encontró archivo con patrón: {patron}")
    archivos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return archivos[0]


def extraer_fecha(texto):
    """Busca una fecha tipo AAAAMMDD dentro de un string."""
    m = re.search(r"(20\d{6})", texto)
    return m.group(1) if m else "SIN_FECHA"


def extraer_fecha_aamm(texto):
    """Extrae fecha AAMM."""
    m = re.search(r"(\d{4})", texto)
    return m.group(1) if m else "SIN_FECHA"


def fecha_aaaammdd_a_ddmmaaaa(aaaammdd):
    """'20260722' -> '22-07-2026'. Si no calza el formato, devuelve tal cual."""
    if not aaaammdd or len(aaaammdd) != 8 or not aaaammdd.isdigit():
        return aaaammdd
    return f"{aaaammdd[6:8]}-{aaaammdd[4:6]}-{aaaammdd[0:4]}"


def col_letra_a_num(letra):
    n = 0
    for c in letra.upper():
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n


# =========================================================
# BÚSQUEDA EN DISCO COMPARTIDO (PLABACOM)
# =========================================================
RAIZ_PLABACOM_DEFAULT = r"T:\Facturacion\Plabacom"
# Se compara contra el nombre normalizado (minúsculas, sin tildes).
PATRON_BALANCES = r"balances_sen.*simplificado.*\.xlsb$"


def subcarpeta_que_contiene(carpeta, *fragmentos):
    """
    Devuelve la primera subcarpeta cuyo nombre normalizado contiene TODOS
    los fragmentos dados. Tolera tildes, mayúsculas y espacios de más.
    """
    if not carpeta.is_dir():
        return None
    frags = [normalizar(f) for f in fragmentos]
    for p in sorted(carpeta.iterdir()):
        if not p.is_dir():
            continue
        nom = normalizar(p.name)
        if all(f in nom for f in frags):
            return p
    return None


def meses_hacia_atras(desde=None, cantidad=36):
    """Genera tuplas (año, mes) desde el mes actual hacia atrás."""
    hoy = desde or datetime.today()
    y, m = hoy.year, hoy.month
    for _ in range(cantidad):
        yield y, m
        m -= 1
        if m == 0:
            m = 12
            y -= 1


def _prioridad_archivo(p):
    """
    Orden de preferencia dentro de una misma carpeta:
    primero los "_def", después los "_pre"; a igualdad, el más reciente.
    OJO: dentro de "02 Definitivo" el archivo igual puede llamarse "_pre".
    """
    nom = normalizar(p.name)
    es_def = 0 if re.search(r"_def\.xlsb$", nom) else 1
    return (es_def, -p.stat().st_mtime)


def archivos_que_matchean(carpeta, patron):
    return sorted(
        [
            f for f in carpeta.iterdir()
            if f.is_file()
            and not f.name.startswith("~$")
            and patron.search(normalizar(f.name))
        ],
        key=_prioridad_archivo,
    )


def archivo_en_carpeta_resultados(carpeta_version, aamm, patron_regex, log=None):
    """
    Dentro de "<version>/Publicar/01 Resultados_AAMM_.../" busca el archivo.
    Si no lo encuentra por esa ruta, hace una búsqueda recursiva de respaldo
    dentro de la carpeta de la versión. Devuelve la ruta o None.
    """
    def _log(msg):
        if log:
            log(msg)

    patron = re.compile(patron_regex, re.IGNORECASE)

    publicar = subcarpeta_que_contiene(carpeta_version, "publicar")
    if publicar is None:
        _log(f"      sin carpeta 'Publicar' en {carpeta_version.name}")
    else:
        # Carpetas "01 Resultados_AAMM_..." (puede haber más de una versión)
        candidatas = [
            p for p in publicar.iterdir()
            if p.is_dir() and "resultados" in normalizar(p.name) and aamm in p.name
        ]
        if not candidatas:
            candidatas = [
                p for p in publicar.iterdir()
                if p.is_dir() and "resultados" in normalizar(p.name)
            ]
        if not candidatas:
            _log(f"      sin carpeta 'Resultados' dentro de Publicar")
        else:
            candidatas.sort(key=lambda p: p.name, reverse=True)
            for carpeta in candidatas:
                archivos = archivos_que_matchean(carpeta, patron)
                if archivos:
                    return archivos[0]
            _log(f"      no hay .xlsb que calce en: "
                 f"{', '.join(c.name for c in candidatas[:3])}")

    # ---- Respaldo: búsqueda recursiva dentro de la versión ----
    try:
        encontrados = [
            f for f in carpeta_version.rglob("*.xlsb")
            if f.is_file()
            and not f.name.startswith("~$")
            and patron.search(normalizar(f.name))
        ]
    except Exception:
        encontrados = []

    if encontrados:
        encontrados.sort(key=_prioridad_archivo)
        _log(f"      encontrado por búsqueda recursiva: "
             f"...\\{encontrados[0].parent.name}\\{encontrados[0].name}")
        return encontrados[0]

    return None


def buscar_plabacom(raiz=RAIZ_PLABACOM_DEFAULT, patron_regex=PATRON_BALANCES,
                    meses=36, log=None):
    """
    Recorre T:\\Facturacion\\Plabacom\\<AAAA>\\<AAMM>\\<02 Definitivo | 01 Preliminar>
             \\Publicar\\01 Resultados_AAMM_...\\<archivo>

    Va del mes más reciente hacia atrás. En cada mes prueba primero
    "02 Definitivo" y luego "01 Preliminar". Devuelve (ruta, descripcion) o
    (None, mensaje de error).

    log: callback opcional para ver el detalle del recorrido.
    """
    def _log(msg):
        if log:
            log(msg)

    raiz = Path(raiz)
    if not raiz.is_dir():
        return None, f"No se puede acceder a {raiz}"

    for y, m in meses_hacia_atras(cantidad=meses):
        aamm = f"{str(y)[2:]}{m:02d}"
        carpeta_anio = raiz / str(y)
        if not carpeta_anio.is_dir():
            _log(f"  {aamm}: no existe la carpeta del año {y}")
            continue

        carpeta_mes = subcarpeta_que_contiene(carpeta_anio, aamm)
        if carpeta_mes is None:
            _log(f"  {aamm}: no existe carpeta del mes")
            continue

        _log(f"  {aamm}: revisando {carpeta_mes.name}")

        encontro_version = False
        for etiqueta, frag in (("Definitivo", "definitivo"), ("Preliminar", "preliminar")):
            carpeta_version = subcarpeta_que_contiene(carpeta_mes, frag)
            if carpeta_version is None:
                _log(f"    sin carpeta {etiqueta}")
                continue
            encontro_version = True
            archivo = archivo_en_carpeta_resultados(
                carpeta_version, aamm, patron_regex, log=_log
            )
            if archivo:
                _log(f"    OK -> {etiqueta}: {archivo.name}")
                return archivo, f"{aamm} / {etiqueta}"
            _log(f"    {etiqueta}: sin archivo")

        if not encontro_version:
            _log(f"    subcarpetas: "
                 f"{', '.join(p.name for p in carpeta_mes.iterdir() if p.is_dir())[:120]}")

    return None, f"No se encontró el archivo en los últimos {meses} meses"


# =========================================================
# BÚSQUEDA CONJUNTA: BALANCES SEN + 1_CUADROS_PAGO_SSCC
# =========================================================
RAIZ_FACTURACION_DEFAULT = r"T:\Facturacion"
NARANJO = "#d97706"   # aviso: el archivo no calza con su carpeta
PATRON_SSCC = r"1_cuadros_pago_sscc.*\.xlsm$"

MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


def carpeta_mes_facturacion(raiz_facturacion, y, m):
    """
    T:\\Facturacion\\<AAAA>\\<MM Mes>   (ej: 2026\\06 Junio)
    Acepta que la carpeta se llame por número, por nombre, o ambos.
    """
    carpeta_anio = Path(raiz_facturacion) / str(y)
    if not carpeta_anio.is_dir():
        return None

    nombre_mes = MESES_ES[m - 1]
    mm = f"{m:02d}"
    candidatas = []
    for p in sorted(carpeta_anio.iterdir()):
        if not p.is_dir():
            continue
        nom = normalizar(p.name)
        if nombre_mes in nom or re.match(rf"^{mm}\b", nom):
            candidatas.append(p)
    # Preferir la que tenga el nombre del mes (más específica)
    candidatas.sort(key=lambda p: 0 if nombre_mes in normalizar(p.name) else 1)
    return candidatas[0] if candidatas else None


def buscar_sscc_en_version(carpeta_version, patron_regex=PATRON_SSCC, log=None):
    """Dentro de "<version>/SSCC/" busca 1_CUADROS_PAGO_SSCC_*.xlsm."""
    def _log(msg):
        if log:
            log(msg)

    patron = re.compile(patron_regex, re.IGNORECASE)

    carpeta_sscc = subcarpeta_que_contiene(carpeta_version, "sscc")
    if carpeta_sscc is not None:
        archivos = archivos_que_matchean(carpeta_sscc, patron)
        if archivos:
            return archivos[0]
        _log(f"      carpeta SSCC sin 1_CUADROS_PAGO_SSCC")
    else:
        _log(f"      sin carpeta 'SSCC' en {carpeta_version.name}")

    # Respaldo: búsqueda recursiva dentro de la versión
    try:
        encontrados = [
            f for f in carpeta_version.rglob("*.xlsm")
            if f.is_file()
            and not f.name.startswith("~$")
            and patron.search(normalizar(f.name))
        ]
    except Exception:
        encontrados = []

    if encontrados:
        encontrados.sort(key=_prioridad_archivo)
        _log(f"      SSCC encontrado por búsqueda recursiva: {encontrados[0].name}")
        return encontrados[0]

    return None


# ---------- Verificación de coherencia nombre de archivo vs carpeta ----------
MESES_ABR = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
}


def aamm_desde_nombre(nombre):
    """
    Intenta sacar el AAMM del nombre del archivo. Reconoce dos formatos:
      - "..._2606_..."          -> 2606
      - "..._jun26_..."         -> 2606
    Devuelve None si no logra determinarlo.
    """
    nom = normalizar(nombre)

    m = re.search(r"(?<!\d)(\d{2})(0[1-9]|1[0-2])(?!\d)", nom)
    if m:
        return m.group(1) + m.group(2)

    m = re.search(r"(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]*[_\-\s]?(\d{2})(?!\d)", nom)
    if m:
        return f"{m.group(2)}{MESES_ABR[m.group(1)]:02d}"

    return None


def sufijo_desde_nombre(nombre):
    """Devuelve 'def', 'pre' o None según el sufijo del archivo."""
    nom = normalizar(nombre)
    m = re.search(r"[_\-\s](def|pre)(?:initivo|liminar)?\.\w+$", nom)
    return m.group(1) if m else None


def verificar_coherencia(ruta, aamm_carpeta, version_carpeta):
    """
    Compara el nombre del archivo contra la carpeta donde está.
    Devuelve una lista de avisos (vacía si todo calza).
    """
    avisos = []
    if ruta is None:
        return avisos

    aamm_arch = aamm_desde_nombre(ruta.name)
    if aamm_arch is None:
        avisos.append(f"no se pudo leer el mes en el nombre de {ruta.name}")
    elif aamm_arch != aamm_carpeta:
        avisos.append(
            f"{ruta.name}: el mes del nombre ({aamm_arch}) "
            f"no calza con la carpeta ({aamm_carpeta})"
        )

    esperado = "def" if normalizar(version_carpeta).startswith("defin") else "pre"
    sufijo = sufijo_desde_nombre(ruta.name)
    if sufijo is None:
        avisos.append(f"{ruta.name}: sin sufijo _def/_pre en el nombre")
    elif sufijo != esperado:
        avisos.append(
            f"{ruta.name}: el nombre dice _{sufijo} pero está en {version_carpeta}"
        )

    return avisos


def buscar_par_mensual(raiz_facturacion=RAIZ_FACTURACION_DEFAULT, meses=36, log=None):
    """
    Busca los DOS archivos exigiendo que vengan del MISMO mes y del MISMO
    proceso (Definitivo o Preliminar):

        Balances SEN : <raiz>\\Plabacom\\<AAAA>\\<AAMM>\\<version>\\Publicar\\01 Resultados_...
        SSCC         : <raiz>\\<AAAA>\\<MM Mes>\\<version>\\SSCC

    Recorre del mes más reciente hacia atrás. En cada mes prueba Definitivo y
    después Preliminar; solo acepta la combinación si ENCUENTRA LOS DOS.

    Devuelve (dict, info) o (None, mensaje).
    """
    def _log(msg):
        if log:
            log(msg)

    raiz = Path(raiz_facturacion)
    if not raiz.is_dir():
        return None, f"No se puede acceder a {raiz}"

    raiz_plabacom = raiz / "Plabacom"
    if not raiz_plabacom.is_dir():
        return None, f"No se puede acceder a {raiz_plabacom}"

    patron_bal = re.compile(PATRON_BALANCES, re.IGNORECASE)

    for y, m in meses_hacia_atras(cantidad=meses):
        aamm = f"{str(y)[2:]}{m:02d}"

        carpeta_anio_plab = raiz_plabacom / str(y)
        carpeta_mes_plab = (
            subcarpeta_que_contiene(carpeta_anio_plab, aamm)
            if carpeta_anio_plab.is_dir() else None
        )
        carpeta_mes_fact = carpeta_mes_facturacion(raiz, y, m)

        if carpeta_mes_plab is None and carpeta_mes_fact is None:
            continue

        _log(f"  {aamm}: "
             f"Plabacom={carpeta_mes_plab.name if carpeta_mes_plab else '—'} | "
             f"Facturacion={carpeta_mes_fact.name if carpeta_mes_fact else '—'}")

        for etiqueta, frag in (("Definitivo", "definitivo"), ("Preliminar", "preliminar")):
            ver_plab = (
                subcarpeta_que_contiene(carpeta_mes_plab, frag)
                if carpeta_mes_plab else None
            )
            ver_fact = (
                subcarpeta_que_contiene(carpeta_mes_fact, frag)
                if carpeta_mes_fact else None
            )

            xlsb = (
                archivo_en_carpeta_resultados(ver_plab, aamm, PATRON_BALANCES, log=_log)
                if ver_plab else None
            )
            xlsm = buscar_sscc_en_version(ver_fact, log=_log) if ver_fact else None

            if xlsb and xlsm:
                avisos_xlsb = verificar_coherencia(xlsb, aamm, etiqueta)
                avisos_xlsm = verificar_coherencia(xlsm, aamm, etiqueta)
                _log(f"    OK {etiqueta}: {xlsb.name} + {xlsm.name}")
                for a in avisos_xlsb + avisos_xlsm:
                    _log(f"    OJO: {a}")
                return (
                    {
                        "xlsb": xlsb,
                        "xlsm": xlsm,
                        "aamm": aamm,
                        "version": etiqueta,
                        "avisos_xlsb": avisos_xlsb,
                        "avisos_xlsm": avisos_xlsm,
                    },
                    f"{aamm} / {etiqueta}",
                )

            faltan = []
            if not xlsb:
                faltan.append("Balances SEN")
            if not xlsm:
                faltan.append("SSCC")
            if ver_plab or ver_fact:
                _log(f"    {etiqueta}: falta {' y '.join(faltan)}")

    return None, f"No se encontró un mes con AMBOS archivos (últimos {meses} meses)"


# =========================================================
# DESCARGA DESDE REUC (con el navegador de siempre)
# =========================================================
URL_REUC_EMPRESAS = (
    "https://reuc.coordinador.cl/maestro_usuarios/empresas/exportar_reuc?&text_search="
)
URL_REUC_REEMPLAZADAS = (
    "https://reuc.coordinador.cl/maestro_usuarios/empresas/"
    "export_reemplazadas_data?&text_search="
)

def carpeta_auxiliares():
    """Carpeta ``__config__/Auxiliares REUC``. La crea si no existe."""
    CARPETA_AUXILIARES.mkdir(parents=True, exist_ok=True)
    return CARPETA_AUXILIARES


# Lo que la carpeta de auxiliares tiene que contener; sirve para saber si en la
# carpeta vieja quedo algo de verdad y no solo basura suelta.
PATRONES_AUXILIARES = ("Reporte_SEC*.xlsx", "datos_reuc*.xlsx", "Reemplazos forzados*.xlsx")


def carpeta_datos_guardada(cfg):
    """Carpeta de datos del config, corrigiendo la que apunta a la carpeta vieja.

    Esto es lo unico que se migra solo: un texto en el JSON de config. Los
    archivos los mueve la persona (ver ``avisar_legado``).
    """
    guardada = str(cfg.get("carpeta_datos") or "").strip()
    if not guardada:
        return str(CARPETA_AUXILIARES)
    try:
        if Path(guardada) == CARPETA_AUXILIARES_LEGADO:
            return str(CARPETA_AUXILIARES)
    except (OSError, ValueError):
        pass
    return guardada


def pendientes_en_legado():
    """Archivos de auxiliares que quedaron en la carpeta vieja, al lado del .py."""
    if not CARPETA_AUXILIARES_LEGADO.is_dir():
        return []
    vistos = {}
    for patron in PATRONES_AUXILIARES:
        for f in CARPETA_AUXILIARES_LEGADO.glob(patron):
            if not f.name.startswith("~$"):
                vistos[f.name] = f
    return [vistos[n] for n in sorted(vistos)]


def avisar_legado(log=print):
    """Avisa que hay auxiliares en la carpeta vieja, al lado del .py.

    No se mueve nada solo: es la misma regla que el resto de ``__config__``,
    donde la migracion la hace la persona. Pero callarse tampoco sirve, porque
    desde afuera solo se ve que "no encuentra los archivos".
    """
    pendientes = pendientes_en_legado()
    if not pendientes:
        return False
    log(f"  !!! Quedaron {len(pendientes)} archivo(s) de auxiliares en la carpeta "
        f"vieja: {CARPETA_AUXILIARES_LEGADO}")
    log(f"      Ahora se leen de: {CARPETA_AUXILIARES}")
    log("      Muevelos a mano (el script no los mueve solo): "
        + ", ".join(f.name for f in pendientes[:6])
        + (" ..." if len(pendientes) > 6 else ""))
    return True


def buscar_archivo_con_respaldo(carpeta_preferida, patron, log=print):
    """
    Busca primero en carpeta_preferida (la que eligio el usuario o quedo
    guardada en config). Si ahi no esta, cae de respaldo a Auxiliares REUC
    -- que es donde SIEMPRE deberia estar-- y avisa por log cual de las
    dos uso, para que quede claro y no parezca que "no encuentra nada".
    """
    aux = carpeta_auxiliares()
    carpeta_preferida = Path(carpeta_preferida) if carpeta_preferida else aux

    try:
        return buscar_archivo(carpeta_preferida, patron)
    except FileNotFoundError:
        pass

    if carpeta_preferida.resolve() != aux.resolve():
        try:
            encontrado = buscar_archivo(aux, patron)
            log(f"    (no estaba en {carpeta_preferida}, se uso {NOMBRE_AUXILIARES}: "
                f"{encontrado.name})")
            return encontrado
        except FileNotFoundError:
            pass

    avisar_legado(log)
    raise FileNotFoundError(
        f"No se encontro '{patron}' ni en {carpeta_preferida} ni en {aux}"
    )


# ---------------------------------------------------------
# La descarga la hace TU navegador de siempre (no uno automatizado).
# Con Playwright, Cloudflare detectaba el navegador automatizado y la
# verificacion fallaba siempre (usuaria, 2026-10-29). Ahora el script solo
# abre los links en el navegador por defecto y mira la carpeta Descargas:
# cuando aparecen Reporte_SEC* y datos_reuc_reemplazos*, los mueve a
# Auxiliares REUC.
# ---------------------------------------------------------
URL_REUC_LOGIN = "https://reuc.coordinador.cl/login/"

# GUID de la carpeta "Descargas" en el registro de Windows (puede estar movida).
_GUID_DESCARGAS = "{374DE290-123F-4565-9164-39C4925E467B}"

# Lo que deja el navegador mientras el archivo todavia se esta bajando.
_EXT_PARCIALES = (".crdownload", ".partial", ".part", ".tmp", ".download")


def carpeta_descargas():
    """Carpeta Descargas del usuario (la del registro si se movio)."""
    if sys.platform.startswith("win"):
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            ) as k:
                valor, _ = winreg.QueryValueEx(k, _GUID_DESCARGAS)
            p = Path(os.path.expandvars(valor))
            if p.is_dir():
                return p
        except OSError:
            pass
    return Path.home() / "Downloads"


def clase_descarga(nombre):
    """'reuc', 'reemplazos' o None segun el nombre del archivo bajado.

    Acepta los nombres que pone el navegador cuando ya habia uno igual,
    como "Reporte_SEC (1).xlsx". El export de empresas se llama Reporte_SEC_*
    (antes datos_reuc_*, usuaria 2026-10-29); el de reemplazos sigue siendo
    datos_reuc_reemplazos_*.
    """
    n = nombre.lower()
    if n.startswith("~$") or not n.endswith(".xlsx"):
        return None
    if n.startswith("reporte_sec"):
        return "reuc"
    if n.startswith("datos_reuc") and "reemplaz" in n:
        return "reemplazos"
    return None


def descargas_nuevas(carpeta, desde):
    """{clase: Path} con el archivo mas nuevo de cada clase modificado desde 'desde'."""
    encontrados = {}
    try:
        archivos = list(Path(carpeta).iterdir())
    except OSError:
        return encontrados
    for f in archivos:
        if f.name.lower().endswith(_EXT_PARCIALES):
            continue
        clase = clase_descarga(f.name)
        if not clase:
            continue
        try:
            st = f.stat()
        except OSError:
            continue
        if st.st_mtime < desde or st.st_size == 0:
            continue
        actual = encontrados.get(clase)
        if actual is None or st.st_mtime > actual[1]:
            encontrados[clase] = (f, st.st_mtime)
    return {c: f for c, (f, _) in encontrados.items()}


def _borrar_viejos(carpeta_aux, patron_glob, excluir=None, log=print):
    """
    Borra archivos previos que calcen con patron_glob, salvo los que
    contengan 'excluir' en el nombre (para no borrar datos_reuc_reemplazos_*
    al limpiar datos_reuc_*). Evita que la carpeta de auxiliares se llene
    de versiones viejas con timestamp distinto.
    """
    borrados = []
    for f in Path(carpeta_aux).glob(patron_glob):
        if excluir and excluir in f.name.lower():
            continue
        try:
            f.unlink()
            borrados.append(f.name)
        except Exception as e:
            log(f"    (no se pudo borrar {f.name}: {e})")
    return borrados


# clase -> (prefijo del nombre final, patrones a borrar, fragmento que no se borra)
# Al bajar un Reporte_SEC se borran tambien los datos_reuc_* viejos (el nombre
# anterior del mismo export), pero nunca los datos_reuc_reemplazos_*.
_DESTINOS_REUC = {
    "reuc": ("Reporte_SEC", ("Reporte_SEC*.xlsx", "datos_reuc*.xlsx"), "reemplazos"),
    "reemplazos": ("datos_reuc_reemplazos", ("datos_reuc_reemplazos*.xlsx",), None),
}


def mover_a_auxiliares(origen, clase, carpeta_aux, log=print):
    """Mueve una descarga a auxiliares, borrando antes las versiones viejas.

    Si el navegador le puso " (1)" o similar, se renombra con la fecha para
    que procesar_datos lo encuentre como siempre.
    """
    prefijo, patrones, excluir = _DESTINOS_REUC[clase]
    nombre = origen.name
    if not re.fullmatch(re.escape(prefijo) + r"[\w\-]*\.xlsx", nombre, re.I) or (
            clase == "reuc" and "reemplaz" in nombre.lower()):
        nombre = f"{prefijo}_{datetime.today().strftime('%Y%m%d_%H%M%S')}.xlsx"
    for patron in patrones:
        for v in _borrar_viejos(carpeta_aux, patron, excluir, log=log):
            log(f"    Reemplazado (borrado): {v}")
    destino = Path(carpeta_aux) / nombre
    shutil.move(str(origen), str(destino))
    return destino


def esperar_descargas(carpeta_bajadas, carpeta_aux, desde, espera_seg=300,
                      log=print, pausa=1.0, reloj=time.time, dormir=time.sleep):
    """Espera los dos exports en Descargas y los mueve a auxiliares.

    Un archivo se da por terminado cuando su tamano no cambia entre dos
    revisiones seguidas. Si se acaba el tiempo, lo que llego igual se mueve y
    se avisa cual falta.
    """
    t_ini = reloj()
    tamanos = {}
    listos = {}
    t_aviso = t_ini
    while len(listos) < len(_DESTINOS_REUC):
        for clase, f in descargas_nuevas(carpeta_bajadas, desde).items():
            if clase in listos:
                continue
            try:
                tam = f.stat().st_size
            except OSError:
                continue
            if tamanos.get(clase) == (f, tam):
                listos[clase] = mover_a_auxiliares(f, clase, carpeta_aux, log=log)
                log(f"    Llego {f.name} -> {listos[clase].name}")
            else:
                tamanos[clase] = (f, tam)
        if len(listos) == len(_DESTINOS_REUC):
            break
        if reloj() - t_ini > espera_seg:
            faltan = [_DESTINOS_REUC[c][0] for c in _DESTINOS_REUC if c not in listos]
            raise TimeoutError(
                f"No llego {' ni '.join(faltan)} a {carpeta_bajadas} en "
                f"{espera_seg // 60} min. Si el navegador lo guardo en otra carpeta, "
                f"muevelo a mano a {carpeta_aux}."
            )
        if reloj() - t_aviso > 30:
            log("    (esperando las descargas en el navegador...)")
            t_aviso = reloj()
        dormir(pausa)
    return listos


def descargar_reuc(carpeta_destino=None, log=print, espera_descarga_seg=300,
                   carpeta_bajadas=None, abrir=None):
    """
    Pide los dos exports al navegador por defecto y los trae a auxiliares:
      - Reporte_SEC_*.xlsx              (exportar_reuc)
      - datos_reuc_reemplazos_*.xlsx    (export_reemplazadas_data)

    Antes de llamarla la persona ya tiene que haber iniciado sesion en REUC
    en ese mismo navegador (la ventana se lo pide). El script no ve ni guarda
    la sesion ni la clave: todo queda en el navegador.

    Devuelve un dict {"reuc": Path, "reemplazos": Path}.
    """
    import webbrowser
    abrir = abrir or webbrowser.open
    carpeta_aux = Path(carpeta_destino) if carpeta_destino else CARPETA_AUXILIARES
    carpeta_aux.mkdir(parents=True, exist_ok=True)
    bajadas = Path(carpeta_bajadas) if carpeta_bajadas else carpeta_descargas()

    desde = time.time() - 5     # margen por relojes de archivo redondeados
    log(f"  Mirando la carpeta de descargas: {bajadas}")
    for url in (URL_REUC_EMPRESAS, URL_REUC_REEMPLAZADAS):
        abrir(url)
    log("  Se pidieron los dos exports en tu navegador "
        "(el servidor los genera, puede demorar).")
    return esperar_descargas(bajadas, carpeta_aux, desde, espera_descarga_seg, log=log)


COLUMNAS_REEMPLAZOS = ("Rut", "Rut Reemplazante", "Empresa", "Reemplazada Por",
                       "Fin de Reemplazo")


def hoja_con_columnas(hojas, columnas, nombre_archivo=""):
    """De {nombre_hoja: DataFrame}, la primera hoja que tenga todas las columnas.

    Si ninguna las tiene, el error dice que hojas y columnas trae el archivo,
    para que se vea al tiro si el REUC cambio el formato.
    """
    for df in hojas.values():
        if all(c in df.columns for c in columnas):
            return df
    detalle = "; ".join(f"{h}: {', '.join(map(str, df.columns))}"
                        for h, df in hojas.items())
    raise ValueError(
        f"{nombre_archivo}: ninguna hoja tiene las columnas "
        f"{', '.join(columnas)}. Trae -> {detalle}"
    )


def ultima_fila(sheet, col_letra, desde):
    """Última fila con contenido en una columna (considera fórmulas)."""
    col_num = col_letra_a_num(col_letra)
    ultima = sheet.cells.last_cell.row
    rng = sheet.range((desde, col_num), (ultima, col_num))
    formulas = rng.formula
    if formulas is None:
        return desde - 1
    if not isinstance(formulas, (list, tuple)):
        formulas = [[formulas]]
    elif not isinstance(formulas[0], (list, tuple)):
        formulas = [[v] for v in formulas]
    fila = desde - 1
    for i, row in enumerate(formulas):
        if row and row[0] not in (None, ""):
            fila = desde + i
    return fila


# =========================================================
# PROCESO DE DATOS (lo que hacía el script original)
# =========================================================
def procesar_datos(carpeta_datos, archivo_cuadros,
                   archivo_xlsb=None, archivo_xlsm=None, log=print):
    """
    Devuelve un dict con los DataFrames que se usan tanto para el archivo
    de salida como para escribir en el archivo destino.

    archivo_xlsb / archivo_xlsm: rutas explícitas (disco compartido). Si vienen
    en None, se buscan en la carpeta de datos como respaldo.
    """
    # Los auxiliares (Reporte_SEC, datos_reuc_reemplazos, reemplazos forzados) viven SIEMPRE en
    # __config__/Auxiliares REUC, compartida por todos.
    avisar_legado(log)
    carpeta_datos = Path(carpeta_datos) if carpeta_datos else carpeta_auxiliares()

    # ---------- Archivos base ----------
    if archivo_xlsm:
        archivo_xlsm = Path(archivo_xlsm)
    else:
        archivo_xlsm = buscar_archivo(carpeta_datos, "1_CUADROS_PAGO_SSCC_*.xlsm")

    if archivo_xlsb:
        archivo_xlsb = Path(archivo_xlsb)
    else:
        archivo_xlsb = buscar_archivo(
            carpeta_datos, "Cuadros de Pago_Balances_SEN_*_Simplificado_def.xlsb"
        )

    log(f"  1_CUADROS_PAGO_SSCC : {archivo_xlsm.name}")
    log(f"  Balances SEN        : {archivo_xlsb.name}")

    # ---------- Hoja EMPRESAS (columnas B y C) ----------
    df_xlsm = pd.read_excel(
        archivo_xlsm, sheet_name="EMPRESAS", usecols="B:C", engine="openpyxl"
    )
    df_xlsb = pd.read_excel(
        archivo_xlsb, sheet_name="EMPRESAS", usecols="B:C", engine="pyxlsb"
    )

    df_xlsm.columns = ["EMPRESA", "RUT"]
    df_xlsb.columns = ["EMPRESA", "RUT"]

    for df in [df_xlsm, df_xlsb]:
        df["EMPRESA"] = df["EMPRESA"].astype(str).str.strip().str.upper()
        df["RUT"] = df["RUT"].astype(str).str.strip().str.upper()

    df_unido = pd.concat([df_xlsb, df_xlsm], ignore_index=True)
    df_unido = df_unido.drop_duplicates(subset=["EMPRESA"])

    df_unido = df_unido[
        ~df_unido["RUT"].isin(["DESACTIVADO", "REEMPLAZADA", ""])
        & df_unido["RUT"].notna()
    ]
    log(f"  Empresas con RUT válido: {len(df_unido)}")

    # ---------- Reemplazos REUC ----------
    archivo_reemplazos = buscar_archivo_con_respaldo(
        carpeta_datos, "datos_reuc_reemplazos_*.xlsx", log=log
    )
    log(f"  Reemplazos REUC     : {archivo_reemplazos.name}")

    # La hoja se llamaba "pyexcel_sheet1"; desde la descarga con el navegador
    # de siempre trae otro nombre (usuaria, 2026-10-29). Se toma la hoja que
    # tenga las columnas, se llame como se llame.
    df_reemplazos = hoja_con_columnas(
        pd.read_excel(archivo_reemplazos, sheet_name=None, usecols="A:H",
                      engine="openpyxl"),
        COLUMNAS_REEMPLAZOS, archivo_reemplazos.name,
    )

    for col in ["Rut", "Rut Reemplazante", "Empresa", "Reemplazada Por"]:
        df_reemplazos[col] = df_reemplazos[col].astype(str).str.strip().str.upper()

    log(f"  Reemplazos leídos   : {len(df_reemplazos)}")

    df_reemplazos_vigentes = df_reemplazos[df_reemplazos["Fin de Reemplazo"].isna()]
    log(f"  Reemplazos vigentes : {len(df_reemplazos_vigentes)}")

    ruts_validos = set(df_unido["RUT"])
    df_reemplazos_vigentes = df_reemplazos_vigentes[
        df_reemplazos_vigentes["Rut"].isin(ruts_validos)
    ]
    df_reemplazos_vigentes = df_reemplazos_vigentes[
        df_reemplazos_vigentes["Rut Reemplazante"].isin(ruts_validos)
    ]
    log(f"  Reemplazos válidos  : {len(df_reemplazos_vigentes)}")

    df_reemplazos_final = df_reemplazos_vigentes[
        ["Rut", "Rut Reemplazante", "Empresa", "Reemplazada Por"]
    ].rename(columns={"Empresa": "Reemplazada", "Reemplazada Por": "Reemplazante"})

    # ---------- REUC ----------
    def _buscar_reuc_sin_reemplazos(carpeta):
        candidatos = [
            p for p in Path(carpeta).iterdir()
            if p.name.lower().startswith("reporte_sec") and p.suffix.lower() == ".xlsx"
        ]
        if not candidatos:
            raise FileNotFoundError(f"No hay Reporte_SEC_*.xlsx en {carpeta}")
        candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return candidatos[0]

    aux_reuc = carpeta_auxiliares()
    carpeta_datos_p = Path(carpeta_datos) if carpeta_datos else aux_reuc
    try:
        archivo_reuc = _buscar_reuc_sin_reemplazos(carpeta_datos_p)
    except FileNotFoundError:
        if carpeta_datos_p.resolve() == aux_reuc.resolve():
            raise
        archivo_reuc = _buscar_reuc_sin_reemplazos(aux_reuc)
        log(f"    (no estaba en {carpeta_datos_p}, se uso {NOMBRE_AUXILIARES}: "
            f"{archivo_reuc.name})")
    log(f"  REUC                : {archivo_reuc.name}")

    df_reuc = pd.read_excel(archivo_reuc, usecols="A,C", engine="openpyxl")
    df_reuc.columns = ["RAZON_SOCIAL", "RUT"]
    df_reuc["RAZON_SOCIAL"] = df_reuc["RAZON_SOCIAL"].astype(str).str.strip().str.upper()
    df_reuc["RUT"] = df_reuc["RUT"].astype(str).str.strip().str.upper()

    # ---------- Planilla 0 (archivo destino, solo lectura acá) ----------
    df_cuadros = pd.read_excel(
        archivo_cuadros,
        sheet_name=2,  # hay hojas ocultas
        usecols="R,U",
        skiprows=3,
        engine="openpyxl",
    )
    df_cuadros.columns = ["EMPRESA", "CLP"]
    df_cuadros["EMPRESA"] = df_cuadros["EMPRESA"].astype(str).str.strip().str.upper()

    # ---------- Reemplazos forzados: PRIMERA pasada ----------
    # Se leen y se aplican ANTES de buscar el RUT, y este orden es el punto.
    # Un reemplazo forzado existe justamente para arreglar un nombre que no
    # coincide con el registro de empresas ("CAMELLO ALTO" en el cuadro contra
    # "CAMELLO_ALTO" en EMPRESAS). Si se aplicara despues del cruce, esa empresa
    # se quedaria sin RUT, la alerta la seguiria denunciando mes a mes y encima
    # la fila se descartaba antes de llegar al reemplazo, asi que el forzado no
    # tenia ninguna chance de arreglarla.
    archivo_forzados = buscar_archivo_con_respaldo(
        carpeta_datos, "Reemplazos forzados*.xlsx", log=log
    )
    log(f"  Reemplazos forzados : {archivo_forzados.name}")

    df_forzados = pd.read_excel(
        archivo_forzados, sheet_name="Reemplazos forzados", engine="openpyxl"
    )
    df_forzados["Reemplazada"] = df_forzados["Reemplazada"].astype(str).str.strip().str.upper()
    df_forzados["Reemplazante"] = df_forzados["Reemplazante"].astype(str).str.strip().str.upper()

    mapa_forzado_empresa = dict(
        zip(df_forzados["Reemplazada"], df_forzados["Reemplazante"])
    )

    # Aviso de cadenas: si un Reemplazante es a su vez una Reemplazada, el mapa
    # se aplica dos veces (aca y mas abajo) y el nombre saltaria dos escalones.
    cadenas = sorted(set(df_forzados["Reemplazante"]) & set(df_forzados["Reemplazada"]))
    if cadenas:
        log("  OJO: hay reemplazos forzados en cadena, revisa la planilla:")
        for nombre in cadenas[:10]:
            log(f"        \"{nombre}\" es Reemplazante y tambien Reemplazada")

    antes = df_cuadros["EMPRESA"].copy()
    df_cuadros["EMPRESA"] = (
        df_cuadros["EMPRESA"].map(mapa_forzado_empresa).fillna(df_cuadros["EMPRESA"])
    )
    cambiadas = antes[antes != df_cuadros["EMPRESA"]]
    if len(cambiadas) > 0:
        pares = sorted(set(zip(cambiadas, df_cuadros.loc[cambiadas.index, "EMPRESA"])))
        log(f"  Reemplazos forzados aplicados al nombre del cuadro: {len(pares)}")
        for viejo, nuevo in pares[:15]:
            log(f"        {viejo}  ->  {nuevo}")
        if len(pares) > 15:
            log(f"        ... y {len(pares) - 15} mas")

    df_con_rut = df_cuadros.merge(
        df_unido[["EMPRESA", "RUT"]], on="EMPRESA", how="left"
    )

    df_empresas_con_rut = df_con_rut[df_con_rut["RUT"].notna()]
    df_empresas_sin_rut = df_con_rut[df_con_rut["RUT"].isna()]

    # La alerta se calcula DESPUES de los forzados, asi que ya no denuncia a las
    # empresas que el forzado arreglo. Las que sigan apareciendo aca es porque de
    # verdad les falta el RUT o el nombre todavia no esta mapeado.
    df_sin_rut_a_pago = df_empresas_sin_rut[
        (df_empresas_sin_rut["CLP"].notna()) & (df_empresas_sin_rut["CLP"] != 0)
    ]
    if len(df_sin_rut_a_pago) > 0:
        faltan = sorted(set(df_sin_rut_a_pago["EMPRESA"]))
        log(f"  Empresas SIN RUT que van a pago: {len(faltan)}")
        for nombre in faltan[:15]:
            log(f"        {nombre}")
        if len(faltan) > 15:
            log(f"        ... y {len(faltan) - 15} mas")

    df_con_rut = df_empresas_con_rut

    # ---------- Reemplazos automáticos (REUC) ----------
    df_final = df_con_rut.copy()

    mapa_rut_reemplazo = dict(
        zip(df_reemplazos_final["Rut"], df_reemplazos_final["Rut Reemplazante"])
    )
    mapa_empresa_reemplazo = dict(
        zip(df_reemplazos_final["Rut"], df_reemplazos_final["Reemplazante"])
    )

    df_final["RUT_FINAL"] = df_final["RUT"].map(mapa_rut_reemplazo).fillna(df_final["RUT"])
    df_final["EMPRESA_FINAL"] = (
        df_final["RUT"].map(mapa_empresa_reemplazo).fillna(df_final["EMPRESA"])
    )

    # ---------- Reemplazos forzados (segunda pasada) ----------
    # La primera pasada ya se hizo ARRIBA, sobre el nombre que viene del cuadro,
    # antes de buscar el RUT. Esta segunda corrige el nombre que quedo despues de
    # aplicar los reemplazos del REUC, que puede ser otro.
    df_final["EMPRESA_FINAL"] = (
        df_final["EMPRESA_FINAL"].map(mapa_forzado_empresa).fillna(df_final["EMPRESA_FINAL"])
    )

    # ---------- RUT final según empresa final ----------
    df_final = df_final.merge(
        df_unido[["EMPRESA", "RUT"]].rename(
            columns={"EMPRESA": "EMPRESA_FINAL", "RUT": "RUT_FINAL_TMP"}
        ),
        on="EMPRESA_FINAL",
        how="left",
    )
    df_final["RUT_FINAL"] = df_final["RUT_FINAL_TMP"].fillna(df_final["RUT_FINAL"])
    df_final = df_final.drop(columns=["RUT_FINAL_TMP"])

    # ---------- Validar contra REUC ----------
    archivo_reuc_df = pd.read_excel(archivo_reuc, engine="openpyxl")
    archivo_reuc_df["Rut"] = archivo_reuc_df["Rut"].astype(str).str.strip().str.upper()
    ruts_reuc = set(archivo_reuc_df["Rut"])

    df_final["EN_REUC"] = df_final["RUT_FINAL"].isin(ruts_reuc).astype(int)

    # ---------- Tabla final ----------
    df_salida = df_final[["EMPRESA_FINAL", "RUT_FINAL", "CLP", "EN_REUC"]].rename(
        columns={"EMPRESA_FINAL": "EMPRESA", "RUT_FINAL": "RUT"}
    )
    df_salida = df_salida[~((df_salida["CLP"] == 0) & (df_salida["EN_REUC"] == 0))]
    df_salida = (
        df_salida.groupby(["EMPRESA", "RUT", "EN_REUC"], as_index=False)
        .agg({"CLP": "sum"})
    )

    # ---------- Reemplazos forzados con RUT ----------
    df_forzados_actualizado = df_forzados.merge(
        df_unido.rename(columns={"EMPRESA": "Reemplazante", "RUT": "RUT Reemplazante"}),
        on="Reemplazante",
        how="left",
    )

    return {
        "df_reemplazos_final": df_reemplazos_final,
        "df_salida": df_salida,
        "df_forzados": df_forzados,
        "df_forzados_actualizado": df_forzados_actualizado,
        "df_reuc": df_reuc,
        "df_unido": df_unido,
        "df_sin_rut_a_pago": df_sin_rut_a_pago,
        "fuentes": [archivo_xlsm, archivo_xlsb, archivo_reuc,
                    archivo_reemplazos, archivo_forzados],
        "fecha_reuc": extraer_fecha(str(archivo_reuc)),
        "fecha_empresas": extraer_fecha_aamm(str(archivo_xlsm)),
    }


# =========================================================
# ARCHIVO DE SALIDA
# =========================================================
def carpeta_reemplazos_reuc(ruta_destino):
    """
    Dos niveles por encima de la carpeta del "0_CUADROS_RELIQUIDACION":

        ...\\SSCC\\02 CASO RELIQUIDACION\\00 Entregables\\0_CUADROS_...xlsm
             ^-- aca se crea "Reemplazos REUC"

    O sea: <carpeta del archivo>.parents[2] / "Reemplazos REUC".
    """
    destino = Path(ruta_destino).resolve()
    try:
        carpeta_sscc = destino.parents[2]        # 00 Entregables -> 02 CASO... -> SSCC
    except IndexError:
        carpeta_sscc = destino.parent
    carpeta = carpeta_sscc / "Reemplazos REUC"
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def respaldar_fuentes(datos, carpeta, log=print):
    """
    Copia a "Reemplazos REUC" los archivos que se usaron para armar el
    resultado, para que quede todo junto y trazable. Si ya existen, se
    reemplazan.
    """
    copiados = []
    for origen in datos.get("fuentes", []):
        if not origen:
            continue
        origen = Path(origen)
        if not origen.is_file():
            log(f"    (no se encontro para copiar: {origen.name})")
            continue
        try:
            shutil.copy2(origen, carpeta / origen.name)
            copiados.append(origen.name)
        except Exception as e:
            log(f"    (no se pudo copiar {origen.name}: {e})")

    for nombre in copiados:
        log(f"    Copiado: {nombre}")
    return copiados


def generar_archivo_salida(datos, carpeta_salida, log=print):
    fecha_hoy = datetime.today().strftime("%Y%m%d")
    nombre_salida = Path(carpeta_salida) / f"Reemplazos_{fecha_hoy}_SSCC.xlsx"

    with pd.ExcelWriter(nombre_salida, engine="openpyxl") as writer:
        datos["df_reemplazos_final"].to_excel(
            writer, sheet_name="Reemplazos validos", index=False
        )
        datos["df_salida"].to_excel(writer, sheet_name="EMPRESAS", index=False)
        datos["df_forzados_actualizado"].to_excel(
            writer, sheet_name="Reemplazos forzados", index=False
        )
        datos["df_reuc"].to_excel(
            writer, sheet_name=f"RUT_REUC_{datos['fecha_reuc']}", index=False
        )
        datos["df_unido"].to_excel(
            writer, sheet_name=f"Empresas_{datos['fecha_empresas']}", index=False
        )

    log(f"  Archivo generado: {nombre_salida.name}")
    return nombre_salida


# =========================================================
# ESCRITURA EN EL ARCHIVO DESTINO (.xlsm con macros)
# =========================================================
def escribir_en_destino(ruta_destino, datos, log=print, dejar_abierto=True):
    """
    Hoja EMPRESAS del archivo destino:
      B:C -> EMPRESA / RUT           (desde df_salida)
      H:I -> Reemplazada / Reemplazante  (reemplazos válidos + forzados)
    Fila 1 = encabezado en ambos bloques. Se limpia el contenido antes de pegar,
    manteniendo el formato de las celdas.
    """
    # --- Bloque B:C (empresas) ---
    empresas = (
        datos["df_salida"][["EMPRESA", "RUT"]]
        .astype(str)
        .values.tolist()
    )

    # --- Bloque H:I (reemplazos) ---
    reemp_validos = datos["df_reemplazos_final"][["Reemplazada", "Reemplazante"]]
    reemp_forzados = datos["df_forzados"][["Reemplazada", "Reemplazante"]]

    df_reemp = pd.concat([reemp_validos, reemp_forzados], ignore_index=True)
    df_reemp = df_reemp.astype(str).drop_duplicates(subset=["Reemplazada"], keep="first")
    reemplazos = df_reemp.values.tolist()

    log(f"  Empresas a pegar    : {len(empresas)}")
    log(f"  Reemplazos a pegar  : {len(reemplazos)} "
        f"({len(reemp_validos)} válidos + {len(reemp_forzados)} forzados)")

    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(str(ruta_destino), update_links=False)
        wb.app.calculation = "manual"

        try:
            sh = wb.sheets["EMPRESAS"]
        except Exception:
            raise RuntimeError("El archivo destino no tiene una hoja llamada 'EMPRESAS'.")

        # ---- Limpiar B:C ----
        fin_b = max(ultima_fila(sh, "B", 2), ultima_fila(sh, "C", 2))
        if fin_b >= 2:
            sh.range((2, col_letra_a_num("B")), (fin_b, col_letra_a_num("C"))).clear_contents()
            log(f"  Limpiado B2:C{fin_b}")

        # ---- Limpiar H:I ----
        fin_h = max(ultima_fila(sh, "H", 2), ultima_fila(sh, "I", 2))
        if fin_h >= 2:
            sh.range((2, col_letra_a_num("H")), (fin_h, col_letra_a_num("I"))).clear_contents()
            log(f"  Limpiado H2:I{fin_h}")

        # ---- Pegar bloques completos ----
        if empresas:
            sh.range((2, col_letra_a_num("B"))).value = empresas
            log(f"  Pegado B2:C{1 + len(empresas)}")

        if reemplazos:
            sh.range((2, col_letra_a_num("H"))).value = reemplazos
            log(f"  Pegado H2:I{1 + len(reemplazos)}")

        # ---- F1: fecha de la descarga REUC (DD-MM-AAAA) ----
        fecha_reuc = fecha_aaaammdd_a_ddmmaaaa(datos.get("fecha_reuc"))
        if fecha_reuc and fecha_reuc != "SIN_FECHA":
            sh.range((1, col_letra_a_num("F"))).value = fecha_reuc
            log(f"  F1 actualizado con fecha REUC: {fecha_reuc}")
        else:
            log("  OJO: no se pudo determinar la fecha REUC, F1 no se tocó.")

        wb.app.calculation = "automatic"
        wb.save()
        log("  Archivo destino guardado.")

        if dejar_abierto:
            app.api.Visible = True
            app.screen_updating = True
            app.display_alerts = True
            wb.activate()
        else:
            wb.close()
            app.quit()

        return True

    except Exception as e:
        log(f"ERROR al escribir en destino: {e}\n{traceback.format_exc()}")
        try:
            if wb:
                wb.close()
            if app:
                app.quit()
        except Exception:
            pass
        return False


# =========================================================
# VENTANA
# =========================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    # modo["traspaso"] se apaga si el usuario elige el destino a mano.
    modo = {"traspaso": traspaso is not None}

    v = VentanaChica("Reemplazos REUC", "Empresas con RUT y reemplazos a la hoja "
                     "EMPRESAS del cuadro 0.", traspaso, alto=430)
    var_carpeta = tk.StringVar(value=carpeta_datos_guardada(cfg))
    var_destino = tk.StringVar(
        value=(traspaso or {}).get("rutas", {}).get("cuadro_0")
              or cfg.get("archivo_destino", "[elige el cuadro 0 con Destino…]"))
    var_raiz = tk.StringVar(value=cfg.get("raiz_facturacion", RAIZ_FACTURACION_DEFAULT))
    var_balances = tk.StringVar(value="[buscando…]")
    var_sscc = tk.StringVar(value="[buscando…]")
    var_avisos = tk.StringVar(value="")
    var_abrir = tk.BooleanVar(value=True)
    avisos_par = {"lista": []}

    fe = v.seccion("Entradas")
    v.archivo(fe, "Datos REUC", var_carpeta)
    fr_reuc = tk.Frame(fe)
    fr_reuc.pack(fill="x", pady=(1, 2))
    btn_reuc = tk.Button(fr_reuc, text="Actualizar data REUC", font=("Segoe UI", 8),
                         command=lambda: actualizar_reuc())
    btn_reuc.pack(side="left", padx=(86, 0))
    tk.Label(fr_reuc, text="con tu navegador",
             font=("Segoe UI", 8), fg="#666666").pack(side="left", padx=4)
    v.archivo(fe, "Balances SEN", var_balances)
    v.archivo(fe, "1_CUADROS", var_sscc)
    tk.Label(fe, textvariable=var_avisos, font=("Segoe UI", 8), fg=NARANJO,
             anchor="w", justify="left", wraplength=500).pack(fill="x")
    fd = v.seccion("Destino")
    v.archivo(fd, "Cuadro 0", var_destino)
    v.casilla(fd, "Dejar Excel abierto al terminar", var_abrir)

    # ---------- disco compartido ----------
    def buscar_par(mostrar_detalle=False):
        var_balances.set("[buscando…]")
        var_sscc.set("[buscando…]")
        var_avisos.set("")
        v.root.update_idletasks()
        detalle = []
        par, info = buscar_par_mensual(var_raiz.get(), log=lambda m: detalle.append(str(m)))
        if par:
            var_balances.set(str(par["xlsb"]))
            var_sscc.set(str(par["xlsm"]))
            avisos_par["lista"] = par["avisos_xlsb"] + par["avisos_xlsm"]
            v.log(f"Disco compartido: {info}")
        else:
            var_balances.set(f"[{info}]")
            var_sscc.set(f"[{info}]")
            avisos_par["lista"] = []
        if mostrar_detalle or not par:
            v.log("== Detalle de la búsqueda en el disco compartido ==")
            for d in detalle:
                v.log(d)
        if avisos_par["lista"]:
            var_avisos.set("⚠ " + "  |  ".join(avisos_par["lista"]))
            for a in avisos_par["lista"]:
                v.log(f"  ⚠ {a}")

    # ---------- REUC ----------
    def actualizar_reuc():
        if not Path(var_carpeta.get()).is_dir():
            messagebox.showerror("Error", "Primero elige la carpeta de datos.")
            return
        import webbrowser
        webbrowser.open(URL_REUC_LOGIN)
        if not messagebox.askokcancel(
                "Iniciar sesión en REUC",
                "Se abrió REUC en tu navegador.\n\nInicia sesión ahí (y pasa la "
                "verificación de Cloudflare). Cuando ya veas REUC, presiona Aceptar "
                "y se bajan los dos archivos solos."):
            return
        v.log("== Actualizar data REUC ==")

        def fin(r):
            for p in r.values():
                v.log(f"  Guardado en auxiliares: {p.name}")
            v.var_estado.set("REUC actualizado")
            messagebox.showinfo("REUC actualizado", "Se descargaron los archivos REUC.")

        v.correr(lambda: descargar_reuc(carpeta_auxiliares(), log=v.log), fin,
                 boton=btn_reuc, estado="Esperando las descargas…")

    # ---------- elegir ----------
    def sel_destino():
        ini = cfg.get("archivo_destino", "")
        ini = str(Path(ini).parent) if ini and Path(ini).exists() else ""
        r = filedialog.askopenfilename(
            title="Seleccionar 0_CUADROS_RELIQUIDACIÓN SSCC", initialdir=ini,
            filetypes=[("Excel con macros", "*.xlsm"), ("Todos", "*.*")])
        if r:
            var_destino.set(r)
            guardar_config({"archivo_destino": r})
            cfg["archivo_destino"] = r
            # Si lo elige a mano, manda el usuario y no el traspaso.
            modo["traspaso"] = False

    def sel_carpeta():
        ini = var_carpeta.get() if Path(var_carpeta.get()).is_dir() else ""
        r = filedialog.askdirectory(
            title=f"Carpeta de datos (por defecto {NOMBRE_AUXILIARES})", initialdir=ini)
        if r:
            var_carpeta.set(r)
            guardar_config({"carpeta_datos": r})

    def sel_raiz():
        ini = var_raiz.get() if Path(var_raiz.get()).is_dir() else ""
        r = filedialog.askdirectory(title="Raíz Facturación (disco compartido)",
                                    initialdir=ini)
        if r:
            var_raiz.set(r)
            guardar_config({"raiz_facturacion": r})
            buscar_par(mostrar_detalle=True)

    # ---------- ejecutar ----------
    def ejecutar():
        carpeta, destino = var_carpeta.get(), var_destino.get()
        if not Path(carpeta).is_dir():
            messagebox.showerror("Error", "La carpeta de datos no es válida.")
            return
        if destino.startswith("[") or not Path(destino).is_file():
            messagebox.showerror("Error", "Elige el 0_CUADROS_RELIQUIDACIÓN SSCC con "
                                 "Destino….")
            return
        balances = var_balances.get()
        balances = None if balances.startswith("[") else balances
        sscc = var_sscc.get()
        sscc = None if sscc.startswith("[") else sscc
        avisos_previos = list(avisos_par["lista"])
        dejar_abierto = var_abrir.get()
        v.limpiar_log()

        def trabajo():
            # xlwings usa COM: en un hilo que no es el principal hay que
            # inicializarlo a mano.
            pythoncom = None
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except ImportError:
                pass
            try:
                v.log("== Procesando datos ==")
                if balances is None or sscc is None:
                    v.log("  (Falta algún archivo del disco compartido; se buscará en "
                          "la carpeta de datos)")
                for a in avisos_previos:
                    v.log(f"  ⚠ {a}")
                datos = procesar_datos(carpeta, destino, archivo_xlsb=balances,
                                       archivo_xlsm=sscc, log=v.log)
                v.estado("Generando archivo de salida…")
                carpeta_rr = carpeta_reemplazos_reuc(destino)
                v.log(f"== Archivo de salida ==\n  Carpeta: {carpeta_rr}")
                generar_archivo_salida(datos, carpeta_rr, log=v.log)
                v.log("== Respaldando archivos usados ==")
                respaldar_fuentes(datos, carpeta_rr, log=v.log)
                v.estado("Escribiendo en el cuadro 0…")
                v.log("== Escribiendo en 0_CUADROS_RELIQUIDACIÓN ==")
                ok = escribir_en_destino(destino, datos, log=v.log,
                                         dejar_abierto=dejar_abierto)
                alertas = []
                sin_rut = datos["df_sin_rut_a_pago"]
                if len(sin_rut) > 0:
                    alertas.append("Empresas a pago SIN RUT:\n" + "\n".join(
                        sin_rut["EMPRESA"].drop_duplicates().head(3).tolist()))
                df = datos["df_salida"]
                fuera = df[(df["EN_REUC"] == 0) & (df["CLP"] != 0)]
                if len(fuera) > 0:
                    alertas.append("Empresas a pago que NO están en REUC:\n" + "\n".join(
                        fuera["EMPRESA"].drop_duplicates().head(3).tolist()))
                return {"ok": ok, "alertas": alertas, "carpeta": str(carpeta_rr)}
            finally:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        def fin(r):
            for a in r["alertas"]:
                v.log("\n⚠ " + a)
            if r["alertas"]:
                messagebox.showwarning("Ojo", "\n\n".join(r["alertas"]))
            v.var_estado.set("Listo" if r["ok"] else "Terminó con errores — revisa el log")
            if r["ok"] and messagebox.askyesno(
                    "Listo", "Cuadro 0 actualizado.\n\n¿Abrir la carpeta Reemplazos REUC?"):
                abrir_en_explorador(r["carpeta"])

        v.correr(trabajo, fin, boton=btn, estado="Procesando datos…")

    btn = v.boton("EJECUTAR", ejecutar, principal=True)
    v.boton("Destino…", sel_destino)
    v.boton("Datos…", sel_carpeta)
    v.boton("Raíz T:…", sel_raiz)
    v.boton("Salir", v.root.destroy)
    avisar_legado(v.log)
    v.root.after(200, buscar_par)   # búsqueda automática al abrir
    v.iniciar()


if __name__ == "__main__":
    main()
