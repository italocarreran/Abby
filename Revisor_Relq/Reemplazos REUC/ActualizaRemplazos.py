"""
ActualizaRemplazos.py

Entradas (por ahora, en la carpeta desde donde se ejecuta el script):
"Reemplazos forzados.xlsx"                                      Copiar del ultimo mes reliquidado y confirmar
"datos_reuc_reemplazos_*.xlsx"                                  descargar de la pag REUC
"datos_reuc_*.xlsx"                                             descargar de la pag REUC
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
import queue
import shutil
import socket
import subprocess
import sys
import threading
import time
import traceback
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import xlwings as xw


# =========================================================
# CONFIG POR PC/USUARIO
# =========================================================
# La carpeta Auxiliares vive AL LADO del .py y es compartida por todos los
# usuarios. Ahi van los datos_reuc_* descargados y el archivo
# "Reemplazos forzados". Como el .py suele estar en un disco compartido, el
# config guarda las rutas SEPARADAS POR PC+USUARIO (ver get_usuario), asi que
# cada persona conserva las suyas sin pisar las de los demas.
CARPETA_AUXILIARES = Path(__file__).parent / "Auxiliares"
CONFIG_PATH = Path(__file__).resolve().parents[2] / "__config__" / "reemplazos_reuc.json"


def get_usuario():
    usuario = os.environ.get("USERNAME") or os.environ.get("USER") or "desconocido"
    return f"{socket.gethostname()}_{usuario}"


def leer_config():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get(get_usuario(), {})
        except Exception:
            return {}
    return {}


def guardar_config(data):
    todo = {}
    if CONFIG_PATH.exists():
        try:
            todo = json.load(open(CONFIG_PATH, "r", encoding="utf-8"))
        except Exception:
            todo = {}
    todo.setdefault(get_usuario(), {}).update(data)
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(todo, f, ensure_ascii=False, indent=2)


# =========================================================
# UTILIDADES
# =========================================================
# ---------------------------------------------------------------------------
# TRASPASO DESDE EL REVISOR
# ---------------------------------------------------------------------------
# El Revisor escribe un JSON en Salidas/AAMM/ y pasa su ruta como argv[1].
# Sin argumento este script funciona como siempre: rutas a mano y su propio
# reemplazos_reuc.json vive en __config__ y NO es el config compartido.
TRASPASO_ORIGEN = "Revisor_Reliquidacion"
TRASPASO_VERSION_MAX = 1


def leer_traspaso(argv):
    """Devuelve el dict del traspaso, o None si no vino o no es valido.
    Nunca lanza: si el JSON esta roto se cae al modo manual."""
    if len(argv) < 2 or not str(argv[1]).strip():
        return None
    ruta = Path(str(argv[1]).strip())
    try:
        if not ruta.is_file():
            return None
        with open(ruta, "r", encoding="utf-8") as f:
            d = json.load(f)
        if not isinstance(d, dict) or d.get("origen") != TRASPASO_ORIGEN:
            return None
        if int(d.get("version", 0)) > TRASPASO_VERSION_MAX:
            return None
        if not isinstance(d.get("rutas"), dict):
            d["rutas"] = {}
        return d
    except Exception:
        return None


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


def normalizar(texto):
    nfkd = unicodedata.normalize("NFKD", texto)
    return re.sub(
        r"\s+", " ", "".join(c for c in nfkd if not unicodedata.combining(c))
    ).strip().lower()


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
# DESCARGA AUTOMÁTICA DESDE REUC (playwright)
# =========================================================
URL_REUC_EMPRESAS = (
    "https://reuc.coordinador.cl/maestro_usuarios/empresas/exportar_reuc?&text_search="
)
URL_REUC_REEMPLAZADAS = (
    "https://reuc.coordinador.cl/maestro_usuarios/empresas/"
    "export_reemplazadas_data?&text_search="
)

def carpeta_auxiliares():
    """Carpeta Auxiliares (al lado del .py). La crea si no existe."""
    CARPETA_AUXILIARES.mkdir(parents=True, exist_ok=True)
    return CARPETA_AUXILIARES


def buscar_archivo_con_respaldo(carpeta_preferida, patron, log=print):
    """
    Busca primero en carpeta_preferida (la que eligio el usuario o quedo
    guardada en config). Si ahi no esta, cae de respaldo a Auxiliares
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
            log(f"    (no estaba en {carpeta_preferida}, se uso Auxiliares: "
                f"{encontrado.name})")
            return encontrado
        except FileNotFoundError:
            pass

    raise FileNotFoundError(
        f"No se encontro '{patron}' ni en {carpeta_preferida} ni en {aux}"
    )


def _sesion_lista(url):
    """
    True cuando el navegador ya volvio a REUC y no esta en una pantalla de
    login. El acceso unificado (hub / sso) rebota entre dominios varias
    veces, asi que solo damos por buena la URL de REUC misma.
    """
    if not url:
        return False
    u = url.lower()
    if "reuc.coordinador.cl" not in u:
        return False          # todavia esta en el acceso unificado
    if "/login" in u or "/logout" in u:
        return False
    return True


def _confirmar_sesion(pagina, timeout_ms=20_000):
    """
    Chequeo liviano y barato: entra a una pagina normal de REUC. Si el
    servidor devuelve la pagina (no rebota a login), la sesion sirve.
    Se usa ANTES de pedir el export, que es pesado de generar.
    """
    try:
        pagina.goto(
            "https://reuc.coordinador.cl/maestro_usuarios/empresas/",
            timeout=timeout_ms,
            wait_until="domcontentloaded",
        )
    except Exception:
        return False
    return _sesion_lista(pagina.url)


def _intentar_descarga(pagina, url, espera_ms=180_000):
    """
    Navega al link de exportacion y espera el evento de descarga del
    NAVEGADOR (igual que si el usuario hiciera clic el mismo: mismos
    headers, mismas cookies, mismo user-agent).

    OJO 1: cuando el link dispara una descarga directa, Chromium ABORTA
    la navegacion (se queda mostrando la pagina anterior). Eso es normal
    y no es un error: se ignora y se confia solo en el evento de descarga.

    OJO 2: el export lo genera el servidor al vuelo y puede demorar. NO
    hay que reintentar con timeouts cortos: cada reintento cancela la
    generacion en curso y obliga al servidor a partir de cero, que era
    justamente lo que hacia que la descarga no terminara nunca. Por eso
    aca se espera con paciencia (por defecto 3 minutos) una sola vez.
    """
    from playwright.sync_api import Error as PWError
    from playwright.sync_api import TimeoutError as PWTimeoutError

    try:
        with pagina.expect_download(timeout=espera_ms) as info:
            try:
                pagina.goto(url, timeout=espera_ms, wait_until="commit")
            except PWError:
                pass  # navegacion abortada por la descarga: esperado
        return info.value
    except PWTimeoutError:
        return None


def _borrar_viejos(carpeta_aux, patron_glob, excluir=None, log=print):
    """
    Borra archivos previos que calcen con patron_glob, salvo los que
    contengan 'excluir' en el nombre (para no borrar datos_reuc_reemplazos_*
    al limpiar datos_reuc_*). Evita que Auxiliares se llene de versiones
    viejas con timestamp distinto.
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


def descargar_reuc(carpeta_destino=None, log=print, timeout_login_seg=600,
                   espera_descarga_seg=180):
    """
    Abre un navegador para que el usuario inicie sesion en REUC y, apenas
    la sesion queda confirmada, descarga los dos exports:
      - datos_reuc_*.xlsx               (exportar_reuc)
      - datos_reuc_reemplazos_*.xlsx    (export_reemplazadas_data)

    Los guarda en la carpeta Auxiliares (por defecto, la que esta al lado
    del .py y es compartida por todos los usuarios).

    Flujo:
      1. Espera a que el usuario termine el login (el acceso unificado
         rebota entre dominios; solo se acepta la URL de REUC misma).
      2. Confirma la sesion con una peticion liviana.
      3. Descarga cada export UNA sola vez, esperando con paciencia.

    Seguridad: no se guarda nada de la sesion. Cada ejecucion abre un
    navegador nuevo y sin memoria (sin perfil, sin cookies persistidas).
    La clave nunca pasa por este script: se escribe directamente en la
    pagina real de REUC, dentro del navegador.

    Devuelve un dict {"reuc": Path, "reemplazos": Path}.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "Falta instalar Playwright. Abre una consola (cmd) y ejecuta:\n"
            "  pip install playwright\n"
            "  playwright install chromium"
        )

    carpeta_aux = Path(carpeta_destino) if carpeta_destino else CARPETA_AUXILIARES
    carpeta_aux.mkdir(parents=True, exist_ok=True)

    # (clave, url, prefijo, patron_a_borrar, fragmento_a_excluir_del_borrado)
    descargas = (
        ("reuc", URL_REUC_EMPRESAS, "datos_reuc",
         "datos_reuc*.xlsx", "reemplazos"),
        ("reemplazos", URL_REUC_REEMPLAZADAS, "datos_reuc_reemplazos",
         "datos_reuc_reemplazos*.xlsx", None),
    )

    resultados = {}

    with sync_playwright() as pw:
        navegador = pw.chromium.launch(headless=False)
        contexto = navegador.new_context(accept_downloads=True,
                                         viewport={"width": 1200, "height": 900})
        try:
            pagina = contexto.new_page()
            pagina.goto("https://reuc.coordinador.cl/login/",
                        wait_until="domcontentloaded")

            log("  Se abrio el navegador: inicia sesion en REUC.")
            log("  Puede pedirte pasar por el acceso unificado un par de veces;")
            log("  apenas entres, la descarga parte sola. No cierres la ventana.")

            # ---- 1. Esperar el login (sin tocar el export) ----
            t_ini = time.time()
            t_ultimo_aviso = time.time()

            while True:
                if time.time() - t_ini > timeout_login_seg:
                    raise TimeoutError(
                        "No se detecto el inicio de sesion a tiempo "
                        f"({timeout_login_seg // 60} min). Intenta de nuevo."
                    )
                try:
                    url_actual = pagina.url
                except Exception:
                    raise RuntimeError(
                        "Se cerro la ventana del navegador antes de terminar."
                    )

                if _sesion_lista(url_actual):
                    break

                if time.time() - t_ultimo_aviso > 30:
                    log(f"    (esperando el login... pagina actual: {url_actual})")
                    t_ultimo_aviso = time.time()

                pagina.wait_for_timeout(1000)

            # ---- 2. Confirmar que la sesion realmente sirve ----
            log("  Sesion detectada, confirmando...")
            if not _confirmar_sesion(pagina):
                raise RuntimeError(
                    "El navegador volvio a REUC pero la sesion no quedo activa. "
                    "Cierra e intenta de nuevo."
                )
            log("  Sesion confirmada.")

            # ---- 3. Descargar cada export, una sola vez y con paciencia ----
            for clave, url, prefijo, patron_borrado, excluir in descargas:
                log(f"  Descargando {prefijo}... "
                    f"(el servidor genera el archivo, puede demorar)")

                t_desc = time.time()
                descarga = _intentar_descarga(
                    pagina, url, espera_ms=espera_descarga_seg * 1000
                )
                if descarga is None:
                    raise RuntimeError(
                        f"No se pudo descargar {prefijo}: el servidor no entrego "
                        f"el archivo en {espera_descarga_seg}s. Puede estar lento; "
                        "intenta de nuevo en un rato."
                    )

                nombre = descarga.suggested_filename or ""
                if not normalizar(nombre).startswith(normalizar(prefijo)):
                    nombre = (f"{prefijo}_"
                              f"{datetime.today().strftime('%Y%m%d_%H%M%S')}.xlsx")

                # Borrar versiones anteriores para no acumular archivos viejos
                viejos = _borrar_viejos(carpeta_aux, patron_borrado, excluir, log=log)
                for v in viejos:
                    log(f"    Reemplazado (borrado): {v}")

                destino = carpeta_aux / nombre
                descarga.save_as(str(destino))
                log(f"    Guardado en {int(time.time() - t_desc)}s: {destino.name}")
                resultados[clave] = destino

        finally:
            contexto.close()
            navegador.close()

    return resultados


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
    # Los auxiliares (datos_reuc, reemplazos forzados) viven SIEMPRE en la
    # carpeta Auxiliares que esta al lado del .py, compartida por todos.
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

    df_reemplazos = pd.read_excel(
        archivo_reemplazos,
        sheet_name="pyexcel_sheet1",
        usecols="A:H",
        engine="openpyxl",
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
            p for p in Path(carpeta).glob("datos_reuc_*.xlsx")
            if "reemplazos" not in p.name.lower() and not p.name.startswith("~$")
        ]
        if not candidatos:
            raise FileNotFoundError(f"No hay datos_reuc_*.xlsx (sin 'reemplazos') en {carpeta}")
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
        log(f"    (no estaba en {carpeta_datos_p}, se uso Auxiliares: {archivo_reuc.name})")
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

    root = tk.Tk()
    root.title("Actualiza Reemplazos SSCC"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("900x620")

    # 1) Botones fijos abajo PRIMERO
    frame_btns_fijo = tk.Frame(root)
    frame_btns_fijo.pack(side="bottom", fill="x", pady=8)

    # 2) Canvas con scrollbar
    canvas = tk.Canvas(root, borderwidth=0, highlightthickness=0)
    scroll = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    contenedor = tk.Frame(canvas)
    canvas_win = canvas.create_window((0, 0), window=contenedor, anchor="nw")

    def _ajustar(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(canvas_win, width=canvas.winfo_width())

    contenedor.bind("<Configure>", _ajustar)
    canvas.bind("<Configure>", _ajustar)
    # La rueda del mouse: el canvas de la ventana solo scrollea si el puntero NO
    # esta sobre un widget que scrollea solo (el log, por ejemplo).
    #
    # bind_all captura la rueda GLOBALMENTE. Sin este filtro, al intentar subir
    # en el log se movia la ventana entera y el log no se movia: los dos
    # respondian a la misma rueda, o directamente ganaba el canvas.
    def _rueda(e, _cv=canvas):
        w = e.widget.winfo_toplevel().winfo_containing(e.x_root, e.y_root)
        pasos = int(-e.delta / 120)
        # Se sube por los padres: el puntero puede estar sobre un hijo del Text.
        while w is not None:
            if isinstance(w, (tk.Text, tk.Listbox)):
                w.yview_scroll(pasos, "units")
                return "break"
            if w is _cv:
                break
            w = getattr(w, "master", None)
        _cv.yview_scroll(pasos, "units")

    canvas.bind_all("<MouseWheel>", _rueda)

    # ---------- Variables ----------
    var_carpeta = tk.StringVar(
        value=cfg.get("carpeta_datos", str(CARPETA_AUXILIARES)))
    var_destino = tk.StringVar(
        value=(traspaso or {}).get("rutas", {}).get("cuadro_0")
              or cfg.get("archivo_destino", "[seleccionar archivo]"))
    var_raiz_facturacion = tk.StringVar(
        value=cfg.get("raiz_facturacion", RAIZ_FACTURACION_DEFAULT))
    var_balances = tk.StringVar(value="[sin buscar]")
    var_sscc = tk.StringVar(value="[sin buscar]")
    var_par_info = tk.StringVar(value="")
    var_aviso_balances = tk.StringVar(value="")
    var_aviso_sscc = tk.StringVar(value="")
    estado_avisos = {"lista": []}
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")
    cola_reuc = queue.Queue()
    cola_ejecutar = queue.Queue()
    ruta_carpeta_salida = {"path": ""}
    var_abrir = tk.BooleanVar(value=True)

    def actualizar_color_label(lbl, valor, es_archivo=False):
        if not valor or valor.startswith("["):
            lbl.config(fg="red")
        elif es_archivo and Path(valor).is_file():
            lbl.config(fg="blue")
        elif not es_archivo and Path(valor).is_dir():
            lbl.config(fg="blue")
        else:
            lbl.config(fg="red")

    # ---------- Aviso de procedencia (solo si vino del Revisor) ----------
    if traspaso:
        fr_aviso = tk.Frame(contenedor, bg="#fff4c2", bd=1, relief="solid")
        fr_aviso.pack(fill="x", padx=20, pady=(10, 0))
        tk.Label(fr_aviso,
                 text=f"Mes {traspaso.get('aamm') or '?'} — enviado por el Revisor",
                 bg="#fff4c2", font=("Segoe UI", 11, "bold")).pack(pady=(6, 0))
        tk.Label(fr_aviso,
                 text="El archivo destino lo resolvió el Revisor.\n"
                      "Las demás rutas se buscan igual que siempre.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    # ---------- Carpeta de datos ----------
    f1 = tk.LabelFrame(
        contenedor,
        text="Carpeta con los archivos de entrada (datos_reuc, reemplazos forzados, cuadros de pago)",
        padx=10, pady=6,
    )
    f1.pack(fill="x", padx=20, pady=6)

    lbl_carpeta = tk.Label(f1, textvariable=var_carpeta, wraplength=760,
                           justify="center", cursor="hand2", font=("Segoe UI", 9))
    lbl_carpeta.pack()
    lbl_carpeta.bind("<Button-1>", lambda e: abrir_en_explorador(var_carpeta.get()))

    def sel_carpeta():
        ini = var_carpeta.get() if Path(var_carpeta.get()).is_dir() else ""
        r = filedialog.askdirectory(title="Carpeta de datos", initialdir=ini)
        if r:
            var_carpeta.set(r)
            actualizar_color_label(lbl_carpeta, r)
            guardar_config({"carpeta_datos": r})

    def usar_auxiliares():
        r = str(carpeta_auxiliares())
        var_carpeta.set(r)
        actualizar_color_label(lbl_carpeta, r)
        guardar_config({"carpeta_datos": r})

    frame_btns1 = tk.Frame(f1)
    frame_btns1.pack(pady=(4, 0))
    tk.Button(frame_btns1, text="Examinar", command=sel_carpeta).pack(side="left", padx=4)
    tk.Button(frame_btns1, text="Usar Auxiliares (default)",
              command=usar_auxiliares).pack(side="left", padx=4)

    var_estado_reuc = tk.StringVar(value="")
    frame_reuc = tk.Frame(f1)
    frame_reuc.pack(pady=(8, 0))

    def _revisar_cola_reuc():
        try:
            while True:
                tipo, dato = cola_reuc.get_nowait()
                if tipo == "log":
                    log(dato)
                elif tipo == "ok":
                    log(f"  Guardado en auxiliares: {dato['reuc'].name}")
                    log(f"  Guardado en auxiliares: {dato['reemplazos'].name}")
                    log("== Fin actualización REUC ==\n")
                    var_estado_reuc.set("")
                    btn_reuc.config(state="normal", bg="#2d7a2d")
                    messagebox.showinfo(
                        "REUC actualizado",
                        "Se descargaron los archivos REUC correctamente."
                    )
                    return
                elif tipo == "error":
                    log(f"ERROR: {dato}")
                    var_estado_reuc.set("")
                    btn_reuc.config(state="normal", bg="#2d7a2d")
                    messagebox.showerror("Error al actualizar REUC", dato.splitlines()[0])
                    return
        except queue.Empty:
            pass
        root.after(300, _revisar_cola_reuc)

    def hilo_reuc(carpeta_base):
        def log_cola(msg):
            cola_reuc.put(("log", msg))
        try:
            resultados = descargar_reuc(carpeta_auxiliares(), log=log_cola)
            cola_reuc.put(("ok", resultados))
        except Exception as e:
            cola_reuc.put(("error", f"{e}\n{traceback.format_exc()}"))

    def actualizar_reuc():
        carpeta = var_carpeta.get()
        if not Path(carpeta).is_dir():
            messagebox.showerror("Error", "Primero selecciona la carpeta de datos.")
            return
        btn_reuc.config(state="disabled", bg="#aaaaaa")
        var_estado_reuc.set("Actualizando...")
        log("== Actualizar data REUC ==")
        threading.Thread(target=hilo_reuc, args=(carpeta,), daemon=True).start()
        root.after(300, _revisar_cola_reuc)

    btn_reuc = tk.Button(frame_reuc, text="Actualizar data REUC", bg="#2d7a2d", fg="white",
                         command=actualizar_reuc)
    btn_reuc.pack(side="left", padx=4)
    tk.Label(frame_reuc, textvariable=var_estado_reuc, fg="#2d7a2d",
             font=("Segoe UI", 9, "italic")).pack(side="left", padx=6)

    # ---------- Disco compartido: Balances SEN + SSCC ----------
    f3 = tk.LabelFrame(
        contenedor,
        text="Disco compartido — Balances SEN (Plabacom) + 1_CUADROS_PAGO_SSCC",
        padx=10, pady=6,
    )
    f3.pack(fill="x", padx=20, pady=6)

    lbl_raiz = tk.Label(f3, textvariable=var_raiz_facturacion, wraplength=760,
                        justify="center", cursor="hand2", font=("Segoe UI", 9))
    lbl_raiz.pack()
    lbl_raiz.bind("<Button-1>", lambda e: abrir_en_explorador(var_raiz_facturacion.get()))

    tk.Label(f3, textvariable=var_par_info,
             font=("Segoe UI", 9, "bold"), fg="#2d7a2d").pack(pady=(4, 2))

    lbl_balances = tk.Label(f3, textvariable=var_balances, wraplength=760,
                            justify="center", cursor="hand2", font=("Segoe UI", 9))
    lbl_balances.pack()
    lbl_balances.bind(
        "<Button-1>",
        lambda e: abrir_en_explorador(var_balances.get(), es_archivo=True)
        if not var_balances.get().startswith("[") else None,
    )

    tk.Label(f3, textvariable=var_aviso_balances, wraplength=760, justify="center",
             font=("Segoe UI", 8), fg=NARANJO).pack()

    lbl_sscc = tk.Label(f3, textvariable=var_sscc, wraplength=760,
                        justify="center", cursor="hand2", font=("Segoe UI", 9))
    lbl_sscc.pack(pady=(2, 0))
    lbl_sscc.bind(
        "<Button-1>",
        lambda e: abrir_en_explorador(var_sscc.get(), es_archivo=True)
        if not var_sscc.get().startswith("[") else None,
    )

    tk.Label(f3, textvariable=var_aviso_sscc, wraplength=760, justify="center",
             font=("Segoe UI", 8), fg=NARANJO).pack()

    def buscar_par(mostrar_detalle=False):
        var_balances.set("[buscando...]")
        var_sscc.set("[buscando...]")
        lbl_balances.config(fg="gray")
        lbl_sscc.config(fg="gray")
        var_par_info.set("")
        var_aviso_balances.set("")
        var_aviso_sscc.set("")
        root.update_idletasks()

        detalle = []

        def _cap(msg):
            detalle.append(str(msg))

        par, info = buscar_par_mensual(var_raiz_facturacion.get(), log=_cap)

        avisos = []
        if par:
            var_balances.set(str(par["xlsb"]))
            var_sscc.set(str(par["xlsm"]))
            var_par_info.set(f"Mes / proceso: {info}")
            avisos = par["avisos_xlsb"] + par["avisos_xlsm"]
        else:
            var_balances.set(f"[{info}]")
            var_sscc.set(f"[{info}]")
            var_par_info.set("")

        if mostrar_detalle or not par:
            log("== Detalle de busqueda en disco compartido ==")
            for d in detalle:
                log(d)
            log("")

        actualizar_color_label(lbl_balances, var_balances.get(), es_archivo=True)
        actualizar_color_label(lbl_sscc, var_sscc.get(), es_archivo=True)

        # ---- Avisos de coherencia nombre vs carpeta ----
        estado_avisos["lista"] = avisos

        if par and par["avisos_xlsb"]:
            lbl_balances.config(fg=NARANJO)
            var_aviso_balances.set("⚠ " + "  |  ".join(par["avisos_xlsb"]))
        else:
            var_aviso_balances.set("")

        if par and par["avisos_xlsm"]:
            lbl_sscc.config(fg=NARANJO)
            var_aviso_sscc.set("⚠ " + "  |  ".join(par["avisos_xlsm"]))
        else:
            var_aviso_sscc.set("")

        if avisos:
            log("== OJO: revisar nombres de archivo ==")
            for a in avisos:
                log(f"  ⚠ {a}")
            log("")

    def sel_raiz():
        ini = var_raiz_facturacion.get() if Path(var_raiz_facturacion.get()).is_dir() else ""
        r = filedialog.askdirectory(title="Raiz Facturacion", initialdir=ini)
        if r:
            var_raiz_facturacion.set(r)
            actualizar_color_label(lbl_raiz, r)
            guardar_config({"raiz_facturacion": r})
            buscar_par()

    frame_btn3 = tk.Frame(f3)
    frame_btn3.pack(pady=(4, 0))
    tk.Button(frame_btn3, text="Cambiar raiz", command=sel_raiz).pack(side="left", padx=4)
    tk.Button(frame_btn3, text="Buscar mas reciente",
              command=lambda: buscar_par(mostrar_detalle=True)).pack(side="left", padx=4)

    # ---------- Archivo destino ----------
    f2 = tk.LabelFrame(
        contenedor,
        text="Archivo destino:  0_CUADROS_RELIQUIDACIÓN SSCC_*.xlsm",
        padx=10, pady=6,
    )
    f2.pack(fill="x", padx=20, pady=6)

    lbl_destino = tk.Label(f2, textvariable=var_destino, wraplength=760,
                           justify="center", cursor="hand2", font=("Segoe UI", 9))
    lbl_destino.pack()
    lbl_destino.bind(
        "<Button-1>",
        lambda e: abrir_en_explorador(var_destino.get(), es_archivo=True)
        if not var_destino.get().startswith("[") else None,
    )

    def sel_destino():
        ini = cfg.get("archivo_destino", "")
        ini = str(Path(ini).parent) if ini and Path(ini).exists() else ""
        r = filedialog.askopenfilename(
            title="Seleccionar 0_CUADROS_RELIQUIDACIÓN SSCC",
            initialdir=ini,
            filetypes=[("Excel con macros", "*.xlsm"), ("Todos", "*.*")],
        )
        if r:
            var_destino.set(r)
            actualizar_color_label(lbl_destino, r, es_archivo=True)
            guardar_config({"archivo_destino": r})
            cfg["archivo_destino"] = r
            # Si lo elige a mano, manda el usuario y no el traspaso.
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Destino elegido a mano: se deja de usar la ruta del Revisor.")

    tk.Button(f2, text="Examinar", command=sel_destino).pack(pady=(4, 0))

    tk.Checkbutton(
        contenedor,
        text="Dejar Excel abierto al terminar",
        variable=var_abrir,
    ).pack(pady=(2, 6))

    # ---------- Progreso ----------
    tk.Label(contenedor, textvariable=var_estado, anchor="w").pack(fill="x", padx=20)
    tk.Label(contenedor, textvariable=var_tiempo,
             font=("Consolas", 10, "bold"), fg="#2d7a2d").pack()
    progress = ttk.Progressbar(contenedor, mode="determinate", length=400, maximum=100)
    progress.pack(fill="x", padx=20, pady=(0, 6))

    txt_log = tk.Text(contenedor, height=14, font=("Consolas", 9))
    txt_log.pack(fill="both", expand=True, padx=20, pady=(0, 10))

    def log(msg):
        txt_log.insert("end", str(msg) + "\n")
        txt_log.see("end")
        root.update_idletasks()

    def fmt_tiempo(seg):
        m, s = divmod(int(seg), 60)
        h, m = divmod(m, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"

    timer_state = {"running": False, "t_ini": 0.0}

    def tick():
        if timer_state["running"]:
            var_tiempo.set(fmt_tiempo(time.time() - timer_state["t_ini"]))
            root.after(500, tick)

    # ---------- Ejecutar ----------
    def hilo_ejecutar(carpeta, destino, balances, sscc, avisos_previos, dejar_abierto):
        """
        Corre TODO el proceso pesado (lectura de Excel, xlwings) en un hilo
        aparte para que la ventana no se congele (se pueda seguir haciendo
        scroll, mover la ventana, etc. mientras trabaja).

        xlwings usa COM (win32com) por debajo: en un hilo que no es el
        principal hay que inicializar COM explicitamente, si no, falla.
        """
        def log_cola(msg):
            cola_ejecutar.put(("log", msg))

        def progreso(v):
            cola_ejecutar.put(("progress", v))

        def estado(msg):
            cola_ejecutar.put(("estado", msg))

        pythoncom = None
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except ImportError:
            pass  # no es Windows / no esta instalado: xlwings igual podria fallar

        try:
            estado("Procesando datos...")
            log_cola("== Procesando datos ==")

            if balances is None or sscc is None:
                log_cola("  (Falta algun archivo del disco compartido; "
                        "se buscara en la carpeta de datos)")

            if avisos_previos:
                log_cola("== OJO: nombres de archivo que no calzan con su carpeta ==")
                for a in avisos_previos:
                    log_cola(f"  ⚠ {a}")
                log_cola("")

            datos = procesar_datos(
                carpeta, destino, archivo_xlsb=balances, archivo_xlsm=sscc, log=log_cola
            )
            progreso(50)

            estado("Generando archivo de salida...")
            log_cola("== Archivo de salida ==")
            carpeta_rr = carpeta_reemplazos_reuc(destino)
            log_cola(f"  Carpeta: {carpeta_rr}")
            generar_archivo_salida(datos, carpeta_rr, log=log_cola)
            cola_ejecutar.put(("carpeta_salida", str(carpeta_rr)))
            progreso(60)

            log_cola("== Respaldando archivos usados ==")
            respaldar_fuentes(datos, carpeta_rr, log=log_cola)
            progreso(70)

            estado("Escribiendo en archivo destino...")
            log_cola("== Escribiendo en 0_CUADROS_RELIQUIDACIÓN ==")
            ok = escribir_en_destino(
                destino, datos, log=log_cola, dejar_abierto=dejar_abierto
            )
            progreso(100)

            # ---------- Alertas ----------
            alertas = []

            df_sin_rut_a_pago = datos["df_sin_rut_a_pago"]
            if len(df_sin_rut_a_pago) > 0:
                ejemplos = df_sin_rut_a_pago["EMPRESA"].drop_duplicates().head(3).tolist()
                alertas.append(
                    "⚠️ Ojito: hay empresas a pago SIN RUT:\n" + "\n".join(ejemplos)
                )

            df_salida = datos["df_salida"]
            fuera_reuc = df_salida[(df_salida["EN_REUC"] == 0) & (df_salida["CLP"] != 0)]
            if len(fuera_reuc) > 0:
                ejemplos = fuera_reuc["EMPRESA"].drop_duplicates().head(3).tolist()
                alertas.append(
                    "⚠️ Ojito: hay empresas a pago que NO están en REUC:\n" + "\n".join(ejemplos)
                )

            estado("Listo" if ok else "Terminó con errores")
            cola_ejecutar.put(("fin", {"ok": ok, "alertas": alertas}))

        except Exception as e:
            cola_ejecutar.put(("error", f"{e}\n{traceback.format_exc()}"))
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _revisar_cola_ejecutar():
        try:
            while True:
                tipo, dato = cola_ejecutar.get_nowait()

                if tipo == "log":
                    log(dato)
                elif tipo == "progress":
                    progress["value"] = dato
                elif tipo == "estado":
                    var_estado.set(dato)
                elif tipo == "carpeta_salida":
                    ruta_carpeta_salida["path"] = dato
                    btn_abrir_carpeta.config(state="normal")
                elif tipo == "fin":
                    if dato["alertas"]:
                        for a in dato["alertas"]:
                            log("")
                            log(a)
                        messagebox.showwarning(
                            "ALERTA DE VALIDACIÓN", "\n\n".join(dato["alertas"])
                        )
                    log("\n== FIN ==")
                    timer_state["running"] = False
                    btn.config(state="normal", bg="#2d7a2d")
                    return
                elif tipo == "error":
                    var_estado.set("Error")
                    log(f"ERROR: {dato}")
                    messagebox.showerror("Error", dato.splitlines()[0])
                    timer_state["running"] = False
                    btn.config(state="normal", bg="#2d7a2d")
                    return
        except queue.Empty:
            pass
        root.after(200, _revisar_cola_ejecutar)

    def ejecutar():
        carpeta = var_carpeta.get()
        destino = var_destino.get()

        if not Path(carpeta).is_dir():
            messagebox.showerror("Error", "La carpeta de datos no es válida.")
            return
        if destino.startswith("[") or not Path(destino).is_file():
            messagebox.showerror("Error", "Debes seleccionar el archivo 0_CUADROS_RELIQUIDACIÓN SSCC.")
            return

        btn.config(state="disabled", bg="#aaaaaa")
        txt_log.delete("1.0", "end")
        progress["value"] = 0
        timer_state["running"] = True
        timer_state["t_ini"] = time.time()
        tick()

        balances = var_balances.get()
        balances = balances if not balances.startswith("[") else None
        sscc = var_sscc.get()
        sscc = sscc if not sscc.startswith("[") else None

        threading.Thread(
            target=hilo_ejecutar,
            args=(carpeta, destino, balances, sscc,
                  list(estado_avisos["lista"]), var_abrir.get()),
            daemon=True,
        ).start()
        root.after(200, _revisar_cola_ejecutar)

    btn = tk.Button(frame_btns_fijo, text="Ejecutar", bg="#2d7a2d", fg="white",
                    font=("Segoe UI", 10, "bold"), command=ejecutar)
    btn.pack(side="left", padx=8, expand=True)

    btn_abrir_carpeta = tk.Button(
        frame_btns_fijo,
        text="Abrir carpeta Reemplazos REUC",
        state="disabled",
        command=lambda: abrir_en_explorador(ruta_carpeta_salida["path"])
        if ruta_carpeta_salida["path"] else None,
    )
    btn_abrir_carpeta.pack(side="left", padx=8, expand=True)

    tk.Button(frame_btns_fijo, text="Salir", command=root.destroy).pack(
        side="right", padx=8
    )

    # ---------- Init ----------
    actualizar_color_label(lbl_carpeta, var_carpeta.get())
    actualizar_color_label(lbl_destino, var_destino.get(), es_archivo=True)
    actualizar_color_label(lbl_raiz, var_raiz_facturacion.get())
    root.after(200, buscar_par)   # busqueda automatica al abrir

    root.mainloop()


if __name__ == "__main__":
    main()