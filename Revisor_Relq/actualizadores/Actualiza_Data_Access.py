"""
Actualiza la tabla [Sobrecostos] de un Access .mdb consolidando la informacion
de tres archivos Excel con macros (.xlsm).

Estructura de carpetas esperada:

    <RAIZ>/
        01 Sobrecostos/                         <- aqui esta el .mdb que se selecciona
        00 Entregables/
            01 Sobrecostos/                     Calculo_SobrecostosSSCC*.xlsm  (hoja SOBRECOSTOS TOTAL)
            02 Costo de Oportunidad/            Calculo_CO*.xlsm               (hoja CO TOTAL)
            03 Costo de Combustible Adicional/  Consolidado_CCA*.xlsm          (hoja CCA)

Los .xlsm SOLO se abren en modo lectura (read_only) y nunca se guardan:
el filtrado y el orden de columnas los hace Python en memoria.

En el Access se reemplaza unicamente la informacion de los tipos seleccionados
(se borran las filas cuyo Tipo_sobrecosto coincide con los tipos que traen los
Excel de esa fuente y luego se insertan las nuevas). Los otros tipos quedan
intactos.

Requiere: xlwings, pyodbc  ->  pip install xlwings pyodbc
Y el driver "Microsoft Access Driver (*.mdb, *.accdb)" de la MISMA arquitectura
(32/64 bits) que el Python que ejecuta el script.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import json
import subprocess
import sys
import re
import socket
import os
import traceback
import unicodedata
import time
import threading
import queue
import datetime
import decimal

# ---------------------------------------------------------------------------
# CONFIGURACION DE FUENTES
# ---------------------------------------------------------------------------
# bloques: lista de (col_ini, col_fin) que se leen y se concatenan EN ESE ORDEN
#          hasta armar las 5 columnas del Access.
# fila_ini: primera fila de datos (encabezado ya omitido).
# filtrar_ceros: si True, se descartan las filas cuya ultima columna (Sobrecosto) sea 0.

# base: de donde cuelga la "carpeta" de cada fuente.
#   "mdb"         -> la carpeta del propio .mdb, que ES "01 Sobrecostos".
#   "entregables" -> la carpeta "00 Entregables".
# El SSCC va por "mdb" a proposito: el archivo que hay que leer es el MAESTRO de
# "01 Sobrecostos" (el que edita el usuario y el que el Revisor suma en V4), no
# la copia de "00 Entregables/01 Sobrecostos". CO y CCA solo existen en
# "00 Entregables", asi que esos si cuelgan de ahi.
FUENTES = {
    "SSCC": {
        "etiqueta": "SSCC  (Calculo_SobrecostosSSCC*.xlsm / SOBRECOSTOS TOTAL)  — MAESTRO de 01 Sobrecostos",
        "base": "mdb",
        "carpeta": None,
        "patron": r"calculo_sobrecostosssc",
        "hoja": "SOBRECOSTOS TOTAL",
        "bloques": [("B", "F")],                      # B,C,D,E,F
        "fila_ini": 5,                                # encabezado en fila 4
        "filtrar_ceros": False,
    },
    "CO": {
        "etiqueta": "CO  (Calculo_CO*.xlsm / CO TOTAL)",
        "base": "entregables",
        "carpeta": "02 Costo de Oportunidad",
        "patron": r"calculo_co",
        "hoja": "CO TOTAL",
        "bloques": [("B", "C"), ("G", "G"), ("E", "F")],   # A,B,G,E,F
        "fila_ini": 5,                                # encabezado en fila 4
        "filtrar_ceros": True,
    },
    "CCA": {
        "etiqueta": "CCA  (Consolidado_CCA*.xlsm / CCA)",
        "base": "entregables",
        "carpeta": "03 Costo de Combustible Adicional",
        "patron": r"consolidado_cca",
        "hoja": "CCA",
        "bloques": [("I", "L"), ("BC", "BC")],        # I,J,K,L,BC
        "fila_ini": 3,                                # encabezado en fila 2
        "filtrar_ceros": True,
    },
}

ORDEN_FUENTES = ["SSCC", "CO", "CCA"]

# Capacidades que este modulo le ofrece a quien lo importe (Actualiza_Energia.py).
# Sirve para que, si alguien copia un .py y no el otro, la falla salga al tiro y
# con un mensaje claro, en vez de a mitad de proceso o -peor- sin fallar y
# metiendo datos sin filtrar en el Access.
#   fuentes_externas   -> proceso(..., fuentes=...) acepta un dict de fuentes propio
#   filtro_por_valores -> leer_fuente respeta cfg["filtro"]
#   borrar_todo        -> proceso(..., borrar_todo=True) vacia la tabla entera
#   cols_no_cero       -> leer_fuente respeta cfg["cols_no_cero"]
#
# CADA VEZ QUE SE AGREGUE UNA CAPACIDAD NUEVA HAY QUE SUMARLA ACA. Si no, el
# script que la use se cae con un TypeError raro en vez de decir "copia el
# archivo actualizado". Y peor: si la capacidad es un FILTRO que no se aplica,
# no falla nada y entran datos de mas.
#   forzar_valores     -> leer_fuente respeta cfg["forzar_valores"], que pisa
#                         una columna con un valor fijo en todas las filas
CAPACIDADES = frozenset({"fuentes_externas", "filtro_por_valores",
                         "borrar_todo", "cols_no_cero", "forzar_valores"})


def _verificar_capacidades():
    """Comprueba que lo declarado en CAPACIDADES exista de verdad en el codigo.

    Existe porque ya paso dos veces: se declara una capacidad y el parametro que
    la implementa no esta, o al reves. Declarar de mas es lo peor, porque el
    script que importa este modulo cree que puede y se cae con un TypeError raro
    en vez de decir "copia el archivo actualizado".

    Corre al importar y cuesta microsegundos. Si algo no calza, lanza al toque.
    """
    import inspect
    faltan = []
    params = set(inspect.signature(proceso).parameters)
    if "fuentes_externas" in CAPACIDADES and "fuentes" not in params:
        faltan.append("fuentes_externas: proceso() no tiene el parametro 'fuentes'")
    if "borrar_todo" in CAPACIDADES and "borrar_todo" not in params:
        faltan.append("borrar_todo: proceso() no tiene el parametro 'borrar_todo'")
    try:
        src_filtrar = inspect.getsource(_filtrar_matriz)
    except (OSError, TypeError):
        # getsource falla si el modulo se cargo sin archivo (exec, un .pyc suelto,
        # un empaquetado). No es motivo para abortar: se comprueban solo las
        # capacidades que se pueden ver por la firma.
        src_filtrar = None
    if src_filtrar is None:
        if faltan:
            raise RuntimeError(
                "Actualiza_Data_Access.py está inconsistente: declara "
                "capacidades que no implementa.\n  - " + "\n  - ".join(faltan))
        return
    if "filtro_por_valores" in CAPACIDADES and 'cfg.get("filtro")' not in src_filtrar:
        faltan.append("filtro_por_valores: _filtrar_matriz no lee cfg['filtro']")
    if "cols_no_cero" in CAPACIDADES and "cols_no_cero" not in src_filtrar:
        faltan.append("cols_no_cero: _filtrar_matriz no lee cfg['cols_no_cero']")
    if "forzar_valores" in CAPACIDADES and "forzar_valores" not in src_filtrar:
        faltan.append(
            "forzar_valores: _filtrar_matriz no lee cfg['forzar_valores']")
    if faltan:
        raise RuntimeError(
            "Actualiza_Data_Access.py está inconsistente: declara capacidades "
            "que no implementa.\n  - " + "\n  - ".join(faltan))

TABLA_ACCESS = "Sobrecostos"
COLUMNAS_ESPERADAS = ["Clave Año_Mes", "Tipo_sobrecosto", "Central",
                      "Hora Mensual", "Sobrecosto"]

# config.json es compartido con el Revisor y el resto de los actualizadores,
# que viven un nivel arriba (en scripts/, junto al Revisor). No es
# Path(__file__).parent porque este script esta en actualizadores/.
CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.json"


# ---------------------------------------------------------------------------
# CONFIG POR PC/USUARIO
# ---------------------------------------------------------------------------
def get_usuario():
    usuario = os.environ.get("USERNAME") or os.environ.get("USER") or "desconocido"
    return f"{socket.gethostname()}_{usuario}"


def leer_config():
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get(get_usuario(), {})
    except Exception:
        pass
    return {}


def escribir_json(ruta, data):
    """Escritura atomica: primero un .tmp y despues os.replace.
    Evita dejar el archivo truncado si algo falla a medio camino."""
    ruta = Path(ruta)
    tmp = ruta.with_suffix(ruta.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, ruta)


def _modificar_config(mutador):
    """config.json lo comparten el Revisor y los dos actualizadores. Solo se
    agregan o actualizan claves, nunca se borra nada, y si el archivo existe
    pero no se puede interpretar NO se escribe (mejor perder un ajuste que el
    archivo entero)."""
    todo = {}
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                todo = json.load(f)
            if not isinstance(todo, dict):
                return False
        except Exception:
            return False
    try:
        mutador(todo)
        escribir_json(CONFIG_PATH, todo)
        return True
    except Exception:
        return False


def guardar_config(data):
    return _modificar_config(
        lambda todo: todo.setdefault(get_usuario(), {}).update(data))


# ---------------------------------------------------------------------------
# TRASPASO DESDE EL REVISOR
# ---------------------------------------------------------------------------
# El Revisor escribe un JSON en ../00_Salidas/AAMM/ y pasa su ruta como argv[1].
# Sin argumento el script funciona como siempre: busca los archivos solo.
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


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------
def abrir_en_explorador(ruta, es_archivo=False):
    if not ruta or ruta.startswith("["):
        return
    p = Path(ruta)
    if not p.exists():
        return
    carpeta = p.parent if es_archivo else p
    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(carpeta)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(carpeta)])
    else:
        subprocess.Popen(["xdg-open", str(carpeta)])


def normalizar(texto):
    nfkd = unicodedata.normalize("NFKD", str(texto))
    limpio = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", limpio).strip().lower()


def buscar_carpeta(base, nombre):
    """Busca una subcarpeta comparando nombres normalizados (tildes/mayusculas)."""
    if not base or not Path(base).is_dir():
        return None
    objetivo = normalizar(nombre)
    candidatos = [d for d in Path(base).iterdir() if d.is_dir()]
    for d in candidatos:
        if normalizar(d.name) == objetivo:
            return d
    for d in candidatos:
        if objetivo in normalizar(d.name):
            return d
    # ultimo intento: coincidencia por el numero inicial ("01", "02", ...)
    m = re.match(r"^(\d+)", objetivo)
    if m:
        for d in candidatos:
            if normalizar(d.name).startswith(m.group(1)):
                return d
    return None


def buscar_archivo(carpeta, patron_regex, extensiones=(".xlsm", ".xlsx", ".xlsb")):
    """Archivo mas reciente cuyo nombre normalizado calza con el patron."""
    if not carpeta or not Path(carpeta).is_dir():
        return None
    patron = re.compile(patron_regex, re.IGNORECASE)
    candidatos = [f for f in Path(carpeta).iterdir()
                  if f.is_file() and f.suffix.lower() in extensiones
                  and not f.name.startswith("~$")
                  and patron.search(normalizar(f.name))]
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]


def col_letra(n):
    """4 -> "D".  Al reves de col_letra_a_num."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_letra_a_num(letra):
    n = 0
    for c in str(letra).upper():
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n


def fmt_tiempo(seg):
    m, s = divmod(int(seg), 60)
    h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def es_vacio(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def es_cero(v):
    if es_vacio(v):
        return False
    if isinstance(v, bool):
        return v is False
    if isinstance(v, (int, float, decimal.Decimal)):
        try:
            return abs(float(v)) < 1e-12
        except Exception:
            return False
    if isinstance(v, str):
        n = a_numero(v)
        return n is not None and abs(float(n)) < 1e-12
    return False


# ---------------------------------------------------------------------------
# LECTURA DE EXCEL (xlwings, solo lectura)
# ---------------------------------------------------------------------------
def buscar_hoja(wb, nombre):
    objetivo = normalizar(nombre)
    for sh in wb.sheets:
        if normalizar(sh.name) == objetivo:
            return sh
    for sh in wb.sheets:
        if objetivo in normalizar(sh.name):
            return sh
    return None


def ultima_fila(sheet, col_num, desde):
    """Ultima fila con contenido en una columna, usando .formula (capta formulas)."""
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


def leer_bloque(sheet, f1, c1, f2, c2):
    """Devuelve siempre una lista de filas (listas)."""
    if f2 < f1:
        return []
    datos = sheet.range((f1, c1), (f2, c2)).options(ndim=2).value
    return datos if datos else []


# ---------------------------------------------------------------------------
#  Lectura rapida: el .xlsx/.xlsm como ZIP, sin abrir Excel
# ---------------------------------------------------------------------------
#  Las planillas son pesadas y aca solo hay que LEERLAS. Abrirlas con xlwings
#  levanta Excel entero por COM, que para leer es lo mas lento que hay: en la T:
#  son varios minutos por planilla.
#
#  Un .xlsx/.xlsm es un ZIP con XML adentro. Se puede leer la hoja escaneando ese
#  XML por trozos, sin Excel y sin cargar el archivo entero en memoria. Es el
#  mismo lector que usa el Revisor.
#
#  Limitacion importante: se lee el RESULTADO guardado de cada celda (el nodo
#  <v>), nunca la formula. Si el archivo nunca se recalculo, los resultados
#  pueden ser viejos. Excel guarda el resultado siempre que se guarde el archivo,
#  asi que en la practica no molesta, pero por eso queda xlwings de respaldo.
NS_XL = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def es_zip_excel(ruta):
    return Path(ruta).suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm")


def ubicar_hoja_xml(z, hoja):
    """Dentro del zip de un .xlsx/.xlsm, devuelve (ruta_del_xml, lista_de_hojas).
    ruta_del_xml es None si la hoja no existe."""
    import xml.etree.ElementTree as ET
    nombres = set(z.namelist())
    if "xl/workbook.xml" not in nombres:
        return None, []
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    lista = list(wb.iter(f"{NS_XL}sheet"))
    hojas = [sh.get("name", "") for sh in lista]
    rid = None
    # "#1", "#2", ... permiten pedir la hoja por posicion cuando no se sabe el
    # nombre o cuando cambia de un mes a otro.
    m_pos = re.fullmatch(r"#(\d+)", str(hoja).strip())
    if m_pos:
        i = int(m_pos.group(1)) - 1
        if 0 <= i < len(lista):
            rid = lista[i].get(f"{NS_REL}id")
    else:
        for sh in lista:
            if normalizar(sh.get("name", "")) == normalizar(hoja):
                rid = sh.get(f"{NS_REL}id")
                break
    if rid is None or "xl/_rels/workbook.xml.rels" not in nombres:
        return None, hojas
    destino = None
    for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
        if r.get("Id") == rid:
            destino = r.get("Target")
            break
    if not destino:
        return None, hojas
    ruta_hoja = destino[1:] if destino.startswith("/") else "xl/" + destino
    ruta_hoja = ruta_hoja.replace("xl/xl/", "xl/")
    return (ruta_hoja if ruta_hoja in nombres else None), hojas


def expandir_columnas(rango):
    """'CF:CI' -> ['CF','CG','CH','CI'].  'CD' -> ['CD']."""
    partes = str(rango).replace(" ", "").replace("$", "").upper().split(":")
    a = col_letra_a_num(partes[0])
    b = col_letra_a_num(partes[-1]) if len(partes) > 1 else a
    letras = []
    for n in range(min(a, b), max(a, b) + 1):
        s, x = "", n
        while x > 0:
            x, r = divmod(x - 1, 26)
            s = chr(ord("A") + r) + s
        letras.append(s)
    return letras


def leer_columnas_rapido(ruta, hoja, columnas, fila_inicio, log):
    """Lee columnas completas de un .xlsx/.xlsm escaneando el XML por trozos.
    Devuelve {"COL": {fila: valor}} con los valores ya calculados, o None.
    Igual que en el resto, se lee SOLO el resultado y nunca el nodo <f>."""
    import zipfile
    import xml.etree.ElementTree as ET

    if not es_zip_excel(ruta):
        log(f"    ! {Path(ruta).name}: formato no soportado para lectura rápida")
        return None
    objetivo = [c.upper() for c in columnas]
    if not objetivo:
        return {}
    alternativas = b"|".join(sorted((c.encode() for c in objetivo),
                                    key=len, reverse=True))
    patron = re.compile(rb'<c r="(' + alternativas + rb')(\d+)"([^>]*?)(?:/>|>(.*?)</c>)',
                        re.S)
    fila_inicio = int(fila_inicio)
    datos = {c: {} for c in objetivo}

    def desescapar(b):
        # Delegado al desescapador de verdad: maneja tambien &#243; y &#x00F3;.
        return desescapar_xml(b)

    try:
        with zipfile.ZipFile(str(ruta)) as z:
            ruta_hoja, hojas = ubicar_hoja_xml(z, hoja)
            if ruta_hoja is None:
                log(f"    ! La hoja '{hoja}' no existe. Hojas: {', '.join(hojas)}")
                return None
            compartidas = []
            if "xl/sharedStrings.xml" in set(z.namelist()):
                ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
                compartidas = ["".join(si.itertext()) for si in ss.iter(f"{NS_XL}si")]

            with z.open(ruta_hoja) as f:
                cola = b""
                while True:
                    trozo = f.read(1 << 20)
                    if not trozo:
                        break
                    buf = cola + trozo
                    fin = 0
                    for m in patron.finditer(buf):
                        fin = m.end()
                        col, nfila = m.group(1).decode(), int(m.group(2))
                        if nfila < fila_inicio:
                            continue
                        attrs, cuerpo = m.group(3), m.group(4) or b""
                        valor = None
                        if b'inlineStr' in attrs:
                            mv = re.search(rb"<is>(.*?)</is>", cuerpo, re.S)
                            if mv:
                                valor = desescapar(re.sub(rb"<[^>]+>", b"", mv.group(1)))
                        else:
                            mv = re.search(rb"<v>(.*?)</v>", cuerpo, re.S)
                            if mv:
                                bruto = mv.group(1)
                                if b't="s"' in attrs:
                                    try:
                                        valor = compartidas[int(bruto)]
                                    except Exception:
                                        valor = None
                                elif b't="e"' in attrs or b't="str"' in attrs:
                                    valor = desescapar(bruto)
                                elif b't="b"' in attrs:
                                    valor = bool(int(bruto))
                                else:
                                    try:
                                        valor = float(bruto)
                                    except ValueError:
                                        valor = desescapar(bruto)
                        if valor is not None and not (isinstance(valor, str)
                                                      and not valor.strip()):
                            datos[col][nfila] = valor
                    cola = buf[fin:] if fin else buf[-8192:]
    except Exception as e:
        log(f"    ! No se pudo leer {Path(ruta).name}: {e}")
        return None
    return datos


_ENT_XML = {"lt": "<", "gt": ">", "quot": '"', "apos": "'", "amp": "&"}
_RE_ENT = re.compile(r"&(?:#(\d+)|#[xX]([0-9a-fA-F]+)|(lt|gt|quot|apos|amp));")


def desescapar_xml(b):
    """Convierte el texto crudo del XML de Excel a texto de verdad.

    Hay que manejar las referencias NUMERICAS (&#243; = o con tilde), no solo las
    cinco entidades con nombre: los nombres de empresa chilenos vienen llenos de
    tildes y ñ, y algunos escritores de Excel las guardan asi. Si no se
    desescapan, "Enel Generaci&#243;n" y "Enel Generación" no se parecen en nada
    al comparar, y el cuadro de pago reporta un descuadre que no existe.

    Se resuelve en UNA pasada a proposito. Reemplazar "&amp;" primero y despues
    "&lt;" convertiria "&amp;lt;" (un literal "&lt;") en "<", que es otra cosa.
    """
    if isinstance(b, bytes):
        b = b.decode("utf-8", "ignore")

    def uno(m):
        dec, hexa, nombre = m.group(1), m.group(2), m.group(3)
        try:
            if dec is not None:
                return chr(int(dec))
            if hexa is not None:
                return chr(int(hexa, 16))
        except (ValueError, OverflowError):
            return m.group(0)
        return _ENT_XML[nombre]

    return _RE_ENT.sub(uno, b)


def leer_matriz_rapida(ruta, cfg, log):
    """La matriz de la fuente leyendo el ZIP, sin Excel. None si no se pudo."""
    cols = []
    for a, b in cfg["bloques"]:
        na, nb = col_letra_a_num(a), col_letra_a_num(b)
        cols += [col_letra(n) for n in range(na, nb + 1)]
    f_ini = int(cfg["fila_ini"])
    datos = leer_columnas_rapido(ruta, cfg["hoja"], cols, f_ini, log)
    if datos is None:
        return None
    filas = set()
    for c in cols:
        filas |= set(datos.get(c.upper(), {}))
    filas = sorted(f for f in filas if f >= f_ini)
    if not filas:
        log("    ADVERTENCIA: no hay datos bajo el encabezado.")
        return []
    log(f"    Filas {f_ini} a {max(filas)} ({len(filas)} con algo)")
    return [[datos.get(c.upper(), {}).get(f) for c in cols] for f in filas]


def leer_fuente(app, ruta_xlsm, cfg, log):
    """La matriz de 5 columnas de una fuente, ya filtrada.

    Primero intenta el lector rapido (ZIP + XML, sin Excel). Si el archivo no es
    un .xlsx/.xlsm o algo falla, cae a xlwings.

    `app` puede ser una instancia de xlwings, None, o una FUNCION que la crea al
    llamarla. Lo ultimo es lo que usa proceso(): asi Excel se abre solo si de
    verdad hace falta, que con planillas normales no pasa nunca.
    """
    matriz = None
    if cfg.get("lectura_rapida", True) and es_zip_excel(ruta_xlsm):
        t0 = time.time()
        log(f"    Leyendo {Path(ruta_xlsm).name} sin abrir Excel...")
        try:
            matriz = leer_matriz_rapida(ruta_xlsm, cfg, log)
        except Exception as e:
            log(f"    ! Falló la lectura rápida ({e}); se abre con Excel.")
            matriz = None
        if matriz is not None:
            log(f"    Leído en {time.time() - t0:.1f} s")
    if matriz is None:
        real = app() if callable(app) else app
        if real is None:
            raise RuntimeError(
                f"No se pudo leer {Path(ruta_xlsm).name} con el lector rápido y "
                f"no hay Excel disponible para el respaldo.")
        matriz = _leer_fuente_xlwings(real, ruta_xlsm, cfg, log)
    return _filtrar_matriz(matriz, cfg, log)


def _filtrar_matriz(matriz, cfg, log):
    """Descarta filas vacias, las del filtro por valores y las que tienen 0 en
    las columnas de control. Es lo mismo para los dos lectores."""
    filtro = cfg.get("filtro")
    col_filtro = filtro["col"] if filtro else None
    permitidos = {normalizar(v) for v in filtro["valores"]} if filtro else None

    # forzar_valores: {indice_de_columna: valor}. Pisa esa columna en TODAS las
    # filas que quedan. Hace falta para el .mdb de la planilla 9, donde la
    # Clave Año_Mes viene mal desde el origen (siempre 23xx) y hay que ponerle el
    # mes de verdad, sacado del nombre de los archivos.
    forzar = cfg.get("forzar_valores") or {}

    n_vacias = n_ceros = n_filtradas = 0
    salida = []
    for fila in matriz:
        if all(es_vacio(v) for v in fila):
            n_vacias += 1
            continue
        if filtro is not None:
            actual = fila[col_filtro] if col_filtro < len(fila) else None
            if normalizar(actual) not in permitidos:
                n_filtradas += 1
                continue
        if cfg["filtrar_ceros"]:
            # Por omision se mira la ULTIMA columna (el monto). Con
            # cfg["cols_no_cero"] se pueden pedir otras: el .mdb de la planilla 9
            # tambien descarta las filas con Central = 0.
            idx = cfg.get("cols_no_cero") or [len(fila) - 1]
            fuera = False
            for i in idx:
                v = fila[i] if 0 <= i < len(fila) else None
                if es_vacio(v) or es_cero(v):
                    fuera = True
                    break
            if fuera:
                n_ceros += 1
                continue
        if forzar:
            fila = list(fila)
            for i, v in forzar.items():
                if 0 <= i < len(fila):
                    fila[i] = v
        salida.append(fila)

    if n_vacias:
        log(f"    Descartadas {n_vacias} filas vacias")
    if n_filtradas:
        log(f"    Descartadas {n_filtradas} filas por el filtro "
            f"{sorted(filtro['valores'])}")
    if n_ceros:
        log(f"    Descartadas {n_ceros} filas con 0 (o vacio) en las columnas "
            f"de control")
    if forzar:
        detalle = ", ".join(f"columna {i + 1} = {v!r}" for i, v in forzar.items())
        log(f"    Valor forzado en las {len(salida)} filas: {detalle}")
    log(f"    Filas utiles: {len(salida)}")
    return salida


def _leer_fuente_xlwings(app, ruta_xlsm, cfg, log):
    """El camino de antes: abrir el archivo con Excel. Queda como respaldo."""
    import xlwings as xw  # noqa: F401  (se importa en el hilo de trabajo)

    wb = None
    try:
        log(f"    Abriendo {Path(ruta_xlsm).name} (solo lectura)...")
        wb = app.books.open(str(ruta_xlsm), read_only=True, update_links=False)
        sh = buscar_hoja(wb, cfg["hoja"])
        if sh is None:
            raise RuntimeError(f"No se encontro la hoja '{cfg['hoja']}' en {Path(ruta_xlsm).name}")
        log(f"    Hoja: {sh.name}")

        fila_ini = cfg["fila_ini"]
        cols = [(col_letra_a_num(a), col_letra_a_num(b)) for a, b in cfg["bloques"]]

        # ultima fila: se mira la primera columna del primer bloque y la ultima del ultimo
        candidatas = [cols[0][0], cols[-1][1]]
        fila_fin = max(ultima_fila(sh, c, fila_ini) for c in candidatas)
        if fila_fin < fila_ini:
            log("    ADVERTENCIA: no hay datos bajo el encabezado.")
            return []
        log(f"    Filas {fila_ini} a {fila_fin} ({fila_fin - fila_ini + 1} filas leidas)")

        bloques = [leer_bloque(sh, fila_ini, c1, fila_fin, c2) for c1, c2 in cols]

        n = fila_fin - fila_ini + 1
        matriz = []
        for i in range(n):
            fila = []
            for b in bloques:
                fila.extend(b[i] if i < len(b) else [])
            matriz.append(fila)
        return matriz
    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# ACCESS (pyodbc)
# ---------------------------------------------------------------------------
def driver_access():
    """El nombre del driver ODBC de Access. Lanza con un mensaje util si no hay.

    El mensaje LISTA los drivers que pyodbc si ve, y dice que Python se esta
    usando. Sin eso no se puede distinguir entre "no hay ningun driver ODBC"
    (instalacion rota o pyodbc mal), "hay drivers pero ninguno de Access"
    (falta el Access Database Engine) y "hay uno de Access pero de otra
    arquitectura" (Office de 32 bits con Python de 64).

    Tambien importa cuando el mismo script funciono antes en el mismo PC: casi
    siempre significa que se lanzo con OTRO Python (otra instalacion, otro
    entorno virtual, o pythonw vs python), y por eso se muestra el ejecutable.
    """
    import pyodbc
    todos = list(pyodbc.drivers())
    drivers = [d for d in todos if "access" in d.lower()]
    if not drivers:
        bits = "64" if sys.maxsize > 2 ** 32 else "32"
        otra = "32" if bits == "64" else "64"
        detalle = ("pyodbc NO ve NINGÚN driver ODBC. Suele ser una instalación "
                   "rota de pyodbc o del sistema."
                   if not todos else
                   "pyodbc ve estos drivers, pero ninguno es de Access:\n   - "
                   + "\n   - ".join(todos))
        raise RuntimeError(
            "No se encontró el driver ODBC de Access.\n\n"
            f"{detalle}\n\n"
            f"Python: {bits} bits\n"
            f"   {sys.executable}\n\n"
            f"Si este mismo script funcionó antes en este PC, lo más probable es "
            f"que se haya lanzado con OTRO Python. Comprobá que sea el mismo "
            f"ejecutable de arriba.\n\n"
            f"Si no, hay que instalar 'Microsoft Access Database Engine' de "
            f"{bits} bits (el de {otra} bits no sirve para este Python)."
        )
    # Se prefiere el moderno, que abre .mdb y .accdb. El viejo solo abre .mdb.
    for d in drivers:
        if "*.mdb, *.accdb" in d:
            return d
    return drivers[0]


def conectar_access(ruta_mdb):
    import pyodbc
    drv = driver_access()
    cs = f"DRIVER={{{drv}}};DBQ={ruta_mdb};ExtendedAnsiSQL=1;"
    cn = pyodbc.connect(cs, autocommit=False)
    return cn, drv


def columnas_tabla(cur):
    """Devuelve [(nombre, tipo_python)] de la tabla, en orden."""
    cur.execute(f"SELECT * FROM [{TABLA_ACCESS}] WHERE 1=0")
    return [(d[0], d[1]) for d in cur.description]


def mapear_columnas(cols_tabla, log):
    """Elige las 5 columnas destino: por nombre si calzan, si no las 5 primeras."""
    nombres = [c[0] for c in cols_tabla]
    norm = {normalizar(n): n for n in nombres}
    elegidas = []
    for esperada in COLUMNAS_ESPERADAS:
        real = norm.get(normalizar(esperada))
        if real is None:
            elegidas = []
            break
        elegidas.append(real)

    if not elegidas:
        log("    Los nombres de columna no calzan exactamente; se usan las primeras 5 en orden.")
        candidatas = [n for n in nombres if normalizar(n) not in ("id", "id_", "codigo", "correlativo")]
        elegidas = candidatas[:5]

    if len(elegidas) != 5:
        raise RuntimeError(
            f"La tabla [{TABLA_ACCESS}] no tiene 5 columnas utilizables. "
            f"Columnas detectadas: {nombres}"
        )
    tipos = {c[0]: c[1] for c in cols_tabla}
    return [(n, tipos.get(n, str)) for n in elegidas]


def a_numero(valor):
    """Convierte a int/float lo que venga (numero, o texto con formato chileno/ingles).
    Devuelve None si no se puede."""
    if es_vacio(valor):
        return None
    if isinstance(valor, bool):
        return int(valor)
    if isinstance(valor, (int, float, decimal.Decimal)):
        try:
            f = float(valor)
        except Exception:
            return None
        return int(f) if f.is_integer() else f
    if not isinstance(valor, str):
        return None

    # limpieza: espacios (incluye no-separable), signo $, %, comillas, apostrofos
    t = re.sub(r"[\s\u00a0\u202f'\"$%]", "", valor).strip()
    if t == "":
        return None
    negativo = t.startswith("(") and t.endswith(")")     # (1.234) = -1234
    if negativo:
        t = t[1:-1]

    hay_punto, hay_coma = "." in t, "," in t
    if hay_punto and hay_coma:
        # el separador decimal es el que aparece mas a la derecha
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif hay_coma:
        t = t.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", t) else t.replace(",", ".")
    elif hay_punto:
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):      # 1.234.567 = miles
            t = t.replace(".", "")

    try:
        f = float(t)
    except Exception:
        return None
    if negativo:
        f = -f
    return int(f) if f.is_integer() else f


def clave_a_numero(valor):
    """La columna 'Clave Año_Mes' SIEMPRE se entrega como numero (AAAAMM).
    Acepta 202506, '202506', '2025-06', '06/2025', ' 202.506 ' o una fecha."""
    if es_vacio(valor):
        return None
    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor.year * 100 + valor.month
    if isinstance(valor, str):
        t = re.sub(r"[\s\u00a0]", "", valor)
        m = re.fullmatch(r"(\d{4})[-/_.](\d{1,2})", t)          # 2025-06
        if m:
            return int(m.group(1)) * 100 + int(m.group(2))
        m = re.fullmatch(r"(\d{1,2})[-/_.](\d{4})", t)          # 06-2025
        if m:
            return int(m.group(2)) * 100 + int(m.group(1))
        m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", t)   # 2025-06-01
        if m:
            return int(m.group(1)) * 100 + int(m.group(2))
        m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", t)   # 01-06-2025
        if m:
            return int(m.group(3)) * 100 + int(m.group(2))
    n = a_numero(valor)
    if n is None:
        return None
    return int(n) if float(n).is_integer() else n


def coercionar(valor, tipo):
    """Adapta el valor leido de Excel al tipo de la columna en Access."""
    if es_vacio(valor):
        return None

    if tipo in (int,):
        n = a_numero(valor)
        return None if n is None else int(round(n))
    if tipo in (float, decimal.Decimal):
        n = a_numero(valor)
        return None if n is None else float(n)
    if tipo in (datetime.datetime, datetime.date):
        if isinstance(valor, (datetime.datetime, datetime.date)):
            return valor
        return None
    if tipo is bool:
        return bool(valor)

    # texto
    if isinstance(valor, float) and float(valor).is_integer():
        return str(int(valor))
    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor.strftime("%Y-%m-%d")
    return str(valor).strip()


def coercionar_clave(valor, tipo):
    """Igual que coercionar pero forzando la clave a numero, sea cual sea el tipo destino."""
    n = clave_a_numero(valor)
    if n is None:
        return None
    if tipo is str:                      # si el campo del Access fuera texto: numero sin '.0'
        return str(int(n)) if float(n).is_integer() else str(n)
    if tipo in (int,):
        return int(round(n))
    return n


def armar_lote(filas, destino, log):
    """Convierte las filas leidas a los tipos del Access. La columna 1 (clave) va a numero.
    Devuelve (lote, excluidas)."""
    lote = []
    n_texto = 0
    ejemplos = []
    excluidas = []
    for f in filas:
        crudo = f[0]
        clave = coercionar_clave(crudo, destino[0][1])
        if clave is None:
            excluidas.append(repr(crudo))
            continue
        if not isinstance(crudo, (int, float, decimal.Decimal)) or isinstance(crudo, bool):
            n_texto += 1
            if len(ejemplos) < 3:
                ejemplos.append(f"{crudo!r} -> {clave}")
        resto = [coercionar(f[j], destino[j][1]) for j in range(1, 5)]
        lote.append(tuple([clave] + resto))

    if n_texto:
        log(f"    Clave Año_Mes: {n_texto} valores no venian como numero y se convirtieron"
            f"  ({'; '.join(ejemplos)})")
    if excluidas:
        log(f"    ADVERTENCIA: {len(excluidas)} filas EXCLUIDAS porque la clave no se pudo "
            f"convertir a numero  ({'; '.join(excluidas[:5])}"
            f"{' ...' if len(excluidas) > 5 else ''})")
    return lote, excluidas


def clave_tipo(t):
    """Clave homogenea para comparar tipos entre el Access y los Excel."""
    if t is None or (isinstance(t, str) and t.strip() == ""):
        return "(vacio)"
    if isinstance(t, float) and float(t).is_integer():
        return str(int(t))
    return str(t).strip()


def estado_access(cur, col_tipo, col_valor, con_suma=True):
    """{clave_tipo: {'n': filas, 'suma': monto o None}} agrupado por Tipo_sobrecosto."""
    if con_suma:
        try:
            cur.execute(f"SELECT [{col_tipo}], COUNT(*), SUM([{col_valor}]) "
                        f"FROM [{TABLA_ACCESS}] GROUP BY [{col_tipo}]")
            out = {}
            for r in cur.fetchall():
                suma = None
                try:
                    suma = float(r[2]) if r[2] is not None else 0.0
                except Exception:
                    suma = None
                out[clave_tipo(r[0])] = {"n": int(r[1]), "suma": suma}
            return out
        except Exception:
            pass  # la columna de monto no es numerica -> solo conteos
    cur.execute(f"SELECT [{col_tipo}], COUNT(*) FROM [{TABLA_ACCESS}] GROUP BY [{col_tipo}]")
    return {clave_tipo(r[0]): {"n": int(r[1]), "suma": None} for r in cur.fetchall()}


def fmt_int(n, signo=False):
    if n is None:
        return "-"
    s = f"{abs(int(n)):,}".replace(",", ".")
    if signo:
        return ("+" if n > 0 else ("-" if n < 0 else " ")) + s
    return ("-" if n < 0 else "") + s


def fmt_monto(v, signo=False):
    if v is None:
        return "-"
    s = f"{abs(float(v)):,.1f}".replace(",", "@").replace(".", ",").replace("@", ".")
    if signo:
        return ("+" if v > 1e-9 else ("-" if v < -1e-9 else " ")) + s
    return ("-" if v < 0 else "") + s


def resumen_cambios(antes, despues, log, proyectado=False):
    """Detalle de diferencias por tipo y total, antes vs despues (o proyectado)."""
    etiqueta = "PROYECTADO (modo prueba, sin escribir)" if proyectado else "REAL"
    log("")
    log("=" * 96)
    log(f"  DETALLE DE DIFERENCIAS - {etiqueta}")
    log("=" * 96)
    log(f"  {'Tipo_sobrecosto':<28} {'Filas antes':>12} {'Filas desp.':>12} {'Dif':>10}"
        f" {'Monto antes':>14} {'Monto desp.':>14} {'Dif monto':>14}")
    log("  " + "-" * 94)

    tipos = sorted(set(antes) | set(despues))
    tot = {"na": 0, "nd": 0, "sa": 0.0, "sd": 0.0, "hay_suma": False}

    for t in tipos:
        a = antes.get(t, {"n": 0, "suma": 0.0})
        d = despues.get(t, {"n": 0, "suma": 0.0})
        na, nd = a["n"], d["n"]
        sa, sd = a.get("suma"), d.get("suma")
        marca = "   " if (na == nd and (sa is None or sd is None or abs((sd or 0) - (sa or 0)) < 0.05)) else " * "
        dif_s = None if (sa is None or sd is None) else (sd - sa)
        log(f"{marca}{str(t)[:28]:<28} {fmt_int(na):>12} {fmt_int(nd):>12} {fmt_int(nd - na, True):>10}"
            f" {fmt_monto(sa):>14} {fmt_monto(sd):>14} {fmt_monto(dif_s, True):>14}")
        tot["na"] += na
        tot["nd"] += nd
        if sa is not None:
            tot["sa"] += sa
            tot["hay_suma"] = True
        if sd is not None:
            tot["sd"] += sd
            tot["hay_suma"] = True

    log("  " + "-" * 94)
    sa_t = tot["sa"] if tot["hay_suma"] else None
    sd_t = tot["sd"] if tot["hay_suma"] else None
    dif_t = None if sa_t is None or sd_t is None else sd_t - sa_t
    log(f"  {'TOTAL':<28} {fmt_int(tot['na']):>12} {fmt_int(tot['nd']):>12}"
        f" {fmt_int(tot['nd'] - tot['na'], True):>10}"
        f" {fmt_monto(sa_t):>14} {fmt_monto(sd_t):>14} {fmt_monto(dif_t, True):>14}")
    log("  (* = tipo con cambios)")

    solo_antes = [t for t in antes if t not in despues]
    solo_despues = [t for t in despues if t not in antes]
    if solo_antes:
        log(f"  Tipos que desaparecieron: {solo_antes}")
    if solo_despues:
        log(f"  Tipos nuevos: {solo_despues}")


# ---------------------------------------------------------------------------
# PROCESO PRINCIPAL
# ---------------------------------------------------------------------------
def proceso(ruta_mdb, archivos, seleccion, solo_lectura, log, progreso,
            fuentes=None, borrar_todo=False):
    """archivos: {clave: ruta_xlsm}; seleccion: lista de claves a actualizar.

    fuentes: diccionario de configuracion de las fuentes. Por omision el FUENTES
    de este modulo (SSCC/CO/CCA); Actualiza_Energia.py y Actualiza_Access_P9.py
    pasan el suyo para reusar todo este motor de Access sin duplicarlo.

    borrar_todo: vaciar la tabla entera en vez de borrar solo los
    Tipo_sobrecosto que vienen en los Excel. Lo usa el .mdb de la planilla 9,
    que se arma COMPLETO desde sus fuentes; ahi el borrado por tipo dejaria
    filas viejas si un tipo desaparece del origen. El vaciado ocurre UNA sola
    vez, antes de insertar (ver ya_vaciada mas abajo).

    Si se agrega un parametro nuevo aca, hay que sumarlo tambien a CAPACIDADES,
    o los scripts que lo usen se caen con un TypeError en vez de decir "copia el
    archivo actualizado"."""
    fuentes = fuentes if fuentes is not None else FUENTES
    import pythoncom
    import xlwings as xw

    pythoncom.CoInitialize()
    app = None
    cn = None
    try:
        # ---- 1) leer los Excel ----
        datos = {}
        log("=" * 70)
        log("PASO 1/2 - Lectura de los Excel (no se modifica ni guarda nada)")
        log("=" * 70)
        # Excel se abre SOLO SI hace falta. Con las planillas normales
        # (.xlsx/.xlsm) el lector rapido alcanza y no se abre nunca. Antes se
        # abria siempre y se cerraba aca, y al cerrarlo se mataba el proceso COM
        # que el script llamador podia estar usando: lo que venia despues
        # reventaba con "El objeto no esta conectado al servidor".
        caja = {"app": None}

        def dame_app():
            if caja["app"] is None:
                log("    (abriendo Excel: hace falta para esta fuente)")
                a = xw.App(visible=False, add_book=False)
                a.display_alerts = False
                a.screen_updating = False
                caja["app"] = a
            return caja["app"]

        try:
            for i, clave in enumerate(seleccion):
                log(f"\n  [{clave}]")
                datos[clave] = leer_fuente(dame_app, archivos[clave],
                                           fuentes[clave], log)
                progreso(int(40 * (i + 1) / len(seleccion)))
        finally:
            try:
                if caja["app"] is not None:
                    caja["app"].quit()
            except Exception:
                pass
            caja["app"] = None

        # ---- 2) escribir en Access ----
        log("")
        log("=" * 70)
        log("PASO 2/2 - Actualizacion del Access")
        log("=" * 70)
        cn, drv = conectar_access(ruta_mdb)
        log(f"  Driver: {drv}")
        cur = cn.cursor()

        cols_tabla = columnas_tabla(cur)
        destino = mapear_columnas(cols_tabla, log)
        nombres_destino = [n for n, _ in destino]
        log(f"  Columnas destino: {nombres_destino}")
        col_tipo = nombres_destino[1]
        col_valor = nombres_destino[4]
        tipo_py_valor = destino[4][1]
        suma_numerica = tipo_py_valor in (int, float, decimal.Decimal)

        antes = estado_access(cur, col_tipo, col_valor, con_suma=suma_numerica)
        log("  Estado actual del Access:")
        for k, v in sorted(antes.items(), key=lambda x: str(x[0])):
            log(f"    {str(k)[:28]:<28} {fmt_int(v['n']):>12} filas   monto {fmt_monto(v['suma'])}")
        log(f"    {'TOTAL':<28} {fmt_int(sum(v['n'] for v in antes.values())):>12} filas")

        sql_ins = (f"INSERT INTO [{TABLA_ACCESS}] ("
                   + ", ".join(f"[{n}]" for n in nombres_destino)
                   + ") VALUES (?, ?, ?, ?, ?)")

        # proyeccion: como quedaria el Access (se usa para el modo prueba)
        # Con borrar_todo se parte de vacio: no queda nada de lo que habia.
        proy = {} if borrar_todo else {t: dict(v) for t, v in antes.items()}

        total_ins = 0
        # Con borrar_todo la tabla se vacia UNA sola vez, antes del bucle. Hacerlo
        # dentro borraria en cada vuelta lo que insertaron las anteriores.
        ya_vaciada = False
        for i, clave in enumerate(seleccion):
            filas = datos[clave]
            log(f"\n  [{clave}]")
            if not filas:
                log("    Sin filas utiles -> NO se borra ni se inserta nada (fuente omitida).")
                continue

            # se convierte todo al tipo de dato de la tabla antes de comparar/insertar
            # (la columna Clave Año_Mes se fuerza a numero)
            lote, excluidas = armar_lote(filas, destino, log)
            if not lote:
                log("    Sin filas validas despues de convertir la clave -> fuente omitida.")
                continue

            # agregado de la fuente por Tipo_sobrecosto (columna 2)
            agg = {}
            valores_tipo = []          # valores crudos (tipados) para el DELETE
            for r in lote:
                if r[1] not in valores_tipo:
                    valores_tipo.append(r[1])
                a = agg.setdefault(clave_tipo(r[1]), {"n": 0, "suma": 0.0 if suma_numerica else None})
                a["n"] += 1
                if suma_numerica:
                    try:
                        a["suma"] += float(r[4]) if r[4] is not None else 0.0
                    except Exception:
                        pass
            log(f"    Tipo_sobrecosto en el Excel: {list(agg.keys())}")
            for t, v in agg.items():
                log(f"      {str(t)[:28]:<28} {fmt_int(v['n']):>12} filas   monto {fmt_monto(v['suma'])}")
                proy[t] = dict(v)      # el tipo se reemplaza completo

            if solo_lectura:
                log(f"    [MODO PRUEBA] se insertarian {len(lote)} filas.")
                progreso(40 + int(60 * (i + 1) / len(seleccion)))
                continue

            # borrar: la tabla entera, o solo los tipos que traen los Excel
            borradas = 0
            if borrar_todo:
                if not ya_vaciada:
                    cur.execute(f"DELETE FROM [{TABLA_ACCESS}]")
                    borradas = cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0
                    log(f"    Tabla vaciada: {borradas} fila(s) borradas "
                        f"(una sola vez, antes de insertar las {len(seleccion)} "
                        f"fuentes)")
                    ya_vaciada = True
                valores_tipo = []
            for t in valores_tipo:
                if t is None:
                    cur.execute(f"DELETE FROM [{TABLA_ACCESS}] WHERE [{col_tipo}] IS NULL")
                else:
                    cur.execute(f"DELETE FROM [{TABLA_ACCESS}] WHERE [{col_tipo}] = ?", t)
                    if clave_tipo(t) not in antes:
                        log(f"    OJO: el tipo '{t}' no existia en el Access (no se borro nada de el).")
                borradas += cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0
            log(f"    Filas borradas: {borradas}")

            paso = 500
            for k in range(0, len(lote), paso):
                cur.executemany(sql_ins, lote[k:k + paso])
                pct = 40 + int(60 * (i + (min(k + paso, len(lote)) / len(lote))) / len(seleccion))
                progreso(pct)
            log(f"    Filas insertadas: {len(lote)}")
            total_ins += len(lote)

        if solo_lectura:
            cn.rollback()
            log("\n  [MODO PRUEBA] No se escribio nada en el Access.")
            resumen_cambios(antes, proy, log, proyectado=True)
        else:
            cn.commit()
            log(f"\n  COMMIT OK. Total insertado: {total_ins} filas.")
            despues = estado_access(cur, col_tipo, col_valor, con_suma=suma_numerica)
            resumen_cambios(antes, despues, log, proyectado=False)

        progreso(100)
        return True, total_ins

    except Exception as e:
        if cn is not None:
            try:
                cn.rollback()
                log("  ROLLBACK: no se aplico ningun cambio al Access.")
            except Exception:
                pass
        log(f"\nERROR: {e}")
        log(traceback.format_exc())
        return False, 0
    finally:
        try:
            if cn is not None:
                cn.close()
        except Exception:
            pass
        try:
            if app is not None:
                app.quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# VENTANA
# ---------------------------------------------------------------------------
def actualizar_color_label(lbl, valor, es_archivo=False):
    if not valor or valor.startswith("["):
        lbl.config(fg="red")
    elif es_archivo and Path(valor).is_file():
        lbl.config(fg="blue")
    elif not es_archivo and Path(valor).is_dir():
        lbl.config(fg="blue")
    else:
        lbl.config(fg="red")


def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    # modo["traspaso"] se apaga si el usuario elige un .mdb a mano: desde ese
    # momento la ventana vuelve a buscar los archivos por su cuenta.
    modo = {"traspaso": traspaso is not None}

    root = tk.Tk()
    root.title("Actualizar Access Sobrecostos"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("980x760")

    # ---- 1) botonera fija abajo ----
    frame_btns_fijo = tk.Frame(root)
    frame_btns_fijo.pack(side="bottom", fill="x", pady=8)

    # ---- 2) canvas con scroll ----
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

    if traspaso:
        fr_aviso = tk.Frame(contenedor, bg="#fff4c2", bd=1, relief="solid")
        fr_aviso.pack(fill="x", padx=20, pady=(10, 0))
        tk.Label(fr_aviso,
                 text=f"Mes {traspaso.get('aamm') or '?'} — enviado por el Revisor",
                 bg="#fff4c2", font=("Segoe UI", 11, "bold")).pack(pady=(6, 0))
        tk.Label(fr_aviso,
                 text="Las rutas las resolvio el Revisor; este script no las vuelve a buscar.\n"
                      "Marca que fuentes quieres reemplazar: los casilleros arrancan en blanco.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    tk.Label(contenedor, text="Actualizacion de la tabla [Sobrecostos] del Access",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(contenedor, text="Los .xlsm se abren solo en lectura: no se modifican ni se guardan.",
             font=("Segoe UI", 8), fg="#555555").pack(pady=(0, 8))

    # ---- variables ----
    var_mdb = tk.StringVar(value=(traspaso or {}).get("rutas", {}).get("mdb_sscc")
                                 or cfg.get("mdb", "[selecciona el Access .mdb]"))
    var_entregables = tk.StringVar(value="[se detecta desde el .mdb]")
    vars_arch = {k: tk.StringVar(value="[pendiente]") for k in ORDEN_FUENTES}
    # Viniendo del Revisor arrancan en blanco a proposito: el usuario tiene que
    # decir que fuentes quiere reemplazar. A mano se recuerda la ultima seleccion
    # como antes.
    vars_sel = {k: tk.BooleanVar(value=False if modo["traspaso"]
                                 else bool(cfg.get(f"sel_{k}", k == "SSCC")))
                for k in ORDEN_FUENTES}
    var_prueba = tk.BooleanVar(value=False)
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    labels = {}

    # ---- selector del .mdb ----
    fr_mdb = tk.LabelFrame(contenedor, text="1) Access .mdb (carpeta '01 Sobrecostos')",
                           padx=10, pady=6)
    fr_mdb.pack(fill="x", padx=20, pady=4)
    lbl_mdb = tk.Label(fr_mdb, textvariable=var_mdb, wraplength=800, justify="center",
                       cursor="hand2", font=("Segoe UI", 9))
    lbl_mdb.pack()
    lbl_mdb.bind("<Button-1>", lambda e: abrir_en_explorador(var_mdb.get(), es_archivo=True))
    labels["mdb"] = lbl_mdb

    # ---- selector carpeta 00 Entregables ----
    fr_ent = tk.LabelFrame(contenedor, text="2) Carpeta '00 Entregables' (automatica, editable)",
                           padx=10, pady=6)
    fr_ent.pack(fill="x", padx=20, pady=4)
    lbl_ent = tk.Label(fr_ent, textvariable=var_entregables, wraplength=800, justify="center",
                       cursor="hand2", font=("Segoe UI", 9))
    lbl_ent.pack()
    lbl_ent.bind("<Button-1>", lambda e: abrir_en_explorador(var_entregables.get()))
    labels["ent"] = lbl_ent

    # ---- fuentes + checkboxes ----
    fr_src = tk.LabelFrame(contenedor, text="3) Que actualizar (puedes marcar 1, 2 o 3)",
                           padx=10, pady=6)
    fr_src.pack(fill="x", padx=20, pady=4)
    for k in ORDEN_FUENTES:
        fila = tk.Frame(fr_src)
        fila.pack(fill="x", pady=2)
        tk.Checkbutton(fila, text=k, variable=vars_sel[k], width=6, anchor="w",
                       font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(fila, text=FUENTES[k]["etiqueta"], font=("Segoe UI", 8),
                 fg="#555555", anchor="w").pack(side="left")
        lb = tk.Label(fr_src, textvariable=vars_arch[k], wraplength=780, justify="left",
                      anchor="w", cursor="hand2", font=("Consolas", 8))
        lb.pack(fill="x", padx=(60, 0))
        lb.bind("<Button-1>", lambda e, kk=k: abrir_en_explorador(vars_arch[kk].get(), es_archivo=True))
        labels[k] = lb

    tk.Checkbutton(contenedor, text="Modo prueba (solo leer y contar, NO escribir en el Access)",
                   variable=var_prueba).pack(pady=(6, 2))

    # ---- progreso ----
    fr_prog = tk.LabelFrame(contenedor, text="Progreso", padx=10, pady=6)
    fr_prog.pack(fill="both", expand=True, padx=20, pady=6)
    tk.Label(fr_prog, textvariable=var_estado, anchor="w").pack(fill="x")
    tk.Label(fr_prog, textvariable=var_tiempo, font=("Consolas", 10, "bold"),
             fg="#2d7a2d").pack()
    progress = ttk.Progressbar(fr_prog, mode="determinate", length=400, maximum=100)
    progress.pack(fill="x", pady=4)
    txt_log = tk.Text(fr_prog, height=18, font=("Consolas", 8))
    txt_log.pack(fill="both", expand=True)

    cola = queue.Queue()

    def log(msg):
        cola.put(("log", str(msg)))

    def progreso(pct):
        cola.put(("pct", pct))

    def bombear():
        try:
            while True:
                tipo, val = cola.get_nowait()
                if tipo == "log":
                    txt_log.insert("end", val + "\n")
                    txt_log.see("end")
                elif tipo == "pct":
                    progress["value"] = val
                elif tipo == "estado":
                    var_estado.set(val)
        except queue.Empty:
            pass
        root.after(120, bombear)

    timer_state = {"running": False, "t_ini": 0.0}

    def tick():
        if timer_state["running"]:
            var_tiempo.set(fmt_tiempo(time.time() - timer_state["t_ini"]))
            root.after(500, tick)

    # ---- deteccion de archivos ----
    MAPA_TRASPASO = {"SSCC": "calculo_sscc_maestro",
                     "CO": "calculo_co",
                     "CCA": "consolidado_cca"}

    def refrescar(*_):
        if modo["traspaso"]:
            # Las rutas las resolvio el Revisor: no se busca nada de nuevo.
            dadas = traspaso.get("rutas", {})
            mdb = dadas.get("mdb_sscc")
            if mdb:
                var_mdb.set(str(mdb))
            var_entregables.set("[no se usa: rutas enviadas por el Revisor]")
            for k in ORDEN_FUENTES:
                val = dadas.get(MAPA_TRASPASO[k])
                vars_arch[k].set(str(val) if val
                                 else f"[el Revisor no mando {MAPA_TRASPASO[k]}]")
            actualizar_color_label(labels["mdb"], var_mdb.get(), es_archivo=True)
            for k in ORDEN_FUENTES:
                actualizar_color_label(labels[k], vars_arch[k].get(), es_archivo=True)
            return

        ruta = var_mdb.get()
        # carpeta_mdb es "01 Sobrecostos": de ahi sale el MAESTRO del SSCC.
        carpeta_mdb = None
        if ruta and not ruta.startswith("[") and Path(ruta).is_file():
            carpeta_mdb = Path(ruta).parent
            raiz = carpeta_mdb.parent
            ent = buscar_carpeta(raiz, "00 Entregables")
            var_entregables.set(str(ent) if ent else "[no encontrada: selecciona manualmente]")
        base = var_entregables.get()
        base = Path(base) if base and not base.startswith("[") and Path(base).is_dir() else None
        for k in ORDEN_FUENTES:
            f = None
            if FUENTES[k].get("base") == "mdb":
                if carpeta_mdb:
                    f = buscar_archivo(carpeta_mdb, FUENTES[k]["patron"])
            elif base:
                sub = buscar_carpeta(base, FUENTES[k]["carpeta"])
                if sub:
                    f = buscar_archivo(sub, FUENTES[k]["patron"])
            vars_arch[k].set(str(f) if f else "[no encontrado]")
        actualizar_color_label(labels["mdb"], var_mdb.get(), es_archivo=True)
        actualizar_color_label(labels["ent"], var_entregables.get())
        for k in ORDEN_FUENTES:
            actualizar_color_label(labels[k], vars_arch[k].get(), es_archivo=True)

    def sel_mdb():
        ini = cfg.get("mdb", "")
        ini = str(Path(ini).parent) if ini and Path(ini).exists() else ""
        r = filedialog.askopenfilename(title="Selecciona el Access", initialdir=ini,
                                       filetypes=[("Access", "*.mdb *.accdb"), ("Todos", "*.*")])
        if r:
            var_mdb.set(r)
            guardar_config({"mdb": r})
            cfg["mdb"] = r
            # Si el usuario elige el .mdb a mano, manda el, no el traspaso.
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Access elegido a mano: se dejan de usar las rutas del Revisor.")
            refrescar()

    def sel_ent():
        ini = var_entregables.get()
        ini = ini if ini and not ini.startswith("[") and Path(ini).is_dir() else ""
        r = filedialog.askdirectory(title="Selecciona la carpeta 00 Entregables", initialdir=ini)
        if r:
            var_entregables.set(r)
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Carpeta elegida a mano: se dejan de usar las rutas del Revisor.")
            refrescar()

    tk.Button(fr_mdb, text="Examinar", command=sel_mdb).pack(pady=(4, 0))
    tk.Button(fr_ent, text="Examinar", command=sel_ent).pack(pady=(4, 0))

    # ---- ejecutar ----
    def ejecutar():
        ruta_mdb = var_mdb.get()
        if not ruta_mdb or ruta_mdb.startswith("[") or not Path(ruta_mdb).is_file():
            messagebox.showerror("Falta el Access", "Selecciona un archivo .mdb valido.")
            return
        seleccion = [k for k in ORDEN_FUENTES if vars_sel[k].get()]
        if not seleccion:
            messagebox.showerror("Sin seleccion", "Marca al menos una fuente (SSCC, CO o CCA).")
            return
        archivos = {}
        for k in seleccion:
            r = vars_arch[k].get()
            if not r or r.startswith("[") or not Path(r).is_file():
                messagebox.showerror("Archivo faltante",
                                     f"No se encontro el Excel de {k}.\nRevisa la carpeta 00 Entregables.")
                return
            archivos[k] = r
        # Solo se recuerda la seleccion cuando el usuario abrio la ventana a mano.
        # Viniendo del Revisor no se guarda, para que la corrida siguiente siga
        # arrancando en blanco.
        if not modo["traspaso"]:
            guardar_config({f"sel_{k}": vars_sel[k].get() for k in ORDEN_FUENTES})

        if not var_prueba.get():
            det = "\n".join(f"  - {k}: {Path(archivos[k]).name}" for k in seleccion)
            if not messagebox.askyesno(
                    "Confirmar",
                    "Se REEMPLAZARA en el Access la informacion de:\n\n" + det +
                    "\n\nLos demas tipos quedan intactos.\n\n(Se recomienda tener un respaldo del .mdb)"):
                return

        txt_log.delete("1.0", "end")
        btn.config(state="disabled", bg="#aaaaaa")
        progress["value"] = 0
        var_estado.set("Procesando...")
        timer_state["running"] = True
        timer_state["t_ini"] = time.time()
        tick()

        def trabajo():
            ok, n = proceso(ruta_mdb, archivos, seleccion, var_prueba.get(), log, progreso)
            def fin():
                timer_state["running"] = False
                btn.config(state="normal", bg="#2d7a2d")
                if ok:
                    var_estado.set(f"Listo - {n} filas insertadas" if not var_prueba.get()
                                   else "Listo - modo prueba (sin cambios)")
                    messagebox.showinfo("Listo", f"Proceso terminado.\nFilas insertadas: {n}"
                                        if not var_prueba.get() else "Modo prueba terminado.")
                else:
                    var_estado.set("Termino con errores - revisa el log")
                    messagebox.showerror("Error", "El proceso fallo. Revisa el log de la ventana.")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns_fijo, text="ACTUALIZAR ACCESS", bg="#2d7a2d", fg="white",
                    font=("Segoe UI", 10, "bold"), command=ejecutar)
    btn.pack(side="left", padx=8, expand=True)
    btn_ref = tk.Button(frame_btns_fijo, text="Refrescar rutas", command=refrescar)
    btn_ref.pack(side="left", padx=8)
    btn_salir = tk.Button(frame_btns_fijo, text="Salir", command=root.destroy)
    btn_salir.pack(side="left", padx=8)

    refrescar()
    bombear()
    root.mainloop()


# Se comprueba al importar: es barato y evita entregar un archivo que promete
# algo que no cumple.
_verificar_capacidades()


if __name__ == "__main__":
    main()
