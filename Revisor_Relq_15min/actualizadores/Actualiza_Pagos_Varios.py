# =============================================================================
#  Arma pagos_varios_AAMM.csv desde las planillas 3, 5 y 6  —  VERSION 15 MIN
# =============================================================================
#  Reemplaza, en el modo 15 minutos, a Actualiza_Access_P9.py: los mismos cuatro
#  bloques de las planillas 3, 5 y 6 ya no se cargan en el Access
#  Ocupar_este_para_Reliquidacion, sino que se escriben en un CSV:
#
#      04 Planilla 9/pagos_varios_AAMM.csv
#      Fecha_Hora,CONFIGURACION,Central,Pago,Tipo_pago,Bloque_15min
#
#  DE DONDE SALE CADA DATO  —  confirmado por la usuaria el 2026-10-28
#  ----------------------------------------------------------------------------
#   | Fuente   | Archivo | Hoja                       | Fecha | Central | Pago | Tipo | Bloque |
#   |----------|---------|----------------------------|-------|---------|------|------|--------|
#   | SUBASTAS | 3_      | DB                         |  (*)  |   BF    |  BH  |  BE  |   BG   |
#   | CRA      | 5_      | CÁLCULO_CRA                |  AV   |   AX    |  AY  |  AZ  |   BA   |
#   | CO_ERNC  | 6_      | ENERGIA_Y_CALCULO_CO_ERNC  |  BB   |   BE    |  BF  |  BG  |   BD   |
#   | REA      | 6_      | CALCULO_REA_CENTRAL        |  matriz (ver abajo)                    |
#
#  Son las columnas de abril 2026 (la 5, las de septiembre) y solo se usan si
#  no se encuentran los encabezados (ver abajo).
#  (*) Sin fecha (trae "Clave Año_Mes"): la Fecha_Hora se arma con el bloque
#      ("Hora Mensual") y el mes del archivo.
#  REA (2026-10-29): matriz. Centrales en D11 hacia la derecha hasta la primera
#      vacia; dia en B13:B, bloque diario 1-96 en C13:C, monto en el cruce.
#      Tipo fijo "REA"; Fecha_Hora = dia + (bloque-1)*15 min; Bloque_15min =
#      (dia-1)*96 + bloque.
#
#  LAS COLUMNAS SE BUSCAN POR ENCABEZADO (2026-10-29): en abril las tablas se
#  corrieron de columna en la 3 y en la 6. En las 20 filas de arriba, columnas
#  A:EZ, se busca un "Pago" y, a no mas de 8 columnas de el, "Central",
#  "Tipo_pago" y el bloque ("Bloque_15min", "Bloque", "Periodo" u "Hora
#  Mensual", en ese orden de preferencia); "Fecha_Hora" es opcional. Si no
#  trae fecha (como la 3 o la REA, que traen "Clave Año_Mes"), la fecha se arma
#  con el bloque y el mes. Si hay mas de una tabla que calce, se toma la mas
#  cercana a las columnas de la tabla de abajo. Si no se encuentra ninguna, se
#  usan esas columnas fijas y se avisa.
#
#  REGLAS
#  ----------------------------------------------------------------------------
#  * CONFIGURACION va vacia (asi viene el pagos_varios de referencia).
#  * Fecha_Hora en una sola forma: DD/MM/AAAA HH:MM.
#  * Se descartan las filas con pago 0 o vacio y las sin central (o central 0),
#    igual que hacia el Access.
#  * Orden: Subastas, CRA, CO_ERNC, REA; dentro de cada una, el de la planilla.
#  * CSV con coma, punto decimal, UTF-8 sin BOM y fin de linea CRLF, como el de
#    referencia. Se escribe en un .tmp y se reemplaza: si algo falla, el archivo
#    anterior queda como estaba.
#  * Solo LEE las planillas: se abren como ZIP, sin Excel. Si eso falla (no es
#    OOXML, por ejemplo) se cae a xlwings, en solo lectura.
#
#  Se puede correr solo, o recibir del Revisor la ruta de un JSON de traspaso.
# =============================================================================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from datetime import datetime, timedelta
import csv, io, os, re, subprocess, sys, threading, time, traceback, queue

DIR_SCRIPT = Path(__file__).resolve().parent

# config.json es compartido con el Revisor y el resto de los actualizadores, y
# vive en __config__, junto a Revisor_Relq_15min.
CONFIG_PATH = DIR_SCRIPT.parent.parent / "__config__" / "config.json"


def _morir(titulo, mensaje):
    """Aborta mostrando el motivo en una ventana. Sin esto, lanzado desde el
    Revisor con pythonw (sin consola) moriria callado. No puede vivir en
    __comun__/: es justamente lo que avisa cuando __comun__/ es lo que falta."""
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(titulo, mensaje)
        r.destroy()
    except Exception:
        pass
    print(f"{titulo}\n\n{mensaje}", file=sys.stderr)
    raise SystemExit(1)


# Va detras de _morir a proposito: lanzado por el Revisor con pythonw no hay
# consola, asi que un ImportError suelto moriria callado.
try:
    sys.path.insert(0, str(DIR_SCRIPT.parent.parent))
    from __comun__ import config as _cfg
    from __comun__ import traspaso as _traspaso
    from __comun__ import texto as _texto
    from __comun__ import excel_xml as _excel_xml
except ImportError as e:
    _morir("Falta la carpeta __comun__/",
           "No se pudo cargar __comun__/.\n\n"
           "Tiene que estar la carpeta '__comun__' hermana de Revisor_Relq_15min.\n"
           "Baja el repositorio completo, no los .py sueltos.\n\n"
           f"Carpeta actual: {DIR_SCRIPT}\n\nDetalle: {e}")


# =============================================================================
#  CONFIGURACION
# =============================================================================
CARPETA_P9 = "04 Planilla 9"
ENCABEZADO_CSV = ["Fecha_Hora", "CONFIGURACION", "Central", "Pago",
                  "Tipo_pago", "Bloque_15min"]
FORMATO_FECHA = "%d/%m/%Y %H:%M"
BLOQUES_POR_DIA = 96
EPS_CERO = 1e-12
FILAS_BUSCAR_ENCABEZADO = 20

PLANILLAS = {
    "p3": {"etiqueta": "Planilla 3 — 3_REMUNERACIÓN_SUBASTAS_E_ID",
           "patron": r"^3_remuneracion_subastas"},
    "p5": {"etiqueta": "Planilla 5 — 5_REMUNERACIÓN_CRA",
           "patron": r"^5_remuneracion_cra"},
    "p6": {"etiqueta": "Planilla 6 — 6_REMUNERACIÓN_REA_Y_CO_ERNC",
           "patron": r"^6_remuneracion_rea"},
}
ORDEN_PL = ["p3", "p5", "p6"]

# Los cuatro bloques, en el orden en que van en el CSV. Las letras son el
# respaldo si no se encuentran los encabezados. "fecha": None quiere decir que
# la Fecha_Hora se arma desde el bloque. "fila_datos" tambien es respaldo.
FUENTES = [
    {"nombre": "SUBASTAS", "planilla": "p3", "hoja": "DB", "fila_datos": 3,
     "fecha": None, "central": "BF", "pago": "BH", "tipo": "BE", "bloque": "BG"},
    {"nombre": "CRA", "planilla": "p5", "hoja": "CÁLCULO_CRA", "fila_datos": 9,
     "fecha": "AV", "central": "AX", "pago": "AY", "tipo": "AZ", "bloque": "BA"},
    {"nombre": "CO_ERNC", "planilla": "p6", "hoja": "ENERGIA_Y_CALCULO_CO_ERNC",
     "fila_datos": 9,
     "fecha": "BB", "central": "BE", "pago": "BF", "tipo": "BG", "bloque": "BD"},
    # REA (usuaria, 2026-10-29): ya no es una tabla sino una MATRIZ en la misma
    # hoja. Centrales en la fila 11 desde la D hasta la primera celda vacia a
    # la derecha; dia en B y bloque diario (1-96) en C desde la fila 13; los
    # montos en el cruce. El tipo es fijo "REA". Ver leer_matriz.
    {"nombre": "REA", "planilla": "p6", "hoja": "CALCULO_REA_CENTRAL",
     "matriz": {"fila_centrales": 11, "col_ini": "D", "col_max": "IV",
                "fila_datos": 13, "col_dia": "B", "col_bloque": "C",
                "tipo": "REA"}},
]

# Encabezados aceptados por campo, en orden de preferencia (se comparan sin
# tildes, espacios ni guiones bajos). "pago" es el ancla de la busqueda.
ENCABEZADOS = {
    "pago":    ["Pago"],
    "central": ["Central"],
    "tipo":    ["Tipo_pago", "Tipo de pago"],
    "bloque":  ["Bloque_15min", "Bloque", "Periodo", "Hora Mensual"],
    "fecha":   ["Fecha_Hora", "Fecha"],
}
CAMPOS_OPCIONALES = ("fecha",)
ZONA_ENCABEZADO = ("A", "EZ")       # columnas donde se buscan
MAX_DISTANCIA_ENCABEZADO = 8        # columnas entre "Pago" y los demas

FUENTES_POR_NOMBRE = {f["nombre"]: f for f in FUENTES}

# AAMM del nombre del archivo: '..._2609_R01D.xlsm' -> 2609.
RE_AAMM = re.compile(r"[_\s](\d{4})(?=[_\s.])")


# =============================================================================
#  CONFIG COMPARTIDO
# =============================================================================
get_usuario = _cfg.clave_equipo
escribir_json = _cfg.escribir_json


# Modo 15 minutos: config.json se comparte con la version horaria, que vive en
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


def abrir_en_explorador(ruta, es_archivo=False):
    if not ruta or str(ruta).startswith("["):
        return
    p = Path(ruta)
    if not p.exists():
        p = p.parent
        if not p.exists():
            return
        es_archivo = False
    carpeta = p.parent if es_archivo else p
    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(carpeta)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(carpeta)])
    else:
        subprocess.Popen(["xdg-open", str(carpeta)])


# =============================================================================
#  TRASPASO DESDE EL REVISOR
# =============================================================================
TRASPASO_ORIGEN = _traspaso.ORIGEN
TRASPASO_VERSION_MAX = _traspaso.VERSION_ACTUAL


def leer_traspaso(argv):
    return _traspaso.leer_argumento(argv)


# =============================================================================
#  UTILIDADES
# =============================================================================
normalizar = _texto.suave


def fmt_tiempo(seg):
    seg = int(seg)
    return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"


def aamm_de_nombre(ruta):
    """El AAMM del nombre de un archivo, o None. Un AAMM valido tiene mes 1..12."""
    for m in RE_AAMM.finditer(Path(ruta).stem + "."):
        v = int(m.group(1))
        if 1 <= v % 100 <= 12:
            return v
    return None


def detectar_aamm(rutas):
    """(aamm, motivo). Si los archivos son de meses distintos, devuelve None."""
    vistos = {}
    for k, r in rutas.items():
        v = aamm_de_nombre(r)
        if v:
            vistos.setdefault(v, []).append(k)
    if not vistos:
        return None, "no se pudo sacar el mes de ningún nombre de archivo"
    if len(vistos) > 1:
        det = "; ".join(f"{v}: {', '.join(ks)}" for v, ks in sorted(vistos.items()))
        return None, f"los archivos son de meses DISTINTOS ({det})"
    v = next(iter(vistos))
    return v, f"{v}, sacado del nombre de {len(vistos[v])} archivo(s)"


def buscar_planilla(carpeta_reliq, patron):
    """El archivo mas reciente de 04 Planilla 9 que calce con el patron,
    descartando temporales y copias de Windows."""
    p9 = None
    for sub in Path(carpeta_reliq).iterdir():
        if sub.is_dir() and normalizar(sub.name) == normalizar(CARPETA_P9):
            p9 = sub
            break
    if p9 is None:
        return None
    pat = re.compile(patron)
    cand = [f for f in p9.iterdir()
            if f.is_file() and f.suffix.lower() in (".xlsx", ".xlsm")
            and not f.name.startswith("~$")
            and not re.search(r"(- ?copia( \(\d+\))?|- ?copy|\(\d+\))$",
                              f.stem, re.IGNORECASE)
            and pat.search(normalizar(f.stem))]
    if not cand:
        return None
    return max(cand, key=lambda f: f.stat().st_mtime)


def ruta_salida(carpeta_p9, aamm):
    return Path(carpeta_p9) / f"pagos_varios_{aamm}.csv"


def _numero(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.strip().replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return None
    return None


def _vacio(v):
    return v is None or (isinstance(v, str) and not v.strip())


def fecha_de(v):
    """Lo que venga en la celda de fecha -> datetime redondeado al minuto, o None.
    Por ZIP llega el numero de serie de Excel; por xlwings, un datetime."""
    if isinstance(v, datetime):
        f = v
    elif isinstance(v, (int, float)) and not isinstance(v, bool):
        f = datetime(1899, 12, 30) + timedelta(days=float(v))
    elif isinstance(v, str):
        for fmt in ("%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M", "%Y-%m-%d %H:%M:%S",
                    "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S"):
            try:
                f = datetime.strptime(v.strip(), fmt)
                break
            except ValueError:
                continue
        else:
            return None
    else:
        return None
    # El serie de Excel arrastra decimales (09:14:59.99...): al minuto mas cercano.
    return (f + timedelta(seconds=30)).replace(second=0, microsecond=0)


def bloque_de_fecha(f):
    """(dia-1)*96 + hora*4 + minuto/15 + 1: la regla con la que calza el
    pagos_varios de referencia en todas sus filas. En un mes con cambio de hora
    no calza los dias siguientes: se avisa."""
    return (f.day - 1) * BLOQUES_POR_DIA + f.hour * 4 + f.minute // 15 + 1


def fecha_de_bloque(bloque, anio, mes):
    b = int(bloque) - 1
    dia, resto = divmod(b, BLOQUES_POR_DIA)
    return datetime(anio, mes, dia + 1) + timedelta(minutes=15 * resto)


def fmt_pago(v):
    """Como el de referencia: la representacion mas corta que vuelve al mismo
    numero (3112.342612166663), sin '.0' en los enteros."""
    if float(v).is_integer():
        return str(int(v))
    r = repr(float(v))
    return r if "e" not in r else f"{v:.15f}".rstrip("0").rstrip(".")


# =============================================================================
#  LECTURA
# =============================================================================
def leer_columnas(ruta, hoja, columnas, log):
    """{COL: {fila: valor}} desde la fila 1. Primero sin Excel (ZIP/XML); si eso
    no funciona, con xlwings en solo lectura."""
    d = _excel_xml.leer_columnas_rapido(ruta, hoja, columnas, 1, log)
    if d is not None:
        return d
    log(f"    lectura sin Excel no disponible: se abre {Path(ruta).name} con Excel")
    import xlwings as xw
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        wb = app.books.open(str(ruta), read_only=True, update_links=False)
        sh = next((s for s in wb.sheets if normalizar(s.name) == normalizar(hoja)), None)
        if sh is None:
            raise RuntimeError(f"No está la hoja '{hoja}' en {Path(ruta).name}")
        ult = sh.used_range.last_cell.row
        out = {}
        for c in columnas:
            vals = sh.range(f"{c}1:{c}{ult}").options(ndim=1).value or []
            out[c.upper()] = {i + 1: v for i, v in enumerate(vals) if v is not None}
        wb.close()
        return out
    finally:
        try:
            app.quit()
        except Exception:
            pass


def _col_num(letra):
    n = 0
    for ch in letra.upper():
        n = n * 26 + ord(ch) - 64
    return n


def _col_letra(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def leer_zona(ruta, hoja, col_ini, col_fin, fila_fin, log):
    """{COL: {fila: valor}} de col_ini:col_fin, filas 1..fila_fin (para buscar
    encabezados). Sin Excel si se puede; si no, con xlwings."""
    cols = [_col_letra(i) for i in range(_col_num(col_ini), _col_num(col_fin) + 1)]
    try:
        d = _excel_xml.leer_columnas_rapido(ruta, hoja, cols, 1, log, fila_fin=fila_fin)
    except TypeError:
        # __comun__/excel_xml.py anterior al 2026-10-29 (sin fila_fin): se lee
        # la hoja entera y se queda con las filas de arriba. Mas lento, igual.
        d = _excel_xml.leer_columnas_rapido(ruta, hoja, cols, 1, log)
        if d is not None:
            d = {c: {f: v for f, v in fs.items() if f <= fila_fin} for c, fs in d.items()}
    if d is not None:
        return d
    import xlwings as xw
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        wb = app.books.open(str(ruta), read_only=True, update_links=False)
        sh = next((s for s in wb.sheets if normalizar(s.name) == normalizar(hoja)), None)
        if sh is None:
            raise RuntimeError(f"No está la hoja '{hoja}' en {Path(ruta).name}")
        vals = sh.range(f"{col_ini}1:{col_fin}{fila_fin}").options(ndim=2).value
        wb.close()
        out = {c: {} for c in cols}
        for f, fila in enumerate(vals, start=1):
            for c, v in zip(cols, fila):
                if v is not None:
                    out[c][f] = v
        return out
    finally:
        try:
            app.quit()
        except Exception:
            pass


def ubicar_encabezado(zona, respaldo=None, encabezados=ENCABEZADOS,
                      opcionales=CAMPOS_OPCIONALES, max_dist=MAX_DISTANCIA_ENCABEZADO):
    """Busca la tabla por sus encabezados. Devuelve (fila, {campo: COL}, n_tablas)
    o None. respaldo: {campo: COL} de la tabla esperada, para elegir si hay mas
    de una que calce (la mas cercana a su "pago")."""
    claves = {campo: [_texto.clave(n) for n in nombres]
              for campo, nombres in encabezados.items()}
    filas = {}
    for col, celdas in zona.items():
        for f, v in celdas.items():
            if isinstance(v, str) and v.strip():
                filas.setdefault(f, {})[_col_num(col)] = _texto.clave(v)
    candidatos = []
    for f in sorted(filas):
        celdas = filas[f]
        for ancla, t in celdas.items():
            if t not in claves["pago"]:
                continue
            mapa, usadas = {"pago": ancla}, {ancla}
            for campo, ks in claves.items():
                if campo == "pago":
                    continue
                opciones = [(ks.index(t2), abs(c - ancla), c) for c, t2 in celdas.items()
                            if t2 in ks and c not in usadas and abs(c - ancla) <= max_dist]
                if opciones:
                    c = min(opciones)[2]
                    mapa[campo] = c
                    usadas.add(c)
            if all(c in mapa for c in claves if c not in opcionales):
                candidatos.append((f, mapa))
    if not candidatos:
        return None
    ref = _col_num(respaldo["pago"]) if respaldo and respaldo.get("pago") else None
    if ref is not None:
        candidatos.sort(key=lambda fm: (abs(fm[1]["pago"] - ref), fm[0]))
    f, mapa = candidatos[0]
    return f, {campo: _col_letra(c) for campo, c in mapa.items()}, len(candidatos)


def _dia_de(v, anio_mes):
    """La celda del dia -> date. Una fecha (serie de Excel o datetime) o el
    numero de dia del mes (1-31, con anio_mes)."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 1 <= v <= 31:
        if anio_mes is None:
            return None
        try:
            return datetime(anio_mes[0], anio_mes[1], int(v))
        except ValueError:
            return None
    f = fecha_de(v)
    return datetime(f.year, f.month, f.day) if f else None


def leer_matriz(ruta, fte, anio_mes, log):
    """REA como matriz: una fila del CSV por central (columna) y bloque (fila)
    con monto distinto de 0. Orden: central por central, como la tabla de antes.
    Si B viene solo en la primera fila de cada dia (celdas combinadas), se
    arrastra hacia abajo."""
    mz = fte["matriz"]
    log(f"  {fte['nombre']}: {Path(ruta).name}, hoja {fte['hoja']} (matriz)")
    avisos = []
    fila_c = mz["fila_centrales"]
    zona = leer_zona(ruta, fte["hoja"], mz["col_ini"], mz["col_max"], fila_c, log)
    if zona is None:
        raise RuntimeError(f"No se pudo leer la hoja '{fte['hoja']}' de "
                           f"{Path(ruta).name}")
    centrales = []                      # [(COL, nombre)] hasta la primera vacia
    for n in range(_col_num(mz["col_ini"]), _col_num(mz["col_max"]) + 1):
        v = zona.get(_col_letra(n), {}).get(fila_c)
        if _vacio(v):
            break
        centrales.append((_col_letra(n), str(v).strip()))
    if not centrales:
        raise RuntimeError(f"{fte['nombre']}: no hay centrales en "
                           f"{mz['col_ini']}{fila_c} hacia la derecha.")
    log(f"    {len(centrales)} central(es) en {centrales[0][0]}{fila_c}:"
        f"{centrales[-1][0]}{fila_c}")

    cols = [mz["col_dia"], mz["col_bloque"]] + [c for c, _ in centrales]
    d = leer_columnas(ruta, fte["hoja"], cols, log)
    if d is None:
        raise RuntimeError(f"No se pudo leer la hoja '{fte['hoja']}' de "
                           f"{Path(ruta).name}")
    g = lambda col, f: d.get(col.upper(), {}).get(f)
    filas_n = sorted({f for c in cols for f in d.get(c.upper(), {})
                      if f >= mz["fila_datos"]})

    # dia y bloque de cada fila
    tiempos, dia, malas = {}, None, 0
    for f in filas_n:
        vd = g(mz["col_dia"], f)
        if not _vacio(vd):
            dia = _dia_de(vd, anio_mes)
            if dia is None:
                malas += 1
                log(f"    OJO: fila {f}: día que no se entiende en {mz['col_dia']}: "
                    f"{vd!r} (se omite)")
                continue
        b = _numero(g(mz["col_bloque"], f))
        if b is None:
            continue                     # fila sin bloque (total, vacia, ...)
        if dia is None or not 1 <= b <= BLOQUES_POR_DIA or not float(b).is_integer():
            malas += 1
            log(f"    OJO: fila {f}: bloque {b!r} o día faltante (se omite)")
            continue
        if anio_mes and (dia.year, dia.month) != anio_mes:
            malas += 1
            log(f"    OJO: fila {f}: {dia:%d/%m/%Y} no es del mes {anio_mes[1]:02d}/"
                f"{anio_mes[0]} (se omite)")
            continue
        b = int(b)
        tiempos[f] = (dia + timedelta(minutes=15 * (b - 1)),
                      (dia.day - 1) * BLOQUES_POR_DIA + b)
    log(f"    {len(tiempos)} bloque(s) desde la fila {mz['fila_datos']}")

    salida, ceros = [], 0
    for col, central in centrales:
        for f, (fecha, bloque) in tiempos.items():
            pago = _numero(g(col, f))
            if pago is None or abs(pago) <= EPS_CERO:
                ceros += 1
                continue
            salida.append([fecha.strftime(FORMATO_FECHA), "", central, pago,
                           mz["tipo"], bloque])
    log(f"    {len(salida)} fila(s) con pago; descartadas {ceros} celdas en 0 o vacías")
    if malas:
        avisos.append(f"{fte['nombre']}: {malas} fila(s) omitidas (ver el log)")
    return salida, avisos


def leer_fuente(ruta, fte, anio_mes, log):
    """Las filas de una fuente, ya en el formato del CSV, y los avisos."""
    if fte.get("matriz"):
        return leer_matriz(ruta, fte, anio_mes, log)
    log(f"  {fte['nombre']}: {Path(ruta).name}, hoja {fte['hoja']}")
    avisos = []

    # --- la tabla, por sus encabezados ---
    zona = leer_zona(ruta, fte["hoja"], *ZONA_ENCABEZADO, FILAS_BUSCAR_ENCABEZADO, log)
    if zona is None:
        raise RuntimeError(f"No se pudo leer la hoja '{fte['hoja']}' de "
                           f"{Path(ruta).name}")
    hallado = ubicar_encabezado(zona, respaldo=fte)
    if hallado:
        enc, cols_de, n_tablas = hallado
        f_ini = enc + 1
        fte = dict(fte, fecha=cols_de.get("fecha"), central=cols_de["central"],
                   pago=cols_de["pago"], tipo=cols_de["tipo"], bloque=cols_de["bloque"])
        log(f"    encabezado en la fila {enc}: "
            + " | ".join(f"{campo} {cols_de[campo]} «{zona[cols_de[campo]].get(enc)}»"
                         for campo in ("fecha", "central", "pago", "tipo", "bloque")
                         if campo in cols_de)
            + ("" if "fecha" in cols_de else " | fecha desde el bloque y el mes"))
        if n_tablas > 1:
            log(f"    (calzaron {n_tablas} tablas; se tomó la más cercana a "
                f"{FUENTES_POR_NOMBRE[fte['nombre']]['pago']})")
    else:
        f_ini = fte["fila_datos"]
        av = (f"{fte['nombre']}: no se encontraron los encabezados (Pago, Central, "
              f"Tipo_pago y bloque en {ZONA_ENCABEZADO[0]}1:{ZONA_ENCABEZADO[1]}"
              f"{FILAS_BUSCAR_ENCABEZADO}); se usan las columnas fijas desde la fila "
              f"{f_ini}")
        avisos.append(av)
        log(f"    OJO: {av}")

    cols = [c for c in (fte["fecha"], fte["central"], fte["pago"], fte["tipo"],
                        fte["bloque"]) if c]
    d = leer_columnas(ruta, fte["hoja"], cols, log)
    if d is None:
        raise RuntimeError(f"No se pudo leer la hoja '{fte['hoja']}' de "
                           f"{Path(ruta).name}")
    g = lambda col, f: d.get(col.upper(), {}).get(f)

    filas_n = sorted({f for c in cols for f in d.get(c.upper(), {}) if f >= f_ini})
    salida, ceros, sin_central, malas, no_calza = [], 0, 0, 0, 0
    for f in filas_n:
        cen, pago = g(fte["central"], f), _numero(g(fte["pago"], f))
        if pago is None or abs(pago) <= EPS_CERO:
            ceros += 1
            continue
        if _vacio(cen) or normalizar(str(cen)) in ("0", "0.0"):
            sin_central += 1
            continue
        bloque = _numero(g(fte["bloque"], f))
        if bloque is None:
            malas += 1
            log(f"    OJO: fila {f}: sin bloque en {fte['bloque']} (se omite)")
            continue
        bloque = int(bloque)
        if fte["fecha"]:
            fecha = fecha_de(g(fte["fecha"], f))
            if fecha is None:
                malas += 1
                log(f"    OJO: fila {f}: fecha que no se entiende en "
                    f"{fte['fecha']}: {g(fte['fecha'], f)!r} (se omite)")
                continue
            if bloque_de_fecha(fecha) != bloque:
                no_calza += 1
        else:
            if anio_mes is None:
                raise RuntimeError(f"{fte['nombre']} no trae fecha: hace falta el mes "
                                   "(AAMM) para armarla desde el bloque.")
            fecha = fecha_de_bloque(bloque, *anio_mes)
        if anio_mes and (fecha.year, fecha.month) != anio_mes:
            malas += 1
            log(f"    OJO: fila {f}: {fecha:%d/%m/%Y} no es del mes {anio_mes[1]:02d}/"
                f"{anio_mes[0]} (se omite)")
            continue
        salida.append([fecha.strftime(FORMATO_FECHA), "", str(cen).strip(),
                       pago, str(g(fte["tipo"], f) or "").strip(), bloque])

    log(f"    {len(salida)} fila(s) con pago; descartadas {ceros} con pago 0 o "
        f"vacío y {sin_central} sin central")
    if malas:
        avisos.append(f"{fte['nombre']}: {malas} fila(s) omitidas (ver el log)")
    if no_calza:
        av = (f"{fte['nombre']}: {no_calza} fila(s) con un bloque que no calza "
              f"con la fecha ((día-1)*96 + hora*4 + min/15 + 1). ¿Mes con cambio "
              f"de hora? Se dejan como vienen.")
        avisos.append(av)
        log(f"    OJO: {av}")
    return salida, avisos


# =============================================================================
#  PROCESO
# =============================================================================
def generar(rutas, aamm, salida, log, progreso=lambda *a: None):
    """rutas: {"p3":..., "p5":..., "p6":...}. Escribe `salida` y devuelve
    (filas, resumen_por_tipo, avisos). Ante un problema lanza la excepcion sin
    tocar el archivo anterior."""
    anio_mes = (2000 + int(aamm) // 100, int(aamm) % 100) if aamm else None
    filas, avisos = [], []
    for i, fte in enumerate(FUENTES):
        progreso(i, len(FUENTES), f"Leyendo {fte['nombre']}...")
        f, av = leer_fuente(rutas[fte["planilla"]], fte, anio_mes, log)
        filas += f
        avisos += av

    if not filas:
        raise RuntimeError("Ninguna fuente trajo filas con pago: no se escribe nada.")

    resumen = {}
    for r in filas:
        n, s = resumen.get(r[4], (0, 0.0))
        resumen[r[4]] = (n + 1, s + r[3])

    progreso(len(FUENTES), len(FUENTES) + 1, "Escribiendo...")
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow(ENCABEZADO_CSV)
    for r in filas:
        w.writerow([r[0], r[1], r[2], fmt_pago(r[3]), r[4], r[5]])
    salida = Path(salida)
    tmp = salida.with_suffix(salida.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        fh.write(buf.getvalue())
    try:
        os.replace(tmp, salida)
    except PermissionError:
        try:
            tmp.unlink()
        except OSError:
            pass
        raise RuntimeError(f"No se puede escribir {salida.name}: ¿está abierto "
                           "en Excel? Ciérralo y vuelve a intentar.")
    log(f"  Escrito: {salida}  ({len(filas)} filas)")
    for tipo, (n, s) in resumen.items():
        log(f"    {tipo:<10} {n:>7} fila(s)   {s:>22,.2f}")
    progreso(1, 1, "Listo")
    return filas, resumen, avisos


# =============================================================================
#  VENTANA
# =============================================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    root = tk.Tk()
    root.title("Arma pagos_varios (planillas 3, 5 y 6)"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("980x680")

    if traspaso:
        fa = tk.Frame(root, bg="#fff4c2", bd=1, relief="solid")
        fa.pack(fill="x", padx=20, pady=(10, 0))
        tk.Label(fa, text=f"Mes {traspaso.get('aamm') or '?'} — enviado por el Revisor",
                 bg="#fff4c2", font=("Segoe UI", 11, "bold")).pack(pady=(6, 0))
        tk.Label(fa, text="Las rutas las resolvió el Revisor.", bg="#fff4c2",
                 font=("Segoe UI", 8), fg="#444444").pack(pady=(0, 6))

    tk.Label(root, text="pagos_varios_AAMM.csv  —  Subastas (3), CRA (5), "
                        "CO_ERNC y REA (6)",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(root, text="Solo lee las planillas (no hace falta cerrarlas). "
                        "Fecha DD/MM/AAAA HH:MM, sin pagos en 0.",
             font=("Segoe UI", 8), fg="#555555").pack(pady=(0, 6))

    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_aamm = tk.StringVar(value=(traspaso or {}).get("aamm") or "")
    vars_pl = {k: tk.StringVar(value="[pendiente]") for k in ORDEN_PL}
    var_salida = tk.StringVar(value="[pendiente]")
    labels = {}

    def fila(titulo, var, clave):
        fr = tk.LabelFrame(root, text=titulo, padx=10, pady=4)
        fr.pack(fill="x", padx=20, pady=3)
        l = tk.Label(fr, textvariable=var, wraplength=880, justify="left",
                     cursor="hand2", font=("Segoe UI", 9), anchor="w")
        l.pack(fill="x")
        l.bind("<Button-1>", lambda e: abrir_en_explorador(var.get(), es_archivo=True))
        labels[clave] = l
        return fr

    for k in ORDEN_PL:
        fila(PLANILLAS[k]["etiqueta"], vars_pl[k], k)
    fila("Salida — 04 Planilla 9/pagos_varios_AAMM.csv  (se crea o se reemplaza)",
                var_salida, "salida")

    fm = tk.Frame(root)
    fm.pack(fill="x", padx=24, pady=(4, 0))
    tk.Label(fm, text="Mes (AAMM):", font=("Segoe UI", 9)).pack(side="left")
    tk.Entry(fm, textvariable=var_aamm, width=8,
             font=("Consolas", 10)).pack(side="left", padx=6)
    tk.Label(fm, text="si queda vacío se saca del nombre de las planillas",
             font=("Segoe UI", 8), fg="#555555").pack(side="left")

    fc = tk.LabelFrame(root, text="Carpeta 02 CASO RELIQUIDACION", padx=10, pady=4)
    fc.pack(fill="x", padx=20, pady=3)
    lc = tk.Label(fc, textvariable=var_carpeta, wraplength=880, justify="left",
                  cursor="hand2", font=("Segoe UI", 9), anchor="w")
    lc.pack(fill="x")
    lc.bind("<Button-1>", lambda e: abrir_en_explorador(var_carpeta.get()))

    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")
    fpr = tk.Frame(root)
    fpr.pack(fill="x", padx=20, pady=(6, 0))
    tk.Label(fpr, textvariable=var_estado, font=("Segoe UI", 9),
             anchor="w").pack(side="left", fill="x", expand=True)
    tk.Label(fpr, textvariable=var_tiempo, font=("Consolas", 10, "bold"),
             fg="#2d7a2d").pack(side="right", padx=8)
    # Sin barra de progreso (usuaria, 2026-10-29): queda solo el tiempo, que
    # es lo que dice que esta corriendo.

    frame_btns = tk.Frame(root)
    frame_btns.pack(side="bottom", fill="x", pady=8)
    fl = tk.LabelFrame(root, text="Progreso detallado", padx=6, pady=4)
    fl.pack(fill="both", expand=True, padx=20, pady=4)
    txt = tk.Text(fl, height=12, font=("Consolas", 9), wrap="none")
    sb = tk.Scrollbar(fl, command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt.pack(fill="both", expand=True)

    cola = queue.Queue()
    log = lambda m="": cola.put(("log", str(m)))
    progreso = lambda a, t=100, m=None: cola.put(("prog", (a, t, m)))

    def bombear():
        try:
            for _ in range(200):
                tipo, val = cola.get_nowait()
                if tipo == "log":
                    txt.insert("end", val + "\n")
                    txt.see("end")
                else:
                    a, t, m = val
                    if m:
                        var_estado.set(m)
        except queue.Empty:
            pass
        root.after(120, bombear)

    timer = {"on": False, "t0": 0.0}

    def tick():
        if timer["on"]:
            var_tiempo.set(fmt_tiempo(time.time() - timer["t0"]))
            root.after(500, tick)

    def rutas_actuales():
        return {k: vars_pl[k].get() for k in ORDEN_PL
                if not vars_pl[k].get().startswith("[")}

    def mes_actual():
        a = var_aamm.get().strip()
        if a:
            return a, "escrito en la ventana"
        v, motivo = detectar_aamm(rutas_actuales())
        return (str(v) if v else None), motivo

    def pintar():
        for k in ORDEN_PL:
            v = vars_pl[k].get()
            ok = not v.startswith("[") and Path(v).is_file()
            labels[k].config(fg="blue" if ok else "red")
        v = var_salida.get()
        # Rojo si todavia no existe: se crea al generar.
        labels["salida"].config(fg="blue" if not v.startswith("[") and Path(v).is_file()
                                else "red")

    def poner_salida():
        aamm, _ = mes_actual()
        base = var_carpeta.get()
        p9 = None
        if base and Path(base).is_dir():
            p9 = next((s for s in Path(base).iterdir() if s.is_dir()
                       and normalizar(s.name) == normalizar(CARPETA_P9)), None)
        if p9 is None:
            r = rutas_actuales().get("p3") or next(iter(rutas_actuales().values()), None)
            p9 = Path(r).parent if r else None
        if p9 is None or not aamm:
            var_salida.set("[falta la carpeta 04 Planilla 9 o el mes]")
        else:
            var_salida.set(str(ruta_salida(p9, aamm)))
        pintar()

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            for k in ORDEN_PL:
                vars_pl[k].set(d.get(k) or f"[el Revisor no mandó {k}]")
        else:
            base = var_carpeta.get()
            if base and Path(base).is_dir():
                for k in ORDEN_PL:
                    r = buscar_planilla(base, PLANILLAS[k]["patron"])
                    vars_pl[k].set(str(r) if r else
                                   f"[{PLANILLAS[k]['etiqueta']}: no encontrada]")
        poner_salida()

    def sel_carpeta():
        ini = var_carpeta.get()
        r = filedialog.askdirectory(title="Selecciona 02 CASO RELIQUIDACION",
                                    initialdir=ini if ini and Path(ini).is_dir() else "")
        if r:
            var_carpeta.set(r)
            guardar_config({"carpeta_reliq": r})
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Carpeta elegida a mano: se dejan de usar las rutas del Revisor.")
            refrescar()

    tk.Button(fc, text="Examinar", command=sel_carpeta).pack(anchor="w", pady=(4, 0))
    var_aamm.trace_add("write", lambda *_: poner_salida())

    def lanzar():
        rutas = rutas_actuales()
        faltan = [PLANILLAS[k]["etiqueta"] for k in ORDEN_PL
                  if k not in rutas or not Path(rutas[k]).is_file()]
        if faltan:
            messagebox.showerror("Faltan entradas",
                                 "No se encontró:\n\n" + "\n".join(faltan))
            return
        aamm, motivo = mes_actual()
        if not aamm or not re.fullmatch(r"\d{4}", aamm) or not 1 <= int(aamm) % 100 <= 12:
            messagebox.showerror("Falta el mes",
                                 f"No se pudo saber el mes ({motivo}).\n\n"
                                 "Escríbelo en AAMM.")
            return
        salida = var_salida.get()
        if salida.startswith("["):
            messagebox.showerror("Falta la salida", salida)
            return

        txt.delete("1.0", "end")
        log(f"Mes {aamm} ({motivo})")
        for k in ORDEN_PL:
            log(f"{k}: {rutas[k]}")
        log(f"Salida: {salida}")
        log("-" * 70)
        btn.config(state="disabled", bg="#aaaaaa")
        timer["on"] = True
        timer["t0"] = time.time()
        tick()

        def trabajo():
            try:
                filas, resumen, avisos = generar(rutas, aamm, salida, log, progreso)
                ok, msg = True, f"{len(filas)} filas" + (
                    f" | {len(avisos)} aviso(s)" if avisos else "")
            except Exception as e:
                log(f"\nERROR: {e}")
                log(traceback.format_exc())
                ok, msg = False, str(e)

            def fin():
                timer["on"] = False
                btn.config(state="normal", bg="#2d7a2d")
                pintar()
                if ok:
                    var_estado.set(f"Listo — {msg}")
                    messagebox.showinfo("Listo", f"pagos_varios generado.\n\n{msg}\n\n"
                                        f"{salida}")
                else:
                    var_estado.set("Terminó con errores — revisa el log")
                    messagebox.showerror("Error", f"Falló.\n\n{msg}")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text="GENERAR pagos_varios", bg="#2d7a2d",
                    fg="white", font=("Segoe UI", 10, "bold"), command=lanzar)
    btn.pack(side="left", padx=8, expand=True)
    tk.Button(frame_btns, text="Refrescar rutas",
              command=refrescar).pack(side="left", padx=8)
    tk.Button(frame_btns, text="Salir", command=root.destroy).pack(side="left", padx=8)

    refrescar()
    bombear()
    root.mainloop()


if __name__ == "__main__":
    main()
