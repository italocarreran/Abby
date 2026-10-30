# =============================================================================
#  Actualiza la hoja "SC y CO" de la planilla 5_  —  VERSION 15 MINUTOS
# =============================================================================
#  Pega los SC y los CO de los EMBALSES, con su prorrata por servicio.
#  Metodologia: la del CRA de Balance_SSCC (trazabilidad v5, seccion 6.6),
#  confirmada por la usuaria el 2026-10-23.
#
#  ORIGEN Y DESTINO
#  ----------------------------------------------------------------------------
#  Los dos bloques van en la MISMA hoja, uno debajo del otro: primero SC, despues
#  CO, desde la fila 9.
#
#  SC — Calculo_SobrecostosSSCC, hoja "SOBRECOSTOS", desde la fila 7:
#      S -> C   Clave Año_Mes
#      T -> D   tipo (SCCF)
#      U -> E   central                       <- por aca se filtra
#      A, R -> F   Clave_Bloque = DIA(A) & "#" & (R - (DIA(A)-1)*96)
#      W -> G   monto
#      AW:BN -> I:Z   las 18 prorratas, tal cual (ID, Pmax y Real; cada grupo
#                     en el orden CPF+, CPF-, CSF+, CTF+, CSF-, CTF-)
#      AA:AC  CPF, CSF y CTF: la SUMA de esas 18, calculada aca y pegada como
#             valor (AA = I,J,O,P,U,V · AB = K,M,Q,S,W,Y · AC = L,N,R,T,X,Z)
#
#  CO — Reporte_CRA_15min_AAMM.csv, en 04 Planilla 9, desde la fila 2
#       (ya NO Calculo_CO):
#      A -> C   Clave Año_Mes
#      B -> D   tipo (CO)
#      H -> E   central                       <- por aca se filtra
#      C, E -> F   Clave_Bloque = DIA(C) & "#" & E   (E = bloque 1..96 del dia)
#      I -> G   monto (RCO)
#      I:Z      vacio: el CSV no trae el desglose ID / Pmax / Real
#      AA:AC  CPF = J+K, CSF = L+M, CTF = N+O del CSV
#
#  OJO con el 96 de la Clave_Bloque de SC: es la formula del Excel y se copia
#  tal cual. En un mes con cambio de hora (dia de 92 o 100 bloques) los dias
#  siguientes quedan corridos; se avisa en el log, no se corrige.
#
#  Las columnas de formula (BLOQUES_FORMULA) se estiran o se cortan para cubrir
#  exactamente las filas que quedaron. Las que estan entre medio (AE, AH) y la
#  tabla dict_SCCO de AT:AV no se tocan.
#
#  POR QUE SE REESCRIBEN LOS DOS BLOQUES SIEMPRE
#  ----------------------------------------------------------------------------
#  El bloque CO va DEBAJO del de SC, asi que no son independientes: si los SC de
#  este mes traen mas o menos embalses que los del mes pasado, el bloque CO tiene
#  que correrse. Por eso, aunque se elija actualizar solo uno, se reescriben los
#  dos: el que no se eligio se lee del propio destino y se vuelve a escribir en su
#  lugar nuevo. Asi nunca queda un hueco ni una fila pisada.
#
#  Los bloques se distinguen por la columna D: "SCCF" son SC y "CO" son CO.
#
#  Se puede correr solo, o recibir del Revisor la ruta de un JSON de traspaso.
# =============================================================================

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime
import json, sys, re, socket, os, traceback, unicodedata, time


DIR_SCRIPT = Path(__file__).resolve().parent

# config.json es compartido con el Revisor y el resto de los actualizadores,
# y ahora vive en __config__, junto a Revisor_Relq. No es
# DIR_SCRIPT / "config.json" porque este script esta en actualizadores/.
CONFIG_PATH = DIR_SCRIPT.parent.parent / "__config__" / "config.json"


def _morir(titulo, mensaje):
    """Aborta mostrando el motivo en una ventana. Sin esto, lanzado desde el
    Revisor con pythonw (sin consola) moriria callado."""
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(titulo, mensaje)
        r.destroy()
    except Exception:
        pass
    print(f"{titulo}\n\n{mensaje}", file=sys.stderr)
    raise SystemExit(1)


# ---------------------------------------------------------------------------
#  Centrales de embalse
# ---------------------------------------------------------------------------
#  OJO: esta lista esta TAMBIEN en Revisor_Reliquidacion_15min.py. Si se agrega o
#  quita una central, hay que cambiarla EN LOS DOS.
#
#  Que pasa si se desincronizan (las dos las caza V10, por suerte):
#    - si aca hay una central que en el otro no: se pega igual, y el revisor la
#      marca en "en E9:E solo hay centrales de embalse".
#    - si en el otro hay una que aca no: no se pega, y el revisor la marca en
#      "no quedo fuera ninguna central «-numero»".
#  O sea que la lista vieja se nota al verificar, no pasa en silencio.
#
#  Al agregar una unidad nueva va con el nombre EXACTO como viene en el origen
#  (columna U de Calculo_SobrecostosSSCC / columna H del Reporte_CRA).
CENTRALES_EMBALSE = [
    "CANUTILLAR-1", "CANUTILLAR-2",
    "ELTORO-1", "ELTORO-2", "ELTORO-3", "ELTORO-4",
    "RALCO-1", "RALCO-2",
    "RAPEL-1", "RAPEL-2", "RAPEL-3", "RAPEL-4", "RAPEL-5",
    "PEHUENCHE-1", "PEHUENCHE-2",
    "COLBUN-1", "COLBUN-2",
    "CIPRESES-1", "CIPRESES-2", "CIPRESES-3",
    "PANGUE-1", "PANGUE-2",
    "ANTUCO-1", "ANTUCO-2",
    "ANGOSTURA-1", "ANGOSTURA-2", "ANGOSTURA-3",
]

# ---------------------------------------------------------------------------
#  Configuracion de origenes y destino
# ---------------------------------------------------------------------------
HOJA_DESTINO = "SC y CO"
FILA_DESTINO = 9

# Tres bloques de VALORES: identificacion (C:G), las 18 prorratas (I:Z) y las
# tres sumas por servicio (AA:AC). La H queda afuera: no se toca.
COLS_ID = ("C", "G")
COLS_PRO = ("I", "Z")
COLS_SUMA = ("AA", "AC")
# Bloques de formulas, EN BLOQUES y no un rango corrido: AE y AH no llevan
# formula, y a la derecha (AT:AV) esta la tabla dict_SCCO, que es un maestro.
# Un AutoFill corrido de AD:AI rellenaria AE y AH con lo que tengan en la 9.
# Es la misma particion que usa V10 del Revisor 15 min.
#   AD = 1 - SUM(AA:AC)                        residuo
#   AF = AG & "_" & F, AG = INDEX(dict_SCCO)   homologacion al nombre CRA
#   AI = control de duplicados
BLOQUES_FORMULA = [("AD", "AD"), ("AF", "AG"), ("AI", "AI")]
# Hasta donde se mira la fila 9 para avisar si hay formulas fuera de la
# particion (la tabla dict_SCCO empieza en AT).
COL_FIN_CONTROL_FORMULAS = "AS"

COL_TIPO = "D"          # SCCF o CO
COL_CLAVE = "C"         # Clave Año_Mes
COL_CENTRAL = "E"
N_ID = 5                # C:G
N_PRO = 18              # I:Z
N_SUMA = 3              # AA:AC

TIPO_SC = "SCCF"
TIPO_CO = "CO"

# Bloques de 15 minutos por dia en la formula de Clave_Bloque de SC (la del
# Excel). Ver el OJO del encabezado sobre el cambio de hora.
BLOQUES_POR_DIA = 96

# Posiciones dentro de I:Z (0 = I) que suma cada servicio, igual que las
# formulas AA:AC del libro. Cada grupo de 6 (ID, Pmax, Real) viene en el orden
# CPF+, CPF-, CSF+, CTF+, CSF-, CTF-: por eso el CSF no son posiciones seguidas.
SUMAS_SC = {
    "CPF": (0, 1, 6, 7, 12, 13),       # I, J, O, P, U, V
    "CSF": (2, 4, 8, 10, 14, 16),      # K, M, Q, S, W, Y
    "CTF": (3, 5, 9, 11, 15, 17),      # L, N, R, T, X, Z
}

FUENTES = {
    "SC": {
        "etiqueta": "SC  —  Calculo_SobrecostosSSCC, hoja SOBRECOSTOS",
        "clave_traspaso": "calculo_sscc_maestro",
        "hoja": "SOBRECOSTOS",
        "fila_ini": 7,
        "col_clave": "S", "col_tipo": "T", "col_central": "U",
        "col_fecha": "A", "col_periodo": "R", "col_monto": "W",
        "bloque_pro": ("AW", "BN"),
    },
    "CO": {
        "etiqueta": "CO  —  Reporte_CRA_15min_AAMM.csv",
        "clave_traspaso": "reporte_cra",
        "fila_ini": 2,
        "col_clave": "A", "col_tipo": "B", "col_central": "H",
        "col_fecha": "C", "col_bloque": "E", "col_monto": "I",
        # CPF = J+K, CSF = L+M, CTF = N+O
        "sumas": {"CPF": ("J", "K"), "CSF": ("L", "M"), "CTF": ("N", "O")},
    },
}
ORDEN = ["SC", "CO"]        # el orden en que quedan en la hoja: SC arriba

# Una central "-numero" que no este en la lista es sospechosa: el sufijo indica
# unidad, y las unidades son justamente lo que tienen los embalses.
RE_UNIDAD = re.compile(r"-\d+\s*$")


# =============================================================================
#  CONFIG COMPARTIDO
# =============================================================================
# El manejo del config.json vive en __comun__/config.py, en la raiz comun.
# Estaba copiado en los 10 scripts y las copias se habian ido separando entre
# si. Los nombres de siempre se conservan como envoltorios, asi que ningun
# punto de llamada cambia. Ver MAPA.md, "El modulo comun".
try:
    sys.path.insert(0, str(DIR_SCRIPT.parent.parent))
    from __comun__ import config as _cfg
    from __comun__ import traspaso as _traspaso
    from __comun__ import texto as _texto
    from __comun__ import csv_columnas as _csv
except ImportError as e:
    _morir("Falta la carpeta __comun__/",
           "No se pudo cargar __comun__/config.py.\n\n"
           "Tiene que estar la carpeta '__comun__' hermana de Revisor_Relq,\n"
           "con config.py adentro. Baja el repositorio completo, no los .py sueltos.\n\n"
           f"Carpeta actual: {DIR_SCRIPT}\n\nDetalle: {e}")

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


clave_central = _texto.clave_mayusculas


def quitar_filtros(sh, log=print):
    """Muestra todas las filas de una hoja que tiene un filtro aplicado.

    Hay que hacerlo ANTES de pegar: con un filtro puesto en el destino, la
    escritura de un bloque de valores no cae en las filas que corresponden
    (en Cálculo_Sobrecostos quedó una fila con la fecha en todas las columnas;
    sin el filtro salió bien, usuaria 2026-10-29). Se limpia solo el criterio:
    las flechas del filtro quedan. Cubre el filtro de la hoja y el de las
    tablas (ListObjects). Devuelve cuantos filtros se limpiaron.
    """
    n = 0
    try:
        api = sh.api
    except Exception:
        return 0
    try:
        if api.FilterMode:
            api.ShowAllData()
            n += 1
    except Exception as e:
        log(f"     (no se pudo quitar el filtro de {sh.name}: {e})")
    try:
        for lo in api.ListObjects:
            af = lo.AutoFilter
            if af is not None and af.FilterMode:
                af.ShowAllData()
                n += 1
    except Exception:
        pass
    if n:
        log(f"     filtro quitado en '{sh.name}' antes de pegar (las flechas quedan)")
    return n


def buscar_hoja(wb, nombre):
    objetivo = normalizar(nombre)
    for sh in wb.sheets:
        if normalizar(sh.name) == objetivo:
            return sh
    return None


def col_num(letra):
    n = 0
    for c in str(letra).upper():
        n = n * 26 + (ord(c) - 64)
    return n


def col_letra(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ultima_fila(sh, col, desde=1):
    try:
        f = sh.api.Cells(sh.api.Rows.Count, col_num(col)).End(-4162).Row   # xlUp
        return max(int(f), desde - 1)
    except Exception:
        return desde - 1


def _vacio(v):
    return v is None or (isinstance(v, str) and not v.strip())


# =============================================================================
#  LECTURA DEL ORIGEN
# =============================================================================
def _numero(v):
    """Valor de celda -> float, o None si no es un numero (como SUMA de Excel,
    que ignora el texto)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def _entero_si_se_puede(x):
    if isinstance(x, float) and x.is_integer():
        return int(x)
    return x


def dia_de(v):
    """DIA() de Excel para lo que venga en la celda de fecha: un datetime (xlwings),
    un numero de serie, o un texto "dd-mm-aaaa" / "dd/mm/aaaa" / "aaaa-mm-dd"
    (el CSV). None si no se entiende."""
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.day
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(v))).day
        except Exception:
            return None
    if isinstance(v, str):
        partes = re.split(r"[-/ ]", v.strip())
        try:
            if len(partes[0]) == 4:          # aaaa-mm-dd
                return int(partes[2])
            return int(partes[0])            # dd-mm-aaaa
        except (ValueError, IndexError):
            return None
    return None


def clave_bloque(dia, bloque):
    """DIA & "#" & bloque, como lo arma el Excel ("1#38")."""
    if dia is None or bloque is None:
        return None
    return f"{dia}#{_entero_si_se_puede(bloque)}"


def _filtrar_embalses(filas, log):
    """filas: lista de (central, fila_id, fila_pro, fila_suma). Devuelve
    (las de embalse, avisos) y deja en el log lo que se descarto."""
    permitidas = {clave_central(c) for c in CENTRALES_EMBALSE}
    quedan, vistas, descartadas, sospechosas = [], set(), {}, {}
    for nom, fid, fpro, fsum in filas:
        if _vacio(nom):
            continue
        k = clave_central(nom)
        if k not in permitidas:
            descartadas[k] = str(nom).strip()
            # Termina en "-numero" pero no esta en la lista: puede ser una unidad
            # de embalse nueva que nadie agrego todavia.
            if RE_UNIDAD.search(str(nom).strip()):
                sospechosas[k] = str(nom).strip()
            continue
        vistas.add(k)
        quedan.append((fid, fpro, fsum))

    log(f"    embalses: {len(quedan)} fila(s) de {len(vistas)} central(es)")
    log(f"    descartadas: {len(descartadas)} central(es) que no son embalse")

    avisos = []
    faltan = [c for c in CENTRALES_EMBALSE if clave_central(c) not in vistas]
    if faltan:
        avisos.append("En este origen no apareció ninguna fila de: "
                      + ", ".join(faltan))
        log(f"    OJO: sin filas para {len(faltan)} central(es) de la lista:")
        for c in faltan[:12]:
            log(f"          {c}")
    if sospechosas:
        avisos.append("Centrales que terminan en «-número» y NO están en la lista "
                      "de embalses: " + ", ".join(sorted(sospechosas.values())))
        log(f"    OJO: {len(sospechosas)} central(es) terminan en «-número» y no")
        log(f"         están en la lista. Puede ser una unidad nueva:")
        for c in sorted(sospechosas.values())[:12]:
            log(f"          {c}")
        if len(sospechosas) > 12:
            log(f"          ... y {len(sospechosas) - 12} más")
    return quedan, avisos


def leer_sc(app, ruta, cfg, log):
    """Devuelve (filas_id, filas_pro, filas_suma, avisos) del Calculo_Sobrecostos.

    filas_id:   5 valores (C:G).
    filas_pro:  las 18 prorratas (I:Z), tal cual.
    filas_suma: CPF, CSF, CTF (AA:AC), sumadas desde las 18.
    """
    log(f"  Abriendo (solo lectura): {Path(ruta).name}")
    wb = app.books.open(str(ruta), read_only=True, update_links=False)
    try:
        sh = buscar_hoja(wb, cfg["hoja"])
        if sh is None:
            raise RuntimeError(f"No está la hoja '{cfg['hoja']}' en "
                               f"{Path(ruta).name}")
        log(f"    hoja: {sh.name}")
        f_ini = cfg["fila_ini"]
        col_c = cfg["col_central"]

        # La ultima fila se busca en la columna de la central y en la de la
        # clave: si una tiene huecos, la otra salva.
        f_fin = max(ultima_fila(sh, col_c, f_ini),
                    ultima_fila(sh, cfg["col_clave"], f_ini))
        if f_fin < f_ini:
            raise RuntimeError(f"No hay datos en {col_c}{f_ini} hacia abajo.")
        log(f"    filas leídas: {f_ini} a {f_fin}  ({f_fin - f_ini + 1})")

        def bloque(c1, c2):
            d = sh.range((f_ini, col_num(c1)), (f_fin, col_num(c2))).options(ndim=2).value
            return d or []

        def col(c):
            return [f[0] if f else None for f in bloque(c, c)]

        claves, tipos, centrales = col(cfg["col_clave"]), col(cfg["col_tipo"]), col(col_c)
        fechas, periodos, montos = (col(cfg["col_fecha"]), col(cfg["col_periodo"]),
                                    col(cfg["col_monto"]))
        pros = bloque(*cfg["bloque_pro"])
    finally:
        try:
            wb.close()
        except Exception:
            pass

    n = f_fin - f_ini + 1
    filas, fuera_de_dia = [], 0
    for i in range(n):
        dia = dia_de(fechas[i] if i < len(fechas) else None)
        per = _numero(periodos[i] if i < len(periodos) else None)
        bloque_dia = None if dia is None or per is None else per - (dia - 1) * BLOQUES_POR_DIA
        if bloque_dia is not None and not (1 <= bloque_dia <= BLOQUES_POR_DIA):
            fuera_de_dia += 1
        fid = [claves[i] if i < len(claves) else None,
               tipos[i] if i < len(tipos) else None,
               centrales[i] if i < len(centrales) else None,
               clave_bloque(dia, bloque_dia),
               montos[i] if i < len(montos) else None]
        fpro = (list(pros[i]) if i < len(pros) and pros[i] else []) + [None] * N_PRO
        fpro = fpro[:N_PRO]
        fsum = [sum(_numero(fpro[j]) or 0.0 for j in SUMAS_SC[s])
                for s in ("CPF", "CSF", "CTF")]
        filas.append((fid[2], fid, fpro, fsum))

    quedan, avisos = _filtrar_embalses(filas, log)
    sin_clave = sum(1 for fid, _, _ in quedan if fid[3] is None)
    if sin_clave:
        avisos.append(f"{sin_clave} fila(s) SC sin fecha (A) o período (R): "
                      "quedan sin Clave_Bloque.")
        log(f"    OJO: {sin_clave} fila(s) sin fecha en A o período en R: "
            "quedan sin Clave_Bloque")
    if fuera_de_dia:
        avisos.append(f"{fuera_de_dia} fila(s) SC con bloque fuera de 1..96: "
                      "¿mes con cambio de hora?")
        log(f"    OJO: {fuera_de_dia} fila(s) dan un bloque fuera de 1..{BLOQUES_POR_DIA}"
            f" con R - (DIA(A)-1)*{BLOQUES_POR_DIA}. Si el mes tiene cambio de")
        log("         hora, la Clave_Bloque de los días siguientes queda corrida.")
    return ([f for f, _, _ in quedan], [p for _, p, _ in quedan],
            [s for _, _, s in quedan], avisos)


def leer_co(ruta, cfg, log):
    """Devuelve (filas_id, filas_pro, filas_suma, avisos) del Reporte_CRA (csv).
    filas_pro va vacia: el CSV no trae el desglose ID / Pmax / Real."""
    log(f"  Leyendo: {Path(ruta).name}")
    cols = [cfg["col_clave"], cfg["col_tipo"], cfg["col_central"],
            cfg["col_fecha"], cfg["col_bloque"], cfg["col_monto"]]
    for a, b in cfg["sumas"].values():
        cols += [a, b]
    d = _csv.leer_columnas_csv(ruta, cols, cfg["fila_ini"], log)
    if d is None:
        raise RuntimeError(f"No se pudo leer {Path(ruta).name}")
    filas_n = sorted(set().union(*(set(v) for v in d.values())) if d else set())
    if not filas_n:
        raise RuntimeError(f"No hay datos en {Path(ruta).name} desde la fila "
                           f"{cfg['fila_ini']}.")
    log(f"    filas leídas: {filas_n[0]} a {filas_n[-1]}  ({len(filas_n)})")

    g = lambda c, f: d.get(c, {}).get(f)
    filas = []
    for f in filas_n:
        fid = [g(cfg["col_clave"], f), g(cfg["col_tipo"], f), g(cfg["col_central"], f),
               clave_bloque(dia_de(g(cfg["col_fecha"], f)), g(cfg["col_bloque"], f)),
               g(cfg["col_monto"], f)]
        fsum = [(_numero(g(a, f)) or 0.0) + (_numero(g(b, f)) or 0.0)
                for a, b in (cfg["sumas"][s] for s in ("CPF", "CSF", "CTF"))]
        filas.append((fid[2], fid, [None] * N_PRO, fsum))

    quedan, avisos = _filtrar_embalses(filas, log)
    sin_clave = sum(1 for fid, _, _ in quedan if fid[3] is None)
    if sin_clave:
        avisos.append(f"{sin_clave} fila(s) CO sin fecha (C) o bloque (E): "
                      "quedan sin Clave_Bloque.")
        log(f"    OJO: {sin_clave} fila(s) sin fecha en C o bloque en E: "
            "quedan sin Clave_Bloque")
    return ([f for f, _, _ in quedan], [p for _, p, _ in quedan],
            [s for _, _, s in quedan], avisos)


def leer_bloque_destino(sh, tipo, log):
    """Lee del propio destino las filas de un tipo (SCCF o CO).

    Sirve para conservar el bloque que el usuario NO eligio actualizar: como los
    dos bloques van uno debajo del otro, hay que reescribirlos juntos."""
    c_id1, c_id2 = (col_num(x) for x in COLS_ID)
    c_p1, c_p2 = (col_num(x) for x in COLS_PRO)
    c_s1, c_s2 = (col_num(x) for x in COLS_SUMA)
    ult = max(ultima_fila(sh, COL_TIPO, FILA_DESTINO),
              ultima_fila(sh, COL_CENTRAL, FILA_DESTINO))
    if ult < FILA_DESTINO:
        log(f"    el destino no tenía filas de {tipo}")
        return [], [], []
    ids = sh.range((FILA_DESTINO, c_id1), (ult, c_id2)).options(ndim=2).value or []
    pros = sh.range((FILA_DESTINO, c_p1), (ult, c_p2)).options(ndim=2).value or []
    sums = sh.range((FILA_DESTINO, c_s1), (ult, c_s2)).options(ndim=2).value or []
    objetivo = normalizar(tipo)
    idx_tipo = col_num(COL_TIPO) - c_id1
    fid, fpro, fsum = [], [], []
    for i, fila in enumerate(ids):
        if not fila or normalizar(fila[idx_tipo] if idx_tipo < len(fila) else None) != objetivo:
            continue
        fid.append((list(fila) + [None] * N_ID)[:N_ID])
        p = pros[i] if i < len(pros) else []
        fpro.append((list(p) + [None] * N_PRO)[:N_PRO])
        q = sums[i] if i < len(sums) else []
        fsum.append((list(q) + [None] * N_SUMA)[:N_SUMA])
    log(f"    conservadas {len(fid)} fila(s) de {tipo} que ya estaban en el destino")
    return fid, fpro, fsum


def controlar_formulas(sh, log):
    """Mira la fila 9 entre AD y COL_FIN_CONTROL_FORMULAS y avisa si la
    particion BLOQUES_FORMULA no calza con el libro: una columna de la particion
    sin formula (se estiraria vacia) o una formula fuera de la particion (no se
    estiraria). No cambia nada: solo avisa."""
    c1, c2 = col_num(BLOQUES_FORMULA[0][0]), col_num(COL_FIN_CONTROL_FORMULAS)
    try:
        fila = sh.range((FILA_DESTINO, c1), (FILA_DESTINO, c2)).formula
    except Exception:
        return []
    if fila and isinstance(fila[0], (list, tuple)):
        fila = fila[0]
    en_particion = set()
    for a, b in BLOQUES_FORMULA:
        en_particion |= set(range(col_num(a), col_num(b) + 1))
    sin_formula, fuera = [], []
    for i, f in enumerate(fila or []):
        c = c1 + i
        es_formula = isinstance(f, str) and f.startswith("=")
        if c in en_particion and not es_formula:
            sin_formula.append(col_letra(c))
        elif c not in en_particion and es_formula:
            fuera.append(col_letra(c))
    avisos = []
    if sin_formula:
        avisos.append(f"En la fila {FILA_DESTINO} no hay fórmula en "
                      f"{', '.join(sin_formula)}, que se esperaba que la tuviera.")
        log(f"    OJO: sin fórmula en la fila {FILA_DESTINO}: {', '.join(sin_formula)}")
    if fuera:
        avisos.append(f"Hay fórmulas en {', '.join(fuera)} que el script no "
                      "estira (no están en BLOQUES_FORMULA).")
        log(f"    OJO: fórmulas que el script NO estira: {', '.join(fuera)}")
    return avisos


# =============================================================================
#  PROCESO
# =============================================================================
def actualizar_en_libro(app, wb_d, rutas, hacer, aamm_respaldo, log, progreso):
    """El trabajo de verdad, sobre un libro YA ABIERTO: lee los origenes, limpia
    y escribe la hoja "SC y CO" y estira las formulas. No cambia el modo de
    calculo, no recalcula ni guarda: eso es de quien abrio el libro.

    Lo usan ejecutar() (la ventana de este script) y Actualiza_datos.py, que en
    15 minutos tiene SC y CO como casillas de la planilla 5 junto con FD y
    prorrata, sobre el mismo libro abierto.

    rutas: {"SC": Calculo_SobrecostosSSCC, "CO": Reporte_CRA (.csv)}; solo hacen
           falta las de lo que se pide en `hacer`.
    Devuelve (hoja, n_sc, n_co, avisos). Ante un problema lanza la excepcion.
    """
    avisos_todos = []
    sh_d = buscar_hoja(wb_d, HOJA_DESTINO)
    if sh_d is None:
        raise RuntimeError(f"No está la hoja '{HOJA_DESTINO}' en el destino.")
    log(f"    hoja destino: {sh_d.name}")
    quitar_filtros(sh_d, log)

    # ---- juntar los dos bloques, se actualicen o no --------------------
    bloques = {}
    for cual in ORDEN:
        log("=" * 70)
        if cual in hacer:
            log(f"{cual}: leyendo del origen")
            if cual == "SC":
                fid, fpro, fsum, av = leer_sc(app, rutas[cual], FUENTES[cual], log)
            else:
                fid, fpro, fsum, av = leer_co(rutas[cual], FUENTES[cual], log)
            avisos_todos += [f"{cual}: {a}" for a in av]
        else:
            log(f"{cual}: NO se actualiza, se conserva lo que ya está")
            tipo = TIPO_SC if cual == "SC" else TIPO_CO
            fid, fpro, fsum = leer_bloque_destino(sh_d, tipo, log)
        bloques[cual] = (fid, fpro, fsum)

    # ---- la Clave Año_Mes ----------------------------------------------
    # En 15 minutos el Reporte_CRA ya la trae. Si a algun CO le faltara, se
    # completa como antes: con la de los SC, o la que tenian los CO, o el
    # mes de la ventana.
    claves = [f[0] for f in bloques["SC"][0] if not _vacio(f[0])]
    clave = None
    if claves:
        distintas = {str(c).strip() for c in claves}
        clave = claves[0]
        if len(distintas) > 1:
            log(f"  OJO: los SC traen más de una Clave Año_Mes: "
                f"{sorted(distintas)}. Se usa {clave!r} para los CO.")
    else:
        # Sin SC no hay de dónde copiarla: se usa la que ya tenían los CO, y
        # si tampoco hay, el mes de la ventana.
        previas = [f[0] for f in bloques["CO"][0] if not _vacio(f[0])]
        clave = previas[0] if previas else (aamm_respaldo or None)
        log(f"  Sin filas SC: la Clave Año_Mes de los CO queda en {clave!r}")
    if clave is None:
        raise RuntimeError(
            "No se pudo determinar la Clave Año_Mes para los CO: no hay filas "
            "SC, el destino no tenía CO y no se indicó el mes.")
    n_completadas = 0
    for f in bloques["CO"][0]:
        if _vacio(f[0]):
            f[0] = clave
            n_completadas += 1
    if n_completadas:
        log(f"  Clave Año_Mes {clave!r} puesta en {n_completadas} fila(s) de CO")

    filas_id = bloques["SC"][0] + bloques["CO"][0]
    filas_pro = bloques["SC"][1] + bloques["CO"][1]
    filas_suma = bloques["SC"][2] + bloques["CO"][2]
    if not filas_id:
        raise RuntimeError("No quedó ninguna fila para escribir.")

    # Un solo mes por hoja: dos Clave Año_Mes distintas es un archivo
    # equivocado (el SC de un mes con el CO de otro).
    meses = sorted({str(_entero_si_se_puede(f[0])).strip()
                    for f in filas_id if not _vacio(f[0])})
    if len(meses) > 1:
        av = (f"Hay más de una Clave Año_Mes en la hoja: {', '.join(meses)}. "
              "¿El SC y el CO son del mismo mes?")
        avisos_todos.append(av)
        log(f"  OJO: {av}")

    # Control AI del libro: la misma Tipo + Unidad + Clave_Bloque dos veces.
    vistas, repetidas = set(), 0
    for f in filas_id:
        k = (normalizar(f[1]), clave_central(f[2]), f[3])
        if f[3] is not None and k in vistas:
            repetidas += 1
        vistas.add(k)
    if repetidas:
        av = f"{repetidas} fila(s) con Tipo + Unidad + Clave_Bloque repetidos."
        avisos_todos.append(av)
        log(f"  OJO: {av}")

    # ---- limpiar y escribir -------------------------------------------
    progreso(50, 100, "Escribiendo...")
    log("=" * 70)
    log("Escribiendo en el destino")
    c_id1, c_id2 = (col_num(x) for x in COLS_ID)
    c_p1, c_p2 = (col_num(x) for x in COLS_PRO)
    c_s1, c_s2 = (col_num(x) for x in COLS_SUMA)
    bloques_f = [(col_num(a), col_num(b)) for a, b in BLOQUES_FORMULA]
    avisos_todos += controlar_formulas(sh_d, log)

    # Se limpia hasta la ultima fila con algo en cualquiera de los tres
    # bloques: el mes anterior pudo ser mas largo, y lo que sobre abajo
    # ensucia los totales.
    previa = 0
    cols_mirar = [c_id1, c_id2, c_p1, c_p2, c_s1, c_s2] + [a for a, _ in bloques_f]
    for c in cols_mirar:
        previa = max(previa, ultima_fila(sh_d, col_letra(c), FILA_DESTINO))
    n = len(filas_id)
    ultima = FILA_DESTINO + n - 1
    if previa >= FILA_DESTINO:
        log(f"    limpiando filas {FILA_DESTINO} a {previa}  "
            f"({previa - FILA_DESTINO + 1} del mes anterior)")
        sh_d.range((FILA_DESTINO, c_id1), (previa, c_id2)).clear_contents()
        sh_d.range((FILA_DESTINO, c_p1), (previa, c_p2)).clear_contents()
        sh_d.range((FILA_DESTINO, c_s1), (previa, c_s2)).clear_contents()
        if previa > ultima:
            # Las formulas que sobran se borran; las que quedan se estiran
            # mas abajo con AutoFill. Se borra bloque por bloque para no
            # tocar AE ni AH.
            for (a, b), (la, lb) in zip(bloques_f, BLOQUES_FORMULA):
                sh_d.range((ultima + 1, a), (previa, b)).clear_contents()
            log(f"    {' y '.join(f'{a}:{b}' for a, b in BLOQUES_FORMULA)}: "
                f"cortado, sobraban {previa - ultima} fila(s)")
    else:
        log("    el destino estaba vacío")

    sh_d.range((FILA_DESTINO, c_id1)).value = filas_id
    sh_d.range((FILA_DESTINO, c_p1)).value = filas_pro
    sh_d.range((FILA_DESTINO, c_s1)).value = filas_suma
    log(f"    {COLS_ID[0]}{FILA_DESTINO}:{COLS_ID[1]}{ultima},  "
        f"{COLS_PRO[0]}{FILA_DESTINO}:{COLS_PRO[1]}{ultima}  y  "
        f"{COLS_SUMA[0]}{FILA_DESTINO}:{COLS_SUMA[1]}{ultima}   "
        f"({len(bloques['SC'][0])} SC + {len(bloques['CO'][0])} CO = {n} filas)")

    # ---- estirar las formulas -----------------------------------------
    if n > 1:
        for a, b in bloques_f:
            origen = sh_d.range((FILA_DESTINO, a), (FILA_DESTINO, b))
            destino = sh_d.range((FILA_DESTINO, a), (ultima, b))
            origen.api.AutoFill(destino.api, 0)     # xlFillDefault
    log(f"    {' y '.join(f'{a}:{b}' for a, b in BLOQUES_FORMULA)} "
        f"estirado hasta la fila {ultima}   (AE y AH no se tocan)")
    return sh_d, len(bloques["SC"][0]), len(bloques["CO"][0]), avisos_todos


def ejecutar(rutas, hacer, aamm_respaldo, log, progreso):
    """hacer: subconjunto de ["SC", "CO"]. Devuelve (ok, resumen)."""
    try:
        import pythoncom
        import xlwings as xw
    except ImportError as e:
        log(f"ERROR: falta una librería: {e}")
        return False, str(e)

    pythoncom.CoInitialize()
    app = None
    wb_d = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        log(f"  Abriendo destino: {Path(rutas['destino']).name}")
        wb_d = app.books.open(str(rutas["destino"]), update_links=False)
        wb_d.app.calculation = "manual"
        sh_d, n_sc, n_co, avisos_todos = actualizar_en_libro(
            app, wb_d, rutas, hacer, aamm_respaldo, log, progreso)
        n = n_sc + n_co

        progreso(85, 100, "Recalculando...")
        wb_d.app.calculation = "automatic"
        t0 = time.time()
        app.calculate()
        log(f"    recálculo: {time.time() - t0:.1f} s")

        log("  Guardando...")
        wb_d.save()
        log("  Guardado.")

        # Queda visible y abierto, igual que los otros actualizadores.
        app.api.Visible = True
        app.screen_updating = True
        app.display_alerts = True
        sh_d.activate()
        wb_d = None
        app = None

        progreso(100, 100, "Listo")
        resumen = f"{n_sc} SC + {n_co} CO = {n} filas"
        if avisos_todos:
            resumen += f" | {len(avisos_todos)} aviso(s)"
        return True, resumen

    except Exception as e:
        log(f"\nERROR: {e}")
        log(traceback.format_exc())
        try:
            if app is not None:
                app.api.Visible = True
                app.screen_updating = True
                app.display_alerts = True
        except Exception:
            pass
        return False, str(e)
    finally:
        pythoncom.CoUninitialize()


# =============================================================================
#  VENTANA
# =============================================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    v = VentanaChica('Hoja "SC y CO" — planilla 5',
                     "Solo embalses. SC arriba y CO abajo; el que no se marca se "
                     "conserva.", traspaso, alto=380)
    var_dest = tk.StringVar(value="[pendiente]")
    var_sc = tk.StringVar(value="[pendiente]")
    var_co = tk.StringVar(value="[pendiente]")
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_aamm = tk.StringVar(value=(traspaso or {}).get("aamm")
                            or cfg.get("ultimo_mes", ""))
    sel_sc = tk.BooleanVar(value=not modo["traspaso"])
    sel_co = tk.BooleanVar(value=not modo["traspaso"])

    fr = v.seccion("Qué actualizar")
    v.casilla(fr, "SC", sel_sc, "desde Cálculo_SobrecostosSSCC")
    v.archivo(fr, "", var_sc, sangria=24)
    v.casilla(fr, "CO", sel_co, "desde Reporte_CRA_15min (.csv)")
    v.archivo(fr, "", var_co, sangria=24)
    fd = v.seccion("Destino")
    v.archivo(fd, "Planilla 5", var_dest)

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            var_dest.set(d.get("p5") or "[el Revisor no la mandó]")
            var_sc.set(d.get("calculo_sscc_maestro") or "[el Revisor no lo mandó]")
            var_co.set(d.get("reporte_cra") or "[el Revisor no lo mandó]")
            return
        base = var_carpeta.get()
        if not base or not Path(base).is_dir():
            for x in (var_dest, var_sc, var_co):
                x.set("[elige la carpeta 02 CASO RELIQUIDACION]")
            return
        base = Path(base)

        def carp(nombre):
            return next((p for p in base.iterdir() if p.is_dir()
                         and normalizar(p.name) == normalizar(nombre)), None)

        def arch(carpeta, pref, ext=(".xlsm", ".xlsx")):
            if carpeta is None:
                return None
            cand = [p for p in carpeta.glob("*")
                    if p.is_file() and normalizar(p.name).startswith(pref)
                    and p.suffix.lower() in ext and not p.name.startswith("~$")]
            return max(cand, key=lambda p: p.stat().st_mtime) if cand else None

        p9, sob = carp("04 Planilla 9"), carp("01 Sobrecostos")
        # El Reporte_CRA reemplaza al Calculo_CO y va en 04 Planilla 9.
        for x, r, falta in ((var_dest, arch(p9, "5_remuneracion_cra"), "5_REMUNERACIÓN_CRA"),
                            (var_sc, arch(sob, "calculo_sobrecostossscc"),
                             "Calculo_SobrecostosSSCC"),
                            (var_co, arch(p9, "reporte_cra", (".csv",)), "Reporte_CRA_15min")):
            x.set(str(r) if r else f"[{falta} no encontrado]")

    def sel_carpeta():
        ini = var_carpeta.get()
        r = filedialog.askdirectory(title="Selecciona 02 CASO RELIQUIDACION",
                                    initialdir=ini if ini and Path(ini).is_dir() else "")
        if r:
            var_carpeta.set(r)
            guardar_config({"carpeta_reliq": r})
            modo["traspaso"] = False
            refrescar()

    def lanzar():
        hacer = [c for c, x in (("SC", sel_sc), ("CO", sel_co)) if x.get()]
        if not hacer:
            messagebox.showwarning("Sin selección", "Marca al menos uno: SC o CO.")
            return
        d = var_dest.get()
        if not d or d.startswith("[") or not Path(d).is_file():
            messagebox.showerror("Falta el archivo", "No se encontró la planilla 5.")
            return
        rutas = {"destino": d}
        for cual in hacer:
            x = {"SC": var_sc, "CO": var_co}[cual].get()
            if not x or x.startswith("[") or not Path(x).is_file():
                messagebox.showerror("Falta el archivo",
                                     f"Para traer los {cual} hace falta su origen.")
                return
            rutas[cual] = x
        try:
            with open(d, "r+b"):
                pass
        except PermissionError:
            messagebox.showwarning("El archivo está en uso",
                                   f"Ciérralo en Excel y vuelve a intentar:\n\n{Path(d).name}")
            return
        except OSError:
            pass
        v.limpiar_log()
        v.log(f"Destino : {d}")
        for cual in hacer:
            v.log(f"Origen {cual}: {rutas[cual]}")

        def fin(r):
            ok, msg = r
            if ok:
                v.var_estado.set(f"Listo — {msg}")
                messagebox.showinfo("Listo", f"{msg}\n\nEl libro quedó abierto en Excel.")
            else:
                v.var_estado.set("Terminó con errores — revisa el log")
                messagebox.showerror("Error", f"Falló.\n\n{msg}")

        v.correr(lambda: ejecutar(rutas, hacer, var_aamm.get().strip(), v.log, v.progreso),
                 fin, boton=btn)

    btn = v.boton("ACTUALIZAR", lanzar, principal=True)
    v.boton("Carpeta…", sel_carpeta)
    v.boton("Salir", v.root.destroy)
    refrescar()
    v.iniciar()


if __name__ == "__main__":
    main()
