# =============================================================================
#  Actualiza la hoja "SC y CO" de la planilla 5_
# =============================================================================
#  Pega los SC y los CO de los EMBALSES, con su prorrata de instruccion directa.
#
#  ORIGEN Y DESTINO
#  ----------------------------------------------------------------------------
#  Los dos bloques van en la MISMA hoja, uno debajo del otro: primero SC, despues
#  CO. Los dos escriben en las mismas columnas.
#
#  SC — Calculo_SobrecostosSSCC, hoja "SOBRECOSTOS", desde la fila 7:
#      S -> C   Clave Año_Mes
#      T -> D   tipo (SCCF)
#      U -> E   central          <- por aca se filtra
#      V -> F   hora mensual
#      W -> G   monto
#      AU:BJ -> I:X   prorrata (16 columnas)
#
#  CO — Calculo_CO, hoja "PRORRATA CO", desde la fila 7:
#      (la Clave Año_Mes no viene: se copia la de los SC)  -> C
#      B -> D   tipo (CO)
#      D -> E   central          <- por aca se filtra
#      F -> F   hora mensual
#      G -> G   monto
#      BM:CB -> I:X   prorrata (16 columnas)
#
#  Y:AF del destino son formulas (8 columnas): se estiran o se cortan para cubrir
#  exactamente las filas que quedaron.
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
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from datetime import datetime
import json, subprocess, sys, re, socket, os, traceback, unicodedata, time
import threading, queue

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
#  OJO: esta lista esta TAMBIEN en Revisor_Reliquidacion.py. Si se agrega o
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
#  (columna U de Calculo_SobrecostosSSCC / columna D de Calculo_CO).
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

# Bloque de identificacion (C:G) y de prorrata (I:X), y las formulas (Y:AF).
COLS_ID = ("C", "G")
COLS_PRO = ("I", "X")
# Bloques de formulas, EN BLOQUES y no un rango corrido: la AC queda AFUERA a
# proposito. Si se hiciera AutoFill de Y:AF de una, la AC se rellenaria tambien
# con lo que tenga en la fila 9, y esa columna no lleva formula.
BLOQUES_FORMULA = [("Y", "AB"), ("AD", "AF")]

COL_TIPO = "D"          # SCCF o CO
COL_CLAVE = "C"         # Clave Año_Mes
COL_CENTRAL = "E"
N_ID = 5                # C:G
N_PRO = 16              # I:X

TIPO_SC = "SCCF"
TIPO_CO = "CO"

FUENTES = {
    "SC": {
        "etiqueta": "SC  —  Calculo_SobrecostosSSCC, hoja SOBRECOSTOS",
        "clave_traspaso": "calculo_sscc_maestro",
        "hoja": "SOBRECOSTOS",
        "fila_ini": 7,
        # bloques de la parte de identificacion, EN EL ORDEN en que se pegan.
        # La Clave Año_Mes viene en el origen (columna S).
        "bloques_id": [("S", "W")],
        "trae_clave": True,
        "col_central": "U",
        "bloques_pro": [("AU", "BJ")],
    },
    "CO": {
        "etiqueta": "CO  —  Calculo_CO, hoja PRORRATA CO",
        "clave_traspaso": "calculo_co",
        "hoja": "PRORRATA CO",
        "fila_ini": 7,
        # Aca NO viene la Clave Año_Mes: se completa con la de los SC.
        "bloques_id": [("B", "B"), ("D", "D"), ("F", "G")],
        "trae_clave": False,
        "col_central": "D",
        "bloques_pro": [("BM", "CB")],
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
except ImportError as e:
    _morir("Falta la carpeta __comun__/",
           "No se pudo cargar __comun__/config.py.\n\n"
           "Tiene que estar la carpeta '__comun__' hermana de Revisor_Relq,\n"
           "con config.py adentro. Baja el repositorio completo, no los .py sueltos.\n\n"
           f"Carpeta actual: {DIR_SCRIPT}\n\nDetalle: {e}")

get_usuario = _cfg.clave_equipo
escribir_json = _cfg.escribir_json


def leer_config():
    return _cfg.leer(CONFIG_PATH)


def _modificar_config(mutador):
    return _cfg.modificar(CONFIG_PATH, mutador)


def guardar_config(data):
    return _cfg.guardar(CONFIG_PATH, data)


def abrir_en_explorador(ruta, es_archivo=False):
    if not ruta or str(ruta).startswith("["):
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


# =============================================================================
#  TRASPASO DESDE EL REVISOR
# =============================================================================
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


# =============================================================================
#  UTILIDADES
# =============================================================================
def normalizar(t):
    if t is None:
        return ""
    t = unicodedata.normalize("NFKD", str(t))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.lower().split())


def clave_central(t):
    """Normaliza un nombre de central para comparar: sin tildes, sin espacios ni
    guiones bajos, en mayusculas. Asi 'El Toro-1' y 'ELTORO-1' son la misma."""
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[\s_]+", "", t).upper()


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


def fmt_tiempo(seg):
    seg = int(seg)
    return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"


def _vacio(v):
    return v is None or (isinstance(v, str) and not v.strip())


# =============================================================================
#  LECTURA DEL ORIGEN
# =============================================================================
def leer_fuente(app, ruta, cfg, log):
    """Devuelve (filas_id, filas_pro, avisos).

    filas_id: lista de listas de 5 valores (C:G). Si la fuente no trae la Clave
              Año_Mes, el primer valor queda en None y se completa despues.
    filas_pro: lista de listas de 16 valores (I:X), alineada con filas_id.
    """
    permitidas = {clave_central(c) for c in CENTRALES_EMBALSE}
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

        # La ultima fila se busca en la columna de la central y en la primera del
        # primer bloque: si una tiene huecos, la otra salva.
        primera = cfg["bloques_id"][0][0]
        f_fin = max(ultima_fila(sh, col_c, f_ini), ultima_fila(sh, primera, f_ini))
        if f_fin < f_ini:
            raise RuntimeError(f"No hay datos en {col_c}{f_ini} hacia abajo.")
        log(f"    filas leídas: {f_ini} a {f_fin}  ({f_fin - f_ini + 1})")

        def bloque(c1, c2):
            d = sh.range((f_ini, col_num(c1)), (f_fin, col_num(c2))).options(ndim=2).value
            return d or []

        partes_id = [bloque(a, b) for a, b in cfg["bloques_id"]]
        partes_pro = [bloque(a, b) for a, b in cfg["bloques_pro"]]
        centrales = bloque(col_c, col_c)
    finally:
        try:
            wb.close()
        except Exception:
            pass

    filas_id, filas_pro = [], []
    n_total = f_fin - f_ini + 1
    vistas, descartadas, sospechosas = set(), {}, {}
    for i in range(n_total):
        nom = centrales[i][0] if i < len(centrales) and centrales[i] else None
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
        fid = []
        for p in partes_id:
            fid.extend(p[i] if i < len(p) else [])
        if not cfg["trae_clave"]:
            fid = [None] + fid          # el hueco de la Clave Año_Mes
        fid = (fid + [None] * N_ID)[:N_ID]
        fpro = []
        for p in partes_pro:
            fpro.extend(p[i] if i < len(p) else [])
        fpro = (fpro + [None] * N_PRO)[:N_PRO]
        filas_id.append(fid)
        filas_pro.append(fpro)

    log(f"    embalses: {len(filas_id)} fila(s) de {len(vistas)} central(es)")
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
    return filas_id, filas_pro, avisos


def leer_bloque_destino(sh, tipo, log):
    """Lee del propio destino las filas de un tipo (SCCF o CO).

    Sirve para conservar el bloque que el usuario NO eligio actualizar: como los
    dos bloques van uno debajo del otro, hay que reescribirlos juntos."""
    c_id1, c_id2 = (col_num(x) for x in COLS_ID)
    c_p1, c_p2 = (col_num(x) for x in COLS_PRO)
    ult = max(ultima_fila(sh, COL_TIPO, FILA_DESTINO),
              ultima_fila(sh, COL_CENTRAL, FILA_DESTINO))
    if ult < FILA_DESTINO:
        log(f"    el destino no tenía filas de {tipo}")
        return [], []
    ids = sh.range((FILA_DESTINO, c_id1), (ult, c_id2)).options(ndim=2).value or []
    pros = sh.range((FILA_DESTINO, c_p1), (ult, c_p2)).options(ndim=2).value or []
    objetivo = normalizar(tipo)
    idx_tipo = col_num(COL_TIPO) - c_id1
    fid, fpro = [], []
    for i, fila in enumerate(ids):
        if not fila or normalizar(fila[idx_tipo] if idx_tipo < len(fila) else None) != objetivo:
            continue
        fid.append((list(fila) + [None] * N_ID)[:N_ID])
        p = pros[i] if i < len(pros) else []
        fpro.append((list(p) + [None] * N_PRO)[:N_PRO])
    log(f"    conservadas {len(fid)} fila(s) de {tipo} que ya estaban en el destino")
    return fid, fpro


# =============================================================================
#  PROCESO
# =============================================================================
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
    avisos_todos = []
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        log(f"  Abriendo destino: {Path(rutas['destino']).name}")
        wb_d = app.books.open(str(rutas["destino"]), update_links=False)
        sh_d = buscar_hoja(wb_d, HOJA_DESTINO)
        if sh_d is None:
            raise RuntimeError(f"No está la hoja '{HOJA_DESTINO}' en el destino.")
        log(f"    hoja destino: {sh_d.name}")
        wb_d.app.calculation = "manual"

        # ---- juntar los dos bloques, se actualicen o no --------------------
        bloques = {}
        for cual in ORDEN:
            log("=" * 70)
            if cual in hacer:
                log(f"{cual}: leyendo del origen")
                fid, fpro, av = leer_fuente(app, rutas[cual], FUENTES[cual], log)
                avisos_todos += [f"{cual}: {a}" for a in av]
            else:
                log(f"{cual}: NO se actualiza, se conserva lo que ya está")
                tipo = TIPO_SC if cual == "SC" else TIPO_CO
                fid, fpro = leer_bloque_destino(sh_d, tipo, log)
            bloques[cual] = (fid, fpro)

        # ---- la Clave Año_Mes de los CO sale de los SC ---------------------
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
        if not filas_id:
            raise RuntimeError("No quedó ninguna fila para escribir.")

        # ---- limpiar y escribir -------------------------------------------
        progreso(50, 100, "Escribiendo...")
        log("=" * 70)
        log("Escribiendo en el destino")
        c_id1, c_id2 = (col_num(x) for x in COLS_ID)
        c_p1, c_p2 = (col_num(x) for x in COLS_PRO)
        bloques_f = [(col_num(a), col_num(b)) for a, b in BLOQUES_FORMULA]

        # Se limpia hasta la ultima fila con algo en cualquiera de los tres
        # bloques: el mes anterior pudo ser mas largo, y lo que sobre abajo
        # ensucia los totales.
        previa = 0
        cols_mirar = [c_id1, c_id2, c_p1, c_p2] + [a for a, _ in bloques_f]
        for c in cols_mirar:
            previa = max(previa, ultima_fila(sh_d, col_letra(c), FILA_DESTINO))
        n = len(filas_id)
        ultima = FILA_DESTINO + n - 1
        if previa >= FILA_DESTINO:
            log(f"    limpiando filas {FILA_DESTINO} a {previa}  "
                f"({previa - FILA_DESTINO + 1} del mes anterior)")
            sh_d.range((FILA_DESTINO, c_id1), (previa, c_id2)).clear_contents()
            sh_d.range((FILA_DESTINO, c_p1), (previa, c_p2)).clear_contents()
            if previa > ultima:
                # Las formulas que sobran se borran; las que quedan se estiran
                # mas abajo con AutoFill. Se borra bloque por bloque para no
                # tocar la AC.
                for (a, b), (la, lb) in zip(bloques_f, BLOQUES_FORMULA):
                    sh_d.range((ultima + 1, a), (previa, b)).clear_contents()
                log(f"    {' y '.join(f'{a}:{b}' for a, b in BLOQUES_FORMULA)}: "
                    f"cortado, sobraban {previa - ultima} fila(s)")
        else:
            log("    el destino estaba vacío")

        sh_d.range((FILA_DESTINO, c_id1)).value = filas_id
        sh_d.range((FILA_DESTINO, c_p1)).value = filas_pro
        log(f"    {COLS_ID[0]}{FILA_DESTINO}:{COLS_ID[1]}{ultima}  y  "
            f"{COLS_PRO[0]}{FILA_DESTINO}:{COLS_PRO[1]}{ultima}   "
            f"({len(bloques['SC'][0])} SC + {len(bloques['CO'][0])} CO = {n} filas)")

        # ---- estirar las formulas -----------------------------------------
        if n > 1:
            for a, b in bloques_f:
                origen = sh_d.range((FILA_DESTINO, a), (FILA_DESTINO, b))
                destino = sh_d.range((FILA_DESTINO, a), (ultima, b))
                origen.api.AutoFill(destino.api, 0)     # xlFillDefault
        log(f"    {' y '.join(f'{a}:{b}' for a, b in BLOQUES_FORMULA)} "
            f"estirado hasta la fila {ultima}   (la AC no se toca)")

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
        resumen = (f"{len(bloques['SC'][0])} SC + {len(bloques['CO'][0])} CO "
                   f"= {n} filas")
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

    root = tk.Tk()
    root.title('Actualiza hoja "SC y CO"'
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("1000x700")

    frame_btns = tk.Frame(root)
    frame_btns.pack(side="bottom", fill="x", pady=8)
    canvas = tk.Canvas(root, borderwidth=0, highlightthickness=0)
    scroll = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    cont = tk.Frame(canvas)
    win = canvas.create_window((0, 0), window=cont, anchor="nw")

    def _aj(_e=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(win, width=canvas.winfo_width())
    cont.bind("<Configure>", _aj)
    canvas.bind("<Configure>", _aj)
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
        fa = tk.Frame(cont, bg="#fff4c2", bd=1, relief="solid")
        fa.pack(fill="x", padx=20, pady=(10, 0))
        tk.Label(fa, text=f"Mes {traspaso.get('aamm') or '?'} — enviado por el Revisor",
                 bg="#fff4c2", font=("Segoe UI", 11, "bold")).pack(pady=(6, 0))
        tk.Label(fa, text="Las rutas las resolvió el Revisor.\n"
                          "Marca qué bloques traer: arrancan en blanco.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    tk.Label(cont, text='Actualizar la hoja "SC y CO" de la planilla 5_',
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(cont, text=f"Solo los embalses: {len(CENTRALES_EMBALSE)} centrales.\n"
                        "Los SC van arriba y los CO abajo; los dos bloques se "
                        "reescriben siempre, para que no queden desalineados.",
             font=("Segoe UI", 8), fg="#555555", justify="center").pack(pady=(0, 8))

    var_dest = tk.StringVar(value="[pendiente]")
    var_sc = tk.StringVar(value="[pendiente]")
    var_co = tk.StringVar(value="[pendiente]")
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_aamm = tk.StringVar(value=(traspaso or {}).get("aamm")
                            or cfg.get("ultimo_mes", ""))
    labels = {}

    def fila(titulo, var, clave):
        fr = tk.LabelFrame(cont, text=titulo, padx=10, pady=6)
        fr.pack(fill="x", padx=20, pady=4)
        l = tk.Label(fr, textvariable=var, wraplength=880, justify="left",
                     cursor="hand2", font=("Segoe UI", 9), anchor="w")
        l.pack(fill="x")
        l.bind("<Button-1>", lambda e: abrir_en_explorador(var.get(), es_archivo=True))
        labels[clave] = l
        return fr

    fila("Destino — 5_REMUNERACIÓN_CRA  (se modifica)", var_dest, "dest")
    fr_sc = fila("Origen SC — " + FUENTES["SC"]["etiqueta"], var_sc, "SC")
    sel_sc = tk.BooleanVar(value=not modo["traspaso"])
    tk.Checkbutton(fr_sc, text="Traer los SC", variable=sel_sc,
                   font=("Segoe UI", 9)).pack(anchor="w")
    fr_co = fila("Origen CO — " + FUENTES["CO"]["etiqueta"], var_co, "CO")
    sel_co = tk.BooleanVar(value=not modo["traspaso"])
    tk.Checkbutton(fr_co, text="Traer los CO", variable=sel_co,
                   font=("Segoe UI", 9)).pack(anchor="w")

    fm = tk.Frame(cont)
    fm.pack(fill="x", padx=24, pady=(4, 0))
    tk.Label(fm, text="Mes (AAMM):", font=("Segoe UI", 9)).pack(side="left")
    tk.Entry(fm, textvariable=var_aamm, width=8,
             font=("Consolas", 10)).pack(side="left", padx=6)
    tk.Label(fm, text="solo se usa si no hay filas SC de donde copiar la clave",
             font=("Segoe UI", 8), fg="#555555").pack(side="left")

    fc = tk.LabelFrame(cont, text="Carpeta 02 CASO RELIQUIDACION", padx=10, pady=6)
    fc.pack(fill="x", padx=20, pady=4)
    lc = tk.Label(fc, textvariable=var_carpeta, wraplength=880, justify="left",
                  cursor="hand2", font=("Segoe UI", 9), anchor="w")
    lc.pack(fill="x")
    lc.bind("<Button-1>", lambda e: abrir_en_explorador(var_carpeta.get()))
    labels["carpeta"] = lc

    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")
    fpr = tk.Frame(cont)
    fpr.pack(fill="x", padx=20, pady=(6, 0))
    tk.Label(fpr, textvariable=var_estado, font=("Segoe UI", 9),
             anchor="w").pack(side="left", fill="x", expand=True)
    tk.Label(fpr, textvariable=var_tiempo, font=("Consolas", 10, "bold"),
             fg="#2d7a2d").pack(side="right", padx=8)
    barra = ttk.Progressbar(cont, mode="determinate", length=400)
    barra.pack(fill="x", padx=20, pady=(2, 4))

    fl = tk.LabelFrame(cont, text="Progreso detallado", padx=6, pady=4)
    fl.pack(fill="both", expand=True, padx=20, pady=4)
    txt = tk.Text(fl, height=15, font=("Consolas", 9), wrap="none")
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
                    barra["maximum"] = t or 100
                    barra["value"] = a
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

    def pintar():
        for k, var in (("dest", var_dest), ("SC", var_sc), ("CO", var_co)):
            v = var.get()
            ok = bool(v) and not v.startswith("[")
            if ok:
                try:
                    ok = Path(v).is_file()
                except Exception:
                    ok = False
            labels[k].config(fg="blue" if ok else "red")
        v = var_carpeta.get()
        labels["carpeta"].config(fg="blue" if v and Path(v).is_dir() else "red")

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            var_dest.set(d.get("p5") or "[el Revisor no mandó p5]")
            var_sc.set(d.get("calculo_sscc_maestro")
                       or "[el Revisor no mandó calculo_sscc_maestro]")
            var_co.set(d.get("calculo_co") or "[el Revisor no mandó calculo_co]")
            pintar()
            return
        base = var_carpeta.get()
        if not base or not Path(base).is_dir():
            pintar()
            return
        base = Path(base)

        def carp(nombre, dentro=None):
            raiz = dentro or base
            for p in raiz.iterdir():
                if p.is_dir() and normalizar(p.name) == normalizar(nombre):
                    return p
            return None

        def arch(carpeta, pref):
            if carpeta is None:
                return None
            cand = [p for p in carpeta.glob("*")
                    if p.is_file() and normalizar(p.name).startswith(pref)
                    and p.suffix.lower() in (".xlsm", ".xlsx")
                    and not p.name.startswith("~$")]
            return max(cand, key=lambda p: p.stat().st_mtime) if cand else None

        p9 = carp("04 Planilla 9")
        sob = carp("01 Sobrecostos")
        ent = carp("00 Entregables")
        co_dir = carp("02 Costo de Oportunidad", ent) if ent else None
        d = arch(p9, "5_remuneracion_cra")
        s = arch(sob, "calculo_sobrecostossscc")
        c = arch(co_dir, "calculo_co")
        var_dest.set(str(d) if d else "[5_REMUNERACIÓN_CRA no encontrado]")
        var_sc.set(str(s) if s else "[Calculo_SobrecostosSSCC no encontrado]")
        var_co.set(str(c) if c else "[Calculo_CO no encontrado]")
        pintar()

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

    def lanzar():
        hacer = [c for c, v in (("SC", sel_sc), ("CO", sel_co)) if v.get()]
        if not hacer:
            messagebox.showwarning("Sin selección",
                                   "Marca al menos uno: SC o CO.")
            return
        d = var_dest.get()
        if not d or d.startswith("[") or not Path(d).is_file():
            messagebox.showerror("Falta el archivo",
                                 f"No se encontró la planilla 5_.\n\n{d}")
            return
        rutas = {"destino": d}
        for cual in hacer:
            v = {"SC": var_sc, "CO": var_co}[cual].get()
            if not v or v.startswith("[") or not Path(v).is_file():
                messagebox.showerror("Falta el archivo",
                                     f"Para traer los {cual} hace falta su origen."
                                     f"\n\n{v}")
                return
            rutas[cual] = v

        p = Path(d)
        try:
            with open(p, "r+b"):
                pass
        except PermissionError:
            messagebox.showwarning("El archivo está en uso",
                                   f"Ciérralo en Excel y volvé a intentar:\n\n{p.name}")
            return
        except OSError:
            pass

        no_elegido = [c for c in ORDEN if c not in hacer]
        extra = ""
        if no_elegido:
            extra = ("\n\nLos " + " y ".join(no_elegido) +
                     " no se traen del origen, pero sus filas se vuelven a escribir "
                     "en su lugar nuevo (si no, quedarían desalineadas).")
        if not messagebox.askyesno(
                "Confirmar",
                f"Sobre {p.name}, hoja «{HOJA_DESTINO}»:\n\n"
                f"  • traer {' y '.join(hacer)} de los embalses\n"
                f"  • se REEMPLAZA todo desde la fila {FILA_DESTINO}\n"
                f"  • se ajustan las fórmulas de "
                + " y ".join(f"{a}:{b}" for a, b in BLOQUES_FORMULA)
                + extra + "\n\n¿Seguir?"):
            return

        txt.delete("1.0", "end")
        log(f"Destino : {d}")
        for cual in hacer:
            log(f"Origen {cual}: {rutas[cual]}")
        if no_elegido:
            log(f"Se conservan las filas de: {', '.join(no_elegido)}")
        if modo["traspaso"]:
            log(f"Rutas enviadas por el Revisor — mes {traspaso.get('aamm') or '?'}")
        log("-" * 70)

        btn.config(state="disabled", bg="#aaaaaa")
        barra["value"] = 0
        var_estado.set("Procesando...")
        timer["on"] = True
        timer["t0"] = time.time()
        tick()

        def trabajo():
            ok, msg = ejecutar(rutas, hacer, var_aamm.get().strip(), log, progreso)

            def fin():
                timer["on"] = False
                btn.config(state="normal", bg="#2d7a2d")
                if ok:
                    var_estado.set(f"Listo — {msg}")
                    messagebox.showinfo(
                        "Listo",
                        f"Terminado.\n\n{msg}\n\nEl libro quedó abierto en Excel.\n"
                        "Si hubo avisos están en el log; después corré Verificar "
                        "en el Revisor.")
                else:
                    var_estado.set("Terminó con errores — revisa el log")
                    messagebox.showerror("Error", f"Falló.\n\n{msg}")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text='ACTUALIZAR "SC y CO"', bg="#2d7a2d",
                    fg="white", font=("Segoe UI", 10, "bold"), command=lanzar)
    btn.pack(side="left", padx=8, expand=True)
    tk.Button(frame_btns, text="Refrescar rutas",
              command=refrescar).pack(side="left", padx=8)
    tk.Button(frame_btns, text="Salir", command=root.destroy).pack(side="left", padx=8)

    refrescar()
    bombear()
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    root.geometry(f"+{(root.winfo_screenwidth() - w) // 2}"
                  f"+{(root.winfo_screenheight() - h) // 2}")
    root.mainloop()


if __name__ == "__main__":
    main()
