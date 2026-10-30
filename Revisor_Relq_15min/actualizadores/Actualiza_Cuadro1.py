# =============================================================================
#  Actualiza la hoja "01.SSCC_Recurso_Técnico" de 1_CUADROS_PAGO_SSCC  (15 MIN)
# =============================================================================
#  Junta, en una sola tabla apilada, los pagos por empresa de:
#
#    PLANILLA 4  (4_REMUNERACIÓN_SC_CO_Y_Pagos_Retiros_*), una hoja por tipo:
#        hojas PAGO_EMPRESAS_*  (SC_SSCC, CO, CO_ERNC, BESS, CPF(-), CRA, ...)
#        B1 (celda combinada B1:D1) -> el tipo de pago
#        A3:D  Empresa | RECIBE | PAGA | Paga+Recibe   (la fila TOTAL no entra)
#
#    PLANILLA 8  (8_REMUNERACIÓN_CONTROL_CONTINGENCIAS_*), hoja RESUMEN, fila 9:
#        C tipo | D empresa | E recibe | F paga | G neto
#
#  DESTINO, hoja 01.SSCC_Recurso_Técnico, desde la fila 9:
#      B  tipo de pago        C  Empresa      D  RECIBE      E  PAGA     F  SEN
#      A y G son formula: se estiran o se cortan al largo de la tabla.
#      Las filas con RECIBE y PAGA en 0 no se pegan (se ve mas limpio).
#      I:K  Empresa | Paga | Recibe, la SUMA por empresa (todos los tipos):
#           tres formulas de matriz dinamica en I9, J9 y K9 que se derraman
#           solas. I9 = UNICOS de C sin vacios y sin las empresas cuyo total da
#           0; J9 / K9 = SUMAR.SI sobre I9#. Los rangos $C$9:$C$<ultima> van al
#           largo de la tabla (antes, fijos hasta la 4346). Se escriben con
#           Formula2: con Formula, Excel les pondria "@" y no se derramarian.
#      F3:G5  los totales de arriba, tambien con el largo de la tabla.
#
#  PLANILLA 4 Y PLANILLA 8, POR SEPARADO
#  ----------------------------------------------------------------------------
#  Se elige cual actualizar (una, la otra o las dos). La que NO se elige se
#  CONSERVA tal como estaba en la planilla 1. Para saber de donde vino cada fila
#  ya pegada se mira su tipo (columna B):
#    * de la planilla 4: los tipos de las hojas PAGO_EMPRESAS_* (el B1 de cada
#      una, leido del archivo de la planilla 4);
#    * de la planilla 8: los tipos de la columna C de su RESUMEN; si ese
#      archivo no esta, todo lo que no sea de la planilla 4.
#  Los tipos se comparan sin espacios ni guiones bajos ("CO ERNC" = "CO_ERNC").
#  Una fila que no es de ninguna de las dos (restos de un formato viejo) se
#  descarta y se avisa.
#  Orden: primero la planilla 4 (sus hojas en el orden del libro), despues la 8.
#
#  Las planillas 4 y 8 solo se LEEN, sin abrir Excel (ZIP/XML). La planilla 1
#  se escribe con xlwings, preservando macros, en calculo manual; al final se
#  recalcula, se guarda y queda abierta, como los demas actualizadores.
#
#  Se puede correr solo, o recibir del Revisor la ruta de un JSON de traspaso.
# =============================================================================

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import re, sys, time, traceback, zipfile

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

# La ventana chica comun de esta version (actualizadores/ventana_chica.py).
try:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "actualizadores"))
    from ventana_chica import VentanaChica
except ImportError as e:
    _morir("Falta ventana_chica.py",
           "No se pudo cargar actualizadores/ventana_chica.py.\n\n"
           "Copia la carpeta Revisor_Relq_15min completa, no los .py sueltos.\n\n"
           f"Detalle: {e}")



# =============================================================================
#  CONFIGURACION
# =============================================================================
HOJA_DESTINO = "01.SSCC_Recurso_Técnico"
FILA_DESTINO = 9
COLS_DATOS = ("B", "F")            # tipo, empresa, recibe, paga, sen
COLS_FORMULA = ("A", "G")          # se estiran desde la fila 9
COLS_EMPRESAS = ("I", "K")         # empresa, paga, recibe (formulas)

# Planilla 4
PREFIJO_HOJAS_P4 = "PAGO_EMPRESA"  # PAGO_EMPRESAS_* (con o sin S, por las dudas)
FILA_DATOS_P4 = 3                  # A3:D
CELDA_TIPO_P4 = ("B", 1)           # combinada B1:D1

# Planilla 8
HOJA_P8 = "RESUMEN"
FILA_DATOS_P8 = 9
COLS_P8 = {"tipo": "C", "empresa": "D", "recibe": "E", "paga": "F", "neto": "G"}

CARPETA_ENT = "00 Entregables"
PATRONES = {
    "p4": r"^4_remuneracion_sc",
    "p8": r"^8_remuneracion_control",
    "cuadro_1": r"^1_cuadros",
}

# Formulas de I9, J9 y K9, en ingles y con coma (por COM no se pueden escribir
# en espanol): UNIQUE = UNICOS, FILTER = FILTRAR, SUMIF = SUMAR.SI. Son de
# matriz dinamica: se derraman hacia abajo solas, una fila por empresa.
# {u} es la ultima fila de la tabla. Una empresa cuyo total da 0 (o casi: el
# mismo 0.00001 de siempre) no entra: queda la lista limpia.
F_EMPRESAS = ('=LET(c,$C$9:$C${u},e,UNIQUE(FILTER(c,c<>"")),'
              's,SUMIF(c,e,$G$9:$G${u}),FILTER(e,ABS(s)>0.00001,""))')
F_PAGA = '=LET(s,SUMIF($C$9:$C${u},I9#,$G$9:$G${u}),IF(s<-0.00001,s,""))'
F_RECIBE = '=LET(s,SUMIF($C$9:$C${u},I9#,$G$9:$G${u}),IF(s>0.00001,s,""))'
# Tolerancia para decir que una fila o una empresa "es 0".
EPS_CERO = 0.00001
# Totales de arriba (F3:G5)
F_TOTALES = {
    "F3": "=+SUM(F9:F{u})", "G3": "=+SUM(G9:G{u})",
    "F4": '=+SUMIF(F$9:F${u},">0")', "G4": '=+SUMIF(G$9:G${u},">0")',
    "F5": '=+SUMIF(F$9:F${u},"<0")', "G5": '=+SUMIF(G$9:G${u},"<0")',
}


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
clave_tipo = _texto.clave_mayusculas     # "CO ERNC" == "CO_ERNC"


def col_num(letra):
    n = 0
    for c in str(letra).upper():
        n = n * 26 + (ord(c) - 64)
    return n


def _vacio(v):
    return v is None or (isinstance(v, str) and not v.strip())


def _es_total(v):
    return isinstance(v, str) and normalizar(v) == "total"


def buscar_archivo(carpeta_reliq, patron):
    """El mas reciente de 00 Entregables que calce, sin temporales ni copias."""
    ent = next((s for s in Path(carpeta_reliq).iterdir() if s.is_dir()
                and normalizar(s.name) == normalizar(CARPETA_ENT)), None)
    if ent is None:
        return None
    pat = re.compile(patron)
    cand = [f for f in ent.iterdir()
            if f.is_file() and f.suffix.lower() in (".xlsx", ".xlsm")
            and not f.name.startswith("~$")
            and not re.search(r"(- ?copia( \(\d+\))?|- ?copy|\(\d+\))$",
                              f.stem, re.IGNORECASE)
            and pat.search(normalizar(f.stem))]
    return max(cand, key=lambda f: f.stat().st_mtime) if cand else None


# =============================================================================
#  LECTURA DE LOS ORIGENES (sin Excel)
# =============================================================================
def hojas_de(ruta):
    with zipfile.ZipFile(str(ruta)) as z:
        return _excel_xml.ubicar_hoja_xml(z, "#1")[1]


def leer_p4(ruta, log):
    """(tipos en orden, filas). filas: [tipo, empresa, recibe, paga, sen]."""
    hojas = [h for h in hojas_de(ruta)
             if normalizar(h).upper().startswith(PREFIJO_HOJAS_P4)]
    if not hojas:
        raise RuntimeError(f"No hay hojas {PREFIJO_HOJAS_P4}* en {Path(ruta).name}")
    log(f"  Planilla 4: {len(hojas)} hoja(s) {PREFIJO_HOJAS_P4}*")
    tipos, filas = [], []
    c_t, f_t = CELDA_TIPO_P4
    for h in hojas:
        d = _excel_xml.leer_columnas_rapido(ruta, h, ["A", "B", "C", "D"], 1, log)
        if d is None:
            raise RuntimeError(f"No se pudo leer la hoja {h} de {Path(ruta).name}")
        tipo = d.get(c_t, {}).get(f_t)
        if _vacio(tipo):
            raise RuntimeError(f"La hoja {h} no tiene el tipo de pago en {c_t}{f_t}")
        tipo = str(tipo).strip()
        tipos.append(tipo)
        n = 0
        for f in sorted(x for x in d.get("A", {}) if x >= FILA_DATOS_P4):
            emp = d["A"][f]
            if _vacio(emp) or _es_total(emp):
                continue
            filas.append([tipo, str(emp).strip(), d.get("B", {}).get(f),
                          d.get("C", {}).get(f), d.get("D", {}).get(f)])
            n += 1
        log(f"    {h:<30} {tipo:<10} {n} empresa(s)")
    return tipos, filas


def leer_p8(ruta, log):
    """(tipos en orden, filas) del RESUMEN de la planilla 8."""
    c = COLS_P8
    cols = [c["tipo"], c["empresa"], c["recibe"], c["paga"], c["neto"]]
    d = _excel_xml.leer_columnas_rapido(ruta, HOJA_P8, cols, FILA_DATOS_P8, log)
    if d is None:
        raise RuntimeError(f"No se pudo leer la hoja {HOJA_P8} de {Path(ruta).name}")
    g = lambda col, f: d.get(col, {}).get(f)
    tipos, filas = [], []
    for f in sorted(d.get(c["empresa"], {})):
        emp = g(c["empresa"], f)
        if _vacio(emp) or _es_total(emp):
            continue
        tipo = g(c["tipo"], f)
        if _vacio(tipo) or _es_total(tipo):
            continue
        tipo = str(tipo).strip()
        if tipo not in tipos:
            tipos.append(tipo)
        filas.append([tipo, str(emp).strip(), g(c["recibe"], f), g(c["paga"], f),
                      g(c["neto"], f)])
    log(f"  Planilla 8: {len(filas)} fila(s), tipo(s) {', '.join(tipos) or '—'}")
    return tipos, filas


# =============================================================================
#  ARMADO (sin Excel)
# =============================================================================
def combinar(existentes, tipos_p4, tipos_p8, nuevas_p4, nuevas_p8,
             hacer_p4, hacer_p8, log):
    """La tabla final, B:F. La planilla que no se actualiza se toma de las filas
    que ya estaban (`existentes`), clasificadas por su tipo.

    tipos_p8 None: no se pudo leer la planilla 8, y entonces es de la 8 todo lo
    que no es de la 4. Devuelve (filas, n_p4, n_p8)."""
    k4 = {clave_tipo(t) for t in tipos_p4}
    k8 = None if tipos_p8 is None else {clave_tipo(t) for t in tipos_p8}
    viejas_p4, viejas_p8, huerfanas = [], [], {}
    for fila in existentes:
        k = clave_tipo(fila[0])
        if k in k4:
            viejas_p4.append(fila)
        elif k8 is None or k in k8:
            viejas_p8.append(fila)
        else:
            huerfanas[str(fila[0])] = huerfanas.get(str(fila[0]), 0) + 1

    bloque4 = nuevas_p4 if hacer_p4 else viejas_p4
    bloque8 = nuevas_p8 if hacer_p8 else viejas_p8
    log(f"  Planilla 4: {len(bloque4)} fila(s) "
        + ("traídas del origen" if hacer_p4 else "conservadas de la planilla 1"))
    log(f"  Planilla 8: {len(bloque8)} fila(s) "
        + ("traídas del origen" if hacer_p8 else "conservadas de la planilla 1"))
    # Lo que se descarta solo importa si no se esta reemplazando todo.
    if huerfanas and not (hacer_p4 and hacer_p8):
        det = ", ".join(f"{t} ({n})" for t, n in huerfanas.items())
        log(f"  OJO: se descartan filas de la planilla 1 que no son de la 4 ni de "
            f"la 8 (¿formato viejo?): {det}")
    return bloque4 + bloque8, len(bloque4), len(bloque8)


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def sin_ceros(filas):
    """(filas sin las que traen RECIBE y PAGA en 0, cuantas se sacaron)."""
    quedan = [f for f in filas
              if abs(_num(f[2])) > EPS_CERO or abs(_num(f[3])) > EPS_CERO]
    return quedan, len(filas) - len(quedan)


def n_empresas(filas):
    """Las empresas que va a listar I9: distintas en C (UNICOS no distingue
    mayusculas, asi que aca tampoco) y con total (columna SEN) distinto de 0.
    Es solo para el log: la lista la arma la formula."""
    tot = {}
    for f in filas:
        if _vacio(f[1]):
            continue
        k = str(f[1]).strip().casefold()
        tot[k] = tot.get(k, 0.0) + _num(f[4])
    return sum(1 for v in tot.values() if abs(v) > EPS_CERO)


def formulas_empresas(ultima):
    """Las formulas de I9, J9 y K9 para una tabla que llega a `ultima`."""
    return [f.format(u=ultima) for f in (F_EMPRESAS, F_PAGA, F_RECIBE)]


# =============================================================================
#  ESCRITURA (xlwings)
# =============================================================================
def ultima_fila(sh, col, desde=1):
    try:
        f = sh.api.Cells(sh.api.Rows.Count, col_num(col)).End(-4162).Row   # xlUp
        return max(int(f), desde - 1)
    except Exception:
        return desde - 1


def actualizar_en_libro(sh, rutas, hacer_p4, hacer_p8, log, progreso=lambda *a: None):
    """Trabaja sobre la hoja destino YA ABIERTA. No recalcula ni guarda.
    Devuelve (n_filas, n_p4, n_p8, n_empresas)."""
    tipos_p4 = nuevas_p4 = tipos_p8 = nuevas_p8 = None

    progreso(10, 100, "Leyendo planilla 4...")
    if rutas.get("p4"):
        tipos_p4, nuevas_p4 = leer_p4(rutas["p4"], log)
    elif hacer_p4:
        raise RuntimeError("Falta la planilla 4.")
    progreso(30, 100, "Leyendo planilla 8...")
    if rutas.get("p8"):
        tipos_p8, nuevas_p8 = leer_p8(rutas["p8"], log)
    elif hacer_p8:
        raise RuntimeError("Falta la planilla 8.")
    if tipos_p4 is None:
        raise RuntimeError("Hace falta la planilla 4 aunque no se actualice: de "
                           "ahí salen los tipos que se conservan.")

    # --- lo que ya estaba -----------------------------------------------------
    c1, c2 = col_num(COLS_DATOS[0]), col_num(COLS_DATOS[1])
    previa = FILA_DESTINO - 1
    for c in ("A", "B", "C", "G", "I", "J", "K"):
        previa = max(previa, ultima_fila(sh, c, FILA_DESTINO))
    existentes = []
    if previa >= FILA_DESTINO:
        bruto = sh.range((FILA_DESTINO, c1), (previa, c2)).options(ndim=2).value or []
        existentes = [list(f) for f in bruto if f and not _vacio(f[0])]
    log(f"  En la planilla 1 había {len(existentes)} fila(s) (hasta la {previa})")

    filas, n4, n8 = combinar(existentes, tipos_p4, tipos_p8, nuevas_p4 or [],
                             nuevas_p8 or [], hacer_p4, hacer_p8, log)
    f4, c4 = sin_ceros(filas[:n4])
    f8, c8 = sin_ceros(filas[n4:])
    filas, n4, n8, n_cero = f4 + f8, len(f4), len(f8), c4 + c8
    if n_cero:
        log(f"  {n_cero} fila(s) con RECIBE y PAGA en 0 no se pegan")
    if not filas:
        raise RuntimeError("No quedó ninguna fila para escribir.")
    n = len(filas)
    ultima = FILA_DESTINO + n - 1
    n_emp = n_empresas(filas)

    # --- limpiar y escribir B:F ---------------------------------------------
    progreso(55, 100, "Escribiendo...")
    if previa >= FILA_DESTINO:
        sh.range((FILA_DESTINO, c1), (previa, c2)).clear_contents()
    sh.range((FILA_DESTINO, c1)).value = filas
    log(f"  B{FILA_DESTINO}:F{ultima}  ({n4} de la 4 + {n8} de la 8 = {n} filas)")

    # --- A y G: estirar o cortar ---------------------------------------------
    for col in COLS_FORMULA:
        c = col_num(col)
        modelo = sh.range((FILA_DESTINO, c)).formula
        if not (isinstance(modelo, str) and modelo.startswith("=")):
            log(f"  OJO: {col}{FILA_DESTINO} no tiene fórmula: la columna {col} no se toca")
            continue
        if previa > ultima:
            sh.range((ultima + 1, c), (previa, c)).clear_contents()
        if n > 1:
            sh.range((FILA_DESTINO, c)).api.AutoFill(
                sh.range((FILA_DESTINO, c), (ultima, c)).api, 0)    # xlFillDefault
    log(f"  A y G estirados hasta la fila {ultima}")

    # --- I:K -------------------------------------------------------------------
    # Se limpia TODO I9:K hacia abajo: una formula que derrama necesita las
    # celdas de abajo vacias, o da #¡DESBORDAMIENTO!.
    i1, i2 = col_num(COLS_EMPRESAS[0]), col_num(COLS_EMPRESAS[1])
    if previa >= FILA_DESTINO:
        sh.range((FILA_DESTINO, i1), (previa, i2)).clear_contents()
    for c, f in zip(range(i1, i2 + 1), formulas_empresas(ultima)):
        sh.range((FILA_DESTINO, c)).api.Formula2 = f
    log(f"  I{FILA_DESTINO}:K{FILA_DESTINO}: UNICOS sin vacíos ni ceros "
        f"(~{n_emp} empresa(s)), rangos hasta la fila {ultima}")

    # --- totales de arriba -----------------------------------------------------
    for celda, f in F_TOTALES.items():
        sh.range(celda).formula = f.format(u=ultima)
    log(f"  F3:G5 sumando hasta la fila {ultima}")
    return n, n4, n8, n_emp


def ejecutar(rutas, hacer_p4, hacer_p8, log, progreso):
    """Abre la planilla 1, actualiza, recalcula y guarda. (ok, resumen)."""
    try:
        import pythoncom
        import xlwings as xw
    except ImportError as e:
        log(f"ERROR: falta una librería: {e}")
        return False, str(e)
    pythoncom.CoInitialize()
    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        log(f"  Abriendo destino: {Path(rutas['cuadro_1']).name}")
        wb = app.books.open(str(rutas["cuadro_1"]), update_links=False)
        wb.app.calculation = "manual"
        sh = next((s for s in wb.sheets
                   if normalizar(s.name) == normalizar(HOJA_DESTINO)), None)
        if sh is None:
            raise RuntimeError(f"No está la hoja '{HOJA_DESTINO}' en el destino.")
        n, n4, n8, n_emp = actualizar_en_libro(sh, rutas, hacer_p4, hacer_p8,
                                               log, progreso)
        progreso(85, 100, "Recalculando...")
        wb.app.calculation = "automatic"
        t0 = time.time()
        app.calculate()
        log(f"  recálculo: {time.time() - t0:.1f} s")
        log("  Guardando...")
        wb.save()
        log("  Guardado.")
        app.api.Visible = True
        app.screen_updating = True
        app.display_alerts = True
        sh.activate()
        progreso(100, 100, "Listo")
        return True, f"{n4} de la 4 + {n8} de la 8 = {n} filas, {n_emp} empresas"
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
#  VENTANA (chica, a proposito)
# =============================================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    v = VentanaChica("Planilla 1 — 01.SSCC_Recurso_Técnico",
                     "Trae los pagos por empresa. La planilla que no se marca se "
                     "conserva como está.", traspaso, alto=360)
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    vars_r = {k: tk.StringVar(value="[pendiente]") for k in ("cuadro_1", "p4", "p8")}
    # Viniendo del Revisor, en blanco (convencion). A mano, la 4 marcada: la 8
    # por lo general no se actualiza.
    sel_p4 = tk.BooleanVar(value=not modo["traspaso"])
    sel_p8 = tk.BooleanVar(value=False)

    fr = v.seccion("Qué actualizar")
    v.casilla(fr, "Planilla 4", sel_p4, "hojas PAGO_EMPRESAS_*")
    v.archivo(fr, "", vars_r["p4"], sangria=24)
    v.casilla(fr, "Planilla 8", sel_p8, "RESUMEN (control de contingencias)")
    v.archivo(fr, "", vars_r["p8"], sangria=24)
    fd = v.seccion("Destino")
    v.archivo(fd, "Planilla 1", vars_r["cuadro_1"])

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            for k in vars_r:
                vars_r[k].set(d.get(k) or "[el Revisor no la mandó]")
        else:
            base = var_carpeta.get()
            if base and Path(base).is_dir():
                for k, pat in PATRONES.items():
                    r = buscar_archivo(base, pat)
                    vars_r[k].set(str(r) if r else f"[no está en {CARPETA_ENT}]")

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
        h4, h8 = sel_p4.get(), sel_p8.get()
        if not (h4 or h8):
            messagebox.showwarning("Sin selección", "Marca la planilla 4, la 8 o las dos.")
            return
        rutas = {k: x.get() for k, x in vars_r.items()
                 if not x.get().startswith("[") and Path(x.get()).is_file()}
        if "cuadro_1" not in rutas:
            messagebox.showerror("Falta el archivo", "No se encontró la planilla 1.")
            return
        if "p4" not in rutas:
            messagebox.showerror("Falta el archivo",
                                 "Hace falta la planilla 4 (aunque no se actualice, "
                                 "de ahí salen los tipos que se conservan).")
            return
        if h8 and "p8" not in rutas:
            messagebox.showerror("Falta el archivo", "No se encontró la planilla 8.")
            return
        try:
            with open(rutas["cuadro_1"], "r+b"):
                pass
        except PermissionError:
            messagebox.showwarning("El archivo está en uso",
                                   "Cierra la planilla 1 en Excel y vuelve a intentar.")
            return
        except OSError:
            pass
        v.limpiar_log()
        v.log(f"Planilla 4: {'ACTUALIZAR' if h4 else 'se conserva'}   |   "
              f"Planilla 8: {'ACTUALIZAR' if h8 else 'se conserva'}")

        def fin(r):
            ok, msg = r
            if ok:
                v.var_estado.set(f"Listo — {msg}")
                messagebox.showinfo("Listo", f"{msg}\n\nLa planilla 1 quedó abierta en Excel.")
            else:
                v.var_estado.set("Terminó con errores — revisa el log")
                messagebox.showerror("Error", f"Falló.\n\n{msg}")

        v.correr(lambda: ejecutar(rutas, h4, h8, v.log, v.progreso), fin, boton=btn)

    btn = v.boton("ACTUALIZAR", lanzar, principal=True)
    v.boton("Carpeta…", sel_carpeta)
    v.boton("Salir", v.root.destroy)
    refrescar()
    v.iniciar()


if __name__ == "__main__":
    main()
