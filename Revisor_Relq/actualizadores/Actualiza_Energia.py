# =============================================================================
#  Actualizar Energia
# =============================================================================
#  Actualiza los dos entregables de "01.a Sobrecostos de Energia" a partir del
#  "02 Consolidado_Tabulado", hoja "Sobrecostos":
#
#    1) 03b ENTRADA_SOB_AAMM_*.mdb   (tabla [Sobrecostos] del Access)
#    2) Consolidado_AAMM_*.xlsm      (hoja "Sobrecostos")
#
#  En los dos casos se toman SOLO las filas cuyo Tipo sobrecosto es SCMT o SCPC,
#  que es justo lo que el revisor comprueba en V5 y V6.
#
#  El motor de Access se reutiliza de Actualiza_Data_Access.py en vez de
#  copiarlo: es el mismo destino (tabla [Sobrecostos], mismas 5 columnas, misma
#  regla de "borrar solo los tipos que trae el Excel"). Los dos .py tienen que
#  estar en la misma carpeta.
#
#  Se puede correr solo, o recibir del Revisor la ruta de un JSON de traspaso
#  como unico argumento.
# =============================================================================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import json, subprocess, sys, re, socket, os, traceback, unicodedata, time
import threading, queue

DIR_SCRIPT = Path(__file__).resolve().parent

# config.json es compartido con el Revisor y el resto de los actualizadores,
# y ahora vive en __config__, junto a Revisor_Relq. No es
# DIR_SCRIPT / "config.json" porque este script esta en actualizadores/.
CONFIG_PATH = DIR_SCRIPT.parent.parent / "__config__" / "config.json"

# --- motor de Access, reutilizado ------------------------------------------
# Este script NO duplica el motor de Access: usa el de Actualiza_Data_Access.py,
# que tiene que estar en la MISMA carpeta y ser una version que soporte fuentes
# externas y filtro por valores. Si se copia un .py y no el otro, la falla se
# avisa aca y no a mitad de proceso.
_AYUDA_COPIAR = (
    "Los dos archivos tienen que estar en la misma carpeta y ser de la misma\n"
    "version. Copia de nuevo Actualiza_Data_Access.py junto a este script.\n\n"
    f"Carpeta actual: {DIR_SCRIPT}"
)

def _morir(titulo, mensaje):
    """Aborta mostrando el motivo en una ventana. Sin esto, lanzado desde el
    Revisor con pythonw (sin consola) el script moriria callado y solo se veria
    que la ventana no aparece."""
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(titulo, mensaje)
        r.destroy()
    except Exception:
        pass
    print(f"{titulo}\n\n{mensaje}", file=sys.stderr)
    raise SystemExit(1)


try:
    import Actualiza_Data_Access as _ADA
    from Actualiza_Data_Access import (
        proceso as proceso_access,
        normalizar, buscar_carpeta, buscar_archivo, buscar_hoja,
        col_letra_a_num, ultima_fila, fmt_tiempo, es_vacio,
        actualizar_color_label,
    )
except ImportError as e:
    _morir("Falta Actualiza_Data_Access.py",
           "No se pudo cargar Actualiza_Data_Access.py.\n\n"
           f"{_AYUDA_COPIAR}\n\nDetalle: {e}")

_NECESITA = {"fuentes_externas", "filtro_por_valores"}
_TIENE = set(getattr(_ADA, "CAPACIDADES", ()))
if not _NECESITA <= _TIENE:
    _faltan = ", ".join(sorted(_NECESITA - _TIENE)) or "(no declara capacidades)"
    _morir("Actualiza_Data_Access.py está desactualizado",
           "El Actualiza_Data_Access.py que hay al lado es una versión ANTIGUA:\n"
           f"le falta {_faltan}.\n\n"
           "Sin eso, Energía no puede filtrar por SCMT/SCPC y el Access quedaría\n"
           "con datos de más.\n\n"
           f"{_AYUDA_COPIAR}")


# =============================================================================
#  CONFIGURACION
# =============================================================================
# El encabezado del "02 Consolidado_Tabulado" (hoja Sobrecostos) esta en la
# fila 2, asi que los datos arrancan en la 3. Si algun mes cambiara, es el unico
# numero que hay que tocar.
FILA_DATOS_TABULADO = 3

# Tipos que se traen. El resto se descarta.
TIPOS = ("SCMT", "SCPC")

HOJA_ORIGEN = "Sobrecostos"
HOJA_DESTINO_CONSOLIDADO = "Sobrecostos"

# ---- 1) Access -------------------------------------------------------------
# Columnas AA:AE del tabulado, que ya vienen en el mismo orden que la tabla:
#   AA Clave Año_Mes | AB Tipo sobrecosto | AC Central | AD Hora Mensual | AE Sobrecosto
# El filtro apunta al indice 1 de ese bloque, que es AB.
FUENTES_ENERGIA = {
    "MDB": {
        "etiqueta": "Access  (02 Consolidado_Tabulado / hoja Sobrecostos / AA:AE)",
        "hoja": HOJA_ORIGEN,
        "bloques": [("AA", "AE")],
        "fila_ini": FILA_DATOS_TABULADO,
        "filtrar_ceros": False,
        "filtro": {"col": 1, "valores": TIPOS},
    },
}

# ---- 2) Consolidado_AAMM ---------------------------------------------------
# Bloques de origen EN ESTE ORDEN -> se pegan corridos en A:J del destino.
# Ojo con el orden: la H del origen queda al final (columna J del destino).
BLOQUES_CONSOLIDADO = [("A", "G"), ("I", "J"), ("H", "H")]
COL_FILTRO_CONSOLIDADO = "AB"          # Tipo sobrecosto en el origen
COL_DESTINO_INI = "A"                  # se pega desde A2
FILA_DESTINO_INI = 2                   # la fila 1 es encabezado
N_COLS_CONSOLIDADO = 10                # A:G (7) + I:J (2) + H (1)


# =============================================================================
#  CONFIG COMPARTIDO  (mismo config.json que el Revisor y los otros dos)
# =============================================================================
def get_usuario():
    try:
        usuario = os.environ.get("USERNAME") or os.environ.get("USER") or "desconocido"
        return f"{socket.gethostname()}_{usuario}"
    except Exception:
        return "desconocido"


def leer_config():
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get(get_usuario(), {})
    except Exception:
        pass
    return {}


def escribir_json(ruta, data):
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_suffix(ruta.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, ruta)


def _modificar_config(mutador):
    """Solo agrega o actualiza claves. Si el archivo existe pero no se puede
    interpretar NO se escribe: mejor perder un ajuste que el archivo entero."""
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
#  BUSQUEDA DE ARCHIVOS  (respaldo para el modo manual)
# =============================================================================
def buscar_tabulado(carpeta_reliq):
    """01 Sobrecostos/<Detalles diarios>/02 Consolidado_Tabulado_AAMM_*"""
    sc = buscar_carpeta(carpeta_reliq, "01 Sobrecostos")
    if sc is None:
        return None
    for nombre in ("Detalles diarios", "Detalle diario",
                   "Detalle diarios", "Detalles diario"):
        det = buscar_carpeta(sc, nombre)
        if det is not None:
            f = buscar_archivo(det, r"consolidado_tabulado")
            if f is not None:
                return f
    return None


def carpeta_energia(carpeta_reliq):
    return buscar_carpeta(carpeta_reliq, "01.a Sobrecostos de Energia")


def buscar_mdb_energia(carpeta_reliq):
    carp = carpeta_energia(carpeta_reliq)
    if carp is None:
        return None
    # "03b entrada_sob_" pero NO el de SSCC, que vive en otra carpeta.
    return buscar_archivo(carp, r"03b entrada_sob_(?!sscc)",
                          extensiones=(".mdb", ".accdb"))


def buscar_consolidado_energia(carpeta_reliq):
    carp = carpeta_energia(carpeta_reliq)
    if carp is None:
        return None
    # "consolidado_" pero no "consolidado_cca" ni el tabulado.
    return buscar_archivo(carp, r"^consolidado_(?!cca|tabulado)")


# =============================================================================
#  ACTUALIZACION DEL Consolidado_AAMM  (Excel -> Excel)
# =============================================================================
def _matriz_filtrada(sh, log):
    """Lee el tabulado y devuelve las filas SCMT/SCPC ya reordenadas a 10
    columnas, en el orden en que se pegan: A:G, I:J, H."""
    cols = [(col_letra_a_num(a), col_letra_a_num(b)) for a, b in BLOQUES_CONSOLIDADO]
    c_filtro = col_letra_a_num(COL_FILTRO_CONSOLIDADO)
    f_ini = FILA_DATOS_TABULADO

    # La ultima fila se busca en la columna del filtro y en la primera del
    # primer bloque: si una tiene huecos, la otra salva.
    f_fin = max(ultima_fila(sh, c_filtro, f_ini),
                ultima_fila(sh, cols[0][0], f_ini))
    if f_fin < f_ini:
        log("    ADVERTENCIA: no hay datos bajo el encabezado.")
        return []
    log(f"    Origen: filas {f_ini} a {f_fin} ({f_fin - f_ini + 1} leidas)")

    def bloque(c1, c2):
        d = sh.range((f_ini, c1), (f_fin, c2)).options(ndim=2).value
        return d or []

    partes = [bloque(c1, c2) for c1, c2 in cols]
    tipos = bloque(c_filtro, c_filtro)

    permitidos = {normalizar(t) for t in TIPOS}
    salida, n_filtradas, n_vacias = [], 0, 0
    for i in range(f_fin - f_ini + 1):
        tipo = tipos[i][0] if i < len(tipos) and tipos[i] else None
        fila = []
        for p in partes:
            fila.extend(p[i] if i < len(p) else [])
        if all(es_vacio(v) for v in fila):
            n_vacias += 1
            continue
        if normalizar(tipo) not in permitidos:
            n_filtradas += 1
            continue
        # Se recorta o rellena para que siempre sean 10 columnas exactas.
        fila = (fila + [None] * N_COLS_CONSOLIDADO)[:N_COLS_CONSOLIDADO]
        salida.append(fila)

    if n_vacias:
        log(f"    Descartadas {n_vacias} filas vacias")
    log(f"    Descartadas {n_filtradas} filas de otro tipo "
        f"(se quedan solo {', '.join(TIPOS)})")
    log(f"    Filas utiles: {len(salida)}")
    return salida


def actualizar_consolidado(app_xw, ruta_tabulado, ruta_destino, log, progreso=None):
    """Pega en el Consolidado_AAMM las filas SCMT/SCPC del tabulado.
    Devuelve la cantidad de filas escritas."""
    wb_o = wb_d = None
    try:
        log(f"  Abriendo origen (solo lectura): {Path(ruta_tabulado).name}")
        wb_o = app_xw.books.open(str(ruta_tabulado), read_only=True, update_links=False)
        sh_o = buscar_hoja(wb_o, HOJA_ORIGEN)
        if sh_o is None:
            raise RuntimeError(f"No esta la hoja '{HOJA_ORIGEN}' en {Path(ruta_tabulado).name}")
        log(f"    Hoja origen: {sh_o.name}")
        if progreso:
            progreso(10, 100, "Leyendo el tabulado...")

        filas = _matriz_filtrada(sh_o, log)
        if not filas:
            raise RuntimeError("El tabulado no dejo ninguna fila SCMT/SCPC: no se escribe nada.")

        log(f"  Abriendo destino: {Path(ruta_destino).name}")
        wb_d = app_xw.books.open(str(ruta_destino), update_links=False)
        sh_d = buscar_hoja(wb_d, HOJA_DESTINO_CONSOLIDADO)
        if sh_d is None:
            raise RuntimeError(f"No esta la hoja '{HOJA_DESTINO_CONSOLIDADO}' "
                               f"en {Path(ruta_destino).name}")
        log(f"    Hoja destino: {sh_d.name}")
        wb_d.app.calculation = "manual"
        if progreso:
            progreso(40, 100, "Limpiando el destino...")

        c_ini = col_letra_a_num(COL_DESTINO_INI)
        c_fin = c_ini + N_COLS_CONSOLIDADO - 1

        # Se limpia lo viejo antes de pegar, porque el mes nuevo puede traer
        # MENOS filas que el anterior: si no, las de abajo quedarian pegadas y
        # el total de la columna E (lo que suma V6) saldria inflado.
        #
        # El largo anterior se toma del used_range y no mirando las columnas A y
        # J: si justo esas dos estan vacias en las ultimas filas pero hay dato en
        # el medio, mirar solo los extremos se queda corto y deja basura.
        # Pasarse no cuesta nada: esas celdas ya estan vacias.
        previa = 0
        try:
            previa = sh_d.used_range.last_cell.row
        except Exception:
            pass
        if previa < FILA_DESTINO_INI:
            # Respaldo por si used_range falla: se mira columna por columna.
            try:
                previa = max(ultima_fila(sh_d, c, FILA_DESTINO_INI)
                             for c in range(c_ini, c_fin + 1))
            except Exception:
                previa = 0

        if previa >= FILA_DESTINO_INI:
            log(f"    Borrando A{FILA_DESTINO_INI}:J{previa} "
                f"({previa - FILA_DESTINO_INI + 1} filas anteriores)")
            sh_d.range((FILA_DESTINO_INI, c_ini), (previa, c_fin)).clear_contents()
        else:
            log("    El destino estaba vacio bajo el encabezado.")

        if progreso:
            progreso(60, 100, "Pegando datos...")
        ultima_nueva = FILA_DESTINO_INI + len(filas) - 1
        log(f"    Pegando {len(filas)} filas en A{FILA_DESTINO_INI}:J{ultima_nueva}")
        sh_d.range((FILA_DESTINO_INI, c_ini)).value = filas
        if previa > ultima_nueva:
            log(f"    (el mes anterior tenia hasta la fila {previa}: "
                f"{previa - ultima_nueva} fila(s) quedaron limpias)")
        elif previa and ultima_nueva > previa:
            log(f"    (el mes anterior llegaba hasta la fila {previa}: "
                f"ahora hay {ultima_nueva - previa} fila(s) mas)")

        wb_d.app.calculation = "automatic"
        if progreso:
            progreso(85, 100, "Guardando...")
        log("    Guardando...")
        wb_d.save()
        log("    Guardado OK")

        try:
            wb_o.close()
            wb_o = None
        except Exception:
            pass
        return len(filas), wb_d
    except Exception:
        for wb in (wb_o, wb_d):
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass
        raise


# =============================================================================
#  PROCESO COMPLETO
# =============================================================================
def ejecutar(rutas, hacer_mdb, hacer_consolidado, solo_prueba, log, progreso):
    """rutas: {"tabulado","mdb","consolidado"}. Devuelve (ok, resumen:str)."""
    import pythoncom
    import xlwings as xw

    resumen = []
    pythoncom.CoInitialize()
    app = None
    wb_abierto = None
    try:
        # ---- 1) Consolidado_AAMM (Excel -> Excel) --------------------------
        if hacer_consolidado:
            log("=" * 70)
            log("Consolidado_AAMM  (hoja Sobrecostos, A:J desde la fila 2)")
            log("=" * 70)
            if solo_prueba:
                log("  [MODO PRUEBA] el Consolidado no se toca en modo prueba.")
                resumen.append("Consolidado: omitido (modo prueba)")
            else:
                app = xw.App(visible=False, add_book=False)
                app.display_alerts = False
                app.screen_updating = False
                n, wb_abierto = actualizar_consolidado(
                    app, rutas["tabulado"], rutas["consolidado"], log, progreso)
                resumen.append(f"Consolidado: {n} filas")
                # Se deja visible y abierto, igual que Actualiza_datos.py.
                app.api.Visible = True
                app.screen_updating = True
                app.display_alerts = True
                wb_abierto.activate()
                wb_abierto = None
                app = None
                log("")

        # ---- 2) Access (reusa el motor de Actualiza_Data_Access) -----------
        if hacer_mdb:
            log("=" * 70)
            log("Access  03b ENTRADA_SOB  (tabla [Sobrecostos])")
            log("=" * 70)
            ok, n = proceso_access(
                rutas["mdb"],
                {"MDB": rutas["tabulado"]},
                ["MDB"],
                solo_prueba,
                log,
                lambda pct: progreso(pct, 100, "Access..."),
                fuentes=FUENTES_ENERGIA,
            )
            if not ok:
                return False, "El Access fallo; revisa el log."
            resumen.append(f"Access: {n} filas insertadas"
                           + (" (modo prueba, sin cambios)" if solo_prueba else ""))

        progreso(100, 100, "Listo")
        return True, " | ".join(resumen) if resumen else "No se hizo nada."

    except Exception as e:
        log(f"\nERROR: {e}")
        log(traceback.format_exc())
        try:
            if wb_abierto is not None:
                wb_abierto.close()
            if app is not None:
                app.quit()
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
    root.title("Actualizar Energía"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("1000x740")

    frame_btns = tk.Frame(root)
    frame_btns.pack(side="bottom", fill="x", pady=8)

    canvas = tk.Canvas(root, borderwidth=0, highlightthickness=0)
    scroll = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    contenedor = tk.Frame(canvas)
    canvas_win = canvas.create_window((0, 0), window=contenedor, anchor="nw")

    def _ajustar(_e=None):
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
                 text="Las rutas las resolvió el Revisor; este script no las vuelve a buscar.\n"
                      "Marca qué quieres actualizar: los casilleros arrancan en blanco.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    tk.Label(contenedor, text="Actualizar los entregables de Sobrecostos de Energía",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(contenedor,
             text=f"Origen: 02 Consolidado_Tabulado, hoja «{HOJA_ORIGEN}», "
                  f"solo los tipos {' y '.join(TIPOS)}.\n"
                  "El tabulado se abre solo en lectura: no se modifica ni se guarda.",
             font=("Segoe UI", 8), fg="#555555", justify="center").pack(pady=(0, 8))

    # ---- variables ----
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_tab = tk.StringVar(value="[pendiente]")
    var_mdb = tk.StringVar(value="[pendiente]")
    var_cons = tk.StringVar(value="[pendiente]")
    # En blanco si viene del Revisor; a mano se recuerda la ultima seleccion.
    sel_mdb = tk.BooleanVar(value=False if modo["traspaso"]
                            else bool(cfg.get("ene_sel_mdb", True)))
    sel_cons = tk.BooleanVar(value=False if modo["traspaso"]
                             else bool(cfg.get("ene_sel_cons", True)))
    var_prueba = tk.BooleanVar(value=False)
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    labels = {}

    def fila_ruta(parent, titulo, var, clave):
        fr = tk.LabelFrame(parent, text=titulo, padx=10, pady=6)
        fr.pack(fill="x", padx=20, pady=4)
        lbl = tk.Label(fr, textvariable=var, wraplength=880, justify="left",
                       cursor="hand2", font=("Segoe UI", 9), anchor="w")
        lbl.pack(fill="x")
        lbl.bind("<Button-1>", lambda e: abrir_en_explorador(var.get(), es_archivo=True))
        labels[clave] = lbl
        return fr

    fila_ruta(contenedor, "Origen — 02 Consolidado_Tabulado", var_tab, "tab")

    fr_mdb = fila_ruta(contenedor, "Destino 1 — 03b ENTRADA_SOB_AAMM_*.mdb", var_mdb, "mdb")
    tk.Checkbutton(fr_mdb, text="Actualizar el Access (reemplaza SCMT y SCPC; "
                                "los otros tipos quedan intactos)",
                   variable=sel_mdb, font=("Segoe UI", 9)).pack(anchor="w")

    fr_cons = fila_ruta(contenedor, "Destino 2 — Consolidado_AAMM_*.xlsm", var_cons, "cons")
    tk.Checkbutton(fr_cons, text=f"Actualizar el Consolidado (hoja «{HOJA_DESTINO_CONSOLIDADO}», "
                                 "A:J desde la fila 2)",
                   variable=sel_cons, font=("Segoe UI", 9)).pack(anchor="w")
    tk.Label(fr_cons, text="Debe estar CERRADO en Excel antes de ejecutar.",
             font=("Segoe UI", 8), fg="#a00000").pack(anchor="w")

    # ---- selector de carpeta ----
    fr_c = tk.LabelFrame(contenedor, text="Carpeta 02 CASO RELIQUIDACION", padx=10, pady=6)
    fr_c.pack(fill="x", padx=20, pady=4)
    lbl_c = tk.Label(fr_c, textvariable=var_carpeta, wraplength=880, justify="left",
                     cursor="hand2", font=("Segoe UI", 9), anchor="w")
    lbl_c.pack(fill="x")
    lbl_c.bind("<Button-1>", lambda e: abrir_en_explorador(var_carpeta.get()))
    labels["carpeta"] = lbl_c

    tk.Checkbutton(contenedor,
                   text="Modo prueba (solo simula el Access; no escribe nada)",
                   variable=var_prueba, font=("Segoe UI", 9)).pack(anchor="w", padx=24, pady=(4, 0))

    # ---- progreso ----
    fr_p = tk.Frame(contenedor)
    fr_p.pack(fill="x", padx=20, pady=(6, 0))
    tk.Label(fr_p, textvariable=var_estado, font=("Segoe UI", 9),
             anchor="w").pack(side="left", fill="x", expand=True)
    tk.Label(fr_p, textvariable=var_tiempo, font=("Consolas", 10, "bold"),
             fg="#2d7a2d").pack(side="right", padx=8)
    barra = ttk.Progressbar(contenedor, mode="determinate", length=400)
    barra.pack(fill="x", padx=20, pady=(2, 4))

    fr_log = tk.LabelFrame(contenedor, text="Progreso detallado", padx=6, pady=4)
    fr_log.pack(fill="both", expand=True, padx=20, pady=4)
    txt_log = tk.Text(fr_log, height=14, font=("Consolas", 9), wrap="none")
    sb = tk.Scrollbar(fr_log, command=txt_log.yview)
    txt_log.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt_log.pack(fill="both", expand=True)

    cola = queue.Queue()

    def log(msg):
        cola.put(("log", str(msg)))

    def progreso(actual, total=100, mensaje=None):
        cola.put(("prog", (actual, total, mensaje)))

    def bombear():
        try:
            for _ in range(200):
                tipo, val = cola.get_nowait()
                if tipo == "log":
                    txt_log.insert("end", val + "\n")
                    txt_log.see("end")
                elif tipo == "prog":
                    actual, total, mensaje = val
                    barra["maximum"] = total or 100
                    barra["value"] = actual
                    if mensaje:
                        var_estado.set(mensaje)
        except queue.Empty:
            pass
        root.after(120, bombear)

    timer = {"on": False, "t0": 0.0}

    def tick():
        if timer["on"]:
            var_tiempo.set(fmt_tiempo(time.time() - timer["t0"]))
            root.after(500, tick)

    # ---- deteccion de rutas ----
    def pintar():
        actualizar_color_label(labels["tab"], var_tab.get(), es_archivo=True)
        actualizar_color_label(labels["mdb"], var_mdb.get(), es_archivo=True)
        actualizar_color_label(labels["cons"], var_cons.get(), es_archivo=True)
        actualizar_color_label(labels["carpeta"], var_carpeta.get())

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            var_tab.set(d.get("consolidado_tabulado") or "[el Revisor no mandó consolidado_tabulado]")
            var_mdb.set(d.get("mdb_sob") or "[el Revisor no mandó mdb_sob]")
            var_cons.set(d.get("consolidado_energia") or "[el Revisor no mandó consolidado_energia]")
            pintar()
            return
        ruta = var_carpeta.get()
        if not ruta or not Path(ruta).is_dir():
            pintar()
            return
        base = Path(ruta)
        t = buscar_tabulado(base)
        m = buscar_mdb_energia(base)
        c = buscar_consolidado_energia(base)
        var_tab.set(str(t) if t else "[02 Consolidado_Tabulado no encontrado]")
        var_mdb.set(str(m) if m else "[03b ENTRADA_SOB no encontrado]")
        var_cons.set(str(c) if c else "[Consolidado_AAMM no encontrado]")
        pintar()

    def sel_carpeta():
        ini = var_carpeta.get()
        ini = ini if ini and Path(ini).is_dir() else ""
        r = filedialog.askdirectory(title="Selecciona 02 CASO RELIQUIDACION", initialdir=ini)
        if r:
            var_carpeta.set(r)
            guardar_config({"carpeta_reliq": r})
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Carpeta elegida a mano: se dejan de usar las rutas del Revisor.")
            refrescar()

    tk.Button(fr_c, text="Examinar", command=sel_carpeta).pack(anchor="w", pady=(4, 0))

    # ---- ejecutar ----
    def lanzar():
        hacer_mdb, hacer_cons = sel_mdb.get(), sel_cons.get()
        if not hacer_mdb and not hacer_cons:
            messagebox.showwarning("Sin selección",
                                   "Marca al menos uno de los dos destinos.")
            return

        rutas = {"tabulado": var_tab.get(),
                 "mdb": var_mdb.get(),
                 "consolidado": var_cons.get()}

        def falta(clave, etiqueta):
            v = rutas[clave]
            if not v or v.startswith("[") or not Path(v).is_file():
                messagebox.showerror("Archivo faltante",
                                     f"No se encontró {etiqueta}.\n\n{v}")
                return True
            return False

        if falta("tabulado", "el 02 Consolidado_Tabulado (origen)"):
            return
        if hacer_mdb and falta("mdb", "el 03b ENTRADA_SOB (.mdb)"):
            return
        if hacer_cons and falta("consolidado", "el Consolidado_AAMM"):
            return

        # El Consolidado se abre para escribir: tiene que estar cerrado.
        if hacer_cons:
            d = Path(rutas["consolidado"])
            try:
                with open(d, "r+b"):
                    pass
            except PermissionError:
                messagebox.showwarning(
                    "El archivo está en uso",
                    f"No se puede escribir en:\n\n{d.name}\n\n"
                    "Ciérralo en Excel y volvé a intentar.")
                return
            except OSError:
                pass

        if not modo["traspaso"]:
            guardar_config({"ene_sel_mdb": hacer_mdb, "ene_sel_cons": hacer_cons})

        if not var_prueba.get():
            que = []
            if hacer_cons:
                que.append(f"  • Consolidado: se REEMPLAZA A2:J de la hoja "
                           f"«{HOJA_DESTINO_CONSOLIDADO}»")
            if hacer_mdb:
                que.append(f"  • Access: se REEMPLAZAN los tipos {', '.join(TIPOS)}")
            if not messagebox.askyesno(
                    "Confirmar",
                    "Se va a hacer:\n\n" + "\n".join(que) +
                    "\n\n(Se recomienda tener respaldo del .mdb)"):
                return

        txt_log.delete("1.0", "end")
        log(f"Origen : {rutas['tabulado']}")
        if hacer_cons:
            log(f"Destino: {rutas['consolidado']}")
        if hacer_mdb:
            log(f"Destino: {rutas['mdb']}")
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
            ok, msg = ejecutar(rutas, hacer_mdb, hacer_cons,
                               var_prueba.get(), log, progreso)

            def fin():
                timer["on"] = False
                btn.config(state="normal", bg="#2d7a2d")
                if ok:
                    var_estado.set(f"Listo — {msg}")
                    messagebox.showinfo("Listo", f"Proceso terminado.\n\n{msg}")
                else:
                    var_estado.set("Terminó con errores — revisa el log")
                    messagebox.showerror("Error", f"El proceso falló.\n\n{msg}")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text="ACTUALIZAR ENERGÍA", bg="#2d7a2d", fg="white",
                    font=("Segoe UI", 10, "bold"), command=lanzar)
    btn.pack(side="left", padx=8, expand=True)
    tk.Button(frame_btns, text="Refrescar rutas", command=refrescar).pack(side="left", padx=8)
    tk.Button(frame_btns, text="Salir", command=root.destroy).pack(side="left", padx=8)

    refrescar()
    bombear()

    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")
    root.mainloop()


if __name__ == "__main__":
    main()
