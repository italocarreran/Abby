# =============================================================================
#  Reorganiza SC_CfrecSinFD_AAMM.csv para Coyailen  —  VERSION 15 MIN
# =============================================================================
#  Integra el reorganizar_sc_cfrec.py que mando la usuaria (2026-10-29). La
#  logica es la misma; cambia solo de donde lee y donde deja la salida:
#
#      01 Sobrecostos/SC_CfrecSinFD_AAMM.csv                 <- entrada
#      01 Sobrecostos/SC_CfrecSinFD_AAMM_para_coyailen.csv   <- salida
#
#  El mes ya no va fijo en el codigo (CLAVE): se toma el SC_CfrecSinFD_* mas
#  nuevo de la carpeta, o el que mande el Revisor.
#
#  Las columnas de la ENTRADA se toman por su encabezado (ver COLUMNAS): en
#  marzo 2026 vino una columna menos antes de "Fecha" y las letras se
#  corrieron. Mapeo (salida <- encabezado de la entrada):
#      A = Fecha
#      B = Hora
#      C = Tipo
#      D = CONFIGURACION si Embalse == 0, UNIDAD GENERADORA si Embalse == 1
#      E = Sobrecosto
#      F = UNIDAD GENERADORA
#      G = GENERACION
#      H = CONFIGURACION           ("Configuración": en embalses mantiene esta)
#      I = CV
#      J = CMg[USD/MWh]
#      K..P = el nombre de CPF(+), CPF(-), CSF(+), CTF(+), CSF(-), CTF(-) si
#             esa columna vale 1, si no vacio (ese orden, como el original)
#      Q = Dolar
#      R = Bloque_15min
#
#  Casos especiales (GEN TG+TV distinta de 0):
#      Las filas de una misma CONFIGURACION en el mismo cuarto de hora
#      (FECHA_HORA) se combinan en una sola: E = suma de Sobrecosto, G = suma
#      de GENERACION; el resto (Central, Unidad, etc.) se toma de la PRIMERA fila del
#      bloque, igual que el BUSCARV del proceso manual. Aplica tambien si viene
#      una sola fila.
#
#  La salida se escribe en un .tmp y se reemplaza: si algo falla, la anterior
#  queda como estaba. Se puede correr solo, o recibir del Revisor la ruta de un
#  JSON de traspaso.
# =============================================================================

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
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
except ImportError as e:
    _morir("Falta la carpeta __comun__/",
           "No se pudo cargar __comun__/.\n\n"
           "Tiene que estar la carpeta '__comun__' hermana de Revisor_Relq_15min.\n"
           "Baja el repositorio completo, no los .py sueltos.\n\n"
           f"Carpeta actual: {DIR_SCRIPT}\n\nDetalle: {e}")


# =============================================================================
#  CONFIGURACION
# =============================================================================
CARPETA_SOB = "01 Sobrecostos"
PREFIJO_ENTRADA = "SC_CfrecSinFD_"
SUFIJO_SALIDA = "_para_coyailen"
ENCODING_SALIDA = "utf-8-sig"   # con BOM: Excel lo abre bien (Ñ, tildes)

# Clave del JSON de traspaso con la entrada (ver CLAVES_TRASPASO del Revisor).
CLAVE_TRASPASO = "sc_cfrec"


# =============================================================================
#  CONFIG COMPARTIDO
# =============================================================================
get_usuario = _cfg.clave_equipo


# Modo 15 minutos: todo lo de esta version va dentro de la clave
# SECCION_CONFIG de config.json (ver Actualiza_Pagos_Varios.py). El valor tiene
# que ser el mismo en todos los scripts de esta version (lo comprueba
# revisor/test_seccion_config.py).
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


def leer_traspaso(argv):
    return _traspaso.leer_argumento(argv)


normalizar = _texto.suave


def fmt_tiempo(seg):
    seg = int(seg)
    return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"


# =============================================================================
#  ARCHIVOS
# =============================================================================
def es_salida(ruta):
    return Path(ruta).stem.lower().endswith(SUFIJO_SALIDA)


def buscar_entrada(carpeta_reliq):
    """El SC_CfrecSinFD_*.csv mas nuevo de 01 Sobrecostos (sin las salidas)."""
    sob = next((s for s in Path(carpeta_reliq).iterdir() if s.is_dir()
                and normalizar(s.name) == normalizar(CARPETA_SOB)), None)
    if sob is None:
        return None
    pref = PREFIJO_ENTRADA.lower()
    cand = [f for f in sob.iterdir()
            if f.is_file() and f.suffix.lower() == ".csv"
            and f.name.lower().startswith(pref) and not es_salida(f)]
    return max(cand, key=lambda f: f.stat().st_mtime) if cand else None


def ruta_salida(ruta_entrada):
    """Misma carpeta, mismo nombre + _para_coyailen."""
    e = Path(ruta_entrada)
    return e.with_name(f"{e.stem}{SUFIJO_SALIDA}.csv")


# =============================================================================
#  REORGANIZACION (la del script de la usuaria)
# =============================================================================
def leer_texto(ruta):
    """Lee el CSV en UTF-8; si no lo es, cae a Windows-1252 (ANSI)."""
    datos = Path(ruta).read_bytes()
    try:
        return datos.decode("utf-8-sig")
    except UnicodeDecodeError:
        return datos.decode("cp1252")


def reparar_texto(v):
    """Repara textos mal codificados, ej: 'DEGAÃ‘' -> 'DEGAÑ'."""
    if "Ã" in v or "Â" in v:
        try:
            return v.encode("cp1252").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            return v
    return v


def clave_col(texto):
    """Encabezado comparable: sin tildes, espacios ni guiones bajos, minusculas."""
    return _texto.clave(texto or "")


# Las columnas de la entrada se toman POR ENCABEZADO, no por letra: en marzo
# 2026 el SinFD trae una columna menos antes de "Fecha" y todo se corrio una
# letra (usuaria, 2026-10-29). Campo -> encabezado de la entrada. "clave" es
# el cuarto de hora con que se agrupan los casos especiales (la columna A del
# script original); si no esta, se usa la primera columna, como antes.
COLUMNAS = {
    "clave":      "FECHA_HORA",
    "bloque":     "Bloque_15min",
    "fecha":      "Fecha",
    "hora":       "Hora",
    "config":     "CONFIGURACION",
    "unidad":     "UNIDAD GENERADORA",
    "generacion": "GENERACION",
    "gen_tg_tv":  "GEN TG+TV",
    "cv":         "CV",
    "cmg":        "CMg[USD/MWh]",
    "embalse":    "Embalse",
    "cpf_mas":    "CPF(+)",
    "cpf_menos":  "CPF(-)",
    "csf_mas":    "CSF(+)",
    "csf_menos":  "CSF(-)",
    "ctf_mas":    "CTF(+)",
    "ctf_menos":  "CTF(-)",
    "dolar":      "Dolar",
    "sobrecosto": "Sobrecosto",
    "tipo":       "Tipo",
}
OPCIONALES = ("clave",)

# Banderas en el orden de salida K..P (asi venian en el original: CTF(+)
# antes que CSF(-)).
BANDERAS = ["cpf_mas", "cpf_menos", "csf_mas", "ctf_mas", "csf_menos", "ctf_menos"]


def ubicar_columnas(encabezado):
    """{campo: indice} segun el encabezado. Levanta ValueError con lo que falta."""
    pos = {}
    for i, h in enumerate(encabezado):
        pos.setdefault(clave_col(h), i)
    ix, faltan = {}, []
    for campo, nombre in COLUMNAS.items():
        i = pos.get(clave_col(nombre))
        if i is None:
            if campo in OPCIONALES:
                i = 0
            else:
                faltan.append(nombre)
                continue
        ix[campo] = i
    if faltan:
        raise ValueError("faltan las columnas " + ", ".join(faltan))
    return ix


def a_float(valor):
    try:
        return float(valor)
    except ValueError:
        return 0.0


def es_uno(valor):
    return a_float(valor) == 1


def es_caso_especial(fila, ix):
    """Caso especial: generacion en GEN TG+TV distinta de 0."""
    return a_float(fila[ix["gen_tg_tv"]]) != 0


def construir_encabezado(h, ix):
    return [
        h[ix["fecha"]],              # A
        h[ix["hora"]],               # B
        h[ix["tipo"]],               # C
        "Central",                   # D
        h[ix["sobrecosto"]],         # E
        h[ix["unidad"]],             # F
        h[ix["generacion"]],         # G
        "Configuración",             # H
        h[ix["cv"]],                 # I
        h[ix["cmg"]],                # J
        *[h[ix[c]] for c in BANDERAS],  # K..P
        h[ix["dolar"]],              # Q
        h[ix["bloque"]],             # R
    ]


def transformar_fila(fila, h, ix):
    d = fila[ix["config"]] if not es_uno(fila[ix["embalse"]]) else fila[ix["unidad"]]
    banderas = [h[ix[c]] if es_uno(fila[ix[c]]) else "" for c in BANDERAS]
    return [
        fila[ix["fecha"]],       # A
        fila[ix["hora"]],        # B
        fila[ix["tipo"]],        # C
        d,                       # D
        fila[ix["sobrecosto"]],  # E
        fila[ix["unidad"]],      # F
        fila[ix["generacion"]],  # G
        fila[ix["config"]],      # H  (configuracion original, tambien en embalses)
        fila[ix["cv"]],          # I
        fila[ix["cmg"]],         # J
        *banderas,               # K..P
        fila[ix["dolar"]],       # Q
        fila[ix["bloque"]],      # R
    ]


def combinar(grupo, h, ix):
    """
    Une en una sola fila las filas de una misma configuracion y cuarto de hora
    con generacion en GEN TG+TV:
      - D, F, H (salida) = de la primera fila del grupo (como el BUSCARV
                           del proceso manual: Unidad = primera unidad)
      - E (salida)     = suma del Sobrecosto
      - G (salida)     = suma de la Generacion
      - banderas K..P  = se marcan si alguna fila del grupo tiene 1
      - resto          = se toma de la primera fila del grupo
    """
    base = grupo[0]
    salida = transformar_fila(base, h, ix)
    salida[4] = repr(sum(a_float(f[ix["sobrecosto"]]) for f in grupo))   # E
    salida[6] = repr(sum(a_float(f[ix["generacion"]]) for f in grupo))   # G
    for i, c in enumerate(BANDERAS):                                    # K..P
        salida[10 + i] = h[ix[c]] if any(es_uno(f[ix[c]]) for f in grupo) else ""
    return salida


def procesar(ruta_entrada, ruta_sal=None, log=print):
    """Escribe la salida y devuelve (ruta_salida, filas_entrada, filas_salida,
    bloques_combinados). Las columnas se ubican por encabezado; si falta
    alguna levanta ValueError y no escribe nada."""
    ruta_entrada = Path(ruta_entrada)
    ruta_sal = Path(ruta_sal) if ruta_sal else ruta_salida(ruta_entrada)
    lector = csv.reader(io.StringIO(leer_texto(ruta_entrada), newline=""))
    encabezado = next(lector, None)
    if not encabezado:
        raise ValueError(f"{ruta_entrada.name} está vacío.")
    encabezado = [reparar_texto(h).strip() for h in encabezado]
    try:
        ix = ubicar_columnas(encabezado)
    except ValueError as e:
        raise ValueError(f"{ruta_entrada.name}: {e}. Encabezado: "
                         f"{', '.join(encabezado)}")

    # Se mantiene el orden original; los casos especiales quedan en la
    # posicion de su primera fila.
    orden = []                       # filas normales o claves de grupo
    grupos = {}
    filas_entrada = 0
    for fila in lector:
        if not any(c.strip() for c in fila):
            continue  # salta filas vacias
        fila += [""] * (len(encabezado) - len(fila))
        fila = [reparar_texto(c) for c in fila]
        filas_entrada += 1
        if es_caso_especial(fila, ix):
            clave = (fila[ix["clave"]], fila[ix["config"]])
            if clave not in grupos:
                grupos[clave] = []
                orden.append(clave)
            grupos[clave].append(fila)
        else:
            orden.append(fila)

    tmp = ruta_sal.with_name(ruta_sal.name + ".tmp")
    try:
        with tmp.open("w", encoding=ENCODING_SALIDA, newline="") as g:
            escritor = csv.writer(g)
            escritor.writerow(construir_encabezado(encabezado, ix))
            for item in orden:
                if isinstance(item, tuple):
                    escritor.writerow(combinar(grupos[item], encabezado, ix))
                else:
                    escritor.writerow(transformar_fila(item, encabezado, ix))
        os.replace(tmp, ruta_sal)
    finally:
        if tmp.exists():
            tmp.unlink()

    log(f"OK  {ruta_entrada.name} -> {ruta_sal.name}")
    log(f"    {filas_entrada} filas de entrada -> {len(orden)} filas de salida "
        f"({len(grupos)} bloques con generación en J combinados)")
    return ruta_sal, filas_entrada, len(orden), len(grupos)


# =============================================================================
#  VENTANA
# =============================================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    root = tk.Tk()
    root.title("Reorganiza SC_CfrecSinFD para Coyailen"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("900x560")

    if traspaso:
        fa = tk.Frame(root, bg="#fff4c2", bd=1, relief="solid")
        fa.pack(fill="x", padx=20, pady=(10, 0))
        tk.Label(fa, text=f"Mes {traspaso.get('aamm') or '?'} — enviado por el Revisor",
                 bg="#fff4c2", font=("Segoe UI", 11, "bold")).pack(pady=(6, 0))
        tk.Label(fa, text="Las rutas las resolvió el Revisor.", bg="#fff4c2",
                 font=("Segoe UI", 8), fg="#444444").pack(pady=(0, 6))

    tk.Label(root, text="SC_CfrecSinFD_AAMM.csv  →  SC_CfrecSinFD_AAMM_para_coyailen.csv",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(root, text="Lee y escribe en 01 Sobrecostos. Une los bloques con "
                        "generación en J (GEN TG+TV).",
             font=("Segoe UI", 8), fg="#555555").pack(pady=(0, 6))

    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_entrada = tk.StringVar(value="[pendiente]")
    var_salida = tk.StringVar(value="[pendiente]")
    labels = {}

    def fila(titulo, var, clave):
        fr = tk.LabelFrame(root, text=titulo, padx=10, pady=4)
        fr.pack(fill="x", padx=20, pady=3)
        l = tk.Label(fr, textvariable=var, wraplength=820, justify="left",
                     cursor="hand2", font=("Segoe UI", 9), anchor="w")
        l.pack(fill="x")
        l.bind("<Button-1>", lambda e: abrir_en_explorador(var.get(), es_archivo=True))
        labels[clave] = l

    fila("Entrada — 01 Sobrecostos/SC_CfrecSinFD_AAMM.csv", var_entrada, "entrada")
    fila("Salida — 01 Sobrecostos/SC_CfrecSinFD_AAMM_para_coyailen.csv  "
         "(se crea o se reemplaza)", var_salida, "salida")

    fc = tk.LabelFrame(root, text="Carpeta 02 CASO RELIQUIDACION", padx=10, pady=4)
    fc.pack(fill="x", padx=20, pady=3)
    lc = tk.Label(fc, textvariable=var_carpeta, wraplength=820, justify="left",
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
    # Sin barra de progreso (usuaria, 2026-10-29): queda solo el tiempo.

    frame_btns = tk.Frame(root)
    frame_btns.pack(side="bottom", fill="x", pady=8)
    fl = tk.LabelFrame(root, text="Progreso detallado", padx=6, pady=4)
    fl.pack(fill="both", expand=True, padx=20, pady=4)
    txt = tk.Text(fl, height=10, font=("Consolas", 9), wrap="none")
    sb = tk.Scrollbar(fl, command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt.pack(fill="both", expand=True)

    cola = queue.Queue()
    log = lambda m="": cola.put(str(m))

    def bombear():
        try:
            for _ in range(200):
                txt.insert("end", cola.get_nowait() + "\n")
                txt.see("end")
        except queue.Empty:
            pass
        root.after(120, bombear)

    timer = {"on": False, "t0": 0.0}

    def tick():
        if timer["on"]:
            var_tiempo.set(fmt_tiempo(time.time() - timer["t0"]))
            root.after(500, tick)

    def pintar():
        for clave, var in (("entrada", var_entrada), ("salida", var_salida)):
            v = var.get()
            # La salida sale en rojo si todavia no existe: se crea al generar.
            ok = not v.startswith("[") and Path(v).is_file()
            labels[clave].config(fg="blue" if ok else "red")

    def refrescar(*_):
        e = None
        if modo["traspaso"]:
            r = traspaso.get("rutas", {}).get(CLAVE_TRASPASO)
            e = Path(r) if r else None
            if e is None:
                var_entrada.set("[el Revisor no mandó SC_CfrecSinFD_*.csv]")
        else:
            base = var_carpeta.get()
            if base and Path(base).is_dir():
                e = buscar_entrada(base)
                if e is None:
                    var_entrada.set("[SC_CfrecSinFD_*.csv no encontrado en 01 Sobrecostos]")
            else:
                var_entrada.set("[elige la carpeta 02 CASO RELIQUIDACION]")
        if e is not None:
            var_entrada.set(str(e))
            var_salida.set(str(ruta_salida(e)))
        else:
            var_salida.set("[falta la entrada]")
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
        entrada = var_entrada.get()
        if entrada.startswith("[") or not Path(entrada).is_file():
            messagebox.showerror("Falta la entrada",
                                 "No se encontró SC_CfrecSinFD_*.csv en 01 Sobrecostos.")
            return
        salida = var_salida.get()

        txt.delete("1.0", "end")
        log(f"Entrada: {entrada}")
        log(f"Salida : {salida}")
        log("-" * 70)
        btn.config(state="disabled", bg="#aaaaaa")
        var_estado.set("Reorganizando...")
        timer["on"] = True
        timer["t0"] = time.time()
        tick()

        def trabajo():
            try:
                _, n_in, n_out, n_gr = procesar(entrada, salida, log)
                ok, msg = True, (f"{n_in} filas de entrada -> {n_out} de salida "
                                 f"({n_gr} bloques combinados)")
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
                    messagebox.showinfo("Listo", f"Salida generada.\n\n{msg}\n\n{salida}")
                else:
                    var_estado.set("Terminó con errores — revisa el log")
                    messagebox.showerror("Error", f"Falló.\n\n{msg}")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text="GENERAR para Coyailen", bg="#2d7a2d",
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
