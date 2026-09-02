# =============================================================================
#  Actualiza el Access de la planilla 9
#  (Ocupar_este_para_Reliquidacion_AAMM_*.mdb)
# =============================================================================
#  Reemplaza a "para ricardo.py" + archivo_de_configuracion.yaml.
#
#  Actualiza DOS tablas:
#
#    Sobrecostos      (Clave Año_Mes, Tipo_sobrecosto, Central, Hora Mensual,
#                      Sobrecosto), desde cuatro bloques de las planillas 3, 5 y 6
#    Central_Empresa  (Central, Empresa), desde tres hojas de propietarios
#
#  DE DONDE SALE CADA DATO  —  replicado del script original
#  ----------------------------------------------------------------------------
#  Sobrecostos (4 bloques de 5 columnas cada uno):
#
#   | Fuente   | Archivo | Hoja                        | Cols    | Datos desde |
#   |----------|---------|-----------------------------|---------|-------------|
#   | SUBASTAS | 3_      | DB                          | BC:BG   | fila 3      |
#   | CRA      | 5_      | CÁLCULO_CRA                 | AQ:AU   | fila 9      |
#   | CO_CT    | 6_      | ENERGIA_Y_CALCULO_CO_ERNC   | AY:BC   | fila 9      |
#   | REA      | 6_      | CALCULO_REA_CENTRAL         | AT:AX   | fila 10     |
#
#  Las filas de datos salen del "header=" del original: header=1 -> datos en 3,
#  header=7 -> datos en 9, header=8 -> datos en 10.
#
#  El bloque BESS del script viejo (planilla 11, PRORRATA_RETIROS, IV:IZ) NO se
#  incluye: incluir_bess estaba en False y no se usa.
#
#  Central_Empresa (2 columnas: Central, Empresa). Se actualiza SIEMPRE junto con
#  su planilla: sale del mismo archivo, no tiene sentido separarlo.
#
#   | Fuente | Archivo | Hoja              | Central | Empresa | Desde |
#   |--------|---------|-------------------|---------|---------|-------|
#   | P3     | 3_      | DB                | M       | L       | 3     |
#   | P5     | 5_      | EMPRESAS          | B       | C       | 9     |
#   | P6     | 6_      | CONSUMOS_PROPIOS  | B       | H       | 9     |
#
#  En P3 salen de la MISMA hoja DB de la que salen los sobrecostos: K, L y M son
#  configuracion, propietario y unidad infotecnica. La "central" que se lleva al
#  Access es la unidad infotecnica (M), asi que el propietario tiene que salir de
#  esa misma tabla o los nombres no calzan con la columna Central del bloque.
#
#  En P6 las dos columnas NO son contiguas, y el largo lo manda la B: se corta en
#  la primera celda vacia de B, aunque la H siga con datos. El motivo es que en
#  esa hoja HAY centrales sin propietario (la H vacia o en 0), y esas filas son
#  validas y hay que conservarlas: si se cortara por la H se perderian.
#
#  Una central sin dueño no es un error por si misma. Lo que SI es un error es que
#  una central con MONTO no tenga dueño, y eso lo comprueba el Revisor cruzando
#  las tablas Sobrecostos y Central_Empresa.
#
#  QUE SE CAMBIO RESPECTO DEL ORIGINAL, Y POR QUE
#  ----------------------------------------------------------------------------
#  1. Se leen los bloques POR POSICION, no por nombre de columna. El original
#     hacia pd.concat de los cuatro dataframes, que alinea por NOMBRE de columna:
#     si una hoja tenia el encabezado escrito distinto ("Pago " con espacio, por
#     ejemplo), aparecian columnas extra llenas de NaN y el INSERT de 5
#     parametros se desalineaba o fallaba. Leyendo por posicion eso no puede
#     pasar. Igual se leen los encabezados y se avisa si no son los esperados.
#  2. Los filtros "Pago != 0" y "Central != 0" pasan a ser por posicion (columna
#     5 y columna 3 del bloque), que es lo mismo pero sin depender del nombre.
#  3. Se reutiliza el motor de Actualiza_Data_Access.py (conexion, tipos,
#     verificacion, rollback) en vez de repetir la parte de pyodbc.
#  4. El original usaba autocommit=True: si el INSERT fallaba despues del DELETE,
#     la tabla quedaba vacia o a medias sin aviso. El motor que se reutiliza
#     trabaja en una transaccion y revierte si algo falla.
#  5. El original escribia un df_ricardo_salida.xlsx de paso. Aca es opcional.
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
CONFIG_PATH = DIR_SCRIPT / "config.json"


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


# --- motor de Access, reutilizado ------------------------------------------
_AYUDA = ("Los dos archivos tienen que estar en la misma carpeta y ser de la\n"
          "misma versión. Copia de nuevo Actualiza_Data_Access.py junto a este\n"
          f"script.\n\nCarpeta actual: {DIR_SCRIPT}")
try:
    import Actualiza_Data_Access as _ADA
    from Actualiza_Data_Access import (
        proceso as proceso_access,
        normalizar, buscar_carpeta, buscar_archivo, buscar_hoja,
        col_letra, col_letra_a_num, ultima_fila, leer_bloque, fmt_tiempo, fmt_int,
        leer_columnas_rapido, es_zip_excel,
        es_vacio, es_cero, actualizar_color_label,
        driver_access, conectar_access, columnas_tabla,
    )
except ImportError as e:
    _morir("Falta Actualiza_Data_Access.py",
           f"No se pudo cargar Actualiza_Data_Access.py.\n\n{_AYUDA}\n\nDetalle: {e}")

# Las cuatro hacen falta: borrar_todo para vaciar la tabla, cols_no_cero para el
# filtro de Central != 0, y las otras dos para pasar fuentes propias. Sin
# cols_no_cero NO falla nada: simplemente entran las filas con Central en 0, en
# silencio. Por eso se exige.
_NECESITA = {"fuentes_externas", "filtro_por_valores", "borrar_todo",
             "cols_no_cero", "forzar_valores"}
_TIENE = set(getattr(_ADA, "CAPACIDADES", ()))
if not _NECESITA <= _TIENE:
    _morir("Actualiza_Data_Access.py está desactualizado",
           "El Actualiza_Data_Access.py que hay al lado es una versión ANTIGUA: "
           f"le falta {', '.join(sorted(_NECESITA - _TIENE))}.\n\n{_AYUDA}")


# =============================================================================
#  CONFIGURACION
# =============================================================================
TABLA_SOB = "Sobrecostos"
TABLA_CE = "Central_Empresa"

# Indices DENTRO del bloque de 5 columnas (0-based), segun el orden de la tabla:
#   0 Clave Año_Mes | 1 Tipo_sobrecosto | 2 Central | 3 Hora Mensual | 4 Sobrecosto
IDX_CLAVE = 0
IDX_CENTRAL = 2
IDX_MONTO = 4

# La Clave Año_Mes viene MAL desde el origen: siempre trae 23xx aunque el mes sea
# otro (2405 llega como 2305). No se lee esa columna: se pisa con el mes sacado
# del NOMBRE de los archivos, que es el dato confiable.
RE_AAMM = re.compile(r"[_\s](\d{4})[_\s]*[Rr]\d")


def aamm_de_nombre(ruta):
    """El AAMM del nombre de un archivo: '..._2502_R01P.xlsm' -> 2502.
    None si no se puede sacar."""
    m = RE_AAMM.search(Path(ruta).stem)
    if not m:
        return None
    v = int(m.group(1))
    # Un AAMM valido tiene mes entre 1 y 12. Asi no se confunde con un numero
    # cualquiera de cuatro cifras que ande por el nombre.
    if 1 <= v % 100 <= 12:
        return v
    return None


def detectar_aamm(rutas, log):
    """El AAMM comun a los archivos. Si no coinciden entre si, lo dice: son
    archivos de meses distintos y eso es un problema en si mismo."""
    vistos = {}
    for k, r in rutas.items():
        if k == "mdb":
            continue
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

# Encabezados que se esperan, solo para avisar si el archivo cambio.
ENCABEZADOS_ESPERADOS = ("clave", "tipo", "central", "hora", "pago")

# Todo se organiza POR PLANILLA: una casilla por planilla, y al marcarla se
# actualizan sus bloques de Sobrecostos Y sus propietarios. No tiene sentido
# separarlos: los propietarios salen del mismo archivo.
#
# La planilla 6 aporta DOS bloques de Sobrecostos (CO_ERNC y REA).
PLANILLAS = {
    "p3": {
        "etiqueta": "Planilla 3 — 3_REMUNERACIÓN_SUBASTAS_E_ID",
        "patron": r"^3_remuneracion_subastas",
        "sobrecostos": [
            {"nombre": "Subastas", "hoja": "DB", "fila_ini": 3,      # header=1
             "bloques": [("BC", "BG")], "filtrar_ceros": True,
             "cols_no_cero": [IDX_MONTO, IDX_CENTRAL]},
        ],
        # Los propietarios salen de la MISMA hoja DB, no de una hoja aparte:
        #     K = configuracion   L = propietario   M = unidad infotecnica
        # La "central" que se lleva al Access es la UNIDAD INFOTECNICA (M), asi
        # que el propietario tiene que salir de esta tabla o los nombres no
        # calzan con los de la columna Central del bloque de sobrecostos.
        # Ojo: la central es la M y la empresa la L, o sea que aca la empresa
        # esta ANTES que la central.
        "propietarios": {"hoja": "DB", "col_central": "M",
                         "col_empresa": "L", "fila_ini": 3},
    },
    "p5": {
        "etiqueta": "Planilla 5 — 5_REMUNERACIÓN_CRA",
        "patron": r"^5_remuneracion_cra",
        "sobrecostos": [
            {"nombre": "CRA", "hoja": "CÁLCULO_CRA", "fila_ini": 9,  # header=7
             "bloques": [("AQ", "AU")], "filtrar_ceros": True,
             "cols_no_cero": [IDX_MONTO, IDX_CENTRAL]},
        ],
        "propietarios": {"hoja": "EMPRESAS", "col_central": "B",
                         "col_empresa": "C", "fila_ini": 9},
    },
    "p6": {
        "etiqueta": "Planilla 6 — 6_REMUNERACIÓN_REA_Y_CO_ERNC",
        "patron": r"^6_remuneracion_rea",
        "sobrecostos": [
            {"nombre": "CO_ERNC", "hoja": "ENERGIA_Y_CALCULO_CO_ERNC",
             "fila_ini": 9,                                          # header=7
             "bloques": [("AY", "BC")], "filtrar_ceros": True,
             "cols_no_cero": [IDX_MONTO, IDX_CENTRAL]},
            {"nombre": "REA", "hoja": "CALCULO_REA_CENTRAL",
             "fila_ini": 10,                                         # header=8
             "bloques": [("AT", "AX")], "filtrar_ceros": True,
             "cols_no_cero": [IDX_MONTO, IDX_CENTRAL]},
        ],
        # En CONSUMOS_PROPIOS hay centrales SIN propietario (empresa 0 o vacia).
        # Son datos validos y hay que conservarlos: por eso el corte es la primera
        # celda vacia de la CENTRAL, no de la empresa.
        "propietarios": {"hoja": "CONSUMOS_PROPIOS", "col_central": "B",
                         "col_empresa": "H", "fila_ini": 9},
    },
}
ORDEN_PL = ["p3", "p5", "p6"]

CARPETA_P9 = "04 Planilla 9"


# =============================================================================
#  CONFIG COMPARTIDO
# =============================================================================
def get_usuario():
    try:
        u = os.environ.get("USERNAME") or os.environ.get("USER") or "desconocido"
        return f"{socket.gethostname()}_{u}"
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
    tmp = ruta.with_suffix(ruta.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, ruta)


def _modificar_config(mutador):
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
#  LECTURA DE LOS PROPIETARIOS
# =============================================================================
def leer_propietarios(ruta, cfg, log):
    """(Central, Empresa) de una hoja de propietarios, SIN abrir Excel.

    El largo lo manda la columna de la CENTRAL: se corta en su primera celda
    vacia aunque la de empresa siga con datos. En CONSUMOS_PROPIOS las dos
    columnas no son contiguas (B y H) y HAY centrales sin propietario, que se
    conservan con la empresa en None.
    """
    c_cen = cfg["col_central"].upper()
    c_emp = cfg["col_empresa"].upper()
    f_ini = int(cfg["fila_ini"])
    datos = leer_columnas_rapido(ruta, cfg["hoja"], [c_cen, c_emp], f_ini, log)
    if datos is None:
        raise RuntimeError(f"No se pudo leer la hoja '{cfg['hoja']}' de "
                           f"{Path(ruta).name}")
    cens = datos.get(c_cen, {})
    emps = datos.get(c_emp, {})
    if not cens:
        log(f"    ADVERTENCIA: no hay centrales en {c_cen}{f_ini} hacia abajo.")
        return []

    # Corte en la PRIMERA fila sin central. Se avisa si mas abajo quedo algo,
    # porque un renglon vacio de mas truncaria todo en silencio.
    ult_con_datos = max(cens)
    f_fin = ult_con_datos
    for f in range(f_ini, ult_con_datos + 1):
        if es_vacio(cens.get(f)):
            f_fin = f - 1
            break
    if f_fin < f_ini:
        log(f"    ADVERTENCIA: {c_cen}{f_ini} está vacía: no se leyó nada.")
        return []
    if f_fin < ult_con_datos:
        quedaron = sum(1 for f in range(f_fin + 1, ult_con_datos + 1)
                       if not es_vacio(cens.get(f)))
        log(f"    OJO: se cortó en {c_cen}{f_fin + 1} (vacía), pero más abajo hay "
            f"{quedaron} celda(s) con central hasta la fila {ult_con_datos}.")
        log(f"         Esas NO se leyeron. Si tienen que entrar, hay que quitar "
            f"el renglón vacío.")

    salida, sin_empresa, vacias = [], 0, 0
    for f in range(f_ini, f_fin + 1):
        cen = cens.get(f)
        emp = emps.get(f)
        if es_vacio(cen) or es_cero(cen):
            vacias += 1
            continue
        # Una central SIN propietario se CONSERVA: es un dato valido. Que eso sea
        # un problema depende de si tiene plata, y eso lo comprueba el Revisor
        # cruzando Sobrecostos con Central_Empresa.
        if es_vacio(emp) or es_cero(emp):
            sin_empresa += 1
            salida.append([str(cen).strip(), None])
        else:
            salida.append([str(cen).strip(), str(emp).strip()])
    log(f"    Filas {f_ini} a {f_fin}: {len(salida)} par(es) Central-Empresa"
        + (f", {vacias} sin central" if vacias else "")
        + (f", {sin_empresa} SIN PROPIETARIO" if sin_empresa else ""))
    return salida


def revisar_encabezados(ruta, cfg, log):
    """Lee la fila de encabezado del bloque y avisa si no se parece a lo
    esperado. No corta el proceso: los datos se leen por POSICION, asi que un
    encabezado distinto no rompe nada, pero conviene saberlo."""
    f_enc = int(cfg["fila_ini"]) - 1
    a, b = cfg["bloques"][0]
    cols = [col_letra(n) for n in range(col_letra_a_num(a), col_letra_a_num(b) + 1)]
    try:
        datos = leer_columnas_rapido(ruta, cfg["hoja"], cols, f_enc, log)
    except Exception as e:
        log(f"    (no se pudo leer el encabezado: {e})")
        return None
    if datos is None:
        return None
    enc = [datos.get(c, {}).get(f_enc) for c in cols]
    enc = ["" if v is None else str(v) for v in enc]
    log(f"    Encabezado fila {f_enc}: {enc}")
    norm = [normalizar(x) for x in enc]
    raros = [i for i, esperado in enumerate(ENCABEZADOS_ESPERADOS)
             if i < len(norm) and esperado not in norm[i]]
    if raros:
        log(f"    OJO: el encabezado no se parece a "
            f"{list(ENCABEZADOS_ESPERADOS)}.")
        log(f"         Los datos se leen por POSICIÓN, así que igual entran en el "
            f"orden Clave, Tipo, Central, Hora, Monto.")
        log(f"         Pero si las columnas del archivo cambiaron de lugar, esto "
            f"es la señal.")
    return enc


# =============================================================================
#  CARGA DE Central_Empresa
# =============================================================================
def cargar_central_empresa(ruta_mdb, filas, solo_lectura, log):
    """Vacia Central_Empresa y carga UNA fila por central.

    filas: lista de (central, empresa, origen). El origen es solo para poder
    decir de que planilla vino cada una cuando hay conflicto.
    Devuelve (ok, n_insertadas, n_sin_dueno).

    Transaccion unica: si algo falla, se revierte y la tabla queda como estaba.
    Las centrales SIN propietario se cargan igual, con la empresa en NULL: son
    datos validos. El Revisor comprueba despues si alguna de esas tiene plata.
    """
    # conectar_access ya comprueba el driver y devuelve (conexion, driver):
    # hay que desempaquetar los dos, o cn queda siendo la tupla y cn.cursor()
    # revienta.
    cn, _drv = conectar_access(ruta_mdb)
    try:
        cn.autocommit = False
        cur = cn.cursor()
        cols = columnas_tabla_de(cur, TABLA_CE)
        log(f"    Columnas de [{TABLA_CE}]: {cols}")
        cc = _buscar_col(cols, "central")
        ce = _buscar_col(cols, "empresa")
        if cc is None or ce is None:
            raise RuntimeError(
                f"No se encontraron las columnas Central y Empresa en "
                f"[{TABLA_CE}]. Hay: {cols}")
        antes = cur.execute(f"SELECT COUNT(*) FROM [{TABLA_CE}]").fetchval()
        log(f"    Filas antes: {fmt_int(antes)}")

        # UNA fila por central. La tabla tiene indice unico en Central: dos filas
        # con la misma central hacen fallar el INSERT entero con
        # "crearian valores duplicados en el indice".
        #
        # No alcanza con quitar los pares (Central, Empresa) repetidos: la misma
        # central puede venir de DOS planillas, o repetida con empresas distintas.
        # Se compara la central normalizada (sin tildes, espacios ni guiones
        # bajos), que es mas estricto que Access y evita colisiones que el indice
        # si consideraria iguales.
        #
        # Al elegir cual queda:
        #   - gana la que TIENE empresa sobre la que no la tiene
        #   - si dos traen empresas distintas, queda la primera y se avisa fuerte
        elegido = {}          # clave -> (central, empresa, origen)
        conflictos, repetidas = [], 0
        for cen, emp, origen in filas:
            k = clave_central(cen)
            if k not in elegido:
                elegido[k] = (cen, emp, origen)
                continue
            repetidas += 1
            cen0, emp0, org0 = elegido[k]
            if emp0 is None and emp is not None:
                elegido[k] = (cen, emp, origen)          # la que tiene dueño gana
            elif emp is not None and emp0 is not None \
                    and normalizar(emp) != normalizar(emp0):
                conflictos.append((cen0, emp0, org0, emp, origen))
        unicas = [(cen, emp) for cen, emp, _ in elegido.values()]
        if repetidas:
            log(f"    {repetidas} fila(s) de centrales ya vistas: se deja una "
                f"por central (la tabla no acepta repetidas)")
        if conflictos:
            log(f"    OJO: {len(conflictos)} central(es) con DOS empresas "
                f"distintas. Queda la primera:")
            for cen, e0, o0, e1, o1 in conflictos[:12]:
                log(f"         {str(cen)[:28]:<30} {o0}:{e0}   vs   {o1}:{e1}"
                    f"   -> queda {e0}")
            if len(conflictos) > 12:
                log(f"         ... y {len(conflictos) - 12} más")
            log(f"         Si alguna es un error de escritura, corregila en la "
                f"planilla de origen.")

        sin_dueno = [c for c, e in unicas if e is None]
        if sin_dueno:
            log(f"    {len(sin_dueno)} central(es) SIN propietario "
                f"(se cargan igual, con la empresa vacía):")
            for c in sin_dueno[:12]:
                log(f"         {c}")
            if len(sin_dueno) > 12:
                log(f"         ... y {len(sin_dueno) - 12} más")
            log(f"    Si alguna de esas tiene monto, el Revisor lo marca al "
                f"verificar.")

        # Una central con DOS empresas distintas si es un problema.
        por_central = {}
        for cen, emp in unicas:
            if emp is None:
                continue
            por_central.setdefault(normalizar(cen), set()).add(normalizar(emp))
        chocan = {c: e for c, e in por_central.items() if len(e) > 1}
        if chocan:
            log(f"    OJO: {len(chocan)} central(es) con MÁS DE UNA empresa:")
            for c in list(chocan)[:10]:
                log(f"         {c}: {sorted(chocan[c])}")

        if solo_lectura:
            log(f"    [MODO PRUEBA] se borrarían {fmt_int(antes)} y se "
                f"insertarían {fmt_int(len(unicas))} filas.")
            cn.rollback()
            return True, len(unicas), len(sin_dueno)

        cur.execute(f"DELETE FROM [{TABLA_CE}]")
        log(f"    Tabla vaciada.")
        sql = f"INSERT INTO [{TABLA_CE}] ([{cc}], [{ce}]) VALUES (?, ?)"
        paso = 500
        try:
            for k in range(0, len(unicas), paso):
                cur.executemany(sql, unicas[k:k + paso])
        except Exception as e:
            if "23000" in str(e) or "duplicad" in str(e).lower():
                # El indice unico rechazo algo. Ya se deja una fila por central,
                # asi que si igual choca el indice es sobre OTRA columna (o es
                # compuesto). Se busca cual para no dejar al usuario a ciegas.
                log("    El Access rechazó la carga por valores duplicados.")
                log("    Ya se manda una sola fila por central, así que el índice")
                log("    debe estar sobre otra columna. Buscando cuál...")
                for col, idx in ((cc, 0), (ce, 1)):
                    vals = [u[idx] for u in unicas if u[idx] is not None]
                    rep_ = len(vals) - len({normalizar(v) for v in vals})
                    log(f"       [{col}]: {len(vals)} valores, "
                        f"{rep_} repetido(s)")
                nulos = sum(1 for u in unicas if u[1] is None)
                if nulos > 1:
                    log(f"       {nulos} filas con la empresa VACÍA: si el índice"
                        f" no acepta nulos repetidos, es por acá.")
            raise
        despues = cur.execute(f"SELECT COUNT(*) FROM [{TABLA_CE}]").fetchval()
        if despues != len(unicas):
            raise RuntimeError(
                f"Quedaron {despues} filas y se insertaron {len(unicas)}: "
                f"no cuadra, se revierte.")
        cn.commit()
        log(f"    Filas insertadas: {fmt_int(len(unicas))}")
        return True, len(unicas), len(sin_dueno)
    except Exception:
        try:
            cn.rollback()
            log("    Se revirtió: la tabla quedó como estaba.")
        except Exception:
            pass
        raise
    finally:
        try:
            cn.close()
        except Exception:
            pass


def clave_central(t):
    """Normaliza el nombre de una central para comparar: sin tildes, sin espacios
    ni guiones bajos, en mayusculas. Asi 'El Toro-1', 'EL_TORO-1' y 'ELTORO-1'
    son la misma.

    Es MAS estricto que la comparacion de Access (que ignora mayusculas pero no
    espacios), y eso conviene: evita mandar dos filas que el indice unico
    consideraria iguales. La misma funcion esta en el Revisor.
    """
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[\s_]+", "", t).upper()


def columnas_tabla_de(cur, tabla):
    """Nombres de columna de una tabla cualquiera del Access."""
    cur.execute(f"SELECT * FROM [{tabla}] WHERE 1=0")
    return [d[0] for d in cur.description]


def _buscar_col(cols, objetivo):
    obj = normalizar(objetivo)
    for c in cols:
        if normalizar(c) == obj:
            return c
    for c in cols:
        if obj in normalizar(c):
            return c
    return None


def guardar_excel_dump(ruta_mdb, tabla, encabezados, filas, log):
    """Guarda en un Excel lo que se cargo, al lado del .mdb.

    Es el equivalente del df_ricardo_salida.xlsx del script viejo, pero con el
    nombre de la tabla y del mes, y junto al Access en vez del directorio de
    trabajo del momento."""
    try:
        from openpyxl import Workbook
    except ImportError as e:
        log(f"    (no se pudo guardar el volcado: falta openpyxl: {e})")
        return
    try:
        destino = (Path(ruta_mdb).parent /
                   f"_cargado_{tabla}_{Path(ruta_mdb).stem}.xlsx")
        wb = Workbook()
        sh = wb.active
        sh.title = tabla[:31]
        sh.append(list(encabezados))
        for f in filas:
            sh.append(list(f))
        wb.save(destino)
        log(f"    Volcado guardado: {destino.name}  ({fmt_int(len(filas))} filas)")
    except Exception as e:
        log(f"    (no se pudo guardar el volcado: {e})")


# =============================================================================
#  PROCESO
# =============================================================================
def ejecutar(rutas, seleccion, solo_lectura, guardar_dump, log, progreso,
             aamm=None):
    """seleccion: lista de planillas ('p3','p5','p6'). Por cada una se cargan sus
    bloques de Sobrecostos Y sus propietarios.

    aamm: el mes que se escribe en la Clave Año_Mes de TODAS las filas. Si no
    viene, se saca del nombre de los archivos. Nunca se usa el valor del origen,
    que viene mal (siempre 23xx).

    Devuelve (ok, resumen).
    """
    # Este script NO abre Excel: todo lo que lee lo hace con el lector rapido
    # (ZIP + XML). El unico que puede llegar a abrirlo es proceso_access, y solo
    # como respaldo si una planilla no fuera .xlsx/.xlsm.
    #
    # Antes se creaba un xw.App aca y se lo pasaba a las funciones de lectura.
    # Eso fallaba: proceso_access crea y CIERRA su propio Excel, y al cerrarlo
    # mataba el proceso COM compartido. Lo que venia despues (los propietarios)
    # reventaba con "El objeto no esta conectado al servidor".
    resumen = []
    try:
        # --- el mes: del nombre de los archivos, no del origen -------------
        if aamm:
            log(f"Clave Año_Mes: {aamm} (indicado en la ventana)")
        else:
            aamm, motivo = detectar_aamm(rutas, log)
            if not aamm:
                return False, (
                    f"No se pudo determinar el mes: {motivo}.\n\n"
                    f"La Clave Año_Mes del origen viene mal (siempre 23xx), así "
                    f"que hace falta el mes de verdad.\n"
                    f"Escribilo a mano en la ventana.")
            log(f"Clave Año_Mes: {motivo}")
        log(f"  Se escribe {aamm} en TODAS las filas, sin leer la columna del "
            f"origen.")

        # --- se arma el dict de fuentes que espera el motor ---------------
        # Una planilla puede aportar mas de un bloque (la 6 aporta dos), asi que
        # la clave del motor es "planilla/bloque".
        fuentes, archivos, orden = {}, {}, []
        for pl in seleccion:
            for b in PLANILLAS[pl]["sobrecostos"]:
                k = f"{pl}/{b['nombre']}"
                fuentes[k] = dict(b,
                                  etiqueta=f"{PLANILLAS[pl]['etiqueta']} — "
                                           f"hoja {b['hoja']}",
                                  forzar_valores={IDX_CLAVE: int(aamm)})
                archivos[k] = rutas[pl]
                orden.append(k)

        log("=" * 70)
        log(f"Tabla [{TABLA_SOB}]  —  {len(orden)} bloque(s) de "
            f"{len(seleccion)} planilla(s)")
        log("=" * 70)
        for k in orden:
            log(f"\n  [{k}]  hoja {fuentes[k]['hoja']}, "
                f"{fuentes[k]['bloques'][0][0]}:{fuentes[k]['bloques'][0][1]}, "
                f"desde la fila {fuentes[k]['fila_ini']}")
            revisar_encabezados(archivos[k], fuentes[k], log)

        progreso(15, 100, "Cargando Sobrecostos...")
        ok, n = proceso_access(
            rutas["mdb"], archivos, orden, solo_lectura, log,
            lambda pct: progreso(15 + int(pct * 0.45), 100, "Sobrecostos..."),
            fuentes=fuentes,
            borrar_todo=True,      # la tabla se arma COMPLETA desde estas fuentes
        )
        if not ok:
            return False, f"falló la carga de [{TABLA_SOB}]"
        resumen.append(f"{TABLA_SOB}: {fmt_int(n)} filas, mes {aamm}"
                       + (" (prueba)" if solo_lectura else ""))

        if guardar_dump:
            filas = []
            for k in orden:
                for f in _ADA.leer_fuente(None, archivos[k], fuentes[k],
                                          lambda *_: None):
                    filas.append([k] + list(f))
            guardar_excel_dump(rutas["mdb"], TABLA_SOB,
                               ["Fuente", "Clave Año_Mes", "Tipo_sobrecosto",
                                "Central", "Hora Mensual", "Sobrecosto"],
                               filas, log)

        # --- propietarios: siempre, junto con su planilla -----------------
        log("\n" + "=" * 70)
        log(f"Tabla [{TABLA_CE}]  —  propietarios de {len(seleccion)} planilla(s)")
        log("=" * 70)
        todas = []
        for pl in seleccion:
            cfg = PLANILLAS[pl]["propietarios"]
            log(f"\n  [{pl}]  hoja {cfg['hoja']}, central {cfg['col_central']}, "
                f"empresa {cfg['col_empresa']}, desde la fila {cfg['fila_ini']}")
            todas += [(c, e, pl) for c, e in
                      leer_propietarios(rutas[pl], cfg, log)]
        progreso(70, 100, "Cargando Central_Empresa...")
        if not todas:
            raise RuntimeError("No se leyó ningún par Central-Empresa: no se "
                               "vacía la tabla.")
        log("")
        ok, n, sin_dueno = cargar_central_empresa(rutas["mdb"], todas,
                                                  solo_lectura, log)
        if not ok:
            return False, f"falló la carga de [{TABLA_CE}]"
        resumen.append(f"{TABLA_CE}: {fmt_int(n)} filas"
                       + (" (prueba)" if solo_lectura else ""))
        if sin_dueno:
            resumen.append(f"{sin_dueno} sin dueño")
        if guardar_dump:
            guardar_excel_dump(rutas["mdb"], TABLA_CE,
                               ["Central", "Empresa", "Planilla"], todas, log)

        progreso(100, 100, "Listo")
        return True, " | ".join(resumen)

    except Exception as e:
        log(f"\nERROR: {e}")
        log(traceback.format_exc())
        return False, str(e)


# =============================================================================
#  VENTANA
# =============================================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    root = tk.Tk()
    root.title("Actualiza el Access de la planilla 9"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("1000x720")

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
                          "Marca qué planillas traer: arrancan en blanco.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    tk.Label(cont, text="Actualizar el Access de la planilla 9",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(cont, text=f"Las planillas se abren SOLO EN LECTURA. Al marcar una se "
                        f"actualizan sus sobrecostos y sus propietarios.\n"
                        f"Las dos tablas se arman completas desde las planillas "
                        f"marcadas, así que se vacían antes de cargar.",
             font=("Segoe UI", 8), fg="#555555", justify="center").pack(pady=(0, 8))

    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_mdb = tk.StringVar(value="[pendiente]")
    vars_arch = {k: tk.StringVar(value="[pendiente]") for k in ORDEN_PL}
    labels = {}
    en_blanco = modo["traspaso"]

    def fila_ruta(parent, titulo, var, clave):
        fr = tk.LabelFrame(parent, text=titulo, padx=10, pady=5)
        fr.pack(fill="x", padx=20, pady=3)
        l = tk.Label(fr, textvariable=var, wraplength=880, justify="left",
                     cursor="hand2", font=("Segoe UI", 9), anchor="w")
        l.pack(fill="x")
        l.bind("<Button-1>", lambda e: abrir_en_explorador(var.get(), es_archivo=True))
        labels[clave] = l
        return fr

    fila_ruta(cont, "Access de destino — Ocupar_este_para_*.mdb", var_mdb, "mdb")

    sel = {}
    for pl in ORDEN_PL:
        c = PLANILLAS[pl]
        fr = fila_ruta(cont, c["etiqueta"], vars_arch[pl], pl)
        sel[pl] = tk.BooleanVar(value=not en_blanco)
        hojas = ", ".join(b["hoja"] for b in c["sobrecostos"])
        pr = c["propietarios"]
        tk.Checkbutton(fr, text=f"Actualizar   (sobrecostos: {hojas}   ·   "
                                f"propietarios: {pr['hoja']} "
                                f"{pr['col_central']}/{pr['col_empresa']})",
                       variable=sel[pl], font=("Segoe UI", 9)).pack(anchor="w")

    # ---- el mes ----
    # La Clave Año_Mes del origen viene mal (siempre 23xx). Se muestra el mes
    # detectado del nombre de los archivos para poder revisarlo antes de cargar.
    fmes = tk.LabelFrame(cont, text="Clave Año_Mes  (la del origen viene mal y "
                                    "se reemplaza)", padx=10, pady=6)
    fmes.pack(fill="x", padx=20, pady=4)
    var_aamm = tk.StringVar(value="")
    fm2 = tk.Frame(fmes)
    fm2.pack(fill="x")
    tk.Label(fm2, text="Mes (AAMM):", font=("Segoe UI", 9)).pack(side="left")
    tk.Entry(fm2, textvariable=var_aamm, width=8,
             font=("Consolas", 11)).pack(side="left", padx=6)
    lbl_mes = tk.Label(fm2, text="", font=("Segoe UI", 8), fg="#555555")
    lbl_mes.pack(side="left")
    tk.Label(fmes, text="Se escribe en TODAS las filas de Sobrecostos. Sale del "
                        "nombre de los archivos; corregilo si no es el que "
                        "corresponde.",
             font=("Segoe UI", 8), fg="#555555").pack(anchor="w", pady=(4, 0))

    fo = tk.Frame(cont)
    fo.pack(fill="x", padx=24, pady=(8, 0))
    var_prueba = tk.BooleanVar(value=True)
    var_dump = tk.BooleanVar(value=False)
    tk.Checkbutton(fo, text="SOLO MIRAR: contar y decir qué haría, sin escribir "
                            "en el Access",
                   variable=var_prueba, font=("Segoe UI", 9, "bold"),
                   fg="#1a5fb4").pack(anchor="w")
    tk.Checkbutton(fo, text="Guardar además un Excel con lo que se cargó",
                   variable=var_dump, font=("Segoe UI", 9)).pack(anchor="w")

    fc = tk.LabelFrame(cont, text="Carpeta 02 CASO RELIQUIDACION", padx=10, pady=5)
    fc.pack(fill="x", padx=20, pady=3)
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
    txt = tk.Text(fl, height=14, font=("Consolas", 9), wrap="none")
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
        actualizar_color_label(labels["mdb"], var_mdb.get(), es_archivo=True)
        for k, v in vars_arch.items():
            actualizar_color_label(labels[k], v.get(), es_archivo=True)
        actualizar_color_label(labels["carpeta"], var_carpeta.get())

    def _refrescar_mes():
        """Muestra el mes sacado de los nombres. Solo si el usuario no escribió
        uno a mano."""
        rr = {pl: vars_arch[pl].get() for pl in ORDEN_PL
              if vars_arch[pl].get() and not vars_arch[pl].get().startswith("[")}
        v, motivo = detectar_aamm(rr, lambda *_: None) if rr else (None, "")
        if v:
            lbl_mes.config(text=f"  detectado: {v}", fg="#2d7a2d")
            if not var_aamm.get().strip():
                var_aamm.set(str(v))
        else:
            lbl_mes.config(text=f"  {motivo or 'sin archivos'}", fg="#a00000")

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            var_mdb.set(d.get("mdb_ocupar") or "[el Revisor no mandó mdb_ocupar]")
            for pl in ORDEN_PL:
                vars_arch[pl].set(d.get(pl) or f"[el Revisor no mandó {pl}]")
            # El Revisor ya detectó el mes: es más confiable que el nombre.
            if traspaso.get("aamm") and not var_aamm.get().strip():
                var_aamm.set(str(traspaso["aamm"]))
                lbl_mes.config(text="  lo mandó el Revisor", fg="#2d7a2d")
            else:
                _refrescar_mes()
            pintar()
            return
        base = var_carpeta.get()
        if not base or not Path(base).is_dir():
            pintar()
            return
        carp = buscar_carpeta(Path(base), CARPETA_P9)
        if carp is None:
            var_mdb.set(f"[no se encontró la carpeta {CARPETA_P9}]")
            for v in vars_arch.values():
                v.set(f"[no se encontró la carpeta {CARPETA_P9}]")
            pintar()
            return
        m = buscar_archivo(carp, PATRON_MDB, extensiones=(".mdb", ".accdb"))
        var_mdb.set(str(m) if m else "[Ocupar_este no encontrado]")
        for pl in ORDEN_PL:
            f = buscar_archivo(carp, PLANILLAS[pl]["patron"])
            vars_arch[pl].set(str(f) if f else "[no encontrado]")
        _refrescar_mes()
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
        elegidas = [pl for pl in ORDEN_PL if sel[pl].get()]
        if not elegidas:
            messagebox.showwarning("Sin selección", "Marca al menos una planilla.")
            return
        mdb = var_mdb.get()
        if not mdb or mdb.startswith("[") or not Path(mdb).is_file():
            messagebox.showerror("Falta el Access",
                                 f"No se encontró el .mdb.\n\n{mdb}")
            return
        rutas = {"mdb": mdb}
        for pl in elegidas:
            v = vars_arch[pl].get()
            if not v or v.startswith("[") or not Path(v).is_file():
                messagebox.showerror("Falta un archivo",
                                     f"No se encontró la {pl}.\n\n{v}")
                return
            rutas[pl] = v

        aamm_txt = var_aamm.get().strip()
        if not aamm_txt:
            messagebox.showerror("Falta el mes",
                                 "Hay que indicar la Clave Año_Mes (AAMM).\n\n"
                                 "La del origen viene mal y se reemplaza por "
                                 "ésta.")
            return
        try:
            aamm_val = int(aamm_txt)
        except ValueError:
            messagebox.showerror("Mes inválido",
                                 f"El mes tiene que ser un número de 4 cifras "
                                 f"(AAMM).\n\nSe escribió: {aamm_txt!r}")
            return
        if not (1 <= aamm_val % 100 <= 12) or not (1000 <= aamm_val <= 9912):
            messagebox.showerror("Mes inválido",
                                 f"{aamm_val} no parece un AAMM válido: el mes "
                                 f"tiene que estar entre 01 y 12.")
            return

        prueba = var_prueba.get()
        faltan = [pl for pl in ORDEN_PL if pl not in elegidas]
        aviso = ""
        if faltan:
            aviso = ("\n\nOJO: las dos tablas se vacían y se arman SOLO con lo "
                     "que traen las planillas marcadas.\nLo que aportaban "
                     + ", ".join(faltan) + " NO va a quedar en el Access.")
        if not prueba:
            if not messagebox.askyesno(
                    "Confirmar escritura en el Access",
                    f"{Path(mdb).name}\n\n"
                    f"  • [{TABLA_SOB}] y [{TABLA_CE}] se VACÍAN y se cargan "
                    f"desde: {', '.join(elegidas)}\n"
                    f"  • la Clave Año_Mes queda en {aamm_val} en todas las filas"
                    + aviso +
                    "\n\n(Se recomienda tener respaldo del .mdb)\n\n¿Seguir?"):
                return

        txt.delete("1.0", "end")
        log(f"Access : {mdb}")
        for pl in elegidas:
            log(f"{pl:7s}: {rutas[pl]}")
        if faltan:
            log(f"NO se marcaron: {', '.join(faltan)}. Lo que aportan no va a "
                f"quedar en el Access.")
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
            ok, msg = ejecutar(rutas, elegidas, prueba, var_dump.get(),
                               log, progreso, aamm=aamm_val)

            def fin():
                timer["on"] = False
                btn.config(state="normal", bg="#2d7a2d")
                if ok and prueba:
                    var_estado.set(f"Solo mirar — {msg}")
                    messagebox.showinfo("Solo mirar",
                                        f"No se tocó el Access.\n\n{msg}\n\n"
                                        "Si está bien, desmarcá «solo mirar».")
                elif ok:
                    var_estado.set(f"Listo — {msg}")
                    messagebox.showinfo(
                        "Listo",
                        f"Terminado.\n\n{msg}\n\nDespués corré Verificar en el "
                        "Revisor: ahí se comprueba que no haya centrales con "
                        "monto y sin dueño.")
                else:
                    var_estado.set("Terminó con errores — revisa el log")
                    messagebox.showerror("Error", f"Falló.\n\n{msg}")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text="ACTUALIZAR EL ACCESS", bg="#2d7a2d",
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
