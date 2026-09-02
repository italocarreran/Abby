# =============================================================================
#  Carga Retiros_h.parquet a SQL Server
# =============================================================================
#  Version con ventana del script original. Lo que hace es lo mismo:
#      1. lee el parquet
#      2. BORRA del servidor los periodos que trae el parquet (o la tabla entera)
#      3. carga por trozos
#      4. verifica que la cuenta cuadre
#
#  ESTO ESCRIBE EN UNA BASE DE DATOS Y BORRA FILAS. Por eso:
#    - antes de borrar muestra exactamente que va a borrar y pide confirmacion
#    - hay un modo SOLO MIRAR que hace los conteos y no toca nada
#    - todo queda en el log con las cuentas antes y despues
#
#  SI FALLA A MITAD DE LA CARGA: el borrado ya se hizo y quedan filas a medias.
#  No es grave: volver a correrlo borra ese periodo otra vez y recarga. Lo que NO
#  hay que hacer es dejarlo asi, porque la tabla queda con datos incompletos.
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

# ---------------------------------------------------------------------------
#  Configuracion
# ---------------------------------------------------------------------------
NOMBRE_PARQUET = "Retiros_h.parquet"
CARPETA_PARQUET = "04 Planilla 9"      # donde lo busca dentro del caso

TABLA = "Retiros"
SERVER = "SRV-DTE"
DRIVER = "ODBC Driver 17 for SQL Server"

# Las bases entre las que se puede elegir. La primera es la de siempre.
# Los nombres van EXACTOS como estan en SQL Server. Ojo: es RELIQUIDACION sin
# tilde. Con tilde el servidor no encuentra la base y el error que devuelve es
# "Login failed for user", que despista: parece un problema de permisos y en
# realidad es el nombre mal escrito.
BASES = ["02_RETIROS", "14_RETIROS_RELIQUIDACION"]

CHUNK = 50_000          # 10k es chico para fast_executemany

# Largo de las columnas de texto SI HAY QUE CREAR la tabla.
#
# Importa mucho: to_sql, sin decirle nada, crea el texto como VARCHAR(MAX), que
# NO SE PUEDE INDEXAR. Y como en la base de sobrecostos las columnas del join
# son NVARCHAR, un VARCHAR obliga ademas a convertir fila por fila, lo que anula
# el indice aunque lo hubiera.
#
# Asi quedo la tabla de 14_RETIROS_RELIQUIDACION: 11,6 millones de filas, sin
# indice y con cinco columnas en varchar(MAX). La prorrata contra esa base
# tardaba 20-50 min contra 5-15 de la otra.
#
# Con if_exists="append" esto SOLO aplica si la tabla no existe: si ya existe se
# respetan sus tipos, y hay que arreglarla en SQL Server
# (ver arreglar_tabla_retiros.sql).
LARGO_TEXTO = 255

# Los nombres de columna NO se escriben igual en todos lados: el parquet a veces
# trae "Clave Año_Mes" (con espacio y ñ) y a veces "Clave_Anio_Mes" o
# "Clave_anio_mes", y la tabla de SQL Server tiene el suyo propio. Por eso no hay
# un nombre fijo: se busca el que corresponde en cada lado comparando sin tildes,
# espacios ni guiones bajos.
#
# Y no alcanza con encontrarlos: si el parquet y la tabla los escriben distinto,
# el to_sql tampoco calza, porque mapea por NOMBRE de columna. Antes de cargar se
# renombran las columnas del parquet a como se llaman en la tabla.
COL_PERIODO = "Clave Año_Mes"           # solo para los mensajes
COL_SUMINISTRADOR = "Suministrador"
COL_HORA = "Hora Mensual"

# El mes del cambio de hora de primavera tiene una hora MENOS: esa hora no
# existe. El archivo suele venir con las horas corridas (1..719 en vez de
# 1..720, con la 145 ocupada por lo que en realidad es la 146), asi que hay que
# empujar todo desde la hora del cambio.
#
# No se toca el parquet: el desplazamiento se hace al cargar.
HORA_CAMBIO_POR_OMISION = 145


def clave_col(t):
    """Normaliza un nombre de columna para comparar.

    Sin tildes, sin espacios ni guiones bajos, en mayusculas. Y ademas trata
    "ANIO" como "ANO", que es lo que hace falta de verdad: la Ñ no se resuelve
    quitando tildes. Descomponer "Año" da "ANO" y "Anio" da "ANIO", y sin este
    paso no coincidirian, que es justo el caso que falla:
        'Clave Año_Mes'  ==  'Clave_Anio_Mes'  ==  'clave_anio_mes'
    """
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"[\s_]+", "", t).upper()
    # La Ñ descompuesta queda como N; "ANIO" se lleva a "ANO" para que las dos
    # formas de escribir "año" den lo mismo.
    return t.replace("ANIO", "ANO")


def resolver_columna(columnas, objetivo):
    """El nombre REAL de la columna, o None. Primero exacto, despues normalizado."""
    for c in columnas:
        if str(c) == objetivo:
            return c
    obj = clave_col(objetivo)
    for c in columnas:
        if clave_col(c) == obj:
            return c
    return None


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
#  UTILIDADES
# =============================================================================
def normalizar(t):
    if t is None:
        return ""
    t = unicodedata.normalize("NFKD", str(t))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.lower().split())


def fmt_tiempo(seg):
    seg = int(seg)
    return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"


def fmt_n(n):
    """12345 -> '12.345'  (separador de miles chileno)."""
    try:
        return f"{int(n):,}".replace(",", ".")
    except (TypeError, ValueError):
        return str(n)


def buscar_parquet(carpeta_reliq):
    """El Retiros_h.parquet dentro de '04 Planilla 9'."""
    base = Path(carpeta_reliq)
    if not base.is_dir():
        return None
    objetivo = normalizar(CARPETA_PARQUET)
    for p in base.iterdir():
        if p.is_dir() and normalizar(p.name) == objetivo:
            exacto = p / NOMBRE_PARQUET
            if exacto.is_file():
                return exacto
            # Por si el nombre trae algun sufijo.
            cand = sorted(p.glob("*.parquet"),
                          key=lambda x: x.stat().st_mtime, reverse=True)
            for c in cand:
                if normalizar(c.stem).startswith(normalizar("retiros")):
                    return c
            return None
    return None


def aplicar_cambio_hora(df, col_hora, hora_cambio, log):
    """Empuja una hora hacia arriba todo lo que este DESDE hora_cambio.

    En el mes del cambio de primavera la hora `hora_cambio` no existe. Si el
    archivo viene corrido (1..719), lo que esta guardado como 145 es en realidad
    la 146, lo que esta como 146 es la 147, y asi. Sumando 1 desde ahi queda
    1..144 y 146..720, con la 145 ausente, que es lo correcto.

    Devuelve (df, aplicado, motivo).

    COMO SE DETECTA SI YA VIENE APLICADO: mirando si la hora del cambio EXISTE
    en los datos.
      - si existe  -> todavia no se aplico, hay que empujar
      - si no esta -> ya se aplico (o no hay retiros a esa hora), no se toca
    Es la comprobacion correcta porque el desplazamiento es justamente lo que
    deja ese hueco: no se puede aplicar dos veces sin que se note.
    """
    horas = df[col_hora].dropna()
    if not len(horas):
        return df, False, "no hay horas en el archivo"
    try:
        horas = horas.astype("int64")
    except Exception:
        return df, False, f"la columna '{col_hora}' no es numérica"

    mn, mx, n_dist = int(horas.min()), int(horas.max()), horas.nunique()
    log(f"    horas en el archivo: {mn} a {mx}, {n_dist} distintas")

    if hora_cambio not in set(horas.unique().tolist()):
        return df, False, (f"la hora {hora_cambio} NO está en el archivo: "
                           f"el desplazamiento ya venía aplicado")

    # Si ya llega al largo de un mes completo, empujar dejaria una hora de mas.
    if mx >= 744:
        return df, False, (f"la hora máxima ya es {mx}: empujar dejaría "
                           f"{mx + 1}, que no es un mes válido. NO se aplica.")

    n_mueve = int((horas >= hora_cambio).sum())
    df = df.copy()
    mask = df[col_hora].astype("int64") >= hora_cambio
    df.loc[mask, col_hora] = df.loc[mask, col_hora].astype("int64") + 1
    nuevas = df[col_hora].astype("int64")
    log(f"    desplazadas {fmt_n(n_mueve)} fila(s) desde la hora "
        f"{hora_cambio}: ahora van {int(nuevas.min())} a {int(nuevas.max())}")
    log(f"    la hora {hora_cambio} queda vacía, que es lo que corresponde")
    return df, True, f"desplazadas {fmt_n(n_mueve)} filas"


# =============================================================================
#  PROCESO
# =============================================================================
def ejecutar(ruta_parquet, base_datos, hora_cambio, log, progreso):
    """Carga el parquet en la tabla, vaciandola antes.

    hora_cambio: None, o la hora mensual del cambio de horario de primavera.
    Si viene, se empuja una hora todo lo que este desde ahi (ver
    aplicar_cambio_hora). El parquet NO se toca.

    Devuelve (ok, resumen).
    """
    try:
        import pandas as pd
        from sqlalchemy import create_engine, text
    except ImportError as e:
        log(f"ERROR: falta una librería: {e}")
        log("       Hacen falta pandas, sqlalchemy, pyodbc y pyarrow.")
        return False, f"falta una librería: {e}"

    t0 = time.time()
    try:
        # ---- 1) el parquet ------------------------------------------------
        progreso(5, 100, "Leyendo el parquet...")
        log(f"Parquet : {ruta_parquet}")
        log(f"Servidor: {SERVER}   |   Base: {base_datos}   |   Tabla: {TABLA}")
        log("-" * 70)
        log("Leyendo parquet...")
        df = pd.read_parquet(ruta_parquet)
        if len(df) == 0:
            return False, ("El parquet no tiene filas. No se borra ni se carga "
                           "nada.")
        # Como se llaman en el PARQUET.
        col_per_pq = resolver_columna(df.columns, COL_PERIODO)
        col_sum_pq = resolver_columna(df.columns, COL_SUMINISTRADOR)
        faltan = [n for n, c in ((COL_PERIODO, col_per_pq),
                                 (COL_SUMINISTRADOR, col_sum_pq)) if c is None]
        if faltan:
            return False, (f"El parquet no tiene la(s) columna(s) "
                           f"{', '.join(faltan)}, ni con otro nombre parecido.\n\n"
                           f"Columnas del parquet: "
                           f"{', '.join(map(str, df.columns))}")
        for pedido, real in ((COL_PERIODO, col_per_pq),
                             (COL_SUMINISTRADOR, col_sum_pq)):
            if str(real) != pedido:
                log(f"  '{pedido}' en el parquet se llama '{real}'")
        n_sum_parquet = df[col_sum_pq].nunique()
        log(f"  {fmt_n(len(df))} filas, {n_sum_parquet} suministradores")
        periodos = sorted(df[col_per_pq].unique().tolist())
        log(f"  período(s) en el archivo: {periodos}")

        # ---- cambio de hora ----------------------------------------------
        aviso_hora = None
        if hora_cambio:
            log("")
            log(f"Cambio de hora (−): hora del mes {hora_cambio}")
            col_h = resolver_columna(df.columns, COL_HORA)
            if col_h is None:
                return False, (f"Se pidió el cambio de hora pero el parquet no "
                               f"tiene la columna '{COL_HORA}'.\n\n"
                               f"Columnas: {', '.join(map(str, df.columns))}")
            if len(periodos) > 1:
                log(f"    OJO: el archivo trae {len(periodos)} períodos y el "
                    f"cambio de hora es de UN mes. Se aplica a todas las filas.")
            df, aplicado, motivo = aplicar_cambio_hora(df, col_h,
                                                       int(hora_cambio), log)
            if aplicado:
                aviso_hora = f"cambio de hora aplicado en la {hora_cambio}"
            else:
                log(f"    NO se aplicó: {motivo}")
                aviso_hora = f"cambio de hora NO aplicado ({motivo})"

        # ---- 2) conexion --------------------------------------------------
        progreso(10, 100, "Conectando...")
        conn_str = (f"mssql+pyodbc://@{SERVER}/{base_datos}"
                    f"?driver={DRIVER}&trusted_connection=yes")
        engine = create_engine(conn_str, fast_executemany=True)

        # ---- como se llaman las columnas en la TABLA ----------------------
        # to_sql mapea por nombre, asi que si el parquet las escribe distinto que
        # la tabla, hay que renombrarlas antes de cargar.
        with engine.connect() as cn:
            hay_tabla = cn.execute(text(
                "SELECT 1 FROM sys.objects o JOIN sys.schemas e "
                "ON e.schema_id = o.schema_id "
                "WHERE o.name = :n AND e.name = 'dbo' AND o.type = 'U'"),
                {"n": TABLA}).first() is not None
            if not hay_tabla:
                return False, (
                    f"La tabla {TABLA} no existe en {base_datos}.\n\n"
                    f"Este script CARGA en una tabla que ya tiene que estar "
                    f"creada, con sus tipos e índices. Crearla al vuelo la "
                    f"dejaría sin índices y con las columnas de texto en "
                    f"VARCHAR(MAX), que es justo lo que hace lenta la prorrata."
                    f"\n\nRevisá que sea la base correcta.")
            cols_tabla = list(pd.read_sql(
                text(f"SELECT TOP 0 * FROM {TABLA}"), cn).columns)
        log(f"  Columnas de la tabla: {', '.join(map(str, cols_tabla))}")

        renombres, sin_par = {}, []
        for c in df.columns:
            real = resolver_columna(cols_tabla, str(c))
            if real is None:
                sin_par.append(str(c))
            elif str(real) != str(c):
                renombres[c] = real
        if sin_par:
            return False, (f"El parquet trae columnas que la tabla {TABLA} no "
                           f"tiene: {', '.join(sin_par)}.\n\n"
                           f"Columnas de la tabla: "
                           f"{', '.join(map(str, cols_tabla))}")
        if renombres:
            log(f"  Renombrando {len(renombres)} columna(s) para que calcen con "
                f"la tabla:")
            for a, b in renombres.items():
                log(f"     '{a}'  ->  '{b}'")
            df = df.rename(columns=renombres)
            col_per_pq = renombres.get(col_per_pq, col_per_pq)
            col_sum_pq = renombres.get(col_sum_pq, col_sum_pq)

        # Los nombres que van dentro del SQL, ya resueltos contra la tabla.
        col_per = resolver_columna(cols_tabla, COL_PERIODO)
        col_sum = resolver_columna(cols_tabla, COL_SUMINISTRADOR)
        if col_per is None or col_sum is None:
            return False, (f"La tabla {TABLA} no tiene las columnas de período o "
                           f"suministrador.\n\nColumnas: "
                           f"{', '.join(map(str, cols_tabla))}")
        faltan_en_pq = [str(c) for c in cols_tabla
                        if resolver_columna(df.columns, str(c)) is None]
        if faltan_en_pq:
            log(f"  OJO: la tabla tiene columnas que el parquet no trae: "
                f"{', '.join(faltan_en_pq)}. Van a quedar en NULL.")

        # ---- 3) borrado ---------------------------------------------------
        # SIEMPRE se vacia la tabla entera: es el unico modo que se usa. Antes
        # habia tambien un borrado por periodo, que nadie usaba y complicaba la
        # verificacion del final.
        progreso(20, 100, "Vaciando la tabla...")
        with engine.begin() as cn:
            antes = cn.execute(text(f"SELECT COUNT(*) FROM {TABLA}")).scalar()
            log(f"Filas en el servidor antes: {fmt_n(antes)}")
            log("Vaciando la tabla completa (TRUNCATE)...")
            cn.execute(text(f"TRUNCATE TABLE {TABLA}"))
            despues = cn.execute(text(f"SELECT COUNT(*) FROM {TABLA}")).scalar()
            log(f"Filas tras el borrado: {fmt_n(despues)}")

        # ---- 4) carga -----------------------------------------------------
        # Tipos por si la tabla NO existe y to_sql tiene que crearla. Si ya
        # existe, to_sql los ignora y usa los que la tabla tenga.
        from sqlalchemy.types import NVARCHAR
        # OJO con la prueba de "es texto": en pandas 3 las columnas de string ya
        # NO son dtype object, son StringDtype. Con "== object" no entraba
        # ninguna y los tipos no se aplicaban. is_string_dtype cubre las dos.
        dtipos = {c: NVARCHAR(length=LARGO_TEXTO)
                  for c in df.columns
                  if pd.api.types.is_string_dtype(df[c])
                  or df[c].dtype == object}
        log(f"  {len(dtipos)} columna(s) de texto -> NVARCHAR({LARGO_TEXTO}) "
            f"si hubiera que crear la tabla")

        log("Cargando...")
        total_chunks = (len(df) + CHUNK - 1) // CHUNK
        cargadas = 0
        for i in range(total_chunks):
            chunk = df.iloc[i * CHUNK:(i + 1) * CHUNK]
            chunk.to_sql(TABLA, engine, if_exists="append", index=False,
                         dtype=dtipos)
            cargadas += len(chunk)
            pct = 100 * (i + 1) // total_chunks
            progreso(30 + int(pct * 0.6), 100,
                     f"Cargando... {pct}%  ({fmt_n(cargadas)} filas)")
            if (i + 1) % max(1, total_chunks // 10) == 0 or i + 1 == total_chunks:
                log(f"  {pct}%  ({fmt_n(cargadas)} filas)")

        # ---- 5) verificacion ----------------------------------------------
        progreso(95, 100, "Verificando...")
        with engine.connect() as cn:
            final = cn.execute(text(f"SELECT COUNT(*) FROM {TABLA}")).scalar()
            n_sum = cn.execute(text(
                f"SELECT COUNT(DISTINCT [{col_sum}]) FROM {TABLA}")).scalar()
            per_srv = [r[0] for r in cn.execute(text(
                f"SELECT DISTINCT [{col_per}] FROM {TABLA} "
                f"ORDER BY 1")).fetchall()]
        log("")
        log(f"Servidor: {fmt_n(final)} filas | {n_sum} suministradores | "
            f"períodos {per_srv}")

        esperado = despues + len(df)
        avisos = []
        if final != esperado:
            m = (f"*** ALERTA: se esperaban {fmt_n(esperado)} filas y hay "
                 f"{fmt_n(final)} (diferencia {final - esperado:+,}) ***")
            log(m)
            avisos.append(f"filas: {final - esperado:+,}")
        elif n_sum != n_sum_parquet:
            # La tabla se vacio, asi que lo unico que hay es lo del parquet: si
            # los suministradores no coinciden, sobro o falto algo.
            m = (f"*** ALERTA: el servidor tiene {n_sum} suministradores y el "
                 f"parquet {n_sum_parquet} — hay residuo de otras cargas ***")
            log(m)
            avisos.append(f"suministradores: {n_sum} vs {n_sum_parquet}")
        else:
            log("Verificación OK: coincide con el parquet.")

        # Se comparan como texto: el parquet puede traer el periodo como numero y
        # el servidor devolverlo como texto (o al reves), y ahi 2412 != "2412"
        # daria una alerta falsa.
        per_srv_txt = {str(x).strip() for x in per_srv}
        falta = [p for p in periodos if str(p).strip() not in per_srv_txt]
        if falta:
            log(f"*** ALERTA: los períodos {falta} no quedaron en el servidor ***")
            avisos.append(f"faltan períodos {falta}")

        log(f"Terminado en {time.time() - t0:.1f} s")
        progreso(100, 100, "Listo")
        resumen = f"{fmt_n(len(df))} filas cargadas, {fmt_n(final)} en total"
        if aviso_hora:
            resumen += f"  |  {aviso_hora}"
        if avisos:
            resumen += "  |  OJO: " + "; ".join(avisos)
        return not avisos, resumen

    except Exception as e:
        log(f"\nERROR: {e}")
        log(traceback.format_exc())
        log("")
        # SQL Server contesta "Login failed for user" tanto cuando faltan
        # permisos como cuando la base NO EXISTE (por ejemplo, el nombre mal
        # escrito). El mensaje crudo despista: parece un problema de cuenta.
        txt_e = str(e)
        if "4060" in txt_e or "Cannot open database" in txt_e:
            log("")
            log(f"El servidor no pudo abrir la base '{base_datos}'.")
            log("Ese error sale por DOS motivos, y el mensaje no los distingue:")
            log("   1. la base NO EXISTE con ese nombre exacto (una tilde de más,")
            log("      un guión bajo distinto), o")
            log("   2. tu usuario no tiene permiso sobre ella.")
            log("Compará el nombre contra el que aparece en Management Studio.")
            return False, (
                f"No se pudo abrir la base '{base_datos}'.\n\n"
                f"O el nombre no es exacto, o no tenés permiso.\n"
                f"Comparalo con el que aparece en Management Studio.\n\n"
                f"El detalle está en el log.")
        log("OJO: si el error fue DESPUÉS del borrado, la tabla quedó a medias.")
        log("     Volvé a correrlo: vacía la tabla y recarga. Lo que no hay que")
        log("     hacer es dejarla así.")
        return False, txt_e


# =============================================================================
#  VENTANA
# =============================================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    root = tk.Tk()
    root.title("Cargar Retiros a SQL Server"
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
        tk.Label(fa, text="La ruta del parquet la resolvió el Revisor.\n"
                          "Elegí la base antes de cargar.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    tk.Label(cont, text="Cargar Retiros_h.parquet a SQL Server",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(cont, text="Esto BORRA y vuelve a cargar filas en la base de datos. "
                        "Antes de escribir muestra qué va a borrar.",
             font=("Segoe UI", 8), fg="#a00000").pack(pady=(0, 8))

    var_parquet = tk.StringVar(value="[pendiente]")
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_base = tk.StringVar(value=cfg.get("retiros_base") or BASES[0])
    labels = {}

    # ---- el parquet ----
    fp = tk.LabelFrame(cont, text=f"Archivo — {NOMBRE_PARQUET}", padx=10, pady=6)
    fp.pack(fill="x", padx=20, pady=4)
    lp = tk.Label(fp, textvariable=var_parquet, wraplength=880, justify="left",
                  cursor="hand2", font=("Segoe UI", 9), anchor="w")
    lp.pack(fill="x")
    lp.bind("<Button-1>", lambda e: abrir_en_explorador(var_parquet.get(),
                                                        es_archivo=True))
    labels["parquet"] = lp

    def pintar():
        v = var_parquet.get()
        ok = bool(v) and not v.startswith("[")
        if ok:
            try:
                ok = Path(v).is_file()
            except Exception:
                ok = False
        labels["parquet"].config(fg="blue" if ok else "red")
        c = var_carpeta.get()
        labels["carpeta"].config(fg="blue" if c and Path(c).is_dir() else "red")

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            var_parquet.set(d.get("retiros_parquet")
                            or "[el Revisor no mandó retiros_parquet]")
            pintar()
            return
        base = var_carpeta.get()
        p = buscar_parquet(base) if base else None
        var_parquet.set(str(p) if p else
                        f"[{NOMBRE_PARQUET} no encontrado en {CARPETA_PARQUET}]")
        pintar()

    def sel_parquet():
        """Elegir el .parquet directo, sin pasar por la carpeta del caso."""
        ini = var_parquet.get()
        ini = str(Path(ini).parent) if ini and not ini.startswith("[") else ""
        r = filedialog.askopenfilename(
            title=f"Seleccionar {NOMBRE_PARQUET}", initialdir=ini,
            filetypes=[("Parquet", "*.parquet"), ("Todos", "*.*")])
        if r:
            var_parquet.set(r)
            guardar_config({"retiros_parquet": r})
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Archivo elegido a mano: se deja de usar la ruta del Revisor.")
            pintar()

    tk.Button(fp, text="Examinar archivo...",
              command=sel_parquet).pack(anchor="w", pady=(4, 0))

    # ---- la base ----
    fb = tk.LabelFrame(cont, text="Base de datos", padx=10, pady=6)
    fb.pack(fill="x", padx=20, pady=4)
    tk.Label(fb, text=f"Servidor {SERVER}   ·   tabla {TABLA}",
             font=("Segoe UI", 8), fg="#555555").pack(anchor="w")
    for b in BASES:
        tk.Radiobutton(fb, text=b, variable=var_base, value=b,
                       font=("Consolas", 10)).pack(anchor="w")

    # ---- cambio de hora ----
    # Ya no hay opciones de borrado: SIEMPRE se vacia la tabla. El borrado por
    # periodo y el modo "solo mirar" se sacaron porque no se usaban.
    fo = tk.LabelFrame(cont, text="Cambio de hora", padx=10, pady=6)
    fo.pack(fill="x", padx=20, pady=4)
    var_cambio = tk.BooleanVar(value=False)
    var_hora = tk.StringVar(value=str(HORA_CAMBIO_POR_OMISION))
    fila_h = tk.Frame(fo)
    fila_h.pack(fill="x")
    tk.Checkbutton(fila_h, text="Cambio de hora (−): esa hora no existe",
                   variable=var_cambio,
                   font=("Segoe UI", 9)).pack(side="left")
    tk.Label(fila_h, text="   hora del mes:",
             font=("Segoe UI", 9)).pack(side="left")
    tk.Entry(fila_h, textvariable=var_hora, width=6,
             font=("Consolas", 10)).pack(side="left", padx=4)
    tk.Label(fo, text="Los retiros desde esa hora se corren una hora hacia "
                      "arriba al cargar, para que la hora del cambio quede "
                      "vacía.\nEl parquet NO se modifica. Si el archivo ya "
                      "viene corrido, se detecta y no se aplica dos veces.",
             font=("Segoe UI", 8), fg="#555555",
             justify="left").pack(anchor="w", pady=(4, 0))

    # ---- carpeta del caso ----
    fc = tk.LabelFrame(cont, text="Carpeta 02 CASO RELIQUIDACION  "
                                  "(para buscar el parquet solo)", padx=10, pady=6)
    fc.pack(fill="x", padx=20, pady=4)
    lc = tk.Label(fc, textvariable=var_carpeta, wraplength=880, justify="left",
                  cursor="hand2", font=("Segoe UI", 9), anchor="w")
    lc.pack(fill="x")
    lc.bind("<Button-1>", lambda e: abrir_en_explorador(var_carpeta.get()))
    labels["carpeta"] = lc

    def sel_carpeta():
        ini = var_carpeta.get()
        r = filedialog.askdirectory(title="Selecciona 02 CASO RELIQUIDACION",
                                    initialdir=ini if ini and Path(ini).is_dir() else "")
        if r:
            var_carpeta.set(r)
            guardar_config({"carpeta_reliq": r})
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Carpeta elegida a mano: se deja de usar la ruta del Revisor.")
            refrescar()

    tk.Button(fc, text="Examinar carpeta...",
              command=sel_carpeta).pack(anchor="w", pady=(4, 0))

    # ---- progreso ----
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

    # ---- lanzar ----
    def lanzar():
        p = var_parquet.get()
        if not p or p.startswith("[") or not Path(p).is_file():
            messagebox.showerror("Falta el archivo",
                                 f"No se encontró el parquet.\n\n{p}")
            return
        base = var_base.get()
        if base not in BASES:
            messagebox.showerror("Base inválida", f"Base desconocida: {base}")
            return
        # El cambio de hora: solo si esta marcado, y la hora tiene que ser un
        # numero razonable. Se valida ANTES de tocar la base.
        hora_cambio = None
        if var_cambio.get():
            txt_h = var_hora.get().strip()
            try:
                hora_cambio = int(txt_h)
            except ValueError:
                messagebox.showerror("Hora inválida",
                                     f"La hora del cambio tiene que ser un "
                                     f"número.\n\nSe escribió: {txt_h!r}")
                return
            if not 1 <= hora_cambio <= 744:
                messagebox.showerror("Hora inválida",
                                     f"La hora del mes tiene que estar entre 1 "
                                     f"y 744.\n\nSe escribió: {hora_cambio}")
                return

        linea_hora = (f"\n\nCambio de hora (−) en la hora {hora_cambio}: los "
                      f"retiros desde ahí se corren una hora."
                      if hora_cambio else
                      "\n\nSin cambio de hora.")
        if not messagebox.askyesno(
                "Confirmar escritura en la base",
                f"Servidor : {SERVER}\n"
                f"Base     : {base}\n"
                f"Tabla    : {TABLA}\n\n"
                f"Se VACÍA la tabla completa (TRUNCATE) y se carga:\n"
                f"   {Path(p).name}"
                + linea_hora +
                "\n\nEsto MODIFICA la base de datos. ¿Seguir?"):
            return

        guardar_config({"retiros_base": base})

        txt.delete("1.0", "end")
        if modo["traspaso"]:
            log(f"Rutas enviadas por el Revisor — mes {traspaso.get('aamm') or '?'}")
        btn.config(state="disabled", bg="#aaaaaa")
        barra["value"] = 0
        var_estado.set("Procesando...")
        timer["on"] = True
        timer["t0"] = time.time()
        tick()

        def trabajo():
            ok, msg = ejecutar(p, base, hora_cambio, log, progreso)

            def fin():
                timer["on"] = False
                btn.config(state="normal", bg="#2d7a2d")
                if ok:
                    var_estado.set(f"Listo — {msg}")
                    messagebox.showinfo("Listo", f"Carga terminada.\n\n{msg}")
                else:
                    var_estado.set("Terminó con problemas — revisa el log")
                    messagebox.showerror("Problema", f"{msg}\n\nEl detalle está "
                                                     "en el log.")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text="CARGAR RETIROS", bg="#2d7a2d", fg="white",
                    font=("Segoe UI", 10, "bold"), command=lanzar)
    btn.pack(side="left", padx=8, expand=True)
    tk.Button(frame_btns, text="Refrescar ruta",
              command=refrescar).pack(side="left", padx=8)
    tk.Button(frame_btns, text="Salir", command=root.destroy).pack(side="left", padx=8)

    # Si vino guardada una ruta a mano y no hay traspaso, se usa esa.
    refrescar()
    if not modo["traspaso"] and var_parquet.get().startswith("["):
        guardada = cfg.get("retiros_parquet")
        if guardada and Path(guardada).is_file():
            var_parquet.set(guardada)
            pintar()

    bombear()
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    root.geometry(f"+{(root.winfo_screenwidth() - w) // 2}"
                  f"+{(root.winfo_screenheight() - h) // 2}")
    root.mainloop()


if __name__ == "__main__":
    main()
