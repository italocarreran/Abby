# -*- coding: utf-8 -*-
# =============================================================================
#  Prorratear: del Access a SQL Server
# =============================================================================
#  Automatiza lo que hoy se hace a mano en SQL Server Management Studio:
#
#    1. Borrar de la base de sobrecostos las tablas:
#         Central_Empresa, Pago_Retiro_reporte_tabla, Sobrecostos, TIPOS
#    2. Importarlas del Access (el "Tasks -> Import Data"):
#         Central_Empresa_Actualizada  ->  Central_Empresa
#           (el .mdb de la planilla 9 no la tiene: ahi la tabla ya se llama
#            Central_Empresa y se copia tal cual)
#         Sobrecostos                  ->  Sobrecostos
#         TIPOS                        ->  TIPOS
#    3. Correr:
#         SELECT Tipo_sobrecosto, Concepto, Barra, Suministrador, Retiro,
#                clave, Tipo, SUM(pago_retiro) AS Pago
#           INTO Pago_Retiro_reporte_tabla
#           FROM dbo.[10_Pago_retiros]
#          GROUP BY ...
#
#  10_Pago_retiros es una VISTA que ya existe en la base: cruza lo que se acaba
#  de importar con los retiros. Por eso el orden importa y el paso 3 va al final.
#
#  LAS BASES VAN DE A PARES  —  la de sobrecostos depende de donde estan los
#  retiros, y elegir mal significa prorratear contra los retiros equivocados:
#
#      02_RETIROS               ->  05_SOBRECOSTOS
#      14_RETIROS_RELIQUIDACION ->  16_SOBRECOSTOS_RELIQUIDACION
#
#  Por eso en la ventana se elige el ESCENARIO, no las dos bases por separado.
#
#  ESTO BORRA TABLAS DE UNA BASE DE DATOS. Por eso:
#    - hay un modo SOLO MIRAR que cuenta todo y no toca nada
#    - antes de borrar se comprueba que la tabla EXISTA
#    - se muestra que se va a borrar y se pide confirmacion
#    - si falla despues del borrado, se avisa fuerte: hay que volver a correrlo
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

# ---------------------------------------------------------------------------
#  Configuracion
# ---------------------------------------------------------------------------
SERVER = "SRV-DTE"
DRIVER_SQL = "ODBC Driver 17 for SQL Server"

# Los dos escenarios. La base de sobrecostos NO se elige suelta: va atada a la
# de retiros, porque la vista 10_Pago_retiros cruza contra esos retiros.
ESCENARIOS = {
    "normal": {
        "etiqueta": "Normal          retiros en 02_RETIROS  ->  05_SOBRECOSTOS",
        "retiros": "02_RETIROS",
        "sobrecostos": "05_SOBRECOSTOS",
    },
    "reliq": {
        "etiqueta": "Reliquidación   retiros en 14_RETIROS_RELIQUIDACION  ->  "
                    "16_SOBRECOSTOS_RELIQUIDACION",
        "retiros": "14_RETIROS_RELIQUIDACION",
        "sobrecostos": "16_SOBRECOSTOS_RELIQUIDACION",
    },
}
ORDEN_ESC = ["normal", "reliq"]

# Tablas que se borran antes de importar, en la base de sobrecostos.
TABLAS_A_BORRAR = ["Central_Empresa", "Pago_Retiro_reporte_tabla",
                   "Sobrecostos", "TIPOS"]

# Que se copia del Access. El primero de "origen" que exista es el que se usa.
# En el .mdb de SSCC la tabla se llama Central_Empresa_Actualizada; en el de la
# planilla 9 se llama Central_Empresa.
IMPORTAR = [
    {"destino": "Central_Empresa",
     "origen": ["Central_Empresa_Actualizada", "Central_Empresa"]},
    {"destino": "Sobrecostos", "origen": ["Sobrecostos"]},
    {"destino": "TIPOS", "origen": ["TIPOS"]},
]

# Columnas del periodo, para comprobar que los retiros y los sobrecostos sean
# del MISMO mes antes de prorratear.
COL_CLAVE_ACCESS = "Clave Año_Mes"      # en la tabla Sobrecostos del .mdb
COL_CLAVE_RETIROS = "Clave_Anio_Mes"    # en la tabla Retiros de SQL Server
TABLA_RETIROS = "Retiros"

VISTA_PAGO = "10_Pago_retiros"
TABLA_REPORTE = "Pago_Retiro_reporte_tabla"
SQL_REPORTE = """
SELECT Tipo_sobrecosto, Concepto, Barra, Suministrador, Retiro, clave, Tipo,
       SUM(pago_retiro) AS Pago
  INTO {destino}
  FROM dbo.[{vista}]
 GROUP BY Tipo_sobrecosto, Concepto, Barra, Suministrador, Retiro, clave, Tipo
"""

CHUNK = 20_000      # filas por lote al escribir en SQL Server

# Largo de las columnas de texto al crear las tablas. 255 es lo que crea el
# asistente "Import Data" y lo que tienen hoy las tablas de las dos bases.
# Conviene que TODAS las columnas del join tengan el mismo tipo y largo en las
# dos bases: si difieren, el servidor convierte fila por fila y el indice deja
# de servir.
LARGO_TEXTO = 255


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
    ruta.parent.mkdir(parents=True, exist_ok=True)
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
    """El dict del traspaso, o None. Nunca lanza: si el JSON esta roto se cae al
    modo manual."""
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
    try:
        return f"{int(n):,}".replace(",", ".")
    except (TypeError, ValueError):
        return str(n)


def driver_access():
    import pyodbc
    todos = list(pyodbc.drivers())
    acc = [d for d in todos if "access" in d.lower()]
    if not acc:
        bits = "64" if sys.maxsize > 2 ** 32 else "32"
        detalle = ("pyodbc NO ve NINGÚN driver ODBC."
                   if not todos else
                   "Drivers que sí ve:\n   - " + "\n   - ".join(todos))
        raise RuntimeError(
            f"No se encontró el driver ODBC de Access.\n\n{detalle}\n\n"
            f"Python: {bits} bits\n   {sys.executable}")
    for d in acc:
        if "*.mdb, *.accdb" in d:
            return d
    return acc[0]


def tablas_access(ruta_mdb):
    """Nombres de las tablas de usuario del .mdb."""
    import pyodbc
    cn = pyodbc.connect(
        f"DRIVER={{{driver_access()}}};DBQ={ruta_mdb};ExtendedAnsiSQL=1;",
        autocommit=True)
    try:
        return [r.table_name for r in cn.cursor().tables(tableType="TABLE")]
    finally:
        cn.close()


def elegir_origen(disponibles, candidatos):
    """El primer candidato que exista, comparando sin mayusculas ni tildes."""
    mapa = {normalizar(t): t for t in disponibles}
    for c in candidatos:
        if normalizar(c) in mapa:
            return mapa[normalizar(c)]
    return None


def objeto_existe(cn, nombre, text, tipos=("U",)):
    """True si el objeto existe en la base conectada.

    OJO con el tipo: en OBJECT_ID(nombre, tipo) la 'U' es TABLA DE USUARIO y la
    'V' es VISTA. Preguntando por 'U' una vista devuelve NULL aunque exista,
    que es justo lo que pasaba con 10_Pago_retiros.

    Se consulta sys.objects, que no obliga a elegir un tipo y ademas permite
    preguntar por varios de una.
    """
    marcas = ", ".join(f"'{t}'" for t in tipos)
    q = text(f"""
        SELECT 1
          FROM sys.objects o
          JOIN sys.schemas e ON e.schema_id = o.schema_id
         WHERE o.name = :n AND e.name = 'dbo' AND o.type IN ({marcas})
    """)
    return cn.execute(q, {"n": nombre}).first() is not None


def tabla_existe(cn, nombre, text):
    """Solo tablas: es lo que se puede borrar con DROP TABLE."""
    return objeto_existe(cn, nombre, text, tipos=("U",))


def vista_o_tabla_existe(cn, nombre, text):
    """Vista o tabla. La 10_Pago_retiros es una VISTA, pero si algun dia fuera
    una tabla el paso final funcionaria igual, asi que se aceptan las dos."""
    return objeto_existe(cn, nombre, text, tipos=("V", "U"))


def contar(cn, nombre, text):
    try:
        return cn.execute(text(f"SELECT COUNT(*) FROM dbo.[{nombre}]")).scalar()
    except Exception:
        return None


def tipos_sql(df):
    """(tipos, avisos) para crear la tabla en SQL Server.

    Hace falta porque to_sql, sin decirle nada, crea las columnas de texto como
    NVARCHAR(MAX). Y NVARCHAR(MAX) **no se puede indexar**: un join sobre esa
    columna obliga al servidor a recorrer la tabla entera. El asistente
    "Import Data" de Management Studio crea longitudes concretas, y por eso las
    tablas que hizo el asistente andan mas rapido que las que haria to_sql a
    secas.

    El texto va en NVARCHAR(255), igual que lo que crea el asistente y que lo
    que tienen hoy las tablas. Solo se agranda si algun dato no entra.
    """
    from sqlalchemy.types import (NVARCHAR, Integer, BigInteger, Float,
                                  DateTime, Boolean)
    import pandas as pd

    tipos, avisos = {}, []
    for col in df.columns:
        ser = df[col]
        if pd.api.types.is_bool_dtype(ser):
            tipos[col] = Boolean()
        elif pd.api.types.is_integer_dtype(ser):
            mx = ser.abs().max() if len(ser.dropna()) else 0
            tipos[col] = BigInteger() if (mx or 0) > 2_000_000_000 else Integer()
        elif pd.api.types.is_float_dtype(ser):
            tipos[col] = Float()
        elif pd.api.types.is_datetime64_any_dtype(ser):
            tipos[col] = DateTime()
        else:
            # NVARCHAR(255), que es lo que crea el asistente "Import Data" y lo
            # que tienen hoy las tablas de las dos bases (nvarchar(510) en
            # sys.columns = 510 bytes = 255 caracteres).
            #
            # Se usa el mismo largo para todas en vez de ajustarlo al dato: si
            # una columna del join quedara con OTRO largo que su par en la otra
            # base, el servidor tendria que convertir y el indice dejaria de
            # servir. Igualar lo que ya funciona vale mas que afinar el tamaño.
            #
            # Si algun texto no entra en 255, se agranda solo y se avisa.
            largos = ser.dropna().astype(str).str.len()
            mx = int(largos.max()) if len(largos) else 0
            n = LARGO_TEXTO
            if mx > LARGO_TEXTO:
                n = min(((mx + 50) // 50 + 1) * 50, 4000)
                avisos.append(f"{col}: hay textos de {mx} caracteres, se usa "
                              f"NVARCHAR({n}) en vez de {LARGO_TEXTO}")
            tipos[col] = NVARCHAR(length=n)
    return tipos, avisos


def periodos_access(ruta_mdb, tabla, columna):
    """Los Clave Año_Mes que hay en una tabla del .mdb."""
    import pyodbc
    cn = pyodbc.connect(
        f"DRIVER={{{driver_access()}}};DBQ={ruta_mdb};ExtendedAnsiSQL=1;",
        autocommit=True)
    try:
        cur = cn.cursor()
        cur.execute(f"SELECT DISTINCT [{columna}] FROM [{tabla}]")
        return sorted({str(r[0]).strip() for r in cur.fetchall()
                       if r[0] is not None and str(r[0]).strip()})
    finally:
        cn.close()


def periodos_retiros(base_retiros, text, create_engine):
    """Los Clave_Anio_Mes que hay hoy en la tabla Retiros de esa base."""
    cs = (f"mssql+pyodbc://@{SERVER}/{base_retiros}"
          f"?driver={DRIVER_SQL}&trusted_connection=yes")
    eng = create_engine(cs)
    try:
        with eng.connect() as cn:
            filas = cn.execute(text(
                f"SELECT [{COL_CLAVE_RETIROS}], COUNT(*) "
                f"FROM dbo.[{TABLA_RETIROS}] "
                f"GROUP BY [{COL_CLAVE_RETIROS}] "
                f"ORDER BY 1")).fetchall()
        return [(str(r[0]).strip(), r[1]) for r in filas if r[0] is not None]
    finally:
        eng.dispose()


# =============================================================================
#  PROCESO
# =============================================================================
def ejecutar(ruta_mdb, escenario, solo_mirar, log, progreso):
    """Devuelve (ok, resumen)."""
    try:
        import pandas as pd
        import pyodbc                                    # noqa: F401
        from sqlalchemy import create_engine, text
    except ImportError as e:
        log(f"ERROR: falta una librería: {e}")
        log("       Hacen falta pandas, sqlalchemy y pyodbc.")
        return False, f"falta una librería: {e}"

    esc = ESCENARIOS[escenario]
    base = esc["sobrecostos"]
    t0 = time.time()
    borrado_hecho = False
    try:
        log(f"Access   : {ruta_mdb}")
        log(f"Escenario: {esc['etiqueta']}")
        log(f"Servidor : {SERVER}   |   Base: {base}")
        log("-" * 70)

        # ---- 1) que hay en el Access -------------------------------------
        log("Mirando el Access...")
        disp = tablas_access(ruta_mdb)
        log(f"  Tablas del .mdb: {', '.join(sorted(disp))}")
        plan = []
        for item in IMPORTAR:
            org = elegir_origen(disp, item["origen"])
            if org is None:
                return False, (f"El Access no tiene ninguna de estas tablas: "
                               f"{', '.join(item['origen'])}.\n\n"
                               f"Tablas del .mdb: {', '.join(sorted(disp))}")
            plan.append((org, item["destino"]))
            flecha = "" if org == item["destino"] else f"  ->  {item['destino']}"
            log(f"    {org}{flecha}")

        # ---- 1.b) el mes tiene que ser el mismo de los dos lados ---------
        # Es lo primero que hay que mirar: prorratear los sobrecostos de un mes
        # contra los retiros de otro no da error, da un resultado MAL. Y como el
        # numero sale igual de plausible, no se nota despues.
        progreso(5, 100, "Comprobando el mes...")
        log("")
        log("Comprobando que los retiros sean del mismo mes que los sobrecostos...")
        try:
            per_acc = periodos_access(ruta_mdb, "Sobrecostos", COL_CLAVE_ACCESS)
        except Exception as ex:
            return False, (f"No se pudo leer la Clave Año_Mes del Access: {ex}")
        log(f"    Access ({esc['sobrecostos']} recibirá esto): {per_acc}")
        if not per_acc:
            return False, ("La tabla Sobrecostos del Access no tiene ninguna "
                           "Clave Año_Mes.")
        if len(per_acc) > 1:
            return False, (f"El Access tiene MÁS DE UN mes en Sobrecostos: "
                           f"{', '.join(per_acc)}.\n\nRevisá el .mdb antes de "
                           f"prorratear.")

        try:
            per_ret = periodos_retiros(esc["retiros"], text, create_engine)
        except Exception as ex:
            return False, (f"No se pudieron leer los períodos de "
                           f"{esc['retiros']}: {ex}")
        log(f"    Retiros ({esc['retiros']}):")
        for pp, nn in per_ret:
            marca = "  <-- coincide" if pp == per_acc[0] else ""
            log(f"       {pp}   {fmt_n(nn)} filas{marca}")
        if not per_ret:
            return False, (f"La tabla {TABLA_RETIROS} de {esc['retiros']} está "
                           f"VACÍA. Hay que cargar los retiros primero.")

        claves_ret = {pp for pp, _ in per_ret}
        if per_acc[0] not in claves_ret:
            return False, (
                f"EL MES NO COINCIDE.\n\n"
                f"El Access tiene sobrecostos del mes {per_acc[0]}, y en "
                f"{esc['retiros']} NO hay retiros de ese mes.\n\n"
                f"Meses que sí hay en los retiros: "
                f"{', '.join(sorted(claves_ret))}\n\n"
                f"Prorratear así daría un resultado equivocado. Cargá los "
                f"retiros del mes {per_acc[0]}, o revisá el escenario elegido.")
        log(f"    OK: los retiros de {per_acc[0]} están cargados.")
        if len(claves_ret) > 1:
            otros = sorted(claves_ret - {per_acc[0]})
            log(f"    (la base tiene además otros meses: {', '.join(otros)}. "
                f"Si la vista no filtra por período, eso hace más lenta la "
                f"prorrata.)")

        # ---- 2) leer del Access ------------------------------------------
        progreso(10, 100, "Leyendo el Access...")
        log("")
        log("Leyendo las tablas del Access...")
        cn_acc = pyodbc.connect(
            f"DRIVER={{{driver_access()}}};DBQ={ruta_mdb};ExtendedAnsiSQL=1;",
            autocommit=True)
        datos = {}
        try:
            for i, (org, dst) in enumerate(plan):
                df = pd.read_sql(f"SELECT * FROM [{org}]", cn_acc)
                datos[dst] = df
                log(f"  {org}: {fmt_n(len(df))} filas, {len(df.columns)} columnas")
                log(f"     {', '.join(map(str, df.columns))}")
                progreso(10 + int(20 * (i + 1) / len(plan)), 100,
                         "Leyendo el Access...")
        finally:
            cn_acc.close()

        # ---- 3) conectar a SQL Server ------------------------------------
        conn_str = (f"mssql+pyodbc://@{SERVER}/{base}"
                    f"?driver={DRIVER_SQL}&trusted_connection=yes")
        engine = create_engine(conn_str, fast_executemany=True)

        log("")
        log(f"Estado actual de [{base}]:")
        with engine.connect() as cn:
            if not vista_o_tabla_existe(cn, VISTA_PAGO, text):
                # Sin la vista el ultimo paso no puede correr, y hay que saberlo
                # ANTES de borrar nada.
                # Se listan las vistas que SI hay: si es un problema de nombre
                # o de esquema, se ve al toque.
                try:
                    hay = [r[0] for r in cn.execute(text(
                        "SELECT o.name FROM sys.objects o "
                        "JOIN sys.schemas e ON e.schema_id = o.schema_id "
                        "WHERE o.type = 'V' AND e.name = 'dbo' "
                        "ORDER BY o.name")).fetchall()]
                except Exception:
                    hay = []
                log(f"  Vistas de dbo en {base}: "
                    + (", ".join(hay) if hay else "(ninguna)"))
                return False, (f"La base {base} no tiene la vista "
                               f"[{VISTA_PAGO}] en el esquema dbo.\n\n"
                               f"Vistas que sí hay: "
                               f"{', '.join(hay) if hay else '(ninguna)'}\n\n"
                               f"Sin ella no se puede armar {TABLA_REPORTE}.")
            existen = {}
            for t in TABLAS_A_BORRAR:
                hay = tabla_existe(cn, t, text)
                n = contar(cn, t, text) if hay else None
                existen[t] = hay
                log(f"    {t:32s} " + (f"{fmt_n(n)} filas" if hay
                                       else "NO existe"))

        # ---- modo prueba --------------------------------------------------
        if solo_mirar:
            log("")
            log("=== SOLO MIRAR: no se borró ni se cargó nada ===")
            for t in TABLAS_A_BORRAR:
                log(f"    {'se borraría' if existen[t] else 'no hay que borrar'}"
                    f"  {t}")
            for org, dst in plan:
                log(f"    se cargarían {fmt_n(len(datos[dst]))} filas en {dst}"
                    f"  (desde {org})")
            log(f"    se armaría {TABLA_REPORTE} desde la vista [{VISTA_PAGO}]")
            progreso(100, 100, "Solo mirar: listo")
            return True, ("solo mirar: " + ", ".join(
                f"{dst} {fmt_n(len(datos[dst]))}" for _, dst in plan))

        # ---- 4) borrar ----------------------------------------------------
        progreso(35, 100, "Borrando las tablas...")
        log("")
        log("Borrando las tablas...")
        with engine.begin() as cn:
            for t in TABLAS_A_BORRAR:
                if not existen[t]:
                    log(f"    {t}: no existía, no se borra")
                    continue
                cn.execute(text(f"DROP TABLE dbo.[{t}]"))
                log(f"    {t}: borrada")
        borrado_hecho = True

        # ---- 5) importar --------------------------------------------------
        progreso(45, 100, "Cargando en SQL Server...")
        log("")
        log("Cargando en SQL Server...")
        for i, (org, dst) in enumerate(plan):
            df = datos[dst]
            dtipos, avisos_t = tipos_sql(df)
            df.to_sql(dst, engine, if_exists="replace", index=False,
                      chunksize=CHUNK, dtype=dtipos)
            log(f"    {dst}: tipos -> "
                + ", ".join(f"{c} {t}" for c, t in list(dtipos.items())[:6])
                + (" ..." if len(dtipos) > 6 else ""))
            for a_ in avisos_t:
                log(f"       OJO: {a_}")
            with engine.connect() as cn:
                n = contar(cn, dst, text)
            log(f"    {dst}: {fmt_n(n)} filas")
            if n != len(df):
                raise RuntimeError(
                    f"En {dst} quedaron {n} filas y se mandaron {len(df)}.")
            progreso(45 + int(35 * (i + 1) / len(plan)), 100,
                     "Cargando en SQL Server...")

        # ---- 6) el reporte ------------------------------------------------
        progreso(85, 100, f"Armando {TABLA_REPORTE}...")
        log("")
        log(f"Armando {TABLA_REPORTE} desde la vista [{VISTA_PAGO}]...")
        t1 = time.time()
        with engine.begin() as cn:
            cn.execute(text(SQL_REPORTE.format(destino=f"dbo.[{TABLA_REPORTE}]",
                                               vista=VISTA_PAGO)))
        with engine.connect() as cn:
            n_rep = contar(cn, TABLA_REPORTE, text)
            total = cn.execute(text(
                f"SELECT SUM(Pago) FROM dbo.[{TABLA_REPORTE}]")).scalar()
        log(f"    {fmt_n(n_rep)} filas en {time.time() - t1:.1f} s")
        log(f"    suma de Pago: {total:,.2f}".replace(",", "@").replace(
            ".", ",").replace("@", ".") if total is not None else
            "    suma de Pago: (vacía)")
        if not n_rep:
            log("    OJO: quedó VACÍA. Suele significar que los retiros de "
                f"{esc['retiros']} no son de este mes, o que la vista no "
                "encontró con qué cruzar.")

        log("")
        log(f"Terminado en {time.time() - t0:.1f} s")
        progreso(100, 100, "Listo")
        resumen = " | ".join(f"{dst} {fmt_n(len(datos[dst]))}" for _, dst in plan)
        resumen += f" | {TABLA_REPORTE} {fmt_n(n_rep)}"
        return bool(n_rep), resumen

    except Exception as e:
        log(f"\nERROR: {e}")
        log(traceback.format_exc())
        txt_e = str(e)
        if "4060" in txt_e or "Cannot open database" in txt_e:
            log("")
            log(f"El servidor no pudo abrir la base '{base}'.")
            log("Ese error sale por DOS motivos y el mensaje no los distingue:")
            log("   1. la base NO EXISTE con ese nombre exacto, o")
            log("   2. tu usuario no tiene permiso sobre ella.")
            log("Compará el nombre contra el que aparece en Management Studio.")
        if borrado_hecho:
            log("")
            log("OJO: las tablas YA SE BORRARON y la carga no terminó.")
            log(f"     La base {base} quedó incompleta.")
            log("     Volvé a correrlo: borra lo que haya y carga de nuevo.")
        return False, str(e)


# =============================================================================
#  VENTANA
# =============================================================================
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    root = tk.Tk()
    root.title("Prorratear — del Access a SQL Server"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("1000x740")

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
    def _rueda(e, _cv=canvas):
        w = e.widget.winfo_toplevel().winfo_containing(e.x_root, e.y_root)
        pasos = int(-e.delta / 120)
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
        tk.Label(fa, text="La ruta del Access la resolvió el Revisor.\n"
                          "Elegí el escenario antes de prorratear.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    tk.Label(cont, text="Prorratear: del Access a SQL Server",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(cont, text="Borra y vuelve a cargar cuatro tablas de la base de "
                        "sobrecostos, y arma el reporte.\n"
                        "El Access solo se LEE.",
             font=("Segoe UI", 8), fg="#555555", justify="center").pack(pady=(0, 8))

    var_mdb = tk.StringVar(value="[pendiente]")
    var_esc = tk.StringVar(value=cfg.get("prorrateo_escenario") or ORDEN_ESC[0])
    labels = {}

    fm = tk.LabelFrame(cont, text="Access con los sobrecostos (solo se lee)",
                       padx=10, pady=6)
    fm.pack(fill="x", padx=20, pady=4)
    lm = tk.Label(fm, textvariable=var_mdb, wraplength=880, justify="left",
                  cursor="hand2", font=("Segoe UI", 9), anchor="w")
    lm.pack(fill="x")
    lm.bind("<Button-1>", lambda e: abrir_en_explorador(var_mdb.get(),
                                                        es_archivo=True))
    labels["mdb"] = lm

    def pintar():
        v = var_mdb.get()
        ok = bool(v) and not v.startswith("[")
        if ok:
            try:
                ok = Path(v).is_file()
            except Exception:
                ok = False
        labels["mdb"].config(fg="blue" if ok else "red")

    def sel_mdb():
        ini = var_mdb.get()
        ini = str(Path(ini).parent) if ini and not ini.startswith("[") else ""
        r = filedialog.askopenfilename(
            title="Seleccionar el .mdb con los sobrecostos", initialdir=ini,
            filetypes=[("Access", "*.mdb *.accdb"), ("Todos", "*.*")])
        if r:
            var_mdb.set(r)
            guardar_config({"prorrateo_mdb": r})
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Archivo elegido a mano: se deja de usar la ruta del Revisor.")
            pintar()

    tk.Button(fm, text="Examinar...", command=sel_mdb).pack(anchor="w", pady=(4, 0))
    tk.Label(fm, text="Puede ser cualquiera de los tres .mdb de sobrecostos.",
             font=("Segoe UI", 8), fg="#555555").pack(anchor="w")

    fe = tk.LabelFrame(cont, text="Escenario  (la base de sobrecostos va atada a "
                                  "la de retiros)", padx=10, pady=6)
    fe.pack(fill="x", padx=20, pady=4)
    for k in ORDEN_ESC:
        tk.Radiobutton(fe, text=ESCENARIOS[k]["etiqueta"], variable=var_esc,
                       value=k, font=("Consolas", 9)).pack(anchor="w")
    tk.Label(fe, text="Elegir mal significa prorratear contra los retiros de la "
                      "otra base.",
             font=("Segoe UI", 8), fg="#a00000").pack(anchor="w", pady=(4, 0))

    fq = tk.LabelFrame(cont, text="Qué va a hacer", padx=10, pady=6)
    fq.pack(fill="x", padx=20, pady=4)
    for linea in (
            "1.  Borra (si existen): " + ", ".join(TABLAS_A_BORRAR),
            "2.  Copia del Access: Central_Empresa_Actualizada → Central_Empresa,",
            "         Sobrecostos y TIPOS",
            f"3.  Arma {TABLA_REPORTE} desde la vista [{VISTA_PAGO}]"):
        tk.Label(fq, text=linea, font=("Consolas", 9), anchor="w").pack(fill="x")

    fo = tk.Frame(cont)
    fo.pack(fill="x", padx=24, pady=(8, 0))
    var_mirar = tk.BooleanVar(value=True)
    tk.Checkbutton(fo, text="SOLO MIRAR: contar y decir qué haría, sin tocar la "
                            "base",
                   variable=var_mirar, font=("Segoe UI", 9, "bold"),
                   fg="#1a5fb4").pack(anchor="w")

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

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            # El Revisor manda en "ruta_nodo" la ruta de LA FILA desde la que se
            # apretó el botón. Sin eso habría que adivinar cuál de los tres .mdb
            # es, y se usaba siempre el de SSCC aunque se hubiera apretado el de
            # la planilla 9.
            r = traspaso.get("ruta_nodo")
            if r:
                log(f"Botón apretado en: {traspaso.get('nodo')}")
            else:
                # Revisor viejo, sin "ruta_nodo": se cae al orden de siempre.
                r = (d.get("mdb_sscc") or d.get("mdb_ocupar")
                     or d.get("mdb_sob"))
                if r:
                    log("El Revisor no mandó de qué fila se apretó el botón "
                        "(versión antigua): se usa el .mdb de SSCC. Si querías "
                        "otro, elegilo con Examinar.")
            var_mdb.set(r or "[el Revisor no mandó ningún .mdb]")
        else:
            g = cfg.get("prorrateo_mdb")
            if g and Path(g).is_file():
                var_mdb.set(g)
        pintar()

    def lanzar():
        m = var_mdb.get()
        if not m or m.startswith("[") or not Path(m).is_file():
            messagebox.showerror("Falta el Access",
                                 f"No se encontró el .mdb.\n\n{m}")
            return
        esc = var_esc.get()
        if esc not in ESCENARIOS:
            messagebox.showerror("Escenario inválido", esc)
            return
        mirar = var_mirar.get()
        base = ESCENARIOS[esc]["sobrecostos"]
        if not mirar:
            if not messagebox.askyesno(
                    "Confirmar",
                    f"Servidor : {SERVER}\n"
                    f"Base     : {base}\n"
                    f"Retiros  : {ESCENARIOS[esc]['retiros']}\n\n"
                    f"Se BORRAN y se vuelven a cargar:\n"
                    + "\n".join(f"   • {t}" for t in TABLAS_A_BORRAR) +
                    f"\n\nDesde:\n   {Path(m).name}\n\n"
                    "Esto MODIFICA la base de datos. ¿Seguir?"):
                return

        guardar_config({"prorrateo_escenario": esc})
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
            ok, msg = ejecutar(m, esc, mirar, log, progreso)

            def fin():
                timer["on"] = False
                btn.config(state="normal", bg="#2d7a2d")
                if ok and mirar:
                    var_estado.set(f"Solo mirar — {msg}")
                    messagebox.showinfo(
                        "Solo mirar",
                        f"No se tocó la base.\n\n{msg}\n\n"
                        "Si está bien, desmarcá «solo mirar».")
                elif ok:
                    var_estado.set(f"Listo — {msg}")
                    messagebox.showinfo("Listo", f"Prorrateo terminado.\n\n{msg}")
                else:
                    var_estado.set("Terminó con problemas — revisa el log")
                    messagebox.showerror("Problema",
                                         f"{msg}\n\nEl detalle está en el log.")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text="PRORRATEAR", bg="#2d7a2d", fg="white",
                    font=("Segoe UI", 10, "bold"), command=lanzar)
    btn.pack(side="left", padx=8, expand=True)
    tk.Button(frame_btns, text="Refrescar ruta",
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
