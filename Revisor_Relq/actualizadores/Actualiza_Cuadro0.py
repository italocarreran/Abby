# =============================================================================
#  Actualiza Cuadro 0   (0_CUADROS_RELIQUIDACION SSCC)
# =============================================================================
#  ALCANCE: deja los DATOS y las FORMULAS al dia, nada mas.
#      - pega la tabla del 1_CUADROS
#      - la tasa de O2, si se le da una
#      - reescribe las listas de empresas (K de la hoja #3, I de 01.SSCC)
#      - estira o corta las formulas de L:P, R:U, A:G y J:K
#  NO llama a Cuadro de pagos, NO llama a Actualiza Rango y NO refresca la tabla
#  dinamica de CPRT: eso se hace a mano en Excel despues de correr esto. La
#  dinamica cuelga de una Power Query ("Consulta - TEE") que depende de un libro
#  externo, y automatizarla escondia los errores en vez de resolverlos.
#
#  ES EL ARCHIVO QUE SE VA A PAGO, asi que cada paso queda escrito en el log con
#  lo que habia antes y lo que quedo despues.
#
#  COMO ESTA ARMADO EL LIBRO  (lo que hay que entender antes de tocar nada)
#  ----------------------------------------------------------------------------
#  Hoja #3 (la tercera del libro; las dos primeras estan OCULTAS):
#      A:C   tabla pegada del 1_CUADROS  (empresa, paga, recibe)
#      F     el otro listado de empresas
#      K5    lista consolidada, formula desbordada
#      L5    =SUMIF(A:A,K5,D:D)         M5  =SUMIF(F:F,K5,I:I)
#      N5    =L5-M5                     O5  =N5*$O$2   <-- LA TASA
#      P5    =N5+O5
#      Q     VACIA a proposito: separa las dos tablas. No se toca.
#      R5    =+K5                       S5  =IF(P5>0,P5,0)
#      T5    =IF(P5<0,-P5,0)            U5  =S5-T5
#      O2    tasa de interes del periodo
#
#  Hoja 01.SSCC_Recurso_Tecnico:
#      A9    =+'<hoja #3>'!R5     D9=S5   E9=T5   F9=U5   G9=+F9
#      C9    =IFNA(VLOOKUP(A9,EMPRESAS!$H:$I,2,0),A9)     <- aplica reemplazos
#      I9    lista unica de C  (aca se pone la formula nueva de UNICOS)
#      J9    =IF(SUMIF(C,I9,G)<-0.00001, SUMIF(...), "")   lo que paga
#      K9    =IF(SUMIF(C,I9,G)> 0.00001, SUMIF(...), "")   lo que recibe
#      N5+   la matriz que arma la macro CuadroPago
#
#  OJO CON LA TASA: O5 = N5*$O$2 y P5 = N5+O5, y de ahi sale todo el cuadro de
#  pago. Si la tasa esta mal, TODOS los montos estan mal, y el descuadre de
#  CPRT!I3 igual sale chico porque la matriz reparte proporcionalmente. O sea
#  que ninguna verificacion posterior salva una tasa equivocada. Por eso existe
#  la opcion "no actualizar la tasa": es mejor dejar la del Excel que poner una
#  a medias.
#
#  ORDEN DE LOS PASOS: no es arbitrario, es la cadena de dependencias.
#      A:C -> K -> L:P y R:U -> A9:G de 01.SSCC -> C -> I -> J:K -> macros
#  Cambiar el orden da resultados mal calculados sin avisar.
#
#  POR QUE SE LIMPIA ANTES DE ESCRIBIR: para reliquidar se parte de un cuadro de
#  un mes pasado y se pisa. Ese archivo llega con las columnas dimensionadas
#  para el mes anterior, y a veces con valores en vez de formulas. Si el mes
#  nuevo tiene menos empresas, lo que sobra abajo hace que la formula desbordada
#  tire #DESBORDAMIENTO. Por eso cada bloque se limpia hasta su ultima celda con
#  algo ANTES de escribir.
#
#  Se puede correr solo, o recibir del Revisor la ruta de un JSON de traspaso
#  como unico argumento.
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

# --- Hojas y celdas ---------------------------------------------------------
IDX_HOJA_3 = 2                  # 0-based: la TERCERA hoja, contando las ocultas
HOJA_SSCC = "01.SSCC_Recurso_Técnico"
HOJA_ORIGEN_CUADRO1 = "01.SSCC_Recurso_Técnico"

CELDA_TASA = "O2"

# Tabla que se pega del 1_CUADROS: I9:K de su hoja -> A5:C de la hoja #3
ORIGEN_TABLA = ("I", "K")
ORIGEN_TABLA_FILA = 9
DESTINO_TABLA = ("A", "C")
DESTINO_TABLA_FILA = 5

# La D de la hoja #3 es una formula que acompaña a la tabla A:C:
#     D5 = B5 + C5      (el neto por empresa: lo que paga mas lo que recibe)
# O sea que su largo lo manda la TABLA PEGADA, no la lista K. Hay que ajustarla
# aparte, porque la tabla A:C se pega como VALORES y el pegado no la toca.
#
# Importa porque la D es lo que suma L5 (=SUMAR.SI(A:A;K5;D:D)): si la formula se
# queda corta, las ultimas empresas suman 0 y el descuadre aparece recien al
# final; si sobra, arrastra filas del mes anterior.
BLOQUES_TABLA = [("D", "D")]

# --- Formulas que se escriben ----------------------------------------------
# IMPORTANTE: por COM las formulas se escriben en INGLES y con COMA como
# separador, aunque en pantalla se vean en espanol con punto y coma. Escribir
# "UNICOS(...;...)" por esta via falla.
#   K5 = lista consolidada de empresas de F y de A, sin ceros ni vacios
#   I9 = lista unica de C (empresas DESPUES de los reemplazos)
# El LET evita calcular el UNIQUE tres veces. El (x<>0)*(x<>"") saca las dos
# basuras: el 0 que deja VSTACK sobre celdas vacias y el "" que devuelven las
# formulas cuando se les acaban las filas.
FILA_K = 5
FORMULA_K = ('=LET(x,UNIQUE(VSTACK(F{f}:F{tope},A{f}:A{tope})),'
             'FILTER(x,(x<>0)*(x<>"")))')
TOPE_K = 1000

FILA_I = 9
FORMULA_I = '=LET(x,UNIQUE(C{f}:C{tope}),FILTER(x,(x<>0)*(x<>"")))'
TOPE_I = 4345

# --- Bloques de formulas que hay que estirar o cortar ----------------------
# (hoja, primera_fila, [(col_ini, col_fin), ...], que define el largo)
#   "K" -> tantas filas como empresas haya en K de la hoja #3
#   "I" -> tantas filas como empresas haya en I de 01.SSCC
BLOQUES = [
    ("#3",   5, [("L", "P"), ("R", "U")], "K"),   # la Q queda afuera: esta vacia
    ("SSCC", 9, [("A", "G")],             "K"),
    ("SSCC", 9, [("J", "K")],             "I"),
]

# Este script NO llama ninguna macro ni refresca la tabla dinamica: eso queda a
# mano. Solo deja los datos y las formulas al dia.


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


def leer_tasa_guardada(cfg, aamm):
    """(valor, fecha) de la tasa guardada para ese mes, o (None, None).

    Se guarda POR MES a proposito: la tasa es del periodo. Guardada suelta, el
    mes siguiente arrastraria la del anterior sin que se note."""
    if not aamm:
        return None, None
    d = (cfg.get("tasa_interes_por_mes") or {}).get(str(aamm))
    if isinstance(d, dict):
        try:
            return float(d.get("valor")), d.get("fecha")
        except (TypeError, ValueError):
            return None, None
    return None, None


def guardar_tasa(aamm, valor):
    if not aamm:
        return False
    marca = f"{datetime.now():%d-%m-%Y %H:%M}"

    def mut(todo):
        mio = todo.setdefault(get_usuario(), {})
        porm = mio.setdefault("tasa_interes_por_mes", {})
        porm[str(aamm)] = {"valor": float(valor), "fecha": marca}
    return _modificar_config(mut)


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


def buscar_hoja(wb, nombre):
    """Busca la hoja tolerando tildes y mayusculas."""
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


def ultima_fila(sh, col, desde=1):
    """Ultima fila con ALGO en la columna, mirando de abajo hacia arriba.
    Cuenta las cadenas vacias como contenido: para el desbordamiento un "" es
    tan estorbo como un numero."""
    try:
        c = col_num(col)
        f = sh.api.Cells(sh.api.Rows.Count, c).End(-4162).Row   # xlUp
        return max(int(f), desde - 1)
    except Exception:
        return desde - 1


def recalcular(app, log, motivo):
    """Recalcula UNA vez y dice cuanto tardo.

    El libro se queda en calculo MANUAL de punta a punta. Con automatico, cada
    AutoFill dispara un recalculo completo, y las formulas de aca son caras: L5
    es =SUMIF(A:A,K5,D:D), un SUMIF de COLUMNA ENTERA repetido por cada empresa,
    y J9/K9 evaluan su SUMIF dos veces (una en la condicion y otra en el
    resultado). Recalcular cuatro veces a proposito, en vez de una por cada
    escritura, es la diferencia entre segundos y minutos.

    Se cronometra y queda en el log: si algun mes se pone lento, se ve donde.
    """
    t0 = time.time()
    app.calculate()
    log(f"    recalculo ({motivo}): {time.time() - t0:.1f} s")


def fmt_tiempo(seg):
    seg = int(seg)
    return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"


# =============================================================================
#  PASOS
# =============================================================================
def paso_pegar_tabla(sh1, sh3, log):
    """I9:K del cuadro 1 -> A5:C de la hoja #3. Solo valores."""
    c1, c2 = (col_num(x) for x in ORIGEN_TABLA)
    ult = max(ultima_fila(sh1, ORIGEN_TABLA[0], ORIGEN_TABLA_FILA),
              ultima_fila(sh1, ORIGEN_TABLA[1], ORIGEN_TABLA_FILA))
    if ult < ORIGEN_TABLA_FILA:
        raise RuntimeError(f"El 1_CUADROS no tiene datos en "
                           f"{ORIGEN_TABLA[0]}{ORIGEN_TABLA_FILA} hacia abajo.")
    datos = sh1.range((ORIGEN_TABLA_FILA, c1), (ult, c2)).options(ndim=2).value
    # Se descartan las filas sin empresa (la cola de vacios del final).
    datos = [f for f in datos if f and f[0] is not None and str(f[0]).strip()]
    if not datos:
        raise RuntimeError("La tabla del 1_CUADROS quedo vacia despues de "
                           "descartar las filas sin empresa.")
    log(f"    origen : {sh1.name}!{ORIGEN_TABLA[0]}{ORIGEN_TABLA_FILA}:"
        f"{ORIGEN_TABLA[1]}{ult}  ->  {len(datos)} empresa(s)")

    d1, d2 = (col_num(x) for x in DESTINO_TABLA)
    previa = max(ultima_fila(sh3, DESTINO_TABLA[0], DESTINO_TABLA_FILA),
                 ultima_fila(sh3, DESTINO_TABLA[1], DESTINO_TABLA_FILA))
    if previa >= DESTINO_TABLA_FILA:
        log(f"    limpiando {DESTINO_TABLA[0]}{DESTINO_TABLA_FILA}:"
            f"{DESTINO_TABLA[1]}{previa}  ({previa - DESTINO_TABLA_FILA + 1} "
            f"fila(s) del mes anterior)")
        sh3.range((DESTINO_TABLA_FILA, d1), (previa, d2)).clear_contents()
    fin = DESTINO_TABLA_FILA + len(datos) - 1
    sh3.range((DESTINO_TABLA_FILA, d1)).value = datos
    log(f"    pegado : {DESTINO_TABLA[0]}{DESTINO_TABLA_FILA}:"
        f"{DESTINO_TABLA[1]}{fin}")
    if previa > fin:
        log(f"    (el mes anterior llegaba a la fila {previa}: "
            f"{previa - fin} fila(s) quedaron limpias)")
    return len(datos)


def paso_tasa(sh3, valor, log):
    anterior = sh3.range(CELDA_TASA).value
    sh3.range(CELDA_TASA).value = float(valor)
    log(f"    {CELDA_TASA}: {anterior!r}  ->  {float(valor)!r}")
    log("    OJO: la tasa multiplica cada fila (O5=N5*$O$2). Si esta mal, todos")
    log("         los montos quedan mal y el descuadre de CPRT!I3 igual sale chico.")


def paso_formula_desbordada(sh, celda, formula, col, fila, log, etiqueta):
    """Limpia la columna y escribe la formula. Devuelve la ultima fila ocupada.

    Limpiar PRIMERO es obligatorio: cualquier celda no vacia debajo hace que la
    formula desbordada tire #DESBORDAMIENTO, y al reliquidar el archivo llega
    con los datos del mes anterior."""
    previa = ultima_fila(sh, col, fila)
    if previa >= fila:
        log(f"    limpiando {col}{fila}:{col}{previa}  "
            f"({previa - fila + 1} celda(s) del mes anterior)")
        sh.range(f"{col}{fila}:{col}{previa}").clear_contents()
    else:
        log(f"    la columna {col} estaba vacia")
    log(f"    escribiendo en {celda}: {formula}")
    r = sh.range(celda)
    try:
        r.formula2 = formula          # la via correcta para matrices dinamicas
    except Exception:
        r.formula = formula
    # Aca SI hace falta recalcular: sin el resultado no se sabe hasta donde
    # desbordo la formula, y ese largo manda todo lo que viene despues.
    recalcular(sh.book.app, log, f"resultado de {celda}")
    ult = ultima_fila(sh, col, fila)
    val = r.value
    if isinstance(val, str) and val.startswith("#"):
        raise RuntimeError(f"{etiqueta}: la formula quedo en error ({val}). "
                           f"Si dice DESBORDAMIENTO, quedo algo debajo de {celda}.")
    n = ult - fila + 1
    log(f"    {etiqueta}: {n} empresa(s), {col}{fila}:{col}{ult}")
    if n <= 0:
        raise RuntimeError(f"{etiqueta}: la formula no devolvio ninguna empresa.")
    return ult


def paso_estirar(sh, fila_ini, bloques, hasta, log):
    """Deja las formulas de esos bloques cubriendo exactamente hasta 'hasta'."""
    for c_ini, c_fin in bloques:
        n1, n2 = col_num(c_ini), col_num(c_fin)
        previa = max(ultima_fila(sh, c, fila_ini) for c in (c_ini, c_fin))
        etiqueta = f"{c_ini}:{c_fin}" if c_ini != c_fin else c_ini
        if previa > hasta:
            sh.range((hasta + 1, n1), (previa, n2)).clear_contents()
            log(f"    {etiqueta}: cortado, sobraban {previa - hasta} fila(s) "
                f"(iba hasta {previa}, ahora hasta {hasta})")
        if hasta > fila_ini:
            # AutoFill copia las formulas con las referencias bien corridas.
            origen = sh.range((fila_ini, n1), (fila_ini, n2))
            destino = sh.range((fila_ini, n1), (hasta, n2))
            origen.api.AutoFill(destino.api, 0)     # xlFillDefault
        if previa < hasta:
            log(f"    {etiqueta}: estirado, faltaban {hasta - previa} fila(s) "
                f"(iba hasta {previa}, ahora hasta {hasta})")
        elif previa == hasta:
            log(f"    {etiqueta}: ya estaba justo hasta la fila {hasta}")


def paso_validar_signos(sh, log):
    """AVISO, no corta. Comprueba lo mismo que valida CuadroPago antes de armar
    la matriz: en la tabla I:K, todo numero de la columna J (paga) tiene que ser
    NEGATIVO y todo numero de la K (recibe) POSITIVO.

    Se avisa aca porque CuadroPago ATRAPA su propio error: muestra un MsgBox y
    sale con Exit Sub sin escribir la matriz. Corriendola a mano se ve el cartel,
    pero conviene saber de antemano si va a fallar. Devuelve True si esta OK.
    """
    ult = max(ultima_fila(sh, "J", FILA_I), ultima_fila(sh, "K", FILA_I))
    if ult < FILA_I:
        log(f"    no hay datos en J{FILA_I}:K{FILA_I} hacia abajo, no se puede "
            "revisar los signos")
        return False
    datos = sh.range((FILA_I, col_num("J")), (ult, col_num("K"))).options(ndim=2).value
    malos_j, malos_k = [], []
    n_j = n_k = 0
    for i, fila in enumerate(datos or []):
        j = fila[0] if len(fila) > 0 else None
        k = fila[1] if len(fila) > 1 else None
        if isinstance(j, (int, float)) and not isinstance(j, bool):
            n_j += 1
            if j >= 0:
                malos_j.append((FILA_I + i, j))
        if isinstance(k, (int, float)) and not isinstance(k, bool):
            n_k += 1
            if k <= 0:
                malos_k.append((FILA_I + i, k))
    log(f"    J (paga): {n_j} monto(s)   |   K (recibe): {n_k} monto(s)")
    if malos_j or malos_k:
        for f, v in (malos_j + malos_k)[:10]:
            col = "J" if (f, v) in malos_j else "K"
            log(f"      {col}{f} = {v}   <- signo al reves")
        log(f"    OJO: {len(malos_j)} monto(s) de J que no son negativos y "
            f"{len(malos_k)} de K que no son positivos.")
        log("         Con esto CuadroPago se va a cortar. Revisa la tabla I:K")
        log("         ANTES de correr la macro.")
        return False
    log("    signos OK: todo J negativo y todo K positivo")
    log("    (o sea que CuadroPago no deberia quejarse por los signos)")
    return True


def ejecutar(rutas, op, log, progreso):
    """Corre la secuencia completa. Lo unico opcional es la tasa, porque puede
    llegar al final del mes, y guardar.

    op: {"tasa": float o None para no tocarla, "guardar": bool}
    Devuelve (ok, resumen)."""
    try:
        import pythoncom
        import xlwings as xw
    except ImportError as e:
        log(f"ERROR: falta una libreria: {e}")
        return False, str(e), True

    pythoncom.CoInitialize()
    app = None
    wb0 = wb1 = None
    resumen = []
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        log(f"  Abriendo el cuadro 0: {Path(rutas['cuadro0']).name}")
        wb0 = app.books.open(str(rutas["cuadro0"]), update_links=False)
        if len(wb0.sheets) <= IDX_HOJA_3:
            raise RuntimeError(f"El libro tiene {len(wb0.sheets)} hoja(s); "
                               f"se esperaba la #{IDX_HOJA_3 + 1}.")
        sh3 = wb0.sheets[IDX_HOJA_3]
        sh_sscc = buscar_hoja(wb0, HOJA_SSCC)
        if sh_sscc is None:
            raise RuntimeError(f"No esta la hoja '{HOJA_SSCC}'.")
        log(f"    hoja #{IDX_HOJA_3 + 1} = '{sh3.name}'   |   '{sh_sscc.name}'")
        wb0.app.calculation = "manual"

        # ---- 1) la tabla del cuadro 1 ------------------------------------
        progreso(10, 100, "Pegando la tabla del 1_CUADROS...")
        log("=" * 70)
        log("1) Tabla del 1_CUADROS -> A5:C de la hoja #3")
        wb1 = app.books.open(str(rutas["cuadro1"]), read_only=True,
                             update_links=False)
        sh1 = buscar_hoja(wb1, HOJA_ORIGEN_CUADRO1)
        if sh1 is None:
            raise RuntimeError(f"El 1_CUADROS no tiene la hoja "
                               f"'{HOJA_ORIGEN_CUADRO1}'.")
        n = paso_pegar_tabla(sh1, sh3, log)
        resumen.append(f"tabla: {n} empresas")
        wb1.close()
        wb1 = None

        # La D acompaña a la tabla que se acaba de pegar.
        hasta_tabla = DESTINO_TABLA_FILA + n - 1
        log(f"    -- {sh3.name}, la fórmula que sigue a la tabla "
            f"({n} empresas -> hasta la fila {hasta_tabla})")
        paso_estirar(sh3, DESTINO_TABLA_FILA, BLOQUES_TABLA, hasta_tabla, log)

        # ---- 2) la tasa ---------------------------------------------------
        log("=" * 70)
        if op["tasa"] is None:
            actual = sh3.range(CELDA_TASA).value
            log(f"2) Tasa: NO se toca. Queda la del Excel ({actual!r}).")
            resumen.append("tasa: sin cambios")
        else:
            log("2) Tasa de interes del periodo")
            paso_tasa(sh3, op["tasa"], log)
            resumen.append(f"tasa: {op['tasa']}")

        # ---- 3) las formulas ---------------------------------------------
        progreso(30, 100, "Reescribiendo las listas de empresas...")
        log("=" * 70)
        log("3) Listas de empresas y largo de las formulas")
        # Se queda en MANUAL. Se recalcula solo en los puntos donde el resultado
        # se necesita para seguir, no despues de cada escritura.
        ult_k = paso_formula_desbordada(
            sh3, f"K{FILA_K}",
            FORMULA_K.format(f=FILA_K, tope=TOPE_K),
            "K", FILA_K, log, "K (lista consolidada)")
        n_k = ult_k - FILA_K + 1

        for hoja, fila, bloques, ref in BLOQUES:
            sh = sh3 if hoja == "#3" else sh_sscc
            if ref == "K":
                hasta = fila + n_k - 1
                log(f"    -- {sh.name}, bloques que siguen a K "
                    f"({n_k} empresas -> hasta la fila {hasta})")
                # Sin recalcular todavia: se hace una sola vez al salir del
                # bucle, porque I necesita que C este calculada.
                paso_estirar(sh, fila, bloques, hasta, log)
        recalcular(wb0.app, log, "columna C, que es la que alimenta a I")

        # I depende de C, que ya quedo calculada arriba
        ult_i = paso_formula_desbordada(
            sh_sscc, f"I{FILA_I}",
            FORMULA_I.format(f=FILA_I, tope=TOPE_I),
            "I", FILA_I, log, "I (empresas tras reemplazos)")
        n_i = ult_i - FILA_I + 1
        if n_i > n_k:
            log(f"    OJO: I tiene {n_i} empresas y K tiene {n_k}. Despues de")
            log("         aplicar reemplazos deberia haber MENOS o iguales.")

        for hoja, fila, bloques, ref in BLOQUES:
            if ref != "I":
                continue
            sh = sh3 if hoja == "#3" else sh_sscc
            hasta = fila + n_i - 1
            log(f"    -- {sh.name}, bloques que siguen a I "
                f"({n_i} empresas -> hasta la fila {hasta})")
            paso_estirar(sh, fila, bloques, hasta, log)
        recalcular(wb0.app, log, "final, con todo estirado")

        resumen.append(f"K: {n_k} empresas, I: {n_i}")

        # ---- 4) las macros y la dinamica ---------------------------------
        # Recien al final vuelve a automatico, para que el libro quede como el
        # usuario lo espera. Sin CalculateFullRebuild: ese reconstruye todo el
        # arbol de dependencias y es el recalculo mas caro que hay; solo sirve si
        # el arbol quedo corrupto, y aca no hay motivo para eso.
        wb0.app.calculation = "automatic"
        # Chequeo de signos: no se arma la matriz aca, pero conviene saber si
        # CuadroPago se va a cortar cuando la corras.
        log("=" * 70)
        log("Antes de que corras CuadroPago: revision de signos en J:K")
        signos_ok = paso_validar_signos(sh_sscc, log)

        log("=" * 70)
        log("LISTO lo que hace este script. Te queda por hacer, en este orden:")
        log("   1. el boton Cuadro de pagos")
        log("   2. el boton Actualiza Rango")
        log("   3. refrescar la tabla dinamica de CPRT")
        log("   4. y despues Verificar en el Revisor")
        if not signos_ok:
            log("   OJO: arregla primero los signos de J:K, ver arriba.")

        if op["guardar"]:
            log("  Guardando...")
            wb0.save()
            log("  Guardado.")
            resumen.append("guardado")
        else:
            log("  NO se guardo (asi lo pediste). El libro queda abierto.")

        # Queda visible y abierto a proposito, para revisar antes de mandarlo.
        app.api.Visible = True
        app.screen_updating = True
        app.display_alerts = True
        sh_sscc.activate()
        wb0 = None
        app = None
        return True, " | ".join(resumen), signos_ok

    except Exception as e:
        log(f"\nERROR: {e}")
        log(traceback.format_exc())
        try:
            if wb1 is not None:
                wb1.close()
        except Exception:
            pass
        # El cuadro 0 NO se cierra ni se guarda si algo fallo: queda visible
        # para ver en que paso quedo.
        try:
            if app is not None:
                app.api.Visible = True
                app.screen_updating = True
                app.display_alerts = True
        except Exception:
            pass
        return False, str(e), True
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
    root.title("Actualiza Cuadro 0"
               + ("  —  enviado por el Revisor" if traspaso else ""))
    root.geometry("1000x640")

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

    aamm_ini = (traspaso or {}).get("aamm") or cfg.get("ultimo_mes", "")

    if traspaso:
        fa = tk.Frame(cont, bg="#fff4c2", bd=1, relief="solid")
        fa.pack(fill="x", padx=20, pady=(10, 0))
        tk.Label(fa, text=f"Mes {traspaso.get('aamm') or '?'} — enviado por el Revisor",
                 bg="#fff4c2", font=("Segoe UI", 11, "bold")).pack(pady=(6, 0))
        tk.Label(fa, text="Las rutas las resolvió el Revisor.\n"
                          "Lo único que hay que decidir es la tasa.",
                 bg="#fff4c2", font=("Segoe UI", 8), fg="#444444",
                 justify="center").pack(pady=(0, 6))

    tk.Label(cont, text="Actualizar el 0_CUADROS_RELIQUIDACIÓN",
             font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
    tk.Label(cont, text="Este archivo se va a pago. Al terminar queda abierto en "
                        "Excel para que lo revises.",
             font=("Segoe UI", 8), fg="#a00000").pack(pady=(0, 8))

    var_c0 = tk.StringVar(value="[pendiente]")
    var_c1 = tk.StringVar(value="[pendiente]")
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    var_aamm = tk.StringVar(value=aamm_ini)
    labels = {}

    def fila(parent, titulo, var, clave):
        fr = tk.LabelFrame(parent, text=titulo, padx=10, pady=6)
        fr.pack(fill="x", padx=20, pady=4)
        l = tk.Label(fr, textvariable=var, wraplength=880, justify="left",
                     cursor="hand2", font=("Segoe UI", 9), anchor="w")
        l.pack(fill="x")
        l.bind("<Button-1>", lambda e: abrir_en_explorador(var.get(), es_archivo=True))
        labels[clave] = l
        return fr

    fila(cont, "Cuadro 0 — 0_CUADROS_RELIQUIDACIÓN (se modifica)", var_c0, "c0")
    fila(cont, "Cuadro 1 — 1_CUADROS_PAGO_SSCC (solo se lee)", var_c1, "c1")

    # ---- opciones ----
    # No hay casilleros por paso a proposito: los pasos dependen unos de otros
    # (la tabla define K, K define el largo de las formulas, A:G calcula C y C
    # define I), asi que correrlos por separado deja el libro a medio calcular.
    # La secuencia completa esta documentada en el encabezado de este archivo y
    # queda escrita en el log en cada corrida.
    v_guardar = tk.BooleanVar(value=True)
    tk.Checkbutton(cont, text="Guardar al terminar (queda abierto en Excel igual)",
                   variable=v_guardar,
                   font=("Segoe UI", 9)).pack(anchor="w", padx=24, pady=(6, 0))

    # ---- tasa ----
    ft = tk.LabelFrame(cont, text=f"Tasa de interés del periodo  (celda "
                                  f"{CELDA_TASA} de la hoja #3)", padx=10, pady=6)
    ft.pack(fill="x", padx=20, pady=4)
    v_no_tasa = tk.BooleanVar(value=True)
    v_tasa = tk.StringVar(value="")
    var_tasa_info = tk.StringVar(value="")

    fr_t = tk.Frame(ft)
    fr_t.pack(fill="x")
    tk.Label(fr_t, text="Tasa:", font=("Segoe UI", 9)).pack(side="left")
    ent_tasa = tk.Entry(fr_t, textvariable=v_tasa, width=14, font=("Consolas", 10))
    ent_tasa.pack(side="left", padx=6)
    tk.Label(fr_t, textvariable=var_tasa_info, font=("Segoe UI", 8),
             fg="#2d7a2d").pack(side="left", padx=8)
    tk.Label(ft, text="Se escribe como fracción, igual que en el Excel: 0,168 va "
                      "como 0.168",
             font=("Segoe UI", 8), fg="#555555").pack(anchor="w")

    def _toggle_tasa(*_):
        ent_tasa.config(state="disabled" if v_no_tasa.get() else "normal")
    tk.Checkbutton(ft, text="NO actualizar la tasa: dejar la que ya tiene el Excel",
                   variable=v_no_tasa, font=("Segoe UI", 9),
                   command=_toggle_tasa).pack(anchor="w", pady=(4, 0))
    tk.Label(ft, text="La tasa multiplica cada fila del cuadro. Si te la dan al "
                      "final, dejá esta casilla marcada:\nmejor la del Excel que "
                      "una a medias, porque ninguna verificación posterior la caza.",
             font=("Segoe UI", 8), fg="#a00000", justify="left").pack(anchor="w")

    def refrescar_tasa(*_):
        val, fecha = leer_tasa_guardada(cfg, var_aamm.get().strip())
        if val is not None:
            var_tasa_info.set(f"guardada para {var_aamm.get().strip()}: {val}"
                              + (f"   ({fecha})" if fecha else ""))
            if not v_tasa.get():
                v_tasa.set(str(val))
        else:
            var_tasa_info.set("no hay tasa guardada para este mes")

    # ---- mes ----
    fm = tk.Frame(cont)
    fm.pack(fill="x", padx=24, pady=(4, 0))
    tk.Label(fm, text="Mes (AAMM):", font=("Segoe UI", 9)).pack(side="left")
    ent_aamm = tk.Entry(fm, textvariable=var_aamm, width=8, font=("Consolas", 10))
    ent_aamm.pack(side="left", padx=6)
    tk.Label(fm, text="define con qué mes se guarda la tasa en config.json",
             font=("Segoe UI", 8), fg="#555555").pack(side="left")
    var_aamm.trace_add("write", lambda *a: refrescar_tasa())

    # ---- carpeta ----
    fc = tk.LabelFrame(cont, text="Carpeta 02 CASO RELIQUIDACION", padx=10, pady=6)
    fc.pack(fill="x", padx=20, pady=4)
    lc = tk.Label(fc, textvariable=var_carpeta, wraplength=880, justify="left",
                  cursor="hand2", font=("Segoe UI", 9), anchor="w")
    lc.pack(fill="x")
    lc.bind("<Button-1>", lambda e: abrir_en_explorador(var_carpeta.get()))
    labels["carpeta"] = lc

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
    txt = tk.Text(fl, height=16, font=("Consolas", 9), wrap="none")
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

    # ---- rutas ----
    def pintar():
        for k in ("c0", "c1"):
            v = labels[k].cget("text")
        for k, var in (("c0", var_c0), ("c1", var_c1)):
            v = var.get()
            ok = bool(v) and not v.startswith("[")
            if ok:
                try:
                    ok = Path(v).is_file()
                except Exception:
                    ok = False
            labels[k].config(fg="blue" if ok else "red")
        v = var_carpeta.get()
        labels["carpeta"].config(
            fg="blue" if v and Path(v).is_dir() else "red")

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            var_c0.set(d.get("cuadro_0") or "[el Revisor no mandó cuadro_0]")
            var_c1.set(d.get("cuadro_1") or "[el Revisor no mandó cuadro_1]")
            pintar()
            return
        base = var_carpeta.get()
        if not base or not Path(base).is_dir():
            pintar()
            return
        ent = None
        for p in Path(base).iterdir():
            if p.is_dir() and normalizar(p.name) == normalizar("00 Entregables"):
                ent = p
                break
        def buscar(pref):
            if ent is None:
                return None
            cand = [p for p in ent.glob("*")
                    if p.is_file() and normalizar(p.name).startswith(pref)
                    and p.suffix.lower() in (".xlsm", ".xlsx")
                    and not p.name.startswith("~$")]
            return max(cand, key=lambda p: p.stat().st_mtime) if cand else None
        c0, c1 = buscar("0_cuadros"), buscar("1_cuadros")
        var_c0.set(str(c0) if c0 else "[0_CUADROS no encontrado]")
        var_c1.set(str(c1) if c1 else "[1_CUADROS no encontrado]")
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

    # ---- lanzar ----
    def lanzar():
        c0 = var_c0.get()
        if not c0 or c0.startswith("[") or not Path(c0).is_file():
            messagebox.showerror("Falta el archivo",
                                 f"No se encontró el cuadro 0.\n\n{c0}")
            return
        c1 = var_c1.get()
        if not c1 or c1.startswith("[") or not Path(c1).is_file():
            messagebox.showerror("Falta el archivo",
                                 f"Hace falta el cuadro 1 para la tabla."
                                 f"\n\n{c1}")
            return
        # Tiene que estar cerrado: se abre para escribir.
        p = Path(c0)
        try:
            with open(p, "r+b"):
                pass
        except PermissionError:
            messagebox.showwarning("El archivo está en uso",
                                   f"Ciérralo en Excel y volvé a intentar:\n\n{p.name}")
            return
        except OSError:
            pass

        tasa = None
        if not v_no_tasa.get():
            crudo = v_tasa.get().strip().replace(",", ".")
            try:
                tasa = float(crudo)
            except ValueError:
                messagebox.showerror("Tasa inválida",
                                     f"No pude leer {v_tasa.get()!r} como número.")
                return
            if not (0 <= tasa < 3):
                if not messagebox.askyesno(
                        "Tasa fuera de lo normal",
                        f"La tasa quedaría en {tasa}, que se sale de lo habitual "
                        f"(se escribe como fracción: 0.168 = 16,8%).\n\n¿Seguir?"):
                    return

        aamm = var_aamm.get().strip()
        que = ["  • pegar la tabla del cuadro 1 (se REEMPLAZA A5:C)",
               (f"  • poner la tasa en {tasa}" if tasa is not None
                else "  • la tasa NO se toca: queda la del Excel"),
               "  • reescribir K5 e I9 y ajustar el largo de las fórmulas",
               "  • Cuadro de pagos, Actualiza Rango y refrescar la dinámica"]
        if v_guardar.get():
            que.append("  • GUARDAR el archivo al terminar")
        if not messagebox.askyesno(
                "Confirmar",
                f"Sobre {Path(c0).name}:\n\n" + "\n".join(que) +
                "\n\nEste archivo se va a pago. ¿Seguir?"):
            return

        if tasa is not None and aamm:
            guardar_tasa(aamm, tasa)
            cfg.setdefault("tasa_interes_por_mes", {})[aamm] = {
                "valor": tasa, "fecha": f"{datetime.now():%d-%m-%Y %H:%M}"}
            refrescar_tasa()

        txt.delete("1.0", "end")
        log(f"Cuadro 0 : {c0}")
        log(f"Cuadro 1 : {c1}")
        if modo["traspaso"]:
            log(f"Rutas enviadas por el Revisor — mes {traspaso.get('aamm') or '?'}")
        log("-" * 70)

        btn.config(state="disabled", bg="#aaaaaa")
        barra["value"] = 0
        var_estado.set("Procesando...")
        timer["on"] = True
        timer["t0"] = time.time()
        tick()

        op = {"tasa": tasa, "guardar": v_guardar.get()}
        rutas = {"cuadro0": c0, "cuadro1": c1}

        def trabajo():
            ok, msg, signos_ok = ejecutar(rutas, op, log, progreso)

            def fin():
                timer["on"] = False
                btn.config(state="normal", bg="#2d7a2d")
                pendientes = ("Ahora, en Excel:\n"
                              "   1. Cuadro de pagos\n"
                              "   2. Actualiza Rango\n"
                              "   3. refrescar la dinámica de CPRT\n"
                              "   4. y después Verificar en el Revisor")
                if ok and not signos_ok:
                    var_estado.set(f"Hecho, pero revisá los signos — {msg}")
                    messagebox.showwarning(
                        "Datos al día, pero ojo con los signos",
                        f"Los datos y las fórmulas quedaron al día.\n\n{msg}\n\n"
                        "Pero hay montos con el signo al revés en J:K, y con eso "
                        "Cuadro de pagos se va a cortar.\n"
                        "Arreglalos primero; el detalle está en el log.\n\n"
                        + pendientes)
                elif ok:
                    var_estado.set(f"Listo — {msg}")
                    messagebox.showinfo(
                        "Listo",
                        f"Datos y fórmulas al día.\n\n{msg}\n\n"
                        "El libro quedó abierto en Excel.\n\n" + pendientes)
                else:
                    var_estado.set("Terminó con errores — revisa el log")
                    messagebox.showerror(
                        "Error", f"Falló.\n\n{msg}\n\nEl libro quedó abierto "
                                 "en el punto donde se cortó.")
            root.after(0, fin)

        threading.Thread(target=trabajo, daemon=True).start()

    btn = tk.Button(frame_btns, text="ACTUALIZAR CUADRO 0", bg="#2d7a2d", fg="white",
                    font=("Segoe UI", 10, "bold"), command=lanzar)
    btn.pack(side="left", padx=8, expand=True)
    tk.Button(frame_btns, text="Refrescar rutas", command=refrescar).pack(side="left", padx=8)
    tk.Button(frame_btns, text="Salir", command=root.destroy).pack(side="left", padx=8)

    refrescar()
    refrescar_tasa()
    _toggle_tasa()
    bombear()
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    root.geometry(f"+{(root.winfo_screenwidth() - w) // 2}"
                  f"+{(root.winfo_screenheight() - h) // 2}")
    root.mainloop()


if __name__ == "__main__":
    main()
