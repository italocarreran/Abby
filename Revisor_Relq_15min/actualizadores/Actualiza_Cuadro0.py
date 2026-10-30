# =============================================================================
#  Actualiza Cuadro 0   (0_CUADROS_RELIQUIDACION SSCC)
# =============================================================================
#  ALCANCE: deja los DATOS y las FORMULAS al dia, nada mas.
#      - pega la tabla del 1_CUADROS
#      - la tasa de O2 NO se toca (15 minutos, usuaria 2026-10-29): queda la que
#        tenga el Excel y solo se anota su valor en el log
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
#  que ninguna verificacion posterior salva una tasa equivocada. En 15 minutos
#  el script ya no la pide ni la escribe: se ajusta a mano en el Excel, y el
#  log deja escrito el valor que tenia O2 para poder revisarlo.
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
from tkinter import filedialog, messagebox
from pathlib import Path
import json, sys, re, socket, os, traceback, unicodedata, time


DIR_SCRIPT = Path(__file__).resolve().parent

# config.json es compartido con el Revisor y el resto de los actualizadores,
# y ahora vive en __config__, junto a Revisor_Relq. No es
# DIR_SCRIPT / "config.json" porque este script esta en actualizadores/.
CONFIG_PATH = DIR_SCRIPT.parent.parent / "__config__" / "config.json"

def _morir(titulo, mensaje):
    """Aborta mostrando el motivo en una ventana. Sin esto, lanzado desde el
    Revisor con pythonw (sin consola) el script moriria callado y solo se veria
    que la ventana no aparece. No puede vivir en __comun__/: es justamente lo
    que avisa cuando __comun__/ es lo que falta."""
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(titulo, mensaje)
        r.destroy()
    except Exception:
        pass
    print(f"{titulo}\n\n{mensaje}", file=sys.stderr)
    raise SystemExit(1)

# Implementaciones compartidas; los envoltorios conservan la interfaz historica.
# Va detras de _morir a proposito: lanzado por el Revisor con pythonw no hay
# consola, asi que un ImportError suelto moriria callado.
_RAIZ_COMUN = Path(__file__).resolve().parents[2]
if str(_RAIZ_COMUN) not in sys.path:
    sys.path.insert(0, str(_RAIZ_COMUN))
try:
    from __comun__ import config as _cfg
    from __comun__ import traspaso as _traspaso
    from __comun__ import texto as _texto
except ImportError as e:
    _morir("Falta la carpeta __comun__/",
           "No se pudieron cargar los modulos de __comun__/.\n\n"
           "Tiene que estar la carpeta '__comun__' hermana de Revisor_Relq.\n"
           "Baja el repositorio completo, no los .py sueltos.\n\n"
           f"Carpeta actual: {_RAIZ_COMUN}\n\nDetalle: {e}")

# La ventana chica comun de esta version (actualizadores/ventana_chica.py).
try:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "actualizadores"))
    from ventana_chica import VentanaChica
except ImportError as e:
    _morir("Falta ventana_chica.py",
           "No se pudo cargar actualizadores/ventana_chica.py.\n\n"
           "Copia la carpeta Revisor_Relq_15min completa, no los .py sueltos.\n\n"
           f"Detalle: {e}")


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
get_usuario = _cfg.clave_equipo


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


escribir_json = _cfg.escribir_json


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
    """Corre la secuencia completa. Lo unico opcional es guardar. La tasa (O2)
    no se toca: solo se anota en el log.

    op: {"guardar": bool}
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

        # ---- 2) la tasa: no se toca, solo se anota ------------------------
        log("=" * 70)
        actual = sh3.range(CELDA_TASA).value
        log(f"2) Tasa ({CELDA_TASA}): no se toca, queda la del Excel ({actual!r}).")

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

    v = VentanaChica("Cuadro 0 — 0_CUADROS_RELIQUIDACIÓN",
                     "Pega la tabla del cuadro 1 y ajusta las fórmulas. Este archivo "
                     "se va a pago: queda abierto en Excel para revisarlo.",
                     traspaso, alto=340)
    var_c0 = tk.StringVar(value="[pendiente]")
    var_c1 = tk.StringVar(value="[pendiente]")
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    # No hay casilleros por paso a proposito: los pasos dependen unos de otros
    # (la tabla define K, K define el largo de las formulas, A:G calcula C y C
    # define I), asi que correrlos por separado deja el libro a medio calcular.
    v_guardar = tk.BooleanVar(value=True)

    fr = v.seccion("Archivos")
    v.archivo(fr, "Cuadro 0", var_c0)
    v.archivo(fr, "Cuadro 1", var_c1)
    v.casilla(v.cuerpo, "Guardar al terminar", v_guardar)

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            var_c0.set(d.get("cuadro_0") or "[el Revisor no lo mandó]")
            var_c1.set(d.get("cuadro_1") or "[el Revisor no lo mandó]")
            return
        base = var_carpeta.get()
        if not base or not Path(base).is_dir():
            var_c0.set("[elige la carpeta 02 CASO RELIQUIDACION]")
            var_c1.set("[elige la carpeta 02 CASO RELIQUIDACION]")
            return
        ent = next((p for p in Path(base).iterdir() if p.is_dir()
                    and normalizar(p.name) == normalizar("00 Entregables")), None)

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
        c0, c1 = var_c0.get(), var_c1.get()
        if not c0 or c0.startswith("[") or not Path(c0).is_file():
            messagebox.showerror("Falta el archivo", "No se encontró el cuadro 0.")
            return
        if not c1 or c1.startswith("[") or not Path(c1).is_file():
            messagebox.showerror("Falta el archivo", "Hace falta el cuadro 1 para la tabla.")
            return
        try:
            with open(c0, "r+b"):
                pass
        except PermissionError:
            messagebox.showwarning("El archivo está en uso",
                                   f"Ciérralo en Excel y vuelve a intentar:\n\n{Path(c0).name}")
            return
        except OSError:
            pass
        if not messagebox.askyesno(
                "Confirmar",
                f"Se reemplaza la tabla A5:C de {Path(c0).name} y se ajustan sus "
                "fórmulas" + (" y se guarda" if v_guardar.get() else "") +
                ".\n\nEste archivo se va a pago. ¿Seguir?"):
            return
        v.limpiar_log()
        v.log(f"Cuadro 0 : {c0}")
        v.log(f"Cuadro 1 : {c1}")
        op = {"guardar": v_guardar.get()}
        rutas = {"cuadro0": c0, "cuadro1": c1}

        def fin(r):
            ok, msg, signos_ok = r
            pendientes = ("Ahora, en Excel: Cuadro de pagos, Actualiza Rango, "
                          "refrescar la dinámica de CPRT y Verificar en el Revisor.")
            if ok and not signos_ok:
                v.var_estado.set("Hecho, pero revisa los signos")
                messagebox.showwarning(
                    "Ojo con los signos",
                    f"{msg}\n\nHay montos con el signo al revés en J:K (detalle en el "
                    f"log): arréglalos antes de Cuadro de pagos.\n\n{pendientes}")
            elif ok:
                v.var_estado.set(f"Listo — {msg}")
                messagebox.showinfo("Listo", f"{msg}\n\n{pendientes}")
            else:
                v.var_estado.set("Terminó con errores — revisa el log")
                messagebox.showerror("Error", f"Falló.\n\n{msg}\n\nEl libro quedó "
                                     "abierto donde se cortó.")

        v.correr(lambda: ejecutar(rutas, op, v.log, v.progreso), fin, boton=btn)

    btn = v.boton("ACTUALIZAR", lanzar, principal=True)
    v.boton("Carpeta…", sel_carpeta)
    v.boton("Salir", v.root.destroy)
    refrescar()
    v.iniciar()


if __name__ == "__main__":
    main()
