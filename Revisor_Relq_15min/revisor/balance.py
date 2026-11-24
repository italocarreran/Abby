# -*- coding: utf-8 -*-
"""El balance del caso: AÑO_MES_PROCESO (Revisor 15 min).

Pedido de la usuaria (2026-11-24): el principio de la ventana del Revisor tiene
que quedar como el de la ventana del BESS (``Balance_SSCC/Balance_BESS.py``):
tres casillas que se leen «2026_Enero_Definitivo», la ruta del balance al lado
y los botones «Ir a este balance», «Examinar» y «Bloquear caso».

Lo que cambia por dentro:

* **Una carpeta por balance**, no por mes: el Preliminar y el Definitivo del
  mismo mes son dos casos. Se recuerdan en ``carpetas_por_balance``
  (``{"2601_Definitivo": ruta}``) del config.json de esta version.
  ``carpetas_por_mes`` se sigue escribiendo (lo leen otros scripts), y una
  carpeta anotada solo por mes sirve para el proceso que dice su nombre.
* **El estado de las verificaciones va por balance** (usuaria): en
  ``__config__/AAAA/MM Mes/<Proceso>/``. Un caso sin proceso (lo de antes de
  este cambio) sigue en ``__config__/AAAA/MM Mes/``, donde ya estaba. La
  primera vez que un balance abre la carpeta que ya tenia ese mes, se COPIA
  ahi el estado viejo del mes (no se mueve ni se borra nada).
* **Bloqueo**: como en el BESS, un archivo ``CASO_BLOQUEADO.txt`` en la
  carpeta del caso, asi viaja con ella y lo ve cualquiera que la abra.

Por dentro el Revisor sigue usando el AAMM («2601»), que sale del año y el mes.
Solo biblioteca estandar; sin tkinter, para poder probarlo sin pantalla.
"""

import shutil
import time
from pathlib import Path

ANO_INICIAL = 2025

MESES = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)

# Los que estan siempre; los demas (R01P, R01D, ...) los agrega la usuaria
# con el "+" de la lista y quedan en config.json (clave CLAVE_PROCESOS).
PROCESOS_FIJOS = ("Preliminar", "Definitivo")
OPCION_AGREGAR_PROCESO = "+ Agregar otro..."
CLAVE_PROCESOS = "procesos"
CLAVE_CARPETAS = "carpetas_por_balance"

ARCHIVO_BLOQUEO = "CASO_BLOQUEADO.txt"


# ── Año, mes y proceso ──────────────────────────────────────────────────────

def anos_disponibles(ano_actual=None, extra=()):
    """['2025', '2026', ...] hasta el año de hoy, mas los años de 'extra'
    (los de balances ya recordados) aunque sean de antes de ANO_INICIAL."""
    actual = int(ano_actual or time.localtime().tm_year)
    anos = set(range(ANO_INICIAL, max(actual, ANO_INICIAL) + 1))
    for a in extra:
        try:
            anos.add(int(a))
        except (TypeError, ValueError):
            pass
    return [str(a) for a in sorted(anos)]


def aamm_de(ano, mes):
    """('2026', 'Enero') -> '2601'. '' si falta algo o no es valido."""
    try:
        numero = MESES.index(str(mes)) + 1
        ano = int(ano)
    except (ValueError, TypeError):
        return ""
    return f"{ano % 100:02d}{numero:02d}"


def ano_mes_de(aamm):
    """'2601' -> ('2026', 'Enero'). ('', '') si no es un AAMM."""
    aamm = str(aamm or "").strip()
    if len(aamm) != 4 or not aamm.isdigit() or not 1 <= int(aamm[2:]) <= 12:
        return "", ""
    return str(2000 + int(aamm[:2])), MESES[int(aamm[2:]) - 1]


def nombre_balance(ano, mes, proceso):
    """'2026_Enero_Definitivo' (lo que falte no se escribe)."""
    return "_".join(str(p) for p in (ano, mes, proceso) if p)


def clave_balance(aamm, proceso):
    """'2601_Definitivo'; solo '2601' si no hay proceso."""
    aamm = str(aamm or "").strip()
    return f"{aamm}_{proceso}" if aamm and proceso else aamm


def partir_clave(clave):
    """'2601_Definitivo' -> ('2601', 'Definitivo'); '2601' -> ('2601', '')."""
    aamm, _, proceso = str(clave or "").strip().partition("_")
    return aamm, proceso


def nombre_de_clave(clave):
    """'2601_Definitivo' -> '2026_Enero_Definitivo'."""
    aamm, proceso = partir_clave(clave)
    ano, mes = ano_mes_de(aamm)
    return nombre_balance(ano, mes, proceso) if ano else str(clave)


def limpiar_proceso(texto):
    """El nombre de un proceso agregado a mano: sin espacios ni caracteres
    que no van en un nombre de carpeta. '' si no queda nada."""
    return "".join(c for c in str(texto or "").strip() if c.isalnum() or c == "-")


def procesos_de(cfg):
    """Preliminar, Definitivo y los que agrego este usuario."""
    procesos = list(PROCESOS_FIJOS)
    extra = (cfg or {}).get(CLAVE_PROCESOS)
    if isinstance(extra, list):
        for p in extra:
            p = limpiar_proceso(p)
            if p and p.lower() not in (x.lower() for x in procesos):
                procesos.append(p)
    return procesos


def con_proceso_nuevo(cfg, texto):
    """(proceso, lista_para_guardar_o_None). La lista es None si el proceso
    ya estaba (no hay que escribir nada); proceso '' si el texto no sirve."""
    nuevo = limpiar_proceso(texto)
    if not nuevo:
        return "", None
    procesos = procesos_de(cfg)
    for p in procesos:
        if p.lower() == nuevo.lower():
            return p, None
    return nuevo, [p for p in procesos if p not in PROCESOS_FIJOS] + [nuevo]


def proceso_en_el_nombre(ruta, procesos=PROCESOS_FIJOS):
    """El proceso que dice el nombre de la carpeta, o '' si no dice ninguno
    (o dice mas de uno). Se mira la carpeta y la de arriba: la del caso
    suele llamarse siempre «02 CASO RELIQUIDACION»."""
    p = Path(str(ruta or ""))
    for nombre in (p.name.lower(), p.parent.name.lower()):
        hallados = [x for x in procesos if x.lower() in nombre]
        if len(hallados) == 1:
            return hallados[0]
    return ""


# ── Carpeta de cada balance ─────────────────────────────────────────────────

def carpetas_de(cfg):
    c = (cfg or {}).get(CLAVE_CARPETAS)
    return dict(c) if isinstance(c, dict) else {}


def _existe(ruta):
    try:
        return bool(ruta) and Path(ruta).is_dir()
    except OSError:
        return False


def carpeta_del_balance(cfg, aamm, proceso):
    """La carpeta de ese balance si ya tiene una y todavia existe; '' si no.
    Una carpeta anotada solo por mes (``carpetas_por_mes``, de antes de que
    hubiera proceso) sirve si su nombre dice ese proceso."""
    ruta = carpetas_de(cfg).get(clave_balance(aamm, proceso), "")
    if not ruta and proceso:
        vieja = ((cfg or {}).get("carpetas_por_mes") or {}).get(str(aamm), "")
        if vieja and proceso_en_el_nombre(vieja, procesos_de(cfg)) == proceso:
            ruta = vieja
    return ruta if _existe(ruta) else ""


def misma_carpeta(a, b):
    """True si las dos rutas son la misma carpeta (resolviendo)."""
    try:
        return Path(a).resolve() == Path(b).resolve()
    except OSError:
        return str(a) == str(b)


def balance_de_la_carpeta(cfg, ruta):
    """La clave del balance que ya usa esa carpeta ('2601_Preliminar'), o ''."""
    if not ruta:
        return ""
    for clave, otra in carpetas_de(cfg).items():
        if otra and misma_carpeta(otra, ruta):
            return clave
    return ""


def carpetas_con(cfg, clave, ruta):
    """``carpetas_por_balance`` con 'ruta' anotada para 'clave'. Una carpeta
    es de un solo balance: si otro la tenia, la pierde."""
    carpetas = {k: v for k, v in carpetas_de(cfg).items()
                if k == clave or not misma_carpeta(v, ruta)}
    if clave and ruta:
        carpetas[clave] = str(ruta)
    return carpetas


# ── Estado de las verificaciones, por balance ───────────────────────────────

def dir_balance(dir_config_mes, clave, crear=False):
    """``__config__/AAAA/MM Mes/<Proceso>`` (o la del mes si no hay proceso).
    ``dir_config_mes`` es la del Revisor, que vive en ``__comun__/salidas.py``:
    la carpeta del mes no se arma a mano (ver AGENTS.md)."""
    aamm, proceso = partir_clave(clave)
    base = dir_config_mes(aamm, crear=crear)
    if not proceso:
        return base
    ruta = Path(base) / proceso
    if crear:
        ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def heredar_estado_del_mes(dir_config_mes, clave, archivos):
    """Copia a la carpeta del balance los archivos de estado que haya a nivel
    de mes y que el balance todavia no tenga. No borra ni mueve nada: lo del
    mes queda donde estaba. Devuelve los nombres copiados.

    El que llama decide si corresponde (la carpeta del caso tiene que ser la
    que ese mes ya tenia en ``carpetas_por_mes``)."""
    aamm, proceso = partir_clave(clave)
    if not proceso:
        return []
    origen = Path(dir_config_mes(aamm))
    pendientes = [a for a in archivos
                  if (origen / a).is_file()
                  and not (dir_balance(dir_config_mes, clave) / a).exists()]
    if not pendientes:
        return []
    destino = dir_balance(dir_config_mes, clave, crear=True)
    for a in pendientes:
        shutil.copy2(origen / a, destino / a)
    return pendientes


# ── Bloqueo del caso ────────────────────────────────────────────────────────

def ruta_bloqueo(carpeta_base):
    return Path(carpeta_base) / ARCHIVO_BLOQUEO


def info_bloqueo(carpeta_base):
    """{"quien", "cuando", "motivo"} si el caso esta bloqueado; None si no
    (o si no hay carpeta)."""
    if not carpeta_base:
        return None
    ruta = ruta_bloqueo(carpeta_base)
    try:
        if not ruta.is_file():
            return None
        texto = ruta.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    datos = {"quien": "", "cuando": "", "motivo": ""}
    for linea in texto.splitlines():
        k, _, v = linea.partition(":")
        k = k.strip().lower()
        if k in datos:
            datos[k] = v.strip()
    return datos


def bloquear(carpeta_base, quien, motivo=""):
    """Deja el archivo de bloqueo. Devuelve su ruta."""
    ruta = ruta_bloqueo(carpeta_base)
    ruta.write_text(
        "Caso bloqueado: el Revisor 15 min no lanza actualizadores ni trae "
        "maestros en este caso.\n"
        "Para desbloquearlo, usar el boton 'Desbloquear caso' (o borrar este "
        "archivo).\n"
        f"quien: {quien}\n"
        f"cuando: {time.strftime('%d-%m-%Y %H:%M')}\n"
        f"motivo: {motivo}\n",
        encoding="utf-8",
    )
    return ruta


def desbloquear(carpeta_base):
    """Borra el archivo de bloqueo. True si habia uno."""
    try:
        ruta_bloqueo(carpeta_base).unlink()
    except FileNotFoundError:
        return False
    return True
