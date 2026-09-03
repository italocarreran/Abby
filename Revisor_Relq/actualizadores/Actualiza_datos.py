# ── Constantes configurables ────────────────────────────────────────────────
INSTRUCCIONES = (
    "Selecciona la carpeta '02 CASO RELIQUIDACION'.\n"
    "El script detecta automáticamente todos los archivos origen y destino.\n"
    "Los .xlsm destino deben estar CERRADOS antes de ejecutar."
)

# ── Imports ─────────────────────────────────────────────────────────────────
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import json, subprocess, sys, re, socket, os, traceback, unicodedata, time

# ── Config por usuario/PC ───────────────────────────────────────────────────
# config.json es compartido con el Revisor y el resto de los actualizadores,
# que viven un nivel arriba (en scripts/, junto al Revisor). No es
# Path(__file__).parent porque este script esta en actualizadores/.
CONFIG_PATH = Path(__file__).resolve().parent.parent.parent / "__config__" / "config.json"

def get_usuario() -> str:
    try:
        usuario = os.environ.get("USERNAME") or os.environ.get("USER") or "desconocido"
        return f"{socket.gethostname()}_{usuario}"
    except Exception:
        return "desconocido"

def leer_config() -> dict:
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get(get_usuario(), {})
    except Exception:
        pass
    return {}

def escribir_json(ruta, data):
    """Escritura atomica: primero un .tmp y despues os.replace.
    Evita dejar el archivo truncado si algo falla a medio camino."""
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_suffix(ruta.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, ruta)

def _modificar_config(mutador):
    """config.json lo comparten el Revisor y los dos actualizadores. Solo se
    agregan o actualizan claves, nunca se borra nada, y si el archivo existe
    pero no se puede interpretar NO se escribe (mejor perder un ajuste que el
    archivo entero)."""
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

def guardar_config(data: dict):
    return _modificar_config(
        lambda todo: todo.setdefault(get_usuario(), {}).update(data))

# ── Traspaso desde el Revisor ───────────────────────────────────────────────
# El Revisor escribe un JSON en Salidas/AAMM/ y pasa su ruta como argv[1].
# Sin argumento el script funciona como siempre: busca los archivos solo.
TRASPASO_ORIGEN = "Revisor_Reliquidacion"
TRASPASO_VERSION_MAX = 1

def leer_traspaso(argv: list) -> dict | None:
    """Devuelve el dict del traspaso, o None si no vino o no es valido.
    Nunca lanza: si el JSON esta roto se cae al modo manual."""
    if len(argv) < 2 or not argv[1].strip():
        return None
    ruta = Path(argv[1].strip())
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

# ── Utilidades ──────────────────────────────────────────────────────────────
def abrir_en_explorador(ruta: str, es_archivo: bool = False):
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

def normalizar(texto: str) -> str:
    nfkd = unicodedata.normalize("NFKD", texto)
    sin_tildes = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", sin_tildes).strip().lower()

# ── Mapeos ───────────────────────────────────────────────────────────────────
# fila_det_origen: columna (letra) usada para detectar última fila en origen
# fila_det_destino: columna (letra) usada para detectar última fila en destino

MAPEO_SOBRECOSTOS_FD = [
    {
        "hoja_origen":    "CT Diario",
        "cols_origen":    [("D", "I")],
        "fila_ini_origen": 12,
        "fila_det_origen": "D",
        "hoja_destino":   "FD_CT",
        "cols_destino":   [("C", "H")],
        "fila_ini_destino": 12,
        "fila_det_destino": "C",
        "cols_formulas":  [("B", "B")],
    },
    {
        "hoja_origen":    "CPF Horario",
        "cols_origen":    [("B", "J")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CPF",
        "cols_destino":   [("C", "K")],
        "fila_ini_destino": 12,
        "fila_det_destino": "C",
        "cols_formulas":  [("B", "B"), ("L", "P")],
    },
    {
        "hoja_origen":    "CSF Horario",
        "cols_origen":    [("B", "H")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CSF",
        "cols_destino":   [("C", "I")],
        "fila_ini_destino": 12,
        "fila_det_destino": "C",
        "cols_formulas":  [("A", "B"), ("J", "M")],
    },
    {
        "hoja_origen":    "CTF Horario",
        "cols_origen":    [("B", "I")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CTF",
        "cols_destino":   [("C", "J")],
        "fila_ini_destino": 12,
        "fila_det_destino": "C",
        "cols_formulas":  [("A", "B"), ("K", "N")],
    },
    {
        "hoja_origen":    "CSF Horario",
        "cols_origen":    [("B", "D"), ("F", "F")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CSF_Disponibilidad",
        "cols_destino":   [("B", "E")],
        "fila_ini_destino": 12,
        "fila_det_destino": "B",
        "cols_formulas":  [("A", "A"), ("F", "H")],
    },
    {
        "hoja_origen":    "CTF Horario",
        "cols_origen":    [("B", "D"), ("G", "G")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CTF_Disponibilidad",
        "cols_destino":   [("B", "E")],
        "fila_ini_destino": 12,
        "fila_det_destino": "B",
        "cols_formulas":  [("A", "A"), ("F", "H")],
    },
]

MAPEO_SOBRECOSTOS_CONSOLIDADO = [
    {
        "hoja_origen":    "Sobrecostos",
        "cols_origen":    [("A", "G"), ("I", "J"), ("Q", "W")],
        "fila_ini_origen": 3,
        "fila_det_origen": "A",
        "hoja_destino":   "SOBRECOSTOS",
        "cols_destino":   [("A", "G"), ("I", "J"), ("K", "Q")],
        "fila_ini_destino": 7,
        "fila_det_destino": "A",
        "cols_formulas":  [("R", "EB")],
        "filtro_col":     "C",        # solo filas donde esta col == filtro_valor
        "filtro_valor":   "C.Frec",
        "detectar_fin_primera_vacia": True,  # parar en primera fila vacía en col det
        "ajustar_formulas": True,    # extender/recortar fórmulas para calzar con filas pegadas
    },
]

MAPEO_P3_FD = [
    {
        "hoja_origen":    "CPF Horario",
        "cols_origen":    [("B", "J")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "CPF_FD",
        "cols_destino":   [("D", "L")],
        "fila_ini_destino": 9,
        "fila_det_destino": "D",
        "cols_formulas":  [("A", "B"), ("N", "P")],
    },
    {
        "hoja_origen":    "CSF Horario",
        "cols_origen":    [("B", "H")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "CSF_FD",
        "cols_destino":   [("E", "K")],
        "fila_ini_destino": 9,
        "fila_det_destino": "E",
        "cols_formulas":  [("A", "D"), ("M", "N"), ("P", "X")],
    },
    {
        "hoja_origen":    "CTF Horario",
        "cols_origen":    [("B", "I")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "CTF_FD",
        "cols_destino":   [("E", "L")],
        "fila_ini_destino": 9,
        "fila_det_destino": "E",
        "cols_formulas":  [("A", "D"), ("N", "P"), ("R", "Y")],
    },
]

MAPEO_P5_FD = [
    {
        "hoja_origen":    "CPF Horario",
        "cols_origen":    [("B", "J")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CPF",
        "cols_destino":   [("B", "J")],
        "fila_ini_destino": 7,
        "fila_det_destino": "B",
        "cols_formulas":  [("A", "A")],
    },
    {
        "hoja_origen":    "CSF Horario",
        "cols_origen":    [("B", "H")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CSF",
        "cols_destino":   [("B", "H")],
        "fila_ini_destino": 7,
        "fila_det_destino": "B",
        "cols_formulas":  [("A", "A")],
    },
    {
        "hoja_origen":    "CTF Horario",
        "cols_origen":    [("B", "I")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD_CTF",
        "cols_destino":   [("B", "I")],
        "fila_ini_destino": 7,
        "fila_det_destino": "B",
        "cols_formulas":  [("A", "A")],
    },
]

MAPEO_P6_FD = [
    {
        "hoja_origen":    "CT Diario",
        "cols_origen":    [("B", "B"), ("D", "I")],
        "fila_ini_origen": 12,
        "fila_det_origen": "B",
        "hoja_destino":   "FD",
        "cols_destino":   [("B", "H")],
        "fila_ini_destino": 7,
        "fila_det_destino": "B",
        "cols_formulas":  [],
    },
]

# ── Búsqueda de archivos ─────────────────────────────────────────────────────
def buscar_sscc_desempeno(carpeta_reliq: Path) -> Path | None:
    """FD está un nivel arriba de 02 CASO RELIQUIDACION."""
    carpeta_fd = carpeta_reliq.parent / "FD"
    if not carpeta_fd.exists():
        # Buscar case-insensitive
        for sub in carpeta_reliq.parent.iterdir():
            if sub.is_dir() and normalizar(sub.name) == "fd":
                carpeta_fd = sub
                break
        else:
            return None
    patron = re.compile(r"sscc_desempeno.*\.(xlsx|xlsm)$", re.IGNORECASE)
    candidatos = [f for f in carpeta_fd.iterdir()
                  if f.is_file() and patron.search(f.name) and not f.name.startswith("~$")]
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]

def buscar_consolidado(carpeta_reliq: Path) -> Path | None:
    """01 Sobrecostos/Detalles diarios/02 Consolidado_Tabulado_AAMM…"""
    objetivo_sc = normalizar("01 sobrecostos")
    carpeta_sc = None
    for sub in carpeta_reliq.iterdir():
        if sub.is_dir() and normalizar(sub.name) == objetivo_sc:
            carpeta_sc = sub
            break
    if carpeta_sc is None:
        return None
    objetivo_det = normalizar("detalles diarios")
    carpeta_det = None
    for sub in carpeta_sc.iterdir():
        if sub.is_dir() and normalizar(sub.name) == objetivo_det:
            carpeta_det = sub
            break
    if carpeta_det is None:
        return None
    patron = re.compile(r"consolidado_tabulado.*\.(xlsx|xlsm)$", re.IGNORECASE)
    candidatos = [f for f in carpeta_det.iterdir()
                  if f.is_file() and patron.search(f.name) and not f.name.startswith("~$")]
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]

def buscar_sobrecostos_xlsm(carpeta_reliq: Path) -> Path | None:
    objetivo_carpeta = normalizar("01 sobrecostos")
    carpeta_sc = None
    for sub in carpeta_reliq.iterdir():
        if sub.is_dir() and normalizar(sub.name) == objetivo_carpeta:
            carpeta_sc = sub
            break
    if carpeta_sc is None:
        return None
    patron = re.compile(r"c[áa]lculo[_ ]sobrecostossscc.*\.xlsm$", re.IGNORECASE)
    candidatos = [f for f in carpeta_sc.iterdir()
                  if f.is_file() and patron.search(f.name) and not f.name.startswith("~$")]
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]

def buscar_planilla9(carpeta_reliq: Path, prefijo: str) -> Path | None:
    p9 = None
    for sub in carpeta_reliq.iterdir():
        if sub.is_dir() and "planilla 9" in normalizar(sub.name):
            p9 = sub
            break
    if p9 is None:
        return None
    prefijo_norm = normalizar(prefijo)
    candidatos = [
        f for f in p9.iterdir()
        if f.is_file() and f.suffix.lower() in (".xlsm", ".xlsx")
        and normalizar(f.name).startswith(prefijo_norm)
        and not f.name.startswith("~$")
    ]
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]

def buscar_prorrata(carpeta_reliq: Path) -> Path | None:
    """04 Planilla 9/Prorrata_Retiros_AAMM…"""
    p9 = None
    for sub in carpeta_reliq.iterdir():
        if sub.is_dir() and "planilla 9" in normalizar(sub.name):
            p9 = sub
            break
    if p9 is None:
        return None
    patron = re.compile(r"prorrata_retiros.*\.(xlsx|xlsm)$", re.IGNORECASE)
    candidatos = [f for f in p9.iterdir()
                  if f.is_file() and patron.search(f.name) and not f.name.startswith("~$")]
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]

# ── Lógica xlwings ───────────────────────────────────────────────────────────
def col_letra_a_num(letra: str) -> int:
    n = 0
    for c in letra.upper():
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n

def expandir_cols(rangos: list) -> list:
    out = []
    for ini, fin in rangos:
        for n in range(col_letra_a_num(ini), col_letra_a_num(fin) + 1):
            out.append(n)
    return out

def ultima_fila(sheet, col_letra: str, desde: int) -> int:
    """Última fila con valor O fórmula en la columna, desde 'desde'."""
    col_num = col_letra_a_num(col_letra)
    ultima = sheet.cells.last_cell.row
    if ultima < desde:
        return desde - 1
    rng = sheet.range((desde, col_num), (ultima, col_num))
    fila = desde - 1
    # Leer fórmulas (devuelve la fórmula si la hay, valor si no)
    formulas = rng.formula
    if formulas is None:
        return desde - 1
    if not isinstance(formulas, (list, tuple)):
        formulas = [[formulas]]
    elif not isinstance(formulas[0], (list, tuple)):
        formulas = [[v] for v in formulas]
    for i, row in enumerate(formulas):
        v = row[0] if row else None
        if v not in (None, ""):
            fila = desde + i
    return fila

def sheet_last_row_col(sheet, col_num: int, desde: int) -> int:
    """Última fila con valor o fórmula en columna por número, desde 'desde'."""
    try:
        ultima = sheet.cells.last_cell.row
        if ultima < desde:
            return desde - 1
        rng = sheet.range((desde, col_num), (ultima, col_num))
        formulas = rng.formula
        if formulas is None:
            return desde - 1
        if not isinstance(formulas, (list, tuple)):
            formulas = [[formulas]]
        elif not isinstance(formulas[0], (list, tuple)):
            formulas = [[v] for v in formulas]
        fila = desde - 1
        for i, row in enumerate(formulas):
            v = row[0] if row else None
            if v not in (None, ""):
                fila = desde + i
        return fila
    except Exception:
        return desde - 1

def ultima_fila_usedrange(sheet, fila_ini: int) -> int:
    """Última fila usada en la hoja según UsedRange (incluye fórmulas y formatos)."""
    try:
        last = sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count - 1
        return max(fila_ini - 1, last)
    except Exception:
        return sheet.cells.last_cell.row

def _pair_dest(cols_origen: list, cols_destino: list) -> list:
    if len(cols_destino) == 1 and len(cols_origen) > 1:
        ini = col_letra_a_num(cols_destino[0][0])
        out = []
        cursor = ini
        for ori_par in cols_origen:
            ancho = col_letra_a_num(ori_par[1]) - col_letra_a_num(ori_par[0]) + 1
            out.append(cursor)
            cursor += ancho
        return out
    return [col_letra_a_num(par[0]) for par in cols_destino]

def fmt_tiempo(segundos):
    m, s = divmod(int(segundos), 60)
    h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def aplicar_mapeo(app_xw, wb_o, wb_d, mapeo: list, log_func=print, progreso_func=None):
    """Aplica una lista de mapeos entre wb_o (origen) y wb_d (destino)."""
    total = len(mapeo)
    for idx, m in enumerate(mapeo, 1):
        pct = int(idx / total * 100)
        t_ini = time.time()
        log_func(f"  [{idx}/{total} - {pct}%] Actualizando {m['hoja_destino']} (desde {m['hoja_origen']})…")
        if progreso_func:
            progreso_func(idx - 1, total, f"Actualizando {m['hoja_destino']}")
        sh_o = wb_o.sheets[m["hoja_origen"]]
        sh_d = wb_d.sheets[m["hoja_destino"]]

        fila_ini_o = m["fila_ini_origen"]
        fila_ini_d = m["fila_ini_destino"]

        # Detectar fin de datos en origen
        if m.get("detectar_fin_primera_vacia"):
            # Primera fila VACÍA en col_det_origen
            col_det_num = col_letra_a_num(m["fila_det_origen"])
            ultima_posible = sh_o.cells.last_cell.row
            rng_det = sh_o.range((fila_ini_o, col_det_num), (ultima_posible, col_det_num))
            vals_det = rng_det.value
            if not isinstance(vals_det, list):
                vals_det = [vals_det]
            fila_fin_o = fila_ini_o - 1
            for i, v in enumerate(vals_det):
                if v in (None, ""):
                    break
                fila_fin_o = fila_ini_o + i
        else:
            fila_fin_o = ultima_fila(sh_o, m["fila_det_origen"], fila_ini_o)

        if fila_fin_o < fila_ini_o:
            log_func("     sin datos en origen, se omite")
            continue

        cols_dest_nums = expandir_cols(m["cols_destino"])
        cols_form_nums = expandir_cols(m["cols_formulas"]) if m["cols_formulas"] else []
        dest_starts = _pair_dest(m["cols_origen"], m["cols_destino"])

        # Si hay filtro, leer columna de filtro y construir lista de filas válidas
        filas_validas = None  # None = todas; lista = subset filtrado
        if m.get("filtro_col") and m.get("filtro_valor") is not None:
            filtro_col_num = col_letra_a_num(m["filtro_col"])
            vals_filtro = sh_o.range((fila_ini_o, filtro_col_num),
                                      (fila_fin_o, filtro_col_num)).value
            if not isinstance(vals_filtro, list):
                vals_filtro = [vals_filtro]
            filas_validas = [i for i, v in enumerate(vals_filtro)
                             if v == m["filtro_valor"]]
            log_func(f"     filtro {m['filtro_col']}={m['filtro_valor']}: {len(filas_validas)} de {len(vals_filtro)} filas")

        # Pegar valores
        for ori_par, dest_ini in zip(m["cols_origen"], dest_starts):
            ori_ini = col_letra_a_num(ori_par[0])
            ori_fin = col_letra_a_num(ori_par[1])
            ancho = ori_fin - ori_ini + 1

            vals = sh_o.range((fila_ini_o, ori_ini), (fila_fin_o, ori_fin)).value
            # Normalizar a lista 2D
            n_total = fila_fin_o - fila_ini_o + 1
            if not isinstance(vals, list):
                vals = [[vals]]
            elif n_total == 1:
                vals = [vals] if not isinstance(vals[0], list) else vals
            elif ancho == 1:
                vals = [[v] for v in vals]

            # Aplicar filtro: tomar solo filas con índices válidos
            if filas_validas is not None:
                vals = [vals[i] for i in filas_validas if i < len(vals)]

            n_filas_pegar = len(vals)
            if n_filas_pegar == 0:
                continue
            fila_fin_d_nueva = fila_ini_d + n_filas_pegar - 1
            sh_d.range((fila_ini_d, dest_ini),
                       (fila_fin_d_nueva, dest_ini + ancho - 1)).value = vals

        # Definir n_filas_nuevas según filtro aplicado o no
        if filas_validas is not None:
            n_filas_nuevas = len(filas_validas)
        else:
            n_filas_nuevas = fila_fin_o - fila_ini_o + 1

        # Extender/recortar fórmulas para que calcen exactamente con las filas
        # pegadas (solo si el mapeo lo pide). Se replica la fórmula de la fila
        # modelo (fila_ini_d) hacia abajo; Excel ajusta las referencias relativas.
        # El borrado posterior recorta cualquier fórmula sobrante más abajo.
        if m.get("ajustar_formulas") and cols_form_nums and n_filas_nuevas >= 1:
            fila_fin_form = fila_ini_d + n_filas_nuevas - 1
            for c in cols_form_nums:
                f_modelo = sh_d.range((fila_ini_d, c)).formula
                if f_modelo in (None, ""):
                    continue
                sh_d.range((fila_ini_d, c), (fila_fin_form, c)).formula = f_modelo

        # Borrar todo debajo de lo pegado hasta UsedRange (cols destino + fórmulas)
        # solo si realmente hay algo abajo
        fila_ini_borrar = fila_ini_d + n_filas_nuevas
        try:
            fila_fin_borrar = sh_d.api.UsedRange.Row + sh_d.api.UsedRange.Rows.Count - 1
        except Exception:
            fila_fin_borrar = 0

        if fila_ini_borrar <= fila_fin_borrar:
            # Verificar que realmente hay contenido en alguna de las cols a borrar
            hay_contenido = False
            for c in cols_dest_nums + cols_form_nums:
                try:
                    rng_check = sh_d.range((fila_ini_borrar, c), (fila_fin_borrar, c))
                    vals = rng_check.formula
                    if vals is None:
                        continue
                    if not isinstance(vals, (list, tuple)):
                        vals = [vals]
                    elif vals and isinstance(vals[0], (list, tuple)):
                        vals = [r[0] for r in vals]
                    if any(v not in (None, "") for v in vals):
                        hay_contenido = True
                        break
                except Exception:
                    pass

            if hay_contenido:
                for c in cols_dest_nums + cols_form_nums:
                    sh_d.range((fila_ini_borrar, c), (fila_fin_borrar, c)).clear_contents()
                log_func(f"     filas pegadas: {n_filas_nuevas} | borrado fila {fila_ini_borrar} a {fila_fin_borrar} | {fmt_tiempo(time.time() - t_ini)}")
            else:
                log_func(f"     filas pegadas: {n_filas_nuevas} | nada que borrar abajo | {fmt_tiempo(time.time() - t_ini)}")
        else:
            log_func(f"     filas pegadas: {n_filas_nuevas} | nada que borrar abajo | {fmt_tiempo(time.time() - t_ini)}")

        if progreso_func:
            progreso_func(idx, total, f"Hoja {m['hoja_destino']} OK")

# ── Configuración Prorrata (pivot tabla→matriz) ──────────────────────────────
# Origen: hoja PRORRATA_HORARIA_TABULAR, fila 2 en adelante
#   Col A: Hora Mensual | Col B: Suministrador | Col C: Prorrata_horaria
# Destino: hoja PRORRATA_RETIROS, matriz desde B8
#   B8='Hora' | C8,D8...=suministradores alfabético | B9,B10...=horas | resto=valores (0 si falta)
PRORRATA_CONFIG = {
    "hoja_origen":   "PRORRATA_HORARIA_TABULAR",
    "fila_ini_origen": 2,
    "col_hora":      "A",
    "col_suministrador": "B",
    "col_valor":     "C",
    "hoja_destino":  "PRORRATA_RETIROS",
    "celda_inicio":  ("B", 8),       # esquina B8 donde va "Hora"
    "borrar_desde":  ("B", 8),       # rango a borrar antes de pegar
    "borrar_hasta":  ("EC", 756),
}

def aplicar_prorrata(app_xw, wb_o, wb_d, cfg: dict, log_func=print, progreso_func=None):
    """Transforma tabla larga (Hora,Suministrador,Valor) a matriz pivote en destino."""
    import time as _t
    t_ini = _t.time()
    log_func(f"  Leyendo {cfg['hoja_origen']}…")
    if progreso_func:
        progreso_func(0, 1, "Leyendo tabla de prorrata")

    sh_o = wb_o.sheets[cfg["hoja_origen"]]
    sh_d = wb_d.sheets[cfg["hoja_destino"]]

    fila_ini_o = cfg["fila_ini_origen"]
    c_hora = col_letra_a_num(cfg["col_hora"])
    c_sum  = col_letra_a_num(cfg["col_suministrador"])
    c_val  = col_letra_a_num(cfg["col_valor"])

    # Detectar fin de datos por col Hora (primera vacía)
    ultima_posible = sh_o.cells.last_cell.row
    col_min = min(c_hora, c_sum, c_val)
    col_max = max(c_hora, c_sum, c_val)
    datos = sh_o.range((fila_ini_o, col_min), (ultima_posible, col_max)).value
    if not isinstance(datos, list):
        datos = [datos]
    if datos and not isinstance(datos[0], list):
        datos = [datos]

    # Índices relativos dentro del bloque leído
    idx_hora = c_hora - col_min
    idx_sum  = c_sum - col_min
    idx_val  = c_val - col_min

    # Construir diccionario {(hora, suministrador): valor}, recolectar horas y suministradores
    valores = {}
    horas = []
    horas_set = set()
    suministradores_set = set()
    for fila in datos:
        h = fila[idx_hora] if idx_hora < len(fila) else None
        if h in (None, ""):
            break  # primera fila vacía = fin de datos
        s = fila[idx_sum] if idx_sum < len(fila) else None
        v = fila[idx_val] if idx_val < len(fila) else None
        if s in (None, ""):
            continue
        valores[(h, s)] = v
        if h not in horas_set:
            horas_set.add(h)
            horas.append(h)
        suministradores_set.add(s)

    suministradores = sorted(suministradores_set)  # orden alfabético
    log_func(f"  {len(horas)} horas × {len(suministradores)} suministradores")

    # ── Borrar contenido previo (B8:EC756) ──────────────────────────────────
    bc, br = cfg["borrar_desde"]
    ec, er = cfg["borrar_hasta"]
    log_func(f"  Borrando rango {bc}{br}:{ec}{er}…")
    sh_d.range((br, col_letra_a_num(bc)), (er, col_letra_a_num(ec))).clear_contents()

    # ── Construir matriz completa de una vez ─────────────────────────────────
    if progreso_func:
        progreso_func(0, 1, "Construyendo matriz de prorrata")
    col_ini_letra, fila_ini_d = cfg["celda_inicio"]
    col_ini_num = col_letra_a_num(col_ini_letra)

    # Fila de encabezado: ["Hora", sum1, sum2, ...]
    matriz = []
    matriz.append(["Hora"] + suministradores)
    # Filas de datos: [hora, val(s1), val(s2), ...]
    for h in horas:
        fila = [h]
        for s in suministradores:
            v = valores.get((h, s), 0)
            if v in (None, ""):
                v = 0
            fila.append(v)
        matriz.append(fila)

    n_filas = len(matriz)
    n_cols = len(suministradores) + 1

    # Escribir todo el bloque de una sola vez (rápido)
    log_func(f"  Escribiendo matriz {n_filas} filas × {n_cols} cols…")
    sh_d.range((fila_ini_d, col_ini_num)).value = matriz

    log_func(f"  Prorrata completada en {fmt_tiempo(_t.time() - t_ini)}")
    if progreso_func:
        progreso_func(1, 1, "Prorrata OK")

def ejecutar_actualizacion(carpeta_reliq: Path, planilla: str,
                            hacer_fd: bool, hacer_otro: bool,
                            log_func=print, progreso_func=None,
                            rutas: dict | None = None) -> tuple[bool, list]:
    """
    planilla: 'sc' | 'p3' | 'p5' | 'p6'
    hacer_fd: actualizar FD
    hacer_otro: para sc=Consolidado; p3/p5/p6=Prorrata (próximamente)
    rutas: dict del traspaso del Revisor. Las claves que vengan se usan tal
           cual y NO se vuelve a buscar el archivo; las que falten caen al
           buscar_* de siempre.
    Retorna (ok, lista_rutas_modificadas)
    """
    try:
        import xlwings as xw
    except ImportError:
        log_func("ERROR: Falta instalar xlwings: pip install xlwings")
        return False, []

    rutas = rutas or {}

    def _resolver(clave: str, buscador, etiqueta: str) -> Path | None:
        """Prefiere la ruta que manda el Revisor; si no vino, la busca."""
        dada = rutas.get(clave)
        if dada:
            p = Path(dada)
            if p.is_file():
                log_func(f"  {etiqueta}: ruta recibida del Revisor")
                return p
            log_func(f"  OJO: la ruta de {etiqueta} que mandó el Revisor no existe:")
            log_func(f"       {p}")
            log_func(f"       se busca el archivo como en modo manual.")
        return buscador()

    app_xw = None
    libros = {}
    rutas_ok = []

    try:
        # Excel invisible para máxima velocidad
        app_xw = xw.App(visible=False, add_book=False)
        app_xw.display_alerts = False
        app_xw.screen_updating = False

        # ── Abrir archivos necesarios ────────────────────────────────────────
        ruta_sscc = _resolver("sscc_desempeno",
                              lambda: buscar_sscc_desempeno(carpeta_reliq),
                              "SSCC_Desempeno")
        ruta_destino = None

        if planilla == "sc":
            ruta_destino = _resolver(
                "calculo_sscc_maestro",
                lambda: buscar_sobrecostos_xlsm(carpeta_reliq),
                "Cálculo_SobrecostosSSCC")
        elif planilla == "p3":
            ruta_destino = _resolver(
                "p3", lambda: buscar_planilla9(carpeta_reliq, "3_remuneracion_subastas"),
                "3_REMUNERACIÓN_SUBASTAS")
        elif planilla == "p5":
            ruta_destino = _resolver(
                "p5", lambda: buscar_planilla9(carpeta_reliq, "5_remuneracion_cra"),
                "5_REMUNERACIÓN_CRA")
        elif planilla == "p6":
            ruta_destino = _resolver(
                "p6", lambda: buscar_planilla9(carpeta_reliq, "6_remuneracion_rea"),
                "6_REMUNERACIÓN_REA")

        if ruta_destino is None or not ruta_destino.exists():
            log_func("ERROR: No se encontró archivo destino.")
            app_xw.quit()
            return False, []

        log_func(f"Abriendo archivo destino (puede tardar)…")
        log_func(f"  {ruta_destino.name}")
        wb_d = app_xw.books.open(str(ruta_destino), update_links=False)
        wb_d.app.calculation = "manual"
        libros["destino"] = wb_d
        log_func("Destino abierto ✓")

        # ── FD ───────────────────────────────────────────────────────────────
        if hacer_fd:
            if ruta_sscc is None or not ruta_sscc.exists():
                log_func("ERROR: No se encontró SSCC_Desempeno.")
                raise FileNotFoundError("SSCC_Desempeno no encontrado")

            if "sscc" not in libros:
                log_func(f"Abriendo origen SSCC_Desempeno (puede tardar)…")
                log_func(f"  {ruta_sscc.name}")
                libros["sscc"] = app_xw.books.open(str(ruta_sscc), read_only=True, update_links=False)
                log_func("Origen SSCC abierto ✓")

            mapeo_fd = {
                "sc": MAPEO_SOBRECOSTOS_FD,
                "p3": MAPEO_P3_FD,
                "p5": MAPEO_P5_FD,
                "p6": MAPEO_P6_FD,
            }[planilla]

            log_func("── Actualizando FD ──")
            aplicar_mapeo(app_xw, libros["sscc"], wb_d, mapeo_fd, log_func, progreso_func)
            rutas_ok.append(ruta_destino)

        # ── Consolidado (solo sc) ────────────────────────────────────────────
        if hacer_otro and planilla == "sc":
            ruta_consol = _resolver("consolidado_tabulado",
                                    lambda: buscar_consolidado(carpeta_reliq),
                                    "Consolidado_Tabulado")
            if ruta_consol is None or not ruta_consol.exists():
                log_func("ERROR: No se encontró Consolidado_Tabulado.")
                raise FileNotFoundError("Consolidado_Tabulado no encontrado")

            if "consol" not in libros:
                log_func(f"Abriendo origen Consolidado (puede tardar)…")
                log_func(f"  {ruta_consol.name}")
                libros["consol"] = app_xw.books.open(str(ruta_consol), read_only=True, update_links=False)
                log_func("Origen Consolidado abierto ✓")

            log_func("── Actualizando Consolidado ──")
            aplicar_mapeo(app_xw, libros["consol"], wb_d, MAPEO_SOBRECOSTOS_CONSOLIDADO, log_func, progreso_func)
            if ruta_destino not in rutas_ok:
                rutas_ok.append(ruta_destino)

        # ── Prorrata (p3, p5, p6) ────────────────────────────────────────────
        if hacer_otro and planilla in ("p3", "p5", "p6"):
            ruta_prorrata = _resolver("prorrata_retiros",
                                      lambda: buscar_prorrata(carpeta_reliq),
                                      "Prorrata_Retiros")
            if ruta_prorrata is None or not ruta_prorrata.exists():
                log_func("ERROR: No se encontró Prorrata_Retiros.")
                raise FileNotFoundError("Prorrata_Retiros no encontrado")

            if "prorrata" not in libros:
                log_func(f"Abriendo origen Prorrata_Retiros (puede tardar)…")
                log_func(f"  {ruta_prorrata.name}")
                libros["prorrata"] = app_xw.books.open(str(ruta_prorrata), read_only=True, update_links=False)
                log_func("Origen Prorrata abierto ✓")

            log_func("── Actualizando Prorrata ──")
            aplicar_prorrata(app_xw, libros["prorrata"], wb_d, PRORRATA_CONFIG, log_func, progreso_func)
            if ruta_destino not in rutas_ok:
                rutas_ok.append(ruta_destino)

        # ── Cerrar orígenes (no el destino) ─────────────────────────────────
        log_func("Cerrando orígenes…")
        for key in ("sscc", "consol", "prorrata"):
            if key in libros:
                libros[key].close()

        # Restaurar cálculo automático y recalcular
        log_func("Recalculando fórmulas…")
        wb_d.app.calculation = "automatic"

        # Guardar el archivo con los cambios aplicados
        log_func("Guardando archivo…")
        wb_d.save()
        log_func("Archivo guardado ✓")

        # Hacer visible — SIN UserControl/Interactive que rompen detección de cambios
        log_func("Mostrando archivo en Excel…")
        app_xw.api.Visible = True
        app_xw.screen_updating = True
        app_xw.display_alerts = True
        wb_d.activate()

        # Liberar referencias pero NO setear app_xw=None ni hacer gc.collect.
        # xlwings mantiene el handle de Excel vivo internamente; al salir de la
        # función Python suelta sus refs y Excel queda como proceso autónomo
        # que SÍ detecta modificaciones del usuario.
        libros.clear()
        wb_d = None

        return True, rutas_ok

    except Exception as e:
        tb = traceback.format_exc()
        log_func(f"EXCEPCIÓN: {e}")
        try:
            for wb in libros.values():
                try:
                    wb.close()
                except Exception:
                    pass
            if app_xw is not None:
                app_xw.quit()
        except Exception:
            pass
        return False, []

# ── Ventana ──────────────────────────────────────────────────────────────────
def main():
    cfg = leer_config()
    traspaso = leer_traspaso(sys.argv)
    # modo["traspaso"] se apaga si el usuario elige otra carpeta a mano: desde
    # ese momento la ventana vuelve a buscar los archivos por su cuenta.
    modo = {"traspaso": traspaso is not None}

    root = tk.Tk()
    root.title("Actualiza FD" + ("  —  enviado por el Revisor" if traspaso else ""))
    root.resizable(True, True)
    root.geometry("1000x700")

    # ── Layout: botones fijos abajo + canvas scrollable arriba ───────────────
    # frame_btns_fijo va PRIMERO al fondo
    frame_btns_fijo = tk.Frame(root)
    frame_btns_fijo.pack(side="bottom", fill="x", pady=8)

    # Canvas + scrollbar para el contenido
    canvas_outer = tk.Canvas(root, borderwidth=0, highlightthickness=0)
    scrollbar_v = tk.Scrollbar(root, orient="vertical", command=canvas_outer.yview)
    canvas_outer.configure(yscrollcommand=scrollbar_v.set)

    scrollbar_v.pack(side="right", fill="y")
    canvas_outer.pack(side="left", fill="both", expand=True)

    # Frame interno que CONTIENE todo el contenido scrollable
    contenedor = tk.Frame(canvas_outer)
    canvas_window = canvas_outer.create_window((0, 0), window=contenedor, anchor="nw")

    def _actualizar_scroll(event=None):
        canvas_outer.configure(scrollregion=canvas_outer.bbox("all"))
        # Hacer que el contenedor tenga el ancho del canvas
        canvas_outer.itemconfig(canvas_window, width=canvas_outer.winfo_width())

    contenedor.bind("<Configure>", _actualizar_scroll)
    canvas_outer.bind("<Configure>", _actualizar_scroll)

    # La rueda del mouse: el canvas de la ventana solo scrollea si el puntero NO
    # esta sobre un widget que scrollea solo (el log, por ejemplo).
    #
    # bind_all captura la rueda GLOBALMENTE. Sin este filtro, al intentar subir
    # en el log se movia la ventana entera y el log no se movia: los dos
    # respondian a la misma rueda, o directamente ganaba el canvas.
    def _rueda(e, _cv=canvas_outer):
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

    canvas_outer.bind_all("<MouseWheel>", _rueda)

    # ── Estado ───────────────────────────────────────────────────────────────
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))

    # Radio: planilla seleccionada (única). El Revisor manda desde qué archivo
    # se apretó el botón, así que se preselecciona esa.
    plan_ini = (traspaso or {}).get("planilla")
    var_planilla = tk.StringVar(
        value=plan_ini if plan_ini in ("sc", "p3", "p5", "p6") else "sc")

    # Checkboxes por planilla: {planilla: {"fd": BoolVar, "otro": BoolVar}}
    # Viniendo del Revisor arrancan en blanco a propósito: el usuario tiene que
    # decir qué quiere actualizar. A mano se mantiene FD premarcado como antes.
    fd_ini = not modo["traspaso"]
    opciones = {
        k: {"fd": tk.BooleanVar(value=fd_ini), "otro": tk.BooleanVar(value=False)}
        for k in ("sc", "p3", "p5", "p6")
    }

    paths_detectados = {
        "sc":    {"destino": tk.StringVar(), "sscc": tk.StringVar(), "consol": tk.StringVar()},
        "p3":    {"destino": tk.StringVar(), "sscc": tk.StringVar(), "prorrata": tk.StringVar()},
        "p5":    {"destino": tk.StringVar(), "sscc": tk.StringVar(), "prorrata": tk.StringVar()},
        "p6":    {"destino": tk.StringVar(), "sscc": tk.StringVar(), "prorrata": tk.StringVar()},
    }

    # Widgets de checkboxes para mostrar/ocultar
    frames_planilla = {}

    # Claves del JSON de traspaso que alimentan cada label de la ventana.
    MAPA_TRASPASO = {
        "sc": {"destino": "calculo_sscc_maestro", "sscc": "sscc_desempeno",
               "consol": "consolidado_tabulado"},
        "p3": {"destino": "p3", "sscc": "sscc_desempeno", "prorrata": "prorrata_retiros"},
        "p5": {"destino": "p5", "sscc": "sscc_desempeno", "prorrata": "prorrata_retiros"},
        "p6": {"destino": "p6", "sscc": "sscc_desempeno", "prorrata": "prorrata_retiros"},
    }

    def aplicar_traspaso():
        """Rellena los labels con las rutas que mandó el Revisor. No busca nada:
        las reglas de búsqueda de este script y las del Revisor no coinciden, y
        el punto del traspaso es justamente usar las del Revisor."""
        dadas = traspaso.get("rutas", {})
        for plan, subs in MAPA_TRASPASO.items():
            for sub, clave in subs.items():
                val = dadas.get(clave)
                paths_detectados[plan][sub].set(
                    str(val) if val else f"[el Revisor no mandó {clave}]")
        actualizar_colores()

    def refrescar_archivos(*_):
        if modo["traspaso"]:
            aplicar_traspaso()
            return
        ruta = var_carpeta.get()
        if not ruta or not Path(ruta).is_dir():
            return
        reliq = Path(ruta)

        sscc = buscar_sscc_desempeno(reliq)
        sscc_str = str(sscc) if sscc else "[SSCC_Desempeno no encontrado]"
        for key in ("sc", "p3", "p5", "p6"):
            paths_detectados[key]["sscc"].set(sscc_str)

        sc = buscar_sobrecostos_xlsm(reliq)
        paths_detectados["sc"]["destino"].set(str(sc) if sc else "[Cálculo_SobrecostosSSCC no encontrado]")
        consol = buscar_consolidado(reliq)
        paths_detectados["sc"]["consol"].set(str(consol) if consol else "[Consolidado_Tabulado no encontrado]")

        p3 = buscar_planilla9(reliq, "3_remuneracion_subastas")
        paths_detectados["p3"]["destino"].set(str(p3) if p3 else "[3_REMUNERACIÓN_SUBASTAS no encontrado]")

        p5 = buscar_planilla9(reliq, "5_remuneracion_cra")
        paths_detectados["p5"]["destino"].set(str(p5) if p5 else "[5_REMUNERACIÓN_CRA no encontrado]")

        p6 = buscar_planilla9(reliq, "6_remuneracion_rea")
        paths_detectados["p6"]["destino"].set(str(p6) if p6 else "[6_REMUNERACIÓN_REA no encontrado]")

        # Origen prorrata (compartido por p3/p5/p6)
        prorrata = buscar_prorrata(reliq)
        prorrata_str = str(prorrata) if prorrata else "[Prorrata_Retiros no encontrado]"
        for key in ("p3", "p5", "p6"):
            paths_detectados[key]["prorrata"].set(prorrata_str)

        actualizar_colores()

    def actualizar_colores():
        for key, lbl_map in labels_archivos.items():
            for subkey, lbl in lbl_map.items():
                val = paths_detectados[key][subkey].get()
                # Rojo tambien si la ruta viene escrita pero el archivo no esta:
                # importa al recibir rutas del Revisor, que las resolvio antes.
                ok = bool(val) and not val.startswith("[")
                if ok:
                    try:
                        ok = Path(val).is_file()
                    except Exception:
                        ok = False
                lbl.config(fg="blue" if ok else "red")

    def mostrar_frame_planilla(*_):
        plan = var_planilla.get()
        for k, fr in frames_planilla.items():
            if k == plan:
                fr.pack(fill="x", padx=6, pady=2)
            else:
                fr.pack_forget()

    # ── Aviso de procedencia (solo si vino del Revisor) ──────────────────────
    # Va arriba de todo y bien visible para que no se actualice el mes equivocado.
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

    # ── Instrucciones ────────────────────────────────────────────────────────
    frame_info = tk.LabelFrame(contenedor, text="Instrucciones", padx=10, pady=6)
    frame_info.pack(fill="x", padx=20, pady=(10, 4))
    tk.Label(frame_info, text=INSTRUCCIONES, wraplength=700,
             justify="left", fg="#333333", font=("Segoe UI", 9)).pack()

    # ── Selector carpeta ─────────────────────────────────────────────────────
    frame_c = tk.LabelFrame(contenedor, text="Carpeta 02 CASO RELIQUIDACION", padx=10, pady=6)
    frame_c.pack(fill="x", padx=20, pady=4)

    lbl_carpeta = tk.Label(frame_c, textvariable=var_carpeta, wraplength=700,
                           justify="center", fg="gray", cursor="hand2",
                           font=("Segoe UI", 9))
    lbl_carpeta.pack()
    lbl_carpeta.bind("<Button-1>",
                     lambda e: abrir_en_explorador(var_carpeta.get()) if var_carpeta.get() else None)

    def sel_carpeta():
        ini = cfg.get("carpeta_reliq", "")
        ini = ini if ini and Path(ini).exists() else ""
        ruta = filedialog.askdirectory(title="Selecciona 02 CASO RELIQUIDACION", initialdir=ini)
        if ruta:
            var_carpeta.set(ruta)
            lbl_carpeta.config(fg="blue" if Path(ruta).is_dir() else "red")
            guardar_config({"carpeta_reliq": ruta})
            # Si el usuario elige carpeta a mano, manda el, no el traspaso:
            # se vuelve a buscar todo dentro de la carpeta que eligio.
            if modo["traspaso"]:
                modo["traspaso"] = False
                log("Carpeta elegida a mano: se dejan de usar las rutas del Revisor.")
            refrescar_archivos()

    tk.Button(frame_c, text="Seleccionar carpeta", command=sel_carpeta).pack(pady=(4, 0))

    # ── Planillas + opciones ─────────────────────────────────────────────────
    frame_plan = tk.LabelFrame(contenedor, text="Planilla a actualizar", padx=10, pady=6)
    frame_plan.pack(fill="x", padx=20, pady=4)

    labels_archivos = {
        "sc": {}, "p3": {}, "p5": {}, "p6": {},
    }

    def hacer_lbl_archivo(parent, var, clave_plan, clave_sub):
        lbl = tk.Label(parent, textvariable=var, wraplength=680,
                       justify="left", fg="gray", font=("Segoe UI", 8), cursor="hand2")
        lbl.pack(anchor="w", padx=30)
        lbl.bind("<Button-1>",
                 lambda e: abrir_en_explorador(var.get(), es_archivo=True)
                 if var.get() and Path(var.get()).exists() else None)
        labels_archivos[clave_plan][clave_sub] = lbl

    def hacer_fila_planilla(parent, key, nombre_radio, nombre_fd, nombre_otro, otro_habilitado):
        rb = tk.Radiobutton(parent, text=nombre_radio, variable=var_planilla,
                            value=key, command=mostrar_frame_planilla,
                            font=("Segoe UI", 9, "bold"))
        rb.pack(anchor="w")

        fr = tk.Frame(parent)
        frames_planilla[key] = fr
        # No se hace pack aquí; lo maneja mostrar_frame_planilla

        # Destino
        tk.Label(fr, text="Destino:", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=14)
        hacer_lbl_archivo(fr, paths_detectados[key]["destino"], key, "destino")

        # Origen SSCC
        tk.Label(fr, text="Origen SSCC_Desempeno:", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=14)
        hacer_lbl_archivo(fr, paths_detectados[key]["sscc"], key, "sscc")

        # Checkbox FD
        tk.Checkbutton(fr, text=nombre_fd, variable=opciones[key]["fd"],
                       font=("Segoe UI", 9)).pack(anchor="w", padx=20)

        # Checkbox otro
        cb_otro = tk.Checkbutton(fr, text=nombre_otro, variable=opciones[key]["otro"],
                                 font=("Segoe UI", 9))
        if not otro_habilitado:
            cb_otro.config(state="disabled")
            opciones[key]["otro"].set(False)
        cb_otro.pack(anchor="w", padx=20)

        # Origen Consolidado (solo sc)
        if key == "sc":
            tk.Label(fr, text="Origen Consolidado:", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=14)
            hacer_lbl_archivo(fr, paths_detectados["sc"]["consol"], "sc", "consol")

        # Origen Prorrata (p3/p5/p6)
        if key in ("p3", "p5", "p6"):
            tk.Label(fr, text="Origen Prorrata_Retiros:", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=14)
            hacer_lbl_archivo(fr, paths_detectados[key]["prorrata"], key, "prorrata")

    hacer_fila_planilla(frame_plan, "sc",
                        "Cálculo_SobrecostosSSCC",
                        "Actualizar FD",
                        "Traer Consolidado",
                        otro_habilitado=True)

    hacer_fila_planilla(frame_plan, "p3",
                        "3_REMUNERACIÓN_SUBASTAS_E_ID",
                        "Actualizar FD",
                        "Actualizar Prorrata",
                        otro_habilitado=True)

    hacer_fila_planilla(frame_plan, "p5",
                        "5_REMUNERACIÓN_CRA",
                        "Actualizar FD",
                        "Actualizar Prorrata",
                        otro_habilitado=True)

    hacer_fila_planilla(frame_plan, "p6",
                        "6_REMUNERACIÓN_REA_Y_CO_ERNC",
                        "Actualizar FD",
                        "Actualizar Prorrata",
                        otro_habilitado=True)

    mostrar_frame_planilla()

    # ── Log ──────────────────────────────────────────────────────────────────
    # Barra de progreso y tiempo
    frame_prog = tk.Frame(contenedor)
    frame_prog.pack(fill="x", padx=20, pady=(4, 0))

    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    tk.Label(frame_prog, textvariable=var_estado, font=("Segoe UI", 9),
             fg="#333333", anchor="w").pack(side="left", fill="x", expand=True)
    tk.Label(frame_prog, textvariable=var_tiempo, font=("Consolas", 10, "bold"),
             fg="#2d7a2d").pack(side="right", padx=8)

    progress_bar = ttk.Progressbar(contenedor, mode="determinate", length=400)
    progress_bar.pack(fill="x", padx=20, pady=(2, 4))

    frame_log = tk.LabelFrame(contenedor, text="Progreso detallado", padx=6, pady=4)
    frame_log.pack(fill="both", expand=True, padx=20, pady=4)
    txt_log = tk.Text(frame_log, height=6, width=70, font=("Consolas", 9))
    scroll = tk.Scrollbar(frame_log, command=txt_log.yview)
    txt_log.config(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    txt_log.pack(fill="both", expand=True)

    def log(msg):
        txt_log.insert("end", str(msg) + "\n")
        txt_log.see("end")
        root.update_idletasks()

    # ── Ejecutar ─────────────────────────────────────────────────────────────
    # Estado del timer
    timer_state = {"running": False, "t_ini": 0.0}

    def actualizar_timer():
        if timer_state["running"]:
            t = time.time() - timer_state["t_ini"]
            var_tiempo.set(fmt_tiempo(t))
            root.after(500, actualizar_timer)

    def callback_progreso(actual, total, mensaje):
        var_estado.set(mensaje)
        progress_bar["maximum"] = total
        progress_bar["value"] = actual
        root.update_idletasks()

    def ejecutar():
        ruta = var_carpeta.get()
        # Con traspaso las rutas son absolutas y la carpeta solo sirve de respaldo,
        # asi que no se exige que exista.
        if not modo["traspaso"] and (not ruta or not Path(ruta).is_dir()):
            messagebox.showerror("Error", "Selecciona la carpeta 02 CASO RELIQUIDACION.")
            return

        plan = var_planilla.get()
        hacer_fd   = opciones[plan]["fd"].get()
        hacer_otro = opciones[plan]["otro"].get()

        if not hacer_fd and not hacer_otro:
            messagebox.showwarning("Aviso", "Marca al menos una opción (FD o la otra).")
            return

        btn_ejecutar.config(state="disabled", bg="#aaaaaa")
        var_estado.set("Iniciando…")
        progress_bar["value"] = 0
        var_tiempo.set("00:00:00")
        timer_state["running"] = True
        timer_state["t_ini"] = time.time()
        root.after(100, actualizar_timer)
        root.update_idletasks()

        txt_log.delete("1.0", "end")
        log(f"Planilla: {plan} | FD: {hacer_fd} | Otro: {hacer_otro}")
        if modo["traspaso"]:
            log(f"Rutas enviadas por el Revisor — mes {traspaso.get('aamm') or '?'}")
        log("─" * 50)

        ok, rutas = ejecutar_actualizacion(
            Path(ruta) if ruta else Path("."), plan, hacer_fd, hacer_otro,
            log_func=log, progreso_func=callback_progreso,
            rutas=traspaso.get("rutas") if modo["traspaso"] else None
        )

        timer_state["running"] = False
        tiempo_total = fmt_tiempo(time.time() - timer_state["t_ini"])
        var_tiempo.set(tiempo_total)
        btn_ejecutar.config(state="normal", bg="#2d7a2d")

        if ok and rutas:
            var_estado.set(f"✓ Completado en {tiempo_total}")
            progress_bar["value"] = progress_bar["maximum"]
            log(f"✓ Proceso completado en {tiempo_total}")
            _mostrar_completado(rutas)
        elif not ok:
            var_estado.set("✗ Error")
            messagebox.showerror("Error", "Hubo un error. Revisa el log.")

    def _mostrar_completado(rutas):
        win = tk.Toplevel(root)
        win.title("Completado")
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="Archivo actualizado (queda abierto en Excel):",
                 font=("Segoe UI", 9, "bold"), pady=12, padx=20).pack()
        for r in rutas:
            lbl = tk.Label(win, text=str(r), fg="blue", cursor="hand2",
                           wraplength=460, justify="center", padx=20)
            lbl.pack(pady=2)
            lbl.bind("<Button-1>", lambda e, p=r: abrir_en_explorador(str(p), True))
        tk.Label(win, text="clic en la ruta para abrir su carpeta",
                 font=("Segoe UI", 7), fg="gray", pady=4).pack()
        tk.Button(win, text="Aceptar", width=10, command=win.destroy).pack(pady=(4, 12))
        win.update_idletasks()
        w, h = win.winfo_width(), win.winfo_height()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    # ── Botones (en frame fijo abajo) ────────────────────────────────────────
    tk.Button(frame_btns_fijo, text="Cerrar", width=12,
              command=root.destroy).pack(side="left", padx=8, expand=True)
    btn_ejecutar = tk.Button(frame_btns_fijo, text="Ejecutar", width=12,
              bg="#2d7a2d", fg="white",
              activebackground="#1f5c1f", activeforeground="white",
              command=ejecutar)
    btn_ejecutar.pack(side="left", padx=8, expand=True)

    # Inicializar
    refrescar_archivos()

    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")
    root.mainloop()

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()