# -*- coding: utf-8 -*-
"""Traspaso y lanzamiento de los actualizadores desde el Revisor.

Las dependencias de UI, rutas y configuración se inyectan desde el punto de
entrada para que este módulo no lo importe ni forme ciclos.
"""

from pathlib import Path
import re

_dependencias = None

def configurar(dependencias):
    """Inyecta el espacio histórico del Revisor sin importar su ventana."""
    global _dependencias
    _dependencias = dependencias
    globals().update({k: v for k, v in dependencias.items()
                      if not k.startswith("__") and k not in globals()})

def bloqueado_para_escritura(self, ruta):
    """True si otro proceso tiene el archivo tomado y no se va a poder
    guardar. Es la prueba DIRECTA de lo que importa: se pide permiso de
    escritura y se cierra al instante, sin escribir ni un byte.

    Es mejor que mirar el "~$": Excel deja ese archivo huérfano cuando se
    cae o lo matan, y entonces el "~$" existe para siempre aunque el libro
    esté cerrado. Esta prueba, en cambio, dice la verdad de ahora."""
    try:
        with open(ruta, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        # No existe, o algo raro con la red. No se bloquea aca: si de verdad
        # hay un problema, el actualizador lo va a decir con mejor detalle.
        return False


def lock_excel(self, ruta):
    """Ruta del "~$" que Excel deja al lado del libro abierto, si existe."""
    try:
        p = Path(ruta)
        lock = p.parent / ("~$" + p.name)
        return lock if lock.exists() else None
    except Exception:
        return None


def dueno_del_lock(self, lock):
    """Nombre de quien tiene el libro abierto. Excel lo guarda dentro del
    "~$": un byte con el largo y despues el nombre, a veces en ANSI y a
    veces en UTF-16. El formato no esta documentado, asi que esto es al
    mejor esfuerzo: si no se entiende se devuelve None y listo."""
    try:
        crudo = Path(lock).read_bytes()[:200]
    except Exception:
        return None
    if not crudo:
        return None
    candidatos = []
    # Se prueban las dos codificaciones y ambos desfases, porque el primer
    # byte es el largo y descuadra el UTF-16 si no se saltea.
    for datos in (crudo, crudo[1:]):
        for codec in ("utf-16-le", "latin-1"):
            try:
                candidatos.append(datos.decode(codec, errors="ignore"))
            except Exception:
                pass
    for texto in candidatos:
        # Los \x00 sobrantes cortarian el nombre en pedazos de 1 letra.
        limpio = texto.replace("\x00", "")
        m = re.search(r"[A-Za-z0-9._\-]{2,}(?:[ ][A-Za-z0-9._\-]+)*", limpio)
        if m:
            nombre = m.group(0).strip()
            if len(nombre) >= 3:
                return nombre
    return None


def armar_traspaso(self, aamm, planilla, nid=None):
    """JSON que reciben los actualizadores. Solo se escriben las rutas que
    el revisor pudo resolver; el actualizador tolera que falte alguna.

    nid: la fila DESDE LA QUE se apreto el boton. Hace falta cuando el mismo
    script cuelga de varias filas (Prorratear esta en los tres .mdb): sin
    esto el script no puede saber a cual le dieron y tiene que adivinar.
    Se manda el id del nodo y tambien la clave de su ruta, ya resuelta.
    """
    rutas = {}
    for nid_, clave in CLAVES_TRASPASO.items():
        r = self.rutas.get(nid_)
        if r is not None:
            rutas[clave] = str(r)
    d = {"origen": _traspaso.ORIGEN,
         "version": TRASPASO_VERSION,
         "aamm": aamm,
         "carpeta_reliq": self.var_base.get(),
         "rutas": rutas}
    if planilla:
        d["planilla"] = planilla
    if nid:
        d["nodo"] = nid
        # La ruta de ESA fila, sin que el script tenga que saber que clave
        # le corresponde.
        r = self.rutas.get(nid)
        if r is not None:
            d["ruta_nodo"] = str(r)
        clave = CLAVES_TRASPASO.get(nid)
        if clave:
            d["clave_nodo"] = clave
    return d


def lanzar_actualizador(self, nid, indice=0):
    spec = ACTUALIZADORES[nid][indice]
    # El script puede estar en una subcarpeta ("Reemplazos REUC/..."), asi
    # que se arma con Path para que la barra funcione igual en Windows.
    script_rel = Path(spec["script"])
    script_nombre = script_rel.name
    planilla = spec.get("planilla")
    nodo = NODO_POR_ID[nid]

    # Se relee el disco ANTES de leer la ruta del destino: si apareció una
    # revisión nueva (R01E donde antes había R01D), la ruta que tenía en
    # memoria apunta al archivo viejo y le mandaríamos ese al actualizador.
    # Se releen el destino y todo lo que va en el JSON de traspaso, que es
    # lo unico que el actualizador va a usar.
    if not self.actualizar(motivo=f"previo a actualizar data de {nid}",
                           solo_ids=set(CLAVES_TRASPASO) | {nid}):
        return
    destino = self.rutas.get(nid)

    if destino is None:
        messagebox.showwarning(
            "Falta el archivo",
            f"No se encontró el archivo de esta fila.\n\n{nodo['texto']}\n\n"
            "Aprieta ACTUALIZAR para volver a buscarlo.")
        return

    script = DIR_SCRIPT / script_rel
    if not script.is_file():
        messagebox.showerror(
            "Falta el actualizador",
            f"No se encontró:\n\n{script_rel}\n\nSe buscó en:\n{DIR_SCRIPT}")
        self.log(f"  ERROR: no se encontró {script_rel} en {DIR_SCRIPT}")
        return

    # Aviso obligatorio cuando el destino es un Excel que se va a escribir:
    # xlwings necesita poder guardarlo. Se prueba la escritura de verdad; el
    # "~$" solo sirve como pista de quien lo tiene, porque queda huerfano
    # cuando Excel se cae. Un .mdb no tiene "~$", asi que no aplica.
    destino_es_excel = destino.suffix.lower() in XL
    if destino_es_excel:
        lock = self._lock_excel(destino)
        if self._bloqueado_para_escritura(destino):
            dueno = self._dueno_del_lock(lock) if lock else None
            self.log(f"  NO SE LANZÓ {script_nombre}: el destino está tomado.")
            self.log(f"     {destino}")
            if dueno:
                self.log(f"     lo tiene abierto: {dueno}")
            messagebox.showwarning(
                "El archivo está en uso",
                f"No se puede escribir en:\n\n{destino.name}\n\n"
                + (f"Lo tiene abierto: {dueno}\n\n" if dueno else "")
                + "Ciérralo (o pídele que lo cierre) y volvé a intentar.")
            return
        if lock is not None:
            # Se puede escribir, así que el "~$" sobró: Excel lo dejó tirado.
            dueno = self._dueno_del_lock(lock)
            self.log(f"  OJO: hay un '~$' huérfano al lado del destino.")
            self.log(f"     {lock}")
            if dueno:
                self.log(f"     quedó a nombre de: {dueno}")
            self.log("     El archivo SÍ se puede escribir, así que se puede seguir.")
            if not messagebox.askyesno(
                    "Quedó un archivo de bloqueo",
                    f"Hay un archivo de bloqueo de Excel al lado del destino:\n\n"
                    f"{lock.name}\n\n"
                    + (f"Quedó a nombre de: {dueno}\n\n" if dueno else "")
                    + "Pero el archivo se puede escribir sin problema, así que "
                      "lo más probable es que sea basura de un Excel que se cerró mal.\n\n"
                      "Podés borrar ese archivo con tranquilidad (está oculto).\n\n"
                      "¿Actualizar igual?"):
                return

    aamm = (self.var_aamm.get() or "").strip()
    if not aamm:
        if not messagebox.askyesno(
                "Sin mes",
                "No hay AAMM definido, así que no se sabe a qué mes pertenece "
                "esto.\n\nEl traspaso se guardará en __config__/sin_mes.\n\n¿Lanzar igual?"):
            return

    # Se deja el mes en curso en config.json para que los actualizadores
    # abiertos a mano después arranquen apuntando al mes correcto.
    cambios = {"carpeta_reliq": self.var_base.get()}
    mdb = self.rutas.get("a_mdb_sscc")
    if mdb is not None:
        cambios["mdb"] = str(mdb)
    guardar_config(cambios)
    self.cfg.update(cambios)

    traspaso = self._armar_traspaso(aamm, planilla, nid)
    carpeta_salida = dir_config_mes(aamm or "sin_mes", crear=True)
    ruta_traspaso = carpeta_salida / ARCHIVO_TRASPASO
    try:
        escribir_json(ruta_traspaso, traspaso)
    except Exception as e:
        messagebox.showerror("No se pudo escribir el traspaso", str(e))
        self.log(f"  ERROR al escribir {ruta_traspaso}: {e}")
        return

    try:
        # cwd = carpeta del propio script, no la del revisor: el de
        # Reemplazos vive en su subcarpeta y resuelve cosas relativas a ella.
        subprocess.Popen([sys.executable, str(script), str(ruta_traspaso)],
                         cwd=str(script.parent))
    except Exception as e:
        messagebox.showerror("No se pudo lanzar el actualizador", str(e))
        self.log(f"  ERROR al lanzar {script_nombre}: {e}")
        return

    # Bitácora: si después algo no cuadra, esto permite reconstruir qué pasó.
    self.log("-" * 96)
    self.log(f"ACTUALIZAR DATA  {datetime.now():%d-%m-%Y %H:%M:%S}")
    self.log(f"  script   : {script_nombre}"
             + (f"   (planilla {planilla})" if planilla else ""))
    self.log(f"  mes      : {aamm or 'sin_mes'}")
    self.log(f"  destino  : {destino}")
    self.log(f"  traspaso : {ruta_traspaso}")
    for clave in sorted(traspaso["rutas"]):
        self.log(f"     {clave:22s} {traspaso['rutas'][clave]}")
    faltan = [c for c in CLAVES_TRASPASO.values() if c not in traspaso["rutas"]]
    if faltan:
        self.log(f"  sin resolver: {', '.join(sorted(faltan))}")
    if destino_es_excel:
        self.log("  OJO: el actualizador deja el archivo GUARDADO y ABIERTO en Excel.")
        self.log("       El revisor lee la versión de disco, así que si sigues editando")
        self.log("       sin guardar, lo que ves y lo que verifica el revisor difieren.")
    self.log("  Cuando termine, aprieta ACTUALIZAR para releer fechas y vencer "
             "las verificaciones que dependan del archivo.")
