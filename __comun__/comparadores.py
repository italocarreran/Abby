"""Infraestructura compartida por los dos comparadores.

No conoce Access, consolidados, SQL ni columnas. Reúne únicamente el puente de
cola hacia tkinter y estado mensual puro, para que los motores de dominio sigan
separados y los workers nunca toquen widgets.
"""

import os
import queue
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from __comun__ import archivos
from __comun__ import texto


class Entrada:
    """Archivo o carpeta con metadatos obtenidos en el mismo ``scandir``."""

    __slots__ = ("nombre", "ruta", "es_dir", "mtime", "size")

    def __init__(self, nombre, ruta, es_dir, mtime, size):
        self.nombre = nombre
        self.ruta = ruta
        self.es_dir = es_dir
        self.mtime = mtime
        self.size = size


class CacheDirectorios:
    """Una consulta de red por carpeta y reutilización hasta ``limpiar``."""

    def __init__(self):
        self.datos = {}

    def limpiar(self):
        self.datos.clear()

    def listar(self, carpeta):
        if carpeta is None:
            return []
        clave = str(carpeta)
        if clave in self.datos:
            return self.datos[clave]
        resultado = []
        try:
            with os.scandir(clave) as entradas:
                for entrada in entradas:
                    try:
                        stat = entrada.stat()
                        resultado.append(Entrada(
                            entrada.name, Path(entrada.path), entrada.is_dir(),
                            int(stat.st_mtime), stat.st_size,
                        ))
                    except OSError:
                        continue
        except (OSError, ValueError):
            resultado = []
        self.datos[clave] = resultado
        return resultado

    def huella_entrada(self, ruta):
        ruta = Path(ruta)
        for entrada in self.listar(ruta.parent):
            if entrada.nombre == ruta.name and not entrada.es_dir:
                return f"{entrada.mtime}_{entrada.size}"
        return None


class CachePorArchivo:
    """Recuerda un cálculo caro por archivo mientras el archivo no cambie.

    Leer el esquema de un parquet abre el archivo y parsea su pie; la ventana
    lo hacía decenas de veces por cada repintado. La firma (mtime + tamaño)
    sale de un ``stat`` local, que es barato, y cambia sola cuando el archivo
    se reescribe: por eso no hace falta invalidar nada a mano después de
    consolidar.

    A diferencia de ``CacheDirectorios``, esta caché NO se limpia en cada
    refresco: la firma ya la mantiene honesta.
    """

    def __init__(self, calcular):
        self.calcular = calcular
        self.datos = {}

    def limpiar(self):
        self.datos.clear()

    def firma(self, ruta):
        try:
            estado = Path(ruta).stat()
        except OSError:
            return None
        return (estado.st_mtime_ns, estado.st_size)

    def __call__(self, ruta):
        clave = str(ruta)
        firma = self.firma(ruta)
        if firma is None:
            # El archivo no está: el cálculo tiene que decidirlo él, y es
            # barato porque no hay nada que leer.
            self.datos.pop(clave, None)
            return self.calcular(ruta)
        previo = self.datos.get(clave)
        if previo is not None and previo[0] == firma:
            return previo[1]
        valor = self.calcular(ruta)
        self.datos[clave] = (firma, valor)
        return valor


def hallar_revisor(raiz):
    """Encuentra la carpeta hermana por su entry point, no por su nombre."""
    raiz = Path(raiz)
    preferida = raiz / "Revisor_Relq"
    if (preferida / "Revisor_Reliquidacion.py").is_file():
        return preferida
    for carpeta in sorted(p for p in raiz.iterdir() if p.is_dir()):
        if (carpeta / "Revisor_Reliquidacion.py").is_file():
            return carpeta
    return None


def abrir_en_explorador(ruta, es_archivo=True):
    """Abre la carpeta existente más cercana sin ejecutar el archivo."""
    if not ruta:
        return
    path = Path(ruta)
    if not path.exists():
        path = path.parent
        if not path.exists():
            return
    carpeta = path.parent if es_archivo and path.is_file() else path
    try:
        if sys.platform == "win32":
            subprocess.Popen(["explorer", str(carpeta)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(carpeta)])
        else:
            subprocess.Popen(["xdg-open", str(carpeta)])
    except Exception:
        pass


def subcarpeta(padre, nombre_buscado, listar):
    """Busca coincidencia exacta normalizada y luego coincidencia contenida."""
    if padre is None:
        return None
    objetivo = texto.suave(nombre_buscado)
    entradas = [entrada for entrada in listar(padre) if entrada.es_dir]
    for entrada in entradas:
        if texto.suave(entrada.nombre) == objetivo:
            return entrada.ruta
    for entrada in entradas:
        if objetivo in texto.suave(entrada.nombre):
            return entrada.ruta
    return None


def buscar_mdb(carpeta, patron, listar):
    """MDB/ACCDB original más reciente usando el listado cacheado del llamador."""
    if carpeta is None:
        return None
    candidatos = [
        entrada for entrada in listar(carpeta)
        if not entrada.es_dir
        and Path(entrada.nombre).suffix.lower() in (".mdb", ".accdb")
        and patron.search(entrada.nombre)
        and not archivos.es_copia(entrada.nombre)
        and not entrada.nombre.startswith("~$")
    ]
    elegido = archivos.mas_reciente(candidatos, lambda entrada: entrada.mtime)
    return elegido.ruta if elegido is not None else None


def huella(ruta, huella_entrada):
    if not ruta:
        return None
    try:
        return huella_entrada(Path(ruta))
    except Exception:
        return None


def tabla_por_nombre(nombres, candidatos, normalizar=texto.clave):
    mapa = {normalizar(nombre): nombre for nombre in nombres}
    for candidato in candidatos:
        if normalizar(candidato) in mapa:
            return mapa[normalizar(candidato)]
    return None


def una_fila(cursor, defecto=None):
    fila = cursor.fetchone()
    return fila if fila is not None else defecto


class ColaTk:
    """Cola de mensajes cuyo ``bombear`` se ejecuta siempre en el hilo de Tk."""

    def __init__(self, root, intervalo_ms=100):
        self.root = root
        self.intervalo_ms = intervalo_ms
        self.cola = queue.Queue()
        self.txt = None
        self.var_estado = None
        self.barra = None

    def conectar(self, txt, var_estado, barra):
        """Conecta widgets una vez construida la ventana e inicia el bombeo."""
        self.txt = txt
        self.var_estado = var_estado
        self.barra = barra
        self.root.after(self.intervalo_ms, self.bombear)

    def log(self, mensaje):
        self.cola.put(("log", str(mensaje)))

    def progreso(self, **opciones):
        self.cola.put(("barra", opciones))

    def estado(self, texto):
        self.cola.put(("estado", texto))

    def llamar(self, funcion, *args):
        self.cola.put(("llamar", (funcion, args)))

    def bombear(self):
        """Aplica mensajes pendientes y vuelve a programarse con ``after``."""
        assert self.txt is not None
        assert self.var_estado is not None
        assert self.barra is not None
        while True:
            try:
                accion, valor = self.cola.get_nowait()
            except queue.Empty:
                break
            if accion == "log":
                self.txt.insert("end", valor + "\n")
                self.txt.see("end")
            elif accion == "estado":
                self.var_estado.set(valor)
            elif accion == "barra":
                if valor.pop("final", False):
                    self.barra.config(value=self.barra["maximum"])
                else:
                    self.barra.config(**valor)
            elif accion == "llamar":
                funcion, args = valor
                funcion(*args)
        self.root.after(self.intervalo_ms, self.bombear)


def meses_del_anio(anio):
    """Devuelve los doce AAMM para un año escrito con dos o cuatro dígitos."""
    texto = str(anio).strip()
    if len(texto) == 4 and texto.isdigit():
        aa = texto[2:]
    elif len(texto) == 2 and texto.isdigit():
        aa = texto
    else:
        return []
    return [f"{aa}{mes:02d}" for mes in range(1, 13)]


def mes_incluido(estado, aamm):
    """Indica si el mes entra al consolidado anual; por omisión entra.

    La marca vive DENTRO del registro del mes (``estado[aamm]["incluir"]``), que
    es donde la escriben los ``estado.json`` que ya existen en el equipo de la
    usuaria. Moverla a un diccionario aparte no rompe nada visible: simplemente
    deja de encontrar las exclusiones viejas y todos los meses vuelven a entrar
    al anual sin aviso.
    """
    registro = estado.get(aamm) or {}
    return bool(registro.get("incluir", True))


def fijar_incluido(estado, aamm, valor):
    estado.setdefault(aamm, {})["incluir"] = bool(valor)


def color_de(estado, colores):
    """Traduce el estado semántico a la paleta activa.

    Son cinco casos, no tres: "desactualizado" es amarillo como "pendiente"
    (el mes está, hay que rehacerlo) y NO rojo, que significa que falta el
    archivo. Cualquier estado desconocido cae en gris, no en rojo.
    """
    return {
        "falta": colores["rojo"],
        "pendiente": colores["amarillo"],
        "desactualizado": colores["amarillo"],
        "ok": colores["verde"],
    }.get(estado, colores["gris"])


def fmt_tiempo(segundos):
    minutos, segundos = divmod(int(segundos), 60)
    horas, minutos = divmod(minutos, 60)
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"


def ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def cargar_estado(ruta, leer_json):
    estado = leer_json(ruta, {})
    return estado if isinstance(estado, dict) else {}


def guardar_estado(ruta, estado, escribir_json):
    escribir_json(ruta, estado)


def firma_vistas(meses, path_vista):
    firma = {}
    for mes in meses:
        vista = path_vista(mes)
        if vista.exists():
            firma[mes] = int(vista.stat().st_mtime)
    return firma


def es_hoja_propia(nombre, hojas_fijas):
    """Reconoce resúmenes y hojas mensuales ``AAMM``/``AAMM_N``."""
    if nombre in hojas_fijas or nombre == "RESUMEN":
        return True
    base, _, resto = nombre.partition("_")
    if not (len(base) == 4 and base.isdigit() and 1 <= int(base[2:]) <= 12):
        return False
    return resto == "" or resto.isdigit()


def hojas_ajenas(destino, openpyxl, es_propia, log=print):
    """Lista hojas que el comparador no puede pisar; falla de forma conservadora."""
    try:
        wb = openpyxl.load_workbook(str(destino), read_only=True)
        try:
            return [nombre for nombre in wb.sheetnames if not es_propia(nombre)]
        finally:
            wb.close()
    except Exception as e:
        log(f"  ! No se pudo revisar {destino.name} ({e}); se reescribe entero.")
        return []


def respaldar(destino, carpeta, log=print, conservar=5):
    """Copia con metadatos y conserva los últimos respaldos del mismo libro."""
    try:
        carpeta.mkdir(parents=True, exist_ok=True)
        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        copia = carpeta / f"{destino.stem}_{marca}{destino.suffix}"
        shutil.copy2(destino, copia)
        previas = sorted(carpeta.glob(f"{destino.stem}_*{destino.suffix}"))
        for vieja in previas[:-conservar]:
            try:
                vieja.unlink()
            except Exception:
                pass
        log(f"  Respaldo: {copia.name}")
        return copia
    except Exception as e:
        log(f"  ! No se pudo respaldar {destino.name}: {e}")
        return None
