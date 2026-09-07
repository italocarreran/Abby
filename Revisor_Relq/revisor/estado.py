"""Persistencia del estado de verificaciones y del caché de valores."""

from datetime import datetime
import json


ARCHIVO_ESTADO = "_revisor_verificaciones.json"
ARCHIVO_CACHE = "_revisor_cache_valores.json"


class Estado:
    """Verificaciones de un mes. Se guardan en ``__config__/AAAA/MM Mes``.

    Las rutas, la escritura y la firma se reciben explícitamente."""

    def __init__(self, dir_mes, escribir_json, firma_verificador):
        self._dir_mes = dir_mes
        self._escribir_json = escribir_json
        self._firma_verificador = firma_verificador
        self.ruta = None
        self.aamm = None
        self.data = {}

    def cargar(self, aamm):
        self.aamm = str(aamm).strip() if aamm else None
        self.data = {}
        self.ruta = None
        if not self.aamm:
            return False
        self.ruta = self._dir_mes(self.aamm) / ARCHIVO_ESTADO
        if self.ruta.exists():
            try:
                with open(self.ruta, "r", encoding="utf-8") as archivo:
                    self.data = json.load(archivo)
                return True
            except Exception:
                self.data = {}
        return False

    def existe(self):
        return bool(self.ruta and self.ruta.exists())

    def guardar(self):
        if not self.aamm:
            return False
        try:
            self._dir_mes(self.aamm, crear=True)
            self._escribir_json(self._dir_mes(self.aamm) / ARCHIVO_ESTADO, self.data)
            return True
        except Exception:
            return False

    def get(self, vid):
        return self.data.get(vid)

    def vigente(self, vid):
        """El registro guardado, si hay alguno.

        No se mira la firma a propósito. Un resultado verificado vale como
        verificado, sin importar con qué versión de la definición se corrió: al
        usuario le da lo mismo esa distinción y solo le ensuciaba la fila. La
        diferencia de definición se avisa en la ventana de detalle, que es donde
        importa, porque ahí sí se muestran hojas y rangos concretos.
        """
        return self.data.get(vid) or None

    def firma_guardada_distinta(self, vid):
        """True si hay un registro pero es de una definición anterior.
        Solo para avisarlo en el detalle, no para invalidar nada."""
        reg = self.data.get(vid)
        return bool(reg) and reg.get("firma") != self._firma_verificador(vid)

    def set(self, vid, registro):
        registro = dict(registro)
        registro["firma"] = self._firma_verificador(vid)
        self.data[vid] = registro
        return self.guardar()


class CacheValores:
    """Guarda el valor ya leído de cada origen junto con la ruta y la fecha de
    modificación del archivo. Si el archivo no cambió y se pide lo mismo, no se
    vuelve a abrir. Se guarda en ``__config__/AAAA/MM Mes`` entre ejecuciones.

    Las dependencias de disco las inyecta el punto de entrada."""

    def __init__(self, dir_mes, escribir_json, mtime, tamano, fmt_fecha):
        self._dir_mes = dir_mes
        self._escribir_json = escribir_json
        self._mtime = mtime
        self._tamano = tamano
        self._fmt_fecha = fmt_fecha
        self.aamm = None
        self.data = {}

    def cargar(self, aamm):
        self.aamm = str(aamm).strip() if aamm else None
        self.data = {}
        if not self.aamm:
            return
        ruta = self._dir_mes(self.aamm) / ARCHIVO_CACHE
        if ruta.exists():
            try:
                with open(ruta, "r", encoding="utf-8") as archivo:
                    self.data = json.load(archivo)
            except Exception:
                self.data = {}

    def guardar(self):
        if not self.aamm:
            return False
        try:
            self._dir_mes(self.aamm, crear=True)
            self._escribir_json(self._dir_mes(self.aamm) / ARCHIVO_CACHE, self.data)
            return True
        except Exception:
            return False

    def obtener(self, clave, ruta, huella):
        reg = self.data.get(clave)
        if not reg or ruta is None:
            return None
        if reg.get("archivo") != ruta.name or reg.get("huella") != huella:
            return None
        # OJO: aca la comparacion es EXACTA, sin la tolerancia de TOL_MTIME.
        # Esa tolerancia sirve para comparar copias entre discos distintos, pero
        # si se usara aca un archivo guardado 1 segundo despues de leerlo
        # devolveria el valor viejo. Se compara ademas el tamaño.
        ts_ahora, tam_ahora = self._mtime(ruta), self._tamano(ruta)
        if ts_ahora is None or reg.get("mtime") is None:
            return None
        if abs(reg["mtime"] - ts_ahora) > 1e-6:
            return None
        if reg.get("tamano") is not None and reg["tamano"] != tam_ahora:
            return None
        if not isinstance(reg.get("valor"), (int, float)):
            return None
        return reg

    def poner(self, clave, ruta, huella, valor, filas=None):
        if ruta is None:
            return
        self.data[clave] = {
            "archivo": ruta.name,
            "ruta": str(ruta),
            "mtime": self._mtime(ruta),
            "mtime_texto": self._fmt_fecha(self._mtime(ruta)),
            "tamano": self._tamano(ruta),
            "huella": huella,
            "valor": valor,
            "filas": filas,
            "leido": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self.guardar()

    def descartar(self, claves=None):
        if claves is None:
            self.data = {}
        else:
            for clave in claves:
                self.data.pop(clave, None)
        self.guardar()


def leer_estado_mes(aamm, dir_mes):
    """Lee el estado de otro mes sin modificar la instancia activa."""
    ruta = dir_mes(aamm) / ARCHIVO_ESTADO
    if not ruta.exists():
        return None
    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            return json.load(archivo)
    except Exception:
        return {}
