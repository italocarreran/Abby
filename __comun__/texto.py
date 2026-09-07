"""Normalización compartida de nombres del dominio y de rutas.

Expone dos contratos distintos a propósito: ``suave`` conserva separadores y
sirve para nombres de hoja/carpeta; ``clave`` elimina espacios y guiones bajos
para comparar centrales, empresas, conceptos o encabezados. Los guiones medios
y signos se conservan porque sí distinguen unidades y conceptos.
"""

import re
import unicodedata


def sin_tildes(texto):
    """Convierte a texto sin marcas diacríticas; ``None`` equivale a vacío."""
    nfkd = unicodedata.normalize("NFKD", str("" if texto is None else texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def suave(texto):
    """Sin tildes, espacios colapsados y en minúsculas."""
    return re.sub(r"\s+", " ", sin_tildes(texto)).strip().lower()


def suave_textual(texto):
    """Variante histórica que convierte también ``None`` en el texto ``none``."""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    limpio = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", limpio).strip().lower()


def suave_requerido(texto):
    """Variante histórica que exige un ``str`` y falla con ``None``."""
    nfkd = unicodedata.normalize("NFKD", texto)
    limpio = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", limpio).strip().lower()


def clave(texto, mayusculas=False):
    """Sin tildes, espacios ni guiones bajos; conserva guiones medios/signos."""
    valor = re.sub(r"[\s_]+", "", sin_tildes(texto)).strip()
    return valor.upper() if mayusculas else valor.lower()


def clave_mayusculas(texto):
    """Variante histórica usada para centrales, empresas y conceptos."""
    return clave(texto or "", mayusculas=True)


def clave_columna(texto):
    """Clave mayúscula que además equipara ``Año`` con ``Anio``."""
    return clave_mayusculas(texto).replace("ANIO", "ANO")
