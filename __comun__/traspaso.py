"""Contrato compartido del JSON que el Revisor pasa a los actualizadores.

El argumento es opcional: si falta, no apunta a un archivo valido, el JSON esta
roto, viene de otro origen o usa una version futura, se devuelve ``None`` para
que el ejecutable conserve su modo manual. Las claves particulares dentro de
``rutas`` siguen siendo responsabilidad de cada actualizador.
"""

import json
from pathlib import Path


ORIGEN = "Revisor_Reliquidacion"
VERSION_ACTUAL = 1


def validar(data, version_max=VERSION_ACTUAL):
    """Devuelve un traspaso normalizado o ``None`` si no cumple el contrato."""
    try:
        if not isinstance(data, dict) or data.get("origen") != ORIGEN:
            return None
        if int(data.get("version", 0)) > version_max:
            return None
        if not isinstance(data.get("rutas"), dict):
            data["rutas"] = {}
        return data
    except Exception:
        return None


def leer(ruta, version_max=VERSION_ACTUAL):
    """Lee y valida ``ruta``; nunca lanza por errores de entrada o de archivo."""
    try:
        ruta = Path(str(ruta).strip())
        if not ruta.is_file():
            return None
        with open(ruta, "r", encoding="utf-8") as f:
            return validar(json.load(f), version_max)
    except Exception:
        return None


def leer_argumento(argv, version_max=VERSION_ACTUAL):
    """Lee ``argv[1]`` o devuelve ``None`` para continuar en modo manual."""
    try:
        if len(argv) < 2 or not str(argv[1]).strip():
            return None
        return leer(argv[1], version_max)
    except Exception:
        return None
