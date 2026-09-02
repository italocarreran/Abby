# -*- coding: utf-8 -*-
"""Lectura y escritura del config.json, indexado por <equipo>_<usuario>.

Esto estaba copiado en los 10 scripts, con cuatro variantes de
`_modificar_config` y cuatro de `get_usuario` que se fueron separando entre si.
Las diferencias eran casi todas cosmeticas (docstrings, type hints, nombres de
variable) salvo dos, que aca quedan resueltas para todos:

  * `ActualizaRemplazos.py` escribia el archivo entero SIN pasar por un .tmp, y
    si el config existia pero no se podia interpretar lo pisaba con {}. O sea:
    un config.json roto le borraba los ajustes a todos los demas scripts. Ahora
    usa la misma regla que el resto, que es no escribir.
  * `get_usuario` tenia versiones con y sin try/except. Queda la defensiva.

REGLAS QUE NO CAMBIAN, porque el archivo es compartido:

  * Solo se agregan o actualizan claves. Nunca se borra nada ajeno.
  * Si el archivo existe pero no se puede interpretar NO se escribe: mejor
    perder un ajuste que el archivo entero.
  * La escritura es atomica (.tmp + os.replace), asi que un corte a mitad de
    camino no deja el archivo truncado.

Solo biblioteca estandar.
"""

from __future__ import annotations

import json
import os
import socket
from pathlib import Path

__all__ = ["clave_equipo", "escribir_json", "leer", "leer_todo", "modificar",
           "guardar"]


def clave_equipo() -> str:
    """Devuelve `<equipo>_<usuario>`, que es como se indexa el config.json.

    No lanza: si el nombre del equipo no se puede averiguar devuelve
    "desconocido", porque quedarse sin ajustes es mejor que no arrancar.
    """
    try:
        usuario = (os.environ.get("USERNAME") or os.environ.get("USER")
                   or "desconocido")
        return f"{socket.gethostname()}_{usuario}"
    except Exception:
        return "desconocido"


def escribir_json(ruta, data) -> None:
    """Escritura atomica: primero un .tmp y despues os.replace.

    Evita dejar el archivo truncado si algo falla a medio camino. Se usa
    tambien para el JSON de traspaso del revisor, no solo para el config.
    """
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_suffix(ruta.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, ruta)


def leer_todo(ruta) -> dict:
    """El config.json completo, con todos los equipos. {} si no se puede leer."""
    ruta = Path(ruta)
    try:
        if ruta.exists():
            with open(ruta, "r", encoding="utf-8") as f:
                todo = json.load(f)
            if isinstance(todo, dict):
                return todo
    except Exception:
        pass
    return {}


def leer(ruta) -> dict:
    """El bloque del equipo actual. {} si no hay archivo o esta roto."""
    return leer_todo(ruta).get(clave_equipo(), {})


def modificar(ruta, mutador) -> bool:
    """Lee el config entero, lo modifica con `mutador(todo)` y lo reescribe.

    `mutador` recibe el dict completo (todos los equipos), no solo el bloque
    propio, porque algunos scripts anidan claves por mes.

    Devuelve True si se escribio. Devuelve False sin tocar el archivo si existe
    pero no se puede interpretar o no es un dict: es un archivo compartido y
    pisarlo le borraria los ajustes a los demas scripts.
    """
    ruta = Path(ruta)
    todo: dict = {}
    if ruta.exists():
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                todo = json.load(f)
        except Exception:
            return False
        if not isinstance(todo, dict):
            return False
    try:
        mutador(todo)
        escribir_json(ruta, todo)
        return True
    except Exception:
        return False


def guardar(ruta, data: dict) -> bool:
    """Agrega o actualiza claves en el bloque del equipo actual."""
    return modificar(ruta,
                     lambda todo: todo.setdefault(clave_equipo(), {}).update(data))
