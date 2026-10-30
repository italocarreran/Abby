# -*- coding: utf-8 -*-
"""Centrales de embalse que se agregaron (o descartaron) desde el Revisor.

La lista base de los 27 embalses vive en el codigo (``CENTRALES_EMBALSE`` del
Revisor y de ``Actualiza_SC_CO.py``). Cuando V10 encuentra en un origen una
central "-numero" que no esta en la lista, el Revisor pregunta (usuaria,
2026-10-30) si es un embalse nuevo:

* si -> va a ``agregadas`` y desde ahi la toman el Revisor y el actualizador
  de la hoja "SC y CO";
* no -> va a ``ignoradas`` y no se vuelve a avisar.

Se guarda en ``__config__/centrales_embalse_15min.json``, compartido por
todos los usuarios: que una central sea o no embalse no depende de quien mira.
Solo biblioteca estandar + ``__comun__/config.py`` (escritura atomica).
"""

from pathlib import Path

from __comun__ import config as _cfg
from __comun__ import texto as _texto

RUTA = Path(__file__).resolve().parents[2] / "__config__" / "centrales_embalse_15min.json"

clave = _texto.clave_mayusculas


def _lista(todo, k):
    v = todo.get(k)
    return [str(x).strip() for x in v if str(x).strip()] if isinstance(v, list) else []


def leer(ruta=None):
    """(agregadas, ignoradas), cada una una lista de nombres."""
    todo = _cfg.leer_todo(ruta or RUTA)
    return _lista(todo, "agregadas"), _lista(todo, "ignoradas")


def ampliar(base, ruta=None):
    """Agrega en el lugar (a la misma lista) las centrales agregadas que no
    esten ya. Devuelve las que agrego."""
    vistas = {clave(x) for x in base}
    nuevas = []
    for n in leer(ruta)[0]:
        if clave(n) not in vistas:
            base.append(n)
            vistas.add(clave(n))
            nuevas.append(n)
    return nuevas


def claves_ignoradas(ruta=None):
    return {clave(x) for x in leer(ruta)[1]}


def guardar(agregar=(), ignorar=(), ruta=None):
    """Anota las decisiones. Una central agregada deja de estar ignorada y al
    reves. Devuelve True si se escribio."""
    agregar = [str(x).strip() for x in agregar if str(x).strip()]
    ignorar = [str(x).strip() for x in ignorar if str(x).strip()]
    if not (agregar or ignorar):
        return True

    def mutar(todo):
        ag, ig = _lista(todo, "agregadas"), _lista(todo, "ignoradas")
        for nuevas, propia, otra in ((agregar, ag, ig), (ignorar, ig, ag)):
            for n in nuevas:
                otra[:] = [x for x in otra if clave(x) != clave(n)]
                if clave(n) not in {clave(x) for x in propia}:
                    propia.append(n)
        todo["agregadas"], todo["ignoradas"] = ag, ig

    ruta = Path(ruta or RUTA)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    return _cfg.modificar(ruta, mutar)
