# -*- coding: utf-8 -*-
"""Cambios de hora de Chile, detectados solos a partir del AAMM.

Copiado de ``Balance_SSCC/Script/nucleo/cambio_hora.py`` (2026-11-24) para el
aviso del recuadro «Balance del caso» del Revisor 15 min, que es igual al de la
ventana del BESS (pedido de la usuaria). Se dejo fuera ``hora_mes_correlativa``
(usa pandas y el Revisor no la necesita) y se agrego ``bloques_del_mes``.

Se recorre el mes hora por hora en UTC y se mira el desfase de
``America/Santiago``: donde cambia, hay cambio de hora. No hay fechas
escritas a mano, asi que si la ley cambia la fecha basta con que la
base de zonas horarias (``tzdata``) este al dia.

Dos clases de cambio:

- **adelanta** (septiembre, horario de verano): el reloj salta de 00:00
  a 01:00. El dia tiene 23 horas y el periodo 1 (00:00-01:00) no
  existe.
- **atrasa** (abril, horario de invierno): el reloj vuelve de 00:00 a
  23:00 del dia anterior. Ese dia tiene 25 horas y el periodo 24
  (23:00-24:00) ocurre dos veces.

En 15 minutos cada hora son 4 bloques: el mes del cambio tiene 4 bloques
de menos (septiembre) o de mas (abril). Lo dice ``describir_periodo``.
"""

import calendar
import datetime
from dataclasses import dataclass
from functools import lru_cache

ZONA_HORARIA = "America/Santiago"

_UTC = datetime.timezone.utc
_UNA_HORA = datetime.timedelta(hours=1)
_TODOS_LOS_PERIODOS = frozenset(range(1, 25))


@dataclass(frozen=True)
class CambioHora:
    """Un cambio de hora y el dia local que lo sufre."""

    fecha: datetime.date          # dia de 23 o 25 horas
    instante: datetime.datetime   # primera hora local con el desfase nuevo
    utc_anterior: datetime.timedelta
    utc_nuevo: datetime.timedelta
    periodos_inexistentes: tuple  # periodos 1..24 que el reloj se salta
    periodos_repetidos: tuple     # periodos 1..24 que ocurren dos veces

    @property
    def tipo(self):
        return "adelanta" if self.utc_nuevo > self.utc_anterior else "atrasa"

    @property
    def horas_del_dia(self):
        return (
            24 - len(self.periodos_inexistentes) + len(self.periodos_repetidos)
        )

    def descripcion(self):
        fecha = self.fecha.strftime("%d-%m-%Y")
        desfase = f"UTC{_texto_desfase(self.utc_anterior)} -> " \
                  f"UTC{_texto_desfase(self.utc_nuevo)}"
        if self.tipo == "adelanta":
            periodos = ", ".join(str(p) for p in self.periodos_inexistentes)
            return (
                f"{fecha}: se adelanta la hora ({desfase}); el dia tiene "
                f"{self.horas_del_dia} horas y no existe la hora {periodos}."
            )
        periodos = ", ".join(str(p) for p in self.periodos_repetidos)
        return (
            f"{fecha}: se atrasa la hora ({desfase}); el dia tiene "
            f"{self.horas_del_dia} horas y la hora {periodos} se repite."
        )


def _texto_desfase(desfase):
    horas = int(desfase.total_seconds() // 3600)
    return f"{horas:+d}"


@lru_cache(maxsize=None)
def _zona():
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo(ZONA_HORARIA)
    except Exception:  # sin base de zonas horarias (Windows sin tzdata)
        return None


def hay_base_de_zonas():
    """False si no se pudo cargar America/Santiago (falta tzdata)."""

    return _zona() is not None


def _periodos_del_dia(fecha, zona):
    """(inexistentes, repetidos) del dia local 'fecha'."""

    inexistentes, repetidos = [], []
    for periodo in range(1, 25):
        local = datetime.datetime(
            fecha.year, fecha.month, fecha.day, periodo - 1, tzinfo=zona
        )
        vuelta = local.astimezone(_UTC).astimezone(zona)
        if vuelta.replace(tzinfo=None) != local.replace(tzinfo=None):
            inexistentes.append(periodo)
        elif local.utcoffset() != local.replace(fold=1).utcoffset():
            repetidos.append(periodo)
    return tuple(inexistentes), tuple(repetidos)


def _cambios_con_zona(anio, mes, zona):
    # Se arranca un dia antes y se termina un dia despues del mes para
    # no perder un cambio pegado al borde; despues se filtra por mes.
    inicio = datetime.datetime(anio, mes, 1, tzinfo=_UTC) - datetime.timedelta(days=1)
    dias = calendar.monthrange(anio, mes)[1]
    fin = inicio + datetime.timedelta(days=dias + 2)

    cambios = []
    actual = inicio
    anterior = actual.astimezone(zona).utcoffset()
    while actual < fin:
        actual += _UNA_HORA
        local = actual.astimezone(zona)
        nuevo = local.utcoffset()
        if nuevo == anterior:
            continue
        # El dia que sufre el cambio: el de 23 o 25 horas, que es el
        # de la hora local justo ANTES del instante del cambio (en
        # abril el cambio cae a las 23:00 del sabado; en septiembre a
        # la 01:00 del domingo, con la 00:00 inexistente del domingo).
        candidatos = {
            (actual - _UNA_HORA).astimezone(zona).date(),
            local.date(),
        }
        for fecha in sorted(candidatos):
            inexistentes, repetidos = _periodos_del_dia(fecha, zona)
            if (inexistentes or repetidos) and (fecha.year, fecha.month) == (anio, mes):
                cambios.append(CambioHora(
                    fecha=fecha, instante=local,
                    utc_anterior=anterior, utc_nuevo=nuevo,
                    periodos_inexistentes=inexistentes,
                    periodos_repetidos=repetidos,
                ))
                break
        anterior = nuevo
    return cambios


def _primer_sabado(anio, mes):
    dia = datetime.date(anio, mes, 1)
    return dia + datetime.timedelta(days=(5 - dia.weekday()) % 7)


def _cambios_por_regla(anio, mes):
    """Respaldo sin tzdata: la regla vigente en Chile continental.

    Septiembre: el primer sabado a las 24:00 el reloj salta a la 01:00
    del domingo. Abril: el primer sabado a las 24:00 vuelve a las 23:00
    del mismo sabado.
    """

    if mes == 9:
        domingo = _primer_sabado(anio, 9) + datetime.timedelta(days=1)
        return [CambioHora(
            fecha=domingo,
            instante=datetime.datetime(domingo.year, 9, domingo.day, 1),
            utc_anterior=datetime.timedelta(hours=-4),
            utc_nuevo=datetime.timedelta(hours=-3),
            periodos_inexistentes=(1,), periodos_repetidos=(),
        )]
    if mes == 4:
        sabado = _primer_sabado(anio, 4)
        return [CambioHora(
            fecha=sabado,
            instante=datetime.datetime(sabado.year, 4, sabado.day, 23),
            utc_anterior=datetime.timedelta(hours=-3),
            utc_nuevo=datetime.timedelta(hours=-4),
            periodos_inexistentes=(), periodos_repetidos=(24,),
        )]
    return []


@lru_cache(maxsize=None)
def _cambios_del_mes(anio, mes):
    zona = _zona()
    if zona is None:
        return tuple(_cambios_por_regla(anio, mes))
    return tuple(_cambios_con_zona(anio, mes, zona))


def cambios_de_hora(anio, mes):
    """Lista de CambioHora del mes (vacia si no hay cambio)."""

    return list(_cambios_del_mes(int(anio), int(mes)))


def cambios_del_periodo(aamm):
    """Lo mismo, a partir del AAMM de la ventana (ej. '2609')."""

    aamm = str(aamm).strip()
    return cambios_de_hora(2000 + int(aamm[:2]), int(aamm[2:]))


def describir_periodo(aamm):
    """Texto para la ventana y el log: si el mes tiene cambio y cuando."""

    cambios = cambios_del_periodo(aamm)
    respaldo = "" if hay_base_de_zonas() else (
        " (sin tzdata: se usa la regla de septiembre/abril; instalar "
        "tzdata con pip)"
    )
    bloques = f" El mes tiene {bloques_del_mes(aamm)} bloques de 15 minutos."
    if not cambios:
        return "Sin cambio de hora en el mes." + bloques + respaldo
    return "Cambio de hora: " + " ".join(
        c.descripcion() for c in cambios
    ) + bloques + respaldo


def bloques_del_mes(aamm):
    """Bloques de 15 minutos del mes: 96 por dia, +-4 por cada hora que el
    cambio de hora salta o repite (2976 en un mes de 31 dias sin cambio)."""

    aamm = str(aamm).strip()
    anio, mes = 2000 + int(aamm[:2]), int(aamm[2:])
    horas = calendar.monthrange(anio, mes)[1] * 24
    for c in _cambios_del_mes(anio, mes):
        horas += len(c.periodos_repetidos) - len(c.periodos_inexistentes)
    return horas * 4


def horas_del_dia(anio, mes, dia):
    """Periodos que existen en el dia local de Chile (1..24 normalmente).

    El dia que se adelanta la hora falta el periodo del salto; el dia
    que se atrasa la hora que se repite se numera despues de la 24
    (1..25), como lo traen las ofertas SSCC en su columna E.
    """

    fecha = datetime.date(int(anio), int(mes), int(dia))
    for cambio in _cambios_del_mes(fecha.year, fecha.month):
        if cambio.fecha == fecha:
            extra = frozenset(
                range(25, 25 + len(cambio.periodos_repetidos))
            )
            return (
                _TODOS_LOS_PERIODOS
                - frozenset(cambio.periodos_inexistentes)
            ) | extra
    return _TODOS_LOS_PERIODOS
