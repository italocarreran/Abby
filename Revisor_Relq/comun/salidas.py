"""Rutas compartidas de ``00_Salidas``.

Centraliza la conversion de AAMM a ``AAAA/MM Mes`` para que el Revisor y los
comparadores lean y escriban exactamente en el mismo lugar. Las carpetas planas
del formato anterior solo se detectan para avisar: nunca se reutilizan.
"""

from pathlib import Path
import re
from typing import List, Optional, Tuple
import unicodedata


MESES = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)


def raiz_salidas(dir_script) -> Path:
    """Devuelve ``00_Salidas``, hermana de la carpeta del script."""
    return Path(dir_script).parent / "00_Salidas"


def partir_aamm(aamm) -> Optional[Tuple[str, int]]:
    """Convierte ``'2407'`` en ``('2024', 7)``; None si no es AAMM valido."""
    texto = str(aamm or "").strip()
    if not re.fullmatch(r"\d{4}", texto):
        return None
    mes = int(texto[2:])
    if not 1 <= mes <= 12:
        return None
    return "20" + texto[:2], mes


def nombre_carpeta_mes(aamm) -> Optional[str]:
    """Devuelve el nombre canonico ``MM Mes``; None si AAMM no es valido."""
    partes = partir_aamm(aamm)
    if partes is None:
        return None
    _, mes = partes
    return f"{mes:02d} {MESES[mes - 1]}"


def carpeta_anio(dir_salidas, aamm) -> Optional[Path]:
    """Devuelve ``00_Salidas/AAAA``; None si AAMM no es valido."""
    partes = partir_aamm(aamm)
    if partes is None:
        return None
    anio, _ = partes
    return Path(dir_salidas) / anio


def _normalizar(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split()).upper()


def _es_variante_mes(nombre, mes) -> bool:
    match = re.fullmatch(r"(\d{1,2})\s+(.+)", str(nombre).strip())
    return bool(
        match
        and int(match.group(1)) == mes
        and _normalizar(match.group(2)) == _normalizar(MESES[mes - 1])
    )


def carpeta_mes(dir_salidas, aamm, crear=False) -> Path:
    """Devuelve la ruta mensual nueva, o una variante ya existente.

    Con ``crear=True`` crea siempre la ruta canonica ``AAAA/MM Mes``. Para un
    valor no valido (incluido ``sin_mes``) conserva la ruta plana historica.
    Las carpetas planas AAMM del formato anterior nunca se devuelven.
    """
    raiz = Path(dir_salidas)
    partes = partir_aamm(aamm)
    if partes is None:
        ruta = raiz / str(aamm).strip()
        if crear:
            ruta.mkdir(parents=True, exist_ok=True)
        return ruta

    anio, mes = partes
    canonica = raiz / anio / f"{mes:02d} {MESES[mes - 1]}"
    if crear:
        canonica.mkdir(parents=True, exist_ok=True)
        return canonica
    if canonica.exists():
        return canonica

    carpeta_del_anio = raiz / anio
    if carpeta_del_anio.is_dir():
        variantes = sorted(
            (p for p in carpeta_del_anio.iterdir()
             if p.is_dir() and _es_variante_mes(p.name, mes)),
            key=lambda p: _normalizar(p.name),
        )
        if variantes:
            return variantes[0]
    return canonica


def carpeta_comparador(dir_salidas, anio, nombre) -> Path:
    """Devuelve una carpeta de comparador dentro del año indicado."""
    return Path(dir_salidas) / str(anio).strip() / str(nombre)


def carpetas_legado(dir_salidas) -> List[Path]:
    """Lista las carpetas planas AAMM del formato anterior que aun existen."""
    raiz = Path(dir_salidas)
    if not raiz.is_dir():
        return []
    return sorted(
        (p for p in raiz.iterdir() if p.is_dir() and partir_aamm(p.name)),
        key=lambda p: p.name,
    )
