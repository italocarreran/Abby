#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera INTERFACES.md recorriendo los .py del repositorio.

Extrae firmas de funciones, clases, constantes de módulo y dependencias
internas, SIN los cuerpos. Es lo que reemplaza pegar código completo en un
chat: para conectar algo nuevo a un script existente no hace falta el script
entero, basta saber qué recibe y qué devuelve.

Solo biblioteca estándar. Requiere Python 3.9 o superior.

Uso:

    python generar_interfaces.py                 genera/actualiza INTERFACES.md
    python generar_interfaces.py --check         no escribe; sale con 1 si está desactualizado
    python generar_interfaces.py --privadas      incluye funciones que empiezan con "_"
    python generar_interfaces.py --esqueleto-mapa   imprime bloques MAPA.md para los .py que faltan
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

# Carpetas que se recorren, en este orden. Las que no existan se saltan.
CARPETAS = ["comun", "scripts", "reemplazos_reuc"]

# Archivos que nunca entran a INTERFACES.md.
EXCLUIDOS = {"generar_interfaces.py"}

# Largo máximo del valor de una constante antes de resumirlo.
MAX_VALOR = 90

# Cuántas líneas puede ocupar una firma antes de que un comentario dentro del
# cuerpo deje de considerarse descripción de la función.
LINEAS_FIRMA_MAX = 6

SALIDA_POR_OMISION = "INTERFACES.md"

ENCABEZADO = """<!-- ARCHIVO GENERADO POR generar_interfaces.py — NO EDITAR A MANO -->
<!-- Para regenerarlo: python generar_interfaces.py -->

# INTERFACES — firmas y contratos

Firmas de funciones, clases y constantes de cada `.py` del repositorio, **sin los
cuerpos**. Sirve para conectar código nuevo con el existente sin abrir los archivos
completos.

> **Regla de expansión.** Leer `MAPA.md` → leer acá solo las entradas que hacen falta
> → abrir completo **únicamente** el archivo que se va a modificar. No abrir los
> archivos vecinos "para tener contexto". Si de verdad hace falta uno más, pedirlo
> explícitamente y decir por qué.

Convenciones de esta página:

- Las funciones que empiezan con `_` son internas del archivo.
- Cuando una función no tiene docstring se muestra el comentario `#` que tenga
  justo encima, si lo tiene.
- Los valores de las constantes largas salen resumidos; el valor exacto está en el
  archivo.

"""


# ---------------------------------------------------------------------------
# Lectura
# ---------------------------------------------------------------------------

def leer_texto(ruta: Path) -> str:
    """Lee un .py probando las codificaciones que aparecen en equipos Windows."""
    datos = ruta.read_bytes()
    for cod in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return datos.decode(cod)
        except UnicodeDecodeError:
            continue
    return datos.decode("utf-8", errors="replace")


def archivos_py(raiz: Path) -> list[Path]:
    """Devuelve los .py a documentar, en el orden de CARPETAS y alfabético dentro."""
    encontrados: list[Path] = []
    for carpeta in CARPETAS:
        base = raiz / carpeta
        if not base.is_dir():
            continue
        for ruta in sorted(base.rglob("*.py")):
            if ruta.name in EXCLUIDOS:
                continue
            encontrados.append(ruta)
    for ruta in sorted(raiz.glob("*.py")):
        if ruta.name not in EXCLUIDOS:
            encontrados.append(ruta)
    return encontrados


# ---------------------------------------------------------------------------
# Extracción
# ---------------------------------------------------------------------------

def firma(nodo: ast.AST) -> str:
    """Arma la línea `def nombre(args) -> retorno:` sin el cuerpo."""
    if isinstance(nodo, ast.ClassDef):
        bases = [ast.unparse(b) for b in nodo.bases]
        bases += [f"{kw.arg}={ast.unparse(kw.value)}" for kw in nodo.keywords]
        return f"class {nodo.name}" + (f"({', '.join(bases)})" if bases else "")

    prefijo = "async def " if isinstance(nodo, ast.AsyncFunctionDef) else "def "
    args = ast.unparse(nodo.args)
    retorno = f" -> {ast.unparse(nodo.returns)}" if nodo.returns else ""
    return f"{prefijo}{nodo.name}({args}){retorno}"


def decoradores(nodo) -> list[str]:
    return ["@" + ast.unparse(d) for d in getattr(nodo, "decorator_list", [])]


def comentario_previo(lineas: list[str], lineno: int) -> str:
    """Junta las líneas `#` contiguas que estén justo encima de `lineno` (1-based).

    Es el reemplazo del docstring cuando el archivo documenta con comentarios.
    """
    juntadas: list[str] = []
    i = lineno - 2  # índice 0-based de la línea anterior
    while i >= 0:
        cruda = lineas[i].strip()
        if cruda.startswith("#"):
            juntadas.append(cruda.lstrip("#").strip())
            i -= 1
        elif cruda == "" and juntadas:
            break
        else:
            break
    return " ".join(reversed(juntadas)).strip()


def descripcion(nodo, lineas: list[str], completa: bool) -> str:
    """Docstring del nodo; si no tiene, el comentario que lo precede."""
    doc = ast.get_docstring(nodo, clean=True)
    if doc:
        doc = doc.strip()
        if not completa:
            doc = doc.split("\n\n", 1)[0].strip()
        return doc
    decos = getattr(nodo, "decorator_list", None)
    inicio = min([d.lineno for d in decos] + [nodo.lineno]) if decos else nodo.lineno
    previo = comentario_previo(lineas, inicio)
    if previo:
        return previo

    # Muchos scripts documentan con # dentro del cuerpo, no con docstring.
    # Solo se toma si el comentario arranca pegado a la firma; si no, es un
    # comentario de la primera instrucción y no describe la función.
    cuerpo = getattr(nodo, "body", None)
    if cuerpo:
        primera = cuerpo[0].lineno
        interno = comentario_previo(lineas, primera)
        if interno and primera - nodo.lineno <= LINEAS_FIRMA_MAX + len(interno.splitlines()):
            return interno
    return ""


def resumen_valor(valor: ast.AST) -> str:
    """Representación corta del valor de una constante de módulo."""
    try:
        texto = ast.unparse(valor)
    except Exception:
        return "…"
    texto = " ".join(texto.split())
    if len(texto) <= MAX_VALOR:
        return texto
    if isinstance(valor, (ast.List, ast.Tuple, ast.Set)):
        clase = {ast.List: "lista", ast.Tuple: "tupla", ast.Set: "conjunto"}[type(valor)]
        muestras = []
        for elemento in valor.elts[:3]:
            try:
                muestras.append(ast.unparse(elemento))
            except Exception:
                break
        cola = ", ".join(muestras)
        return f"{clase} de {len(valor.elts)} elementos: {cola}, …"
    if isinstance(valor, ast.Dict):
        claves = []
        for clave in valor.keys[:3]:
            try:
                claves.append(ast.unparse(clave))
            except Exception:
                break
        return f"dict de {len(valor.keys)} claves: {', '.join(claves)}, …"
    return texto[: MAX_VALOR - 1] + "…"


def es_constante(nombre: str) -> bool:
    """Constante de módulo: MAYUSCULAS_CON_GUION_BAJO, al menos dos caracteres."""
    return len(nombre) > 1 and nombre.isupper() and not nombre.startswith("__")


def analizar(ruta: Path, raiz: Path, incluir_privadas: bool) -> dict:
    texto = leer_texto(ruta)
    lineas = texto.splitlines()
    relativa = ruta.relative_to(raiz).as_posix()

    try:
        arbol = ast.parse(texto, filename=str(ruta))
    except SyntaxError as err:
        return {
            "ruta": relativa,
            "error": f"no se pudo interpretar: {err.msg} (línea {err.lineno})",
            "doc": "",
            "constantes": [],
            "miembros": [],
            "imports": set(),
            "lineas": len(lineas),
        }

    doc = (ast.get_docstring(arbol, clean=True) or "").strip()

    constantes: list[tuple[str, str, str]] = []
    miembros: list[dict] = []
    imports: set[str] = set()

    for nodo in arbol.body:
        if isinstance(nodo, ast.Import):
            for alias in nodo.names:
                imports.add(alias.name.split(".")[0])

        elif isinstance(nodo, ast.ImportFrom):
            if nodo.level and nodo.level > 0:
                if nodo.module:
                    imports.add(nodo.module.split(".")[0])
            elif nodo.module:
                imports.add(nodo.module.split(".")[0])

        elif isinstance(nodo, ast.Assign):
            for destino in nodo.targets:
                if isinstance(destino, ast.Name) and es_constante(destino.id):
                    constantes.append(
                        (destino.id, resumen_valor(nodo.value),
                         comentario_previo(lineas, nodo.lineno))
                    )

        elif isinstance(nodo, ast.AnnAssign):
            if isinstance(nodo.target, ast.Name) and es_constante(nodo.target.id):
                valor = resumen_valor(nodo.value) if nodo.value else ""
                constantes.append(
                    (nodo.target.id, valor, comentario_previo(lineas, nodo.lineno))
                )

        elif isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if nodo.name.startswith("_") and not incluir_privadas:
                continue
            miembros.append({
                "tipo": "funcion",
                "privada": nodo.name.startswith("_"),
                "firma": firma(nodo),
                "decoradores": decoradores(nodo),
                "doc": descripcion(nodo, lineas, completa=not nodo.name.startswith("_")),
                "metodos": [],
            })

        elif isinstance(nodo, ast.ClassDef):
            metodos = []
            for hijo in nodo.body:
                if isinstance(hijo, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    if hijo.name.startswith("_") and hijo.name != "__init__":
                        if not incluir_privadas:
                            continue
                    metodos.append({
                        "firma": firma(hijo),
                        "decoradores": decoradores(hijo),
                        "doc": descripcion(hijo, lineas, completa=False),
                    })
            miembros.append({
                "tipo": "clase",
                "privada": nodo.name.startswith("_"),
                "firma": firma(nodo),
                "decoradores": decoradores(nodo),
                "doc": descripcion(nodo, lineas, completa=True),
                "metodos": metodos,
            })

    return {
        "ruta": relativa,
        "error": "",
        "doc": doc,
        "constantes": constantes,
        "miembros": miembros,
        "imports": imports,
        "lineas": len(lineas),
    }


# ---------------------------------------------------------------------------
# Render
# ---------------------------------------------------------------------------

def ancla_github(titulo: str) -> str:
    """Ancla que GitHub genera para un encabezado `titulo` entre backticks."""
    limpio = titulo.lower().replace("`", "")
    limpio = re.sub(r"[^a-z0-9 _-]", "", limpio)
    return limpio.replace(" ", "-")


def bloque_cita(texto: str, sangria: str = "") -> list[str]:
    if not texto:
        return []
    return [f"{sangria}{linea}".rstrip() for linea in texto.splitlines()]


def render(modulos: list[dict]) -> str:
    nombres_modulo = {Path(m["ruta"]).stem for m in modulos}
    partes: list[str] = [ENCABEZADO]

    if not modulos:
        partes.append(
            "## Todavía no hay `.py` en el repositorio\n\n"
            "En cuanto se suba el primer script a `scripts/`, `comun/` o "
            "`reemplazos_reuc/`, correr:\n\n"
            "```\npython generar_interfaces.py\n```\n\n"
            "Mientras tanto, lo que hace cada script está descrito en `MAPA.md`.\n"
        )
        return "\n".join(partes).rstrip() + "\n"

    partes.append("## Índice\n")
    for m in modulos:
        ancla = ancla_github(m["ruta"])
        partes.append(f"- [`{m['ruta']}`](#{ancla}) — {m['lineas']} líneas")
    partes.append("")

    for m in modulos:
        partes.append("\n---\n")
        partes.append(f"## `{m['ruta']}`\n")

        if m["error"]:
            partes.append(f"> **No se pudo analizar** — {m['error']}\n")
            continue

        if m["doc"]:
            partes.extend(bloque_cita(m["doc"]))
            partes.append("")

        propios = sorted(m["imports"] & nombres_modulo)
        if propios:
            partes.append("**Depende de (del repo):** " + ", ".join(f"`{d}`" for d in propios) + "\n")

        externos = sorted(d for d in m["imports"] if d not in nombres_modulo)
        if externos:
            partes.append("**Importa:** " + ", ".join(f"`{d}`" for d in externos) + "\n")

        if m["constantes"]:
            partes.append("### Constantes\n")
            partes.append("| Nombre | Valor | |")
            partes.append("|---|---|---|")
            for nombre, valor, nota in m["constantes"]:
                celda_valor = f"`{valor}`" if valor else ""
                celda_valor = celda_valor.replace("|", "\\|")
                nota = nota.replace("|", "\\|")
                partes.append(f"| `{nombre}` | {celda_valor} | {nota} |")
            partes.append("")

        funciones = [x for x in m["miembros"] if x["tipo"] == "funcion"]
        clases = [x for x in m["miembros"] if x["tipo"] == "clase"]

        if clases:
            partes.append("### Clases\n")
            for c in clases:
                for deco in c["decoradores"]:
                    partes.append(f"`{deco}`")
                partes.append(f"#### `{c['firma']}`\n")
                if c["doc"]:
                    partes.extend(bloque_cita(c["doc"]))
                    partes.append("")
                for met in c["metodos"]:
                    linea = f"- `{met['firma']}`"
                    if met["doc"]:
                        linea += f" — {met['doc'].splitlines()[0]}"
                    partes.append(linea)
                partes.append("")

        if funciones:
            partes.append("### Funciones\n")
            for f in funciones:
                marca = " *(interna)*" if f["privada"] else ""
                for deco in f["decoradores"]:
                    partes.append(f"`{deco}`")
                partes.append(f"#### `{f['firma']}`{marca}\n")
                if f["doc"]:
                    partes.extend(bloque_cita(f["doc"]))
                    partes.append("")

        if not m["constantes"] and not m["miembros"]:
            partes.append("*Sin funciones, clases ni constantes de módulo.*\n")

    duplicadas = constantes_duplicadas(modulos)
    if duplicadas:
        partes.append("\n---\n")
        partes.append("## Constantes definidas en más de un archivo\n")
        partes.append(
            "Cada una de estas es un punto donde un cambio hay que hacerlo en varios "
            "lados a la vez. Candidatas a mudarse a `comun/`.\n"
        )
        partes.append("| Constante | Archivos |")
        partes.append("|---|---|")
        for nombre, rutas in duplicadas:
            partes.append(f"| `{nombre}` | " + ", ".join(f"`{r}`" for r in rutas) + " |")
        partes.append("")

    return "\n".join(partes).rstrip() + "\n"


def constantes_duplicadas(modulos: list[dict]) -> list[tuple[str, list[str]]]:
    donde: dict[str, list[str]] = {}
    for m in modulos:
        for nombre, _valor, _nota in m["constantes"]:
            donde.setdefault(nombre, []).append(m["ruta"])
    return sorted((n, r) for n, r in donde.items() if len(r) > 1)


def esqueleto_mapa(modulos: list[dict], raiz: Path) -> str:
    """Bloques MAPA.md para los .py que todavía no aparecen en MAPA.md."""
    mapa = raiz / "MAPA.md"
    texto_mapa = leer_texto(mapa) if mapa.exists() else ""
    faltantes = [m for m in modulos if Path(m["ruta"]).name not in texto_mapa]
    if not faltantes:
        return "MAPA.md ya menciona todos los .py del repositorio.\n"
    salida = ["Falta el bloque de MAPA.md para estos archivos:\n"]
    for m in faltantes:
        salida.append(f"""### `{m['ruta']}`

- **Qué hace:** …
- **Consume:** …
- **Produce:** …
- **Expone:** …
- **Depende de:** …
""")
    return "\n".join(salida)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    if sys.version_info < (3, 9):
        print("Hace falta Python 3.9 o superior (usa ast.unparse).", file=sys.stderr)
        return 2

    par = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    par.add_argument("--raiz", default=".", help="raíz del repositorio (por omisión: .)")
    par.add_argument("--salida", default=SALIDA_POR_OMISION, help="archivo a escribir")
    par.add_argument("--check", action="store_true",
                     help="no escribe; sale con 1 si el archivo está desactualizado")
    par.add_argument("--privadas", action="store_true",
                     help="incluye funciones y métodos que empiezan con _")
    par.add_argument("--esqueleto-mapa", action="store_true",
                     help="imprime bloques MAPA.md para los .py que falten ahí")
    args = par.parse_args(argv)

    raiz = Path(args.raiz).resolve()
    rutas = archivos_py(raiz)

    if not rutas:
        print(f"No hay .py en {', '.join(CARPETAS)} bajo {raiz}.")
        print("Se escribe igual un INTERFACES.md vacío para que quede el archivo.")

    modulos = [analizar(r, raiz, args.privadas) for r in rutas]
    contenido = render(modulos)

    if args.esqueleto_mapa:
        print(esqueleto_mapa(modulos, raiz))
        return 0

    destino = raiz / args.salida

    if args.check:
        actual = leer_texto(destino) if destino.exists() else ""
        if actual == contenido:
            print(f"{args.salida} está al día ({len(modulos)} archivo(s)).")
            return 0
        print(f"{args.salida} está DESACTUALIZADO. Correr: python generar_interfaces.py",
              file=sys.stderr)
        return 1

    destino.write_text(contenido, encoding="utf-8")
    print(f"{args.salida} escrito: {len(modulos)} archivo(s), "
          f"{sum(len(m['miembros']) for m in modulos)} funciones/clases.")

    duplicadas = constantes_duplicadas(modulos)
    if duplicadas:
        print(f"Aviso: {len(duplicadas)} constante(s) definidas en más de un archivo:")
        for nombre, donde in duplicadas:
            print(f"   {nombre}: {', '.join(donde)}")

    faltan = [m["ruta"] for m in modulos
              if (raiz / "MAPA.md").exists()
              and Path(m["ruta"]).name not in leer_texto(raiz / "MAPA.md")]
    if faltan:
        print(f"Aviso: {len(faltan)} archivo(s) sin bloque en MAPA.md: {', '.join(faltan)}")
        print("   Correr con --esqueleto-mapa para tener los bloques listos.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
