#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera INTERFACES.md recorriendo los .py del repositorio.

Extrae firmas de funciones, clases, constantes de modulo y dependencias
internas, SIN los cuerpos. Es lo que reemplaza pegar codigo completo en un
chat: para conectar algo nuevo a un script existente no hace falta el script
entero, basta saber que recibe y que devuelve.

Esta ajustado a como documentan los scripts de este repositorio:

  * El encabezado del archivo puede ser un docstring o un bloque de comentarios
    "#" arriba de todo. Los dos se usan igual.
  * Las lineas de guiones o iguales son separadores y se descartan.
  * Un banner de seccion ("# === UTILIDADES ===") no es la descripcion de la
    funcion que viene abajo: se muestra como divisor.
  * Cuando una funcion no tiene docstring se usa el comentario que tenga encima.

Solo biblioteca estandar. Requiere Python 3.9 o superior.

Uso:

    python generar_interfaces.py                    genera/actualiza INTERFACES.md
    python generar_interfaces.py --check            no escribe; sale con 1 si esta desactualizado
    python generar_interfaces.py --privadas         incluye funciones que empiezan con "_"
    python generar_interfaces.py --esqueleto-mapa   imprime bloques MAPA.md para los .py que faltan
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

# Carpetas que se recorren, en este orden. Las que no existan se saltan, y un
# archivo ya visto no se repite: "scripts/comun" va primero solo para que el
# modulo compartido quede arriba en el indice.
CARPETAS = ["Revisor_Relq/comun", "Revisor_Relq"]

# Archivos que nunca entran a INTERFACES.md. __init__.py de un paquete
# es solo el "que es esta carpeta"; su contenido ya esta en MAPA.md.
EXCLUIDOS = {"generar_interfaces.py", "__init__.py"}

# Prefijos de archivo que tampoco entran (las pruebas no son interfaz).
PREFIJOS_EXCLUIDOS = ("test_",)

# Largo maximo del valor de una constante antes de resumirlo.
MAX_VALOR = 90

# Largo maximo de la nota de una constante en la tabla.
MAX_NOTA = 200

# Cuantas lineas del encabezado del archivo se muestran antes de cortar.
MAX_LINEAS_ENCABEZADO = 15

# Cuantas lineas puede ocupar una firma antes de que un comentario dentro del
# cuerpo deje de considerarse descripcion de la funcion.
LINEAS_FIRMA_MAX = 6

# Caracteres con los que estan hechas las lineas separadoras de los banners.
# Ademas de los ASCII van los de dibujo de cajas (U+2500..U+257F), que es lo
# que usa Actualiza_datos.py para sus separadores.
CHARS_SEPARADOR = set("=-*_~#<> ") | {chr(c) for c in range(0x2500, 0x2580)}

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

- Del encabezado de cada archivo se muestran las primeras líneas; el resto está
  arriba de todo en el `.py`.
- Las funciones que empiezan con `_` son internas del archivo.
- Cuando una función no tiene docstring se muestra el comentario que tenga encima.
- Los `— TÍTULO —` son los banners de sección del propio archivo.
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


def documentable(ruta: Path) -> bool:
    return (ruta.name not in EXCLUIDOS
            and not ruta.name.startswith(PREFIJOS_EXCLUIDOS)
            and "__pycache__" not in ruta.parts)


def archivos_py(raiz: Path) -> list[Path]:
    """Los .py a documentar, en el orden de CARPETAS y alfabetico dentro.

    Un archivo que ya salio en una carpeta anterior no se repite: las carpetas
    de CARPETAS pueden estar una dentro de otra.
    """
    encontrados: list[Path] = []
    vistos: set[Path] = set()
    for carpeta in list(CARPETAS) + ["."]:
        base = raiz / carpeta
        if not base.is_dir():
            continue
        candidatos = (sorted(base.glob("*.py")) if carpeta == "."
                      else sorted(base.rglob("*.py")))
        for ruta in candidatos:
            if ruta not in vistos and documentable(ruta):
                vistos.add(ruta)
                encontrados.append(ruta)
    return encontrados


# ---------------------------------------------------------------------------
# Comentarios: separadores, banners y limpieza
# ---------------------------------------------------------------------------

def partir_banner(linea: str) -> tuple[str, bool]:
    """Separa el titulo de un banner. Devuelve (titulo, es_banner).

    Reconoce los dos estilos que hay en el repo: la fila sola de separadores
    ("# ====="), que devuelve titulo vacio, y el que lleva el texto adentro
    ("# -- Titulo -------"), que devuelve el titulo suelto.
    """
    s = linea.strip()
    if not s:
        return "", False
    izq = 0
    while izq < len(s) and s[izq] in CHARS_SEPARADOR:
        izq += 1
    der = len(s)
    while der > izq and s[der - 1] in CHARS_SEPARADOR:
        der -= 1
    nucleo = s[izq:der].strip()
    if not nucleo:
        return "", len(s) >= 3
    if izq >= 3 or (len(s) - der) >= 3:
        return nucleo, True
    return s, False


def es_separador(linea: str) -> bool:
    """True si la linea es solo una fila de separadores, sin texto."""
    titulo, banner = partir_banner(linea)
    return banner and not titulo


def limpiar_comentario(crudas: list[str]) -> tuple[list[str], bool]:
    """Saca las lineas separadoras. Devuelve (lineas utiles, habia separadores)."""
    utiles = [ln for ln in crudas if not es_separador(ln)]
    while utiles and not utiles[0].strip():
        utiles.pop(0)
    while utiles and not utiles[-1].strip():
        utiles.pop()
    return utiles, len(utiles) != len(crudas)


def es_banner_seccion(utiles: list[str], habia_separadores: bool) -> bool:
    """True si el comentario es un titulo de seccion y no una descripcion.

    Lo es cuando queda una sola linea, y esa linea o venia encerrada entre
    filas de separadores o los trae adentro. Ese comentario divide el archivo;
    no describe la funcion que viene abajo.
    """
    con_texto = [ln for ln in utiles if ln.strip()]
    if len(con_texto) != 1:
        return False
    return habia_separadores or partir_banner(con_texto[0])[1]


def titulo_seccion(utiles: list[str]) -> str:
    """El texto del banner, ya sin los separadores de los costados."""
    return partir_banner(utiles[0])[0] or utiles[0].strip()


def es_titulo_mayusculas(linea: str) -> bool:
    """True si la linea es un subtitulo en MAYUSCULAS dentro de un encabezado."""
    limpia = linea.strip()
    if not (3 <= len(limpia) <= 60):
        return False
    letras = [c for c in limpia if c.isalpha()]
    return bool(letras) and all(c.isupper() for c in letras)


def sin_sangria_comun(lineas: list[str]) -> list[str]:
    """Saca la sangria que comparten todas las lineas con texto."""
    con_texto = [ln for ln in lineas if ln.strip()]
    if not con_texto:
        return lineas
    comun = min(len(ln) - len(ln.lstrip()) for ln in con_texto)
    return [ln[comun:] if ln.strip() else "" for ln in lineas]


def primera_frase(texto: str, largo: int = MAX_NOTA) -> str:
    """Primera oracion del texto, en una linea, recortada a `largo`."""
    plano = " ".join(texto.split())
    if not plano:
        return ""
    corte = re.search(r"(?<=\.)\s", plano)
    if corte and corte.start() + 1 <= largo:
        plano = plano[: corte.start() + 1]
    if len(plano) > largo:
        plano = plano[: largo - 1].rstrip() + "…"
    return plano


# ---------------------------------------------------------------------------
# Extraccion
# ---------------------------------------------------------------------------

def firma(nodo: ast.AST) -> str:
    """Arma la linea `def nombre(args) -> retorno` sin el cuerpo."""
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


def comentario_previo(lineas: list[str], lineno: int) -> list[str]:
    """Lineas `#` contiguas justo encima de `lineno` (1-based), sin el `#`."""
    juntadas: list[str] = []
    i = lineno - 2  # indice 0-based de la linea anterior
    while i >= 0:
        cruda = lineas[i].strip()
        if cruda.startswith("#"):
            juntadas.append(cruda.lstrip("#").rstrip())
            i -= 1
        else:
            break
    return list(reversed(juntadas))


def descripcion(nodo, lineas: list[str]) -> tuple[str, str]:
    """Devuelve (descripcion, banner de seccion) del nodo.

    La descripcion sale del docstring; si no hay, del comentario que tenga
    encima. Si ese comentario resulta ser un banner de seccion se devuelve
    aparte, porque no describe a la funcion.
    """
    doc = ast.get_docstring(nodo, clean=True)
    if doc:
        return doc.strip(), ""

    decos = getattr(nodo, "decorator_list", None)
    inicio = min([d.lineno for d in decos] + [nodo.lineno]) if decos else nodo.lineno
    crudas = comentario_previo(lineas, inicio)
    utiles, hubo_sep = limpiar_comentario(crudas)

    if utiles:
        if es_banner_seccion(utiles, hubo_sep):
            return "", titulo_seccion(utiles)
        return "\n".join(ln.strip() for ln in utiles).strip(), ""

    # Muchas funciones documentan con # dentro del cuerpo, no con docstring.
    # Solo vale si el comentario arranca pegado a la firma; mas abajo ya es un
    # comentario de la primera instruccion y no describe a la funcion.
    cuerpo = getattr(nodo, "body", None)
    if cuerpo:
        primera = cuerpo[0].lineno
        crudas_int = comentario_previo(lineas, primera)
        utiles_int, hubo_sep_int = limpiar_comentario(crudas_int)
        if utiles_int and not es_banner_seccion(utiles_int, hubo_sep_int):
            if primera - nodo.lineno <= LINEAS_FIRMA_MAX + len(crudas_int):
                return "\n".join(ln.strip() for ln in utiles_int).strip(), ""
    return "", ""


def nota_constante(lineas: list[str], lineno: int) -> tuple[str, str]:
    """Devuelve (nota corta para la tabla, banner de seccion) de una constante."""
    crudas = comentario_previo(lineas, lineno)
    utiles, hubo_sep = limpiar_comentario(crudas)
    if not utiles:
        return "", ""
    if es_banner_seccion(utiles, hubo_sep):
        return "", titulo_seccion(utiles)
    return primera_frase(" ".join(ln.strip() for ln in utiles)), ""


def encabezado_modulo(arbol: ast.Module, lineas: list[str]) -> tuple[str, bool]:
    """Descripcion del archivo, del docstring o del bloque `#` de arriba de todo.

    Devuelve (texto, se_corto). Se corta en el primer subtitulo en MAYUSCULAS o
    a las MAX_LINEAS_ENCABEZADO lineas: el encabezado completo se lee en el .py.
    """
    doc = ast.get_docstring(arbol, clean=True)
    if doc:
        crudas = doc.splitlines()
    else:
        crudas = []
        for cruda in lineas:
            limpia = cruda.strip()
            if not limpia:
                if crudas:
                    break
                continue
            if limpia.startswith("#!") or "coding" in limpia[:20]:
                continue
            if limpia.startswith("#"):
                crudas.append(limpia.lstrip("#").rstrip())
            else:
                break

    utiles, _ = limpiar_comentario(crudas)
    # Un archivo puede arrancar directo con un banner de seccion y no tener
    # encabezado propio (Actualiza_datos.py). Ese banner no lo describe.
    while utiles and partir_banner(utiles[0])[1]:
        utiles.pop(0)
    if not utiles:
        return "", False
    utiles = sin_sangria_comun(utiles)

    recorte: list[str] = []
    for ln in utiles:
        if recorte and es_titulo_mayusculas(ln):
            return "\n".join(recorte).strip(), True
        recorte.append(ln.rstrip())
        if len(recorte) >= MAX_LINEAS_ENCABEZADO:
            return "\n".join(recorte).strip(), len(recorte) < len(utiles)
    return "\n".join(recorte).strip(), False


def resumen_valor(valor: ast.AST) -> str:
    """Representacion corta del valor de una constante de modulo."""
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
        return f"{clase} de {len(valor.elts)} elementos: {', '.join(muestras)}, …"
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
    """Constante de modulo: MAYUSCULAS_CON_GUION_BAJO, al menos dos caracteres."""
    return len(nombre) > 1 and nombre.isupper() and not nombre.startswith("__")


def analizar(ruta: Path, raiz: Path, incluir_privadas: bool) -> dict:
    texto = leer_texto(ruta)
    lineas = texto.splitlines()
    relativa = ruta.relative_to(raiz).as_posix()
    vacio = {"ruta": relativa, "doc": "", "doc_cortado": False, "constantes": [],
             "miembros": [], "imports": set(), "lineas": len(lineas)}

    try:
        arbol = ast.parse(texto, filename=str(ruta))
    except SyntaxError as err:
        return {**vacio,
                "error": f"no se pudo interpretar: {err.msg} (linea {err.lineno})"}

    doc, cortado = encabezado_modulo(arbol, lineas)

    constantes: list[dict] = []
    miembros: list[dict] = []
    imports: set[str] = set()

    for nodo in arbol.body:
        if isinstance(nodo, ast.Import):
            for alias in nodo.names:
                imports.add(alias.name.split(".")[0])

        elif isinstance(nodo, ast.ImportFrom):
            if nodo.module:
                imports.add(nodo.module.split(".")[0])

        elif isinstance(nodo, (ast.Assign, ast.AnnAssign)):
            destinos = nodo.targets if isinstance(nodo, ast.Assign) else [nodo.target]
            for destino in destinos:
                if isinstance(destino, ast.Name) and es_constante(destino.id):
                    nota, seccion = nota_constante(lineas, nodo.lineno)
                    valor = resumen_valor(nodo.value) if nodo.value else ""
                    constantes.append({"nombre": destino.id, "valor": valor,
                                       "nota": nota, "seccion": seccion})

        elif isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
            desc, seccion = descripcion(nodo, lineas)
            if nodo.name.startswith("_") and not incluir_privadas:
                continue
            miembros.append({"tipo": "funcion", "privada": nodo.name.startswith("_"),
                             "firma": firma(nodo), "decoradores": decoradores(nodo),
                             "doc": desc, "seccion": seccion, "metodos": []})

        elif isinstance(nodo, ast.ClassDef):
            metodos = []
            for hijo in nodo.body:
                if isinstance(hijo, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    if hijo.name.startswith("_") and hijo.name != "__init__":
                        if not incluir_privadas:
                            continue
                    doc_met, _ = descripcion(hijo, lineas)
                    metodos.append({"firma": firma(hijo), "doc": doc_met})
            desc, seccion = descripcion(nodo, lineas)
            miembros.append({"tipo": "clase", "privada": nodo.name.startswith("_"),
                             "firma": firma(nodo), "decoradores": decoradores(nodo),
                             "doc": desc, "seccion": seccion, "metodos": metodos})

    return {**vacio, "error": "", "doc": doc, "doc_cortado": cortado,
            "constantes": constantes, "miembros": miembros, "imports": imports}


# ---------------------------------------------------------------------------
# Render
# ---------------------------------------------------------------------------

def ancla_github(titulo: str) -> str:
    """Ancla que GitHub genera para un encabezado `titulo` entre backticks.

    GitHub conserva las letras con tilde (son letras, no puntuacion): solo saca
    los signos y el "/" de las rutas. Por eso NO se puede filtrar con un regex
    ASCII tipo [^a-z0-9]: eso le borraria la tilde a "Reliquidación" y el link
    del indice quedaria apuntando a un ancla que no existe.
    """
    limpio = titulo.lower().replace("`", "")
    limpio = "".join(c for c in limpio if c.isalnum() or c in " _-")
    return limpio.replace(" ", "-")


def escapar_tabla(texto: str) -> str:
    return texto.replace("|", "\\|")


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
        resumen = primera_frase(m["doc"].splitlines()[0] if m["doc"] else "", 80)
        linea = f"- [`{m['ruta']}`](#{ancla_github(m['ruta'])}) — {m['lineas']} líneas"
        partes.append(f"{linea} — {resumen}" if resumen else linea)
    partes.append("")

    for m in modulos:
        partes.append("\n---\n")
        partes.append(f"## `{m['ruta']}`\n")

        if m["error"]:
            partes.append(f"> **No se pudo analizar** — {m['error']}\n")
            continue

        if m["doc"]:
            partes.extend(f"> {ln}".rstrip() for ln in m["doc"].splitlines())
            if m["doc_cortado"]:
                partes.append(">")
                partes.append("> *(el encabezado sigue arriba de todo en el archivo)*")
            partes.append("")

        propios = sorted(m["imports"] & nombres_modulo)
        if propios:
            partes.append("**Depende de (del repo):** "
                          + ", ".join(f"`{d}`" for d in propios) + "\n")
        externos = sorted(d for d in m["imports"] if d not in nombres_modulo)
        if externos:
            partes.append("**Importa:** " + ", ".join(f"`{d}`" for d in externos) + "\n")

        if m["constantes"]:
            partes.append("### Constantes\n")
            partes.append("| Nombre | Valor | |")
            partes.append("|---|---|---|")
            for c in m["constantes"]:
                if c["seccion"]:
                    partes.append(f"| **— {escapar_tabla(c['seccion'])} —** | | |")
                valor = f"`{escapar_tabla(c['valor'])}`" if c["valor"] else ""
                partes.append(f"| `{c['nombre']}` | {valor} | "
                              f"{escapar_tabla(c['nota'])} |")
            partes.append("")

        clases = [x for x in m["miembros"] if x["tipo"] == "clase"]
        funciones = [x for x in m["miembros"] if x["tipo"] == "funcion"]

        if clases:
            partes.append("### Clases\n")
            for c in clases:
                if c["seccion"]:
                    partes.append(f"**— {c['seccion']} —**\n")
                for deco in c["decoradores"]:
                    partes.append(f"`{deco}`")
                partes.append(f"#### `{c['firma']}`\n")
                if c["doc"]:
                    partes.extend(c["doc"].splitlines())
                    partes.append("")
                for met in c["metodos"]:
                    linea = f"- `{met['firma']}`"
                    if met["doc"]:
                        linea += f" — {primera_frase(met['doc'], 120)}"
                    partes.append(linea)
                partes.append("")

        if funciones:
            partes.append("### Funciones\n")
            for f in funciones:
                if f["seccion"]:
                    partes.append(f"**— {f['seccion']} —**\n")
                for deco in f["decoradores"]:
                    partes.append(f"`{deco}`")
                marca = " *(interna)*" if f["privada"] else ""
                partes.append(f"#### `{f['firma']}`{marca}\n")
                if f["doc"]:
                    partes.extend(f["doc"].splitlines())
                    partes.append("")

        if not m["constantes"] and not m["miembros"]:
            partes.append("*Sin funciones, clases ni constantes de módulo.*\n")

    duplicadas = constantes_duplicadas(modulos)
    if duplicadas:
        partes.append("\n---\n")
        partes.append("## Constantes definidas en más de un archivo\n")
        partes.append(
            "Cada una es un punto donde un cambio hay que hacerlo en varios lados a "
            "la vez. Candidatas a mudarse a `comun/`.\n"
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
        for c in m["constantes"]:
            donde.setdefault(c["nombre"], []).append(m["ruta"])
    return sorted((n, r) for n, r in donde.items() if len(r) > 1)


def esqueleto_mapa(modulos: list[dict], raiz: Path) -> str:
    """Bloques MAPA.md para los .py que todavia no aparecen en MAPA.md."""
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

    par = argparse.ArgumentParser(description="Genera INTERFACES.md.")
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

    mapa = raiz / "MAPA.md"
    if mapa.exists():
        texto_mapa = leer_texto(mapa)
        faltan = [m["ruta"] for m in modulos
                  if Path(m["ruta"]).name not in texto_mapa]
        if faltan:
            print(f"Aviso: {len(faltan)} archivo(s) sin bloque en MAPA.md: "
                  f"{', '.join(faltan)}")
            print("   Correr con --esqueleto-mapa para tener los bloques listos.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
