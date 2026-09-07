"""Lectura rápida de columnas OOXML sin abrir Excel.

Concentra la implementación idéntica que tenían el Revisor y
``Actualiza_Data_Access.py``. Solo lee resultados calculados; nunca evalúa el
nodo de fórmula. Ante un formato no soportado, una hoja ausente o un error de
lectura devuelve ``None`` para que el consumidor conserve su fallback actual a
xlwings/COM.
"""

import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from __comun__ import texto as _texto


NS_XL = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_ENT_XML = {"lt": "<", "gt": ">", "quot": '"', "apos": "'", "amp": "&"}
_RE_ENT = re.compile(r"&(?:#(\d+)|#[xX]([0-9a-fA-F]+)|(lt|gt|quot|apos|amp));")


def _normalizar(texto):
    # La copia histórica usaba str(texto), incluso para None.
    return _texto.suave_textual(texto)


def _col_letra_a_num(letra):
    n = 0
    for c in str(letra).upper():
        if "A" <= c <= "Z":
            n = n * 26 + (ord(c) - ord("A") + 1)
    return n


def es_zip_excel(ruta):
    """Indica si la extensión puede contener un libro OOXML."""
    return Path(ruta).suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm")


def ubicar_hoja_xml(z, hoja):
    """Devuelve ``(ruta_xml, hojas)``; acepta nombre normalizado o ``#N``."""
    nombres = set(z.namelist())
    if "xl/workbook.xml" not in nombres:
        return None, []
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    lista = list(wb.iter(f"{NS_XL}sheet"))
    hojas = [sh.get("name", "") for sh in lista]
    rid = None
    m_pos = re.fullmatch(r"#(\d+)", str(hoja).strip())
    if m_pos:
        i = int(m_pos.group(1)) - 1
        if 0 <= i < len(lista):
            rid = lista[i].get(f"{NS_REL}id")
    else:
        for sh in lista:
            if _normalizar(sh.get("name", "")) == _normalizar(hoja):
                rid = sh.get(f"{NS_REL}id")
                break
    if rid is None or "xl/_rels/workbook.xml.rels" not in nombres:
        return None, hojas
    destino = None
    for relacion in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
        if relacion.get("Id") == rid:
            destino = relacion.get("Target")
            break
    if not destino:
        return None, hojas
    ruta_hoja = destino[1:] if destino.startswith("/") else "xl/" + destino
    ruta_hoja = ruta_hoja.replace("xl/xl/", "xl/")
    return (ruta_hoja if ruta_hoja in nombres else None), hojas


def expandir_columnas(rango):
    """Convierte ``CF:CI`` en ``[CF, CG, CH, CI]``."""
    partes = str(rango).replace(" ", "").replace("$", "").upper().split(":")
    a = _col_letra_a_num(partes[0])
    b = _col_letra_a_num(partes[-1]) if len(partes) > 1 else a
    letras = []
    for n in range(min(a, b), max(a, b) + 1):
        s, x = "", n
        while x > 0:
            x, resto = divmod(x - 1, 26)
            s = chr(ord("A") + resto) + s
        letras.append(s)
    return letras


def desescapar_xml(valor):
    """Desescapa entidades XML, incluidas referencias numéricas dec/hex."""
    if isinstance(valor, bytes):
        valor = valor.decode("utf-8", "ignore")

    def uno(m):
        dec, hexa, nombre = m.group(1), m.group(2), m.group(3)
        try:
            if dec is not None:
                return chr(int(dec))
            if hexa is not None:
                return chr(int(hexa, 16))
        except (ValueError, OverflowError):
            return m.group(0)
        return _ENT_XML[nombre]

    return _RE_ENT.sub(uno, valor)


def leer_columnas_rapido(ruta, hoja, columnas, fila_inicio, log):
    """Lee columnas calculadas como ``{COL: {fila: valor}}`` o devuelve None."""
    if not es_zip_excel(ruta):
        log(f"    ! {Path(ruta).name}: formato no soportado para lectura rápida")
        return None
    objetivo = [c.upper() for c in columnas]
    if not objetivo:
        return {}
    alternativas = b"|".join(sorted((c.encode() for c in objetivo),
                                     key=len, reverse=True))
    patron = re.compile(rb'<c r="(' + alternativas + rb')(\d+)"([^>]*?)(?:/>|>(.*?)</c>)',
                        re.S)
    fila_inicio = int(fila_inicio)
    datos = {c: {} for c in objetivo}
    try:
        with zipfile.ZipFile(str(ruta)) as z:
            ruta_hoja, hojas = ubicar_hoja_xml(z, hoja)
            if ruta_hoja is None:
                log(f"    ! La hoja '{hoja}' no existe. Hojas: {', '.join(hojas)}")
                return None
            compartidas = []
            if "xl/sharedStrings.xml" in set(z.namelist()):
                ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
                compartidas = ["".join(si.itertext()) for si in ss.iter(f"{NS_XL}si")]
            with z.open(ruta_hoja) as archivo:
                cola = b""
                while True:
                    trozo = archivo.read(1 << 20)
                    if not trozo:
                        break
                    buf = cola + trozo
                    fin = 0
                    for m in patron.finditer(buf):
                        fin = m.end()
                        col, nfila = m.group(1).decode(), int(m.group(2))
                        if nfila < fila_inicio:
                            continue
                        attrs, cuerpo = m.group(3), m.group(4) or b""
                        valor = None
                        if b"inlineStr" in attrs:
                            mv = re.search(rb"<is>(.*?)</is>", cuerpo, re.S)
                            if mv:
                                valor = desescapar_xml(re.sub(rb"<[^>]+>", b"", mv.group(1)))
                        else:
                            mv = re.search(rb"<v>(.*?)</v>", cuerpo, re.S)
                            if mv:
                                bruto = mv.group(1)
                                if b't="s"' in attrs:
                                    try:
                                        valor = compartidas[int(bruto)]
                                    except Exception:
                                        valor = None
                                elif b't="e"' in attrs or b't="str"' in attrs:
                                    valor = desescapar_xml(bruto)
                                elif b't="b"' in attrs:
                                    valor = bool(int(bruto))
                                else:
                                    try:
                                        valor = float(bruto)
                                    except ValueError:
                                        valor = desescapar_xml(bruto)
                        if valor is not None and not (
                            isinstance(valor, str) and not valor.strip()
                        ):
                            datos[col][nfila] = valor
                    cola = buf[fin:] if fin else buf[-8192:]
    except Exception as e:
        log(f"    ! No se pudo leer {Path(ruta).name}: {e}")
        return None
    return datos
