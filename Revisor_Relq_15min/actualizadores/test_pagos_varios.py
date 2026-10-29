"""Pruebas de Actualiza_Pagos_Varios.py (solo biblioteca estandar).

Arma a mano tres .xlsx minimos con las columnas que dio la usuaria y comprueba
el CSV: encabezado, orden de las fuentes, fecha DD/MM/AAAA HH:MM, filtro de
pagos en 0 y de centrales vacias, bloque de la planilla 3 -> fecha, CRLF y sin
BOM. tkinter se reemplaza por un modulo falso: el script lo importa arriba,
pero estas pruebas no abren ventanas.

    python Revisor_Relq_15min/actualizadores/test_pagos_varios.py
"""

import csv
import importlib.util
import sys
import tempfile
import types
import unittest
import zipfile
from pathlib import Path
from unittest import mock

AQUI = Path(__file__).resolve().parent

_tk = types.ModuleType("tkinter")
for _n in ("filedialog", "messagebox", "ttk"):
    _m = mock.MagicMock()
    setattr(_tk, _n, _m)
    sys.modules.setdefault("tkinter." + _n, _m)
sys.modules.setdefault("tkinter", _tk)

_spec = importlib.util.spec_from_file_location("pagos_varios",
                                               AQUI / "Actualiza_Pagos_Varios.py")
pv = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pv)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _celda(ref, v):
    if isinstance(v, str):
        return f'<c r="{ref}" t="inlineStr"><is><t>{v}</t></is></c>'
    return f'<c r="{ref}"><v>{v!r}</v></c>'


def escribir_xlsx(ruta, hojas):
    """hojas: {nombre: {(fila, "COL"): valor}}"""
    with zipfile.ZipFile(ruta, "w") as z:
        sheets, rels = [], []
        for i, (nombre, celdas) in enumerate(hojas.items(), start=1):
            sheets.append(f'<sheet name="{nombre}" sheetId="{i}" r:id="rId{i}"/>')
            rels.append(f'<Relationship Id="rId{i}" Target="worksheets/sheet{i}.xml"/>')
            filas = {}
            for (f, c), v in celdas.items():
                filas.setdefault(f, []).append(_celda(f"{c}{f}", v))
            datos = "".join(f'<row r="{f}">{"".join(cs)}</row>'
                            for f, cs in sorted(filas.items()))
            z.writestr(f"xl/worksheets/sheet{i}.xml",
                       f'<worksheet xmlns="{NS}"><sheetData>{datos}</sheetData></worksheet>')
        z.writestr("xl/workbook.xml",
                   f'<workbook xmlns="{NS}" xmlns:r="{NSR}"><sheets>{"".join(sheets)}'
                   "</sheets></workbook>")
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                   f'2006/relationships">{"".join(rels)}</Relationships>')


# Encabezados como vienen en las planillas de abril 2026.
TITULOS = {"pago": "Pago", "central": "Central", "tipo": "Tipo_pago",
           "fecha": "Fecha_Hora", "bloque": "Bloque_15min", "clave": "Clave Año_Mes",
           "hora": "Hora Mensual", "periodo": "Periodo", "aux": "Aux",
           "config": "CONFIGURACION", "empresa": "Empresa"}


def bloque(cols, fila_enc, filas, titulos=TITULOS):
    """cols: {campo: letra}; filas: lista de dicts con los campos. Pone el
    encabezado de cada campo en fila_enc (titulos=None: sin encabezados)."""
    c = {}
    if titulos:
        for campo, letra in cols.items():
            c[(fila_enc, letra)] = titulos[campo]
    for i, r in enumerate(filas, start=fila_enc + 1):
        for campo, letra in cols.items():
            if campo in r:
                c[(i, letra)] = r[campo]
    return c


# 2026-09-01 18:00 -> serie Excel 46266.75 ; bloque (1-1)*96 + 18*4 + 0 + 1 = 73
SERIE_0109_1800 = 46266.75


class PagosVarios(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: [p.unlink() for p in self.tmp.iterdir()])
        # Planilla 3, hoja DB (abril): Clave Año_Mes | Tipo_pago | Central |
        # Hora Mensual | Pago, en BD:BH. Sin fecha: sale del bloque.
        p3 = bloque({"clave": "BD", "tipo": "BE", "central": "BF", "hora": "BG",
                     "pago": "BH"}, 2, [
            {"clave": 2609, "tipo": "CPF(+)", "central": "HE RALCO U2", "pago": 1794.478875,
             "hora": 37},
            {"clave": 2609, "tipo": "CPF(+)", "central": "HE RALCO U2", "pago": 0,
             "hora": 38},                                                       # 0: fuera
            {"clave": 2609, "tipo": "CSF(-)", "central": "", "pago": 5.0, "hora": 39},
            {"clave": 2609, "tipo": "CSF(-)", "central": "SAE DEL DESIERTO",
             "pago": 3112.342612166663, "hora": 2880},
        ])
        # Otra tabla en la misma hoja, lejos: no se tiene que tomar.
        p3.update(bloque({"tipo": "D", "central": "E", "hora": "F", "pago": "G"}, 2, [
            {"tipo": "OTRA", "central": "NO VA", "pago": 1.0, "hora": 1}]))
        p5 = bloque({"fecha": "AV", "central": "AX", "pago": "AY", "tipo": "AZ",
                     "bloque": "BA"}, 8, [
            {"fecha": SERIE_0109_1800 + 1e-9, "central": "RALCO_U1", "pago": 185.02,
             "tipo": "CRA", "bloque": 73},
        ])
        # CO_ERNC (abril): Fecha_Hora | CONFIGURACION | Periodo | Central | Pago |
        # Tipo_pago | Empresa, en BB:BH.
        co = bloque({"fecha": "BB", "config": "BC", "periodo": "BD", "central": "BE",
                     "pago": "BF", "tipo": "BG", "empresa": "BH"}, 8, [
            {"fecha": SERIE_0109_1800, "central": "PFV CEME 1", "pago": 9929.683898156996,
             "tipo": "CO_ERNC", "periodo": 73, "empresa": "GM_HOLDINGS"},
            {"fecha": SERIE_0109_1800, "central": "PFV CEME 2", "pago": 1.5,
             "tipo": "CO_ERNC", "periodo": 74},                                 # no calza
        ])
        # REA (abril), corrida 3 columnas a la derecha de AS:AX: se encuentra
        # igual por los encabezados.
        rea = bloque({"aux": "AV", "clave": "AW", "tipo": "AX", "central": "AY",
                      "hora": "AZ", "pago": "BA"}, 9, [
            {"aux": 0, "clave": SERIE_0109_1800, "central": "PFV EL ROMERO",
             "pago": 1000.0, "tipo": "REA", "hora": 74},
        ])
        self.rutas = {
            "p3": self.tmp / "3_REMUNERACION_SUBASTAS_E_ID_2609_R01D.xlsm",
            "p5": self.tmp / "5_REMUNERACION_CRA_2609_R01D.xlsx",
            "p6": self.tmp / "6_REMUNERACION_REA_Y_CO_ERNC_2609_R01D.xlsx",
        }
        escribir_xlsx(self.rutas["p3"], {"DB": p3})
        escribir_xlsx(self.rutas["p5"], {"CÁLCULO_CRA": p5})
        escribir_xlsx(self.rutas["p6"], {"ENERGIA_Y_CALCULO_CO_ERNC": co,
                                         "CALCULO_REA_CENTRAL": rea})

    def test_csv_completo(self):
        salida = self.tmp / "pagos_varios_2609.csv"
        log = []
        filas, resumen, avisos = pv.generar(self.rutas, "2609", salida, log.append)
        crudo = salida.read_bytes()
        self.assertFalse(crudo.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(crudo.count(b"\r\n"), 7)
        leido = list(csv.reader(crudo.decode("utf-8").splitlines()))
        self.assertEqual(leido[0], pv.ENCABEZADO_CSV)
        self.assertEqual(leido[1:], [
            ["01/09/2026 09:00", "", "HE RALCO U2", "1794.478875", "CPF(+)", "37"],
            ["30/09/2026 23:45", "", "SAE DEL DESIERTO", "3112.342612166663", "CSF(-)", "2880"],
            ["01/09/2026 18:00", "", "RALCO_U1", "185.02", "CRA", "73"],
            ["01/09/2026 18:00", "", "PFV CEME 1", "9929.683898156996", "CO_ERNC", "73"],
            ["01/09/2026 18:00", "", "PFV CEME 2", "1.5", "CO_ERNC", "74"],
            ["01/09/2026 18:15", "", "PFV EL ROMERO", "1000", "REA", "74"],
        ])
        self.assertEqual(resumen["CPF(+)"][0], 1)
        # Un bloque que no calza con su fecha: se avisa, no se toca.
        self.assertTrue(any("CO_ERNC" in a and "no calza" in a for a in avisos))
        self.assertFalse(any("no se encontraron" in a for a in avisos))
        self.assertTrue(any("calzaron 2 tablas" in l for l in log))
        rea = next(l for l in log if "encabezado en la fila 9" in l)
        self.assertIn("central AY", rea)
        self.assertIn("pago BA", rea)
        self.assertIn("bloque AZ", rea)
        self.assertIn("fecha desde el bloque", rea)

    def test_sin_encabezados_usa_las_columnas_fijas(self):
        rea = bloque({"tipo": "AU", "central": "AV", "hora": "AW", "pago": "AX"}, 9, [
            {"central": "PFV EL ROMERO", "pago": 7.0, "tipo": "REA", "hora": 1}],
            titulos=None)
        co = bloque({"fecha": "BB", "periodo": "BD", "central": "BE", "pago": "BF",
                     "tipo": "BG"}, 8, [
            {"fecha": SERIE_0109_1800, "central": "C", "pago": 2.0, "tipo": "CO_ERNC",
             "periodo": 73}])
        escribir_xlsx(self.rutas["p6"], {"ENERGIA_Y_CALCULO_CO_ERNC": co,
                                         "CALCULO_REA_CENTRAL": rea})
        filas, _, avisos = pv.generar(self.rutas, "2609", self.tmp / "x.csv",
                                      lambda *_: None)
        self.assertIn(["01/09/2026 00:00", "", "PFV EL ROMERO", 7.0, "REA", 1], filas)
        self.assertTrue(any("REA" in a and "no se encontraron" in a for a in avisos))

    def test_ubicar_encabezado(self):
        zona = {"B": {3: "Pago"}, "C": {3: "Central"}, "D": {3: "Tipo_pago"},
                "E": {3: "Hora Mensual"}, "F": {3: "Bloque_15min"}}
        f, cols, n = pv.ubicar_encabezado(zona)
        self.assertEqual((f, n), (3, 1))
        # Bloque_15min se prefiere a Hora Mensual; sin fecha no es error.
        self.assertEqual(cols, {"pago": "B", "central": "C", "tipo": "D", "bloque": "F"})
        # "Tipo_pago" no se confunde con "Pago"; sin Central no hay tabla.
        self.assertIsNone(pv.ubicar_encabezado({"A": {1: "Tipo_pago"}, "B": {1: "Pago"},
                                                "C": {1: "Periodo"}}))

    def test_falla_no_pisa_el_anterior(self):
        salida = self.tmp / "pagos_varios_2609.csv"
        salida.write_text("anterior", encoding="utf-8")
        rutas = dict(self.rutas, p5=self.tmp / "no_existe.xlsx")
        with self.assertRaises(Exception):
            pv.generar(rutas, "2609", salida, lambda *_: None)
        self.assertEqual(salida.read_text(encoding="utf-8"), "anterior")

    def test_funciones(self):
        self.assertEqual(pv.aamm_de_nombre("5_REMUNERACIÓN_CRA_2609_R01D.xlsx"), 2609)
        self.assertEqual(pv.detectar_aamm(self.rutas)[0], 2609)
        f = pv.fecha_de_bloque(2880, 2026, 9)
        self.assertEqual(f.strftime(pv.FORMATO_FECHA), "30/09/2026 23:45")
        self.assertEqual(pv.bloque_de_fecha(f), 2880)
        self.assertEqual(pv.fecha_de("2026-09-01 18:00:00").hour, 18)
        self.assertEqual(pv.fmt_pago(3.0), "3")
        self.assertEqual(pv.fmt_pago(10711.122449999999), "10711.122449999999")


if __name__ == "__main__":
    unittest.main()
