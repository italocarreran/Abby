"""Pruebas stdlib del lector OOXML compartido."""

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __comun__ import excel_xml


WORKBOOK = f'''<workbook xmlns="{excel_xml.NS_XL[1:-1]}"
 xmlns:r="{excel_xml.NS_REL[1:-1]}"><sheets>
 <sheet name="Generación Ñ" sheetId="1" r:id="rId1"/></sheets></workbook>'''
RELS = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'''
SHEET = f'''<worksheet xmlns="{excel_xml.NS_XL[1:-1]}"><sheetData><row r="2">
 <c r="A2" t="s"><v>0</v></c><c r="B2"><v>12.5</v></c>
 <c r="C2" t="inlineStr"><is><t>Pehuenche &#209;</t></is></c>
 <c r="D2" t="b"><v>1</v></c></row></sheetData></worksheet>'''
SHARED = f'''<sst xmlns="{excel_xml.NS_XL[1:-1]}"><si><t>Enel Generaci&#243;n</t></si></sst>'''


class TestExcelXml(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ruta = Path(self.tmp.name) / "libro.xlsm"
        with zipfile.ZipFile(self.ruta, "w") as z:
            z.writestr("xl/workbook.xml", WORKBOOK)
            z.writestr("xl/_rels/workbook.xml.rels", RELS)
            z.writestr("xl/worksheets/sheet1.xml", SHEET)
            z.writestr("xl/sharedStrings.xml", SHARED)

    def tearDown(self):
        self.tmp.cleanup()

    def test_extensiones_ooxml(self):
        self.assertTrue(excel_xml.es_zip_excel("x.xlsm"))
        self.assertFalse(excel_xml.es_zip_excel("x.xlsb"))

    def test_expande_en_ambos_sentidos(self):
        self.assertEqual(excel_xml.expandir_columnas("$Z : AB"), ["Z", "AA", "AB"])
        self.assertEqual(excel_xml.expandir_columnas("C:A"), ["A", "B", "C"])

    def test_ubica_por_nombre_normalizado_y_posicion(self):
        with zipfile.ZipFile(self.ruta) as z:
            esperado = "xl/worksheets/sheet1.xml"
            self.assertEqual(excel_xml.ubicar_hoja_xml(z, "generacion n")[0], esperado)
            self.assertEqual(excel_xml.ubicar_hoja_xml(z, "#1")[0], esperado)

    def test_lee_tipos_y_entidades(self):
        mensajes = []
        datos = excel_xml.leer_columnas_rapido(
            self.ruta, "Generación Ñ", ["A", "B", "C", "D"], 2, mensajes.append
        )
        self.assertEqual(datos, {
            "A": {2: "Enel Generación"}, "B": {2: 12.5},
            "C": {2: "Pehuenche Ñ"}, "D": {2: True},
        })
        self.assertEqual(mensajes, [])

    def test_fallos_devuelven_none_y_avisan(self):
        mensajes = []
        self.assertIsNone(excel_xml.leer_columnas_rapido("x.xlsb", "H", ["A"], 1, mensajes.append))
        self.assertIsNone(excel_xml.leer_columnas_rapido(self.ruta, "No existe", ["A"], 1, mensajes.append))
        self.assertEqual(len(mensajes), 2)

    def test_desescape_en_una_pasada(self):
        self.assertEqual(excel_xml.desescapar_xml("&#243; &#xD1; &amp;"), "ó Ñ &")
        self.assertEqual(excel_xml.desescapar_xml("&amp;lt;"), "&lt;")


if __name__ == "__main__":
    unittest.main()
