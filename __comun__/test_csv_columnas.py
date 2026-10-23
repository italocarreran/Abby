"""Pruebas stdlib del lector de .csv por letra de columna."""

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __comun__ import csv_columnas as cc

# Las dos primeras filas reales del Reporte_CRA_15min_2512.csv (Balance_SSCC).
REPORTE = (
    "Clave año_mes;Tipo;FECHA;Hora;Bloque_15min;Hora_Mensual_15min;FECHA_HORA;"
    "UNIDAD GENERADORA;RCO;Prorrata_CPF(+);Prorrata_CPF(-);Prorrata_CSF(+);"
    "Prorrata_CSF(-);Prorrata_CTF(+);Prorrata_CTF(-);Prorrata_CPF;Prorrata_CSF;"
    "Prorrata_CTF\r\n"
    "2512;CO;01-12-2025;10;38;38;01-12-2025 9:15;CANUTILLAR-1;3202.189645;"
    "0.912266283;0.087733717;0;0;0;0;1;0;0\r\n"
    "2512;CO;01-12-2025;10;39;39;01-12-2025 9:30;CANUTILLAR-1;3382.51529;"
    "0.912266283;0.087733717;0;0;0;0;1;0;0\r\n"
)


class TestCsvColumnas(unittest.TestCase):
    def escribir(self, texto, cod="utf-8"):
        tmp = tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False,
                                          encoding=cod, newline="")
        tmp.write(texto)
        tmp.close()
        self.addCleanup(Path(tmp.name).unlink)
        return tmp.name

    def test_reporte_cra_real(self):
        d = cc.leer_columnas_csv(self.escribir(REPORTE), ["A", "C", "E", "H", "I", "J"], 2)
        self.assertEqual(d["A"], {2: 2512, 3: 2512})
        self.assertEqual(d["C"][2], "01-12-2025")
        self.assertEqual(d["E"], {2: 38, 3: 39})
        self.assertEqual(d["H"][3], "CANUTILLAR-1")
        self.assertAlmostEqual(d["I"][2], 3202.189645)
        self.assertAlmostEqual(d["J"][2], 0.912266283)
        self.assertNotIn(1, d["A"])   # el encabezado queda fuera con fila 2

    def test_todas_las_columnas_y_vacias(self):
        d = cc.leer_columnas_csv(self.escribir("a;b;;d\n1;;x;\n"))
        self.assertEqual(d["A"], {1: "a", 2: 1})
        self.assertEqual(d["C"], {2: "x"})
        self.assertNotIn(2, d["B"])

    def test_coma_decimal_y_cp1252(self):
        d = cc.leer_columnas_csv(self.escribir("Año;Monto\n2512;3,5\n", "cp1252"))
        self.assertEqual(d["A"][1], "Año")
        self.assertEqual(d["B"][2], 3.5)

    def test_separador_coma(self):
        d = cc.leer_columnas_csv(self.escribir("x,y\n1.5,2\n"))
        self.assertEqual(d["A"][2], 1.5)
        self.assertEqual(d["B"][2], 2)

    def test_archivo_que_no_existe(self):
        self.assertIsNone(cc.leer_columnas_csv("/no/existe.csv", ["A"]))

    def test_letras(self):
        self.assertEqual([cc.letra_columna(n) for n in (1, 26, 27, 52)],
                         ["A", "Z", "AA", "AZ"])


if __name__ == "__main__":
    unittest.main()
