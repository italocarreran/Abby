# -*- coding: utf-8 -*-
"""Pruebas stdlib de los adaptadores de lectura aislados de la ventana."""

from pathlib import Path
import sys
import unittest

RAIZ = Path(__file__).resolve().parents[2]
DIR_REVISOR = Path(__file__).resolve().parents[1]
for carpeta in (RAIZ, DIR_REVISOR):
    if str(carpeta) not in sys.path:
        sys.path.insert(0, str(carpeta))

from revisor import lectores


class LectoresTest(unittest.TestCase):
    def test_columnas_ida_y_vuelta(self):
        for numero, letra in ((1, "A"), (26, "Z"), (27, "AA"), (703, "AAA")):
            self.assertEqual(lectores.col_letra(numero), letra)
            self.assertEqual(lectores.col_letra_a_num(letra), numero)

    def test_resolver_hoja_por_nombre_normalizado_y_posicion(self):
        nombres = ["Resumen", "CÁLCULO_CO"]
        self.assertEqual(lectores.resolver_hoja(nombres, "#2"), "CÁLCULO_CO")
        self.assertEqual(lectores.resolver_hoja(nombres, "calculo_co"), "CÁLCULO_CO")
        self.assertIsNone(lectores.resolver_hoja(nombres, "#3"))

    def test_armar_tabla_conserva_contrato_historico(self):
        mensajes = []
        info = {}
        datos = {
            "A": {5: " Empresa 1 ", 6: "", 7: 0, 8: "Empresa 1"},
            "B": {5: 10, 8: 20},
        }
        tabla = lectores.armar_tabla(datos, "A", ["B"], mensajes.append,
                                      etiqueta="origen", info=info)
        self.assertEqual(tabla, {"empresa 1": ("Empresa 1", [20.0])})
        self.assertEqual(info["vacias"], 1)
        self.assertEqual(info["excluidas"], [(7, 0)])
        self.assertEqual(info["duplicadas"], [("Empresa 1", 8)])
        self.assertTrue(any("EMPRESAS REPETIDAS" in m for m in mensajes))

    def test_es_significativo(self):
        for valor in (None, "", "  ", "#REF!", 0, 0.0):
            self.assertFalse(lectores.es_significativo(valor))
        for valor in ("empresa", 1, -1, True):
            self.assertTrue(lectores.es_significativo(valor))


if __name__ == "__main__":
    unittest.main()
