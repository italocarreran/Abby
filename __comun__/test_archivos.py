"""Pruebas stdlib de filtros de archivos compartidos."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __comun__ import archivos


class Entrada:
    def __init__(self, nombre, mtime):
        self.nombre = nombre
        self.mtime = mtime


class TestArchivos(unittest.TestCase):
    def test_temporales(self):
        self.assertTrue(archivos.es_temporal("~$libro.xlsx"))
        self.assertTrue(archivos.es_temporal(".oculto"))
        self.assertFalse(archivos.es_temporal("libro.xlsx"))

    def test_variantes_de_copia(self):
        for nombre in ("a - copia.xlsx", "a - copia (2).xlsm", "a - Copy.mdb", "a (20).xlsx"):
            self.assertTrue(archivos.es_copia(nombre), nombre)
        self.assertFalse(archivos.es_copia("a original.xlsx"))

    def test_original_gana_a_copia_mas_nueva(self):
        entradas = [Entrada("a.xlsx", 1), Entrada("a - copia.xlsx", 2)]
        filtradas = archivos.preferir_originales(entradas, lambda e: e.nombre)
        self.assertEqual([e.nombre for e in filtradas], ["a.xlsx"])

    def test_si_solo_hay_copias_conserva_la_via_de_escape(self):
        entradas = [Entrada("a - copia.xlsx", 1)]
        self.assertEqual(archivos.preferir_originales(entradas, lambda e: e.nombre), entradas)

    def test_mas_reciente_usa_mtime_entregado_sin_stat(self):
        entradas = [Entrada("a", 1), Entrada("b", 3)]
        elegida = archivos.mas_reciente(entradas, lambda e: e.mtime)
        self.assertIsNotNone(elegida)
        assert elegida is not None
        self.assertEqual(elegida.nombre, "b")
        self.assertIsNone(archivos.mas_reciente([], lambda e: e.mtime))


if __name__ == "__main__":
    unittest.main()
