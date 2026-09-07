"""Pruebas stdlib de las normalizaciones compartidas."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __comun__ import texto


class TestTexto(unittest.TestCase):
    def test_suave(self):
        self.assertEqual(texto.suave("  Generación   Ñuble "), "generacion nuble")
        self.assertEqual(texto.suave(None), "")
        self.assertEqual(texto.suave(0), "0")

    def test_variantes_historicas_de_none(self):
        self.assertEqual(texto.suave_textual(None), "none")
        with self.assertRaises(TypeError):
            texto.suave_requerido(None)

    def test_clave_elimina_espacios_y_guiones_bajos(self):
        esperada = "eltoro-1"
        self.assertEqual(texto.clave("El Toro-1"), esperada)
        self.assertEqual(texto.clave("EL_TORO-1"), esperada)
        self.assertEqual(texto.clave("ELTORO-1"), esperada)

    def test_clave_conserva_signos_del_concepto(self):
        self.assertNotEqual(texto.clave("CSF(+)"), texto.clave("CSF(-)"))

    def test_mayusculas(self):
        self.assertEqual(texto.clave_mayusculas("Pehuenche_1"), "PEHUENCHE1")
        self.assertEqual(texto.clave_mayusculas(0), "")

    def test_anio_y_ano_son_la_misma_columna(self):
        valores = ["Clave Año_Mes", "Clave_Anio_Mes", "clave ano mes"]
        self.assertEqual({texto.clave_columna(v) for v in valores}, {"CLAVEANOMES"})


if __name__ == "__main__":
    unittest.main()
