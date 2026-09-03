"""Pruebas del armado compartido de rutas de 00_Salidas."""

from pathlib import Path
import tempfile
import unittest

from salidas import (
    carpeta_anio,
    carpeta_comparador,
    carpeta_mes,
    carpetas_legado,
    nombre_carpeta_mes,
    partir_aamm,
    raiz_salidas,
)


class SalidasTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.raiz = Path(self.tmp.name) / "00_Salidas"

    def tearDown(self):
        self.tmp.cleanup()

    def test_partir_aamm(self):
        self.assertEqual(partir_aamm("2407"), ("2024", 7))
        self.assertEqual(partir_aamm("2312"), ("2023", 12))
        for invalido in ("sin_mes", "", "24", "240713", "2400", "2413"):
            with self.subTest(invalido=invalido):
                self.assertIsNone(partir_aamm(invalido))

    def test_nombre_carpeta_mes(self):
        self.assertEqual(nombre_carpeta_mes("2401"), "01 Enero")
        self.assertEqual(nombre_carpeta_mes("2409"), "09 Septiembre")

    def test_rutas_auxiliares(self):
        script = Path(self.tmp.name) / "Revisor_Relq"
        self.assertEqual(raiz_salidas(script), self.raiz)
        self.assertEqual(carpeta_anio(self.raiz, "2407"), self.raiz / "2024")
        self.assertEqual(
            carpeta_comparador(self.raiz, "2024", "_comparador"),
            self.raiz / "2024" / "_comparador",
        )

    def test_carpeta_mes_canonica_sin_disco(self):
        self.assertEqual(
            carpeta_mes(self.raiz, "2407"), self.raiz / "2024" / "07 Julio"
        )

    def test_carpeta_mes_encuentra_variantes(self):
        for nombre in ("7 Julio", "07 julio"):
            with self.subTest(nombre=nombre):
                variante = self.raiz / "2024" / nombre
                variante.mkdir(parents=True)
                self.assertEqual(carpeta_mes(self.raiz, "2407"), variante)
                variante.rmdir()

    def test_sin_mes_conserva_ruta_plana(self):
        self.assertEqual(carpeta_mes(self.raiz, "sin_mes"), self.raiz / "sin_mes")

    def test_crear_usa_siempre_la_canonica(self):
        variante = self.raiz / "2024" / "7 Julio"
        variante.mkdir(parents=True)
        canonica = self.raiz / "2024" / "07 Julio"
        self.assertEqual(carpeta_mes(self.raiz, "2407", crear=True), canonica)
        self.assertTrue(canonica.is_dir())

    def test_no_reutiliza_carpeta_legado(self):
        legado = self.raiz / "2407"
        legado.mkdir(parents=True)
        self.assertEqual(
            carpeta_mes(self.raiz, "2407"), self.raiz / "2024" / "07 Julio"
        )

    def test_carpetas_legado(self):
        legado = self.raiz / "2407"
        legado.mkdir(parents=True)
        (self.raiz / "2024").mkdir()
        self.assertEqual(carpetas_legado(self.raiz), [legado])


if __name__ == "__main__":
    unittest.main()
