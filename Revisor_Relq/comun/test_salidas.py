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
    normalizar_anio,
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

    # Antes esta prueba exigia que crear=True armara SIEMPRE la canonica, aun
    # habiendo una variante al lado. Se cambio a proposito: partia el mes en
    # dos carpetas. Ver test_crear_respeta_una_variante_ya_existente.

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


    # --- casos agregados al revisar la Tarea 1 -----------------------------

    def test_crear_respeta_una_variante_ya_existente(self):
        """El mes no se puede partir en dos carpetas.

        Si el usuario ya escribio "7 Julio" a mano, escribir NO debe crear
        "07 Julio" al lado: el estado se leeria de una y se escribiria en la
        otra, y nadie se enteraria.
        """
        variante = self.raiz / "2024" / "7 Julio"
        variante.mkdir(parents=True)
        leer = carpeta_mes(self.raiz, "2407")
        escribir = carpeta_mes(self.raiz, "2407", crear=True)
        self.assertEqual(leer, variante)
        self.assertEqual(escribir, variante)
        self.assertEqual(sorted(p.name for p in (self.raiz / "2024").iterdir()),
                         ["7 Julio"])

    def test_crear_hace_la_canonica_si_no_hay_variante(self):
        ruta = carpeta_mes(self.raiz, "2407", crear=True)
        self.assertTrue(ruta.is_dir())
        self.assertEqual(ruta.name, "07 Julio")

    def test_normalizar_anio(self):
        self.assertEqual(normalizar_anio("2024"), "2024")
        self.assertEqual(normalizar_anio("24"), "2024")
        self.assertEqual(normalizar_anio(" 25 "), "2025")
        for malo in ("", None, "204", "20244", "ab", "sin_mes"):
            self.assertIsNone(normalizar_anio(malo), malo)

    def test_carpeta_comparador_acepta_dos_o_cuatro_digitos(self):
        """La ventana acepta el anio de las dos formas; la carpeta es una sola."""
        cuatro = carpeta_comparador(self.raiz, "2025", "_comparador")
        dos = carpeta_comparador(self.raiz, "25", "_comparador")
        self.assertEqual(cuatro, dos)
        self.assertEqual(cuatro, self.raiz / "2025" / "_comparador")

    def test_carpeta_comparador_falla_fuerte_con_anio_raro(self):
        with self.assertRaises(ValueError):
            carpeta_comparador(self.raiz, "pepe", "_comparador")


if __name__ == "__main__":
    unittest.main()
