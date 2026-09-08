"""Pruebas de ActualizaRemplazos que no necesitan Excel ni navegador.

Se carga el .py a mano y con pandas/xlwings/tkinter simulados: en un
contenedor sin esas librerias (y sin Excel) el modulo no se puede importar de
la forma normal, pero lo que se prueba aca es solo armado de rutas y texto.
"""

import importlib.util
from pathlib import Path
import sys
import tempfile
import types
import unittest


def _cargar_modulo():
    for nombre in ("tkinter", "tkinter.filedialog", "tkinter.messagebox",
                   "tkinter.ttk", "pandas", "xlwings"):
        if nombre not in sys.modules:
            mod = types.ModuleType(nombre)
            mod.__getattr__ = lambda n: types.SimpleNamespace()
            sys.modules[nombre] = mod
    for hijo in ("filedialog", "messagebox", "ttk"):
        setattr(sys.modules["tkinter"], hijo, sys.modules[f"tkinter.{hijo}"])
    ruta = Path(__file__).with_name("ActualizaRemplazos.py")
    spec = importlib.util.spec_from_file_location("_actualiza_remplazos", ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


ar = _cargar_modulo()


class AuxiliaresTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.raiz = Path(self.tmp.name)
        self.nueva = self.raiz / "__config__" / "Auxiliares REUC"
        self.vieja = self.raiz / "Reemplazos REUC" / "Auxiliares"
        self._original = (ar.CARPETA_AUXILIARES, ar.CARPETA_AUXILIARES_LEGADO)
        ar.CARPETA_AUXILIARES = self.nueva
        ar.CARPETA_AUXILIARES_LEGADO = self.vieja

    def tearDown(self):
        ar.CARPETA_AUXILIARES, ar.CARPETA_AUXILIARES_LEGADO = self._original
        self.tmp.cleanup()

    def _poblar_vieja(self):
        self.vieja.mkdir(parents=True)
        (self.vieja / "datos_reuc_20260101.xlsx").write_text("x")
        (self.vieja / "Reemplazos forzados 2026.xlsx").write_text("x")
        (self.vieja / "~$datos_reuc_abierto.xlsx").write_text("x")
        (self.vieja / "notas.txt").write_text("x")

    def test_vive_en_config(self):
        """La de verdad, no la del tmp: cuelga de __config__, no del .py."""
        raiz = Path(ar.__file__).resolve().parents[2]
        self.assertEqual(self._original[0], raiz / "__config__" / "Auxiliares REUC")
        self.assertEqual(self._original[1],
                         Path(ar.__file__).parent / "Auxiliares")

    def test_carpeta_auxiliares_la_crea(self):
        self.assertFalse(self.nueva.exists())
        self.assertEqual(ar.carpeta_auxiliares(), self.nueva)
        self.assertTrue(self.nueva.is_dir())

    def test_sin_carpeta_vieja_no_avisa(self):
        lineas = []
        self.assertFalse(ar.avisar_legado(lineas.append))
        self.assertEqual(lineas, [])

    def test_avisa_lo_que_quedo_y_no_mueve_nada(self):
        self._poblar_vieja()
        self.assertEqual([f.name for f in ar.pendientes_en_legado()],
                         ["Reemplazos forzados 2026.xlsx", "datos_reuc_20260101.xlsx"])
        lineas = []
        self.assertTrue(ar.avisar_legado(lineas.append))
        self.assertTrue(any("carpeta vieja" in l for l in lineas))
        self.assertTrue((self.vieja / "datos_reuc_20260101.xlsx").exists())
        self.assertFalse(self.nueva.exists())

    def test_config_viejo_se_corrige_solo(self):
        self.assertEqual(ar.carpeta_datos_guardada({}), str(self.nueva))
        self.assertEqual(ar.carpeta_datos_guardada({"carpeta_datos": "  "}),
                         str(self.nueva))
        self.assertEqual(ar.carpeta_datos_guardada({"carpeta_datos": str(self.vieja)}),
                         str(self.nueva))
        self.assertEqual(ar.carpeta_datos_guardada({"carpeta_datos": "T:/otra"}),
                         "T:/otra")

    def test_respaldo_cae_en_la_nueva(self):
        ar.carpeta_auxiliares()
        (self.nueva / "datos_reuc_20260202.xlsx").write_text("x")
        elegida = self.raiz / "elegida"
        elegida.mkdir()
        lineas = []
        hallado = ar.buscar_archivo_con_respaldo(elegida, "datos_reuc_*.xlsx",
                                                 log=lineas.append)
        self.assertEqual(hallado.name, "datos_reuc_20260202.xlsx")
        self.assertTrue(any("Auxiliares REUC" in l for l in lineas))

    def test_si_no_esta_en_ninguna_avisa_del_legado(self):
        self._poblar_vieja()
        elegida = self.raiz / "elegida"
        elegida.mkdir()
        lineas = []
        with self.assertRaises(FileNotFoundError) as caso:
            ar.buscar_archivo_con_respaldo(elegida, "no_existe_*.xlsx",
                                           log=lineas.append)
        self.assertIn("Auxiliares REUC", str(caso.exception))
        self.assertTrue(any("carpeta vieja" in l for l in lineas))


class _CampoFalso:
    """Un <input> de mentira, con lo justo que usa el codigo de login."""

    def __init__(self, valor="", visible=True):
        self.valor = valor
        self.visible = visible
        self.teclas = []

    # --- API de Playwright que se usa ---
    def is_visible(self):
        return self.visible

    def input_value(self):
        return self.valor

    def fill(self, texto):
        self.valor = texto

    def press(self, tecla):
        self.teclas.append(tecla)


class _LocatorFalso:
    def __init__(self, campos):
        self.campos = campos

    def count(self):
        return len(self.campos)

    @property
    def first(self):
        return self.campos[0]


class _PaginaFalsa:
    """Devuelve campos segun el selector, como haria pagina.locator()."""

    def __init__(self, por_selector):
        self.por_selector = por_selector

    def locator(self, selector):
        return _LocatorFalso(self.por_selector.get(selector, []))


class CorreoTest(unittest.TestCase):
    def test_correo_valido(self):
        for bueno in ("a@b.cl", "nombre.apellido@empresa.co.uk"):
            self.assertTrue(ar.correo_valido(bueno), bueno)
        for malo in ("", None, "   ", "sin arroba.cl", "a@b", "a b@c.cl", "@b.cl"):
            self.assertFalse(ar.correo_valido(malo), malo)

    def test_escribe_el_correo_y_salta_a_la_clave(self):
        campo = _CampoFalso()
        pagina = _PaginaFalsa({'input[type="email"]': [campo]})
        self.assertEqual(ar._rellenar_correo(pagina, "yo@empresa.cl"), "escrito")
        self.assertEqual(campo.valor, "yo@empresa.cl")
        self.assertEqual(campo.teclas, ["Tab"])

    def test_no_pisa_lo_que_la_persona_escribio(self):
        campo = _CampoFalso(valor="otra@empresa.cl")
        pagina = _PaginaFalsa({'input[type="email"]': [campo]})
        self.assertEqual(ar._rellenar_correo(pagina, "yo@empresa.cl"), "ya_tenia")
        self.assertEqual(campo.valor, "otra@empresa.cl")
        self.assertEqual(campo.teclas, [])

    def test_sin_campo_todavia_no_escribe_nada(self):
        pagina = _PaginaFalsa({})
        self.assertEqual(ar._rellenar_correo(pagina, "yo@empresa.cl"), "")
        self.assertEqual(ar._correo_en_pantalla(pagina), "")

    def test_sin_correo_recordado_no_toca_la_pagina(self):
        campo = _CampoFalso()
        pagina = _PaginaFalsa({'input[type="email"]': [campo]})
        self.assertEqual(ar._rellenar_correo(pagina, ""), "")
        self.assertEqual(campo.valor, "")

    def test_campo_invisible_se_ignora(self):
        escondido = _CampoFalso(visible=False)
        visible = _CampoFalso()
        pagina = _PaginaFalsa({
            'input[type="email"]': [escondido],
            'input[name*="usuario" i]': [visible],
        })
        self.assertEqual(ar._rellenar_correo(pagina, "yo@empresa.cl"), "escrito")
        self.assertEqual(escondido.valor, "")
        self.assertEqual(visible.valor, "yo@empresa.cl")

    def test_lee_el_correo_que_escribio_la_persona(self):
        campo = _CampoFalso(valor="  yo@empresa.cl ")
        pagina = _PaginaFalsa({'input[type="email"]': [campo]})
        self.assertEqual(ar._correo_en_pantalla(pagina), "yo@empresa.cl")

    def test_no_recuerda_algo_a_medio_escribir(self):
        campo = _CampoFalso(valor="yo@")
        pagina = _PaginaFalsa({'input[type="email"]': [campo]})
        self.assertEqual(ar._correo_en_pantalla(pagina), "")

    def test_un_campo_que_revienta_no_rompe_nada(self):
        class Explosivo(_CampoFalso):
            def is_visible(self):
                raise RuntimeError("el selector ya no existe")

        campo = _CampoFalso()
        pagina = _PaginaFalsa({
            'input[type="email"]': [Explosivo()],
            'input[name*="email" i]': [campo],
        })
        self.assertEqual(ar._rellenar_correo(pagina, "yo@empresa.cl"), "escrito")
        self.assertEqual(campo.valor, "yo@empresa.cl")


if __name__ == "__main__":
    unittest.main()
