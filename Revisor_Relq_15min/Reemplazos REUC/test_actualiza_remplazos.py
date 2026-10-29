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


class DescargaNavegadorTest(unittest.TestCase):
    """La descarga ya no automatiza el navegador: se mira la carpeta Descargas."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        raiz = Path(self.tmp.name)
        self.bajadas = raiz / "Downloads"
        self.aux = raiz / "Auxiliares REUC"
        self.bajadas.mkdir()
        self.aux.mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def _bajar(self, nombre, texto="xlsx"):
        f = self.bajadas / nombre
        f.write_text(texto)
        return f

    def test_clase_descarga(self):
        self.assertEqual(ar.clase_descarga("datos_reuc_20261029.xlsx"), "reuc")
        self.assertEqual(ar.clase_descarga("datos_reuc (1).xlsx"), "reuc")
        self.assertEqual(ar.clase_descarga("Datos_REUC_reemplazos_2026.xlsx"), "reemplazos")
        self.assertIsNone(ar.clase_descarga("datos_reuc.xlsx.crdownload"))
        self.assertIsNone(ar.clase_descarga("~$datos_reuc.xlsx"))
        self.assertIsNone(ar.clase_descarga("otra_cosa.xlsx"))

    def test_solo_las_nuevas(self):
        import os
        import time
        viejo = self._bajar("datos_reuc_viejo.xlsx")
        os.utime(viejo, (time.time() - 3600,) * 2)
        self._bajar("datos_reuc_reemplazos_1.xlsx")
        self._bajar("Unconfirmed 123.crdownload")
        self._bajar("datos_reuc_vacio.xlsx", "")
        nuevas = ar.descargas_nuevas(self.bajadas, time.time() - 60)
        self.assertEqual({c: f.name for c, f in nuevas.items()},
                         {"reemplazos": "datos_reuc_reemplazos_1.xlsx"})

    def test_espera_mueve_y_borra_viejos(self):
        (self.aux / "datos_reuc_20260101.xlsx").write_text("viejo")
        (self.aux / "datos_reuc_reemplazos_20260101.xlsx").write_text("viejo")
        self._bajar("datos_reuc_20261029.xlsx")
        self._bajar("datos_reuc_reemplazos (2).xlsx")
        lineas = []
        res = ar.esperar_descargas(self.bajadas, self.aux, 0, 60, log=lineas.append,
                                   dormir=lambda _: None)
        self.assertEqual(res["reuc"].name, "datos_reuc_20261029.xlsx")
        self.assertTrue(res["reemplazos"].name.startswith("datos_reuc_reemplazos_"))
        self.assertNotIn(" ", res["reemplazos"].name)
        self.assertEqual(sorted(f.name for f in self.aux.iterdir()),
                         sorted([res["reuc"].name, res["reemplazos"].name]))
        self.assertEqual(list(self.bajadas.iterdir()), [])

    def test_sin_descarga_avisa_cual_falta(self):
        self._bajar("datos_reuc_20261029.xlsx")
        t = iter(range(0, 1000, 10))
        with self.assertRaises(TimeoutError) as e:
            ar.esperar_descargas(self.bajadas, self.aux, 0, 60, log=lambda *_: None,
                                 reloj=lambda: next(t), dormir=lambda _: None)
        self.assertIn("datos_reuc_reemplazos", str(e.exception))
        self.assertTrue((self.aux / "datos_reuc_20261029.xlsx").exists())

    def test_descargar_abre_los_dos_exports(self):
        abiertos = []

        def abrir(url):
            abiertos.append(url)
            nombre = ("datos_reuc_reemplazos_x.xlsx" if "reemplazadas" in url
                      else "datos_reuc_x.xlsx")
            self._bajar(nombre)

        res = ar.descargar_reuc(self.aux, log=lambda *_: None,
                                carpeta_bajadas=self.bajadas, abrir=abrir)
        self.assertEqual(abiertos, [ar.URL_REUC_EMPRESAS, ar.URL_REUC_REEMPLAZADAS])
        self.assertEqual(set(res), {"reuc", "reemplazos"})


if __name__ == "__main__":
    unittest.main()
