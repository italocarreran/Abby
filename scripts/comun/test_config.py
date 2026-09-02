# -*- coding: utf-8 -*-
"""Pruebas de comun/config.py. Solo stdlib, se corre con: python test_config.py"""
import json, os, sys, tempfile, unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from comun import config


class TestConfig(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.ruta = Path(self.dir.name) / "config.json"
        self.yo = config.clave_equipo()

    def tearDown(self):
        self.dir.cleanup()

    # --- lectura -----------------------------------------------------------
    def test_sin_archivo_devuelve_vacio(self):
        self.assertEqual(config.leer(self.ruta), {})
        self.assertEqual(config.leer_todo(self.ruta), {})

    def test_lee_solo_el_bloque_propio(self):
        self.ruta.write_text(json.dumps({
            self.yo: {"carpeta_base": "T:/mio"},
            "otroPC_otro": {"carpeta_base": "T:/ajeno"},
        }), encoding="utf-8")
        self.assertEqual(config.leer(self.ruta), {"carpeta_base": "T:/mio"})

    def test_archivo_roto_no_revienta(self):
        self.ruta.write_text("{ esto no es json", encoding="utf-8")
        self.assertEqual(config.leer(self.ruta), {})

    def test_json_que_no_es_dict(self):
        self.ruta.write_text("[1, 2, 3]", encoding="utf-8")
        self.assertEqual(config.leer(self.ruta), {})

    # --- la regla que importa ---------------------------------------------
    def test_no_pisa_un_archivo_roto(self):
        """Un config.json ilegible NO se sobreescribe: se pierde el ajuste,
        no el archivo. Es el bug que tenia ActualizaRemplazos."""
        original = "{ roto a proposito"
        self.ruta.write_text(original, encoding="utf-8")
        self.assertFalse(config.guardar(self.ruta, {"x": 1}))
        self.assertEqual(self.ruta.read_text(encoding="utf-8"), original)

    def test_no_borra_lo_de_otros_equipos(self):
        self.ruta.write_text(json.dumps({
            "otroPC_otro": {"carpeta_base": "T:/ajeno", "mdb": "x.mdb"},
        }), encoding="utf-8")
        self.assertTrue(config.guardar(self.ruta, {"carpeta_base": "T:/mio"}))
        todo = json.loads(self.ruta.read_text(encoding="utf-8"))
        self.assertEqual(todo["otroPC_otro"], {"carpeta_base": "T:/ajeno", "mdb": "x.mdb"})
        self.assertEqual(todo[self.yo], {"carpeta_base": "T:/mio"})

    def test_no_borra_las_claves_propias_que_no_se_tocan(self):
        config.guardar(self.ruta, {"carpeta_base": "T:/a", "ultimo_mes": "2407"})
        config.guardar(self.ruta, {"carpeta_base": "T:/b"})
        self.assertEqual(config.leer(self.ruta),
                         {"carpeta_base": "T:/b", "ultimo_mes": "2407"})

    # --- mutador (tasa por mes, que es el uso real de Cuadro0) -------------
    def test_mutador_anida_por_mes(self):
        def mut(todo):
            mio = todo.setdefault(config.clave_equipo(), {})
            porm = mio.setdefault("tasa_interes_por_mes", {})
            porm["2407"] = {"valor": 0.168, "fecha": "02-09-2026 10:00"}
        self.assertTrue(config.modificar(self.ruta, mut))
        self.assertEqual(config.leer(self.ruta)["tasa_interes_por_mes"]["2407"]["valor"],
                         0.168)

    def test_mutador_que_falla_no_escribe(self):
        config.guardar(self.ruta, {"carpeta_base": "T:/a"})
        antes = self.ruta.read_text(encoding="utf-8")
        def mut(todo):
            raise RuntimeError("algo se rompio")
        self.assertFalse(config.modificar(self.ruta, mut))
        self.assertEqual(self.ruta.read_text(encoding="utf-8"), antes)

    # --- escritura ---------------------------------------------------------
    def test_escritura_es_atomica_no_deja_tmp(self):
        config.guardar(self.ruta, {"x": 1})
        sobrantes = [p.name for p in self.ruta.parent.iterdir() if ".tmp" in p.name]
        self.assertEqual(sobrantes, [])

    def test_crea_la_carpeta_si_no_existe(self):
        """El de Reemplazos REUC vive en Auxiliares/, que puede no estar."""
        ruta = Path(self.dir.name) / "Auxiliares" / "config.json"
        self.assertTrue(config.guardar(ruta, {"raiz": "T:/Facturacion"}))
        self.assertEqual(config.leer(ruta), {"raiz": "T:/Facturacion"})

    def test_conserva_acentos_y_enes(self):
        config.guardar(self.ruta, {"carpeta": "T:/Cálculo Año_Mes/Ñuñoa"})
        crudo = self.ruta.read_text(encoding="utf-8")
        self.assertIn("Cálculo Año_Mes/Ñuñoa", crudo)
        self.assertEqual(config.leer(self.ruta)["carpeta"], "T:/Cálculo Año_Mes/Ñuñoa")

    def test_clave_equipo_tiene_la_forma_esperada(self):
        self.assertIn("_", config.clave_equipo())
        self.assertEqual(config.clave_equipo(), config.clave_equipo())


if __name__ == "__main__":
    unittest.main(verbosity=2)
