"""Pruebas de revisor/embalses.py (solo biblioteca estandar)."""

import json
import sys
import tempfile
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(RAIZ / "Revisor_Relq_15min"))
from revisor import embalses  # noqa: E402


class Embalses(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.ruta = Path(self.tmp.name) / "__config__" / "centrales_embalse_15min.json"

    def test_vive_en_config(self):
        self.assertEqual(embalses.RUTA, RAIZ / "__config__" / "centrales_embalse_15min.json")

    def test_sin_archivo(self):
        self.assertEqual(embalses.leer(self.ruta), ([], []))
        base = ["RALCO-1"]
        self.assertEqual(embalses.ampliar(base, self.ruta), [])
        self.assertEqual(base, ["RALCO-1"])

    def test_agregar_ignorar_y_cambiar_de_idea(self):
        self.assertTrue(embalses.guardar(["RALCO-3"], ["PE-CANELA-2"], self.ruta))
        self.assertEqual(embalses.leer(self.ruta), (["RALCO-3"], ["PE-CANELA-2"]))
        base = ["RALCO-1", "RALCO-2"]
        self.assertEqual(embalses.ampliar(base, self.ruta), ["RALCO-3"])
        self.assertEqual(base, ["RALCO-1", "RALCO-2", "RALCO-3"])
        self.assertEqual(embalses.ampliar(base, self.ruta), [])    # no duplica
        self.assertEqual(embalses.claves_ignoradas(self.ruta),
                         {embalses.clave("PE-CANELA-2")})
        # cambia de idea (mayusculas distintas = la misma central): pasa a agregadas
        embalses.guardar(agregar=["pe-canela-2"], ruta=self.ruta)
        ag, ig = embalses.leer(self.ruta)
        self.assertEqual(ig, [])
        self.assertEqual(ag, ["RALCO-3", "pe-canela-2"])

    def test_no_pisa_un_archivo_roto(self):
        self.ruta.parent.mkdir(parents=True)
        self.ruta.write_text("{roto", encoding="utf-8")
        self.assertFalse(embalses.guardar(["X-1"], ruta=self.ruta))
        self.assertEqual(self.ruta.read_text(encoding="utf-8"), "{roto")

    def test_formato(self):
        embalses.guardar(["A-1"], ["B-2"], self.ruta)
        self.assertEqual(json.loads(self.ruta.read_text(encoding="utf-8")),
                         {"agregadas": ["A-1"], "ignoradas": ["B-2"]})


if __name__ == "__main__":
    unittest.main()
