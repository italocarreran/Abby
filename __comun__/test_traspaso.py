"""Pruebas stdlib del contrato opcional de traspaso."""

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __comun__ import traspaso


class TestTraspaso(unittest.TestCase):
    def test_sin_argumento_es_modo_manual(self):
        self.assertIsNone(traspaso.leer_argumento(["script.py"]))
        self.assertIsNone(traspaso.leer_argumento(["script.py", "  "]))

    def test_archivo_inexistente_es_modo_manual(self):
        self.assertIsNone(traspaso.leer_argumento(["script.py", "no-existe.json"]))

    def test_json_roto_es_modo_manual(self):
        with tempfile.TemporaryDirectory() as td:
            ruta = Path(td) / "traspaso.json"
            ruta.write_text("{roto", encoding="utf-8")
            self.assertIsNone(traspaso.leer(ruta))

    def test_rechaza_origen_ajeno(self):
        self.assertIsNone(traspaso.validar({"origen": "otro", "version": 1}))

    def test_rechaza_version_futura(self):
        data = {"origen": traspaso.ORIGEN, "version": 2, "rutas": {}}
        self.assertIsNone(traspaso.validar(data))

    def test_version_invalida_es_modo_manual(self):
        data = {"origen": traspaso.ORIGEN, "version": "invalida", "rutas": {}}
        self.assertIsNone(traspaso.validar(data))

    def test_normaliza_rutas_ausentes_o_invalidas(self):
        for rutas in (None, [], "ruta"):
            data = {"origen": traspaso.ORIGEN, "version": 1, "rutas": rutas}
            validado = traspaso.validar(data)
            self.assertIsNotNone(validado)
            assert validado is not None
            self.assertEqual(validado["rutas"], {})

    def test_conserva_datos_validos(self):
        with tempfile.TemporaryDirectory() as td:
            ruta = Path(td) / "traspaso.json"
            esperado = {
                "origen": traspaso.ORIGEN,
                "version": 1,
                "aamm": "2609",
                "rutas": {"cuadro_0": "destino.xlsm"},
            }
            ruta.write_text(json.dumps(esperado), encoding="utf-8")
            self.assertEqual(
                traspaso.leer_argumento(["script.py", str(ruta)]), esperado
            )


if __name__ == "__main__":
    unittest.main()
