"""Pruebas stdlib de persistencia interna del Revisor."""

import json
from pathlib import Path
import tempfile
import sys
import unittest
from pathlib import Path as _Path

# Para poder correrlo suelto: la raiz del repo (para __comun__) y
# Revisor_Relq (para el paquete revisor).
_AQUI = _Path(__file__).resolve()
for _p in (_AQUI.parents[2], _AQUI.parents[1]):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from revisor.estado import ARCHIVO_CACHE, ARCHIVO_ESTADO, CacheValores, Estado, leer_estado_mes


def escribir(ruta, data):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(json.dumps(data), encoding="utf-8")


class EstadoTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.raiz = Path(self.tmp.name)
        self.dir_mes = lambda aamm, crear=False: self._dir_mes(aamm, crear)

    def tearDown(self):
        self.tmp.cleanup()

    def _dir_mes(self, aamm, crear=False):
        ruta = self.raiz / str(aamm)
        if crear:
            ruta.mkdir(parents=True, exist_ok=True)
        return ruta

    def test_estado_conserva_nombre_formato_y_firma(self):
        estado = Estado(self.dir_mes, escribir, lambda vid: "firma-" + vid)
        estado.cargar("2407")
        self.assertTrue(estado.set("V4", {"estado": "OK"}))
        ruta = self.raiz / "2407" / ARCHIVO_ESTADO
        self.assertEqual(json.loads(ruta.read_text())["V4"]["firma"], "firma-V4")
        self.assertEqual(leer_estado_mes("2407", self.dir_mes), estado.data)

    def test_cache_valida_huella_fecha_y_tamano(self):
        origen = self.raiz / "origen.xlsx"
        origen.write_bytes(b"abc")
        cache = CacheValores(
            self.dir_mes, escribir, lambda _ruta: 12.5,
            lambda ruta: ruta.stat().st_size, lambda ts: f"fecha:{ts}",
        )
        cache.cargar("2407")
        cache.poner("TOTAL", origen, "hoja|A1", 9.0, filas=2)
        registro = cache.obtener("TOTAL", origen, "hoja|A1")
        self.assertIsNotNone(registro)
        assert registro is not None
        self.assertEqual(registro["valor"], 9.0)
        self.assertIsNone(cache.obtener("TOTAL", origen, "otra"))
        self.assertTrue((self.raiz / "2407" / ARCHIVO_CACHE).exists())
        cache.descartar(["TOTAL"])
        self.assertEqual(cache.data, {})

    def test_json_roto_no_se_propaga(self):
        ruta = self.raiz / "2407" / ARCHIVO_ESTADO
        ruta.parent.mkdir()
        ruta.write_text("{", encoding="utf-8")
        estado = Estado(self.dir_mes, escribir, lambda _vid: "firma")
        self.assertFalse(estado.cargar("2407"))
        self.assertEqual(estado.data, {})


if __name__ == "__main__":
    unittest.main()
