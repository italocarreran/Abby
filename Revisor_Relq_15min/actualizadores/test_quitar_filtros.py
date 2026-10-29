"""Pruebas de quitar_filtros (Actualiza_SC_CO.py) con hojas falsas, sin Excel.

    python Revisor_Relq_15min/actualizadores/test_quitar_filtros.py
"""

import importlib.util
import sys
import types
import unittest
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import test_pagos_varios  # noqa: E402,F401  (falsea tkinter)

sys.modules.setdefault("xlwings", types.ModuleType("xlwings"))
_spec = importlib.util.spec_from_file_location("scco", AQUI / "Actualiza_SC_CO.py")
scco = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(scco)


class Filtro:
    def __init__(self, activo):
        self.FilterMode = activo
        self.limpiado = False

    def ShowAllData(self):
        if not self.FilterMode:
            raise RuntimeError("ShowAllData sin filtro falla en Excel")
        self.FilterMode = False
        self.limpiado = True


def hoja(filtro_hoja, tablas=()):
    api = Filtro(filtro_hoja)
    api.ListObjects = [types.SimpleNamespace(AutoFilter=t) for t in tablas]
    return types.SimpleNamespace(name="FD_CPF", api=api)


class QuitarFiltros(unittest.TestCase):
    def test_hoja_filtrada(self):
        sh, log = hoja(True), []
        self.assertEqual(scco.quitar_filtros(sh, log.append), 1)
        self.assertTrue(sh.api.limpiado)
        self.assertTrue(any("FD_CPF" in l for l in log))

    def test_sin_filtro_no_toca_ni_avisa(self):
        sh, log = hoja(False, [None, Filtro(False)]), []
        self.assertEqual(scco.quitar_filtros(sh, log.append), 0)
        self.assertEqual(log, [])

    def test_tabla_filtrada(self):
        t = Filtro(True)
        self.assertEqual(scco.quitar_filtros(hoja(False, [t]), lambda *_: None), 1)
        self.assertTrue(t.limpiado)


if __name__ == "__main__":
    unittest.main()
