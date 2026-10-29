"""Pruebas del pegado de SOBRECOSTOS desde SC_CfrecSinFD_*_para_coyailen.csv
(Actualiza_datos.py), con una hoja falsa: sin Excel.

    python Revisor_Relq_15min/actualizadores/test_sobrecostos_coyailen.py
"""

import importlib.util
import shutil
import sys
import tempfile
import types
import unittest
from datetime import datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import test_pagos_varios  # noqa: E402,F401  (falsea tkinter)
import test_sc_cfrec  # noqa: E402  (arma la entrada y la salida reales)

_spec = importlib.util.spec_from_file_location("act_datos", AQUI / "Actualiza_datos.py")
ad = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ad)


class Rango:
    def __init__(self, hoja, a, b):
        self.hoja, self.a, self.b = hoja, a, b

    @property
    def value(self):
        raise NotImplementedError

    @value.setter
    def value(self, filas):
        (r0, c0), (r1, c1) = self.a, self.b
        assert len(filas) == r1 - r0 + 1 and all(len(f) == c1 - c0 + 1 for f in filas)
        for i, f in enumerate(filas):
            for j, v in enumerate(f):
                self.hoja.celdas[(r0 + i, c0 + j)] = v

    def clear_contents(self):
        (r0, c0), (r1, c1) = self.a, self.b
        for k in list(self.hoja.celdas):
            if r0 <= k[0] <= r1 and c0 <= k[1] <= c1:
                del self.hoja.celdas[k]


class Hoja:
    def __init__(self, name, celdas):
        self.name, self.celdas = name, dict(celdas)
        self.api = types.SimpleNamespace(FilterMode=False, ListObjects=[])
        self.api.UsedRange = self

    # UsedRange falso: Row y Rows.Count sobre las celdas con algo
    @property
    def Row(self):
        return min(r for r, _ in self.celdas)

    @property
    def Rows(self):
        return types.SimpleNamespace(Count=max(r for r, _ in self.celdas) - self.Row + 1)

    def range(self, a, b):
        return Rango(self, a, b)


class SobrecostosCoyailen(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        caso = test_sc_cfrec.ScCfrec("test_reorganiza")
        caso.setUp()
        self.addCleanup(caso.doCleanups)
        self.csv, *_ = test_sc_cfrec.sc.procesar(caso.entrada, log=lambda *_: None)
        self.caso_dir = caso.tmp

    def test_valor_csv(self):
        self.assertEqual(ad.valor_csv("3.75"), 3.75)
        self.assertEqual(ad.valor_csv("-1e-3"), -0.001)
        self.assertEqual(ad.valor_csv("01-04-2026"), datetime(2026, 4, 1))
        self.assertEqual(ad.valor_csv("01/04/2026 09:15"), datetime(2026, 4, 1, 9, 15))
        self.assertEqual(ad.valor_csv("2026-04-01"), datetime(2026, 4, 1))
        self.assertIsNone(ad.valor_csv(""))
        self.assertEqual(ad.valor_csv("CPF(+)"), "CPF(+)")
        self.assertEqual(ad.valor_csv("TER QUINTERO U1"), "TER QUINTERO U1")

    def test_pega_desde_a7_y_borra_sobrante(self):
        viejas = {(6, 1): "Fecha", (6, 30): "encabezado AD"}
        for r in range(7, 20):                      # pegada anterior, mas larga
            for c in range(1, 19):
                viejas[(r, c)] = "viejo"
            viejas[(r, 30)] = "fuera de A:R"
        sh = Hoja("SOBRECOSTOS", viejas)
        sh.api.FilterMode = True
        sh.api.ShowAllData = lambda: setattr(sh.api, "FilterMode", False)
        wb = types.SimpleNamespace(name="Calculo.xlsm", sheets=[Hoja("FD_CT", {(1, 1): 1}), sh])
        log = []
        n = ad.pegar_sobrecostos_coyailen(wb, self.csv, log_func=log.append)
        self.assertEqual(n, 3)
        self.assertFalse(sh.api.FilterMode)                 # se quito el filtro
        self.assertEqual(sh.celdas[(6, 1)], "Fecha")        # encabezado intacto
        self.assertEqual(sh.celdas[(7, 1)], datetime(2026, 1, 1))
        self.assertEqual(sh.celdas[(7, 4)], "CONF1")
        self.assertEqual(sh.celdas[(7, 5)], 3.5)
        self.assertEqual(sh.celdas[(9, 5)], 3.75)           # bloque combinado
        self.assertEqual(sh.celdas[(9, 18)], 2.0)           # R = bloque
        self.assertIsNone(sh.celdas[(7, 12)])               # bandera vacia: celda vacia
        self.assertFalse(any(r >= 10 and c <= 18 for r, c in sh.celdas))
        self.assertEqual(sh.celdas[(19, 30)], "fuera de A:R")

    def test_busca_la_salida_en_01_sobrecostos(self):
        self.assertEqual(ad.buscar_sc_cfrec_coyailen(self.caso_dir), self.csv)

    def test_sin_hoja_sobrecostos(self):
        wb = types.SimpleNamespace(name="x.xlsm", sheets=[Hoja("OTRA", {(1, 1): 1})])
        with self.assertRaises(RuntimeError):
            ad.pegar_sobrecostos_coyailen(wb, self.csv, log_func=lambda *_: None)


if __name__ == "__main__":
    unittest.main()
