"""Pruebas de Actualiza_Cuadro1.py sin Excel (solo biblioteca estandar).

Lectura de las hojas PAGO_EMPRESAS_* (tipo en B1, sin la fila TOTAL) y del
RESUMEN de la planilla 8, la combinacion que conserva la planilla que no se
actualiza, y las formulas de I:K con el largo justo.

    python Revisor_Relq_15min/actualizadores/test_cuadro1.py
"""

import importlib.util
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
from test_pagos_varios import escribir_xlsx  # noqa: E402  (tambien falsea tkinter)

_spec = importlib.util.spec_from_file_location("cuadro1", AQUI / "Actualiza_Cuadro1.py")
c1 = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(c1)


def hoja_p4(tipo, filas):
    c = {(1, "A"): "Empresa", (1, "B"): tipo,
         (2, "B"): "RECIBE", (2, "C"): "PAGA", (2, "D"): "Paga+Recibe"}
    for i, (emp, rec, pag) in enumerate(filas, start=3):
        c[(i, "A")], c[(i, "B")], c[(i, "C")], c[(i, "D")] = emp, rec, pag, rec - pag
    f = 3 + len(filas)
    c[(f, "A")], c[(f, "B")] = "TOTAL", 999.0
    return c


class Cuadro1(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        self.p4 = self.tmp / "4_REMUNERACION_SC_CO_Y_Pagos_Retiros_2601.xlsx"
        escribir_xlsx(self.p4, {
            "PAGO_Configuracion": {(1, "A"): "x"},
            "PAGO_EMPRESAS_SC_SSCC": hoja_p4("SC_SSCC", [("ABASTIBLE", 0.0, 10.0),
                                                         ("ENEL", 10.0, 0.0)]),
            "PAGO_EMPRESAS_CO_ERNC": hoja_p4("CO_ERNC", [("ABASTIBLE", 0.0, 0.0)]),
        })
        self.p8 = self.tmp / "8_REMUNERACION_CONTROL_CONTINGENCIAS_2601.xlsx"
        escribir_xlsx(self.p8, {"RESUMEN": {
            (8, "C"): "Tipo", (8, "D"): "Empresa",
            (9, "C"): "EDAC-BF", (9, "D"): "COLBUN", (9, "E"): 5.0, (9, "F"): 0.0,
            (9, "G"): 5.0,
            (10, "C"): "Total", (10, "D"): "TOTAL", (10, "E"): 5.0,
        }})

    def test_leer_p4(self):
        tipos, filas = c1.leer_p4(self.p4, lambda *_: None)
        self.assertEqual(tipos, ["SC_SSCC", "CO_ERNC"])
        self.assertEqual(filas, [["SC_SSCC", "ABASTIBLE", 0.0, 10.0, -10.0],
                                 ["SC_SSCC", "ENEL", 10.0, 0.0, 10.0],
                                 ["CO_ERNC", "ABASTIBLE", 0.0, 0.0, 0.0]])

    def test_leer_p8(self):
        tipos, filas = c1.leer_p8(self.p8, lambda *_: None)
        self.assertEqual(tipos, ["EDAC-BF"])
        self.assertEqual(filas, [["EDAC-BF", "COLBUN", 5.0, 0.0, 5.0]])

    def test_combinar_conserva_la_que_no_se_actualiza(self):
        viejas = [["CO ERNC", "VIEJA4", 1, 1, 0],        # de la 4 ("CO ERNC" = "CO_ERNC")
                  ["EDAC-BF", "VIEJA8", 2, 2, 0],         # de la 8
                  ["CPF(-) ID", "FORMATO_VIEJO", 0, 0, 0]]
        nuevas4 = [["SC_SSCC", "N4", 0, 1, -1]]
        nuevas8 = [["EDAC-BF", "N8", 1, 0, 1]]
        t4, t8 = ["SC_SSCC", "CO_ERNC"], ["EDAC-BF"]
        log = []
        f, n4, n8 = c1.combinar(viejas, t4, t8, nuevas4, nuevas8, True, False, log.append)
        self.assertEqual([x[1] for x in f], ["N4", "VIEJA8"])
        self.assertTrue(any("FORMATO_VIEJO" not in l and "CPF(-) ID (1)" in l for l in log))
        f, _, _ = c1.combinar(viejas, t4, t8, nuevas4, nuevas8, False, True, lambda *_: None)
        self.assertEqual([x[1] for x in f], ["VIEJA4", "N8"])
        f, _, _ = c1.combinar(viejas, t4, t8, nuevas4, nuevas8, True, True, lambda *_: None)
        self.assertEqual([x[1] for x in f], ["N4", "N8"])
        # Sin la planilla 8 a mano: todo lo que no es de la 4 se toma como de la 8.
        f, _, n8 = c1.combinar(viejas, t4, None, nuevas4, [], True, False, lambda *_: None)
        self.assertEqual(n8, 2)

    def test_formulas_ik(self):
        self.assertEqual(c1.n_empresas([["a", "Enel"], ["b", "ENEL"], ["c", "Colbun"]]), 2)
        f = c1.formulas_empresas(2, 120)
        self.assertEqual(len(f), 2)
        self.assertIn("COUNTIF($I$8:I9,$C$9:$C$120)", f[1][0])
        self.assertIn("SUMIF($C$9:$C$120,I10,$G$9:$G$120)<-0.00001", f[1][1])
        self.assertIn(">0.00001", f[1][2])


if __name__ == "__main__":
    unittest.main()
