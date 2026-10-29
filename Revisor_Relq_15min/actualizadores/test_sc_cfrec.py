"""Pruebas de Reorganiza_SC_Cfrec.py (solo biblioteca estandar).

Arma un SC_CfrecSinFD minimo con columnas A..Z y comprueba el mapeo, la union
de los bloques con generacion en J, la busqueda en 01 Sobrecostos y que una
entrada mala no pisa la salida anterior. tkinter se falsea como en
test_pagos_varios.py.

    python Revisor_Relq_15min/actualizadores/test_sc_cfrec.py
"""

import csv
import importlib.util
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import test_pagos_varios  # noqa: E402,F401  (falsea tkinter)

_spec = importlib.util.spec_from_file_location("sc_cfrec", AQUI / "Reorganiza_SC_Cfrec.py")
sc = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(sc)

LETRAS = [chr(c) for c in range(ord("A"), ord("Z") + 1)]
ENC = {"B": "Bloque_15min", "D": "Fecha", "E": "Hora", "F": "CONFIGURACION",
       "G": "UNIDAD GENERADORA", "I": "GENERACION", "J": "GEN TG+TV", "K": "CV",
       "L": "CMg[USD/MWh]", "M": "Embalse", "N": "CPF(+)", "O": "CPF(-)",
       "P": "CSF(+)", "Q": "CSF(-)", "R": "CTF(+)", "S": "CTF(-)", "X": "Dolar",
       "Y": "Sobrecosto", "Z": "Tipo"}


def fila(**v):
    return [str(v.get(l, "")) for l in LETRAS]


class ScCfrec(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        self.sob = self.tmp / "01 Sobrecostos"
        self.sob.mkdir()
        self.entrada = self.sob / "SC_CfrecSinFD_2601.csv"
        filas = [
            [ENC.get(l, l) for l in LETRAS],
            # normal, no embalse: Central = F
            fila(A="k1", B=1, D="01-01-2026", E=1, F="CONF1", G="U1", I=10, J=0,
                 K=5, L=7, M=0, N=1, X=900, Y=3.5, Z="SC"),
            # embalse: Central = G, H mantiene F
            fila(A="k1", B=1, D="01-01-2026", E=1, F="EMB", G="U_EMB", I=4, J=0,
                 M=1, Q=1, Y=1, Z="SC"),
            # dos filas con J != 0, misma config y cuarto de hora: se unen
            fila(A="k2", B=2, D="01-01-2026", E=1, F="CC", G="TG1", I=2, J=5,
                 M=0, Y=1.5, Z="SC"),
            [""] * 26,
            fila(A="k2", B=2, D="01-01-2026", E=1, F="CC", G="TV1", I=3, J=5,
                 M=0, S=1, Y=2.25, Z="SC"),
        ]
        with self.entrada.open("w", encoding="cp1252", newline="") as g:
            csv.writer(g).writerows(filas)

    def test_reorganiza(self):
        sal, n_in, n_out, n_gr = sc.procesar(self.entrada, log=lambda *_: None)
        self.assertEqual(sal.name, "SC_CfrecSinFD_2601_para_coyailen.csv")
        self.assertEqual((n_in, n_out, n_gr), (4, 3, 1))
        self.assertTrue(sal.read_bytes().startswith(b"\xef\xbb\xbf"))
        leido = list(csv.reader(sal.read_text(encoding="utf-8-sig").splitlines()))
        self.assertEqual(leido[0][:8], ["Fecha", "Hora", "Tipo", "Central", "Sobrecosto",
                                        "UNIDAD GENERADORA", "GENERACION", "Configuración"])
        self.assertEqual(leido[0][10:], ["CPF(+)", "CPF(-)", "CSF(+)", "CTF(+)", "CSF(-)",
                                         "CTF(-)", "Dolar", "Bloque_15min"])
        self.assertEqual(leido[1], ["01-01-2026", "1", "SC", "CONF1", "3.5", "U1", "10",
                                    "CONF1", "5", "7", "CPF(+)", "", "", "", "", "", "900",
                                    "1"])
        self.assertEqual(leido[2][3], "U_EMB")
        self.assertEqual(leido[2][7], "EMB")
        self.assertEqual(leido[2][14], "CSF(-)")
        self.assertEqual(leido[3][3:7], ["CC", "3.75", "TG1", "5.0"])
        self.assertEqual(leido[3][15], "CTF(-)")

    def test_busca_en_01_sobrecostos_sin_la_salida(self):
        sc.procesar(self.entrada, log=lambda *_: None)
        self.assertEqual(sc.buscar_entrada(self.tmp), self.entrada)
        self.assertIsNone(sc.buscar_entrada(self.sob))

    def test_entrada_mala_no_pisa(self):
        sal = sc.ruta_salida(self.entrada)
        sal.write_text("anterior", encoding="utf-8")
        self.entrada.write_text("a,b,c\n1,2,3\n", encoding="utf-8")
        with self.assertRaises(ValueError):
            sc.procesar(self.entrada, log=lambda *_: None)
        self.assertEqual(sal.read_text(encoding="utf-8"), "anterior")
        self.assertEqual(sorted(p.name for p in self.sob.iterdir()),
                         [self.entrada.name, sal.name])


if __name__ == "__main__":
    unittest.main()
