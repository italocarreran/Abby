# -*- coding: utf-8 -*-
"""Pruebas de la comprobacion ``pago_por_concepto`` (V15, modo 15 minutos):
total por tipo de pago de la planilla 9 contra la planilla 1."""

from pathlib import Path
import sys
import tempfile
import unittest

RAIZ = Path(__file__).resolve().parents[2]
REVISOR = Path(__file__).resolve().parents[1]
for carpeta in (RAIZ, REVISOR):
    if str(carpeta) not in sys.path:
        sys.path.insert(0, str(carpeta))

from __comun__ import texto as _texto
from revisor import lectores, verificaciones

P9 = """Clave_Anio_Mes,nombre_barra,tension,nombre_barra_cmg,Suministrador,Retiro,clave,Tipo_pago,Tipo_ret,pago
2601,B,13,B,EMP1,R,K,CRA,L_D,100.5
2601,B,13,B,EMP2,R,K,CRA,L_D,50
2601,B,13,B,EMP1,R,K,CO_ERNC,L_D,30
2601,B,13,B,EMP1,R,K,REA,L_D,7
2601,B,13,B,EMP1,R,K,C.Frec,L_D,999
"""

# Planilla 1 como csv: B concepto, E PAGA, desde la fila 9 (1-8 de relleno:
# el lector de csv salta las lineas vacias y saca el separador de la 1).
P1_OK = "relleno,,,,\n" * 8 + """,CRA,EMP1,0,120.5
,CRA,EMP2,0,30
,CO ERNC,EMP1,0,30
,REA,EMP1,0,-7
,SC_SSCC,EMP1,0,-5
,EDAC-BF,EMP1,0,4
"""


def comprobacion():
    return {"tipo": "pago_por_concepto", "desc": "total por tipo",
            "detalle": {"archivo": "p9", "hoja": "(csv)", "fila_inicio": 2,
                        "nombre": "P9", "col_concepto": "H", "col_monto": "J"},
            "excluir_conceptos": ["C.Frec", "SCMT", "OTRO", "EDAC-BF"],
            "resumen": {"archivo": "p1", "hoja": "(csv)", "fila_inicio": 9,
                        "nombre": "P1", "col_concepto": "B", "col_monto": "E"},
            "tolerancia": 1.0}


class Revisor:
    def __init__(self, rutas):
        self.rutas = rutas


class PagoPorConceptoTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        verificaciones.configurar({
            "clave_concepto": _texto.clave_mayusculas,
            "leer_columnas_rapido": lectores.leer_columnas_rapido,
            "fmt_monto": lambda v: f"{v:.2f}",
            "TOLERANCIA": 1.0,
        })

    def correr(self, p1):
        with tempfile.TemporaryDirectory() as d:
            r9, r1 = Path(d, "p9.csv"), Path(d, "p1.csv")
            r9.write_text(P9, encoding="utf-8")
            r1.write_text(p1, encoding="utf-8")
            log = []
            res = verificaciones.comprobar(Revisor({"p9": r9, "p1": r1}),
                                           comprobacion(), {}, log.append)
            return res, "\n".join(log)

    def test_totales_iguales_aunque_difiera_por_empresa(self):
        # CRA: 150,5 en los dos lados, repartido distinto entre empresas.
        res, _ = self.correr(P1_OK)
        self.assertEqual(res["estado"], "OK")
        self.assertEqual(sorted(res["comunes"]), ["CO_ERNC", "CRA", "REA"])

    def test_co_ernc_con_y_sin_guion_bajo(self):
        res, _ = self.correr(P1_OK)
        self.assertIn("CO_ERNC", res["comunes"])

    def test_excluidos_y_los_de_un_solo_lado_no_se_comparan(self):
        res, log = self.correr(P1_OK)
        self.assertNotIn("C.Frec", res["comunes"])
        self.assertNotIn("EDAC-BF", res["comunes"] + res["solo_resumen"])
        self.assertEqual(res["solo_resumen"], ["SC_SSCC"])

    def test_descuadre_de_un_tipo(self):
        res, log = self.correr(P1_OK.replace("CRA,EMP2,0,30", "CRA,EMP2,0,40"))
        self.assertEqual(res["estado"], "NO CUADRA")
        self.assertEqual([d[0] for d in res["difs"]], ["CRA"])
        self.assertEqual(res["n_ok"], 2)

    def test_sin_tipos_en_comun(self):
        res, _ = self.correr("relleno,,,,\n" * 8 + ",SC_SSCC,EMP1,0,5\n")
        self.assertEqual(res["estado"], "SIN DATOS")


if __name__ == "__main__":
    unittest.main()
