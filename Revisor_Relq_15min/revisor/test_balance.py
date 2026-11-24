# -*- coding: utf-8 -*-
"""Pruebas de revisor/balance.py y revisor/cambio_hora.py (recuadro «Balance
del caso» del Revisor 15 min)."""

from pathlib import Path
import sys
import tempfile
import unittest

RAIZ = Path(__file__).resolve().parents[2]
REVISOR = Path(__file__).resolve().parents[1]
for carpeta in (RAIZ, REVISOR):
    if str(carpeta) not in sys.path:
        sys.path.insert(0, str(carpeta))

from __comun__ import salidas as _sal
from revisor import balance as B
from revisor import cambio_hora as CH


class NombresTests(unittest.TestCase):
    def test_aamm_ida_y_vuelta(self):
        self.assertEqual(B.aamm_de("2026", "Enero"), "2601")
        self.assertEqual(B.aamm_de("2025", "Diciembre"), "2512")
        self.assertEqual(B.ano_mes_de("2601"), ("2026", "Enero"))
        self.assertEqual(B.aamm_de("", "Enero"), "")
        self.assertEqual(B.ano_mes_de("2613"), ("", ""))
        self.assertEqual(B.ano_mes_de(None), ("", ""))

    def test_nombre_y_clave(self):
        self.assertEqual(B.nombre_balance("2026", "Enero", "Definitivo"),
                         "2026_Enero_Definitivo")
        self.assertEqual(B.clave_balance("2601", "Definitivo"), "2601_Definitivo")
        self.assertEqual(B.clave_balance("2601", ""), "2601")
        self.assertEqual(B.partir_clave("2601_R01P"), ("2601", "R01P"))
        self.assertEqual(B.partir_clave("2601"), ("2601", ""))
        self.assertEqual(B.nombre_de_clave("2601_R01P"), "2026_Enero_R01P")

    def test_anos_incluyen_los_recordados(self):
        self.assertEqual(B.anos_disponibles(2026), ["2025", "2026"])
        self.assertEqual(B.anos_disponibles(2026, extra=["2024"]),
                         ["2024", "2025", "2026"])

    def test_procesos(self):
        self.assertEqual(B.procesos_de({}), ["Preliminar", "Definitivo"])
        cfg = {"procesos": ["R01P", "r01p", "mal nombre!", ""]}
        self.assertEqual(B.procesos_de(cfg),
                         ["Preliminar", "Definitivo", "R01P", "malnombre"])
        self.assertEqual(B.con_proceso_nuevo({}, " R01D "),
                         ("R01D", ["R01D"]))
        self.assertEqual(B.con_proceso_nuevo(cfg, "definitivo"),
                         ("Definitivo", None))
        self.assertEqual(B.con_proceso_nuevo({}, "  "), ("", None))

    def test_proceso_en_el_nombre(self):
        self.assertEqual(B.proceso_en_el_nombre("T:/2601 Definitivo/02 CASO"),
                         "Definitivo")
        self.assertEqual(B.proceso_en_el_nombre("T:/2601/02 CASO RELIQUIDACION"), "")
        self.assertEqual(B.proceso_en_el_nombre("T:/Preliminar y Definitivo"), "")


class CarpetasTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = Path(self.tmp.name)
        self.pre = self.d / "2601 Preliminar" / "02 CASO RELIQUIDACION"
        self.dfn = self.d / "2601 Definitivo" / "02 CASO RELIQUIDACION"
        self.pre.mkdir(parents=True)
        self.dfn.mkdir(parents=True)

    def tearDown(self):
        self.tmp.cleanup()

    def test_carpeta_por_balance(self):
        cfg = {"carpetas_por_balance": {"2601_Preliminar": str(self.pre)}}
        self.assertEqual(B.carpeta_del_balance(cfg, "2601", "Preliminar"), str(self.pre))
        self.assertEqual(B.carpeta_del_balance(cfg, "2601", "Definitivo"), "")

    def test_carpeta_vieja_por_mes_si_el_nombre_dice_el_proceso(self):
        cfg = {"carpetas_por_mes": {"2601": str(self.dfn)}}
        self.assertEqual(B.carpeta_del_balance(cfg, "2601", "Definitivo"), str(self.dfn))
        self.assertEqual(B.carpeta_del_balance(cfg, "2601", "Preliminar"), "")

    def test_carpeta_que_ya_no_existe(self):
        cfg = {"carpetas_por_balance": {"2601_Preliminar": str(self.d / "no")}}
        self.assertEqual(B.carpeta_del_balance(cfg, "2601", "Preliminar"), "")

    def test_una_carpeta_es_de_un_solo_balance(self):
        cfg = {"carpetas_por_balance": {"2601_Preliminar": str(self.pre),
                                        "2512_Definitivo": "X"}}
        self.assertEqual(B.balance_de_la_carpeta(cfg, self.pre), "2601_Preliminar")
        nuevas = B.carpetas_con(cfg, "2601_Definitivo", str(self.pre))
        self.assertEqual(nuevas, {"2601_Definitivo": str(self.pre),
                                  "2512_Definitivo": "X"})


class EstadoPorBalanceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.raiz = Path(self.tmp.name)
        self.dir_mes = lambda aamm, crear=False: _sal.carpeta_mes(self.raiz, aamm, crear=crear)

    def tearDown(self):
        self.tmp.cleanup()

    def test_ruta_del_estado(self):
        mes = self.raiz / "2026" / "01 Enero"
        self.assertEqual(B.dir_balance(self.dir_mes, "2601"), mes)
        self.assertEqual(B.dir_balance(self.dir_mes, "2601_Definitivo"),
                         mes / "Definitivo")
        self.assertFalse((mes / "Definitivo").exists())
        B.dir_balance(self.dir_mes, "2601_Definitivo", crear=True)
        self.assertTrue((mes / "Definitivo").is_dir())

    def test_hereda_copiando_sin_pisar(self):
        mes = self.dir_mes("2601", crear=True)
        (mes / "a.json").write_text("viejo", encoding="utf-8")
        (mes / "b.json").write_text("viejo", encoding="utf-8")
        destino = B.dir_balance(self.dir_mes, "2601_Definitivo", crear=True)
        (destino / "b.json").write_text("nuevo", encoding="utf-8")
        copiados = B.heredar_estado_del_mes(
            self.dir_mes, "2601_Definitivo", ["a.json", "b.json", "c.json"])
        self.assertEqual(copiados, ["a.json"])
        self.assertEqual((destino / "a.json").read_text(encoding="utf-8"), "viejo")
        self.assertEqual((destino / "b.json").read_text(encoding="utf-8"), "nuevo")
        self.assertTrue((mes / "a.json").exists())        # no se movio
        self.assertEqual(B.heredar_estado_del_mes(self.dir_mes, "2601", ["a.json"]), [])


class BloqueoTests(unittest.TestCase):
    def test_bloquear_y_desbloquear(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertIsNone(B.info_bloqueo(d))
            B.bloquear(d, "PC_usuaria")
            info = B.info_bloqueo(d)
            self.assertEqual(info["quien"], "PC_usuaria")
            self.assertTrue(info["cuando"])
            self.assertTrue(B.desbloquear(d))
            self.assertIsNone(B.info_bloqueo(d))
            self.assertFalse(B.desbloquear(d))
        self.assertIsNone(B.info_bloqueo(""))


class CambioHoraTests(unittest.TestCase):
    def test_bloques_del_mes(self):
        self.assertEqual(CH.bloques_del_mes("2601"), 31 * 96)
        self.assertEqual(CH.bloques_del_mes("2602"), 28 * 96)
        self.assertEqual(CH.bloques_del_mes("2604"), 30 * 96 + 4)
        self.assertEqual(CH.bloques_del_mes("2609"), 30 * 96 - 4)

    def test_descripcion(self):
        self.assertIn("Sin cambio de hora", CH.describir_periodo("2601"))
        self.assertIn("se atrasa la hora", CH.describir_periodo("2604"))
        self.assertIn("se adelanta la hora", CH.describir_periodo("2609"))

    def test_regla_sin_tzdata_da_lo_mismo(self):
        for mes in (4, 9):
            regla = CH._cambios_por_regla(2026, mes)
            real = CH.cambios_de_hora(2026, mes)
            self.assertEqual([c.fecha for c in regla], [c.fecha for c in real])


if __name__ == "__main__":
    unittest.main()
