"""Pruebas stdlib de búsqueda y caché internos del Revisor."""

import os
from pathlib import Path
import tempfile
import unittest

from revisor import archivos


class ArchivosTests(unittest.TestCase):
    def test_busqueda_normalizada_y_descarta_copias(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            carpeta = base / "01 Sobrecóstos"
            carpeta.mkdir()
            original = carpeta / "Calculo_SSCC_2407.xlsm"
            copia = carpeta / "Calculo_SSCC_2407 - copia.xlsm"
            original.write_text("original", encoding="utf-8")
            copia.write_text("copia", encoding="utf-8")
            os.utime(original, (1, 1))
            os.utime(copia, (2, 2))
            self.assertEqual(archivos.buscar_carpeta(base, "01 sobrecostos"), carpeta)
            self.assertEqual(
                archivos.buscar_archivo(carpeta, r"calculo_sscc", (".xlsm",)),
                original,
            )

    def test_cache_recorre_una_vez_y_se_apaga_al_salir(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            (base / "uno.txt").write_text("1", encoding="utf-8")
            with archivos.cache_directorios() as cache:
                archivos.leer_dir(base)
                archivos.leer_dir(base)
            self.assertEqual(cache.stats, (1, 1))
            archivos.leer_dir(base)
            self.assertEqual(cache.stats, (1, 1))

    def test_diarios_prefiere_original(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            (base / "Detalle_20240701 - copia.xlsx").touch()
            original = base / "Detalle_20240701.xlsx"
            original.touch()
            self.assertEqual(
                archivos.listar_diarios(base, r"(\d{8})", (".xlsx",))["20240701"],
                original,
            )


if __name__ == "__main__":
    unittest.main()
