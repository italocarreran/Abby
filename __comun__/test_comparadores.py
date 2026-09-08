"""Pruebas stdlib de infraestructura común de comparadores."""

import sys
import tempfile
import threading
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __comun__ import comparadores


class FakeRoot:
    def __init__(self): self.after_calls = []
    def after(self, demora, funcion): self.after_calls.append((demora, funcion))


class FakeText:
    def __init__(self): self.insertados = []; self.seen = []
    def insert(self, donde, texto): self.insertados.append((donde, texto))
    def see(self, donde): self.seen.append(donde)


class FakeVar:
    def __init__(self): self.valor = None
    def set(self, valor): self.valor = valor


class FakeBarra:
    def __init__(self): self.maximum = 10; self.configs = []
    def __getitem__(self, clave): return self.maximum
    def config(self, **kw): self.configs.append(kw)


class TestComparadores(unittest.TestCase):
    def test_meses(self):
        self.assertEqual(comparadores.meses_del_anio("2026")[0:2], ["2601", "2602"])
        self.assertEqual(comparadores.meses_del_anio("26")[-1], "2612")
        self.assertEqual(comparadores.meses_del_anio(""), [])

    def test_inclusion(self):
        estado = {}
        self.assertTrue(comparadores.mes_incluido(estado, "2601"))
        comparadores.fijar_incluido(estado, "2601", False)
        self.assertFalse(comparadores.mes_incluido(estado, "2601"))

    def test_inclusion_usa_el_formato_en_disco_de_siempre(self):
        """La marca va en estado[aamm]["incluir"], que es lo que hay escrito en
        los estado.json que ya existen. Si se mueve a otra clave, las
        exclusiones viejas dejan de leerse y los meses vuelven al anual sin
        que nada avise."""
        estado = {}
        comparadores.fijar_incluido(estado, "2601", False)
        self.assertEqual(estado, {"2601": {"incluir": False}})
        # un estado.json ya escrito por una version anterior se sigue leyendo
        viejo = {"2602": {"incluir": False, "etapa1": {"consolidado": True}}}
        self.assertFalse(comparadores.mes_incluido(viejo, "2602"))
        # y fijar la marca no pisa el resto del registro del mes
        comparadores.fijar_incluido(viejo, "2602", True)
        self.assertTrue(comparadores.mes_incluido(viejo, "2602"))
        self.assertEqual(viejo["2602"]["etapa1"], {"consolidado": True})

    def test_cola_no_toca_widgets_hasta_bombear(self):
        root, txt, var, barra = FakeRoot(), FakeText(), FakeVar(), FakeBarra()
        puente = comparadores.ColaTk(root)
        puente.conectar(txt, var, barra)
        llamado = []

        def worker():
            puente.log("uno")
            puente.estado("ocupado")
            puente.progreso(value=3)
            puente.llamar(llamado.append, "hecho")

        hilo = threading.Thread(target=worker)
        hilo.start(); hilo.join()
        self.assertEqual(txt.insertados, [])
        puente.bombear()
        self.assertEqual(txt.insertados, [("end", "uno\n")])
        self.assertEqual(var.valor, "ocupado")
        self.assertEqual(barra.configs, [{"value": 3}])
        self.assertEqual(llamado, ["hecho"])

    def test_final_completa_barra(self):
        root, txt, var, barra = FakeRoot(), FakeText(), FakeVar(), FakeBarra()
        puente = comparadores.ColaTk(root)
        puente.conectar(txt, var, barra)
        puente.progreso(final=True)
        puente.bombear()
        self.assertEqual(barra.configs, [{"value": 10}])

    def test_color_semantico(self):
        colores = {"verde": "v", "amarillo": "a", "rojo": "r", "gris": "g"}
        self.assertEqual(comparadores.color_de("ok", colores), "v")
        self.assertEqual(comparadores.color_de("pendiente", colores), "a")
        self.assertEqual(comparadores.color_de("falta", colores), "r")
        # Los dos casos que se habian perdido: "desactualizado" es amarillo
        # como "pendiente" (el mes esta, hay que rehacerlo), no rojo, que
        # significa que falta el archivo; y lo desconocido cae en gris.
        self.assertEqual(comparadores.color_de("desactualizado", colores), "a")
        self.assertEqual(comparadores.color_de("otro", colores), "g")
        self.assertEqual(comparadores.color_de(None, colores), "g")

    def test_estado_y_firma(self):
        guardado = []
        self.assertEqual(comparadores.cargar_estado("x", lambda *_: []), {})
        self.assertEqual(comparadores.cargar_estado("x", lambda *_: {"a": 1}), {"a": 1})
        comparadores.guardar_estado("x", {"a": 1}, lambda *args: guardado.append(args))
        self.assertEqual(guardado, [("x", {"a": 1})])

    def test_hojas_propias(self):
        fijas = ["RESUMEN SSCC"]
        for nombre in ("RESUMEN", "RESUMEN SSCC", "2601", "2601_2"):
            self.assertTrue(comparadores.es_hoja_propia(nombre, fijas))
        for nombre in ("Notas", "2613", "2601_extra"):
            self.assertFalse(comparadores.es_hoja_propia(nombre, fijas))

    def test_respaldo_conserva_cinco(self):
        with tempfile.TemporaryDirectory() as td:
            raiz = Path(td)
            destino = raiz / "resultado.xlsx"
            destino.write_text("actual", encoding="utf-8")
            respaldos = raiz / "respaldos"
            for i in range(6):
                (respaldos / f"resultado_20000101_00000{i}.xlsx").parent.mkdir(exist_ok=True)
                (respaldos / f"resultado_20000101_00000{i}.xlsx").write_text("viejo")
            self.assertIsNotNone(comparadores.respaldar(destino, respaldos, lambda _: None))
            self.assertEqual(len(list(respaldos.glob("resultado_*.xlsx"))), 5)

    def test_cache_directorios(self):
        with tempfile.TemporaryDirectory() as td:
            raiz = Path(td)
            (raiz / "a.txt").write_text("a")
            cache = comparadores.CacheDirectorios()
            self.assertEqual([e.nombre for e in cache.listar(raiz)], ["a.txt"])
            (raiz / "b.txt").write_text("b")
            self.assertEqual([e.nombre for e in cache.listar(raiz)], ["a.txt"])
            cache.limpiar()
            self.assertEqual({e.nombre for e in cache.listar(raiz)}, {"a.txt", "b.txt"})

    def test_busqueda_mdb_descarta_copia_mas_nueva(self):
        entradas = [
            comparadores.Entrada("base.mdb", Path("base.mdb"), False, 1, 10),
            comparadores.Entrada("base - copia.mdb", Path("copia.mdb"), False, 2, 10),
        ]
        import re
        elegido = comparadores.buscar_mdb("carpeta", re.compile("base"), lambda _: entradas)
        self.assertEqual(elegido, Path("base.mdb"))

    def test_subcarpeta_normalizada(self):
        entradas = [comparadores.Entrada("Detalles Diários", Path("d"), True, 1, 0)]
        self.assertEqual(
            comparadores.subcarpeta("raiz", "detalles diarios", lambda _: entradas),
            Path("d"),
        )

    def test_cache_por_archivo_no_recalcula_si_no_cambio(self):
        vueltas = []

        def calcular(ruta):
            vueltas.append(str(ruta))
            return Path(ruta).read_text()

        cache = comparadores.CachePorArchivo(calcular)
        with tempfile.TemporaryDirectory() as tmp:
            archivo = Path(tmp) / "datos.parquet"
            archivo.write_text("uno")
            self.assertEqual(cache(archivo), "uno")
            self.assertEqual(cache(archivo), "uno")
            self.assertEqual(len(vueltas), 1)   # la segunda salio de la cache

    def test_cache_por_archivo_se_entera_si_el_archivo_cambio(self):
        """Reescribir el parquet tiene que dar la respuesta nueva, sin invalidar
        nada a mano: es lo que pasa despues de consolidar un mes."""
        cache = comparadores.CachePorArchivo(lambda ruta: Path(ruta).read_text())
        with tempfile.TemporaryDirectory() as tmp:
            archivo = Path(tmp) / "datos.parquet"
            archivo.write_text("viejo")
            self.assertEqual(cache(archivo), "viejo")
            archivo.write_text("nuevo y mas largo")
            self.assertEqual(cache(archivo), "nuevo y mas largo")

    def test_cache_por_archivo_delega_cuando_no_existe(self):
        cache = comparadores.CachePorArchivo(lambda ruta: Path(ruta).exists())
        with tempfile.TemporaryDirectory() as tmp:
            archivo = Path(tmp) / "no_esta.parquet"
            self.assertFalse(cache(archivo))
            archivo.write_text("ahora si")
            self.assertTrue(cache(archivo))

    def test_tabla_por_nombre(self):
        self.assertEqual(
            comparadores.tabla_por_nombre(["Central_Empresa"], ["central empresa"]),
            "Central_Empresa",
        )


if __name__ == "__main__":
    unittest.main()
