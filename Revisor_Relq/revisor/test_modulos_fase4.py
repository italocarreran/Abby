# -*- coding: utf-8 -*-
"""Pruebas de las fronteras extraídas al completar la Fase 4."""

from pathlib import Path
import sys
import tempfile
import unittest

RAIZ = Path(__file__).resolve().parents[2]
REVISOR = Path(__file__).resolve().parents[1]
for carpeta in (RAIZ, REVISOR):
    if str(carpeta) not in sys.path:
        sys.path.insert(0, str(carpeta))

from revisor import lanzamiento, lectores, verificaciones


class Variable:
    def __init__(self, valor):
        self.valor = valor

    def get(self):
        return self.valor


class RevisorMinimo:
    def __init__(self):
        self.rutas = {"archivo": Path("origen.xlsx")}
        self.var_base = Variable("C:/caso")


class Fase4Tests(unittest.TestCase):
    def test_configurar_no_pisa_las_funciones_extraidas(self):
        originales = (
            lectores.conexion_mdb,
            lanzamiento.armar_traspaso,
            verificaciones.comprobar,
        )
        dependencias = {
            "conexion_mdb": object(),
            "armar_traspaso": object(),
            "comprobar": object(),
            "UNA_DEPENDENCIA_NUEVA": 42,
        }
        lectores.configurar(dependencias)
        lanzamiento.configurar(dependencias)
        verificaciones.configurar(dependencias)
        self.assertEqual(
            originales,
            (lectores.conexion_mdb,
             lanzamiento.armar_traspaso,
             verificaciones.comprobar),
        )
        self.assertEqual(lanzamiento.UNA_DEPENDENCIA_NUEVA, 42)

    def test_traspaso_conserva_sobre_y_ruta_de_la_fila(self):
        lanzamiento.configurar({
            "CLAVES_TRASPASO": {"archivo": "origen"},
            "TRASPASO_VERSION": 1,
            "_traspaso": type("Contrato", (), {"ORIGEN": "revisor"}),
        })
        dato = lanzamiento.armar_traspaso(
            RevisorMinimo(), "2608", "p3", "archivo")
        self.assertEqual(dato["origen"], "revisor")
        self.assertEqual(dato["version"], 1)
        self.assertEqual(dato["aamm"], "2608")
        self.assertEqual(dato["carpeta_reliq"], "C:/caso")
        self.assertEqual(dato["rutas"], {"origen": "origen.xlsx"})
        self.assertEqual(dato["ruta_nodo"], "origen.xlsx")
        self.assertEqual(dato["clave_nodo"], "origen")

    def test_prueba_directa_de_archivo_escribible(self):
        with tempfile.TemporaryDirectory() as td:
            ruta = Path(td) / "destino.xlsm"
            ruta.write_bytes(b"contenido")
            self.assertFalse(
                lanzamiento.bloqueado_para_escritura(object(), ruta))
            self.assertEqual(ruta.read_bytes(), b"contenido")

    def test_motor_informa_tipo_desconocido(self):
        lineas = []
        resultado = verificaciones.comprobar(
            RevisorMinimo(), {"tipo": "inventado", "desc": "x"}, {}, lineas.append)
        self.assertEqual(resultado["estado"], "SIN DATOS")
        self.assertIn("tipo desconocido", lineas[-1])


if __name__ == "__main__":
    unittest.main()
