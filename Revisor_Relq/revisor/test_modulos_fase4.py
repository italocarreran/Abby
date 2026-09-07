# -*- coding: utf-8 -*-
"""Pruebas de las fronteras extraídas al completar la Fase 4."""

import ast
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


class WrappersReenvianArgumentos(unittest.TestCase):
    """Los métodos que quedaron en Revisor_Reliquidacion.py delegando en estos
    módulos tienen que REENVIAR sus argumentos, no repetir el valor por omisión.

    Es la clase de error que no falla ni avisa: `_lanzar_actualizador` pasando
    `indice=0` fijo lanzaba siempre el primer script de la fila (el botón
    "Prorratear" de los tres .mdb abría Actualiza_Data_Access), y
    `_armar_traspaso` pasando `nid=None` dejaba el JSON sin `nodo`, que es
    justo lo que el script necesita para saber desde qué fila lo llamaron.
    """

    def _wrappers(self):
        ruta = REVISOR / "Revisor_Reliquidacion.py"
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        for n in ast.walk(arbol):
            if not isinstance(n, ast.FunctionDef):
                continue
            cuerpo = [s for s in n.body
                      if not (isinstance(s, ast.Expr)
                              and isinstance(s.value, ast.Constant))]
            if len(cuerpo) != 1:
                continue
            s = cuerpo[0]
            valor = s.value if isinstance(s, (ast.Return, ast.Expr)) else None
            if not (isinstance(valor, ast.Call)
                    and isinstance(valor.func, ast.Attribute)
                    and isinstance(valor.func.value, ast.Name)
                    and valor.func.value.id in ("_lectores", "_verificaciones",
                                                "_lanzamiento")):
                continue
            yield n, valor

    def test_ningun_wrapper_pierde_un_argumento(self):
        perdidos = {}
        for definicion, llamada in self._wrappers():
            reenviados = set()
            for a in llamada.args:
                if isinstance(a, ast.Name):
                    reenviados.add(a.id)
            for k in llamada.keywords:
                if isinstance(k.value, ast.Name):
                    reenviados.add(k.value.id)
            faltan = [a.arg for a in definicion.args.args
                      if a.arg != "self" and a.arg not in reenviados]
            if faltan:
                perdidos[definicion.name] = faltan
        self.assertEqual(perdidos, {}, f"wrappers que no reenvían: {perdidos}")

    def test_hay_wrappers_que_revisar(self):
        # Si la búsqueda deja de encontrarlos, la prueba de arriba pasa vacía.
        self.assertGreaterEqual(len(list(self._wrappers())), 10)


if __name__ == "__main__":
    unittest.main()
