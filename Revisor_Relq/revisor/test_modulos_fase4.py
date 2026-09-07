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

    def test_el_punto_de_entrada_resuelve_todos_sus_nombres(self):
        """Ningún nombre que use Revisor_Reliquidacion.py puede haberse quedado
        del otro lado de la mudanza.

        `_suma_rango_openpyxl` se fue a revisor/lectores.py pero
        `leer_valor_excel` se quedó acá y lo usa; como el uso está dentro de un
        try/except, no reventaba: caía al camino de Excel por COM. El número
        salía bien y por eso no se notaba, pero cada rango abría Excel en la T:,
        que es lo que ese camino rápido existe para evitar.
        """
        import builtins
        ruta = REVISOR / "Revisor_Reliquidacion.py"
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        definidos = set()
        for n in ast.walk(arbol):
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                definidos.add(n.name)
            elif isinstance(n, ast.Name) and isinstance(n.ctx, ast.Store):
                definidos.add(n.id)
            elif isinstance(n, ast.arg):
                definidos.add(n.arg)
            elif isinstance(n, (ast.Import, ast.ImportFrom)):
                for a in n.names:
                    definidos.add((a.asname or a.name).split(".")[0])
            elif isinstance(n, ast.ExceptHandler) and n.name:
                definidos.add(n.name)
            elif isinstance(n, ast.Global):
                definidos.update(n.names)
        usados = {n.id for n in ast.walk(arbol)
                  if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load)}
        # Los dunder de módulo (__file__, __name__…) los pone Python.
        sueltos = sorted(n for n in usados - definidos
                         if not hasattr(builtins, n)
                         and not (n.startswith("__") and n.endswith("__")))
        self.assertEqual(sueltos, [],
                         f"nombres que el punto de entrada usa y no tiene: {sueltos}")

    def test_hay_wrappers_que_revisar(self):
        # Si la búsqueda deja de encontrarlos, la prueba de arriba pasa vacía.
        self.assertGreaterEqual(len(list(self._wrappers())), 10)


if __name__ == "__main__":
    unittest.main()
