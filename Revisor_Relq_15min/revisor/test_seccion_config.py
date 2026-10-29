# -*- coding: utf-8 -*-
"""La version 15 minutos guarda su config dentro de SECCION_CONFIG.

La horaria y la de 15 minutos comparten raiz: el mismo __comun__/ y el mismo
__config__/config.json (y reemplazos_reuc.json). Si un script de esta version
escribiera afuera de su seccion, le pisaria a la horaria la carpeta, el ultimo
mes o las rutas recordadas sin ningun error visible. Y si un script usara otro
nombre de seccion, el Revisor y ese actualizador dejarian de ver lo mismo.

Las funciones se sacan de cada .py con ast y se ejecutan contra un archivo
temporal con el __comun__/config.py real: importar los scripts exige tkinter,
xlwings y pandas, que no estan en los contenedores.

    python Revisor_Relq/revisor/test_seccion_config.py
"""

import ast
import json
import sys
import tempfile
import unittest
from pathlib import Path

REVISOR = Path(__file__).resolve().parents[1]
RAIZ = REVISOR.parent
sys.path.insert(0, str(RAIZ))

from __comun__ import config as _cfg  # noqa: E402

SCRIPTS = [
    REVISOR / "Revisor_Reliquidacion_15min.py",
    REVISOR / "actualizadores" / "Actualiza_datos.py",
    REVISOR / "actualizadores" / "Actualiza_Cuadro0.py",
    REVISOR / "actualizadores" / "Actualiza_SC_CO.py",
    REVISOR / "actualizadores" / "Actualiza_Pagos_Varios.py",
    REVISOR / "actualizadores" / "Actualiza_Cuadro1.py",
    REVISOR / "Reemplazos REUC" / "ActualizaRemplazos.py",
]
FUNCIONES = {"_seccion_config", "leer_config", "_modificar_config",
             "guardar_config", "leer_valores_cfg", "guardar_valores_cfg"}


def cargar(ruta, config_path):
    arbol = ast.parse(ruta.read_text(encoding="utf-8"))
    cuerpo = []
    for n in arbol.body:
        if isinstance(n, ast.FunctionDef) and n.name in FUNCIONES:
            cuerpo.append(n)
        elif (isinstance(n, ast.Assign) and len(n.targets) == 1
              and isinstance(n.targets[0], ast.Name)
              and n.targets[0].id == "SECCION_CONFIG"):
            cuerpo.append(n)
    ns = {"_cfg": _cfg, "CONFIG_PATH": config_path,
          "get_usuario": _cfg.clave_equipo, "json": json}
    exec(compile(ast.Module(body=cuerpo, type_ignores=[]), str(ruta), "exec"), ns)
    return ns


class SeccionConfig(unittest.TestCase):
    def test_todos_usan_la_misma_seccion(self):
        with tempfile.TemporaryDirectory() as tmp:
            valores = {r.name: cargar(r, Path(tmp) / "c.json").get("SECCION_CONFIG")
                       for r in SCRIPTS}
        self.assertEqual(set(valores.values()), {"modo_15min"}, valores)

    def test_nadie_llama_al_config_sin_seccion(self):
        for r in SCRIPTS:
            texto = r.read_text(encoding="utf-8")
            for prohibido in ("_cfg.leer(", "_cfg.guardar("):
                self.assertNotIn(prohibido, texto, f"{r.name}: {prohibido}")

    def test_no_pisa_lo_de_la_horaria(self):
        yo = _cfg.clave_equipo()
        for r in SCRIPTS:
            with tempfile.TemporaryDirectory() as tmp:
                ruta = Path(tmp) / "config.json"
                horaria = {yo: {"carpeta_base": "HORARIA", "ultimo_mes": "2401"},
                           "_valores": {"P5_D7": {"celda": "D7"}},
                           "otro_pc_otro_usuario": {"x": 1}}
                ruta.write_text(json.dumps(horaria), encoding="utf-8")
                ns = cargar(r, ruta)
                # Arranca vacia: no hereda lo de la horaria.
                self.assertEqual(ns["leer_config"](), {}, r.name)
                self.assertTrue(ns["guardar_config"]({"carpeta_base": "QUINCE"}))
                self.assertEqual(ns["leer_config"]()["carpeta_base"], "QUINCE")
                todo = json.loads(ruta.read_text(encoding="utf-8"))
                for k, v in horaria.items():
                    self.assertEqual(todo[k], v, f"{r.name} tocó {k}")
                self.assertEqual(todo["modo_15min"][yo]["carpeta_base"], "QUINCE")

    def test_valores_del_revisor_van_en_la_seccion(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "config.json"
            ruta.write_text(json.dumps({"_valores": {"P5_D7": {"celda": "Z9"}}}),
                            encoding="utf-8")
            ns = cargar(SCRIPTS[0], ruta)
            self.assertEqual(ns["leer_valores_cfg"](), {})
            ns["guardar_valores_cfg"]("P5_D7", {"celda": "D7"})
            self.assertEqual(ns["leer_valores_cfg"](), {"P5_D7": {"celda": "D7"}})
            todo = json.loads(ruta.read_text(encoding="utf-8"))
            self.assertEqual(todo["_valores"], {"P5_D7": {"celda": "Z9"}})

    def test_seccion_rota_no_se_pisa(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "config.json"
            ruta.write_text(json.dumps({"modo_15min": "basura"}), encoding="utf-8")
            ns = cargar(SCRIPTS[1], ruta)
            self.assertEqual(ns["leer_config"](), {})
            self.assertFalse(ns["guardar_config"]({"a": 1}))
            self.assertEqual(json.loads(ruta.read_text(encoding="utf-8")),
                             {"modo_15min": "basura"})


if __name__ == "__main__":
    unittest.main()
