# -*- coding: utf-8 -*-
"""Pruebas de que la ventana de los comparadores no consulta el disco.

Son de lectura estatica (``ast``) a proposito: importar los comparadores exige
pandas, pyarrow y Access, que no estan en los contenedores donde corre
``scripts/verificar.sh``.

QUE CUIDAN. ``pintar()`` se llama en cada click —marcar un mes, expandirlo,
terminar una etapa— y llego a preguntarle al disco decenas de veces por vuelta:
un ``is_file()`` por archivo contra la carpeta de red, mas el esquema de cada
parquet. En el NAS eso hacia que marcar una casilla tardara lo mismo que la
busqueda completa. La regla que lo evita es simple y facil de romper sin
notarlo: lo que se sabe del disco se calcula en ``instantanea_mes()``, desde el
hilo de fondo, y ``pintar()`` solo mueve widgets leyendo ``self.instant``.

Ojo: se mira SOLO el cuerpo de ``pintar``. ``Comparador_Etapas.pintar_actual``
hace un ``exists()`` local (en ``__config__``, no en la red) y eso esta bien.
"""

import ast
import sys
import unittest
from pathlib import Path

DIR = Path(__file__).resolve().parent
COMPARADORES = ["Comparador_Etapas.py", "Comparador_Tabulado.py"]

# Nombres que significan "esto va al disco". Si aparecen dentro de pintar(),
# volvimos al problema que esta prueba existe para que no vuelva.
PROHIBIDOS = {
    "is_file", "exists", "stat", "iterdir", "scandir", "glob",
    "estado_etapa", "estado_mes", "datos_completos", "vista_completa",
    "huella", "listar", "cargar_estado", "leer_json", "resolver_rutas",
    "instantanea_mes", "refrescar_instantaneas",
}


def arbol(nombre):
    return ast.parse((DIR / nombre).read_text(encoding="utf-8"))


def metodo(nombre_archivo, nombre_metodo):
    for nodo in ast.walk(arbol(nombre_archivo)):
        if isinstance(nodo, ast.ClassDef) and nodo.name == "App":
            for hijo in nodo.body:
                if isinstance(hijo, ast.FunctionDef) and hijo.name == nombre_metodo:
                    return hijo
    return None


def nombres_llamados(nodo):
    usados = set()
    for hijo in ast.walk(nodo):
        if isinstance(hijo, ast.Call):
            f = hijo.func
            if isinstance(f, ast.Attribute):
                usados.add(f.attr)
            elif isinstance(f, ast.Name):
                usados.add(f.id)
    return usados


class TestVentanaSinDisco(unittest.TestCase):
    def test_pintar_no_consulta_el_disco(self):
        for archivo in COMPARADORES:
            pintar = metodo(archivo, "pintar")
            self.assertIsNotNone(pintar, f"{archivo}: no se encontro App.pintar")
            usados = nombres_llamados(pintar) & PROHIBIDOS
            self.assertEqual(
                usados, set(),
                f"{archivo}: pintar() volvio a tocar el disco ({sorted(usados)}). "
                "Eso va en instantanea_mes(), que corre en el hilo de fondo.")

    def test_pintar_lee_la_instantanea(self):
        for archivo in COMPARADORES:
            fuente = ast.dump(metodo(archivo, "pintar"))
            self.assertIn("instant", fuente,
                          f"{archivo}: pintar() ya no lee self.instant")

    def test_la_instantanea_se_calcula_al_buscar(self):
        """refrescar() la deja lista para cada mes que busco; si no, la ventana
        pintaria con datos viejos o vacios."""
        for archivo in COMPARADORES:
            refrescar = metodo(archivo, "refrescar")
            self.assertIn("instantanea_mes", nombres_llamados(refrescar),
                          f"{archivo}: refrescar() no rehace la instantanea")

    def test_lanzar_rehace_la_instantanea_antes_de_pintar(self):
        """Consolidar cambia el estado del mes: si lanzar() no la rehace, los
        colores quedan mostrando lo de antes hasta el proximo refresco."""
        for archivo in COMPARADORES:
            lanzar = metodo(archivo, "lanzar")
            self.assertIn("refrescar_instantaneas", nombres_llamados(lanzar),
                          f"{archivo}: lanzar() no rehace la instantanea")

    def test_no_quedo_nada_del_tema(self):
        """El piloto de tema claro/oscuro se saco: repintar el arbol entero de
        widgets en cada refresco costaba mas de lo que aportaba."""
        for archivo in COMPARADORES:
            fuente = (DIR / archivo).read_text(encoding="utf-8")
            for rastro in ("_tema", "aplicar_tema", "tema_oscuro"):
                self.assertNotIn(rastro, fuente, f"{archivo}: quedo {rastro}")


if __name__ == "__main__":
    unittest.main()
