"""Pruebas sin pantalla para comun.tema."""

import unittest
from unittest.mock import patch

from tema import CLARO, aplicar, paleta


class _Root:
    def __init__(self):
        self.config = {}

    def configure(self, **kwargs):
        self.config.update(kwargs)


class _Style:
    def __init__(self, root):
        self.root = root
        self.usado = None
        self.estilos = {}

    def theme_use(self, nombre):
        self.usado = nombre

    def configure(self, nombre, **kwargs):
        self.estilos[nombre] = kwargs


class TemaTests(unittest.TestCase):
    def test_claro_conserva_los_colores_historicos(self):
        colores = paleta("claro")
        self.assertEqual(colores["rojo"], "#c0392b")
        self.assertEqual(colores["amarillo"], "#b8860b")
        self.assertEqual(colores["verde"], "#1e7a1e")
        self.assertEqual(colores["gris"], "#7f8c8d")
        self.assertEqual(colores["enlace"], "#1f3864")
        self.assertEqual(colores, CLARO)

    def test_valor_ausente_o_invalido_es_claro(self):
        self.assertEqual(paleta(None), CLARO)
        self.assertEqual(paleta("cualquier cosa"), CLARO)

    def test_aplicar_con_root_simulado(self):
        root = _Root()
        with patch("tema.ttk.Style", _Style):
            colores = aplicar(root, "oscuro")
        self.assertEqual(root.config["bg"], colores["fondo"])
        self.assertEqual(colores["modo"], "oscuro")


if __name__ == "__main__":
    unittest.main()
