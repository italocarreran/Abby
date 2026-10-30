"""Pruebas de ventana_chica.py. Las de texto corren siempre que haya tkinter;
las de la ventana, solo si hay pantalla (en el PC, o con xvfb-run).

    python Revisor_Relq_15min/actualizadores/test_ventana_chica.py
"""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ventana_chica as vc  # noqa: E402


def _hay_pantalla():
    try:
        r = vc.tk.Tk()
        r.destroy()
        return True
    except Exception:
        return False


class Texto(unittest.TestCase):
    def test_solo_el_nombre_y_el_color(self):
        with tempfile.TemporaryDirectory() as d:
            f = Path(d) / "5_REMUNERACIÓN_CRA_2603.xlsx"
            f.write_text("x")
            self.assertEqual(vc.texto_archivo(str(f)), (f.name, vc.VERDE))
            falta = str(Path(d) / "otro.csv")
            self.assertEqual(vc.texto_archivo(falta), ("otro.csv  (no está)", vc.ROJO))
            self.assertEqual(vc.texto_archivo(falta, crea=True),
                             ("otro.csv  (se crea)", vc.GRIS))
        self.assertEqual(vc.texto_archivo("[no encontrado]"), ("no encontrado", vc.ROJO))
        self.assertEqual(vc.texto_archivo(""), ("—", vc.ROJO))
        self.assertEqual(vc.fmt_tiempo(3725), "01:02:05")


@unittest.skipUnless(_hay_pantalla(), "sin pantalla")
class Ventana(unittest.TestCase):
    def setUp(self):
        self.v = vc.VentanaChica("Prueba", "desc", {"aamm": "2603"})
        self.addCleanup(self.v.root.destroy)
        self.v._bombear()          # lo que hace iniciar(), sin el mainloop

    def _esperar(self, cond):
        for _ in range(200):
            self.v.root.update()
            if cond():
                return
            self.v.root.after(10)
        self.fail("no terminó")

    def test_correr_en_hilo_y_en_el_principal(self):
        b = self.v.boton("X", lambda: None, principal=True)
        for hilo in (True, False):
            r = []
            self.v.correr(lambda: (self.v.log("hola"), 7)[1], r.append, boton=b, hilo=hilo)
            self._esperar(lambda: r)
            self.assertEqual(r, [7])
            self.assertEqual(str(b.cget("state")), "normal")
        self._esperar(lambda: "hola" in self.v.txt.get("1.0", "end"))

    def test_error_va_al_log_y_no_llama_fin(self):
        r = []
        with mock.patch.object(vc.messagebox, "showerror") as err:
            self.v.correr(lambda: 1 / 0, r.append)
            self._esperar(lambda: err.called)
        self.assertEqual(r, [])
        self._esperar(lambda: "ZeroDivisionError" in self.v.txt.get("1.0", "end"))

    def test_archivo_se_repinta_al_cambiar(self):
        var = vc.tk.StringVar(master=self.v.root, value="[falta]")
        l = self.v.archivo(self.v.cuerpo, "Algo", var)
        self.assertEqual(l.cget("fg"), vc.ROJO)
        var.set(__file__)
        self.assertEqual((l.cget("text"), l.cget("fg")), (Path(__file__).name, vc.VERDE))


if __name__ == "__main__":
    unittest.main()
