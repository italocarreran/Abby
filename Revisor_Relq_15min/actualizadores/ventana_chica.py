# =============================================================================
#  Ventana chica comun de los actualizadores  —  VERSION 15 MIN
# =============================================================================
#  La usuaria pidio (2026-10-30) ventanas mas simples, intuitivas y chicas,
#  como la de la planilla 1: los botones justos, con su descripcion justa, sin
#  tantas rutas ni tanta informacion. Esta es esa ventana, armada una sola vez:
#
#      Titulo                                  (una linea de descripcion)
#      [Mes 2603 · desde el Revisor]           (solo si vino del Revisor)
#      ┌ seccion ─────────────────────────────┐
#      │ ☐ Casilla             nota corta     │
#      │   Etiqueta   nombre_del_archivo.xlsx │  <- verde si esta, rojo si no;
#      └──────────────────────────────────────┘     clic abre la carpeta
#      Estado ...                     00:00:12
#      [ log chico ]
#      [ ACCION ]  [Carpeta…]  [Salir]
#
#  Se muestra solo el NOMBRE del archivo, no la ruta; la ruta completa queda en
#  el log al correr. Vive en Revisor_Relq_15min/ (no en __comun__/) porque
#  usa tkinter y es solo de esta version.
# =============================================================================

import queue
import subprocess
import sys
import threading
import time
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

VERDE = "#1e7b34"
ROJO = "#c0392b"
GRIS = "#666666"
AMARILLO = "#fff4c2"
BOTON = "#2d7a2d"
FUENTE = "Segoe UI"


def fmt_tiempo(seg):
    seg = int(seg)
    return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"


def abrir_carpeta(ruta):
    """Abre la carpeta de un archivo (o la carpeta misma), subiendo hasta la
    primera que exista. No hace nada con los textos entre corchetes."""
    if not ruta or str(ruta).startswith("["):
        return
    p = Path(ruta)
    if p.is_file():
        p = p.parent
    while not p.exists() and p != p.parent:
        p = p.parent
    if not p.exists():
        return
    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(p)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(p)])
    else:
        subprocess.Popen(["xdg-open", str(p)])


def texto_archivo(valor, crea=False):
    """(texto, color) para mostrar una ruta: solo el nombre del archivo. Un
    valor entre corchetes es un aviso ("[no encontrado]") y va en rojo. crea:
    es una salida, que puede no existir todavia."""
    v = (valor or "").strip()
    if not v or v.startswith("["):
        return (v.strip("[]") or "—"), ROJO
    p = Path(v)
    try:
        existe = p.exists()
    except OSError:
        existe = False
    if existe:
        return p.name, VERDE
    if crea:
        return f"{p.name}  (se crea)", GRIS
    return f"{p.name}  (no está)", ROJO


class VentanaChica:
    def __init__(self, titulo, descripcion="", traspaso=None, ancho=560, alto=400):
        self.root = tk.Tk()
        self.root.title(titulo)
        self.root.geometry(f"{ancho}x{alto}")
        self.root.minsize(460, 300)
        self._ancho = ancho
        self._archivos = []          # [(var, label, crea)]
        self._cola = queue.Queue()
        self._timer = {"on": False, "t0": 0.0}
        self._hilo = True

        top = tk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=(8, 2))
        tk.Label(top, text=titulo, font=(FUENTE, 10, "bold"),
                 anchor="w").pack(fill="x")
        if descripcion:
            tk.Label(top, text=descripcion, font=(FUENTE, 8), fg=GRIS, anchor="w",
                     justify="left", wraplength=ancho - 30).pack(fill="x")
        if traspaso:
            tk.Label(top, text=f"Mes {traspaso.get('aamm') or '?'} · desde el Revisor",
                     bg=AMARILLO, font=(FUENTE, 8, "bold"),
                     anchor="w").pack(fill="x", pady=(4, 0))

        # abajo primero, para que el log sea lo que se estira
        self._botones = tk.Frame(self.root)
        self._botones.pack(side="bottom", fill="x", padx=10, pady=6)
        self.cuerpo = tk.Frame(self.root)
        self.cuerpo.pack(fill="x", padx=10)

        fe = tk.Frame(self.root)
        fe.pack(fill="x", padx=10, pady=(6, 0))
        self.var_estado = tk.StringVar(value="Listo")
        self.var_tiempo = tk.StringVar(value="00:00:00")
        tk.Label(fe, textvariable=self.var_estado, font=(FUENTE, 8),
                 anchor="w").pack(side="left", fill="x", expand=True)
        tk.Label(fe, textvariable=self.var_tiempo, font=("Consolas", 9, "bold"),
                 fg=BOTON).pack(side="right")

        fl = tk.Frame(self.root)
        fl.pack(fill="both", expand=True, padx=10, pady=(2, 0))
        self.txt = tk.Text(fl, height=6, font=("Consolas", 8), wrap="none")
        sb = tk.Scrollbar(fl, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(fill="both", expand=True)

    # ---------------------------------------------------------------- piezas
    def seccion(self, titulo, parent=None):
        fr = tk.LabelFrame(parent or self.cuerpo, text=titulo, font=(FUENTE, 8, "bold"),
                           padx=6, pady=2)
        fr.pack(fill="x", pady=(4, 0))
        return fr

    def archivo(self, parent, etiqueta, var, crea=False, sangria=0):
        """Una linea: etiqueta gris + nombre del archivo en color. Clic: carpeta.
        Sin etiqueta va debajo de una casilla: usar sangria=24."""
        fr = tk.Frame(parent)
        fr.pack(fill="x", padx=(sangria, 0))
        if etiqueta:
            tk.Label(fr, text=etiqueta, font=(FUENTE, 8), fg=GRIS, width=12,
                     anchor="w").pack(side="left")
        l = tk.Label(fr, font=(FUENTE, 8), anchor="w", cursor="hand2")
        l.pack(side="left", fill="x", expand=True)
        l.bind("<Button-1>", lambda e: abrir_carpeta(var.get()))
        self._archivos.append((var, l, crea))
        var.trace_add("write", lambda *_: self.pintar())
        self.pintar()
        return l

    def casilla(self, parent, texto, var, nota=None):
        fr = tk.Frame(parent)
        fr.pack(fill="x")
        tk.Checkbutton(fr, text=texto, variable=var,
                       font=(FUENTE, 9)).pack(side="left")
        if nota:
            tk.Label(fr, text=nota, font=(FUENTE, 8), fg=GRIS).pack(side="left", padx=4)
        return fr

    def nota(self, parent, texto):
        l = tk.Label(parent, text=texto, font=(FUENTE, 8), fg=GRIS, anchor="w",
                     justify="left", wraplength=self._ancho - 40)
        l.pack(fill="x")
        return l

    def boton(self, texto, comando, principal=False):
        if principal:
            b = tk.Button(self._botones, text=texto, bg=BOTON, fg="white",
                          activebackground="#1f5c1f", activeforeground="white",
                          font=(FUENTE, 9, "bold"), command=comando)
            b.pack(side="left", padx=4, expand=True, fill="x")
        else:
            b = tk.Button(self._botones, text=texto, font=(FUENTE, 8), command=comando)
            b.pack(side="left", padx=4)
        return b

    def pintar(self):
        for var, l, crea in self._archivos:
            t, c = texto_archivo(var.get(), crea)
            l.config(text=t, fg=c)

    # ------------------------------------------------------------------- log
    def log(self, m=""):
        """Seguro desde cualquier hilo."""
        if self._hilo:
            self._cola.put(("log", str(m)))
        else:
            self._escribir(str(m))
            self.root.update_idletasks()

    def estado(self, m):
        if self._hilo:
            self._cola.put(("estado", str(m)))
        else:
            self.var_estado.set(str(m))
            self.root.update_idletasks()

    def progreso(self, a=0, t=100, m=None):
        if m:
            self.estado(m)

    def limpiar_log(self):
        self.txt.delete("1.0", "end")

    def _escribir(self, m):
        self.txt.insert("end", m + "\n")
        self.txt.see("end")

    def _bombear(self):
        try:
            for _ in range(200):
                tipo, val = self._cola.get_nowait()
                if tipo == "log":
                    self._escribir(val)
                elif tipo == "fin":
                    # El fin de un trabajo en hilo llega por la cola: asi
                    # corre en el hilo de la ventana, sin tocar tkinter
                    # desde el hilo del trabajo.
                    terminar, ok, r = val
                    terminar(ok, r)
                else:
                    self.var_estado.set(val)
        except queue.Empty:
            pass
        self.root.after(120, self._bombear)

    def _tick(self):
        if self._timer["on"]:
            self.var_tiempo.set(fmt_tiempo(time.time() - self._timer["t0"]))
            self.root.after(500, self._tick)

    # ----------------------------------------------------------------- correr
    def correr(self, trabajo, fin, boton=None, hilo=True, estado="Procesando..."):
        """Corre trabajo() (en un hilo, o en el principal si hilo=False, que es
        lo que necesitan los que manejan Excel desde ahi) y despues fin(r) en
        el hilo de la ventana. Si trabajo levanta, el error va al log y a una
        ventana, y fin no se llama."""
        if boton is not None:
            boton.config(state="disabled", bg="#aaaaaa")
        self.var_estado.set(estado)
        self._timer.update(on=True, t0=time.time())
        self._tick()

        def terminar(ok, r):
            self._timer["on"] = False
            self.var_tiempo.set(fmt_tiempo(time.time() - self._timer["t0"]))
            if boton is not None:
                boton.config(state="normal", bg=BOTON)
            self.pintar()
            if ok:
                fin(r)
            else:
                self.var_estado.set("Terminó con errores — revisa el log")
                messagebox.showerror("Error", f"Falló.\n\n{r}")

        def envuelto():
            try:
                r, ok = trabajo(), True
            except Exception as e:
                self.log(f"\nERROR: {e}")
                self.log(traceback.format_exc())
                r, ok = e, False
            if hilo:
                self._cola.put(("fin", (terminar, ok, r)))
            else:
                terminar(ok, r)

        self._hilo = hilo
        if hilo:
            threading.Thread(target=envuelto, daemon=True).start()
        else:
            self.root.update_idletasks()
            envuelto()
            self._hilo = True

    def iniciar(self):
        self._bombear()
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        self.root.geometry(f"+{(self.root.winfo_screenwidth() - w) // 2}"
                           f"+{(self.root.winfo_screenheight() - h) // 2}")
        self.root.mainloop()
