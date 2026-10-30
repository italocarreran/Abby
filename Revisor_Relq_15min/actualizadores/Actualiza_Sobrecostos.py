# =============================================================================
#  Actualiza Cálculo_SobrecostosSSCC  —  VERSION 15 MIN
# =============================================================================
#  Ventana propia del maestro 01 Sobrecostos/Cálculo_SobrecostosSSCC_AAMM_*.xlsm,
#  separada de la de las planillas 3, 5 y 6 (usuaria, 2026-10-30). Dos casillas:
#
#    FD           hojas FD_* desde SSCC_Desempeno (MAPEO_SOBRECOSTOS_FD)
#    SOBRECOSTOS  hoja SOBRECOSTOS!A7:R desde
#                 01 Sobrecostos/SC_CfrecSinFD_AAMM_para_coyailen.csv, con las
#                 formulas de S:EH al largo de lo pegado
#
#  El motor es el mismo de Actualiza_datos.py (ejecutar_actualizacion con la
#  planilla "sc"): aca solo esta la ventana. Corre en un hilo aparte, asi que
#  la ventana no se congela mientras trabaja.
#
#  Se puede correr solo, o recibir del Revisor la ruta de un JSON de traspaso.
# =============================================================================

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

DIR_SCRIPT = Path(__file__).resolve().parent


def _morir(titulo, mensaje):
    """Aborta mostrando el motivo en una ventana (lanzado con pythonw no hay
    consola)."""
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(titulo, mensaje)
        r.destroy()
    except Exception:
        pass
    print(f"{titulo}\n\n{mensaje}", file=sys.stderr)
    raise SystemExit(1)


# El motor y las busquedas viven en Actualiza_datos.py, en esta misma carpeta
# (que a su vez carga __comun__/ y ventana_chica.py, y avisa si faltan).
try:
    sys.path.insert(0, str(DIR_SCRIPT))
    import Actualiza_datos as _datos
    from ventana_chica import VentanaChica
except ImportError as e:
    _morir("Falta un archivo",
           "No se pudo cargar Actualiza_datos.py o ventana_chica.py.\n\n"
           "Copia la carpeta Revisor_Relq_15min completa, no los .py sueltos.\n\n"
           f"Detalle: {e}")


# Claves del JSON de traspaso de cada archivo de la ventana.
CLAVES = {"destino": "calculo_sscc_maestro", "sscc": "sscc_desempeno",
          "coyailen": "sc_cfrec_coyailen"}


def main():
    cfg = _datos.leer_config()
    traspaso = _datos.leer_traspaso(sys.argv)
    modo = {"traspaso": traspaso is not None}

    v = VentanaChica("Cálculo_SobrecostosSSCC",
                     "Trae los FD y la hoja SOBRECOSTOS. El Excel tiene que estar "
                     "cerrado.", traspaso, alto=360)
    var_carpeta = tk.StringVar(
        value=(traspaso or {}).get("carpeta_reliq") or cfg.get("carpeta_reliq", ""))
    rutas_v = {k: tk.StringVar(value="[pendiente]") for k in CLAVES}
    # Viniendo del Revisor en blanco (convencion); a mano, FD marcado.
    sel_fd = tk.BooleanVar(value=not modo["traspaso"])
    sel_sob = tk.BooleanVar(value=False)

    fr = v.seccion("Qué actualizar")
    v.casilla(fr, "FD", sel_fd, "hojas FD_*, desde SSCC_Desempeno")
    v.archivo(fr, "", rutas_v["sscc"], sangria=24)
    v.casilla(fr, "SOBRECOSTOS", sel_sob, "A7:R, desde el csv para Coyailen")
    v.archivo(fr, "", rutas_v["coyailen"], sangria=24)
    fd = v.seccion("Destino")
    v.archivo(fd, "Sobrecostos", rutas_v["destino"])

    def refrescar(*_):
        if modo["traspaso"]:
            d = traspaso.get("rutas", {})
            for k, clave in CLAVES.items():
                rutas_v[k].set(d.get(clave) or "[el Revisor no lo mandó]")
            return
        base = var_carpeta.get()
        if not base or not Path(base).is_dir():
            for x in rutas_v.values():
                x.set("[elige la carpeta 02 CASO RELIQUIDACION]")
            return
        reliq = Path(base)
        for k, r, falta in (
                ("destino", _datos.buscar_sobrecostos_xlsm(reliq), "Cálculo_SobrecostosSSCC"),
                ("sscc", _datos.buscar_sscc_desempeno(reliq), "SSCC_Desempeno"),
                ("coyailen", _datos.buscar_sc_cfrec_coyailen(reliq),
                 "SC_CfrecSinFD_*_para_coyailen.csv")):
            rutas_v[k].set(str(r) if r else f"[{falta} no encontrado]")

    def sel_carpeta():
        ini = var_carpeta.get()
        r = filedialog.askdirectory(title="Selecciona 02 CASO RELIQUIDACION",
                                    initialdir=ini if ini and Path(ini).is_dir() else "")
        if r:
            var_carpeta.set(r)
            _datos.guardar_config({"carpeta_reliq": r})
            modo["traspaso"] = False
            refrescar()

    def ejecutar():
        ruta = var_carpeta.get()
        if not modo["traspaso"] and (not ruta or not Path(ruta).is_dir()):
            messagebox.showerror("Error", "Elige la carpeta 02 CASO RELIQUIDACION.")
            return
        hacer_fd, hacer_sob = sel_fd.get(), sel_sob.get()
        if not (hacer_fd or hacer_sob):
            messagebox.showwarning("Aviso", "Marca FD, SOBRECOSTOS o los dos.")
            return
        v.limpiar_log()
        v.log(f"FD: {hacer_fd} | SOBRECOSTOS: {hacer_sob}")

        def trabajo():
            return _datos.ejecutar_en_hilo(
                Path(ruta) if ruta else Path("."), "sc", hacer_fd, hacer_sob,
                log_func=v.log, progreso_func=lambda a, t=100, m=None: v.estado(m or ""),
                rutas=traspaso.get("rutas") if modo["traspaso"] else None)

        def fin(r):
            ok, rutas = r
            if ok and rutas:
                v.var_estado.set(f"Listo en {v.var_tiempo.get()}")
                messagebox.showinfo("Listo", "Actualizado (queda abierto en Excel):\n\n"
                                    + "\n".join(Path(x).name for x in rutas))
            elif not ok:
                v.var_estado.set("Terminó con errores — revisa el log")
                messagebox.showerror("Error", "Hubo un error. Revisa el log.")

        v.correr(trabajo, fin, boton=btn, estado="Iniciando…")

    btn = v.boton("EJECUTAR", ejecutar, principal=True)
    v.boton("Carpeta…", sel_carpeta)
    v.boton("Salir", v.root.destroy)
    refrescar()
    v.iniciar()


if __name__ == "__main__":
    main()
