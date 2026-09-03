# -*- coding: utf-8 -*-
"""Tema claro/oscuro compartido para las ventanas tkinter.

No depende de librerias externas.  ``aplicar`` configura ttk y devuelve la
paleta; ``pintar_tk`` completa el trabajo para los widgets tk clasicos, que no
obedecen los estilos ttk.
"""

from tkinter import ttk


CLARO = {
    "modo": "claro",
    "fondo": "SystemButtonFace",
    "panel": "SystemButtonFace",
    "texto": "SystemWindowText",
    "entrada_fondo": "SystemWindow",
    "entrada_texto": "SystemWindowText",
    "seleccion": "SystemHighlight",
    "seleccion_texto": "SystemHighlightText",
    "borde": "SystemButtonShadow",
    "enlace": "#1f3864",
    "rojo": "#c0392b",
    "amarillo": "#b8860b",
    "verde": "#1e7a1e",
    "gris": "#7f8c8d",
    "texto_estado": "white",
}

OSCURO = {
    "modo": "oscuro",
    "fondo": "#1e1f22",
    "panel": "#292b30",
    "texto": "#f2f3f5",
    "entrada_fondo": "#151619",
    "entrada_texto": "#f2f3f5",
    "seleccion": "#3d6fa8",
    "seleccion_texto": "#ffffff",
    "borde": "#555960",
    "enlace": "#8ab4f8",
    "rojo": "#ff7b72",
    "amarillo": "#f2cc60",
    "verde": "#7ee787",
    "gris": "#b1bac4",
    "texto_estado": "#111317",
}


def paleta(modo="oscuro"):
    """Devuelve una copia de la paleta pedida; cualquier otro valor es claro."""
    return dict(OSCURO if str(modo).lower() == "oscuro" else CLARO)


def aplicar(root, modo="oscuro") -> dict:
    """Configura los estilos ttk de la ventana y devuelve la paleta.

    El llamador puede usar el resultado para pintar tambien los widgets ``tk``
    clasicos. El modo claro conserva exactamente los colores historicos.
    """
    colores = paleta(modo)
    root.configure(bg=colores["fondo"])
    estilo = ttk.Style(root)
    if colores["modo"] == "oscuro":
        try:
            estilo.theme_use("clam")
        except Exception:
            pass
    estilo.configure(".", background=colores["fondo"], foreground=colores["texto"])
    estilo.configure("TFrame", background=colores["fondo"])
    estilo.configure("TLabel", background=colores["fondo"], foreground=colores["texto"])
    estilo.configure("TButton", background=colores["panel"], foreground=colores["texto"])
    estilo.configure("TCheckbutton", background=colores["fondo"], foreground=colores["texto"])
    estilo.configure("TLabelframe", background=colores["fondo"])
    estilo.configure("TLabelframe.Label", background=colores["fondo"],
                     foreground=colores["texto"])
    estilo.configure("TEntry", fieldbackground=colores["entrada_fondo"],
                     foreground=colores["entrada_texto"])
    estilo.configure("TProgressbar", background=colores["verde"],
                     troughcolor=colores["panel"])
    return colores


def pintar_tk(widget, colores):
    """Pinta recursivamente widgets tk clasicos con una paleta de ``aplicar``."""
    clase = widget.winfo_class()
    opciones = {}
    semanticos = ("enlace", "rojo", "amarillo", "verde", "gris", "texto_estado")

    def color_semantico(opcion):
        try:
            actual = str(widget.cget(opcion)).lower()
        except Exception:
            return None
        for nombre in semanticos:
            if actual in {str(CLARO[nombre]).lower(), str(OSCURO[nombre]).lower()}:
                return colores[nombre]
        return None

    if clase in {"Frame", "Canvas"}:
        opciones["bg"] = colores["fondo"]
    elif clase in {"Labelframe", "LabelFrame"}:
        opciones.update(bg=colores["fondo"], fg=colores["texto"])
    elif clase in {"Label", "Checkbutton", "Radiobutton"}:
        opciones.update(bg=colores["fondo"], fg=colores["texto"])
        if clase in {"Checkbutton", "Radiobutton"}:
            opciones.update(activebackground=colores["fondo"],
                             activeforeground=colores["texto"],
                             selectcolor=colores["entrada_fondo"])
    elif clase == "Button":
        opciones.update(bg=colores["panel"], fg=colores["texto"],
                        activebackground=colores["borde"],
                        activeforeground=colores["texto"])
    elif clase in {"Entry", "Text", "Listbox", "Spinbox"}:
        opciones.update(bg=colores["entrada_fondo"], fg=colores["entrada_texto"],
                        insertbackground=colores["entrada_texto"],
                        selectbackground=colores["seleccion"],
                        selectforeground=colores["seleccion_texto"])
    elif clase == "Scrollbar":
        opciones.update(bg=colores["panel"], troughcolor=colores["fondo"],
                        activebackground=colores["borde"])
    fg_semantico = color_semantico("fg")
    bg_semantico = color_semantico("bg")
    if fg_semantico is not None:
        opciones["fg"] = fg_semantico
    if bg_semantico is not None:
        opciones["bg"] = bg_semantico
    if opciones:
        widget.configure(**opciones)
    for hijo in widget.winfo_children():
        pintar_tk(hijo, colores)
