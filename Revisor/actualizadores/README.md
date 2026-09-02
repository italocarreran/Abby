# `actualizadores/` — los 8 que el Revisor lanza por botón

```
Actualiza_datos.py           Actualiza_Cuadro0.py        Carga_Retiros.py
Actualiza_Data_Access.py     Actualiza_SC_CO.py          Prorratear.py
Actualiza_Energia.py         Actualiza_Access_P9.py
```

`Actualiza_Energia.py` y `Actualiza_Access_P9.py` importan el motor de
`Actualiza_Data_Access.py` por nombre (`import Actualiza_Data_Access as _ADA`), así
que los tres tienen que quedar siempre en esta misma carpeta.

**`config.json` NO está acá.** Es compartido con el Revisor y vive un nivel arriba,
en `scripts/`. Cada script lo resuelve con `DIR_SCRIPT.parent / "config.json"` —
si algún día se mueve un script a otra carpeta, esa línea tiene que ajustarse para
seguir apuntando al mismo archivo.

`Actualiza_SC_CO.py` ya usa el módulo común (`scripts/comun/config.py`) en vez de
tener su propia copia del manejo del config. Los otros 7 todavía tienen la copia
propia — ver `MAPA.md` → "El módulo común".
