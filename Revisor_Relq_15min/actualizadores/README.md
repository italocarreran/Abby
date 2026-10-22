# `actualizadores/` — los 4 que el Revisor lanza por botón

```
Actualiza_datos.py           Actualiza_Cuadro0.py
Actualiza_SC_CO.py           Carga_Retiros.py
```

Modo 15 minutos: `Actualiza_Data_Access.py` (el motor de los Access),
`Actualiza_Energia.py`, `Actualiza_Access_P9.py` y `Prorratear.py` se sacaron,
porque ya no existen las bases Access ni el 02 Consolidado_Tabulado de los que
leían o en los que escribían.

**`config.json` NO está acá.** Es compartido con el Revisor y los comparadores y
vive en `../../__config__/config.json`. Si algún día se mueve un script, su
`CONFIG_PATH` tiene que seguir resolviendo exactamente ese archivo.

Todos usan el módulo común (`../../__comun__/config.py`) para el manejo del
config — ver `MAPA.md` → "El módulo común".
