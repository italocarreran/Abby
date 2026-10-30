# `actualizadores/` — los que el Revisor lanza por botón

```
Actualiza_datos.py           (planillas 3, 5 y 6: FD, prorrata, SC y CO)
Actualiza_Sobrecostos.py     (Cálculo_SobrecostosSSCC: FD y SOBRECOSTOS)
Actualiza_Cuadro0.py
Actualiza_Pagos_Varios.py    (arma 04 Planilla 9/pagos_varios_AAMM.csv)
Actualiza_Cuadro1.py         (planilla 1: PAGO_EMPRESAS_* de la 4 y RESUMEN de la 8)
Reorganiza_SC_Cfrec.py       (01 Sobrecostos/SC_CfrecSinFD_AAMM.csv -> ..._para_coyailen.csv)
Actualiza_SC_CO.py           (sin botón propio: lo usa Actualiza_datos)
ventana_chica.py             (la ventana común de todos)
```

La hoja "SC y CO" de la planilla 5 se actualiza desde la ventana de
`Actualiza_datos.py` (casillas «Actualizar SC» / «Actualizar CO»), que importa
el motor de `Actualiza_SC_CO.py`: **los dos tienen que quedar en esta misma
carpeta.** `Actualiza_SC_CO.py` se puede seguir abriendo solo.

`Carga_Retiros.py` se sacó el 2026-10-23: el `Retiros_h.parquet` ya no se usa
en 15 minutos.

Modo 15 minutos: `Actualiza_Data_Access.py` (el motor de los Access),
`Actualiza_Energia.py`, `Actualiza_Access_P9.py` y `Prorratear.py` se sacaron,
porque ya no existen las bases Access ni el 02 Consolidado_Tabulado de los que
leían o en los que escribían.

**`config.json` NO está acá.** Es compartido con el Revisor y los comparadores y
vive en `../../__config__/config.json`. Si algún día se mueve un script, su
`CONFIG_PATH` tiene que seguir resolviendo exactamente ese archivo.

Todos usan el módulo común (`../../__comun__/config.py`) para el manejo del
config — ver `MAPA.md` → "El módulo común".
