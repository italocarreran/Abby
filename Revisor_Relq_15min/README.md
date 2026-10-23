# `Revisor_Relq/` — la carpeta de trabajo

```
Revisor_Reliquidacion_15min.py ← solo, en la raíz — es el que se abre siempre
actualizadores/              ← los 3 que el Revisor lanza por botón
Reemplazos REUC/             ← el de reemplazos, aparte
```

El código común vive afuera, en `../__comun__/`. Toda configuración vive también
afuera, en `../__config__/`: `config.json` es compartido por el Revisor, los tres
actualizadores y los comparadores; `reemplazos_reuc.json` pertenece al de reemplazos.

Modo 15 minutos: ya no hay detalles diarios, CCA, 02 Consolidado_Tabulado ni
bases Access, así que tampoco están `Actualiza_Data_Access.py`,
`Actualiza_Energia.py`, `Actualiza_Access_P9.py` ni `Prorratear.py`.

Después de subir o modificar cualquier `.py`, correr `generar_interfaces.py` (en la
raíz del repositorio) y subir también el `INTERFACES.md` resultante.
