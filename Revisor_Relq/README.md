# `Revisor_Relq/` — la carpeta de trabajo

```
Revisor_Reliquidacion.py     ← solo, en la raíz — es el que se abre siempre
actualizadores/              ← los 8 que el Revisor lanza por botón
Reemplazos REUC/             ← el noveno, aparte
```

El código común vive afuera, en `../__comun__/`. Toda configuración vive también
afuera, en `../__config__/`: `config.json` es compartido por el Revisor, los ocho
actualizadores y los comparadores; `reemplazos_reuc.json` pertenece al noveno.

`Actualiza_Energia.py` y `Actualiza_Access_P9.py`, dentro de `actualizadores/`,
importan el motor de `Actualiza_Data_Access.py` — los tres tienen que quedar en la
misma carpeta.

Después de subir o modificar cualquier `.py`, correr `generar_interfaces.py` (en la
raíz del repositorio) y subir también el `INTERFACES.md` resultante.
