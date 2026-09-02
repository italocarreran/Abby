# `scripts/` — la carpeta de trabajo

```
Revisor_Reliquidacion.py     ← solo, en la raíz — es el que se abre siempre
comun/                       ← lo compartido
actualizadores/              ← los 8 que el Revisor lanza por botón
Reemplazos REUC/             ← el noveno, aparte, con su propio config.json
```

`config.json` (que no se sube: ver `.gitignore`) vive **acá**, compartido entre el
Revisor y los 8 de `actualizadores/`.

`Actualiza_Energia.py` y `Actualiza_Access_P9.py`, dentro de `actualizadores/`,
importan el motor de `Actualiza_Data_Access.py` — los tres tienen que quedar en la
misma carpeta.

Después de subir o modificar cualquier `.py`, correr `generar_interfaces.py` (en la
raíz del repositorio) y subir también el `INTERFACES.md` resultante.
