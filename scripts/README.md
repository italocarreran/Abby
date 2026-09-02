# `scripts/` — los .py de trabajo

Acá van, junto a su `config.json` compartido (que no se sube: ver `.gitignore`):

```
Revisor_Reliquidacion.py     Actualiza_Data_Access.py    Carga_Retiros.py
Actualiza_datos.py           Actualiza_Energia.py        Prorratear.py
Actualiza_Cuadro0.py         Actualiza_SC_CO.py          Actualiza_Access_P9.py
```

`Actualiza_Energia.py` y `Actualiza_Access_P9.py` importan el motor de
`Actualiza_Data_Access.py`, así que los tres tienen que quedar en la misma carpeta.

Después de subir o modificar cualquiera de estos, correr `generar_interfaces.py` en
la raíz y subir también el `INTERFACES.md` resultante.
