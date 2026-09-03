# `Comparadores/` — carpeta hermana de `Revisor_Relq/`

Dos herramientas anuales, parte de la reliquidación pero fuera del circuito
mensual del Revisor:

- `Comparador_Etapas.py` — consolida los sobrecostos horarios de los `.mdb` de
  las tres etapas (Definitivo, Reliquidación Preliminar, Reliquidación
  Definitiva) para los 12 meses de un año.
- `Comparador_Tabulado.py` — lo mismo sobre los Consolidados Tabulados, con las
  variables (Generación, CV, CMg, USD). Lee el `rutas.json` del otro comparador
  para no pedir dos veces las mismas carpetas.

## Estado: cableados al Revisor

Los dos localizan la carpeta hermana que contiene `Revisor_Reliquidacion.py`,
usan su `config.json` compartido e importan de ahí `comun/salidas.py`. Sus datos
anuales quedan en `00_Salidas/AAAA/_comparador*`; los Excel y JSON mensuales,
en `00_Salidas/AAAA/MM Mes/`.
