# `Comparadores/` — carpeta hermana de `Revisor_Relq/`

Dos herramientas anuales, parte de la reliquidación pero fuera del circuito
mensual del Revisor:

- `Comparador_Etapas.py` compara los `.mdb` de Definitivo, Rpre y Rdef.
- `Comparador_Tabulado.py` compara los Consolidados Tabulados y reutiliza el
  `rutas.json` del primer comparador.

Los dos localizan la carpeta hermana que contiene `Revisor_Reliquidacion.py`,
usan `../__config__/config.json` e importan el paquete hermano `../__comun__/`.
Estado, rutas, parquet y vistas quedan en `__config__/AAAA/_comparador*`; los JSON
mensuales, en `__config__/AAAA/MM Mes/`. Solo los Excel mensuales y anuales se
guardan en `00_Salidas/`.
