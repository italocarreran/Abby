# `Comparadores/` — carpeta hermana de `Revisor_Relq/`

Dos herramientas anuales, parte de la reliquidación pero fuera del circuito
mensual del Revisor:

- `Comparador_Etapas.py` — consolida los sobrecostos horarios de los `.mdb` de
  las tres etapas (Definitivo, Reliquidación Preliminar, Reliquidación
  Definitiva) para los 12 meses de un año.
- `Comparador_Tabulado.py` — lo mismo sobre los Consolidados Tabulados, con las
  variables (Generación, CV, CMg, USD). Lee el `rutas.json` del otro comparador
  para no pedir dos veces las mismas carpetas.

## ⚠️ Estado: todavía NO están cableados

Los dos archivos están acá **tal cual los entregó el usuario**, sin modificar.
Hoy resuelven sus rutas como si vivieran dentro de `Revisor_Relq/`
(`BASE / "Salidas"`, `BASE / "config.json"`), así que **en esta ubicación no
funcionan**.

Lo que falta está especificado en **`docs/PLAN_comparadores.md`, Tarea 2**:
apuntarlos al `config.json` compartido del Revisor y a `00_Salidas/`, y volver
anuales sus carpetas `_comparador` / `_comparador_tabulado`.

La Tarea 1 de ese mismo plan (`comun/salidas.py` + el Revisor) va primero.
