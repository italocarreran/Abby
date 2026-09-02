<!-- ARCHIVO GENERADO POR generar_interfaces.py — NO EDITAR A MANO -->
<!-- Para regenerarlo: python generar_interfaces.py -->

# INTERFACES — firmas y contratos

Firmas de funciones, clases y constantes de cada `.py` del repositorio, **sin los
cuerpos**. Sirve para conectar código nuevo con el existente sin abrir los archivos
completos.

> **Regla de expansión.** Leer `MAPA.md` → leer acá solo las entradas que hacen falta
> → abrir completo **únicamente** el archivo que se va a modificar. No abrir los
> archivos vecinos "para tener contexto". Si de verdad hace falta uno más, pedirlo
> explícitamente y decir por qué.

Convenciones de esta página:

- Las funciones que empiezan con `_` son internas del archivo.
- Cuando una función no tiene docstring se muestra el comentario `#` que tenga
  justo encima, si lo tiene.
- Los valores de las constantes largas salen resumidos; el valor exacto está en el
  archivo.


## Todavía no hay `.py` en el repositorio

En cuanto se suba el primer script a `scripts/`, `comun/` o `reemplazos_reuc/`, correr:

```
python generar_interfaces.py
```

Mientras tanto, lo que hace cada script está descrito en `MAPA.md`.
