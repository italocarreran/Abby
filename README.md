# Reliquidación SSCC — repositorio de códigos

Scripts en Python del proceso de reliquidación: el revisor, los actualizadores, el
prorrateador y los cargadores de datos.

## Si sos un asistente y entrás por primera vez

Leé **`AGENTS.md`** entero — es corto y trae la regla de expansión, las convenciones
del código y las trampas conocidas. Después:

| Querés… | Leé |
|---|---|
| entender el sistema | `MAPA.md` — un bloque por script, dos páginas |
| conectar código nuevo con el existente | `INTERFACES.md` — firmas sin cuerpos |
| saber en qué hoja/celda/columna vive un dato | `docs/ESTRUCTURA_CASO_RELIQUIDACION.md` (buscá la sección, no lo leas entero) |
| modificar un script | ahí sí, abrí ese `.py` completo — **y solo ese** |

## Estructura

```
AGENTS.md                cómo se trabaja acá (documento vivo)
MAPA.md                  qué hace cada script y de qué depende
INTERFACES.md            generado — firmas, constantes y dependencias
generar_interfaces.py    el generador de INTERFACES.md
docs/                    referencia de dominio
comun/                   lo compartido (ventana, config, log, excel) — por construir
scripts/                 los .py de trabajo
reemplazos_reuc/         ActualizaRemplazos.py y su propio config
```

## Después de tocar cualquier `.py`

```
python generar_interfaces.py
```

Solo biblioteca estándar, Python 3.9 o superior. Subí el `INTERFACES.md` resultante
junto con el `.py`.
