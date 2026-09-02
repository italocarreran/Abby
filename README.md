# Reliquidación SSCC — repositorio de códigos

Scripts en Python del proceso de reliquidación: el revisor, los actualizadores, el
prorrateador y los cargadores de datos.

## Para descargar y usar

**`scripts/` es la carpeta de trabajo.** Descargás el repositorio
(`Code` → `Download ZIP`), lo descomprimís, y te llevás la carpeta `scripts/`
entera. Ahí adentro está todo lo que corre, con la estructura que los scripts
esperan.

Lo que no se sube y aparece solo al usarlo: `config.json`, `Salidas/` y
`Reemplazos REUC/Auxiliares/config.json`.

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
scripts/                 ← LA CARPETA DE TRABAJO
├── comun/               lo compartido, con sus pruebas
├── Revisor_Reliquidacion.py
├── Actualiza_*.py, Carga_Retiros.py, Prorratear.py
└── Reemplazos REUC/     su propio config, en Auxiliares/
```

El nombre `Reemplazos REUC`, con espacio y mayúsculas, **no es un descuido**: el
Revisor lo busca así, literal, en `DIR_SCRIPT / "Reemplazos REUC/..."`. Cambiarlo
rompe el botón "Actualizar reemplazos".

## Después de tocar cualquier `.py`

```
python generar_interfaces.py          # regenera INTERFACES.md
python scripts/comun/test_config.py   # pruebas del módulo común
```

Solo biblioteca estándar, Python 3.9 o superior. Subí el `INTERFACES.md` resultante
junto con el `.py`.
