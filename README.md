# Reliquidación SSCC — repositorio de códigos

Scripts en Python del proceso de reliquidación: el revisor, los actualizadores, el
prorrateador y los cargadores de datos.

## Para descargar y usar

**`Revisor Reliquidación/` es la carpeta de trabajo.** Descargás el repositorio
(`Code` → `Download ZIP`), lo descomprimís, y te llevás la carpeta `Revisor Reliquidación/`
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
AGENTS.md                       cómo se trabaja acá (documento vivo)
MAPA.md                         qué hace cada script y de qué depende
INTERFACES.md                   generado — firmas, constantes y dependencias
generar_interfaces.py           el generador de INTERFACES.md
docs/                           referencia de dominio
Revisor Reliquidación/          ← LA CARPETA DE TRABAJO
├── Revisor_Reliquidacion.py   solo, en la raíz — es el que se abre siempre
├── comun/                     lo compartido, con sus pruebas
├── actualizadores/            los 8 que el Revisor lanza por botón
│   ├── Actualiza_datos.py
│   ├── Actualiza_Data_Access.py    el motor que reutilizan Energia y P9
│   ├── Actualiza_Energia.py
│   ├── Actualiza_Cuadro0.py
│   ├── Actualiza_SC_CO.py
│   ├── Actualiza_Access_P9.py
│   ├── Carga_Retiros.py
│   └── Prorratear.py
└── Reemplazos REUC/           el noveno, aparte porque tiene su propio config
    └── ActualizaRemplazos.py
```

Dos cosas que **no son un descuido** y no conviene reordenar a mano:

- `Reemplazos REUC`, con espacio y mayúsculas: el Revisor lo busca así, literal.
- `config.json` vive en `Revisor Reliquidación/`, **compartido** entre el Revisor y los 8 de
  `actualizadores/`. Si movés uno de esos 8 a otra carpeta, tiene que seguir
  apuntando a ese mismo archivo (no crearse el suyo propio), o el traspaso de
  rutas entre el Revisor y los actualizadores se rompe sin avisar.

## Después de tocar cualquier `.py`

```
python generar_interfaces.py          # regenera INTERFACES.md
python "Revisor Reliquidación/comun/test_config.py"   # pruebas del módulo común
```

Solo biblioteca estándar, Python 3.9 o superior. Subí el `INTERFACES.md` resultante
junto con el `.py`.
