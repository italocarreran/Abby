# Reliquidación SSCC — repositorio de códigos

Scripts en Python del proceso de reliquidación: el revisor, los actualizadores, el
prorrateador y los cargadores de datos.

## Para descargar y usar

**`Revisor_Relq/` y `Comparadores/` son las dos carpetas de programas.**
Descargás el repositorio (`Code` → `Download ZIP`), lo descomprimís y conservás
ambas como hermanas. `00_Salidas/`, `__config__/` y `__comun__/` viven al lado. `00_Salidas/`
contiene solamente Excel/resultados; `__config__/` concentra todos los JSON,
estados y datos intermedios; `__comun__/` contiene el código compartido.

Lo que no se sube y aparece solo al usarlo: `00_Salidas/` y `__config__/`.
Ambas usan la estructura `AAAA/MM Mes` cuando corresponde. La segunda guarda
`config.json`, `reemplazos_reuc.json`, los estados, rutas, parquet y vistas.

## Si sos un asistente y entrás — Claude, ChatGPT, el que sea

**Leé `REGLAS.md` primero, entero, siempre — antes que cualquier otra cosa de
esta tabla.** Es obligatorio: dos asistentes distintos pueden estar editando
este repositorio, y esas reglas son lo único que evita que uno pise el
trabajo del otro sin enterarse.

Después, `AGENTS.md` entero — es corto y trae la regla de expansión, las
convenciones del código y las trampas conocidas. Después:

| Querés… | Leé |
|---|---|
| saber qué se hizo en sesiones anteriores y qué quedó pendiente | `BITACORA.md` |
| entender el sistema | `MAPA.md` — un bloque por script, dos páginas |
| conectar código nuevo con el existente | `INTERFACES.md` — firmas sin cuerpos |
| saber en qué hoja/celda/columna vive un dato | `docs/ESTRUCTURA_CASO_RELIQUIDACION.md` (buscá la sección, no lo leas entero) |
| modificar un script | ahí sí, abrí ese `.py` completo — **y solo ese** |

## Estructura

```
REGLAS.md                       obligatorio — leer primero, siempre
BITACORA.md                     registro de sesiones, qué quedó pendiente
AGENTS.md                       cómo se trabaja acá (documento vivo)
MAPA.md                         qué hace cada script y de qué depende
INTERFACES.md                   generado — firmas, constantes y dependencias
generar_interfaces.py           el generador de INTERFACES.md
docs/                           referencia de dominio
__comun__/                     código común y sus pruebas
__config__/                     configuración/estado/caché local (ignorada por Git)
Comparadores/                   comparadores anuales (hermana del Revisor)
├── Comparador_Etapas.py
└── Comparador_Tabulado.py
00_Salidas/                     salidas compartidas, organizadas por AAAA/MM Mes
Revisor_Relq/                   ← CARPETA DEL REVISOR
├── Revisor_Reliquidacion.py   solo, en la raíz — es el que se abre siempre
├── actualizadores/            los 8 que el Revisor lanza por botón
│   ├── Actualiza_datos.py
│   ├── Actualiza_Data_Access.py    el motor que reutilizan Energia y P9
│   ├── Actualiza_Energia.py
│   ├── Actualiza_Cuadro0.py
│   ├── Actualiza_SC_CO.py
│   ├── Actualiza_Access_P9.py
│   ├── Carga_Retiros.py
│   └── Prorratear.py
└── Reemplazos REUC/           el noveno; usa __config__/reemplazos_reuc.json
    └── ActualizaRemplazos.py
```

Dos cosas que **no son un descuido** y no conviene reordenar a mano:

- `Reemplazos REUC`, con espacio y mayúsculas: el Revisor lo busca así, literal.
- `__config__/config.json` es **compartido** entre el Revisor, los 8
  actualizadores y los comparadores. Ningún JSON se guarda fuera de `__config__/`.

`__config__/` se crea automáticamente al primer guardado y está ignorada por Git,
por lo que no viene en el ZIP. Su estructura refleja la de las salidas: los JSON
mensuales están en `AAAA/MM Mes/`, y cada comparador guarda estado, rutas, parquet
y vistas en `AAAA/_comparador*`. Los Excel finales siguen en `00_Salidas/`.

### Si VS Code subraya `from __comun__ import ...`

Está resuelto y **no hay que tocar el código**. Cada carpeta trae un
`.vscode/settings.json` con `python.analysis.extraPaths`, que es lo único que
hace falta: los scripts agregan `__comun__/` al `sys.path` al arrancar, pero
Pylance no ejecuta nada, solo mira esas rutas. Si el subrayado aparece igual,
es una de estas tres:

1. `__comun__/` no está todavía al lado de `Revisor_Relq/` — bajá el
   repositorio de nuevo, esa carpeta es código y viaja con los programas.
2. Abriste una carpeta que no trae el `.vscode/` (por ejemplo una subcarpeta
   suelta). Abrí la carpeta de trabajo, `Revisor_Relq/` o `Comparadores/`.
3. VS Code quedó con el análisis viejo: `Ctrl+Shift+P` →
   *Developer: Reload Window*.

## Después de tocar cualquier `.py`

```
python generar_interfaces.py          # regenera INTERFACES.md
python __comun__/test_config.py   # pruebas del módulo común
```

Solo biblioteca estándar, Python 3.9 o superior. Subí el `INTERFACES.md` resultante
junto con el `.py`.
