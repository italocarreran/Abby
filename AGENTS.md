# AGENTS.md — cómo se trabaja en este repositorio

> Documento vivo. Si en una sesión se descubre algo que la siguiente necesita saber,
> se agrega acá antes de cerrar.
>
> Léelo completo. Es corto a propósito.

---

## 1. Qué es este repositorio

Los scripts en Python con los que se hace la **reliquidación de SSCC**: un revisor,
varios actualizadores, un prorrateador y cargadores de datos. Son interdependientes
y comparten un `config.json`.

El repositorio no existe para "guardar el código". Existe para que un asistente
pueda entender el sistema **sin leerlo entero**. De ahí sale la única regla que no
se negocia:

## 2. Regla de expansión

```
1. Leer MAPA.md                       (dos páginas: qué hace cada script)
2. Leer de INTERFACES.md SOLO las entradas que hacen falta
3. Abrir completo ÚNICAMENTE el archivo que se va a modificar
```

**Dónde vive el código:** `Revisor_Relq/` es a la vez la carpeta del
repositorio y la carpeta de trabajo del usuario. Lo que está ahí adentro corre tal cual al
descargarlo. Adentro:

- `Revisor_Reliquidacion.py` va **solo**, en la raíz de `Revisor_Relq/`.
- `comun/` es el módulo compartido.
- `actualizadores/` tiene los 8 que el Revisor lanza por botón.
- `Reemplazos REUC/` — con ese nombre exacto, espacio y mayúsculas, porque el
  Revisor lo busca literal — tiene el noveno, con su propio `config.json`.

**`config.json` es compartido** entre el Revisor y los 8 de `actualizadores/`.
Vive en `Revisor_Relq/`, un nivel arriba de ellos. Un script movido a una subcarpeta
nueva **tiene que seguir resolviendo esa ruta compartida**, no la suya propia
(`DIR_SCRIPT.parent / "config.json"`, no `DIR_SCRIPT / "config.json"`) — si no,
cada uno termina con su copia vacía y el traspaso de rutas entre el Revisor y los
actualizadores se rompe en silencio, sin ningún error visible.

**No abrir los archivos vecinos "para tener contexto".** Si de verdad hace falta uno
más, pedirlo explícitamente y decir por qué. Leer de más es exactamente el problema
que este repositorio resuelve.

Para el dominio (dónde vive cada dato: hoja, celda, columna, tabla) está
`docs/ESTRUCTURA_CASO_RELIQUIDACION.md`. Es largo: buscar la sección puntual con
grep o por su encabezado, no leerlo entero.

---

## 3. Cómo se entregan los cambios

**Este repositorio lo editan dos asistentes con acceso de escritura por git:
Claude y ChatGPT.** Ninguno de los dos ve lo que hace el otro en tiempo real.

La lista concreta de qué hacer al empezar, mientras se trabaja y antes de
cerrar está en **`REGLAS.md`** — es obligatoria, no un resumen de esta
sección. Se lee entera, siempre, antes que este archivo. No se duplica acá
para no tener dos copias de la misma lista desincronizándose con el tiempo.

`BITACORA.md` es el registro de qué se hizo en cada sesión y qué quedó
pendiente — lo que un `git log` no cuenta. Toda sesión que cambia algo real
le agrega una entrada, según `REGLAS.md`.

> **Si en algún momento se vuelve al flujo original** (el usuario sube y baja
> archivos por la web, sin que un asistente tenga push directo): entregar
> siempre el archivo completo, nunca un diff ni "reemplazá la línea 240 por
> esto" — el usuario lo descarga y lo corre tal cual, y editar a mano es donde
> se cuelan los errores. Y como el repo puede estar desactualizado respecto de lo
> que el usuario corre en su equipo, conviene preguntar si la versión del repo es
> la vigente antes de modificar un script, y recordarle subir lo que confirme que
> funciona.

---

## 4. Convenciones de los scripts

Estas ya están establecidas en el código existente. **Todo script nuevo las respeta.**

- **Ventana de escritorio (tkinter)** para elegir rutas y archivos, con selectores
  tipo "examinar", indicador visual de si el archivo requerido está o no, y click en
  la ruta para abrir la carpeta.
- **Las rutas se recuerdan entre ejecuciones** en `config.json`, indexado por
  `<host>_<usuario>`. Se escribe de forma **atómica** (`.tmp` + `os.replace`) y **no
  se escribe** si el archivo existe pero no se puede interpretar: mejor perder un
  ajuste que el archivo entero. Solo se agregan o actualizan claves propias, nunca
  se borra nada ajeno.
- **Log de progreso** con timer y barra, visible durante toda la corrida.
- **Tema oscuro optativo.** El piloto vive en los dos comparadores y usa
  `comun/tema.py`; la clave compartida `tema` vale `"claro"` por omisión. Los
  widgets `tk` clásicos se pintan además de los estilos `ttk`, y un fallo del
  tema nunca debe impedir que la herramienta arranque.
- **Los `.xlsm` se modifican preservando las macros** (xlwings / COM). Los destinos
  tienen que estar cerrados antes de correr; al terminar el archivo queda guardado y
  abierto en Excel a propósito.
- **Leer sin abrir Excel cuando solo se lee**: abrir el `.xlsx`/`.xlsm` como ZIP y
  escanear el XML. Levantar Excel por COM en la unidad de red T: cuesta minutos por
  archivo. Si el escaneo falla o el archivo no es OOXML, se cae a xlwings solo.
  `cfg["lectura_rapida"] = False` fuerza el camino viejo. Los archivos que se
  **escriben** siguen yendo por xlwings.
- **Descartar copias de Windows** al buscar por patrón: nombres que terminen en
  `- copia`, `- copia (2)`, `- Copy`, `(2)`, y los temporales `~$`. Una copia más
  nueva le gana al original si no se filtran.
- **Modo SOLO MIRAR**, marcado por defecto, en todo script que escriba en una base
  de datos.
- **Comparar nombres normalizados** (sin tildes, sin espacios ni guiones bajos, en
  mayúsculas) siempre que se comparen centrales o empresas. `El Toro-1`, `EL_TORO-1`
  y `ELTORO-1` son la misma central. Ojo: la ñ no se arregla quitando tildes —
  descomponer `Año` da `ANO` y `Anio` da `ANIO`, hay que tratarlos como iguales.
- **Por COM las fórmulas se escriben en inglés y con coma**
  (`=LET(x,UNIQUE(VSTACK(...)),FILTER(...))`), aunque en pantalla se vean en español
  con punto y coma. Escribirlas en español por esa vía falla.
- **Limpiar antes de escribir.** Reliquidar es reciclar el cuadro de un mes pasado:
  los archivos llegan dimensionados para el mes anterior. Si el mes nuevo trae menos
  filas y no se limpia, quedan las viejas abajo y los totales salen inflados.
- **Los scripts hermanos localizan al Revisor por su archivo.** Un script fuera de
  `Revisor_Relq/` busca entre sus carpetas hermanas la que contiene
  `Revisor_Reliquidacion.py`; no depende del nombre `Revisor_Relq`, que ya cambió.
- **El JSON de traspaso es opcional.** El revisor puede pasar
  `00_Salidas/AAAA/MM Mes/_traspaso_actualizador.json` como único argumento; **sin argumento
  cada script tiene que seguir funcionando solo**, buscando los archivos por su
  cuenta. Es la vía de escape si el revisor no está.

Hay una skill del usuario, `ventana-xlwings-cl`, que tiene el patrón de ventana +
xlwings ya escrito. Usarla al crear un script nuevo en vez de reinventarlo.

---

## 5. Trampas conocidas (ya costaron caro una vez)

| Trampa | Regla |
|---|---|
| `CENTRALES_EMBALSE` está duplicada en `Actualiza_SC_CO.py` y `Revisor_Reliquidacion.py` | Al tocarla hay que cambiarla en **los dos**, con el nombre exacto del origen. V10 caza la desincronización en ambos sentidos, pero se ve como descuadre de monto. |
| Capacidad nueva en `Actualiza_Data_Access.py` | Sumarla a `CAPACIDADES` (hoy son **5**). Si no, el script que la importa se cae con un `TypeError` raro en vez de decir "copiá el archivo actualizado". Peor si la capacidad es un **filtro**: no falla nada y entran datos de más. Hay un `_verificar_capacidades()` que comprueba al arrancar que lo declarado exista de verdad; ya hizo falta dos veces. |
| Columnas que no llevan fórmula | La `Q` de la hoja #3 del cuadro cero y la `AC` de la hoja "SC y CO" están vacías **a propósito**. Nunca hacer un AutoFill corrido que las pise: hay que tratar los bloques por separado. El actualizador y el verificador usan la misma partición. |
| La `D` de la hoja #3 | Va contra la tabla `A:C`, **no** contra `K`. `K` junta las empresas de `A` y de `F`, así que casi siempre es más larga. |
| `Clave Año_Mes` del origen de la planilla 9 | Viene mal (siempre `23xx`). Se pisa con el mes sacado del **nombre de los archivos**. |
| Orden de los pasos de `Actualiza_Cuadro0.py` | Es la cadena de dependencias, no una preferencia. Correrlos por separado deja el libro a medio calcular **sin ninguna señal**. |
| `Actualiza_Cuadro0.py` con cálculo automático | Va en **manual** de punta a punta, con recálculo explícito en los cuatro puntos donde hace falta. `L5`/`M5` son `SUMAR.SI` de columna entera repetidos por empresa. |
| Sufijo de revisión distinto entre archivos | Un `.mdb` en `R01P` con el resto en `R01D` no es error del árbol, pero suele explicar descuadres de monto. |
| Copiar un maestro a su copia | `shutil.copy2`, que conserva la fecha. Con `copy()` a secas el revisor la sigue marcando en amarillo, porque compara por fecha. |
| El `~$` de Excel | No sirve para saber si un libro está abierto: Excel lo deja huérfano cuando se cae. Comprobar que se pueda escribir abriéndolo en `r+b`. |

---

## 6. INTERFACES.md se genera, no se escribe

```
python generar_interfaces.py            # regenera INTERFACES.md
python generar_interfaces.py --check    # sale con 1 si quedó desactualizado
python generar_interfaces.py --esqueleto-mapa   # bloques para los .py que faltan en MAPA.md
python Revisor_Relq/comun/test_config.py   # pruebas del módulo común
```

Solo biblioteca estándar, Python 3.9+. **Correrlo después de cualquier cambio de
firma** y subir el `INTERFACES.md` resultante junto con el `.py`.

El generador está ajustado a cómo documentan estos scripts, así que **conviene
seguir el mismo estilo** para que siga saliendo bien:

- **Encabezado arriba de todo**, docstring o bloque `#`: los dos se usan igual. De
  ahí sale la descripción del archivo (las primeras líneas, hasta el primer
  subtítulo en MAYÚSCULAS).
- **Banners para separar secciones**, en cualquiera de los dos estilos que ya usás:
  la fila sola (`# ===== TÍTULO =====` entre líneas de guiones) o el texto adentro
  (`# ── Título ──────`). El generador los reconoce y los muestra como divisores, no
  como descripción de la función que viene abajo.
- **Comentario justo encima** de una función o constante cuando no hay docstring:
  eso es lo que se usa como descripción.

Además avisa de dos cosas que conviene mirar: constantes definidas en más de un
archivo (la clase de error de `CENTRALES_EMBALSE`) y archivos `.py` que todavía no
tienen bloque en `MAPA.md`.

---

## 7. Decisiones ya tomadas

- **No usar servidores MCP de indexado de código** (`codebase-memory` y similares):
  corren sobre una copia local del repositorio, que el usuario no tiene, y su índice
  queda en una sola máquina y un solo cliente. El usuario trabaja con dos asistentes
  en paralelo y los dos tienen que ver el mismo estado. `MAPA.md` e `INTERFACES.md`
  cumplen la misma función y viven dentro del repositorio.
- **El módulo común se migra por partes, nunca de golpe.** Una pieza a la vez, un
  script a la vez, verificando que sigue corriendo antes de seguir con el próximo.
- **`comun/config.py` ya está**, con 13 pruebas en `comun/test_config.py`. Los
  scripts migrados conservan sus nombres de siempre (`leer_config`,
  `guardar_config`, `_modificar_config`, `escribir_json`, `get_usuario`) como
  envoltorios de dos líneas, así que **ningún punto de llamada cambia**. Migrar un
  script es reemplazar esas cinco funciones por los envoltorios y agregar
  `from comun import config as _cfg`. Nada más.
- **Un script de `actualizadores/` está una carpeta más lejos de `comun/` que el
  Revisor.** `comun/` vive en `Revisor_Relq/`, junto al Revisor, no dentro de
  `actualizadores/`. Python solo agrega al `sys.path` la carpeta del propio
  script, así que `from comun import config` no la encuentra sola: hay que
  agregar la carpeta padre antes,
  `sys.path.insert(0, str(DIR_SCRIPT.parent))`. Ver cómo lo hace
  `actualizadores/Actualiza_SC_CO.py`.
- **Al migrar una pieza, la versión que queda en `comun/` es la más defensiva de
  las que había**, y las diferencias se anotan en el docstring del módulo. No
  "elegir la del archivo más nuevo": hay que mirarlas todas.
- **`docs/ESTRUCTURA_CASO_RELIQUIDACION.md` entra tal cual y es la referencia de
  dominio.** No hay que rehacerlo. Si algo del dominio cambia, se corrige ahí, no en
  una copia.
- **Donde el código y el documento de dominio difieren, manda el código.** Ya hay
  cinco diferencias detectadas, anotadas en `MAPA.md` → "Diferencias con el
  documento de dominio". Al corregir el documento, borrar la fila de esa tabla.

---

## 8. Al empezar y al cerrar cada sesión

El checklist obligatorio — `git log`, `BITACORA.md`, `git pull` al empezar;
`INTERFACES.md`, `MAPA.md`, la entrada en `BITACORA.md`, commit y push al
cerrar — está en **`REGLAS.md`**. Se sigue siempre, sin excepción; no se
repite acá.
