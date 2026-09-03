# Plan: `__comun__/` y `__config__/` afuera de `Revisor_Relq/`

> **Este documento es para Codex.** Se ejecuta Tarea por Tarea, en orden (cada
> una depende de la anterior). Cada Tarea es un commit. Al terminar cada una,
> correr su "Verificación" antes de pasar a la siguiente. Cuando las tres
> estén aplicadas y verificadas, este archivo se borra (mismo ciclo de vida
> que tuvo `docs/PLAN_comparadores.md`) y lo que quede vivo se resume en
> `MAPA.md` / `AGENTS.md` / `BITACORA.md`.

## Por qué

La usuaria vio el warning del editor sobre `sys.path.insert(0, ...)` +
`from comun import ...` en los comparadores y pidió sacar esa costura de
raíz, no parchearla. Instrucción textual (traducida a decisiones concretas
abajo):

> "todo lo que sea común y configuración quede afuera de Revisor_Reliq. Que
> sea una carpeta que se llame `__config__` ahí que se guarden los .json de
> todos los programas [...] no quiero ni un solo .json fuera de esa carpeta
> [...] La carpeta comun de Revisor_relq [...] dejarla afuera como `__comun__`
> no dentro de `__config__` porque `__config__` no la descargo [...] En la
> carpeta `00_Salidas` solo que se guarden resultados, nada más."

Decisiones que se derivan de esto, ya tomadas (no hay que volver a
preguntarlas):

- `__config__/` y `__comun__/` son carpetas **hermanas** de `Revisor_Relq/`,
  `Comparadores/` y `00_Salidas/` — mismo nivel, la "carpeta de trabajo".
- `__config__/` la arma y ordena la usuaria a mano, con la misma estructura
  `AAAA/MM Mes` que ya usa `00_Salidas`. Acá no se construye tooling de
  migración de sus archivos existentes — el código solo tiene que apuntar al
  lugar nuevo.
- `__config__/` **no se sube al repo** (se agrega a `.gitignore` completa).
  `__comun__/` **sí** se sube, es código como `Revisor_Relq/` o
  `Comparadores/`.
- Dentro de `__config__/` hay un `config.json` común (compartido entre
  Revisor, los 8 actualizadores y los 2 comparadores) y JSON por mes/año,
  "como se hace ahora" — mismo esquema de hoy, cambia la raíz nada más.
- `00_Salidas/` termina con **solo** los `.xlsx` finales. Todo lo que hoy es
  estado/caché/traspaso (JSON, `parquet/`, `vistas/`) se muda a `__config__/`.
  Confirmado con la usuaria (eligió la opción recomendada): los `parquet/` y
  `vistas/` de los comparadores van con el estado, a `__config__`.

Árbol final (nivel de trabajo, "2_0" en palabras de la usuaria):

```
__comun__/                       CÓDIGO — se sube al repo
├── __init__.py
├── config.py
├── salidas.py
├── tema.py
├── test_config.py
├── test_salidas.py
└── test_tema.py

__config__/                      DATOS — NO se sube (gitignored entero)
├── config.json                          compartido: Revisor + actualizadores + comparadores
├── reemplazos_reuc/
│   └── config.json                      propio de ActualizaRemplazos.py
└── AAAA/
    ├── MM Mes/
    │   ├── _revisor_verificaciones.json
    │   ├── _revisor_cache_valores.json
    │   └── _traspaso_actualizador.json
    ├── _comparador/                     Comparador_Etapas
    │   ├── estado.json
    │   ├── rutas.json
    │   ├── parquet/
    │   ├── vistas/
    │   └── actual/
    └── _comparador_tabulado/            Comparador_Tabulado
        ├── estado.json
        ├── rutas.json
        ├── parquet_variables/
        └── vistas_variables/

00_Salidas/                      RESULTADOS — se sube la estructura, no el contenido (ya gitignorado)
└── AAAA/
    ├── MM Mes/
    │   ├── Comparacion_AAMM.xlsx            Comparador_Etapas, mensual
    │   └── Comparacion_Variables_AAMM.xlsx  Comparador_Tabulado, mensual
    ├── Comparacion_Etapas_AAAA.xlsx         Comparador_Etapas, anual
    └── Comparacion_Variables_AAAA.xlsx      Comparador_Tabulado, anual

Revisor_Relq/                    (sin comun/ adentro — se fue a __comun__/)
├── Revisor_Reliquidacion.py
├── actualizadores/
└── Reemplazos REUC/
    └── Auxiliares/              sigue acá: datos_reuc_*.xlsx, "Reemplazos forzados.xlsx"
                                  (NO son .json — solo su config.json se mudó)

Comparadores/
├── Comparador_Etapas.py
└── Comparador_Tabulado.py
```

Nota al pasar (no es una tarea, la usuaria solo avisó "por si"): el
`__pycache__/` que aparece en la carpeta de trabajo ya está cubierto por la
regla `__pycache__/` que existe en `.gitignore` desde siempre — Git la aplica
en cualquier profundidad, así que no hace falta tocar nada por eso.

---

## Tarea 1 — Mover `comun/` a `__comun__/`, retirar `_hallar_revisor`

Solo mueve **código e imports**. Ningún `CONFIG_PATH` ni ruta de `.json`
cambia todavía en esta tarea — eso es la Tarea 2. Así queda cada tarea chica
y verificable por separado.

### 1.1 — Mover el módulo

`git mv Revisor_Relq/comun __comun__` (mueve `__init__.py`, `config.py`,
`salidas.py`, `tema.py`, `test_config.py`, `test_salidas.py`, `test_tema.py`
tal cual están, sin tocar su contenido).

### 1.2 — Borrar la carpeta huérfana `comun/` de la raíz

Existe hoy un `comun/README.md` en la raíz del repo (al lado de `AGENTS.md`),
un placeholder de una etapa anterior del proyecto, nunca usado por ningún
script (el `comun/` real siempre vivió dentro de `Revisor_Relq/`). Borrar esa
carpeta entera (`comun/README.md` incluido) — es basura vieja, no una
carpeta en uso, y confunde tener dos `comun` en el árbol.

### 1.3 — Bootstrap nuevo, igual en los seis puntos de entrada

Hoy `_hallar_revisor()` busca una carpeta hermana que tenga
`Revisor_Reliquidacion.py` (duplicado en los dos comparadores). Ya no sirve:
después de este cambio, los comparadores no necesitan encontrar
`Revisor_Relq/` para nada — lo único que necesitan encontrar es `__comun__/`.
Y como el nombre de `Revisor_Relq/` ya cambió tres veces, conviene que la
búsqueda sea por la marca (`__comun__/` al lado), no por una cuenta fija de
niveles `.parent.parent`.

Reemplaza `_hallar_revisor()` por esta función, **copiada igual** (mismo
cuerpo, se explica una sola vez acá) en cada uno de estos seis archivos:
`Revisor_Relq/Revisor_Reliquidacion.py`,
`Revisor_Relq/Reemplazos REUC/ActualizaRemplazos.py`,
`Comparadores/Comparador_Etapas.py`, `Comparadores/Comparador_Tabulado.py`,
y en los dos actualizadores que hoy importan `comun`
(`Revisor_Relq/actualizadores/Actualiza_SC_CO.py`; los otros 7 actualizadores
no importan `comun` todavía — no les toca esta función en esta tarea, ver
nota al final de 1.3):

```python
def _hallar_workroot(desde):
    """Sube desde `desde` hasta encontrar la carpeta que tiene __comun__ al
    lado. No asume nombres de carpetas intermedias (Revisor_Relq ya se
    renombro varias veces) ni cuantos niveles de profundidad hay: busca la
    marca, no una ruta fija."""
    actual = Path(desde).resolve()
    for candidata in (actual, *actual.parents):
        if (candidata / "__comun__").is_dir():
            return candidata
    return None
```

Se define ANTES de cualquier `import __comun__`, porque hasta no encontrar
`WORKROOT` no se puede armar el `sys.path` para importarlo — mismo motivo por
el que `_hallar_revisor` estaba donde estaba.

Uso, en cada uno de los seis archivos (`_morir` ya existe en todos, no se
toca):

```python
WORKROOT = _hallar_workroot(DIR_SCRIPT)   # o BASE, segun como se llame ahi
if WORKROOT is None:
    _morir(
        "No se encontro __comun__",
        "Este script tiene que estar dentro de la carpeta de trabajo, la que "
        "tiene __comun__/ como hermana (junto con Revisor_Relq/, "
        "Comparadores/ y 00_Salidas/).\n\n"
        f"Se busco subiendo desde: {DIR_SCRIPT}",
    )

sys.path.insert(0, str(WORKROOT))
from __comun__ import salidas as _sal      # + tema, config, lo que ya importe cada uno
```

`from __comun__ import ...` es una importación normal — los guiones bajos
dobles al principio y al final de un nombre de paquete no activan el
"name mangling" de Python (eso solo aplica a atributos con un solo guion bajo
al final, dentro de una clase). No hace falta nada especial.

Detalle por archivo:

- **`Revisor_Reliquidacion.py`**: hoy hace `from comun import salidas as
  _sal` directo (línea 28), sin `sys.path.insert`, porque el script vive
  adentro de `Revisor_Relq/` al lado de `comun/` y Python agrega la carpeta
  del script a `sys.path` solo. Ahora `__comun__` ya no es hermano del
  script, así que sí necesita el bootstrap completo: `DIR_SCRIPT =
  Path(__file__).resolve().parent` ya existe (línea 1100, más abajo en el
  archivo — subir esa línea o reutilizarla antes del import, lo que quede
  más prolijo) — buscar `WORKROOT` con `_hallar_workroot(DIR_SCRIPT)`,
  insertar en `sys.path`, recién ahí `from __comun__ import salidas as _sal`.
  Ese script no tiene `_morir()` como los demás: agregar una función
  equivalente ahí (mismo patrón, ventana de error + `raise SystemExit(1)`) o,
  si ya existe algo parecido, reusarlo — revisar antes de duplicar.
- **`Comparador_Etapas.py` / `Comparador_Tabulado.py`**: reemplazan el bloque
  actual (`_hallar_revisor`, `DIR_REVISOR`, el `if DIR_REVISOR is None:
  _morir(...)`, el `sys.path.insert(0, str(DIR_REVISOR))`) por el bloque de
  arriba con `WORKROOT`. `BASE` (que ya existe, `= Path(__file__).resolve()
  .parent`) hace de `desde`. Ya no queda ninguna referencia a
  `Revisor_Relq` ni a `Revisor_Reliquidacion.py` en estos dos scripts.
- **`ActualizaRemplazos.py`**: hoy no importa `comun` para nada (tiene su
  propio manejo de config inline). No hace falta agregarle el bootstrap en
  esta tarea si no va a usar `__comun__` — pero si en la Tarea 2 se decide
  que reutilice `comun/salidas.py` para su propia carpeta dentro de
  `__config__/`, ahí sí le hace falta. Ver Tarea 2.2.
- **`Actualiza_SC_CO.py`**: hoy hace `sys.path.insert(0,
  str(DIR_SCRIPT.parent))` seguido de `from comun import config as _cfg`
  (asume que `comun` es hermano de `Revisor_Relq/`, un nivel arriba de
  `actualizadores/`). Cambia a `_hallar_workroot(DIR_SCRIPT)` (que ya
  sube los niveles que hagan falta solo) + `from __comun__ import config as
  _cfg`.
- **Los otros 7 actualizadores** (`Actualiza_datos.py`,
  `Actualiza_Data_Access.py`, `Actualiza_Energia.py`, `Actualiza_Cuadro0.py`,
  `Actualiza_Access_P9.py`, `Carga_Retiros.py`, `Prorratear.py`): no importan
  `comun` hoy (tienen su propio código de config inline, migración pendiente
  y ya documentada aparte en `BITACORA.md`). No les toca nada en esta tarea.

### 1.4 — `generar_interfaces.py`

`CARPETAS` tiene hoy `["Revisor_Relq/comun", "Revisor_Relq"]` (falta
`"Comparadores"` — confirmar si ya está, si no agregarlo también).
Cambiarla a `["__comun__", "Revisor_Relq", "Comparadores"]`. Regenerar
`INTERFACES.md` (`python generar_interfaces.py`) y confirmar con
`python generar_interfaces.py --check` que queda limpio.

### Verificación de la Tarea 1

- `python __comun__/test_config.py`, `python __comun__/test_salidas.py`,
  `python __comun__/test_tema.py` — los tres pasan igual que antes de mover
  la carpeta (son pruebas de comportamiento puro, no dependen de dónde vive
  el módulo).
- `python generar_interfaces.py --check` no marca diferencias.
- Los seis scripts arrancan sin `ModuleNotFoundError` ni
  `ImportError` — probar levantando cada uno (puede ser solo hasta que
  aparezca la ventana, no hace falta operarlos a fondo en esta tarea).
- Ningún `.py` del repo importa `comun` (con ese nombre) en ningún lado:
  `grep -rn "from comun import\|import comun" --include=*.py .` no devuelve
  nada.
- No quedan referencias a `_hallar_revisor` ni a `DIR_REVISOR` en ningún
  archivo.
- No queda ninguna carpeta `comun/` en el repo (ni en la raíz ni dentro de
  `Revisor_Relq/`) — solo `__comun__/`.

---

## Tarea 2 — `__config__/`: mover ahí todos los `.json`, dejar `00_Salidas/` solo con resultados

Depende de la Tarea 1 (usa `__comun__` ya movido). Esta es la tarea grande:
toca `CONFIG_PATH` en los 10 scripts que lo tienen, más las rutas de estado
de los 2 comparadores.

### 2.1 — `Revisor_Reliquidacion.py`

Reemplazar (alrededor de la línea 1100-1103):

```python
DIR_SCRIPT = Path(__file__).resolve().parent
CONFIG_PATH = DIR_SCRIPT / "config.json"
DIR_SALIDAS = DIR_SCRIPT.parent / "00_Salidas"
ARCHIVO_ESTADO = "_revisor_verificaciones.json"
```

por (usando el `WORKROOT` que ya se calculó en la Tarea 1.3, cerca del
`import` de `__comun__`, más arriba en el archivo):

```python
DIR_SCRIPT = Path(__file__).resolve().parent
DIR_SALIDAS = WORKROOT / "00_Salidas"      # resultados finales — sin cambios de fondo
DIR_CONFIG = WORKROOT / "__config__"       # config y estado — nuevo
CONFIG_PATH = DIR_CONFIG / "config.json"
ARCHIVO_ESTADO = "_revisor_verificaciones.json"
```

`ARCHIVO_CACHE` (línea 1578, `_revisor_cache_valores.json`) y
`ARCHIVO_TRASPASO` (línea 1081, `_traspaso_actualizador.json`) no cambian de
nombre, solo de raíz.

`dir_mes(aamm, crear=False)` (línea 1106) sigue existiendo tal cual, sigue
usando `_sal.carpeta_mes(DIR_SALIDAS, aamm, crear=crear)` — es la carpeta de
**resultados** del mes (el botón "abrir en el Explorador" y, después de la
Tarea 2.4, donde los comparadores dejan sus `.xlsx` mensuales). Agregar una
función hermana, misma forma, para la carpeta de **config** del mes:

```python
def dir_mes_config(aamm, crear=False):
    """__config__/AAAA/MM Mes, hermana de Revisor_Relq. Mismo AAAA/MM Mes que
    dir_mes(), pero para el JSON de estado/cache/traspaso, no para resultados.
    """
    return _sal.carpeta_mes(DIR_CONFIG, aamm, crear=crear)
```

Todos los usos de `dir_mes(aamm) / ARCHIVO_ESTADO`,
`dir_mes(aamm) / ARCHIVO_CACHE` y `dir_mes(aamm) / ARCHIVO_TRASPASO` (o
`carpeta_salida / ARCHIVO_TRASPASO` en la línea ~4050, donde `carpeta_salida
= dir_mes(aamm or "sin_mes", crear=True)`) cambian de `dir_mes(...)` a
`dir_mes_config(...)`. Revisar con
`grep -n "ARCHIVO_ESTADO\|ARCHIVO_CACHE\|ARCHIVO_TRASPASO" Revisor_Relq/Revisor_Reliquidacion.py`
— son las líneas 1528, 1545-1546, 1627, 1639-1640, 1696, 3750, 4049-4050,
4581-4582, 4929, 4935, 4941 (los números pueden correrse un poco por los
cambios de 1.3/2.1, pero el patrón es ese). Los usos de `dir_mes(...)` que
**no** están seguidos de `/ ARCHIVO_*` (el botón de "abrir en el Explorador",
líneas ~3442-3443) se quedan como están — apuntan a resultados, no a config.

### 2.2 — `ActualizaRemplazos.py`

Reemplazar (línea 54):

```python
CARPETA_AUXILIARES = Path(__file__).parent / "Auxiliares"
CONFIG_PATH = CARPETA_AUXILIARES / "config.json"
```

por:

```python
DIR_SCRIPT = Path(__file__).resolve().parent
CARPETA_AUXILIARES = DIR_SCRIPT / "Auxiliares"   # sigue igual: datos_reuc_*, Reemplazos forzados.xlsx
WORKROOT = _hallar_workroot(DIR_SCRIPT)
if WORKROOT is None:
    _morir("No se encontro __comun__", "...")    # mismo mensaje que en 1.3
CONFIG_PATH = WORKROOT / "__config__" / "reemplazos_reuc" / "config.json"
```

`CARPETA_AUXILIARES` **no se mueve** — sigue al lado del script, guardando
los `.xlsx` descargados de la web REUC y "Reemplazos forzados.xlsx". La
única pieza de esa carpeta que era un `.json` (`config.json`) es la que se
va a `__config__/reemplazos_reuc/`. Este script no importaba `comun` antes;
ahora necesita `_hallar_workroot` (copiar la función de 1.3, este archivo
quedó afuera de esa tarea porque entonces no la necesitaba) — no necesita
importar nada más de `__comun__`, ya que tiene su propio manejo de config
inline (fuera del alcance de este plan; ver pendiente ya anotado en
`BITACORA.md` sobre migrar esto a `comun/config.py` más adelante).

### 2.3 — Los 8 actualizadores (`CONFIG_PATH`)

En los 8 archivos de `Revisor_Relq/actualizadores/`, `CONFIG_PATH` hoy es
`DIR_SCRIPT.parent / "config.json"` (o
`Path(__file__).resolve().parent.parent / "config.json"` en
`Actualiza_datos.py` y `Actualiza_Data_Access.py`, mismo resultado escrito
distinto). Cambia a apuntar a `__config__/config.json`, que está dos
niveles arriba de `actualizadores/` (`actualizadores/` → `Revisor_Relq/` →
carpeta de trabajo → `__config__/`):

```python
DIR_SCRIPT = Path(__file__).resolve().parent
CONFIG_PATH = DIR_SCRIPT.parent.parent / "__config__" / "config.json"
```

Esto es un cálculo directo por niveles fijos (no busca con
`_hallar_workroot`) porque estos 8 scripts siempre viven en
`Revisor_Relq/actualizadores/`, a profundidad fija — igual que hoy usan
`.parent.parent` sin buscar nada. Si más adelante alguno de estos 8 pasa a
importar `__comun__` (parte del pendiente ya anotado en `BITACORA.md` de
migrar los 7 que faltan a `comun/config.py`), ahí conviene que también sume
el bootstrap de `_hallar_workroot` y calcule `CONFIG_PATH` desde `WORKROOT`
en vez de contar niveles a mano — no es parte de esta tarea, pero dejarlo
anotado para quien la haga.

Los comentarios que hoy explican por qué `CONFIG_PATH` no es
`DIR_SCRIPT / "config.json"` (ej. "config.json es compartido con el Revisor
y el resto de los actualizadores [...] no es DIR_SCRIPT / "config.json"
porque este script esta en actualizadores/") hay que actualizarlos: ya no
dicen "vive un nivel arriba, junto al Revisor" sino "vive en `__config__/`,
hermana de `Revisor_Relq/`".

Archivos: `Actualiza_datos.py`, `Actualiza_Data_Access.py`,
`Actualiza_Energia.py`, `Actualiza_Cuadro0.py`, `Actualiza_SC_CO.py`,
`Actualiza_Access_P9.py`, `Carga_Retiros.py`, `Prorratear.py`.

### 2.4 — Los 2 comparadores

Reemplazar, en ambos archivos, el tramo que hoy tiene
`CONFIG_PATH = DIR_REVISOR / "config.json"` y
`SALIDAS = _sal.raiz_salidas(BASE)` por:

```python
CONFIG_PATH = WORKROOT / "__config__" / "config.json"
SALIDAS = WORKROOT / "00_Salidas"
CONFIG = WORKROOT / "__config__"
```

(`WORKROOT` ya quedó definido en la Tarea 1.3; `_sal.raiz_salidas()` deja de
usarse acá — ver 2.6 sobre borrarla de `salidas.py` si nada más la usa).

**`Comparador_Etapas.py`** — separar `cdir(anio)` (que hoy mezcla estado y
resultado en la misma carpeta `00_Salidas/AAAA/_comparador/`) en dos raíces
distintas:

```python
def cdir(anio):
    """__config__/AAAA/_comparador — estado, cache y datos intermedios."""
    return _sal.carpeta_comparador(CONFIG, anio, "_comparador")


def dir_parquet(anio):
    return cdir(anio) / "parquet"


def dir_sob_raiz(anio):
    return dir_parquet(anio) / "sobrecostos"


def dir_cen_raiz(anio):
    return dir_parquet(anio) / "centrales"


def dir_vistas(anio):
    return cdir(anio) / "vistas"


def dir_actual(anio):
    return cdir(anio) / "actual"


def actual_parquet(anio):
    return dir_actual(anio) / "central_empresa_actual.parquet"


def estado_path(anio):
    return cdir(anio) / "estado.json"


def rutas_path(anio):
    return cdir(anio) / "rutas.json"


def xlsx_anual(anio):
    """00_Salidas/AAAA/Comparacion_Etapas_AAAA.xlsx — el resultado, directo
    bajo el año, sin la subcarpeta _comparador (esa quedo en __config__)."""
    return SALIDAS / _sal.normalizar_anio(anio) / f"Comparacion_Etapas_{_sal.normalizar_anio(anio)}.xlsx"
```

Buscar dónde se arma hoy el nombre del Excel anual (línea ~1111, algo como
`cdir(anio_aa) / f"Comparacion_Etapas_{...}.xlsx"`) y usar `xlsx_anual(anio)`
en su lugar. El Excel **mensual** (`Comparacion_{aamm}.xlsx`, línea ~1106,
`_sal.carpeta_mes(SALIDAS, aamm) / f"Comparacion_{aamm}.xlsx"`) no cambia —
ya apunta a `SALIDAS` (resultados), sigue igual.

`NOMBRE_JSON_MES = "_traspaso_actualizador.json"` (línea 232) — donde hoy se
lee con `_sal.carpeta_mes(SALIDAS, aamm) / NOMBRE_JSON_MES` (línea 498),
cambiar `SALIDAS` por `CONFIG`: ese archivo lo escribe el Revisor en
`__config__/AAAA/MM Mes/` después de la Tarea 2.1, así que el comparador
tiene que leerlo del mismo lugar. **Esto tiene que aplicarse junto con la
2.1 y verificarse junto** (ver más abajo, "Trampa" — es exactamente el tipo
de bug que ya está anotado en `AGENTS.md` → Trampas conocidas sobre armar
rutas por separado).

**`Comparador_Tabulado.py`** — mismo patrón:

```python
def cdir(anio):
    """__config__/AAAA/_comparador_tabulado."""
    return _sal.carpeta_comparador(CONFIG, anio, "_comparador_tabulado")


def dir_parquet(anio):
    return cdir(anio) / "parquet_variables"


def dir_vistas(anio):
    return cdir(anio) / "vistas_variables"


def estado_path(anio):
    return cdir(anio) / "estado.json"


def rutas_path(anio):
    return cdir(anio) / "rutas.json"


def cdir_mdb(anio):
    """__config__/AAAA/_comparador — carpeta del OTRO comparador, se consulta
    en modo solo lectura para reusar sus rutas.json de los .mdb."""
    return _sal.carpeta_comparador(CONFIG, anio, "_comparador")


def rutas_mdb_path(anio):
    return cdir_mdb(anio) / "rutas.json"


def xlsx_anual(anio):
    return SALIDAS / _sal.normalizar_anio(anio) / f"Comparacion_Variables_{_sal.normalizar_anio(anio)}.xlsx"
```

Mismo cambio de `SALIDAS` a `CONFIG` en la lectura de `NOMBRE_JSON_MES`
(línea ~631) y en el Excel anual (línea ~931, hoy
`cdir(aa) / f"Comparacion_Variables_{...}.xlsx"` → `xlsx_anual(aa)`). El
Excel mensual (línea ~927, `_sal.carpeta_mes(SALIDAS, aamm) / ...`) no
cambia.

### 2.5 — `Comparadores/README.md`

Actualizar el bloque de comentarios/README que hoy describe
`00_Salidas/AAAA/_comparador/estado.json` etc. (líneas ~22-23 de
`Comparador_Etapas.py` y ~35-36 de `Comparador_Tabulado.py`, más
`Comparadores/README.md` si repite lo mismo) para que diga
`__config__/AAAA/_comparador/` en vez de `00_Salidas/AAAA/_comparador/`.

### 2.6 — Limpieza en `__comun__/salidas.py`

Después de 2.4, `raiz_salidas()` queda sin ningún llamador (los dos
comparadores pasan a calcular `SALIDAS` y `CONFIG` directo desde `WORKROOT`,
que ya tienen). Confirmar con
`grep -rn "raiz_salidas" --include=*.py .` que efectivamente no queda nadie
usándola, y si es así, borrarla de `salidas.py` (y su prueba en
`test_salidas.py`) — código muerto, no se deja "por si acaso".

El resto de `salidas.py` (`carpeta_mes`, `carpeta_comparador`,
`carpetas_legado`, `partir_aamm`, `nombre_carpeta_mes`, `normalizar_anio`,
`carpeta_anio`) **no cambia**: ya reciben la raíz como parámetro, por eso
sirven igual para `00_Salidas` y para `__config__` sin tocar una línea de
lógica — es la razón por la que este plan no necesita un módulo nuevo para
`__config__`.

### 2.7 — `.gitignore`

Reemplazar:

```
config.json
00_Salidas/
```

por:

```
__config__/
00_Salidas/
```

(`config.json` como regla suelta ya no hace falta — todos los `config.json`
del repo, sin excepción, viven ahora dentro de `__config__/`, que se ignora
entera). El resto del archivo no cambia.

### Verificación de la Tarea 2

- `grep -rn "\.json" --include=*.py . | grep -iE "CONFIG_PATH ="` — las 10
  líneas (Revisor + 8 actualizadores + ActualizaRemplazos, ojo que son 10
  con el propio Revisor) apuntan todas dentro de `__config__/`.
- `grep -rn "00_Salidas" --include=*.py Comparadores` no debería devolver
  ninguna línea relacionada con `estado.json`, `rutas.json`, `parquet`,
  `vistas`, `_traspaso_actualizador.json` — solo con los `.xlsx`.
- **Prueba cruzada, la que importa de verdad**: armar una carpeta temporal
  con la estructura completa (`__comun__/`, `__config__/`, `00_Salidas/`,
  `Revisor_Relq/`, `Comparadores/`), simular que el Revisor escribe un
  `_traspaso_actualizador.json` para un AAMM cualquiera vía
  `dir_mes_config(aamm, crear=True) / ARCHIVO_TRASPASO`, y confirmar que
  `Comparador_Etapas.py` / `Comparador_Tabulado.py` lo encuentran en
  exactamente esa misma ruta (`CONFIG` + `carpeta_mes` con el mismo AAMM).
  Es la misma clase de prueba que ya se hizo para verificar la Tarea 1 de
  `PLAN_comparadores.md` — no alcanza con leer el código, hay que ejecutarlo.
- Ningún `.py` del repo, después de esta tarea, escribe ni lee un `.json`
  fuera de `__config__/` — repasar cada `CONFIG_PATH`, `ARCHIVO_ESTADO`,
  `ARCHIVO_CACHE`, `ARCHIVO_TRASPASO`, `estado_path`, `rutas_path`,
  `rutas_mdb_path` uno por uno contra este criterio.
- `00_Salidas/`, después de correr el Revisor y los dos comparadores contra
  datos de prueba, contiene solo `.xlsx` (y las carpetas `AAAA/`, `AAAA/MM
  Mes/` que los contienen) — ni un `.json`, ni `parquet/`, ni `vistas/`.
- `python __comun__/test_salidas.py` sigue pasando después de borrar
  `raiz_salidas()` (o se lo actualiza si su prueba la usaba).

---

## Tarea 3 — Documentación y verificación final

Depende de que la 1 y la 2 estén aplicadas y verificadas.

### 3.1 — `README.md` (raíz)

- El árbol de "Estructura" ya no tiene `comun/` dentro de `Revisor_Relq/`;
  agregar `__comun__/` y `__config__/` como hermanas al principio del árbol.
- La sección "Para descargar y usar": aclarar que se descargan
  `__comun__/`, `Revisor_Relq/` y `Comparadores/` (las tres son código);
  `__config__/` la arma la usuaria a mano, no se descarga del repo, con la
  misma estructura `AAAA/MM Mes` que `00_Salidas/`.
- "Lo que no se sube y aparece solo al usarlo": reemplazar la lista actual
  (`config.json`, `00_Salidas/`, `Reemplazos REUC/Auxiliares/config.json`)
  por `__config__/` entera (con lo que contiene: `config.json` compartido,
  `reemplazos_reuc/config.json`, y el árbol `AAAA/...` de cada programa) y
  `00_Salidas/` (con su árbol `AAAA/MM Mes` de resultados).
- La nota final ("`config.json` vive en `Revisor_Relq/`, compartido...") —
  reescribirla: ahora vive en `__config__/`, sigue siendo compartido entre
  el Revisor y los 8 de `actualizadores/`.

### 3.2 — `AGENTS.md`

- Sección de convenciones de config: donde dice que `config.json` es
  compartido y vive en `Revisor_Relq/`, corregir a `__config__/`.
- Sección "Dónde vive el código" (o equivalente): agregar `__comun__/` como
  carpeta de primer nivel, sacar `comun/` de adentro de `Revisor_Relq/`.
- Tabla de "Trampas conocidas": la fila que ya existe sobre armar
  `00_Salidas/AAAA/MM Mes` a mano en vez de vía `comun/salidas.py` — ahora
  aplica IGUAL a `__config__/AAAA/MM Mes` (mismo riesgo, misma causa: dos
  scripts tienen que ponerse de acuerdo en la ruta sin verse). Ampliar esa
  fila para que mencione las dos raíces, o agregar una fila nueva —
  cualquiera de las dos formas sirve, lo que no puede pasar es que quede
  documentado solo para `00_Salidas` cuando ahora aplica a dos carpetas.
- Si `AGENTS.md` menciona en algún lado el patrón
  `sys.path.insert(0, ...)` + `from comun import ...` como convención del
  repo (es probable, documentaba el patrón de los comparadores), actualizarlo
  al patrón nuevo: `_hallar_workroot()` + `from __comun__ import ...`.

### 3.3 — `MAPA.md`

- El diagrama "Dónde vive cada uno": agregar `__comun__/` y `__config__/`
  como hermanas de nivel superior, sacar `comun/` de adentro de
  `Revisor_Relq/`.
- La sección "El módulo común" (`comun/config.py`, `comun/salidas.py`,
  `comun/tema.py`): renombrar las referencias a `__comun__/config.py`, etc.
  Actualizar lo que diga sobre `raiz_salidas()` si esa función se borró en
  2.6.
- Los bloques de cada script (`Revisor_Reliquidacion.py`,
  `ActualizaRemplazos.py`, los 8 actualizadores, los 2 comparadores): donde
  digan de dónde leen/escriben su `config.json` o su estado, actualizar a
  `__config__/...`.

### 3.4 — `Revisor_Relq/README.md`, `Revisor_Relq/actualizadores/README.md`, `Comparadores/README.md`

Repasar cada uno por si mencionan `comun/` o rutas de `config.json` /
`00_Salidas` para estado — corregir lo que corresponda con el mismo criterio
de arriba.

### 3.5 — `generar_interfaces.py` / `INTERFACES.md`

Ya se tocó en la Tarea 1.4 (movió `comun` → `__comun__` en `CARPETAS`). Acá
solo confirmar que sigue estando bien después de todos los cambios de
código de la Tarea 2 (que no tocan firmas, así que no debería haber diff,
pero regenerar y correr `--check` de nuevo para estar seguros).

### 3.6 — `BITACORA.md`

Agregar la entrada de esta sesión (qué se hizo, qué se verificó, qué quedó
pendiente) y actualizar "Pendientes abiertos ahora mismo":

- El ítem "El usuario tiene que mover a mano las carpetas de `00_Salidas` al
  formato `AAAA/MM Mes`" sigue vigente tal cual.
- Agregar un ítem nuevo: la usuaria tiene que armar `__config__/` a mano con
  la misma estructura `AAAA/MM Mes`, migrando ahí sus `config.json`,
  `_revisor_verificaciones.json`, `_revisor_cache_valores.json`,
  `_traspaso_actualizador.json`, `estado.json`, `rutas.json`, `parquet/` y
  `vistas/` existentes (los que ya tenga de antes de este cambio) — el
  código no los migra solo.
- El ítem sobre `leer_config`/`guardar_config`/`escribir_json_atomico`
  duplicados en los comparadores sigue pendiente igual (no es parte de este
  plan).
- Los ítems sobre los 7 actualizadores sin migrar a `comun/config.py` y
  sobre `docs/ESTRUCTURA_CASO_RELIQUIDACION.md` siguen igual, solo que ahora
  el módulo se llama `__comun__/config.py`.

### 3.7 — Borrar este plan

Con las Tareas 1, 2 y 3 aplicadas y verificadas, borrar
`docs/PLAN_reorganizacion_config.md` en el mismo commit final (mismo ciclo
que tuvo `docs/PLAN_comparadores.md`).

### Verificación final (antes de borrar el plan)

- `python generar_interfaces.py --check` limpio.
- Las tres suites de `__comun__/test_*.py` pasan.
- Ningún `.py` del repo importa `comun` (sin dunder) en ningún lado.
- Ningún `.py` del repo escribe o lee un `.json` fuera de `__config__/`
  (repasado uno por uno en la Tarea 2, reconfirmar acá con una pasada
  final).
- `00_Salidas/` solo tiene `.xlsx` después de una corrida de prueba
  completa (Revisor + los dos comparadores).
- `.gitignore` ignora `__config__/` entera y NO ignora `__comun__/`
  (confirmar con `git status` después de crear un `__config__/` de prueba
  con algún `.json` adentro: no debe aparecer como candidato a `git add`).
- Los seis scripts con `_hallar_workroot()` (Revisor, ActualizaRemplazos,
  los 2 comparadores, Actualiza_SC_CO, y cualquier otro que la Tarea 2 haya
  sumado) muestran el error explicado (no un traceback pelado) si se los
  corre desde una copia a la que le falta `__comun__/`.
