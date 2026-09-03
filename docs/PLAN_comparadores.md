# PLAN — comparadores + reorganización de `00_Salidas`

> **Documento de trabajo, temporal.** Es el plan para dos tareas concretas.
> Cuando las dos estén aplicadas y verificadas, este archivo se borra y lo que
> valga la pena queda en `MAPA.md` / `AGENTS.md`.
>
> Lo escribió Claude; lo aplica Codex. Antes de empezar, `REGLAS.md` como
> siempre.

---

## 1. Qué se quiere

Tres cosas, pedidas por el usuario:

1. **Los dos comparadores** (`Comparador_Etapas.py`, `Comparador_Tabulado.py`)
   pasan a vivir en `Comparadores/`, **carpeta hermana de `Revisor_Relq/`**.
2. **Las rutas salen del mismo origen que las del Revisor** — el `config.json`
   compartido y los `_traspaso_actualizador.json` por mes — para no tener un
   almacén de rutas por herramienta.
3. **`00_Salidas` se reorganiza por año**: hoy es una carpeta plana por mes
   (`00_Salidas/2407/`); pasa a ser `00_Salidas/AAAA/MM Mes/`. Las carpetas
   `_comparador` y `_comparador_tabulado` son **anuales**, así que quedan
   dentro de su año.

El usuario mueve las carpetas a mano en su equipo. El código tiene que hablar
el mismo idioma que esa reorganización.

---

## 2. Estado actual — hechos verificados, no supuestos

### El Revisor

- `DIR_SALIDAS = DIR_SCRIPT.parent / "00_Salidas"` (línea ~1100). Ya está bien:
  es hermana de `Revisor_Relq/`. **No se toca.**
- **`dir_mes(aamm, crear=False)` es el único punto por donde pasa el armado de
  la carpeta del mes.** Sus ~10 usos (líneas 1525, 1542, 1543, 1624, 1636,
  1637, 1693, 4046, 4578, 4926) llaman siempre a esa función. Cambiar la
  estructura de carpetas es cambiar **una sola función**.
- **Ningún actualizador usa `DIR_SALIDAS` ni `dir_mes`.** Reciben la ruta del
  JSON de traspaso por `argv[1]`, absoluta. No hay que tocarlos.
- Ojo con la línea 4046: `dir_mes(aamm or "sin_mes", crear=True)`. `"sin_mes"`
  **no** es un AAMM válido y tiene que seguir funcionando.

### Los comparadores

Los dos están escritos para vivir **al lado** de `Salidas/` y del
`config.json` — es decir, pensados para estar dentro de `Revisor_Relq/`:

```
BASE        = Path(__file__).resolve().parent
SALIDAS     = BASE / "Salidas"              ← Etapas 142 · Tabulado 157
CONFIG_PATH = BASE / "config.json"          ← Etapas 152 · Tabulado 163
CDIR        = SALIDAS / "_comparador"           ← Etapas 143
CDIR        = SALIDAS / "_comparador_tabulado"  ← Tabulado 158
```

**Lo bueno: el punto 2 del pedido ya está casi resuelto en el código.** Los dos
leen el `config.json` compartido con el mismo esquema `<host>_<usuario>` que
el resto del repo (`leer_config` / `guardar_config`), y los dos leen y
escriben los `_traspaso_actualizador.json` del Revisor:

- `ruta_json_mes(aamm)` → `SALIDAS / aamm / "_traspaso_actualizador.json"`
  (Etapas 424 · Tabulado 565).
- `Comparador_Etapas` **escribe** en ese JSON bajo su propia clave
  `comparador_etapas`, sin tocar `rutas`, `planilla` ni nada ajeno
  (`guardar_rpre_en_json_mes`, línea ~452). Ese contrato **se conserva tal
  cual**.

Así que del punto 2 solo falta que, después de la mudanza, `CONFIG_PATH`
siga apuntando al `config.json` del Revisor y no a uno nuevo.

- **Ninguno de los dos recorre la carpeta `Salidas`.** Arman la lista de meses
  con `meses_del_anio(anio)`, que es aritmética pura. Los únicos accesos a
  carpetas de mes son `ruta_json_mes(aamm)` y `path_excel_mes(aamm)`
  (Etapas 1032 · Tabulado 840). Eso hace la mudanza mucho más barata de lo que
  parece.

- El año lo elige el usuario en la ventana (`self.var_anio`, guardado en
  `config.json` como `comp_anio`). O sea: **la herramienta ya trabaja sobre un
  año a la vez**, que es justo lo que hace natural volver anuales las carpetas.

- `Comparador_Tabulado` lee la carpeta del otro comparador:
  `CDIR_MDB = SALIDAS / "_comparador"` y `RUTAS_MDB_PATH = CDIR_MDB /
  "rutas.json"` (líneas 167-168), "para no pedirle dos veces lo mismo al
  usuario". **Esa dependencia cruzada tiene que seguir apuntando al mismo
  año.**

### ⚠️ Un bug que hay que arreglar de paso (`Comparador_Tabulado.py`)

`DIR_PARQUET` y `DIR_VISTAS` están definidos **dos veces**:

| Línea | Definición |
|---|---|
| 159-160 | `CDIR / "parquet"` · `CDIR / "vistas"` |
| 827-828 | `CDIR / "parquet_variables"` · `CDIR / "vistas_variables"` |

Todos los usos (832, 836, 914, 917, 1710) están **después** de la 827, así que
lo que corre de verdad es `parquet_variables` / `vistas_variables`. Las líneas
159-160 son código muerto **y engañan**: quien lee el bloque de constantes de
arriba cree que los datos van a `_comparador_tabulado/parquet/`.

**Al refactorizar hay que preservar el comportamiento real
(`parquet_variables` / `vistas_variables`) y borrar las definiciones muertas de
las líneas 159-160.** No al revés.

---

## 3. La estructura nueva de `00_Salidas`

```
00_Salidas/
├── 2024/
│   ├── 01 Enero/
│   ├── ...
│   ├── 07 Julio/
│   │   ├── _revisor_verificaciones.json
│   │   ├── _revisor_cache_valores.json
│   │   ├── _traspaso_actualizador.json
│   │   ├── Comparacion_2407.xlsx
│   │   └── Comparacion_Variables_2407.xlsx
│   ├── 12 Diciembre/
│   ├── _comparador/              ← anual
│   │   ├── estado.json
│   │   ├── rutas.json
│   │   ├── parquet/ · vistas/ · actual/
│   │   └── Comparacion_Etapas_2024.xlsx
│   └── _comparador_tabulado/     ← anual
│       ├── estado.json · rutas.json
│       ├── parquet_variables/ · vistas_variables/
│       └── Comparacion_Variables_2024.xlsx
├── 2025/
└── sin_mes/                      ← el caso sin AAMM, fuera de los años
```

**Nombres de mes**, con cero adelante y mayúscula inicial:

```
01 Enero      04 Abril    07 Julio      10 Octubre
02 Febrero    05 Mayo     08 Agosto     11 Noviembre
03 Marzo      06 Junio    09 Septiembre 12 Diciembre
```

Es el mismo formato que ya usa el árbol `T:\Facturacion\<AAAA>\<MM Mes>\` que
`Comparador_Etapas` recorre hoy (ver `PAT_ANIO_DIR`, `PAT_MES_DIR`,
`RAMAS_FACT`). No es una convención nueva.

`AAAA` sale de `AAMM` como `"20" + aamm[:2]` (`2407` → `2024`).

---

## 4. Decisión de diseño: dónde vive la lógica de carpetas

El Revisor y los dos comparadores tienen que estar **exactamente de acuerdo**
sobre cómo se llama la carpeta de un mes, para siempre. Si se separan, el
comparador lee o escribe en la carpeta equivocada **sin ningún error visible**
— y esa es la clase de bug que este repositorio ya documentó con
`CENTRALES_EMBALSE` (ver `AGENTS.md`, "Trampas conocidas").

**Por eso la lógica va en `Revisor_Relq/comun/salidas.py`, una sola vez**, y
los tres la importan. Es la segunda pieza del módulo común, y se justifica por
un riesgo real, no por prolijidad.

### Cómo llegan los comparadores a `comun/`

`Comparadores/` es hermana de `Revisor_Relq/`, así que ni `comun/` ni el
`config.json` están al lado. Y **el nombre `Revisor_Relq` ya cambió tres
veces**, así que hardcodearlo es frágil.

Solución: un pequeño localizador al inicio de cada comparador, que busca entre
las carpetas hermanas la que contenga `Revisor_Reliquidacion.py`:

```python
BASE = Path(__file__).resolve().parent

def _hallar_revisor(raiz):
    """La carpeta hermana que contiene Revisor_Reliquidacion.py.

    Se busca por el archivo y no por el nombre de la carpeta porque ese
    nombre ya cambio tres veces (scripts -> Revisor Reliquidacion ->
    Revisor -> Revisor_Relq).
    """
    preferida = raiz / "Revisor_Relq"           # camino rapido
    if (preferida / "Revisor_Reliquidacion.py").is_file():
        return preferida
    for d in sorted(p for p in raiz.iterdir() if p.is_dir()):
        if (d / "Revisor_Reliquidacion.py").is_file():
            return d
    return None

DIR_REVISOR = _hallar_revisor(BASE.parent)
if DIR_REVISOR is None:
    _morir("No se encontró la carpeta del Revisor",
           "Este comparador tiene que estar en una carpeta hermana de la del\n"
           "Revisor (la que contiene Revisor_Reliquidacion.py).\n\n"
           f"Se buscó en: {BASE.parent}")

sys.path.insert(0, str(DIR_REVISOR))
from comun import salidas as _sal

CONFIG_PATH = DIR_REVISOR / "config.json"
SALIDAS = _sal.raiz_salidas(BASE)      # = BASE.parent / "00_Salidas"
```

`_morir` ya existe en el repo (ver `Revisor_Relq/actualizadores/Actualiza_SC_CO.py`):
abre una ventana con el motivo en vez de morir callado bajo `pythonw`. Los
comparadores ya tienen una función equivalente; **usar la que ya tengan, no
inventar otra**.

> Nota: `raiz_salidas(dir_script)` devuelve `dir_script.parent / "00_Salidas"`,
> y da lo mismo para los dos casos — desde `Revisor_Relq/` y desde
> `Comparadores/`, el padre es la misma carpeta de trabajo.

---

## 5. TAREA 1 — `comun/salidas.py` + el Revisor

**Alcance: solo el Revisor y el módulo común. No tocar los comparadores.**
Al terminar esta tarea el Revisor tiene que funcionar con la estructura nueva,
verificado, commiteado y pusheado. Es autocontenida.

### 5.1 Crear `Revisor_Relq/comun/salidas.py`

Solo biblioteca estándar. API:

```python
MESES = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")

def raiz_salidas(dir_script) -> Path
    """00_Salidas, hermana de la carpeta del script. dir_script.parent / '00_Salidas'."""

def partir_aamm(aamm) -> tuple[str, int] | None
    """'2407' -> ('2024', 7). None si no es un AAMM valido (incluye 'sin_mes')."""

def nombre_carpeta_mes(aamm) -> str | None
    """'2407' -> '07 Julio'. None si el aamm no es valido."""

def carpeta_anio(dir_salidas, aamm) -> Path | None
    """00_Salidas/2024. None si el aamm no es valido."""

def carpeta_mes(dir_salidas, aamm, crear=False) -> Path
    """00_Salidas/2024/07 Julio.

    Si la carpeta canonica no existe pero hay una variante escrita distinto
    en el mismo anio ('7 Julio', '07 julio'), devuelve esa: el usuario mueve
    estas carpetas a mano. Si el aamm no es valido ('sin_mes'), devuelve
    dir_salidas / aamm, plana, como antes.
    """

def carpeta_comparador(dir_salidas, anio, nombre) -> Path
    """00_Salidas/2024/_comparador (o el nombre que se pida), por anio."""

def carpetas_legado(dir_salidas) -> list[Path]
    """Las carpetas planas con el formato viejo (00_Salidas/2407) que sigan
    existiendo. Sirve para avisar una vez que hay que moverlas; nunca se
    escribe en ellas."""
```

Reglas que tienen que quedar en el código:

- **`carpeta_mes` nunca devuelve una carpeta con el formato viejo.** Si
  devolviera la vieja para leer y la nueva para escribir, el estado quedaría
  partido en dos lugares sin que nadie se entere. Para el formato viejo está
  `carpetas_legado`, que solo sirve para avisar.
- La comparación tolerante de nombres usa la misma normalización que ya usa el
  repo (sin tildes, sin espacios de más, en mayúsculas). Hay ejemplos en
  `comun/config.py` y en los comparadores (`normalizar_suave`, `subcarpeta`).
- `crear=True` crea siempre la **canónica**, con `parents=True, exist_ok=True`.

### 5.2 Pruebas: `Revisor_Relq/comun/test_salidas.py`

Mismo estilo que `test_config.py` (solo stdlib, `unittest`, carpeta temporal).
Como mínimo:

- `partir_aamm`: `'2407'` → `('2024', 7)`; `'2312'` → `('2023', 12)`;
  `'sin_mes'`, `''`, `'24'`, `'240713'`, `'2400'`, `'2413'` → `None`.
- `nombre_carpeta_mes('2401')` → `'01 Enero'`; `('2409')` → `'09 Septiembre'`.
- `carpeta_mes` devuelve la canónica cuando no hay nada en disco.
- `carpeta_mes` encuentra la variante `'7 Julio'` y `'07 julio'` si ya existen.
- `carpeta_mes(..., 'sin_mes')` → `00_Salidas/sin_mes`, plana.
- `carpeta_mes(..., crear=True)` crea la canónica aunque el año no exista.
- **`carpeta_mes` NO devuelve la carpeta vieja `00_Salidas/2407` aunque exista.**
- `carpetas_legado` sí la encuentra, y no confunde `2024/` (año) con `2407/`
  (formato viejo).

### 5.3 Cambiar `dir_mes` en el Revisor

```python
from comun import salidas as _sal      # junto a los imports de arriba

def dir_mes(aamm, crear=False):
    """00_Salidas/AAAA/MM Mes, hermana de Revisor_Relq.

    La logica vive en comun/salidas.py porque los comparadores tienen que
    armar exactamente la misma ruta; si se separan, uno lee donde el otro no
    escribe y no se entera nadie.
    """
    return _sal.carpeta_mes(DIR_SALIDAS, aamm, crear=crear)
```

**No tocar ninguno de los ~10 puntos de llamada.** Siguen igual.

`DIR_SALIDAS` se puede dejar como está o pasarlo a `_sal.raiz_salidas(DIR_SCRIPT)`;
las dos dan lo mismo. Si se cambia, que sea por claridad, no por necesidad.

### 5.4 Aviso de carpetas viejas

Una vez, al arrancar (donde el Revisor ya escribe en su bitácora de pantalla),
si `carpetas_legado(DIR_SALIDAS)` devuelve algo:

```
OJO: hay 3 carpeta(s) con el formato viejo en 00_Salidas (2407, 2408, 2409).
     La estructura nueva es 00_Salidas/AAAA/MM Mes (ej: 2024/07 Julio).
     Mové el contenido a mano; desde acá no se lee ni se escribe en ellas.
```

Barato, y evita que el usuario pierda el estado de verificaciones de un mes sin
darse cuenta.

### 5.5 Documentación de la Tarea 1

- `MAPA.md`: el bloque del Revisor dice hoy `00_Salidas/AAMM/`. Actualizar a la
  estructura nueva. Agregar `comun/salidas.py` a la sección "El módulo común",
  con el mismo formato que tiene `comun/config.py`.
- `AGENTS.md`: la línea del JSON de traspaso (sección 4) dice
  `00_Salidas/AAMM/_traspaso_actualizador.json` → actualizar.
- `README.md`: si menciona la estructura de `00_Salidas`, actualizar.
- `INTERFACES.md`: regenerar (`python generar_interfaces.py`).
- `BITACORA.md`: entrada nueva.

### 5.6 Verificación de la Tarea 1

```
python Revisor_Relq/comun/test_salidas.py     # las pruebas nuevas
python Revisor_Relq/comun/test_config.py      # que no se rompió lo de antes
python generar_interfaces.py --check
python -c "import ast;ast.parse(open('Revisor_Relq/Revisor_Reliquidacion.py',encoding='utf-8').read())"
```

Y una comprobación de que el Revisor arma la ruta esperada, sin abrir la
ventana (hay ejemplos de este patrón, con `tkinter` simulado, en `BITACORA.md`
→ entrada "separa el Revisor de los actualizadores"):

- `dir_mes("2407")` termina en `00_Salidas/2024/07 Julio`.
- `dir_mes("sin_mes")` termina en `00_Salidas/sin_mes`.

---

## 6. TAREA 2 — los dos comparadores

**No empezar hasta que la Tarea 1 esté pusheada y verificada.** La Tarea 2
depende de `comun/salidas.py`.

Los dos archivos **ya están en `Comparadores/`, sin modificar**, tal como los
entregó el usuario. Hoy están rotos ahí (apuntan a `BASE / "Salidas"`, que no
existe en esa ubicación). Esta tarea los deja andando.

### 6.1 Cabecera común a los dos

Reemplazar el bloque de constantes de rutas por el localizador de la
sección 4: `DIR_REVISOR`, `CONFIG_PATH`, `SALIDAS`, y el `sys.path.insert`
antes de `from comun import salidas as _sal`.

Actualizar también el docstring de arriba de cada archivo: hoy los dos dicen
*"Ubicacion: este .py va en la misma carpeta que contiene la carpeta Salidas"*,
que deja de ser cierto.

### 6.2 Las carpetas del comparador pasan a ser anuales

Hoy son constantes de módulo colgando de `CDIR`. Pasan a ser **funciones del
año**. El año siempre está disponible:

- Las funciones que ya reciben `aamm` lo derivan (`"20" + aamm[:2]`); conviene
  un helper local `_anio_de(aamm)` que use `_sal.partir_aamm`.
- Las funciones de nivel año (`cargar_estado`, `guardar_estado`, las de
  `rutas.json`, `ACTUAL_PARQUET`) **reciben el año como parámetro**. Sus
  llamadores están en la clase de la ventana, donde `self.var_anio` ya existe.

`Comparador_Etapas.py` — constantes a convertir y cuántos usos tiene cada una:

| Constante | Usos | Pasa a |
|---|---|---|
| `CDIR` | 10 | `cdir(anio)` → `_sal.carpeta_comparador(SALIDAS, anio, "_comparador")` |
| `DIR_PARQUET` | 3 | `cdir(anio) / "parquet"` |
| `DIR_SOB` | 4 | `… / "sobrecostos"` |
| `DIR_CEN` | 4 | `… / "centrales"` |
| `DIR_VISTAS` | 4 | `cdir(anio) / "vistas"` |
| `DIR_ACTUAL` | 3 | `cdir(anio) / "actual"` |
| `ACTUAL_PARQUET` | 7 | `dir_actual(anio) / "central_empresa_actual.parquet"` |
| `ESTADO_PATH` | 3 | `cdir(anio) / "estado.json"` |
| `RUTAS_PATH` | 4 | `cdir(anio) / "rutas.json"` |

`Comparador_Tabulado.py`:

| Constante | Usos | Pasa a |
|---|---|---|
| `CDIR` | 11 | `cdir(anio)` → `… "_comparador_tabulado"` |
| `DIR_PARQUET` | 5 | `cdir(anio) / "parquet_variables"` ← **el real, no `parquet`** |
| `DIR_VISTAS` | 5 | `cdir(anio) / "vistas_variables"` ← **idem** |
| `ESTADO_PATH` | 3 | `cdir(anio) / "estado.json"` |
| `RUTAS_PATH` | 5 | `cdir(anio) / "rutas.json"` |
| `CDIR_MDB` | 2 | `_sal.carpeta_comparador(SALIDAS, anio, "_comparador")` |
| `RUTAS_MDB_PATH` | 2 | `cdir_mdb(anio) / "rutas.json"` |

Y **borrar las definiciones muertas de las líneas 159-160** (ver sección 2).

### 6.3 Las rutas de mes

```python
def ruta_json_mes(aamm):
    return _sal.carpeta_mes(SALIDAS, aamm) / NOMBRE_JSON_MES

def path_excel_mes(aamm):
    return _sal.carpeta_mes(SALIDAS, aamm) / f"Comparacion_{aamm}.xlsx"
```

(en Tabulado, `f"Comparacion_Variables_{aamm}.xlsx"`)

`guardar_rpre_en_json_mes` tiene que seguir creando la carpeta del mes si no
existe — hoy `escribir_json_atomico` hace `mkdir(parents=True)` sobre el padre,
así que sigue andando, pero **verificarlo**, porque ahora el padre son dos
niveles y no uno.

### 6.4 Lo que NO se cambia

- **El contrato del JSON de traspaso.** `Comparador_Etapas` sigue escribiendo
  solo bajo `comparador_etapas`, sin tocar `rutas`, `planilla`, `nodo` ni nada
  del Revisor. Ese acuerdo es lo que hace que los dos convivan.
- **`rutas.json` se queda como archivo**, dentro del `_comparador*` del año.
  No se funde en `config.json`. Razón: ahora que la carpeta es anual, ese
  archivo es exactamente lo que su nombre dice — las rutas elegidas a mano
  **para ese año** — y ese es su lugar natural. Meterlo en `config.json`
  mezclaría datos por año con configuración por equipo. (Si el usuario
  prefiere lo contrario, se hace después: es un cambio chico y aislado.)
- **`estado.json`** también se queda: es el índice de lo ya consolidado, no
  configuración.
- La lógica de comparación, lectura de Access, parquet, duckdb y armado de
  Excel **no se toca**. Esta tarea es de rutas.

### 6.5 Registrar los comparadores en el resto del repo

- `generar_interfaces.py`: `CARPETAS = ["Revisor_Relq/comun", "Revisor_Relq"]`
  → agregar `"Comparadores"`.
- `MAPA.md`: un bloque por comparador, con el mismo formato que los demás
  (**qué hace · consume · produce · expone · depende de**). Van después de los
  actualizadores. Mencionar la dependencia cruzada: `Comparador_Tabulado` lee
  el `rutas.json` de `_comparador` del mismo año.
- `MAPA.md`, sección "Dónde vive cada uno": agregar `Comparadores/` al árbol.
- `README.md`: agregar `Comparadores/` al árbol de estructura y a la
  explicación de qué se descarga (ahora son **dos** carpetas hermanas, más
  `00_Salidas/`).
- `AGENTS.md`: si hace falta, una línea en las convenciones sobre que un script
  fuera de `Revisor_Relq/` localiza al Revisor por el archivo
  `Revisor_Reliquidacion.py`, no por el nombre de la carpeta.
- `INTERFACES.md`: regenerar.
- `BITACORA.md`: entrada nueva.

### 6.6 Verificación de la Tarea 2

```
python generar_interfaces.py --check
python Revisor_Relq/comun/test_salidas.py
python Revisor_Relq/comun/test_config.py
python -c "import ast;ast.parse(open('Comparadores/Comparador_Etapas.py',encoding='utf-8').read())"
python -c "import ast;ast.parse(open('Comparadores/Comparador_Tabulado.py',encoding='utf-8').read())"
```

Y, con `tkinter` y las dependencias pesadas simuladas (`pandas`, `pyodbc`,
`duckdb`, `pyarrow`, `xlsxwriter` — los comparadores ya las cargan de forma
perezosa con `_Perezoso`, así que importarlos no debería requerirlas):

- Que `DIR_REVISOR` se resuelva a la carpeta del Revisor.
- Que `CONFIG_PATH` sea **el mismo archivo** que el `CONFIG_PATH` del Revisor.
  Esta es la comprobación central del pedido 2: si no dan el mismo `Path`, el
  comparador se armó su propio `config.json` y hay que corregirlo.
- Que `ruta_json_mes("2407")` termine en `00_Salidas/2024/07 Julio/_traspaso_actualizador.json`.
- Que `cdir("2024")` termine en `00_Salidas/2024/_comparador`.
- Que en Tabulado, `DIR_PARQUET` del año termine en `parquet_variables` (no en
  `parquet`).

Hay un ejemplo de este tipo de prueba con módulos simulados en `BITACORA.md`,
entrada "separa el Revisor de los actualizadores".

---

## 7. Trampas — leer antes de tocar

| Trampa | Qué pasa si se ignora |
|---|---|
| Duplicar el armado de la carpeta del mes en vez de importarlo de `comun/salidas.py` | El comparador lee donde el Revisor no escribe. **Sin error visible.** Es el caso `CENTRALES_EMBALSE` otra vez. |
| Preservar `CDIR / "parquet"` en Tabulado | Se pierde de vista el parquet real (`parquet_variables`) y el comparador queda mirando una carpeta vacía. |
| Hardcodear `"Revisor_Relq"` en los comparadores | El nombre ya cambió 3 veces. Al cuarto, los comparadores dejan de encontrar el `config.json`. |
| Hacer que `carpeta_mes` devuelva la carpeta vieja si existe | El estado del mes queda partido: se lee de una y se escribe en otra. |
| Romper `dir_mes("sin_mes")` | El Revisor revienta al lanzar un actualizador sin AAMM definido (línea 4046). |
| Tocar la lógica de comparación "ya que estamos" | Fuera de alcance. Esta es una tarea de rutas. |

---

## 8. Qué queda pendiente después, y no es parte de esto

- El usuario tiene que **mover a mano** las carpetas de `00_Salidas` al formato
  nuevo. El aviso de la sección 5.4 le dice cuáles faltan.
- Los comparadores todavía tienen su propia copia de `leer_config` /
  `guardar_config` / `escribir_json_atomico`. Podrían pasar a
  `comun/config.py` como ya hizo `Actualiza_SC_CO.py`. **No en estas dos
  tareas** — se anota en `BITACORA.md` → "Pendientes abiertos" y se hace
  cuando toque, una pieza a la vez.

---

## 9. TAREA 3 — bugs encontrados en los comparadores

Revisión hecha sobre `Comparadores/*.py` sin correrlos (no hay Windows, Excel ni
Access acá). Cada hallazgo dice cómo verificarlo. **Van en orden de gravedad.**

Se puede hacer junto con la Tarea 2 o después, pero **no antes**: varios tocan
las mismas funciones que la Tarea 2 mueve de lugar.

### 3.1 🔴 Se escribe en la ventana desde el hilo de trabajo (LOS DOS)

`App.log()` (Etapas 1922 · Tabulado 1320) hace, tal cual:

```python
self.txt.insert("end", str(msg) + "\n")
self.txt.see("end")
self.root.update_idletasks()
```

y se lo llama **desde dentro de los hilos**: `correr` (Et 1957 · Ta 1353) y
`buscar` (Et 2103 · Ta 1438), en las líneas 1961-1962, 2113-2114, 2126 (Etapas)
y 1357-1358, 1449-1450, 1462 (Tabulado).

**Tkinter no es thread-safe.** Tocar widgets desde otro hilo da cuelgues y
crashes intermitentes — y `update_idletasks()` es lo peor de todo, porque
reentra al loop de eventos desde el hilo equivocado. No falla siempre: falla en
las corridas largas, que es justo cuando estas herramientas se usan (consolidar
12 meses de Access).

Los autores conocían la regla — usaron `self.root.after(0, ...)` para `pintar()`
y `botones()` (Ta 1362-1364, 1464). Se les escapó el log, que es lo que más se
llama.

**Arreglo: copiar el patrón que el Revisor ya usa y que funciona.** En
`Revisor_Reliquidacion.py`: `self.cola = queue.Queue()` (3290), un
`self.root.after(300, self._bombear_cola)` (3324) y `_bombear_cola` (6791) que
vacía la cola **desde el hilo principal**. El hilo solo hace
`self.cola.put(("log", mensaje))`. Ninguno de los dos comparadores importa
`queue` hoy (0 ocurrencias).

Verificación: que `grep -n "queue" Comparadores/*.py` deje de dar 0, y que
ninguna función lanzada con `threading.Thread(target=...)` llame a `self.log`
ni a ningún método de un widget de forma directa.

### 3.2 🔴 `Comparador_Tabulado` corta el Excel en silencio

Al pasar `LIMITE_FILAS_HOJA`, los dos hacen cosas distintas:

| | Qué hace |
|---|---|
| **Etapas** (1502) | abre una hoja nueva (`AAMM_2`, `_3`…) y sigue. **Correcto.** |
| **Tabulado** (1141) | `break` y **descarta el resto de las filas**. |

Dos problemas:

1. **El Excel queda incompleto pero parece completo.** En una herramienta cuyo
   único trabajo es encontrar diferencias entre etapas, eso significa poder
   concluir "acá no hay diferencias" cuando estaban en las filas cortadas. Para
   este dominio, es el peor tipo de error: silencioso y con plata detrás.
2. Ese `break` sale solo del `for reg in lote`, **no del `while True`**: el
   cursor se sigue vaciando lote por lote sin escribir nada, y el aviso se
   repite una vez por lote.

**Arreglo: que Tabulado haga lo mismo que Etapas** — hoja nueva y seguir. Es el
comportamiento correcto y además ya está escrito al lado, en el archivo hermano.

### 3.3 🟠 La hora mensual se calcula sin comprobar que el mes venga completo

`Comparador_Tabulado.acumulado_por_dia` (815) arma
`{día: horas_acumuladas_antes}` haciendo `groupby("dia")["hora_dia"].max()`, y
`hora_mensual` (804) suma eso a la hora del día.

El comentario explica bien por qué usa `max()` y no la cantidad de filas (para
que un día de 23 o 25 horas por cambio de hora no desalinee). Está bien pensado.
**Lo que falta es comprobar que los días estén completos y contiguos.**

Si al archivo de una etapa le falta un día entero (o le faltan las últimas horas
de un día), el acumulado de todos los días siguientes se corre — y `hora_mes` es
**la clave con la que se cruzan las etapas** (`SELECT etapa, central, tipo,
hora_mes`, líneas 923-997). O sea: se comparan horas distintas entre sí. El
resultado no es un error, es una **tabla de diferencias inventadas**, o peor,
diferencias reales tapadas.

El dominio ya documenta esta trampa: `docs/ESTRUCTURA_CASO_RELIQUIDACION.md`,
"El cambio de hora", con la hora 145 que no existe.

**Arreglo:** después de armar `acc`, avisar **fuerte** en el log (no cortar, que
el usuario decida) si:

- los días no van de 1 a `max(dias)` sin huecos;
- dentro de un día las horas no van de 1 a `max` sin huecos;
- `min(hora_dia) != 1` (si algún archivo viniera con horas desde 0, **todo** se
  corre una hora y hoy no se notaría);
- el total de horas del mes no es 672/696/720/744 ±1 (el ±1 es el cambio de hora).

Con el nombre del archivo y el mes en el mensaje, para que se pueda ir a mirar.

### 3.4 🟡 283 líneas duplicadas entre los dos comparadores

33 funciones son **idénticas carácter por carácter** en los dos archivos:
`leer_config`, `guardar_config`, `escribir_json_atomico`, `get_usuario`,
`leer_json`, `normalizar`, `normalizar_suave`, `listar`, `huella`,
`huella_entrada`, `limpiar_cache`, `subcarpeta`, `buscar_mdb`, `es_copia`,
`carpeta_definitivo`, `meses_del_anio`, `fmt_tiempo`, `ahora`,
`abrir_en_explorador`, `respaldar`, `cargar_estado`, `guardar_estado`,
`firma_vistas`, `color_de`, `mes_incluido`, `fijar_incluido`, `una_fila`,
`tabla_por_nombre`, `es_hoja_propia`, `hojas_ajenas`, `path_vista`,
`ruta_json_mes`, `_falta`.

Otras 12 comparten nombre y difieren; se revisaron las dos más delicadas:
`rutas_desde_json_mes` difiere **solo en el nombre de una variable** (`rutas` vs
`r`), y `estado_etapa` difiere por una razón real (Etapas tiene 2 archivos por
etapa, Tabulado 1). Ninguna de las dos es un bug hoy.

Pero es exactamente la trampa de `CENTRALES_EMBALSE` a mayor escala: arreglar
algo en uno y no en el otro. **No hacer esta migración dentro de esta tarea** —
va a `comun/`, una pieza a la vez, como manda `AGENTS.md`. Se anota acá para que
quede el número medido y el orden sugerido:

1. `leer_config` / `guardar_config` / `escribir_json_atomico` / `get_usuario` →
   ya existen en `comun/config.py`. Es reemplazo directo, igual que se hizo con
   `Actualiza_SC_CO.py`.
2. `normalizar` / `normalizar_suave` / `es_copia` / `subcarpeta` / `listar` /
   `huella` → un `comun/archivos.py`.
3. El resto, cuando toque.

### 3.5 Lo que se revisó y está bien

Para que nadie lo revise dos veces:

- **Cierre de conexiones Access y de libros Excel**: todas las aperturas tienen
  su `try/finally` con `con.close()` / `wb.close()`.
- **El marshalling a la ventana de `pintar()`, `botones()` y `set_estado()`**:
  correcto, va por `self.root.after(0, ...)`. El problema es solo el log (3.1).
- **`rutas_desde_json_mes`** en los dos archivos: mismo comportamiento.
- **El contrato del JSON de traspaso**: `Comparador_Etapas` escribe solo bajo su
  clave `comparador_etapas`, sin tocar nada del Revisor. Correcto.
- **`hora_mensual` con `max()` en vez de contar filas**: la decisión es la
  correcta para el cambio de hora. Lo que falta es la validación (3.3).

---

## 10. TAREA 4 — tema oscuro (experimental, va al final)

El usuario dijo que las ventanas le parecen feas y quiere probar un tema oscuro.
Es **para probar**, así que la regla es: que se pueda volver atrás sin tocar
código.

### 10.1 Piloto primero, no las 12 ventanas

**Aplicarlo solo a los dos comparadores.** Son los que menos se usan, así que si
queda mal no molesta en el trabajo diario. Si al usuario le gusta, después se
lleva al Revisor y a los actualizadores; si no, se descarta y no se tocó nada
importante.

### 10.2 `Revisor_Relq/comun/tema.py`

Una paleta y una función que la aplica:

```python
def aplicar(root, modo="oscuro") -> dict
    """Configura los estilos ttk de la ventana y devuelve la paleta.

    Devuelve el dict de colores para que quien lo llame pinte tambien los
    widgets tk clasicos (Text, Label, Entry, Canvas), que NO siguen los
    estilos ttk.
    """
```

Tres cosas que hay que resolver sí o sí, porque son las que hacen que un tema
oscuro en tkinter quede a medias:

1. **Los widgets `tk.` clásicos no siguen a ttk.** Estos scripts usan mezcla:
   `tk.Text` (el log), `tk.Label`, `tk.Entry`, `tk.Frame`. Hay que pintarlos a
   mano con `bg`/`fg`/`insertbackground` desde la paleta que devuelve `aplicar`.
2. **Los colores de estado están calibrados para fondo claro** y sobre oscuro no
   se leen: `COLOR_ROJO = "#c0392b"`, `COLOR_AMARILLO = "#b8860b"`,
   `COLOR_VERDE = "#1e7a1e"` en los comparadores, y `C_OK`, `C_FALTA`,
   `C_AMARILLO`, `C_VENCIDA` en el Revisor. La paleta tiene que traer **su
   propia versión de cada color de estado**, más clara, y el código usarla desde
   ahí en vez de las constantes de módulo.
3. **`C_NEUTRO = "SystemButtonFace"`** (Revisor) es un color del sistema: en
   tema oscuro queda un parche gris claro. Reemplazarlo por el de la paleta.

### 10.3 Que se pueda apagar

El modo sale de `config.json`, clave `tema` (`"claro"` | `"oscuro"`), por
equipo/usuario como todo lo demás. **Por omisión `"claro"`**, o sea: si el
usuario no hace nada, todo se ve igual que hoy. Un menú o casilla en la ventana
lo cambia y lo guarda.

### 10.4 Qué NO hacer

- **No reacomodar la disposición de los controles.** Esto es color, no rediseño.
  Mover botones de lugar en una herramienta que el usuario ya tiene aprendida es
  una molestia, no una mejora.
- **No instalar librerías de temas** (`ttkbootstrap`, `sv-ttk`). El repo es solo
  biblioteca estándar y estas herramientas corren en equipos donde instalar algo
  es un trámite.
- **No tocar el tema si `aplicar()` falla.** Envolver en `try/except` y seguir
  con el aspecto de siempre: nadie se puede quedar sin poder trabajar porque un
  color no se pudo aplicar.

### 10.5 Verificación

No hay ventana en este entorno, así que la verificación real la hace el usuario
mirando. Lo que sí se puede comprobar acá:

- Que con `tema` sin definir en `config.json`, los colores que se usan sean
  exactamente los de hoy (o sea: por omisión no cambia nada).
- Que `aplicar()` con un `root` simulado no lance.
- Que ninguna de las constantes de color viejas quede sin usar y sin reemplazo.
