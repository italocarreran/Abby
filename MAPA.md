# MAPA — qué hace cada script y de qué depende

> **Regla de expansión.** Leer este archivo → leer de `INTERFACES.md` solo las
> entradas que hacen falta → abrir completo **únicamente** el archivo que se va a
> modificar. No abrir los vecinos "para tener contexto".

**Estado:** los 10 scripts están en el repositorio y los bloques de abajo están
**validados contra el código**. Donde el código difiere de
`docs/ESTRUCTURA_CASO_RELIQUIDACION.md`, manda el código y la diferencia está
anotada al final, en "Diferencias con el documento de dominio".

Convención de cada bloque: **qué hace · consume · produce · expone · depende de**.

**Dónde vive cada uno:**

```
Revisor Reliquidación/
├── Revisor_Reliquidacion.py     ← el único que va suelto en la raíz
├── comun/                        lo compartido (ver más abajo)
├── actualizadores/                los 8 que el Revisor lanza por botón
│   ├── Actualiza_datos.py
│   ├── Actualiza_Data_Access.py
│   ├── Actualiza_Energia.py
│   ├── Actualiza_Cuadro0.py
│   ├── Actualiza_SC_CO.py
│   ├── Actualiza_Access_P9.py
│   ├── Carga_Retiros.py
│   └── Prorratear.py
└── Reemplazos REUC/               aparte porque tiene su propio config.json
    └── ActualizaRemplazos.py
```

`config.json` es **compartido** entre el Revisor y los 8 de `actualizadores/`:
vive en `Revisor Reliquidación/`, un nivel arriba de ellos, y cada uno lo resuelve con
`DIR_SCRIPT.parent / "config.json"` (antes era `DIR_SCRIPT / "config.json"`,
porque vivían junto al Revisor). `Reemplazos REUC/` no: el suyo está en su
propia `Auxiliares/`, como siempre.

---

## Cómo se conecta todo

```
                     02 Consolidado_Tabulado  ← la fuente de casi todo
                       │
   Actualiza_datos ────┤            Actualiza_Energia ──→ 03b ENTRADA_SOB.mdb
   (FD + consolidado   │                              └─→ Consolidado_AAMM
    + prorrata)        │
        │              └── Actualiza_Data_Access ──→ 03b ENTRADA_SOB_SSCC.mdb
        ↓                        ▲ (motor reutilizado por Energia y P9)
   Cálculo_SobrecostosSSCC
   planillas 3 / 5 / 6           Actualiza_Access_P9 ──→ Ocupar_este...mdb
        │                                                     │
        ├── Actualiza_SC_CO ──→ hoja "SC y CO" de la 5_       │
        │                                                     ↓
   Carga_Retiros ──→ SQL Server (Retiros) ──→ Prorratear ──→ Pago_Retiro_reporte_tabla
                                                              │
   ActualizaRemplazos ──→ empresas + reemplazos ──┐           │
                                                  ↓           ↓
                                     Actualiza_Cuadro0 ──→ 0_CUADROS_RELIQUIDACIÓN
                                                              (el que se va a pago)

   Revisor_Reliquidacion  ← orquesta todo: verifica V4..V17 y lanza los de arriba
                            pasándoles Salidas/AAMM/_traspaso_actualizador.json
```

---

## `Revisor Reliquidación/Revisor_Reliquidacion.py`

- **Qué hace:** es la consola del proceso. Arma el árbol de archivos del mes,
  verifica por fecha de modificación (copias contra sus maestros) y por valores
  (V4…V17, con cadena de dependencias entre verificaciones), y desde cada fila lanza
  el actualizador que corresponde.
- **Consume:** todo el árbol de `02 CASO RELIQUIDACION` + `FD/`, las tres bases
  Access, `config.json` (`carpeta_base`, `ultimo_mes`, `carpetas_por_mes`,
  `_valores`).
- **Produce:** `Salidas/AAMM/` — estado, caché y `_traspaso_actualizador.json`, que
  le pasa como único argumento al actualizador que lanza.
- **Expone:** el JSON de traspaso (`origen`, `version`, `aamm`, `carpeta_reliq`,
  `planilla`, `rutas`, `nodo`, `clave_nodo`, `ruta_nodo`); `plan_traer_maestro()`
  separado del ejecutor; `cache_directorios()` como context manager;
  `CENTRALES_EMBALSE` (27); `UMBRAL_DESCUADRE_CPRT` (**1000.0**, no 500) y
  `UMBRAL_PAR_SEGURO` (100.0); `TOLERANCIA` (1.0), `TOL_PAGO_EMPRESA` (150.0),
  `TOL_SOBRECOSTO_FILA` (1.0), `TOL_MTIME` (2 s); `NODOS` (36) y `VERIFICADORES` (14).
- **Depende de:** los 9 scripts que lanza, pero solo por línea de comandos — no los
  importa. **Duplica `CENTRALES_EMBALSE` con `Actualiza_SC_CO.py`.**
- **Detalles que importan:** V17 es la raíz de casi todo (si el tabulado tiene el
  sobrecosto mal calculado, todos los totales cuadran igual y el error llega al
  cuadro de pago). El caché de directorios está apagado por omisión y solo se
  enciende dentro del `with`, porque `mtime()` se usa después para decidir si una
  verificación venció. La relectura es parcial: V9 relee 1 nodo, V16 relee 13 (el
  árbol tiene 36 nodos). `cache_directorios` es una clase, no una función: al salir
  deja `self.stats = (scans, hits)` para poder decir en la bitácora cuánto se ahorró.

## `Revisor Reliquidación/actualizadores/Actualiza_datos.py`

- **Qué hace:** trae las hojas FD desde `SSCC_Desempeno*`, el consolidado tabulado y
  la prorrata de retiros, a `Cálculo_SobrecostosSSCC` y a las planillas 3, 5 y 6.
  Tres opciones: "Actualizar FD", "Traer Consolidado", "Actualizar Prorrata".
- **Consume:** `FD/SSCC_Desempeno*.xlsx|xlsm` (hojas `CT Diario`, `CPF Horario`,
  `CSF Horario`, `CTF Horario`, desde la fila 12); `02 Consolidado_Tabulado`, hoja
  `Sobrecostos`; `04 Planilla 9/Prorrata_Retiros_*.xlsx`, hoja
  `PRORRATA_HORARIA_TABULAR`.
- **Produce:** escribe en `Cálculo_SobrecostosSSCC` (`FD_CT`, `FD_CPF`, `FD_CSF`,
  `FD_CTF`, `FD_CSF_Disponibilidad`, `FD_CTF_Disponibilidad`, `SOBRECOSTOS`) y en las
  planillas 3/5/6 (`CPF_FD`, `CSF_FD`, `CTF_FD`, `FD_CPF`, `FD_CSF`, `FD_CTF`, `FD`,
  `PRORRATA_RETIROS`).
- **Expone:** el parámetro `planilla` con valores `sc`, `p3`, `p5`, `p6`; es como lo
  llama el revisor desde cada fila.
- **Depende de:** `config.json` (clave `carpeta_reliq`).
- **Detalles que importan:** la prorrata se transforma de tabla larga a matriz
  pivote, se pega desde `B8` y se borra antes `B8:EC756`; suministradores ordenados
  alfabéticamente, faltantes en 0. El "Traer Consolidado" filtra solo las filas con
  columna `C = C.Frec`. Los `.xlsm` destino tienen que estar cerrados.

## `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py`

- **Qué hace:** consolida tres Excel (SSCC, CO, CCA) en la tabla `Sobrecostos` del
  `.mdb` de `01 Sobrecostos`. Es además **el motor** que reutilizan
  `Actualiza_Energia.py` y `Actualiza_Access_P9.py`.
- **Consume:** `01 Sobrecostos/Cálculo_SobrecostosSSCC*.xlsm` (hoja
  `SOBRECOSTOS TOTAL`, B:F, desde la fila 5) — **el maestro, no la copia de
  Entregables**; `Cálculo_CO*.xlsm` (hoja `CO TOTAL`); `Consolidado_CCA*.xlsm`
  (hoja `CCA`).
- **Produce:** filas en `03b ENTRADA_SOB_SSCC_*.mdb`, tabla `Sobrecostos`, en el
  orden `Clave Año_Mes, Tipo_sobrecosto, Central, Hora Mensual, Sobrecosto`. Borra
  solo los tipos que vienen en los Excel seleccionados; los demás quedan intactos.
- **Expone:** `proceso(..., fuentes=..., borrar_todo=...)`, `leer_fuente(cfg)` (que
  respeta `cfg["filtro"]` y `cfg["cols_no_cero"]`), la verificación posterior a la
  carga, el lector de Excel por ZIP/XML (`NS_XL`, `NS_REL`, `_ENT_XML`, `_RE_ENT`,
  compartido con el Revisor), y
  `CAPACIDADES = frozenset({"fuentes_externas", "filtro_por_valores", "borrar_todo",
  "cols_no_cero", "forzar_valores"})` — **son 5**, el documento de dominio todavía
  lista 4.
- **Depende de:** driver `Microsoft Access Driver (*.mdb, *.accdb)` de la **misma
  arquitectura** (32/64 bits) que el Python que lo ejecuta.
- **⚠️ Al agregar una capacidad hay que sumarla a `CAPACIDADES`.** Si no, el script
  que la use falla con un `TypeError` raro en vez de decir "copiá el archivo
  actualizado"; y si la capacidad es un filtro, no falla nada y entran datos de más.
  Hay un `_verificar_capacidades()` que comprueba al arrancar que lo declarado
  exista de verdad en el código, porque el error ya pasó dos veces en los dos
  sentidos — y declarar de más es lo peor, porque el que importa el motor cree que
  puede usarla.

## `Revisor Reliquidación/actualizadores/Actualiza_Energia.py`

- **Qué hace:** los dos entregables de `01.a Sobrecostos de Energia`, los dos desde
  el `02 Consolidado_Tabulado`, tomando **solo SCMT y SCPC** (columna `AB`).
- **Consume:** `02 Consolidado_Tabulado`, hoja `Sobrecostos`, desde la fila 3.
- **Produce:** `03b ENTRADA_SOB_*.mdb` tabla `Sobrecostos` (origen `AA:AE`, ya en el
  orden de la tabla) y `Consolidado_AAMM`, hoja `Sobrecostos` (origen `A:G` + `I:J`
  + `H` **en ese orden**, a destino `A:J` desde la fila 2; la `H` queda al final).
- **Expone:** —
- **Depende de:** **el motor de `Actualiza_Data_Access.py`**, así que los dos `.py`
  tienen que vivir en la misma carpeta. Comprueba `CAPACIDADES` antes de usarlo.
- **Detalles que importan:** limpia `A2:J` midiendo con el `used_range` antes de
  pegar; sin eso, un mes con menos filas deja las viejas abajo y el total de la
  columna `E` (lo que suma V6) sale inflado. **Caso no cubierto:** el `DELETE` del
  Access borra solo los tipos que vienen en el Excel, así que si un mes desaparece
  SCPC del tabulado, las filas viejas de SCPC quedan; V5 lo caza pero se ve como
  descuadre de monto.

## `Revisor Reliquidación/actualizadores/Actualiza_Cuadro0.py`

- **Qué hace:** deja datos y fórmulas del cuadro cero al día. **No** llama macros ni
  refresca la tabla dinámica: eso se hace a mano.
- **Consume:** `1_CUADROS_PAGO_SSCC` (hoja `01.SSCC_Recurso_Técnico`, `I9:K`), la
  tasa de interés que se le dé, y `config.json` (`tasa_interes_por_mes`).
- **Produce:** el `0_CUADROS_RELIQUIDACIÓN SSCC` — hoja #3 (`A5:C`, `D`, `K5`,
  `L:P`, `R:U`, `O2`) y hoja `01.SSCC_Recurso_Técnico` (`A:G`, `I9`, `J:K`). Guarda
  y deja el libro abierto.
- **Expone:** —
- **Depende de:** que `ActualizaRemplazos.py` ya haya dejado las empresas y los
  reemplazos, porque la `C` es el `BUSCARV` que los aplica.
- **⚠️ El orden de los 8 pasos es la cadena de dependencias, no una preferencia:**
  `A:C → K → L:P y R:U → A:G de 01.SSCC → C → I → J:K`. Por eso la ventana no tiene
  casilleros por paso: correrlos sueltos deja el libro a medio calcular sin señal.
- **Detalles que importan:** cálculo en **manual** de punta a punta con recálculo
  explícito en cuatro puntos; las fórmulas se escriben **en inglés y con coma** por
  COM; la `D` se ajusta contra la tabla `A:C`, no contra `K`; la `Q` queda vacía a
  propósito; la tasa se guarda por mes, no suelta.

## `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py`

- **Qué hace:** reescribe la hoja **"SC y CO"** de la planilla 5, con los SC y los CO
  de los embalses y su prorrata de instrucción directa. SC arriba, CO abajo.
- **Consume:** `Calculo_SobrecostosSSCC` hoja `SOBRECOSTOS` (`S`,`T`,`U`,`V`,`W` y
  `AU:BJ`, desde la fila 7) y `Calculo_CO` hoja `PRORRATA CO` (`B`,`D`,`F:G` y
  `BM:CB`, desde la fila 7).
- **Produce:** planilla 5, hoja "SC y CO", desde la fila 9: `C:G`, `I:X` (prorrata,
  16 columnas) y las fórmulas `Y:AB` y `AD:AF`.
- **Expone:** `CENTRALES_EMBALSE` — 27 unidades, con ANGOSTURA-1/2/3.
- **Depende de:** nada del repo por import.
  **⚠️ Duplica `CENTRALES_EMBALSE` con `Revisor_Reliquidacion.py`** — al tocar la
  lista hay que cambiarla en los dos, con el nombre exacto del origen.
- **Detalles que importan:** **los dos bloques se reescriben siempre**, aunque se
  elija uno solo: el CO va debajo del SC, así que si los SC cambian de largo el CO
  tiene que correrse. El que no se eligió se lee del propio destino por la columna
  `D` (`SCCF` / `CO`). La `AC` queda afuera de las fórmulas: hay que tratar `Y:AB` y
  `AD:AF` por separado, nunca un `Y:AF` corrido. La Clave Año_Mes de los CO sale de
  los SC (no viene en el origen); si no hay SC, la que ya tenía el destino; si
  tampoco, el mes de la ventana.

## `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`

- **Qué hace:** carga el `.mdb` de la planilla 9 desde las planillas 3, 5 y 6.
  Reemplaza a `para ricardo.py` + `archivo_de_configuracion.yaml`. Una casilla por
  planilla; al marcar una se actualizan sus sobrecostos **y** sus propietarios.
- **Consume:** planilla 3 (hoja `DB`, `BC:BG` desde la fila 3; propietarios `M`/`L`),
  planilla 5 (`CÁLCULO_CRA`, `AQ:AU` desde la 9; `EMPRESAS`, `B`/`C`), planilla 6
  (`ENERGIA_Y_CALCULO_CO_ERNC` `AY:BC` desde la 9 y `CALCULO_REA_CENTRAL` `AT:AX`
  desde la 10; `CONSUMOS_PROPIOS`, `B`/`H`).
- **Produce:** `Ocupar_este_para_Reliquidacion_*.mdb`, tablas de sobrecostos y de
  propietarios, **vaciadas completas una sola vez** antes de insertar. Opcionalmente
  `df_ricardo_salida.xlsx` junto al `.mdb`.
- **Expone:** —
- **Depende de:** el motor y la verificación de `Actualiza_Data_Access.py`. Exige
  las **5** capacidades (`fuentes_externas`, `filtro_por_valores`, `borrar_todo`,
  `cols_no_cero`, `forzar_valores`) y si falta alguna corta con un mensaje que dice
  cuál y que hay que copiar el archivo actualizado al lado. `Actualiza_Energia.py`
  solo exige las dos primeras.
- **Detalles que importan:** la `Clave Año_Mes` del origen **no se usa** (viene mal,
  siempre `23xx`): se pisa con el mes sacado del nombre de los archivos. Los bloques
  se leen **por posición, no por nombre de columna** (un `pd.concat` alinea por
  nombre y un encabezado escrito distinto desalineaba el `INSERT`). Transacción única
  con rollback. Modo SOLO MIRAR por defecto. **Una sola fila por central**, comparando
  normalizado: si hay dos empresas distintas gana la primera en el orden 3→5→6 y se
  avisa. Las centrales **sin** propietario se conservan (se cortan por la columna de
  la central, no por la de la empresa); una central sin dueño no es error, lo es una
  central **con monto** sin dueño, y eso lo caza V12. El bloque BESS del script viejo
  no se incluye.

## `Revisor Reliquidación/actualizadores/Prorratear.py`

- **Qué hace:** automatiza en SQL Server lo que se hacía a mano en Management
  Studio. Está como botón en los **tres** `.mdb`.
- **Consume:** el `.mdb` que corresponda (**solo lectura**): tablas
  `Central_Empresa_Actualizada` o `Central_Empresa`, `Sobrecostos`, `TIPOS`; y la
  vista `dbo.[10_Pago_retiros]` de la base de sobrecostos.
- **Produce:** en la base de sobrecostos, `Central_Empresa`, `Sobrecostos`, `TIPOS`
  importadas, y `Pago_Retiro_reporte_tabla` armada con el `SELECT … INTO … GROUP BY`.
- **Expone:** los dos escenarios, que van **de a pares** — Normal (`02_RETIROS` +
  `05_SOBRECOSTOS`) y Reliquidación (`14_RETIROS_RELIQUIDACION` +
  `16_SOBRECOSTOS_RELIQUIDACION`). En la ventana se elige el escenario, no las bases
  sueltas, porque elegir mal significa prorratear contra los retiros equivocados.
- **Depende de:** que `Carga_Retiros.py` haya cargado los retiros del mes;
  `ruta_nodo` del JSON de traspaso para saber a cuál de los tres `.mdb` le dieron.
- **Detalles que importan:** **lo primero que comprueba es el mes** — prorratear
  contra los retiros de otro mes no da error, da un resultado mal que se ve
  plausible. Corta si el mes del Access no está entre los retiros, si el Access
  tiene más de un mes, o si `Retiros` está vacía. Comprueba que la vista y las tres
  tablas existan **antes de borrar nada**. Modo SOLO MIRAR por defecto. Si falla
  después del borrado la base queda incompleta: hay que volver a correrlo.

## `Revisor Reliquidación/actualizadores/Carga_Retiros.py`

- **Qué hace:** carga `Retiros_h.parquet` a SQL Server. Es el único script que
  escribe en una base y no en un Excel.
- **Consume:** `04 Planilla 9/Retiros_h.parquet`; `config.json` (`retiros_base`).
- **Produce:** tabla `Retiros` en `SRV-DTE`, base `02_RETIROS` o
  `14_RETIROS_RELIQUIDACION` (**sin tilde** en el código; el documento la escribe con
  tilde en un lado y sin tilde en otro). Driver `ODBC Driver 17 for SQL Server`,
  conexión de confianza. **Siempre `TRUNCATE`** y carga de cero, en trozos de
  `CHUNK` = 50.000, verificando que la cuenta cuadre. La hora del cambio por omisión
  es la 145.
- **Expone:** la casilla «Cambio de hora (−)» y la hora del mes.
- **Depende de:** la ruta del parquet, que puede venir del revisor.
- **Detalles que importan:** los nombres de columna no se escriben igual en todos
  lados (`Clave Año_Mes` / `Clave_Anio_Mes` / `Clave_anio_mes`): se resuelven
  comparando normalizado y tratando `ANIO` como `ANO`, y **se renombran las del
  parquet a como se llaman en la tabla** antes de cargar, porque `to_sql` mapea por
  nombre. Si el parquet trae una columna que la tabla no tiene, aborta antes de
  tocar la base. El cambio de hora empuja una hora todo lo que esté desde la hora que
  no existe; **el parquet no se modifica**. No la aplica si el archivo ya viene
  corrido (la hora del cambio no existe en los datos) ni si el mes está completo.

## `Revisor Reliquidación/Reemplazos REUC/ActualizaRemplazos.py`

- **Qué hace:** arma la tabla de empresas con RUT y la de reemplazos, y las escribe
  en la hoja `EMPRESAS` del cuadro cero. Hace bastante más de lo que dice el
  documento de dominio: además descarga los datos del REUC solo y busca en el disco
  T: los dos archivos grandes que necesita.
- **Consume:**
  - `datos_reuc_*.xlsx` y `datos_reuc_reemplazos_*.xlsx` — **los descarga solo**
    desde `reuc.coordinador.cl` con **playwright** (pide `pip install playwright` y
    `playwright install chromium`); si no está, se pueden bajar a mano.
  - `Reemplazos forzados*.xlsx`, hoja `Reemplazos forzados`, columnas
    `Reemplazada` / `Reemplazante`. Se copia del último mes reliquidado y se
    confirma.
  - `Cuadros de Pago_Balances_SEN_*_Simplificado_def.xlsb` — de PLABACOM, buscado
    bajo `T:\Facturacion\Plabacom` con el patrón `balances_sen.*simplificado.*\.xlsb$`.
  - `1_CUADROS_PAGO_SSCC_*.xlsm` — buscado bajo `T:\Facturacion`, con búsqueda
    conjunta por mes (`MESES_ES`, `MESES_ABR`). **No** es el de `00 Entregables` que
    tiene el revisor; sigue sin confirmarse si es el mismo archivo, y por eso el
    revisor manda `cuadro_1` en el JSON pero el script no la usa.
- **Produce:** `Reemplazos_AAAAMMDD_SSCC.xlsx` (el entregable suelto) y, en la hoja
  `EMPRESAS` del cuadro cero, `B:C` = EMPRESA / RUT y `H:I` = Reemplazada /
  Reemplazante. **Los dos bloques se limpian antes de pegar, conservando el formato
  de las celdas.**
- **Expone:** —
- **Depende de:** `pandas`, `xlwings` y `playwright` (este último opcional). Su
  **propio** `config.json`, en `Reemplazos REUC/Auxiliares/`, no el compartido.
- **⚠️ Orden corregido:** los reemplazos forzados se aplican **antes** del cruce con
  el registro de empresas. Aplicándolos después, las filas sin RUT ya se habían
  descartado y un forzado nunca podía arreglar el nombre que no coincidía: la alerta
  volvía todos los meses y la empresa desaparecía del entregable en silencio (3
  empresas por $6.000 salían como $3.000). El mapa se aplica dos veces, así que un
  `Reemplazante` que sea a la vez `Reemplazada` salta dos escalones; el script lo
  avisa en el log.

---

## Diferencias con el documento de dominio

`docs/ESTRUCTURA_CASO_RELIQUIDACION.md` se quedó atrás en estos puntos. **Manda el
código.** Corregir el documento cuando haya oportunidad:

| Dónde | Dice el documento | Dice el código |
|---|---|---|
| V16, comprobación 6 | `UMBRAL_DESCUADRE_CPRT` = 500 | **1000.0** — es el redondeo acumulado del reparto proporcional, que crece con la cantidad de empresas; en 2312, con 481 pares, dio 31,3 pesos |
| `Actualiza_Data_Access.py` | `CAPACIDADES` tiene 4 | **5**: falta listar `forzar_valores`. Y hay un `_verificar_capacidades()` que el documento no menciona |
| Caché de directorios | "27 nodos" | el árbol tiene **36** nodos (`NODOS`) |
| `Carga_Retiros.py` | `14_RETIROS_RELIQUIDACIÓN` con tilde | **`14_RETIROS_RELIQUIDACION`** sin tilde, igual que en `Prorratear.py` |
| `ActualizaRemplazos.py` | solo pega empresas y reemplazos | además **descarga del REUC con playwright**, busca los Balances SEN de PLABACOM y el `1_CUADROS_PAGO` en `T:\Facturacion`, y genera `Reemplazos_AAAAMMDD_SSCC.xlsx`. Escribe en la hoja `EMPRESAS`, no en una hoja sin nombrar |

---

## El módulo común

### `Revisor Reliquidación/comun/config.py` — **hecho**

- **Qué hace:** lee y escribe el `config.json`, indexado por `<equipo>_<usuario>`.
- **Expone:** `clave_equipo()`, `leer(ruta)`, `leer_todo(ruta)`,
  `guardar(ruta, data)`, `modificar(ruta, mutador)`, `escribir_json(ruta, data)`.
- **Reglas que no cambian:** solo se agregan o actualizan claves, nunca se borra
  nada ajeno; si el archivo existe pero no se puede interpretar **no se escribe**;
  la escritura es atómica (`.tmp` + `os.replace`).
- **Pruebas:** `Revisor Reliquidación/comun/test_config.py`, 13 casos, solo stdlib.
- **Dos bugs que se arreglaron al juntar las copias:**
  1. `ActualizaRemplazos.py` escribía el archivo entero **sin `.tmp`**, y si el
     config existía pero estaba roto lo pisaba con `{}`. O sea: un `config.json`
     ilegible le borraba los ajustes a los otros nueve scripts.
  2. `get_usuario` tenía versiones con y sin `try/except`. Quedó la defensiva.
- **Cómo migrar el resto:** reemplazar las cinco funciones del script por
  envoltorios de dos líneas sobre `_cfg`, y agregar `from comun import config as
  _cfg`. Los nombres viejos se conservan, así que **ningún punto de llamada
  cambia** — y `_modificar_config` se llama directo en los 10 scripts, con
  mutadores propios (la tasa por mes de `Actualiza_Cuadro0`, por ejemplo).

| Script | Estado |
|---|---|
| `actualizadores/Actualiza_SC_CO.py` | migrado — 42 líneas menos |
| los otros 9 | pendientes |

### Lo que sigue

El generador detectó **15 constantes definidas
en más de un archivo** (la tabla completa está al final de `INTERFACES.md`). Ordenadas
por cuánto rinde sacarlas:

| Constante | En cuántos archivos | Qué es |
|---|---|---|
| `CONFIG_PATH` | **10 de 10** | la ruta del `config.json` |
| `TRASPASO_ORIGEN`, `TRASPASO_VERSION_MAX` | 9 | el contrato del JSON del revisor |
| `DIR_SCRIPT` | 7 | la carpeta del script |
| `NS_XL`, `NS_REL`, `_ENT_XML`, `_RE_ENT` | 2 | el lector de Excel por ZIP/XML, duplicado entre el Revisor y `Actualiza_Data_Access` |
| `CENTRALES_EMBALSE` | 2 | el caso documentado, el único que ya obligó a defenderse con una verificación |
| `SERVER`, `DRIVER`, `CHUNK`, `LARGO_TEXTO` | 2 | la conexión a SQL Server |

**El orden de migración no es `CENTRALES_EMBALSE`**, aunque sea el caso más famoso:
es el que menos riesgo tiene de quedar mal, porque V10 lo caza en los dos sentidos.
La próxima pieza, por tamaño de la ganancia, es el **lector de Excel por ZIP/XML**
(`NS_XL`, `NS_REL`, `_ENT_XML`, `_RE_ENT` y sus funciones): son varios cientos de
líneas idénticas entre el Revisor y `Actualiza_Data_Access`.

**Este cambio toca todos los scripts a la vez, así que va por partes.** Una pieza a
la vez, un script a la vez, verificando que sigue corriendo antes de seguir con el
próximo.
