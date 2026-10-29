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
scripts/                           utilidades de sesión (sincronizar, verificar)
__comun__/                        código compartido y sus pruebas
__config__/                        JSON, estados y datos intermedios locales
Comparadores/                      herramientas anuales, hermanas del Revisor
├── Comparador_Etapas.py
└── Comparador_Tabulado.py
00_Salidas/                        salidas compartidas por año y mes
Revisor_Relq/
├── Revisor_Reliquidacion.py     ← el único que va suelto en la raíz
├── revisor/                       componentes internos (no son ejecutables)
│   ├── archivos.py              búsqueda y caché temporal de directorios
│   └── estado.py                estado mensual y caché persistente de valores
├── actualizadores/                los 8 que el Revisor lanza por botón
│   ├── Actualiza_datos.py
│   ├── Actualiza_Data_Access.py
│   ├── Actualiza_Energia.py
│   ├── Actualiza_Cuadro0.py
│   ├── Actualiza_SC_CO.py
│   ├── Actualiza_Access_P9.py
│   ├── Carga_Retiros.py
│   └── Prorratear.py
└── Reemplazos REUC/               usa __config__/reemplazos_reuc.json
                                   y __config__/Auxiliares REUC/
    └── ActualizaRemplazos.py
Revisor_Relq_15min/                la versión 15 minutos (ver su sección, al final)
├── Revisor_Reliquidacion_15min.py
├── revisor/
├── actualizadores/                Actualiza_datos, Actualiza_Cuadro0,
│                                  Actualiza_SC_CO
└── Reemplazos REUC/
```

**Dos versiones, una raíz.** `Revisor_Relq/` es la reliquidación **horaria** y
`Revisor_Relq_15min/` la de **bloques de 15 minutos**. Comparten `__comun__/`,
`__config__/`, `00_Salidas/` y `Comparadores/`. Todo lo que sigue hasta "El
módulo común" describe la horaria; la de 15 minutos tiene su propia sección al
final de este archivo, con lo que cambia respecto de la horaria.

`__config__/config.json` es **compartido** entre el Revisor, los 8
actualizadores y los comparadores. El noveno usa `__config__/reemplazos_reuc.json`
y deja sus insumos descargados en `__config__/Auxiliares REUC/`.
Los JSON mensuales replican `AAAA/MM Mes`; estado, rutas, parquet y vistas de los
comparadores van en `__config__/AAAA/_comparador*`. `00_Salidas/` queda reservado
para Excel y demás resultados entregables.

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
                            pasándoles __config__/AAAA/MM Mes/_traspaso_actualizador.json
```

---

## `Revisor_Relq/Revisor_Reliquidacion.py`

- **Qué hace:** es la consola del proceso. Arma el árbol de archivos del mes,
  verifica por fecha de modificación (copias contra sus maestros) y por valores
  (V4…V17, con cadena de dependencias entre verificaciones), y desde cada fila lanza
  el actualizador que corresponde.
- **Consume:** todo el árbol de `02 CASO RELIQUIDACION` + `FD/`, las tres bases
  Access, `config.json` (`carpeta_base`, `ultimo_mes`, `carpetas_por_mes`,
  `_valores`).
- **Produce:** `00_Salidas/AAAA/MM Mes/`, en la carpeta hermana de `Revisor_Relq/` —
  estado, caché y `_traspaso_actualizador.json`, que
  le pasa como único argumento al actualizador que lanza.
- **Expone:** el JSON de traspaso (`origen`, `version`, `aamm`, `carpeta_reliq`,
  `planilla`, `rutas`, `nodo`, `clave_nodo`, `ruta_nodo`); `plan_traer_maestro()`
  separado del ejecutor; `cache_directorios()` como context manager;
  `CENTRALES_EMBALSE` (27); `UMBRAL_DESCUADRE_CPRT` (**1000.0**, no 500) y
  `UMBRAL_PAR_SEGURO` (100.0); `TOLERANCIA` (1.0), `TOL_PAGO_EMPRESA` (150.0),
  `TOL_SOBRECOSTO_FILA` (1.0), `TOL_MTIME` (2 s); `NODOS` (36) y `VERIFICADORES` (14).
- **Depende de:** los 9 scripts que lanza, pero solo por línea de comandos — no los
  importa. **Duplica `CENTRALES_EMBALSE` con `Actualiza_SC_CO.py`.**
- **División interna:** `revisor/archivos.py` contiene la búsqueda y el caché
  acotado de directorios; `revisor/estado.py` contiene estado mensual y caché
  persistente, con sus dependencias inyectadas. El archivo principal conserva
  todos los nombres históricos mediante imports o wrappers breves.
- **Detalles que importan:** V17 es la raíz de casi todo (si el tabulado tiene el
  sobrecosto mal calculado, todos los totales cuadran igual y el error llega al
  cuadro de pago). El caché de directorios está apagado por omisión y solo se
  enciende dentro del `with`, porque `mtime()` se usa después para decidir si una
  verificación venció. La relectura es parcial: V9 relee 1 nodo, V16 relee 13 (el
  árbol tiene 36 nodos). `cache_directorios` es una clase, no una función: al salir
  deja `self.stats = (scans, hits)` para poder decir en la bitácora cuánto se ahorró.

## `Revisor_Relq/actualizadores/Actualiza_datos.py`

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

## `Revisor_Relq/actualizadores/Actualiza_Data_Access.py`

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

## `Revisor_Relq/actualizadores/Actualiza_Energia.py`

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

## `Revisor_Relq/actualizadores/Actualiza_Cuadro0.py`

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

## `Revisor_Relq/actualizadores/Actualiza_SC_CO.py`

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

## `Revisor_Relq/actualizadores/Actualiza_Access_P9.py`

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

## `Revisor_Relq/actualizadores/Prorratear.py`

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

## `Revisor_Relq/actualizadores/Carga_Retiros.py`

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

## `Comparadores/Comparador_Etapas.py`

- **Qué hace:** consolida y compara por central y hora los sobrecostos de los dos
  Access de las etapas Definitivo, Rpre y Rdef para los doce meses de un año.
- **Consume:** los `.mdb` de cada etapa, el JSON mensual de traspaso del Revisor,
  el `config.json` compartido y el maestro vigente de central/empresa.
- **Produce:** parquet, vistas, estado y rutas en
  `__config__/AAAA/_comparador/`; Excel mensual en `00_Salidas/AAAA/MM Mes/` y
  consolidado anual directo en `00_Salidas/AAAA/` (sin subcarpeta — ahí no
  queda nada más que los `.xlsx`).
- **Expone:** `cdir(anio)`, `ruta_json_mes(aamm)`, `path_excel_mes(aamm)` y el
  bloque propio `comparador_etapas` del JSON mensual.
- **Depende de:** `__comun__/salidas.py`; localiza la carpeta del Revisor
  por `Revisor_Reliquidacion.py`, no por su nombre.
- **Detalles que importan:** el hilo de trabajo comunica log, estado, progreso y
  repintado mediante una cola; solo el hilo principal toca tkinter. **`pintar()`
  no toca el disco**: lee `self.instant`, la instantánea que `instantanea_mes()`
  calcula en el hilo de fondo. Sin eso cada click repetía más de cien `is_file()`
  contra el NAS y marcar una casilla tardaba lo mismo que la búsqueda entera.

## `Comparadores/Comparador_Tabulado.py`

- **Qué hace:** compara sobrecosto, generación, CV, CMg y USD entre Definitivo,
  Rpre y Rdef, por central y hora, desde los Consolidados Tabulados de un año.
- **Consume:** Consolidados Tabulados, JSON mensuales, el `config.json` compartido
  y `__config__/AAAA/_comparador/rutas.json` del comparador de Access.
- **Produce:** parquet de variables, vistas, estado y rutas en
  `__config__/AAAA/_comparador_tabulado/`; Excel mensual en
  `00_Salidas/AAAA/MM Mes/` y consolidado anual directo en `00_Salidas/AAAA/`
  (sin subcarpeta).
- **Expone:** `cdir(anio)`, `dir_parquet(anio)`, `ruta_json_mes(aamm)` y
  `path_excel_mes(aamm)`.
- **Depende de:** `__comun__/salidas.py` y, para no pedir las mismas
  carpetas dos veces, del `rutas.json` de `_comparador` **del mismo año**.
- **Detalles que importan:** el hilo de trabajo se comunica con tkinter mediante
  una cola; un detalle que supera el límite de Excel continúa en hojas `_2`,
  `_3`, etc. Antes de calcular `hora_mes` se advierte si faltan días u horas o si
  el total mensual no corresponde a un mes completo (incluido cambio de hora).
  Cada fila del Excel lleva `Fecha` y `Hora Dia` además de `Hora Mensual`, para
  ubicar la diferencia en el calendario sin tener que contar horas; las tres
  columnas van juntas al principio y quedan inmovilizadas. Un parquet escrito por
  una versión anterior (sin `fecha` ni `hora_dia`) se marca **desactualizado** y
  se reconsolida solo.
  **`pintar()` no toca el disco**: lee `self.instant`, la instantánea que
  `instantanea_mes()` calcula en el hilo de fondo; y el esquema de cada parquet
  se recuerda por firma del archivo (`__comun__/comparadores.py`,
  `CachePorArchivo`), en vez de releerse 36 veces por repintado.

## `Revisor_Relq/Reemplazos REUC/ActualizaRemplazos.py`

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
  **propio** `reemplazos_reuc.json`, en `__config__/`, no el compartido.
- **Correo del REUC:** se guarda en `reemplazos_reuc.json` (clave `correo_reuc`,
  por PC+usuario) y se escribe solo en el campo del login cuando se abre el
  navegador, sin pisar lo que la persona haya empezado a escribir; si lo escribe
  ella, queda recordado para la próxima. **La clave nunca se guarda ni pasa por
  el script:** se escribe en la página real del REUC, dentro del navegador, que
  además sigue abriéndose sin perfil ni cookies persistidas.
- **Carpeta de auxiliares:** `__config__/Auxiliares REUC/`, compartida por todos
  los usuarios: ahí caen los `datos_reuc_*` descargados y ahí se busca
  `Reemplazos forzados*`. Antes estaba al lado del `.py`
  (`Reemplazos REUC/Auxiliares/`); esa carpeta ya no se lee ni se escribe, solo
  se mira para **avisar** que quedaron archivos por mover — el script no los
  mueve solo, igual que con el resto de `__config__`. Lo único que se corrige
  solo es la ruta guardada en el config si apuntaba a la carpeta vieja.
- **⚠️ Orden corregido:** los reemplazos forzados se aplican **antes** del cruce con
  el registro de empresas. Aplicándolos después, las filas sin RUT ya se habían
  descartado y un forzado nunca podía arreglar el nombre que no coincidía: la alerta
  volvía todos los meses y la empresa desaparecía del entregable en silencio (3
  empresas por $6.000 salían como $3.000). El mapa se aplica dos veces, así que un
  `Reemplazante` que sea a la vez `Reemplazada` salta dos escalones; el script lo
  avisa en el log.

---

## Versión 15 minutos — `Revisor_Relq_15min/`

Nació como copia de `Revisor_Relq/` (2026-10-22) y se adapta de a poco, a
medida que la usuaria necesita cada actualizador. **Lo que no se describe acá
es igual que en la horaria** (Cuadro0, ActualizaRemplazos
y el paquete `revisor/`, salvo lo indicado).

- **Lo que no existe en 15 minutos y se sacó:** detalles diarios, CCA, 02
  Consolidado_Tabulado y las tres bases Access, con sus nodos, las
  verificaciones V4, V5, V6, V12 y V17, y los actualizadores
  `Actualiza_Data_Access`, `Actualiza_Energia`, `Actualiza_Access_P9` y
  `Prorratear`.
- **Config propio:** los cinco scripts de esta versión leen y escriben
  `config.json` y `reemplazos_reuc.json` **solo dentro de la clave
  `"modo_15min"`**, con la misma forma que el archivo entero. Lo de afuera es
  de la horaria. Está en los envoltorios de cada script (`SECCION_CONFIG`), no
  en `__comun__/`. Lo comprueba `Revisor_Relq_15min/revisor/test_seccion_config.py`.
- **Carpetas por mes compartidas:** `__config__/AAAA/MM Mes` y
  `00_Salidas/AAAA/MM Mes` son las mismas para las dos versiones; no chocan
  porque un mes es de un solo modo.
- **Comparadores:** siguen siendo de la horaria y buscan
  `Revisor_Reliquidacion.py`, o sea `Revisor_Relq/`.
- **Tampoco se usan en 15 min** (usuaria, 2026-10-23): la carpeta `01.a
  Sobrecostos de Energia` (Consolidado_AAMM y Pago_Sobrecostos, con su V7),
  `00 Entregables/02 Costo de Oportunidad` y el `Retiros_h.parquet` (con
  `Carga_Retiros.py`, que se borró). El árbol queda en 20 nodos.
- **Pendiente** (ver `BITACORA.md`): qué reemplaza el "Traer Consolidado", el
  `Prorratear` y las comprobaciones contra los `.mdb`.

### `Revisor_Relq_15min/actualizadores/Actualiza_SC_CO.py` — hoja "SC y CO" en 15 min

- **Cómo se usa (2026-10-28):** ya **no es un botón aparte** del Revisor. Es
  una casilla más de la planilla 5 en la ventana de `Actualiza_datos.py`
  («Actualizar SC» y «Actualizar CO», junto con FD y prorrata), que importa
  este archivo y llama a **`actualizar_en_libro(app, wb_d, rutas, hacer,
  aamm_respaldo, log, progreso)`** sobre el mismo libro abierto, después de FD
  y prorrata. Esa función no cambia el modo de cálculo, no recalcula ni
  guarda. `ejecutar()` y la ventana propia siguen funcionando solos (abren,
  llaman a `actualizar_en_libro`, recalculan y guardan). **Los dos `.py`
  tienen que estar en la misma carpeta.**

- **Metodología:** la del CRA de `Balance_SSCC` (trazabilidad v5, §6.6),
  confirmada por la usuaria el 2026-10-23. SC arriba, CO abajo, desde la fila
  9, solo embalses.
- **SC** — `Cálculo_SobrecostosSSCC`, hoja `SOBRECOSTOS`, fila 7: `S,T,U → C,D,E`;
  `F = DIA(A) & "#" & (R − (DIA(A)−1)×96)` (Clave_Bloque); `W → G`;
  **`AW:BN → I:Z`** (las 18 prorratas ID / Pmax / Real, tal cual) y en
  **`AA:AC` la suma calculada** CPF/CSF/CTF. Cada grupo de 6 viene en el orden
  CPF+, CPF−, CSF+, **CTF+**, CSF−, CTF−, así que CPF = I,J,O,P,U,V,
  CSF = K,M,Q,S,W,Y y CTF = L,N,R,T,X,Z (igual que las fórmulas del libro).
- **CO** — **`Reporte_CRA_15min_AAMM.csv`** (ya no `Calculo_CO`), en
  **`04 Planilla 9/`**, fila 2: `A,B,H → C,D,E`;
  `F = DIA(C) & "#" & E`; `I → G`; `I:Z` vacío; `AA:AC` = `J+K`, `L+M`, `N+O`.
  Se lee con `__comun__/csv_columnas.py`.
- **Fórmulas que se estiran:** `AD`, `AF:AG`, `AI` (`BLOQUES_FORMULA`). AE y AH
  no llevan fórmula y la tabla `dict_SCCO` (AT:AV) no se toca. Antes de
  escribir avisa si la fila 9 no calza con esa partición. V10 usa la misma.
- **Avisos:** más de una Clave Año_Mes, Tipo + Unidad + Clave_Bloque repetidos
  (control AI del libro), y bloques de SC fuera de 1..96 (mes con cambio de
  hora: el 96 es el del Excel y se copia tal cual).
- **Revisor 15 min:** el nodo `a_calc_co` pasó a **`a_reporte_cra`** (clave de
  traspaso `reporte_cra`). V10 compara el CO contra el CSV, revisa que
  `AA:AC` sume 1 por fila y que `AD`, `AF:AG`, `AI` cubran las filas.
  `revisor/lectores.leer_columnas_rapido` despacha los `.csv` a
  `__comun__/csv_columnas.py`.

### `Revisor_Relq_15min/actualizadores/Actualiza_Pagos_Varios.py` — pagos_varios (2026-10-28)

- **Qué hace:** reemplaza, en 15 minutos, a `Actualiza_Access_P9.py`: los mismos
  cuatro bloques de las planillas 3, 5 y 6 ya no van al Access
  `Ocupar_este...`, sino a **`04 Planilla 9/pagos_varios_AAMM.csv`**
  (`Fecha_Hora,CONFIGURACION,Central,Pago,Tipo_pago,Bloque_15min`). Sin
  Central_Empresa.
- **Consume** (columnas dadas por la usuaria; el encabezado se busca en las 20
  primeras filas por un "Pago" y los datos empiezan en la siguiente):

  | Fuente | Hoja | Fecha | Central | Pago | Tipo | Bloque |
  |---|---|---|---|---|---|---|
  | Subastas (3) | `DB` | — | BF | BH | BE | BI |
  | CRA (5) | `CÁLCULO_CRA` | AV | AX | AY | AZ | BA |
  | CO_ERNC (6) | `ENERGIA_Y_CALCULO_CO_ERNC` | BB | BD | BE | BF | BG |
  | REA (6) | `CALCULO_REA_CENTRAL` | AT | AV | AW | AX | AY |

  La planilla 3 no trae fecha (BG es la hora mensual): se arma con el bloque y
  el mes (AAMM del nombre de los archivos o de la ventana).
- **Reglas:** Fecha_Hora siempre `DD/MM/AAAA HH:MM`; CONFIGURACION vacía; fuera
  los pagos 0 y las centrales vacías; orden Subastas, CRA, CO_ERNC, REA; coma,
  punto decimal, UTF-8 sin BOM, CRLF; escritura `.tmp` + `os.replace`. Avisa si
  un bloque no calza con su fecha (`(día−1)×96 + hora×4 + min/15 + 1`: cambio
  de hora) o si una fecha no es del mes. Solo lee: ZIP/XML, con xlwings de
  respaldo.
- **Revisor 15 min:** nodo `a_pagos_varios` (`crea_salida=True`): en rojo si no
  existe, pero con el botón «Actualizar data» activo, que es el que lo crea.
  `revisor/lanzamiento.py` deja lanzar sin destino cuando el nodo es
  `crea_salida`.
- **Verificador V12** (botones «Verificar» y «detalle» en la fila del csv):
  cada fuente del csv contra el total de su RESUMEN — Subastas (los seis
  CPF/CSF/CTF ±) = planilla 3 `U7`, CRA = planilla 5 `D7`, CO_ERNC =
  planilla 6 `L7`, REA = planilla 6 `E7` (los mismos totales con que el V12
  horario cuadraba el `Ocupar_este.mdb`). Previas V9, V10, V11. Las sumas son
  `excel_col` sobre el csv (`PV_*`, hoja `(csv)`), que
  `lectores.leer_columna_excel` lee con `__comun__/csv_columnas.py`.
- **Validado** contra el `pagos_varios_2609.csv` de la usuaria (planillas
  sintéticas con esos datos): mismas 40.248 filas, orden, centrales, tipos,
  bloques y totales por tipo. Pruebas: `actualizadores/test_pagos_varios.py`.

### `Revisor_Relq_15min/actualizadores/Actualiza_Cuadro1.py` — planilla 1 (2026-10-29)

- **Qué hace:** arma la hoja `01.SSCC_Recurso_Técnico` del `1_CUADROS_PAGO`
  apilando los pagos por empresa de la **planilla 4** (todas las hojas
  `PAGO_EMPRESAS_*`: tipo en `B1`, combinada `B1:D1`; tabla `A3:D` Empresa,
  RECIBE, PAGA, Paga+Recibe, sin la fila TOTAL, con las filas en 0) y, si se
  elige, de la **planilla 8** (`8_REMUNERACIÓN_CONTROL_CONTINGENCIAS`, hoja
  `RESUMEN` desde la fila 9: C tipo, D empresa, E recibe, F paga, G neto).
- **Destino** desde la fila 9: `B` tipo, `C:F` empresa / recibe / paga / SEN,
  **sin las filas con RECIBE y PAGA en 0**; `A` y `G` fórmula, estiradas o
  cortadas; **`I:K`** empresa / paga / recibe (suma de todos los tipos) con
  tres fórmulas de matriz dinámica en `I9`, `J9`, `K9` (`UNIQUE`/`FILTER`/
  `SUMIF` sobre `I9#`, escritas con `Formula2`), sin vacíos ni empresas con
  total 0, y rangos `$C$9:$C$<última>` (antes fijos a 4346); se limpia todo
  `I9:K` hacia abajo para que derramen. `F3:G5` también al largo de la tabla.
- **4 y 8 por separado:** ventana chica con dos casillas. La que no se marca
  se conserva de la planilla 1, reconociendo sus filas por el tipo (columna B)
  contra los tipos de la planilla 4 (sus `B1`) y de la 8 (su columna C),
  comparando con `clave_mayusculas` (`CO ERNC` = `CO_ERNC`). Sin el archivo
  de la 8, es de la 8 todo lo que no sea de la 4. Filas de ninguna de las dos
  (formato viejo) se descartan y se avisa. Orden: la 4 y después la 8. La
  planilla 4 hace falta siempre (de ahí salen los tipos que se conservan).
- **Lectura** de las planillas 4 y 8 sin Excel (ZIP/XML; la 4 real de 27 MB en
  0,2 s); escritura de la 1 con xlwings, cálculo manual, recalcula y guarda.
- **Revisor 15 min:** botón «Actualizar data» en la fila de la planilla 1;
  nodo nuevo `a_8_cc` (`00 Entregables/8_REMUNERACIÓN_CONTROL_CONTINGENCIAS*`);
  claves de traspaso `p4`, `p8`, `cuadro_1`.
- **Pruebas:** `actualizadores/test_cuadro1.py`. Probado además contra los
  archivos reales de 2601 que mandó la usuaria (no versionados), con la
  escritura simulada.

### Revisor 15 min: planilla 9, V15 y «Traer maestro» (2026-10-29)

- **Planilla 9** = `00 Entregables/9_Pagos_Retiros_AAMM_*.csv` (nodo
  `a_9_pagos`, `.csv`). **V13** (total A/B = D/E de su hoja VERIFICADORES) se
  sacó. `P9_TOTAL` = suma de J del csv.
- **V15**: pago por suministrador y tipo de pago, planilla 9 (csv, fila 2:
  E suministrador, H tipo, J pago) contra la planilla 1 (B tipo, C empresa,
  E PAGA), en magnitud y con `TOL_PAGO_EMPRESA`; **EDAC-BF fuera**
  (`excluir_conceptos`, nuevo en `pago_por_empresa`). Previas V14 y V12.
- **«Traer maestro»**: si la copia no existe copia sin preguntar ni avisar; si
  pisa o borra algo, una pregunta corta («X ya existe. ¿Seguro que desea
  reemplazarlo?»); si sale bien no muestra aviso (queda en la bitácora).

### `Revisor_Relq_15min/Revisor_Reliquidacion_15min.py`

- **Qué hace:** es la consola del proceso. Arma el árbol de archivos del mes,
  verifica por fecha de modificación (copias contra sus maestros) y por valores
  (V8…V16, con cadena de dependencias entre verificaciones), y desde cada fila lanza
  el actualizador que corresponde.
- **Consume:** el árbol de `02 CASO RELIQUIDACION` + `FD/` (solo Excel y el
  parquet de retiros; ya no hay `.mdb`), `config.json` (`carpeta_base`,
  `ultimo_mes`, `carpetas_por_mes`, `_valores`).
- **Produce:** `__config__/AAAA/MM Mes/` — estado, caché y
  `_traspaso_actualizador.json`, que le pasa como único argumento al
  actualizador que lanza. El CSV del CPRT va a `00_Salidas/`.
- **Expone:** el JSON de traspaso (`origen`, `version`, `aamm`, `carpeta_reliq`,
  `planilla`, `rutas`, `nodo`, `clave_nodo`, `ruta_nodo`; el `origen` sigue
  siendo `"Revisor_Reliquidacion"`, es el contrato de `__comun__/traspaso.py`);
  `plan_traer_maestro(nodo, ruta_copia, ruta_maestro, carpeta_copia)` separado
  del ejecutor; `cache_directorios()` como context manager; `CENTRALES_EMBALSE`
  (27); `UMBRAL_DESCUADRE_CPRT` (**1000.0**, no 500) y `UMBRAL_PAR_SEGURO`
  (100.0); `TOLERANCIA` (1.0), `TOL_PAGO_EMPRESA` (150.0), `TOL_MTIME` (2 s);
  `NODOS` (**22**) y `VERIFICADORES` (**8**: V8, V9, V10, V11, V12, V14,
  V15, V16).
- **Depende de:** los 5 scripts que lanza, pero solo por línea de comandos — no
  los importa. **Duplica `CENTRALES_EMBALSE` con `Actualiza_SC_CO.py`.**
- **División interna:** `revisor/archivos.py` contiene la búsqueda y el caché
  acotado de directorios; `revisor/estado.py` contiene estado mensual y caché
  persistente, con sus dependencias inyectadas. El archivo principal conserva
  todos los nombres históricos mediante imports o wrappers breves.
- **Lo que se sacó en el modo 15 minutos:** el tipo de nodo `diarios` (y con él
  `listar_diarios`, el botón «detalle» y la detección del mes por las fechas
  AAAAMMDD), el tipo de valor `mdb` y sus lectores Access, y los tipos de
  comprobación `sobrecosto_por_fila` (era V17) y `centrales_sin_dueno` (era
  V12). V7 quedó solo con `M5 = 0`; V8 sin el `EE6 = SCAGC` del tabulado; V13
  sin la suma del `Ocupar_este.mdb` y colgando de V9/V10/V11; V14 sin la suma
  del `03b ENTRADA_SOB_SSCC.mdb` y colgando de V8.
- **Detalles que importan:** la cadena del cuadro cero es ahora
  `V16 -> V15 -> V14 -> V8` (y `V15 -> V13 -> V9/V10/V11`). El caché de
  directorios está apagado por omisión y solo se enciende dentro del `with`,
  porque `mtime()` se usa después para decidir si una verificación venció.
  `cache_directorios` es una clase, no una función: al salir deja
  `self.stats = (scans, hits)` para poder decir en la bitácora cuánto se ahorró.

### `Revisor_Relq_15min/actualizadores/Actualiza_datos.py`

- **Qué hace:** trae las hojas FD desde `SSCC_Desempeno*` y la prorrata de
  retiros, a `Cálculo_SobrecostosSSCC` y a las planillas 3, 5 y 6. Dos opciones:
  "Actualizar FD" y "Actualizar Prorrata" (esta última solo en 3/5/6). El "Traer
  Consolidado" del modo horario se sacó junto con el tabulado.
- **Consume:** `FD/SSCC_Desempeno*.xlsx|xlsm` (hojas `CT Diario`, `CPF Horario`,
  `CSF Horario`, `CTF Horario`, desde la fila 12);
  `04 Planilla 9/Prorrata_Retiros_*.xlsx`, hoja
  `PRORRATA_HORARIA_TABULAR`.
- **Produce:** escribe en `Cálculo_SobrecostosSSCC` (`FD_CT`, `FD_CPF`, `FD_CSF`,
  `FD_CTF`, `FD_CSF_Disponibilidad`, `FD_CTF_Disponibilidad`) y en las
  planillas 3/5/6 (`CPF_FD`, `CSF_FD`, `CTF_FD`, `FD_CPF`, `FD_CSF`, `FD_CTF`, `FD`,
  `PRORRATA_RETIROS`).
- **Expone:** el parámetro `planilla` con valores `sc`, `p3`, `p5`, `p6`; es como lo
  llama el revisor desde cada fila.
- **Depende de:** `config.json` (clave `carpeta_reliq`).
- **Planilla 3, FD en 15 minutos:** en `CPF_FD`, `CSF_FD` y `CTF_FD` el
  bloque pegado empieza en la **F** (`F:N`, `F:L`, `F:M`; antes `D:L`, `E:K`,
  `E:L`). El origen se lee igual. Las columnas de fórmula y del rótulo fijo
  (`CPF`/`CSF`/`CTF`) solo se limpian abajo, no se extienden: la planilla
  tiene que traerlas hasta donde lleguen los datos. Las planillas 5 y 6
  siguen con los mapeos de siempre (confirmado por la usuaria); el Cálculo
  SSCC todavía no se revisó.
- **Planilla 5, hoja "SC y CO" (2026-10-28):** casillas «Actualizar SC» y
  «Actualizar CO» con sus orígenes (`Cálculo_SobrecostosSSCC` maestro y
  `04 Planilla 9/Reporte_CRA_15min*.csv`, `buscar_reporte_cra`; con traspaso,
  `calculo_sscc_maestro` y `reporte_cra`). `ejecutar_actualizacion(...,
  hacer_sc_co=["SC", "CO"])` corre el motor de `Actualiza_SC_CO.py` después
  de FD y prorrata, sobre el mismo libro, y guarda una sola vez.
- **Detalles que importan:** la prorrata se transforma de tabla larga a matriz
  pivote y se pega desde `B8`; suministradores ordenados alfabéticamente,
  faltantes en 0. **Modo 15 minutos:** la columna A del origen es el bloque de
  15 minutos (hasta 2976 filas por mes, `MAX_FILAS_PRORRATA` = 2980 con el
  cambio de hora). Antes de pegar se borra `B8:EC` hasta la fila más larga entre
  la matriz vieja (columna B corrida) y la nueva — ya no hasta la 756 fija. Las
  columnas siguen siendo `B:EC` porque a la derecha las planillas 5 y 6 tienen
  bloques de montos. Si el origen trae 745 filas o menos avisa que parece
  horario; si trae más de 2980, corta. La hoja de origen es
  **`Prorrata 15min`** (antes `PRORRATA_HORARIA_TABULAR`), con las mismas
  columnas A/B/C desde la fila 2; se acepta también con otras mayúsculas,
  espacios o guiones bajos. V9/V10/V11 del Revisor 15 min leen la misma hoja. Los `.xlsm` destino tienen que estar cerrados.
  El motor `aplicar_mapeo` todavía acepta `filtro_col`, `detectar_fin_primera_vacia`
  y `ajustar_formulas`, que hoy no usa ningún mapeo (los usaba el del tabulado).

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

### `__comun__/salidas.py` — **hecho**

- **Qué hace:** centraliza las raíces de `00_Salidas` y `__config__`, y la
  estructura `AAAA/MM Mes`, compartida por el Revisor y los comparadores.
- **Expone:** `raiz_salidas(dir_script)`, `partir_aamm(aamm)`,
  `nombre_carpeta_mes(aamm)`, `carpeta_anio(dir_salidas, aamm)`,
  `carpeta_mes(dir_salidas, aamm, crear=False)`,
  `carpeta_comparador(dir_salidas, anio, nombre)` y
  `carpetas_legado(dir_salidas)`.
- **Reglas:** nunca lee ni escribe en carpetas planas AAMM antiguas; las detecta
  solo para avisar. Tolera variantes existentes como `7 Julio` o `07 julio`,
  y, también con `crear=True`, reutiliza esa variante para no partir el mes en dos carpetas.
- **Pruebas:** `__comun__/test_salidas.py`, solo stdlib.

### `__comun__/config.py` — **hecho**

- **Qué hace:** lee y escribe el `config.json`, indexado por `<equipo>_<usuario>`.
- **Expone:** `clave_equipo()`, `leer(ruta)`, `leer_todo(ruta)`,
  `guardar(ruta, data)`, `modificar(ruta, mutador)`, `escribir_json(ruta, data)`.
- **Reglas que no cambian:** solo se agregan o actualizan claves, nunca se borra
  nada ajeno; si el archivo existe pero no se puede interpretar **no se escribe**;
  la escritura es atómica (`.tmp` + `os.replace`).
- **Pruebas:** `__comun__/test_config.py`, 14 casos, solo stdlib.
- **Dos bugs que se arreglaron al juntar las copias:**
  1. `ActualizaRemplazos.py` escribía el archivo entero **sin `.tmp`**, y si el
     config existía pero estaba roto lo pisaba con `{}`. O sea: un `config.json`
     ilegible le borraba los ajustes a los otros nueve scripts.
  2. `get_usuario` tenía versiones con y sin `try/except`. Quedó la defensiva.
- **Migración:** los doce programas que manejan configuración ya delegan acá.
  Conservan sus nombres históricos como aliases o envoltorios breves, así que
  ningún punto de llamada cambió. `ActualizaRemplazos.py` sigue usando su JSON
  propio; compartir el motor no mezcla los datos.

| Script | Estado |
|---|---|
| Revisor, 9 actualizadores y 2 comparadores | migrados |

### `__comun__/traspaso.py` — **hecho**

- **Qué hace:** centraliza el sobre del JSON opcional que el Revisor pasa por
  `argv[1]`: origen, versión máxima y normalización de `rutas`.
- **Expone:** `ORIGEN`, `VERSION_ACTUAL`, `validar(data, version_max=...)`,
  `leer(ruta, version_max=...)` y `leer_argumento(argv, version_max=...)`.
- **Regla que no cambia:** cualquier argumento ausente, archivo inválido, JSON
  roto, origen ajeno o versión futura devuelve `None`; el actualizador continúa
  funcionando solo y busca sus archivos manualmente.
- **Pruebas:** `__comun__/test_traspaso.py`, 8 casos, solo stdlib.

### `__comun__/excel_xml.py` — **hecho**

- **Qué hace:** lee columnas calculadas de `.xlsx`/`.xlsm` escaneando el ZIP y
  XML, sin abrir Excel; ante un fallo devuelve `None` para conservar el fallback.
- **Lo usan:** Revisor y `Actualiza_Data_Access.py`, mediante aliases históricos.
- **Pruebas:** `__comun__/test_excel_xml.py`, con un OOXML sintético.

### `__comun__/csv_columnas.py` — **hecho** (2026-10-23)

- **Qué hace:** lee un `.csv` como `{COL: {fila: valor}}`, igual que
  `excel_xml.leer_columnas_rapido`, con la fila contada como en Excel.
  Separador `;`, `,` o tabulador; coma decimal con `;`; UTF-8 o cp1252.
- **Lo usan:** el `Actualiza_SC_CO.py` y el Revisor de la versión 15 minutos
  (CO desde `Reporte_CRA_15min`).
- **Pruebas:** `__comun__/test_csv_columnas.py`, con las filas reales del
  Reporte_CRA de diciembre 2025.

### `__comun__/texto.py` y `__comun__/archivos.py` — **hecho**

- **Qué hacen:** separan normalización suave de claves estrictas y centralizan
  temporales/copias de Windows. Conservan explícitamente las variantes antiguas
  para `None` y `0`, en vez de cambiar casos límite al migrar.
- **Pruebas:** `__comun__/test_texto.py` y `__comun__/test_archivos.py`.

### `__comun__/comparadores.py` — **hecho**

- **Qué hace:** infraestructura compartida por composición: caché `scandir`,
  estado/inclusión, respaldo y hojas ajenas, helpers
  de rutas/Access y puente `ColaTk` para que workers no toquen tkinter.
- **Regla:** no contiene SQL, columnas, vistas ni lectura de dominio; Access y
  Tabulado siguen siendo motores independientes.
- **Pruebas:** `__comun__/test_comparadores.py`, incluidas cola con hilo real,
  caché, copias, estado y respaldos.

### `_morir()` se queda duplicado a propósito

Cada ejecutable define su propio `_morir()` (ventana + `SystemExit(1)`) y hace
`from __comun__ import ...` **adentro de un `try/except ImportError`**. No se
puede centralizar: es justamente lo que avisa cuando lo que falta es
`__comun__/`. Lanzados por el Revisor con `pythonw` no hay consola, así que un
`ImportError` suelto mata el script en silencio y solo se ve que la ventana
nunca apareció.

### `Revisor_Relq/revisor/` — Fase 4, **implementada**

Paquete interno del Revisor. `Revisor_Reliquidacion.py` sigue siendo el único
punto de entrada y el archivo literal que localizan los hermanos; reexporta
todos los nombres históricos, así que nada de afuera cambió.

- **`revisor/archivos.py`:** búsqueda de carpetas/archivos y el caché de
  directorios (`cache_directorios`, `leer_dir`, `buscar_archivo`,
  `listar_diarios`, `mtime`, `tamano`, `fmt_fecha`, `fmt_monto`,
  `iguales_mtime`). **`TOL_MTIME` vive acá y en ningún otro lado.**
- **`revisor/estado.py`:** `Estado` y `CacheValores` con sus dependencias de
  disco inyectadas, más `leer_estado_mes`. El punto de entrada las subclasea
  para conservar los constructores sin argumentos de siempre.
- **`revisor/lectores.py`:** adaptadores de lectura Excel y MDB:
  columnas con fallback openpyxl/xlwings, escaneos y diagnósticos OOXML,
  resolución de hojas, armado de tablas, inspección Access y resolución cacheada
  de los valores configurados.
- **`revisor/verificaciones.py`:** motor de los tipos de comprobación usados por
  V4…V17. No importa la ventana; recibe sus dependencias desde el punto de entrada.
- **`revisor/lanzamiento.py`:** bloqueo de destinos Excel, armado del JSON de
  traspaso y lanzamiento de actualizadores, sin importar el punto de entrada.
- **Pruebas:** `revisor/test_archivos.py`, `revisor/test_estado.py` y
  `revisor/test_lectores.py` y `revisor/test_modulos_fase4.py`, que se
  corren igual que las de `__comun__`: `python Revisor_Relq/revisor/test_*.py`.

La extracción de la Fase 4 está completa. Falta probar el Revisor de punta a
punta en Windows con un mes real, Excel y Access. Ver
`docs/PLAN_MODULARIZACION_TOKENS.md`.

### Lo que sigue

El generador detecta **39 nombres de constante repetidos** (la tabla completa
está al final de `INTERFACES.md`). Es una detección sintáctica, y tres de esos
nombres **tienen que** seguir repetidos: `CONFIG_PATH` y `_RAIZ_COMUN` se
calculan desde la ubicación de cada ejecutable, y los dos `TRASPASO_*` ya son
aliases del contrato único, no valores duplicados.
Entre las duplicaciones todavía reales, las que más rinden son:

| Constante | En cuántos archivos | Qué es |
|---|---|---|
| `DIR_SCRIPT` | 7 | la carpeta del script |
| `CENTRALES_EMBALSE` | 2 | el caso documentado, el único que ya obligó a defenderse con una verificación |
| `SERVER`, `DRIVER`, `CHUNK`, `LARGO_TEXTO` | 2 | la conexión a SQL Server |

**El orden de migración no es `CENTRALES_EMBALSE`**, aunque sea el caso más famoso:
es el que menos riesgo tiene de quedar mal, porque V10 lo caza en los dos sentidos.
El lector de Excel por ZIP/XML y la infraestructura duplicada de comparadores ya
se migraron. La Fase 4 dividió internamente el Revisor y queda pendiente su
validación integral en Windows, la frontera de riesgo alto que el contenedor no
puede reproducir.

El orden y los límites de las fases siguientes quedaron en
`docs/PLAN_MODULARIZACION_TOKENS.md`. Aunque una pieza tenga varios consumidores,
se migra y verifica uno antes de avanzar al próximo.
