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

- Del encabezado de cada archivo se muestran las primeras líneas; el resto está
  arriba de todo en el `.py`.
- Las funciones que empiezan con `_` son internas del archivo.
- Cuando una función no tiene docstring se muestra el comentario que tenga encima.
- Los `— TÍTULO —` son los banners de sección del propio archivo.
- Los valores de las constantes largas salen resumidos; el valor exacto está en el
  archivo.


## Índice

- [`Revisor Reliquidación/comun/config.py`](#revisor-reliquidacióncomunconfigpy) — 117 líneas — Lectura y escritura del config.json, indexado por <equipo>_<usuario>.
- [`Revisor Reliquidación/Reemplazos REUC/ActualizaRemplazos.py`](#revisor-reliquidaciónreemplazos-reucactualizaremplazospy) — 1860 líneas — ActualizaRemplazos.py
- [`Revisor Reliquidación/Revisor_Reliquidacion.py`](#revisor-reliquidaciónrevisor_reliquidacionpy) — 6840 líneas — Revisor de entregables - CASO RELIQUIDACION
- [`Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`](#revisor-reliquidaciónactualizadoresactualiza_access_p9py) — 1135 líneas — Actualiza el Access de la planilla 9
- [`Revisor Reliquidación/actualizadores/Actualiza_Cuadro0.py`](#revisor-reliquidaciónactualizadoresactualiza_cuadro0py) — 1030 líneas — Actualiza Cuadro 0 (0_CUADROS_RELIQUIDACION SSCC)
- [`Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py`](#revisor-reliquidaciónactualizadoresactualiza_data_accesspy) — 1598 líneas — Actualiza la tabla [Sobrecostos] de un Access .mdb consolidando la informacion
- [`Revisor Reliquidación/actualizadores/Actualiza_Energia.py`](#revisor-reliquidaciónactualizadoresactualiza_energiapy) — 809 líneas — Actualizar Energia
- [`Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py`](#revisor-reliquidaciónactualizadoresactualiza_sc_copy) — 906 líneas — Actualiza la hoja "SC y CO" de la planilla 5_
- [`Revisor Reliquidación/actualizadores/Actualiza_datos.py`](#revisor-reliquidaciónactualizadoresactualiza_datospy) — 1343 líneas
- [`Revisor Reliquidación/actualizadores/Carga_Retiros.py`](#revisor-reliquidaciónactualizadorescarga_retirospy) — 890 líneas — Carga Retiros_h.parquet a SQL Server
- [`Revisor Reliquidación/actualizadores/Prorratear.py`](#revisor-reliquidaciónactualizadoresprorratearpy) — 920 líneas — Prorratear: del Access a SQL Server


---

## `Revisor Reliquidación/comun/config.py`

> Lectura y escritura del config.json, indexado por <equipo>_<usuario>.
>
> Esto estaba copiado en los 10 scripts, con cuatro variantes de
> `_modificar_config` y cuatro de `get_usuario` que se fueron separando entre si.
> Las diferencias eran casi todas cosmeticas (docstrings, type hints, nombres de
> variable) salvo dos, que aca quedan resueltas para todos:
>
>   * `ActualizaRemplazos.py` escribia el archivo entero SIN pasar por un .tmp, y
>     si el config existia pero no se podia interpretar lo pisaba con {}. O sea:
>     un config.json roto le borraba los ajustes a todos los demas scripts. Ahora
>     usa la misma regla que el resto, que es no escribir.
>   * `get_usuario` tenia versiones con y sin try/except. Queda la defensiva.
>
> REGLAS QUE NO CAMBIAN, porque el archivo es compartido:
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `__future__`, `json`, `os`, `pathlib`, `socket`

### Funciones

#### `def clave_equipo() -> str`

Devuelve `<equipo>_<usuario>`, que es como se indexa el config.json.

No lanza: si el nombre del equipo no se puede averiguar devuelve
"desconocido", porque quedarse sin ajustes es mejor que no arrancar.

#### `def escribir_json(ruta, data) -> None`

Escritura atomica: primero un .tmp y despues os.replace.

Evita dejar el archivo truncado si algo falla a medio camino. Se usa
tambien para el JSON de traspaso del revisor, no solo para el config.

#### `def leer_todo(ruta) -> dict`

El config.json completo, con todos los equipos. {} si no se puede leer.

#### `def leer(ruta) -> dict`

El bloque del equipo actual. {} si no hay archivo o esta roto.

#### `def modificar(ruta, mutador) -> bool`

Lee el config entero, lo modifica con `mutador(todo)` y lo reescribe.

`mutador` recibe el dict completo (todos los equipos), no solo el bloque
propio, porque algunos scripts anidan claves por mes.

Devuelve True si se escribio. Devuelve False sin tocar el archivo si existe
pero no se puede interpretar o no es un dict: es un archivo compartido y
pisarlo le borraria los ajustes a los demas scripts.

#### `def guardar(ruta, data: dict) -> bool`

Agrega o actualiza claves en el bloque del equipo actual.


---

## `Revisor Reliquidación/Reemplazos REUC/ActualizaRemplazos.py`

> ActualizaRemplazos.py
>
> Entradas (por ahora, en la carpeta desde donde se ejecuta el script):
> "Reemplazos forzados.xlsx"                                      Copiar del ultimo mes reliquidado y confirmar
> "datos_reuc_reemplazos_*.xlsx"                                  descargar de la pag REUC
> "datos_reuc_*.xlsx"                                             descargar de la pag REUC
> "Cuadros de Pago_Balances_SEN_*_Simplificado_def.xlsb"          de PLABACOM, el más actualizado
> "1_CUADROS_PAGO_SSCC_*.xlsm"                                    del disco T, el más actualizado
>
> Archivo destino (se selecciona en la ventana):
> "0_CUADROS_RELIQUIDACIÓN SSCC_*.xlsm"  -> el del mes que estamos trabajando.
>
> El programa:
>   1. Genera el archivo de salida "Reemplazos_AAAAMMDD_SSCC.xlsx" (igual que antes).
>   2. Escribe directamente en la hoja EMPRESAS del archivo destino:
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `datetime`, `json`, `os`, `pandas`, `pathlib`, `queue`, `re`, `shutil`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`, `xlwings`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `CARPETA_AUXILIARES` | `Path(__file__).parent / 'Auxiliares'` | CONFIG POR PC/USUARIO La carpeta Auxiliares vive AL LADO del .py y es compartida por todos los usuarios.  |
| `CONFIG_PATH` | `CARPETA_AUXILIARES / 'config.json'` |  |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` | UTILIDADES TRASPASO DESDE EL REVISOR El Revisor escribe un JSON en Salidas/AAMM/ y pasa su ruta como argv[1].  |
| `TRASPASO_VERSION_MAX` | `1` |  |
| **— BÚSQUEDA EN DISCO COMPARTIDO (PLABACOM) —** | | |
| `RAIZ_PLABACOM_DEFAULT` | `'T:\\Facturacion\\Plabacom'` |  |
| `PATRON_BALANCES` | `'balances_sen.*simplificado.*\\.xlsb$'` | Se compara contra el nombre normalizado (minúsculas, sin tildes). |
| **— BÚSQUEDA CONJUNTA: BALANCES SEN + 1_CUADROS_PAGO_SSCC —** | | |
| `RAIZ_FACTURACION_DEFAULT` | `'T:\\Facturacion'` |  |
| `NARANJO` | `'#d97706'` |  |
| `PATRON_SSCC` | `'1_cuadros_pago_sscc.*\\.xlsm$'` |  |
| `MESES_ES` | `lista de 12 elementos: 'enero', 'febrero', 'marzo', …` |  |
| **— Verificación de coherencia nombre de archivo vs carpeta —** | | |
| `MESES_ABR` | `dict de 13 claves: 'ene', 'feb', 'mar', …` |  |
| **— DESCARGA AUTOMÁTICA DESDE REUC (playwright) —** | | |
| `URL_REUC_EMPRESAS` | `'https://reuc.coordinador.cl/maestro_usuarios/empresas/exportar_reuc?&text_search='` |  |
| `URL_REUC_REEMPLAZADAS` | `'https://reuc.coordinador.cl/maestro_usuarios/empresas/export_reemplazadas_data?&text_sea…` |  |

### Funciones

#### `def get_usuario()`

#### `def leer_config()`

#### `def guardar_config(data)`

#### `def leer_traspaso(argv)`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def normalizar(texto)`

#### `def buscar_archivo(ruta, patron)`

Busca por glob y devuelve el más reciente (ignora temporales ~$).

#### `def extraer_fecha(texto)`

Busca una fecha tipo AAAAMMDD dentro de un string.

#### `def extraer_fecha_aamm(texto)`

Extrae fecha AAMM.

#### `def fecha_aaaammdd_a_ddmmaaaa(aaaammdd)`

'20260722' -> '22-07-2026'. Si no calza el formato, devuelve tal cual.

#### `def col_letra_a_num(letra)`

#### `def subcarpeta_que_contiene(carpeta, *fragmentos)`

Devuelve la primera subcarpeta cuyo nombre normalizado contiene TODOS
los fragmentos dados. Tolera tildes, mayúsculas y espacios de más.

#### `def meses_hacia_atras(desde=None, cantidad=36)`

Genera tuplas (año, mes) desde el mes actual hacia atrás.

#### `def archivos_que_matchean(carpeta, patron)`

#### `def archivo_en_carpeta_resultados(carpeta_version, aamm, patron_regex, log=None)`

Dentro de "<version>/Publicar/01 Resultados_AAMM_.../" busca el archivo.
Si no lo encuentra por esa ruta, hace una búsqueda recursiva de respaldo
dentro de la carpeta de la versión. Devuelve la ruta o None.

#### `def buscar_plabacom(raiz=RAIZ_PLABACOM_DEFAULT, patron_regex=PATRON_BALANCES, meses=36, log=None)`

Recorre T:\Facturacion\Plabacom\<AAAA>\<AAMM>\<02 Definitivo | 01 Preliminar>
         \Publicar\01 Resultados_AAMM_...\<archivo>

Va del mes más reciente hacia atrás. En cada mes prueba primero
"02 Definitivo" y luego "01 Preliminar". Devuelve (ruta, descripcion) o
(None, mensaje de error).

log: callback opcional para ver el detalle del recorrido.

#### `def carpeta_mes_facturacion(raiz_facturacion, y, m)`

T:\Facturacion\<AAAA>\<MM Mes>   (ej: 2026\06 Junio)
Acepta que la carpeta se llame por número, por nombre, o ambos.

#### `def buscar_sscc_en_version(carpeta_version, patron_regex=PATRON_SSCC, log=None)`

Dentro de "<version>/SSCC/" busca 1_CUADROS_PAGO_SSCC_*.xlsm.

#### `def aamm_desde_nombre(nombre)`

Intenta sacar el AAMM del nombre del archivo. Reconoce dos formatos:
  - "..._2606_..."          -> 2606
  - "..._jun26_..."         -> 2606
Devuelve None si no logra determinarlo.

#### `def sufijo_desde_nombre(nombre)`

Devuelve 'def', 'pre' o None según el sufijo del archivo.

#### `def verificar_coherencia(ruta, aamm_carpeta, version_carpeta)`

Compara el nombre del archivo contra la carpeta donde está.
Devuelve una lista de avisos (vacía si todo calza).

#### `def buscar_par_mensual(raiz_facturacion=RAIZ_FACTURACION_DEFAULT, meses=36, log=None)`

Busca los DOS archivos exigiendo que vengan del MISMO mes y del MISMO
proceso (Definitivo o Preliminar):

    Balances SEN : <raiz>\Plabacom\<AAAA>\<AAMM>\<version>\Publicar\01 Resultados_...
    SSCC         : <raiz>\<AAAA>\<MM Mes>\<version>\SSCC

Recorre del mes más reciente hacia atrás. En cada mes prueba Definitivo y
después Preliminar; solo acepta la combinación si ENCUENTRA LOS DOS.

Devuelve (dict, info) o (None, mensaje).

#### `def carpeta_auxiliares()`

Carpeta Auxiliares (al lado del .py). La crea si no existe.

#### `def buscar_archivo_con_respaldo(carpeta_preferida, patron, log=print)`

Busca primero en carpeta_preferida (la que eligio el usuario o quedo
guardada en config). Si ahi no esta, cae de respaldo a Auxiliares
-- que es donde SIEMPRE deberia estar-- y avisa por log cual de las
dos uso, para que quede claro y no parezca que "no encuentra nada".

#### `def descargar_reuc(carpeta_destino=None, log=print, timeout_login_seg=600, espera_descarga_seg=180)`

Abre un navegador para que el usuario inicie sesion en REUC y, apenas
la sesion queda confirmada, descarga los dos exports:
  - datos_reuc_*.xlsx               (exportar_reuc)
  - datos_reuc_reemplazos_*.xlsx    (export_reemplazadas_data)

Los guarda en la carpeta Auxiliares (por defecto, la que esta al lado
del .py y es compartida por todos los usuarios).

Flujo:
  1. Espera a que el usuario termine el login (el acceso unificado
     rebota entre dominios; solo se acepta la URL de REUC misma).
  2. Confirma la sesion con una peticion liviana.
  3. Descarga cada export UNA sola vez, esperando con paciencia.

Seguridad: no se guarda nada de la sesion. Cada ejecucion abre un
navegador nuevo y sin memoria (sin perfil, sin cookies persistidas).
La clave nunca pasa por este script: se escribe directamente en la
pagina real de REUC, dentro del navegador.

Devuelve un dict {"reuc": Path, "reemplazos": Path}.

#### `def ultima_fila(sheet, col_letra, desde)`

Última fila con contenido en una columna (considera fórmulas).

#### `def procesar_datos(carpeta_datos, archivo_cuadros, archivo_xlsb=None, archivo_xlsm=None, log=print)`

Devuelve un dict con los DataFrames que se usan tanto para el archivo
de salida como para escribir en el archivo destino.

archivo_xlsb / archivo_xlsm: rutas explícitas (disco compartido). Si vienen
en None, se buscan en la carpeta de datos como respaldo.

#### `def carpeta_reemplazos_reuc(ruta_destino)`

Dos niveles por encima de la carpeta del "0_CUADROS_RELIQUIDACION":

    ...\SSCC\02 CASO RELIQUIDACION\00 Entregables\0_CUADROS_...xlsm
         ^-- aca se crea "Reemplazos REUC"

O sea: <carpeta del archivo>.parents[2] / "Reemplazos REUC".

#### `def respaldar_fuentes(datos, carpeta, log=print)`

Copia a "Reemplazos REUC" los archivos que se usaron para armar el
resultado, para que quede todo junto y trazable. Si ya existen, se
reemplazan.

#### `def generar_archivo_salida(datos, carpeta_salida, log=print)`

#### `def escribir_en_destino(ruta_destino, datos, log=print, dejar_abierto=True)`

Hoja EMPRESAS del archivo destino:
  B:C -> EMPRESA / RUT           (desde df_salida)
  H:I -> Reemplazada / Reemplazante  (reemplazos válidos + forzados)
Fila 1 = encabezado en ambos bloques. Se limpia el contenido antes de pegar,
manteniendo el formato de las celdas.

**— VENTANA —**

#### `def main()`


---

## `Revisor Reliquidación/Revisor_Reliquidacion.py`

> Revisor de entregables - CASO RELIQUIDACION
> Ventana que replica la estructura de carpetas de "02 CASO RELIQUIDACION" y permite:
>
>   * ACTUALIZAR  -> ubica todos los archivos, compara copias contra su maestro por
>                    fecha y hora de modificacion (pinta AMARILLO las copias
>                    desactualizadas, ROJO lo que falta) y revisa la vigencia de
>                    las verificaciones de valores.
>   * VERIFICAR   -> boton por archivo que compara sumas de sobrecostos entre
>                    archivos (.xlsm / .mdb) y deja registrada la verificacion
>                    con su fecha y las fechas de modificacion de sus dependencias.
>
> Requisitos:  pip install xlwings openpyxl pyodbc
> Para .mdb se necesita el "Microsoft Access Driver (*.mdb, *.accdb)" con la misma
> arquitectura (32/64 bits) que el Python que ejecuta el script.

**Importa:** `csv`, `datetime`, `json`, `os`, `pathlib`, `queue`, `re`, `shutil`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `TOLERANCIA` | `1.0` |  |
| `UMBRAL_DESCUADRE_CPRT` | `1000.0` | Residuo maximo aceptado en el descuadre del cuadro de pago (CPRT!I3).  |
| `TOL_SOBRECOSTO_FILA` | `1.0` | Diferencia maxima aceptada en UNA fila al recalcular el sobrecosto desde sus componentes.  |
| `TOL_PAGO_EMPRESA` | `150.0` | Diferencia maxima aceptada al comparar el pago por empresa y concepto entre la planilla 1 y la 9 (o la 4).  |
| `TOTALES_PRORRATA` | `(1.0, 100.0)` | Las filas de prorrata suman el 100%, escrito como 1 o como 100 segun la planilla.  |
| `TOL_PRORRATA` | `0.0001` |  |
| `TOL_PRORRATA_SUMA` | `0.0001` | Diferencia maxima al comparar la suma por suministrador de la prorrata de una planilla contra el Prorrata_Retiros.  |
| `RE_UNIDAD_CENTRAL` | `re.compile('-\\d+\\s*$')` | Una central que termina en "-numero" es una unidad, y las unidades son lo que tienen los embalses.  |
| `CENTRALES_EMBALSE` | `lista de 27 elementos: 'CANUTILLAR-1', 'CANUTILLAR-2', 'ELTORO-1', …` | Centrales de embalse OJO: esta lista esta TAMBIEN en Actualiza_SC_CO.py.  |
| `TOL_MTIME` | `2` |  |
| `VALORES` | `dict de 56 claves: 'TOTAL_SSCC', 'TOTAL_CO', 'TOTAL_CCA', …` |  |
| `VERIFICADORES` | `dict de 14 claves: 'V8', 'V9', 'V10', …` |  |
| `XL` | `('.xlsm', '.xlsx', '.xlsb')` |  |
| `DB` | `('.mdb', '.accdb')` |  |
| `DD` | `('Detalles diarios', 'Detalle diario', 'Detalle diarios', 'Detalles diario')` | La carpeta de detalles se llama distinto segun el modulo ("Detalles diarios" en Sobrecostos, "Detalle diario" en CO y CCA).  |
| `NODOS` | `lista de 36 elementos: dict(id='c_fd', tipo='carpeta', pref='', texto='../FD/          « ORIGEN — fuera de 02 CASO RELIQUIDACION »'), dict(id='a_sscc_desempeno', tipo='archivo', pref='    └── ', texto='SSCC_Desempeno*.xlsx\|xlsm   (origen de las hojas FD)', carpeta=['FD'], sube=True, solo_info=True, patron='^sscc_desempeno', ext=XL, espejo=None), dict(id='c_ent', tipo='carpeta', pref='├── ', texto='00 Entregables/'), …` |  |
| `NODO_POR_ID` | `{n['id']: n for n in NODOS}` |  |
| `ACTUALIZADORES` | `dict de 10 claves: 'a_calc_sscc_01', 'a_3_p9', 'a_retiros_parq', …` | Botón "Actualizar data": qué script lanza cada archivo maestro Solo va en los MAESTROS, que son los que el usuario edita.  |
| `CLAVES_TRASPASO` | `dict de 16 claves: 'a_sscc_desempeno', 'a_cons_tab', 'a_prorrata', …` | Traduccion de los id del arbol a las claves del JSON de traspaso, que son las que esperan los actualizadores. |
| `ARCHIVO_TRASPASO` | `'_traspaso_actualizador.json'` |  |
| `TRASPASO_VERSION` | `1` |  |
| `ACCIONES_INTERNAS` | `{'a_0_cuadros': [('Exportar CPRT', '_exportar_cprt')]}` | Acciones que corren DENTRO del revisor, sin lanzar otro proceso.  |
| `C_OK` | `'#1a7f1a'` | colores |
| `C_FALTA` | `'#c00000'` |  |
| `C_AMARILLO` | `'#ffe600'` |  |
| `C_VENCIDA` | `'#ff9900'` |  |
| `C_GRIS` | `'#777777'` |  |
| `C_NEUTRO` | `'SystemButtonFace'` |  |
| `DIR_SCRIPT` | `Path(__file__).resolve().parent` |  |
| `CONFIG_PATH` | `DIR_SCRIPT / 'config.json'` |  |
| `DIR_SALIDAS` | `DIR_SCRIPT / 'Salidas'` |  |
| `ARCHIVO_ESTADO` | `'_revisor_verificaciones.json'` |  |
| `_DIR_CACHE` | `{'on': False, 'datos': {}, 'hits': 0, 'scans': 0}` | Cache de directorios Una relectura completa hacia 68 recorridos de carpeta para 13 carpetas distintas: cada nodo del arbol recorria la carpeta entera de nuevo, y encima resolver_carpeta recorria la r… |
| `RE_COPIA` | `re.compile('(-\\s*cop(?:ia\|y)(?:\\s*\\(\\d+\\))?\|\\(\\d+\\))\\s*$')` | Sufijos que deja Windows al copiar: "archivo - copia.mdb", "archivo - copia (2).mdb", "archivo - Copy.xlsm". |
| `ARCHIVO_CACHE` | `'_revisor_cache_valores.json'` |  |
| `CACHE` | `CacheValores()` |  |
| `ESTADO` | `Estado()` |  |
| `CACHE_COLUMNAS` | `{}` |  |
| `NS_XL` | `'{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'` |  |
| `NS_REL` | `'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'` |  |
| `_ENT_XML` | `{'lt': '<', 'gt': '>', 'quot': '"', 'apos': "'", 'amp': '&'}` |  |
| `_RE_ENT` | `re.compile('&(?:#(\\d+)\|#[xX]([0-9a-fA-F]+)\|(lt\|gt\|quot\|apos\|amp));')` |  |
| `C_LOG_MALO` | `'#c00000'` | Coloreado del log y del detalle Los mensajes ya vienen rotulados: ">>" es fallo, "OK" es bien, "?" es sin datos y ".." es "trabajando".  |
| `C_LOG_DUDA` | `'#b45309'` |  |
| `C_LOG_BIEN` | `'#1d6b1d'` |  |
| `_PALABRAS_MALO` | `tupla de 13 elementos: 'FALTA ', 'FALTA EL ARCHIVO', 'NO CUADRA', …` | Palabras que pintan la linea sin importar el bloque en que caiga. |
| `_PALABRAS_DUDA` | `tupla de 10 elementos: 'OJO', 'AMARILLO', 'ADVERTENCIA', …` | Ojo: aca NO van "omitidas" ni "descartadas".  |
| `HOJA_CPRT` | `'CPRT'` | Exportar el CPRT a csv El cuadro cero trae en su hoja "CPRT" el cuadro de pago ya armado.  |
| `CPRT_FILAS_META` | `5` |  |
| `CPRT_FILA_ENCABEZADO` | `6` |  |
| `CPRT_N_CAMPOS` | `7` | El csv lleva 7 campos: A:F mas UNA columna de monto.  |
| `CPRT_COL_MONTO` | `8` |  |
| `CPRT_ROTULO_DESDE_G` | `True` | El ROTULO de esa columna se sigue tomando de la G, o sea que el encabezado dice "Monto" y no "Monto retenido".  |
| `CPRT_CODIF` | `'cp1252'` |  |
| `UMBRAL_PAR_SEGURO` | `100.0` | El cuadro de pagos: la matriz y la tabla dinamica CuadroPago arma en 01.SSCC_Recurso_Tecnico una matriz cruzada: N8 = "Pagan" O8:..8 = las empresas que RECIBEN N9:N..  |

### Clases

#### `class cache_directorios`

Enciende el cache mientras dura el bloque. Devuelve (scans, hits) al salir
en self.stats, para poder decir en la bitacora cuanto se ahorro.


#### `class Estado`

Verificaciones de un mes. Se guardan en Salidas/AAMM junto al .py.

- `def __init__(self)`
- `def cargar(self, aamm)`
- `def existe(self)`
- `def guardar(self)`
- `def get(self, vid)`
- `def vigente(self, vid)` — El registro guardado, si hay alguno. 
- `def firma_guardada_distinta(self, vid)` — True si hay un registro pero es de una definicion anterior. 
- `def set(self, vid, registro)`

#### `class CacheValores`

Guarda el valor ya leido de cada origen junto con la ruta y la fecha de
modificacion del archivo. Si el archivo no cambio y se pide lo mismo, no se
vuelve a abrir. Se guarda en Salidas/AAMM para que sirva entre ejecuciones.

- `def __init__(self)`
- `def cargar(self, aamm)`
- `def guardar(self)`
- `def obtener(self, clave, ruta, huella)`
- `def poner(self, clave, ruta, huella, valor, filas=None)`
- `def descartar(self, claves=None)`

#### `class Revisor`

- `def __init__(self, root)`
- `def log(self, msg='')`
- `def examinar(self)`
- `def actualizar(self, motivo=None, solo_ids=None)` — Relee la carpeta: reubica los archivos, repinta fechas y recalcula la vigencia de las verificaciones. 
- `def estado_verificador(self, vid, _visitados=None)` — OK | VENCIDA | NO CUADRA | NO SE PUDO | SIN VERIFICAR. 
- `def configurar_valores(self)` — Ventana para indicar donde esta cada valor, sin editar el codigo. 
- `def ir_a_mes(self)` — Lleva TODA la ventana al mes escrito: cambia la carpeta del caso a la que se usó ese mes, recarga el árbol y su estado…
- `def reiniciar_mes(self)` — Borra el estado y los valores guardados de un mes, para partir limpio.
- `def ver_detalle_verificacion(self, vid)` — Totales, comprobaciones y bitácora de la última corrida.
- `def ver_estado_mes(self)` — Abre el estado guardado del AAMM escrito en el cuadro.
- `def verificar(self, vid)`
- `def verificar_todo(self)`

### Funciones

#### `def clave_concepto(t)`

Normaliza un concepto o una empresa para comparar entre planillas.

Saca tildes, espacios y guiones BAJOS, y pasa a mayusculas. Hace falta porque
las dos planillas escriben lo mismo distinto: la 1 dice "CO ERNC" con espacio
y la 9 dice "CO_ERNC" con guion bajo. Sin esto no se cruza ni un concepto.
Se conservan los parentesis y el signo, que SI distinguen: CSF(+) y CSF(-) son
conceptos diferentes.

#### `def clave_central(t)`

Normaliza el nombre de una central para comparar: sin tildes, sin espacios
ni guiones bajos, en mayusculas. Asi 'El Toro-1' y 'ELTORO-1' son la misma.

#### `def dir_mes(aamm, crear=False)`

Salidas/AAMM junto al .py. Solo la crea si crear=True.

#### `def escribir_json(ruta, data)`

Escritura atomica: primero un .tmp y despues os.replace.
Evita dejar el archivo truncado si algo falla a medio camino.

#### `def get_usuario()`

#### `def leer_config()`

#### `def guardar_config(data)`

#### `def leer_valores_cfg()`

La ubicacion de los valores es estructural, no por usuario: va en la
clave compartida '_valores' de config.json.

#### `def guardar_valores_cfg(clave, campos)`

#### `def aplicar_valores_cfg()`

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def normalizar(texto)`

#### `def leer_dir(carpeta)`

El listado de una carpeta, del cache si esta encendido.

#### `def buscar_carpeta(base, nombre)`

Busca subcarpeta tolerando tildes, mayusculas y espacios extra.

#### `def resolver_carpeta(base, partes)`

Cada parte puede ser un nombre o una tupla de nombres alternativos.

#### `def es_temporal(nombre)`

#### `def es_copia(nombre_sin_extension)`

#### `def buscar_archivo(carpeta, patron_regex, extensiones)`

Devuelve el archivo mas reciente que calza el patron (sobre nombre normalizado).

#### `def listar_diarios(carpeta, patron_regex, extensiones)`

{fecha_AAAAMMDD: Path} para las planillas diarias de una carpeta.

#### `def mtime(p)`

Fecha de modificacion. Con el cache encendido sale del recorrido de la
carpeta, sin un stat por archivo.

#### `def tamano(p)`

#### `def fmt_fecha(ts)`

#### `def iguales_mtime(a, b)`

#### `def fmt_monto(v)`

#### `def detectar_aamm(rutas, diarios)`

Deduce el AAMM reliquidado. Primero del sufijo _AAMM_ de los nombres de
archivo (se queda con el mas repetido); si no, de las fechas AAAAMMDD de
los detalles diarios. Devuelve (aamm, motivo) o (None, motivo).

#### `def firma_verificador(vid)`

Huella corta y estable de COMO esta definida una verificacion.

Existe para que un resultado guardado con una definicion vieja no se muestre
nunca como si fuera del chequeo actual. Si se cambian las hojas, columnas,
filas o comprobaciones de un verificador, la huella cambia y el registro
anterior se descarta en vez de quedar mostrando datos de rangos que ya no se
leen (por ejemplo columnas que ni existen en la definicion nueva).

#### `def partir_signo(clave)`

En las listas de comprobaciones una clave puede llevar '-' delante para
entrar con signo cambiado.  '-P4_L6' -> ('P4_L6', -1.0)

#### `def huella_spec(spec)`

Firma de QUE se lee de un archivo. Si cambia la hoja, la celda, la
columna o el filtro, la huella cambia y el cache no aplica.

#### `def leer_estado_mes(aamm)`

Lee el estado de cualquier mes sin tocar el estado en uso.

#### `def col_letra(n)`

4 -> "D".  Al reves de col_letra_a_num.

#### `def col_letra_a_num(letra)`

#### `def leer_columna_excel(ruta, hoja, columna, fila_inicio, col_filtro, valores_filtro, log)`

Suma una columna completa desde fila_inicio hacia abajo.
Si se indica col_filtro, suma solo las filas cuyo valor de esa columna esta
en valores_filtro (comparacion sin tildes ni mayusculas).
Devuelve (suma, n_filas, {valor_filtro: suma}) o (None, 0, {}).

#### `def es_zip_excel(ruta)`

#### `def ubicar_hoja_xml(z, hoja)`

Dentro del zip de un .xlsx/.xlsm, devuelve (ruta_del_xml, lista_de_hojas).
ruta_del_xml es None si la hoja no existe.

#### `def expandir_columnas(rango)`

'CF:CI' -> ['CF','CG','CH','CI'].  'CD' -> ['CD'].

#### `def buscar_marcas_rapido(ruta, hoja, fila_inicio, reglas, log, tope_detalle=30)`

Busca errores de fórmula y textos prohibidos en columnas puntuales,
escaneando el XML de la hoja por trozos en vez de cargar el libro.

reglas: [{"rangos": ["CF:CI"], "errores": True, "textos": ["REVISAR"]}, ...]
Devuelve {"conteo": {motivo: n}, "marcas": [(celda, motivo, valor)]} o None.

#### `def leer_celdas_rapido(ruta, hoja, celdas)`

Lee celdas puntuales de un .xlsx/.xlsm sin cargar el libro completo.

Un .xlsm es un ZIP con XML adentro. Se recorre el XML de la hoja en
streaming y se corta en cuanto se pasa de la ultima fila pedida, asi que
para celdas de las primeras filas (H1, EE6) casi no se lee nada, aunque el
archivo tenga miles de filas y millones de formulas.

Devuelve {celda: valor} con los valores YA CALCULADOS que Excel dejo
guardados. Si el archivo nunca fue calculado y guardado, no habra valores.
Devuelve None si no se pudo (formato distinto, hoja inexistente, etc.).

#### `def diagnosticar_celda(ruta, hoja, celda, log)`

Cuando una celda no entrega valor, explica POR QUE mirando el XML crudo:
si esta vacia, si tiene formula sin resultado guardado, si trae texto, si es
un error, o si es parte de una celda combinada.

#### `def resolver_hoja(nombres, hoja)`

Traduce lo pedido al nombre real de la hoja: acepta '#1' (por posicion) y
tolera tildes y mayusculas. Devuelve None si no calza ninguna.

#### `def es_significativo(v)`

Sirve para decidir si una celda 'cuenta' al buscar el ultimo dato de una
columna: se omiten vacios, ceros y errores de formula (#REF!, #N/D...).

#### `def leer_columnas_rapido(ruta, hoja, columnas, fila_inicio, log)`

Lee columnas completas de un .xlsx/.xlsm escaneando el XML por trozos.
Devuelve {"COL": {fila: valor}} con los valores ya calculados, o None.
Igual que en el resto, se lee SOLO el resultado y nunca el nodo <f>.

#### `def desescapar_xml(b)`

Convierte el texto crudo del XML de Excel a texto de verdad.

Hay que manejar las referencias NUMERICAS (&#243; = o con tilde), no solo las
cinco entidades con nombre: los nombres de empresa chilenos vienen llenos de
tildes y ñ, y algunos escritores de Excel las guardan asi. Si no se
desescapan, "Enel Generaci&#243;n" y "Enel Generación" no se parecen en nada
al comparar, y el cuadro de pago reporta un descuadre que no existe.

Se resuelve en UNA pasada a proposito. Reemplazar "&amp;" primero y despues
"&lt;" convertiria "&amp;lt;" (un literal "&lt;") en "<", que es otra cosa.

#### `def leer_formulas_rapido(ruta, hoja, columnas, fila_inicio, log)`

Devuelve {"COL": set(filas_que_TIENEN_formula)}.

A diferencia de leer_columnas_rapido, que lee el resultado, esto detecta la
PRESENCIA del nodo <f>, o sea si la celda es una formula o un valor escrito
a mano. Sirve para saber hasta donde se arrastro una formula.

Ojo con las formulas compartidas: la primera celda trae la formula completa
(<f t="shared" ref="L5:L120" si="0">...) y las siguientes solo <f t="shared"
si="0"/> sin texto. Como aca solo importa que exista un <f>, las dos formas
cuentan igual.

#### `def armar_tabla(datos, col_clave, cols_valor, log, etiqueta='', excluir=('0',), info=None)`

Convierte {"A":{fila:val}} en {clave_normalizada: (nombre, [valores])}.
Se salta las filas sin clave, que es lo que aparece como vacios al final.

excluir: claves normalizadas que NO son empresas y hay que descartar. Por
         defecto el "0", porque ninguna empresa se llama asi y se cuela
         cuando sobran formulas arrastrando ceros.
info:    dict opcional que se rellena con {"duplicadas", "excluidas",
         "vacias"} para que quien llame decida si eso es un fallo.

#### `def configurar_tags_log(widget)`

Deja el widget listo para recibir lineas con color.

#### `def clasificar_linea(linea, bloque=None)`

(tag, bloque_nuevo) para una linea del log.

bloque recuerda si venimos de un ">>" o de un "OK", para que las lineas
indentadas que siguen hereden el color. Devolver el bloque permite pintar
de a una linea (log en vivo) o una lista entera (ventana de detalle).

#### `def insertar_con_color(widget, lineas, bloque=None)`

Inserta lineas ya coloreadas. Devuelve el bloque en que quedo.

#### `def exportar_cprt(ruta_xlsm, ruta_csv, log)`

Escribe el csv del CPRT. Devuelve (n_filas_datos, n_con_retencion).

#### `def columnas_de_fila(ruta, hoja, fila, log)`

Letras de las columnas que tienen algo en esa fila, en orden.

Existe para no poner un tope fijo de columnas. La matriz del cuadro de pagos
tiene una columna por cada empresa que RECIBE, y eso cambia todos los meses:
con un tope escrito a mano, el mes que se pase se pierden receptores en
silencio y las verificaciones reportan faltantes que en realidad estan ahi.

#### `def leer_matriz_pago(ruta, hoja, log, fila_enc=8, col_ini='N')`

Lee la matriz cruzada. Devuelve dict con:
    pagan     : {nombre: fila}
    reciben   : {nombre: columna}
    montos    : {(pagador, receptor): monto}
    ultima_fila / ultima_col  (numero de columna)
o None si no se pudo leer.

Se descartan la fila y la columna de totales: CuadroPago escribe un "Total"
al final de cada lado, y esos no son pares de pago.

#### `def leer_nombre_definido(ruta, nombre)`

El rango al que apunta un nombre definido, tal como esta guardado.
Puede venir en A1 ($N$8:$T$89) o en R1C1 localizado (F8C14:F84C22).

#### `def celdas_de_rango(texto)`

(fila_fin, col_fin) del final de un rango, sea A1 o R1C1 localizado.
Devuelve None si no se entiende.

#### `def plan_traer_maestro(nodo, ruta_copia, ruta_maestro, diarios_copia, diarios_maestro, carpeta_copia)`

Arma la lista de operaciones, sin tocar el disco.

Devuelve (acciones, avisos). Cada accion es
    ("copiar"|"reemplazar"|"borrar", origen_o_None, destino, etiqueta)

Se separa del ejecutor a proposito: asi se puede mostrar al usuario exactamente
que se va a hacer ANTES de hacerlo, y se puede probar sin borrar nada.

#### `def ids_de_verificacion(vid)`

Nodos del arbol que un verificador necesita leer.

Se usa para releer del disco SOLO lo que hace falta antes de verificar. La
lectura completa recorre todas las carpetas de diarios, que en un disco de
red es lo que se lleva el tiempo, y para verificar un archivo no aporta nada.

Se juntan de todos lados para no dejar ninguno afuera: el archivo propio, los
"depende", los archivos que nombra cada comprobacion (incluidos los anidados
en referencia/lado_a/lado_b) y los de los VALORES que use.

#### `def ids_de_verificaciones(vids)`

Union de los nodos que necesitan varios verificadores, con sus previas.

#### `def ultimo_significativo(datos, columna)`

(fila, valor) del ultimo dato util de una columna, omitiendo vacios,
ceros y errores. (None, None) si no hay ninguno.

#### `def leer_valor_por_etiqueta(ruta, hoja, col_etiqueta, texto, col_valor, fila_inicio, log)`

Busca la fila cuya col_etiqueta diga `texto` y devuelve el numero que hay
en col_valor de esa misma fila. Sirve para tablas que cambian de largo: no
importa en que fila quede el "Total general", se lo busca por el rotulo.
Si aparece mas de una vez se usa la ULTIMA.

#### `def leer_valor_excel(ruta, hoja, celda, log)`

Lee una celda o la suma de un rango. Tres caminos, del mas rapido al mas
lento: streaming del XML, openpyxl, y por ultimo Excel via xlwings.

#### `def obtener_hojas(ruta)`

#### `def listar_hojas(ruta, log)`

#### `def conexion_mdb(ruta)`

#### `def listar_tablas_mdb(ruta, log)`

#### `def obtener_tablas_columnas(ruta)`

{tabla: [columnas]} de una base Access. {} si no se pudo abrir.

#### `def desglose_por_tipo(ruta, tabla, columna, columna_tipo, where, log)`

Escribe en el log la suma agrupada por tipo. Solo informativo.

#### `def leer_valor_mdb(ruta, tabla, columna, where, log)`

#### `def obtener_valor(clave, rutas, log, usar_cache=True)`

Devuelve (valor, mensaje_error_o_None).
Si el archivo no cambio desde la ultima lectura y se pide exactamente lo
mismo, devuelve el valor guardado sin abrir el archivo.

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`

> Actualiza el Access de la planilla 9
> (Ocupar_este_para_Reliquidacion_AAMM_*.mdb)
> Reemplaza a "para ricardo.py" + archivo_de_configuracion.yaml.
>
> Actualiza DOS tablas:
>
>   Sobrecostos      (Clave Año_Mes, Tipo_sobrecosto, Central, Hora Mensual,
>                     Sobrecosto), desde cuatro bloques de las planillas 3, 5 y 6
>   Central_Empresa  (Central, Empresa), desde tres hojas de propietarios
>
> DE DONDE SALE CADA DATO  —  replicado del script original
> Sobrecostos (4 bloques de 5 columnas cada uno):
>
>  | Fuente   | Archivo | Hoja                        | Cols    | Datos desde |
>  |----------|---------|-----------------------------|---------|-------------|
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `datetime`, `json`, `os`, `pathlib`, `queue`, `re`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `DIR_SCRIPT` | `Path(__file__).resolve().parent` |  |
| `CONFIG_PATH` | `DIR_SCRIPT.parent / 'config.json'` | config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto al Revisor).  |
| **— motor de Access, reutilizado —** | | |
| `_AYUDA` | `f'Los dos archivos tienen que estar en la misma carpeta y ser de la\nmisma versión. Copia…` |  |
| `_NECESITA` | `conjunto de 5 elementos: 'fuentes_externas', 'filtro_por_valores', 'borrar_todo', …` | Las cuatro hacen falta: borrar_todo para vaciar la tabla, cols_no_cero para el filtro de Central != 0, y las otras dos para pasar fuentes propias.  |
| `_TIENE` | `set(getattr(_ADA, 'CAPACIDADES', ()))` |  |
| **— CONFIGURACION —** | | |
| `TABLA_SOB` | `'Sobrecostos'` |  |
| `TABLA_CE` | `'Central_Empresa'` |  |
| `IDX_CLAVE` | `0` | Indices DENTRO del bloque de 5 columnas (0-based), segun el orden de la tabla: 0 Clave Año_Mes \| 1 Tipo_sobrecosto \| 2 Central \| 3 Hora Mensual \| 4 Sobrecosto |
| `IDX_CENTRAL` | `2` |  |
| `IDX_MONTO` | `4` |  |
| `RE_AAMM` | `re.compile('[_\\s](\\d{4})[_\\s]*[Rr]\\d')` | La Clave Año_Mes viene MAL desde el origen: siempre trae 23xx aunque el mes sea otro (2405 llega como 2305).  |
| `ENCABEZADOS_ESPERADOS` | `('clave', 'tipo', 'central', 'hora', 'pago')` | Encabezados que se esperan, solo para avisar si el archivo cambio. |
| `PLANILLAS` | `dict de 3 claves: 'p3', 'p5', 'p6', …` | Todo se organiza POR PLANILLA: una casilla por planilla, y al marcarla se actualizan sus bloques de Sobrecostos Y sus propietarios.  |
| `ORDEN_PL` | `['p3', 'p5', 'p6']` |  |
| `CARPETA_P9` | `'04 Planilla 9'` |  |
| **— TRASPASO DESDE EL REVISOR —** | | |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` |  |
| `TRASPASO_VERSION_MAX` | `1` |  |

### Funciones

#### `def aamm_de_nombre(ruta)`

El AAMM del nombre de un archivo: '..._2502_R01P.xlsm' -> 2502.
None si no se puede sacar.

#### `def detectar_aamm(rutas, log)`

El AAMM comun a los archivos. Si no coinciden entre si, lo dice: son
archivos de meses distintos y eso es un problema en si mismo.

**— CONFIG COMPARTIDO —**

#### `def get_usuario()`

#### `def leer_config()`

#### `def escribir_json(ruta, data)`

#### `def guardar_config(data)`

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def leer_traspaso(argv)`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

#### `def leer_propietarios(ruta, cfg, log)`

(Central, Empresa) de una hoja de propietarios, SIN abrir Excel.

El largo lo manda la columna de la CENTRAL: se corta en su primera celda
vacia aunque la de empresa siga con datos. En CONSUMOS_PROPIOS las dos
columnas no son contiguas (B y H) y HAY centrales sin propietario, que se
conservan con la empresa en None.

#### `def revisar_encabezados(ruta, cfg, log)`

Lee la fila de encabezado del bloque y avisa si no se parece a lo
esperado. No corta el proceso: los datos se leen por POSICION, asi que un
encabezado distinto no rompe nada, pero conviene saberlo.

#### `def cargar_central_empresa(ruta_mdb, filas, solo_lectura, log)`

Vacia Central_Empresa y carga UNA fila por central.

filas: lista de (central, empresa, origen). El origen es solo para poder
decir de que planilla vino cada una cuando hay conflicto.
Devuelve (ok, n_insertadas, n_sin_dueno).

Transaccion unica: si algo falla, se revierte y la tabla queda como estaba.
Las centrales SIN propietario se cargan igual, con la empresa en NULL: son
datos validos. El Revisor comprueba despues si alguna de esas tiene plata.

#### `def clave_central(t)`

Normaliza el nombre de una central para comparar: sin tildes, sin espacios
ni guiones bajos, en mayusculas. Asi 'El Toro-1', 'EL_TORO-1' y 'ELTORO-1'
son la misma.

Es MAS estricto que la comparacion de Access (que ignora mayusculas pero no
espacios), y eso conviene: evita mandar dos filas que el indice unico
consideraria iguales. La misma funcion esta en el Revisor.

#### `def columnas_tabla_de(cur, tabla)`

Nombres de columna de una tabla cualquiera del Access.

#### `def guardar_excel_dump(ruta_mdb, tabla, encabezados, filas, log)`

Guarda en un Excel lo que se cargo, al lado del .mdb.

Es el equivalente del df_ricardo_salida.xlsx del script viejo, pero con el
nombre de la tabla y del mes, y junto al Access en vez del directorio de
trabajo del momento.

#### `def ejecutar(rutas, seleccion, solo_lectura, guardar_dump, log, progreso, aamm=None)`

seleccion: lista de planillas ('p3','p5','p6'). Por cada una se cargan sus
bloques de Sobrecostos Y sus propietarios.

aamm: el mes que se escribe en la Clave Año_Mes de TODAS las filas. Si no
viene, se saca del nombre de los archivos. Nunca se usa el valor del origen,
que viene mal (siempre 23xx).

Devuelve (ok, resumen).

**— VENTANA —**

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Actualiza_Cuadro0.py`

> Actualiza Cuadro 0   (0_CUADROS_RELIQUIDACION SSCC)
> ALCANCE: deja los DATOS y las FORMULAS al dia, nada mas.
>     - pega la tabla del 1_CUADROS
>     - la tasa de O2, si se le da una
>     - reescribe las listas de empresas (K de la hoja #3, I de 01.SSCC)
>     - estira o corta las formulas de L:P, R:U, A:G y J:K
> NO llama a Cuadro de pagos, NO llama a Actualiza Rango y NO refresca la tabla
> dinamica de CPRT: eso se hace a mano en Excel despues de correr esto. La
> dinamica cuelga de una Power Query ("Consulta - TEE") que depende de un libro
> externo, y automatizarla escondia los errores en vez de resolverlos.
>
> ES EL ARCHIVO QUE SE VA A PAGO, asi que cada paso queda escrito en el log con
> lo que habia antes y lo que quedo despues.
>
> COMO ESTA ARMADO EL LIBRO  (lo que hay que entender antes de tocar nada)
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `datetime`, `json`, `os`, `pathlib`, `queue`, `re`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `DIR_SCRIPT` | `Path(__file__).resolve().parent` |  |
| `CONFIG_PATH` | `DIR_SCRIPT.parent / 'config.json'` | config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto al Revisor).  |
| **— Hojas y celdas —** | | |
| `IDX_HOJA_3` | `2` |  |
| `HOJA_SSCC` | `'01.SSCC_Recurso_Técnico'` |  |
| `HOJA_ORIGEN_CUADRO1` | `'01.SSCC_Recurso_Técnico'` |  |
| `CELDA_TASA` | `'O2'` |  |
| `ORIGEN_TABLA` | `('I', 'K')` | Tabla que se pega del 1_CUADROS: I9:K de su hoja -> A5:C de la hoja #3 |
| `ORIGEN_TABLA_FILA` | `9` |  |
| `DESTINO_TABLA` | `('A', 'C')` |  |
| `DESTINO_TABLA_FILA` | `5` |  |
| `BLOQUES_TABLA` | `[('D', 'D')]` | La D de la hoja #3 es una formula que acompaña a la tabla A:C: D5 = B5 + C5 (el neto por empresa: lo que paga mas lo que recibe) O sea que su largo lo manda la TABLA PEGADA, no la lista K.  |
| `FILA_K` | `5` | --- Formulas que se escriben ---------------------------------------------- IMPORTANTE: por COM las formulas se escriben en INGLES y con COMA como separador, aunque en pantalla se vean en espanol con… |
| `FORMULA_K` | `'=LET(x,UNIQUE(VSTACK(F{f}:F{tope},A{f}:A{tope})),FILTER(x,(x<>0)*(x<>"")))'` |  |
| `TOPE_K` | `1000` |  |
| `FILA_I` | `9` |  |
| `FORMULA_I` | `'=LET(x,UNIQUE(C{f}:C{tope}),FILTER(x,(x<>0)*(x<>"")))'` |  |
| `TOPE_I` | `4345` |  |
| `BLOQUES` | `lista de 3 elementos: ('#3', 5, [('L', 'P'), ('R', 'U')], 'K'), ('SSCC', 9, [('A', 'G')], 'K'), ('SSCC', 9, [('J', 'K')], 'I'), …` | --- Bloques de formulas que hay que estirar o cortar ---------------------- (hoja, primera_fila, [(col_ini, col_fin), ...], que define el largo) "K" -> tantas filas como empresas haya en K de la hoja… |
| **— TRASPASO DESDE EL REVISOR —** | | |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` |  |
| `TRASPASO_VERSION_MAX` | `1` |  |

### Funciones

**— CONFIG COMPARTIDO —**

#### `def get_usuario()`

#### `def leer_config()`

#### `def escribir_json(ruta, data)`

#### `def guardar_config(data)`

#### `def leer_tasa_guardada(cfg, aamm)`

(valor, fecha) de la tasa guardada para ese mes, o (None, None).

Se guarda POR MES a proposito: la tasa es del periodo. Guardada suelta, el
mes siguiente arrastraria la del anterior sin que se note.

#### `def guardar_tasa(aamm, valor)`

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def leer_traspaso(argv)`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

**— UTILIDADES —**

#### `def normalizar(t)`

#### `def buscar_hoja(wb, nombre)`

Busca la hoja tolerando tildes y mayusculas.

#### `def col_num(letra)`

#### `def ultima_fila(sh, col, desde=1)`

Ultima fila con ALGO en la columna, mirando de abajo hacia arriba.
Cuenta las cadenas vacias como contenido: para el desbordamiento un "" es
tan estorbo como un numero.

#### `def recalcular(app, log, motivo)`

Recalcula UNA vez y dice cuanto tardo.

El libro se queda en calculo MANUAL de punta a punta. Con automatico, cada
AutoFill dispara un recalculo completo, y las formulas de aca son caras: L5
es =SUMIF(A:A,K5,D:D), un SUMIF de COLUMNA ENTERA repetido por cada empresa,
y J9/K9 evaluan su SUMIF dos veces (una en la condicion y otra en el
resultado). Recalcular cuatro veces a proposito, en vez de una por cada
escritura, es la diferencia entre segundos y minutos.

Se cronometra y queda en el log: si algun mes se pone lento, se ve donde.

#### `def fmt_tiempo(seg)`

#### `def paso_pegar_tabla(sh1, sh3, log)`

I9:K del cuadro 1 -> A5:C de la hoja #3. Solo valores.

#### `def paso_tasa(sh3, valor, log)`

#### `def paso_formula_desbordada(sh, celda, formula, col, fila, log, etiqueta)`

Limpia la columna y escribe la formula. Devuelve la ultima fila ocupada.

Limpiar PRIMERO es obligatorio: cualquier celda no vacia debajo hace que la
formula desbordada tire #DESBORDAMIENTO, y al reliquidar el archivo llega
con los datos del mes anterior.

#### `def paso_estirar(sh, fila_ini, bloques, hasta, log)`

Deja las formulas de esos bloques cubriendo exactamente hasta 'hasta'.

#### `def paso_validar_signos(sh, log)`

AVISO, no corta. Comprueba lo mismo que valida CuadroPago antes de armar
la matriz: en la tabla I:K, todo numero de la columna J (paga) tiene que ser
NEGATIVO y todo numero de la K (recibe) POSITIVO.

Se avisa aca porque CuadroPago ATRAPA su propio error: muestra un MsgBox y
sale con Exit Sub sin escribir la matriz. Corriendola a mano se ve el cartel,
pero conviene saber de antemano si va a fallar. Devuelve True si esta OK.

#### `def ejecutar(rutas, op, log, progreso)`

Corre la secuencia completa. Lo unico opcional es la tasa, porque puede
llegar al final del mes, y guardar.

op: {"tasa": float o None para no tocarla, "guardar": bool}
Devuelve (ok, resumen).

**— VENTANA —**

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py`

> Actualiza la tabla [Sobrecostos] de un Access .mdb consolidando la informacion
> de tres archivos Excel con macros (.xlsm).
>
> Estructura de carpetas esperada:
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `datetime`, `decimal`, `json`, `os`, `pathlib`, `queue`, `re`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `FUENTES` | `dict de 3 claves: 'SSCC', 'CO', 'CCA', …` | base: de donde cuelga la "carpeta" de cada fuente.  |
| `ORDEN_FUENTES` | `['SSCC', 'CO', 'CCA']` |  |
| `CAPACIDADES` | `frozenset({'fuentes_externas', 'filtro_por_valores', 'borrar_todo', 'cols_no_cero', 'forz…` | Capacidades que este modulo le ofrece a quien lo importe (Actualiza_Energia.py).  |
| `TABLA_ACCESS` | `'Sobrecostos'` |  |
| `COLUMNAS_ESPERADAS` | `['Clave Año_Mes', 'Tipo_sobrecosto', 'Central', 'Hora Mensual', 'Sobrecosto']` |  |
| `CONFIG_PATH` | `Path(__file__).resolve().parent.parent / 'config.json'` | config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto al Revisor).  |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` | TRASPASO DESDE EL REVISOR El Revisor escribe un JSON en Salidas/AAMM/ y pasa su ruta como argv[1].  |
| `TRASPASO_VERSION_MAX` | `1` |  |
| `NS_XL` | `'{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'` | Lectura rapida: el .xlsx/.xlsm como ZIP, sin abrir Excel Las planillas son pesadas y aca solo hay que LEERLAS.  |
| `NS_REL` | `'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'` |  |
| `_ENT_XML` | `{'lt': '<', 'gt': '>', 'quot': '"', 'apos': "'", 'amp': '&'}` |  |
| `_RE_ENT` | `re.compile('&(?:#(\\d+)\|#[xX]([0-9a-fA-F]+)\|(lt\|gt\|quot\|apos\|amp));')` |  |

### Funciones

**— CONFIG POR PC/USUARIO —**

#### `def get_usuario()`

#### `def leer_config()`

#### `def escribir_json(ruta, data)`

Escritura atomica: primero un .tmp y despues os.replace.
Evita dejar el archivo truncado si algo falla a medio camino.

#### `def guardar_config(data)`

#### `def leer_traspaso(argv)`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

**— UTILIDADES —**

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def normalizar(texto)`

#### `def buscar_carpeta(base, nombre)`

Busca una subcarpeta comparando nombres normalizados (tildes/mayusculas).

#### `def buscar_archivo(carpeta, patron_regex, extensiones=('.xlsm', '.xlsx', '.xlsb'))`

Archivo mas reciente cuyo nombre normalizado calza con el patron.

#### `def col_letra(n)`

4 -> "D".  Al reves de col_letra_a_num.

#### `def col_letra_a_num(letra)`

#### `def fmt_tiempo(seg)`

#### `def es_vacio(v)`

#### `def es_cero(v)`

**— LECTURA DE EXCEL (xlwings, solo lectura) —**

#### `def buscar_hoja(wb, nombre)`

#### `def ultima_fila(sheet, col_num, desde)`

Ultima fila con contenido en una columna, usando .formula (capta formulas).

#### `def leer_bloque(sheet, f1, c1, f2, c2)`

Devuelve siempre una lista de filas (listas).

#### `def es_zip_excel(ruta)`

#### `def ubicar_hoja_xml(z, hoja)`

Dentro del zip de un .xlsx/.xlsm, devuelve (ruta_del_xml, lista_de_hojas).
ruta_del_xml es None si la hoja no existe.

#### `def expandir_columnas(rango)`

'CF:CI' -> ['CF','CG','CH','CI'].  'CD' -> ['CD'].

#### `def leer_columnas_rapido(ruta, hoja, columnas, fila_inicio, log)`

Lee columnas completas de un .xlsx/.xlsm escaneando el XML por trozos.
Devuelve {"COL": {fila: valor}} con los valores ya calculados, o None.
Igual que en el resto, se lee SOLO el resultado y nunca el nodo <f>.

#### `def desescapar_xml(b)`

Convierte el texto crudo del XML de Excel a texto de verdad.

Hay que manejar las referencias NUMERICAS (&#243; = o con tilde), no solo las
cinco entidades con nombre: los nombres de empresa chilenos vienen llenos de
tildes y ñ, y algunos escritores de Excel las guardan asi. Si no se
desescapan, "Enel Generaci&#243;n" y "Enel Generación" no se parecen en nada
al comparar, y el cuadro de pago reporta un descuadre que no existe.

Se resuelve en UNA pasada a proposito. Reemplazar "&amp;" primero y despues
"&lt;" convertiria "&amp;lt;" (un literal "&lt;") en "<", que es otra cosa.

#### `def leer_matriz_rapida(ruta, cfg, log)`

La matriz de la fuente leyendo el ZIP, sin Excel. None si no se pudo.

#### `def leer_fuente(app, ruta_xlsm, cfg, log)`

La matriz de 5 columnas de una fuente, ya filtrada.

Primero intenta el lector rapido (ZIP + XML, sin Excel). Si el archivo no es
un .xlsx/.xlsm o algo falla, cae a xlwings.

`app` puede ser una instancia de xlwings, None, o una FUNCION que la crea al
llamarla. Lo ultimo es lo que usa proceso(): asi Excel se abre solo si de
verdad hace falta, que con planillas normales no pasa nunca.

#### `def driver_access()`

El nombre del driver ODBC de Access. Lanza con un mensaje util si no hay.

El mensaje LISTA los drivers que pyodbc si ve, y dice que Python se esta
usando. Sin eso no se puede distinguir entre "no hay ningun driver ODBC"
(instalacion rota o pyodbc mal), "hay drivers pero ninguno de Access"
(falta el Access Database Engine) y "hay uno de Access pero de otra
arquitectura" (Office de 32 bits con Python de 64).

Tambien importa cuando el mismo script funciono antes en el mismo PC: casi
siempre significa que se lanzo con OTRO Python (otra instalacion, otro
entorno virtual, o pythonw vs python), y por eso se muestra el ejecutable.

#### `def conectar_access(ruta_mdb)`

#### `def columnas_tabla(cur)`

Devuelve [(nombre, tipo_python)] de la tabla, en orden.

#### `def mapear_columnas(cols_tabla, log)`

Elige las 5 columnas destino: por nombre si calzan, si no las 5 primeras.

#### `def a_numero(valor)`

Convierte a int/float lo que venga (numero, o texto con formato chileno/ingles).
Devuelve None si no se puede.

#### `def clave_a_numero(valor)`

La columna 'Clave Año_Mes' SIEMPRE se entrega como numero (AAAAMM).
Acepta 202506, '202506', '2025-06', '06/2025', ' 202.506 ' o una fecha.

#### `def coercionar(valor, tipo)`

Adapta el valor leido de Excel al tipo de la columna en Access.

#### `def coercionar_clave(valor, tipo)`

Igual que coercionar pero forzando la clave a numero, sea cual sea el tipo destino.

#### `def armar_lote(filas, destino, log)`

Convierte las filas leidas a los tipos del Access. La columna 1 (clave) va a numero.
Devuelve (lote, excluidas).

#### `def clave_tipo(t)`

Clave homogenea para comparar tipos entre el Access y los Excel.

#### `def estado_access(cur, col_tipo, col_valor, con_suma=True)`

{clave_tipo: {'n': filas, 'suma': monto o None}} agrupado por Tipo_sobrecosto.

#### `def fmt_int(n, signo=False)`

#### `def fmt_monto(v, signo=False)`

#### `def resumen_cambios(antes, despues, log, proyectado=False)`

Detalle de diferencias por tipo y total, antes vs despues (o proyectado).

#### `def proceso(ruta_mdb, archivos, seleccion, solo_lectura, log, progreso, fuentes=None, borrar_todo=False)`

archivos: {clave: ruta_xlsm}; seleccion: lista de claves a actualizar.

fuentes: diccionario de configuracion de las fuentes. Por omision el FUENTES
de este modulo (SSCC/CO/CCA); Actualiza_Energia.py y Actualiza_Access_P9.py
pasan el suyo para reusar todo este motor de Access sin duplicarlo.

borrar_todo: vaciar la tabla entera en vez de borrar solo los
Tipo_sobrecosto que vienen en los Excel. Lo usa el .mdb de la planilla 9,
que se arma COMPLETO desde sus fuentes; ahi el borrado por tipo dejaria
filas viejas si un tipo desaparece del origen. El vaciado ocurre UNA sola
vez, antes de insertar (ver ya_vaciada mas abajo).

Si se agrega un parametro nuevo aca, hay que sumarlo tambien a CAPACIDADES,
o los scripts que lo usen se caen con un TypeError en vez de decir "copia el
archivo actualizado".

**— VENTANA —**

#### `def actualizar_color_label(lbl, valor, es_archivo=False)`

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Actualiza_Energia.py`

> Actualizar Energia
> Actualiza los dos entregables de "01.a Sobrecostos de Energia" a partir del
> "02 Consolidado_Tabulado", hoja "Sobrecostos":
>
>   1) 03b ENTRADA_SOB_AAMM_*.mdb   (tabla [Sobrecostos] del Access)
>   2) Consolidado_AAMM_*.xlsm      (hoja "Sobrecostos")
>
> En los dos casos se toman SOLO las filas cuyo Tipo sobrecosto es SCMT o SCPC,
> que es justo lo que el revisor comprueba en V5 y V6.
>
> El motor de Access se reutiliza de Actualiza_Data_Access.py en vez de
> copiarlo: es el mismo destino (tabla [Sobrecostos], mismas 5 columnas, misma
> regla de "borrar solo los tipos que trae el Excel"). Los dos .py tienen que
> estar en la misma carpeta.
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `json`, `os`, `pathlib`, `queue`, `re`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `DIR_SCRIPT` | `Path(__file__).resolve().parent` |  |
| `CONFIG_PATH` | `DIR_SCRIPT.parent / 'config.json'` | config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto al Revisor).  |
| `_AYUDA_COPIAR` | `f'Los dos archivos tienen que estar en la misma carpeta y ser de la misma\nversion. Copia…` | --- motor de Access, reutilizado ------------------------------------------ Este script NO duplica el motor de Access: usa el de Actualiza_Data_Access.py, que tiene que estar en la MISMA carpeta y se… |
| `_NECESITA` | `{'fuentes_externas', 'filtro_por_valores'}` |  |
| `_TIENE` | `set(getattr(_ADA, 'CAPACIDADES', ()))` |  |
| `FILA_DATOS_TABULADO` | `3` | CONFIGURACION El encabezado del "02 Consolidado_Tabulado" (hoja Sobrecostos) esta en la fila 2, asi que los datos arrancan en la 3.  |
| `TIPOS` | `('SCMT', 'SCPC')` | Tipos que se traen.  |
| `HOJA_ORIGEN` | `'Sobrecostos'` |  |
| `HOJA_DESTINO_CONSOLIDADO` | `'Sobrecostos'` |  |
| `FUENTES_ENERGIA` | `dict de 1 claves: 'MDB', …` | ---- 1) Access ------------------------------------------------------------- Columnas AA:AE del tabulado, que ya vienen en el mismo orden que la tabla: AA Clave Año_Mes \| AB Tipo sobrecosto \| AC Cent… |
| `BLOQUES_CONSOLIDADO` | `[('A', 'G'), ('I', 'J'), ('H', 'H')]` | ---- 2) Consolidado_AAMM --------------------------------------------------- Bloques de origen EN ESTE ORDEN -> se pegan corridos en A:J del destino.  |
| `COL_FILTRO_CONSOLIDADO` | `'AB'` |  |
| `COL_DESTINO_INI` | `'A'` |  |
| `FILA_DESTINO_INI` | `2` |  |
| `N_COLS_CONSOLIDADO` | `10` |  |
| **— TRASPASO DESDE EL REVISOR —** | | |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` |  |
| `TRASPASO_VERSION_MAX` | `1` |  |

### Funciones

**— CONFIG COMPARTIDO  (mismo config.json que el Revisor y los otros dos) —**

#### `def get_usuario()`

#### `def leer_config()`

#### `def escribir_json(ruta, data)`

#### `def guardar_config(data)`

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def leer_traspaso(argv)`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

#### `def buscar_tabulado(carpeta_reliq)`

01 Sobrecostos/<Detalles diarios>/02 Consolidado_Tabulado_AAMM_*

#### `def carpeta_energia(carpeta_reliq)`

#### `def buscar_mdb_energia(carpeta_reliq)`

#### `def buscar_consolidado_energia(carpeta_reliq)`

#### `def actualizar_consolidado(app_xw, ruta_tabulado, ruta_destino, log, progreso=None)`

Pega en el Consolidado_AAMM las filas SCMT/SCPC del tabulado.
Devuelve la cantidad de filas escritas.

#### `def ejecutar(rutas, hacer_mdb, hacer_consolidado, solo_prueba, log, progreso)`

rutas: {"tabulado","mdb","consolidado"}. Devuelve (ok, resumen:str).

**— VENTANA —**

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py`

> Actualiza la hoja "SC y CO" de la planilla 5_
> Pega los SC y los CO de los EMBALSES, con su prorrata de instruccion directa.
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `datetime`, `json`, `os`, `pathlib`, `queue`, `re`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `DIR_SCRIPT` | `Path(__file__).resolve().parent` |  |
| `CONFIG_PATH` | `DIR_SCRIPT.parent / 'config.json'` | config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto al Revisor).  |
| `CENTRALES_EMBALSE` | `lista de 27 elementos: 'CANUTILLAR-1', 'CANUTILLAR-2', 'ELTORO-1', …` | Centrales de embalse OJO: esta lista esta TAMBIEN en Revisor_Reliquidacion.py.  |
| **— Configuracion de origenes y destino —** | | |
| `HOJA_DESTINO` | `'SC y CO'` |  |
| `FILA_DESTINO` | `9` |  |
| `COLS_ID` | `('C', 'G')` | Bloque de identificacion (C:G) y de prorrata (I:X), y las formulas (Y:AF). |
| `COLS_PRO` | `('I', 'X')` |  |
| `BLOQUES_FORMULA` | `[('Y', 'AB'), ('AD', 'AF')]` | Bloques de formulas, EN BLOQUES y no un rango corrido: la AC queda AFUERA a proposito.  |
| `COL_TIPO` | `'D'` |  |
| `COL_CLAVE` | `'C'` |  |
| `COL_CENTRAL` | `'E'` |  |
| `N_ID` | `5` |  |
| `N_PRO` | `16` |  |
| `TIPO_SC` | `'SCCF'` |  |
| `TIPO_CO` | `'CO'` |  |
| `FUENTES` | `dict de 2 claves: 'SC', 'CO', …` |  |
| `ORDEN` | `['SC', 'CO']` |  |
| `RE_UNIDAD` | `re.compile('-\\d+\\s*$')` | Una central "-numero" que no este en la lista es sospechosa: el sufijo indica unidad, y las unidades son justamente lo que tienen los embalses. |
| **— TRASPASO DESDE EL REVISOR —** | | |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` |  |
| `TRASPASO_VERSION_MAX` | `1` |  |

### Funciones

#### `def leer_config()`

#### `def guardar_config(data)`

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def leer_traspaso(argv)`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

**— UTILIDADES —**

#### `def normalizar(t)`

#### `def clave_central(t)`

Normaliza un nombre de central para comparar: sin tildes, sin espacios ni
guiones bajos, en mayusculas. Asi 'El Toro-1' y 'ELTORO-1' son la misma.

#### `def buscar_hoja(wb, nombre)`

#### `def col_num(letra)`

#### `def col_letra(n)`

#### `def ultima_fila(sh, col, desde=1)`

#### `def fmt_tiempo(seg)`

#### `def leer_fuente(app, ruta, cfg, log)`

Devuelve (filas_id, filas_pro, avisos).

filas_id: lista de listas de 5 valores (C:G). Si la fuente no trae la Clave
          Año_Mes, el primer valor queda en None y se completa despues.
filas_pro: lista de listas de 16 valores (I:X), alineada con filas_id.

#### `def leer_bloque_destino(sh, tipo, log)`

Lee del propio destino las filas de un tipo (SCCF o CO).

Sirve para conservar el bloque que el usuario NO eligio actualizar: como los
dos bloques van uno debajo del otro, hay que reescribirlos juntos.

#### `def ejecutar(rutas, hacer, aamm_respaldo, log, progreso)`

hacer: subconjunto de ["SC", "CO"]. Devuelve (ok, resumen).

**— VENTANA —**

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Actualiza_datos.py`

**Importa:** `json`, `os`, `pathlib`, `re`, `socket`, `subprocess`, `sys`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| **— Constantes configurables —** | | |
| `INSTRUCCIONES` | `"Selecciona la carpeta '02 CASO RELIQUIDACION'.\nEl script detecta automáticamente todos …` |  |
| `CONFIG_PATH` | `Path(__file__).resolve().parent.parent / 'config.json'` | ── Config por usuario/PC ─────────────────────────────────────────────────── config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto a… |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` | ── Traspaso desde el Revisor ─────────────────────────────────────────────── El Revisor escribe un JSON en Salidas/AAMM/ y pasa su ruta como argv[1].  |
| `TRASPASO_VERSION_MAX` | `1` |  |
| `MAPEO_SOBRECOSTOS_FD` | `lista de 6 elementos: {'hoja_origen': 'CT Diario', 'cols_origen': [('D', 'I')], 'fila_ini_origen': 12, 'fila_det_origen': 'D', 'hoja_destino': 'FD_CT', 'cols_destino': [('C', 'H')], 'fila_ini_destino': 12, 'fila_det_destino': 'C', 'cols_formulas': [('B', 'B')]}, {'hoja_origen': 'CPF Horario', 'cols_origen': [('B', 'J')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'FD_CPF', 'cols_destino': [('C', 'K')], 'fila_ini_destino': 12, 'fila_det_destino': 'C', 'cols_formulas': [('B', 'B'), ('L', 'P')]}, {'hoja_origen': 'CSF Horario', 'cols_origen': [('B', 'H')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'FD_CSF', 'cols_destino': [('C', 'I')], 'fila_ini_destino': 12, 'fila_det_destino': 'C', 'cols_formulas': [('A', 'B'), ('J', 'M')]}, …` |  |
| `MAPEO_SOBRECOSTOS_CONSOLIDADO` | `lista de 1 elementos: {'hoja_origen': 'Sobrecostos', 'cols_origen': [('A', 'G'), ('I', 'J'), ('Q', 'W')], 'fila_ini_origen': 3, 'fila_det_origen': 'A', 'hoja_destino': 'SOBRECOSTOS', 'cols_destino': [('A', 'G'), ('I', 'J'), ('K', 'Q')], 'fila_ini_destino': 7, 'fila_det_destino': 'A', 'cols_formulas': [('R', 'EB')], 'filtro_col': 'C', 'filtro_valor': 'C.Frec', 'detectar_fin_primera_vacia': True, 'ajustar_formulas': True}, …` |  |
| `MAPEO_P3_FD` | `lista de 3 elementos: {'hoja_origen': 'CPF Horario', 'cols_origen': [('B', 'J')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'CPF_FD', 'cols_destino': [('D', 'L')], 'fila_ini_destino': 9, 'fila_det_destino': 'D', 'cols_formulas': [('A', 'B'), ('N', 'P')]}, {'hoja_origen': 'CSF Horario', 'cols_origen': [('B', 'H')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'CSF_FD', 'cols_destino': [('E', 'K')], 'fila_ini_destino': 9, 'fila_det_destino': 'E', 'cols_formulas': [('A', 'D'), ('M', 'N'), ('P', 'X')]}, {'hoja_origen': 'CTF Horario', 'cols_origen': [('B', 'I')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'CTF_FD', 'cols_destino': [('E', 'L')], 'fila_ini_destino': 9, 'fila_det_destino': 'E', 'cols_formulas': [('A', 'D'), ('N', 'P'), ('R', 'Y')]}, …` |  |
| `MAPEO_P5_FD` | `lista de 3 elementos: {'hoja_origen': 'CPF Horario', 'cols_origen': [('B', 'J')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'FD_CPF', 'cols_destino': [('B', 'J')], 'fila_ini_destino': 7, 'fila_det_destino': 'B', 'cols_formulas': [('A', 'A')]}, {'hoja_origen': 'CSF Horario', 'cols_origen': [('B', 'H')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'FD_CSF', 'cols_destino': [('B', 'H')], 'fila_ini_destino': 7, 'fila_det_destino': 'B', 'cols_formulas': [('A', 'A')]}, {'hoja_origen': 'CTF Horario', 'cols_origen': [('B', 'I')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'FD_CTF', 'cols_destino': [('B', 'I')], 'fila_ini_destino': 7, 'fila_det_destino': 'B', 'cols_formulas': [('A', 'A')]}, …` |  |
| `MAPEO_P6_FD` | `lista de 1 elementos: {'hoja_origen': 'CT Diario', 'cols_origen': [('B', 'B'), ('D', 'I')], 'fila_ini_origen': 12, 'fila_det_origen': 'B', 'hoja_destino': 'FD', 'cols_destino': [('B', 'H')], 'fila_ini_destino': 7, 'fila_det_destino': 'B', 'cols_formulas': []}, …` |  |
| `PRORRATA_CONFIG` | `dict de 9 claves: 'hoja_origen', 'fila_ini_origen', 'col_hora', …` | ── Configuración Prorrata (pivot tabla→matriz) ────────────────────────────── Origen: hoja PRORRATA_HORARIA_TABULAR, fila 2 en adelante Col A: Hora Mensual \| Col B: Suministrador \| Col C: Prorrata_ho… |

### Funciones

#### `def get_usuario() -> str`

#### `def leer_config() -> dict`

#### `def escribir_json(ruta, data)`

Escritura atomica: primero un .tmp y despues os.replace.
Evita dejar el archivo truncado si algo falla a medio camino.

#### `def guardar_config(data: dict)`

#### `def leer_traspaso(argv: list) -> dict | None`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

**— Utilidades —**

#### `def abrir_en_explorador(ruta: str, es_archivo: bool=False)`

#### `def normalizar(texto: str) -> str`

#### `def buscar_sscc_desempeno(carpeta_reliq: Path) -> Path | None`

FD está un nivel arriba de 02 CASO RELIQUIDACION.

#### `def buscar_consolidado(carpeta_reliq: Path) -> Path | None`

01 Sobrecostos/Detalles diarios/02 Consolidado_Tabulado_AAMM…

#### `def buscar_sobrecostos_xlsm(carpeta_reliq: Path) -> Path | None`

#### `def buscar_planilla9(carpeta_reliq: Path, prefijo: str) -> Path | None`

#### `def buscar_prorrata(carpeta_reliq: Path) -> Path | None`

04 Planilla 9/Prorrata_Retiros_AAMM…

**— Lógica xlwings —**

#### `def col_letra_a_num(letra: str) -> int`

#### `def expandir_cols(rangos: list) -> list`

#### `def ultima_fila(sheet, col_letra: str, desde: int) -> int`

Última fila con valor O fórmula en la columna, desde 'desde'.

#### `def sheet_last_row_col(sheet, col_num: int, desde: int) -> int`

Última fila con valor o fórmula en columna por número, desde 'desde'.

#### `def ultima_fila_usedrange(sheet, fila_ini: int) -> int`

Última fila usada en la hoja según UsedRange (incluye fórmulas y formatos).

#### `def fmt_tiempo(segundos)`

#### `def aplicar_mapeo(app_xw, wb_o, wb_d, mapeo: list, log_func=print, progreso_func=None)`

Aplica una lista de mapeos entre wb_o (origen) y wb_d (destino).

#### `def aplicar_prorrata(app_xw, wb_o, wb_d, cfg: dict, log_func=print, progreso_func=None)`

Transforma tabla larga (Hora,Suministrador,Valor) a matriz pivote en destino.

#### `def ejecutar_actualizacion(carpeta_reliq: Path, planilla: str, hacer_fd: bool, hacer_otro: bool, log_func=print, progreso_func=None, rutas: dict | None=None) -> tuple[bool, list]`

planilla: 'sc' | 'p3' | 'p5' | 'p6'
hacer_fd: actualizar FD
hacer_otro: para sc=Consolidado; p3/p5/p6=Prorrata (próximamente)
rutas: dict del traspaso del Revisor. Las claves que vengan se usan tal
       cual y NO se vuelve a buscar el archivo; las que falten caen al
       buscar_* de siempre.
Retorna (ok, lista_rutas_modificadas)

**— Ventana —**

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Carga_Retiros.py`

> Carga Retiros_h.parquet a SQL Server
> Version con ventana del script original. Lo que hace es lo mismo:
>     1. lee el parquet
>     2. BORRA del servidor los periodos que trae el parquet (o la tabla entera)
>     3. carga por trozos
>     4. verifica que la cuenta cuadre
>
> ESTO ESCRIBE EN UNA BASE DE DATOS Y BORRA FILAS. Por eso:
>   - antes de borrar muestra exactamente que va a borrar y pide confirmacion
>   - hay un modo SOLO MIRAR que hace los conteos y no toca nada
>   - todo queda en el log con las cuentas antes y despues
>
> SI FALLA A MITAD DE LA CARGA: el borrado ya se hizo y quedan filas a medias.
> No es grave: volver a correrlo borra ese periodo otra vez y recarga. Lo que NO
> hay que hacer es dejarlo asi, porque la tabla queda con datos incompletos.
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `datetime`, `json`, `os`, `pathlib`, `queue`, `re`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `DIR_SCRIPT` | `Path(__file__).resolve().parent` |  |
| `CONFIG_PATH` | `DIR_SCRIPT.parent / 'config.json'` | config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto al Revisor).  |
| **— Configuracion —** | | |
| `NOMBRE_PARQUET` | `'Retiros_h.parquet'` |  |
| `CARPETA_PARQUET` | `'04 Planilla 9'` |  |
| `TABLA` | `'Retiros'` |  |
| `SERVER` | `'SRV-DTE'` |  |
| `DRIVER` | `'ODBC Driver 17 for SQL Server'` |  |
| `BASES` | `['02_RETIROS', '14_RETIROS_RELIQUIDACION']` | Las bases entre las que se puede elegir.  |
| `CHUNK` | `50000` |  |
| `LARGO_TEXTO` | `255` | Largo de las columnas de texto SI HAY QUE CREAR la tabla.  |
| `COL_PERIODO` | `'Clave Año_Mes'` | Los nombres de columna NO se escriben igual en todos lados: el parquet a veces trae "Clave Año_Mes" (con espacio y ñ) y a veces "Clave_Anio_Mes" o "Clave_anio_mes", y la tabla de SQL Server tiene el… |
| `COL_SUMINISTRADOR` | `'Suministrador'` |  |
| `COL_HORA` | `'Hora Mensual'` |  |
| `HORA_CAMBIO_POR_OMISION` | `145` | El mes del cambio de hora de primavera tiene una hora MENOS: esa hora no existe.  |
| **— TRASPASO DESDE EL REVISOR —** | | |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` |  |
| `TRASPASO_VERSION_MAX` | `1` |  |

### Funciones

#### `def clave_col(t)`

Normaliza un nombre de columna para comparar.

Sin tildes, sin espacios ni guiones bajos, en mayusculas. Y ademas trata
"ANIO" como "ANO", que es lo que hace falta de verdad: la Ñ no se resuelve
quitando tildes. Descomponer "Año" da "ANO" y "Anio" da "ANIO", y sin este
paso no coincidirian, que es justo el caso que falla:
    'Clave Año_Mes'  ==  'Clave_Anio_Mes'  ==  'clave_anio_mes'

#### `def resolver_columna(columnas, objetivo)`

El nombre REAL de la columna, o None. Primero exacto, despues normalizado.

**— CONFIG COMPARTIDO —**

#### `def get_usuario()`

#### `def leer_config()`

#### `def escribir_json(ruta, data)`

#### `def guardar_config(data)`

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def leer_traspaso(argv)`

Devuelve el dict del traspaso, o None si no vino o no es valido.
Nunca lanza: si el JSON esta roto se cae al modo manual.

**— UTILIDADES —**

#### `def normalizar(t)`

#### `def fmt_tiempo(seg)`

#### `def fmt_n(n)`

12345 -> '12.345'  (separador de miles chileno).

#### `def buscar_parquet(carpeta_reliq)`

El Retiros_h.parquet dentro de '04 Planilla 9'.

#### `def aplicar_cambio_hora(df, col_hora, hora_cambio, log)`

Empuja una hora hacia arriba todo lo que este DESDE hora_cambio.

En el mes del cambio de primavera la hora `hora_cambio` no existe. Si el
archivo viene corrido (1..719), lo que esta guardado como 145 es en realidad
la 146, lo que esta como 146 es la 147, y asi. Sumando 1 desde ahi queda
1..144 y 146..720, con la 145 ausente, que es lo correcto.

Devuelve (df, aplicado, motivo).

COMO SE DETECTA SI YA VIENE APLICADO: mirando si la hora del cambio EXISTE
en los datos.
  - si existe  -> todavia no se aplico, hay que empujar
  - si no esta -> ya se aplico (o no hay retiros a esa hora), no se toca
Es la comprobacion correcta porque el desplazamiento es justamente lo que
deja ese hueco: no se puede aplicar dos veces sin que se note.

#### `def ejecutar(ruta_parquet, base_datos, hora_cambio, log, progreso)`

Carga el parquet en la tabla, vaciandola antes.

hora_cambio: None, o la hora mensual del cambio de horario de primavera.
Si viene, se empuja una hora todo lo que este desde ahi (ver
aplicar_cambio_hora). El parquet NO se toca.

Devuelve (ok, resumen).

**— VENTANA —**

#### `def main()`


---

## `Revisor Reliquidación/actualizadores/Prorratear.py`

> Prorratear: del Access a SQL Server
> Automatiza lo que hoy se hace a mano en SQL Server Management Studio:
>
>   1. Borrar de la base de sobrecostos las tablas:
>        Central_Empresa, Pago_Retiro_reporte_tabla, Sobrecostos, TIPOS
>   2. Importarlas del Access (el "Tasks -> Import Data"):
>        Central_Empresa_Actualizada  ->  Central_Empresa
>          (el .mdb de la planilla 9 no la tiene: ahi la tabla ya se llama
>           Central_Empresa y se copia tal cual)
>        Sobrecostos                  ->  Sobrecostos
>
> *(el encabezado sigue arriba de todo en el archivo)*

**Importa:** `datetime`, `json`, `os`, `pathlib`, `queue`, `re`, `socket`, `subprocess`, `sys`, `threading`, `time`, `tkinter`, `traceback`, `unicodedata`

### Constantes

| Nombre | Valor | |
|---|---|---|
| `DIR_SCRIPT` | `Path(__file__).resolve().parent` |  |
| `CONFIG_PATH` | `DIR_SCRIPT.parent / 'config.json'` | config.json es compartido con el Revisor y el resto de los actualizadores, que viven un nivel arriba (en scripts/, junto al Revisor).  |
| **— Configuracion —** | | |
| `SERVER` | `'SRV-DTE'` |  |
| `DRIVER_SQL` | `'ODBC Driver 17 for SQL Server'` |  |
| `ESCENARIOS` | `dict de 2 claves: 'normal', 'reliq', …` | Los dos escenarios.  |
| `ORDEN_ESC` | `['normal', 'reliq']` |  |
| `TABLAS_A_BORRAR` | `['Central_Empresa', 'Pago_Retiro_reporte_tabla', 'Sobrecostos', 'TIPOS']` | Tablas que se borran antes de importar, en la base de sobrecostos. |
| `IMPORTAR` | `lista de 3 elementos: {'destino': 'Central_Empresa', 'origen': ['Central_Empresa_Actualizada', 'Central_Empresa']}, {'destino': 'Sobrecostos', 'origen': ['Sobrecostos']}, {'destino': 'TIPOS', 'origen': ['TIPOS']}, …` | Que se copia del Access.  |
| `COL_CLAVE_ACCESS` | `'Clave Año_Mes'` | Columnas del periodo, para comprobar que los retiros y los sobrecostos sean del MISMO mes antes de prorratear. |
| `COL_CLAVE_RETIROS` | `'Clave_Anio_Mes'` |  |
| `TABLA_RETIROS` | `'Retiros'` |  |
| `VISTA_PAGO` | `'10_Pago_retiros'` |  |
| `TABLA_REPORTE` | `'Pago_Retiro_reporte_tabla'` |  |
| `SQL_REPORTE` | `'\nSELECT Tipo_sobrecosto, Concepto, Barra, Suministrador, Retiro, clave, Tipo,\n SUM(pag…` |  |
| `CHUNK` | `20000` |  |
| `LARGO_TEXTO` | `255` | Largo de las columnas de texto al crear las tablas.  |
| **— TRASPASO DESDE EL REVISOR —** | | |
| `TRASPASO_ORIGEN` | `'Revisor_Reliquidacion'` |  |
| `TRASPASO_VERSION_MAX` | `1` |  |

### Funciones

**— CONFIG COMPARTIDO —**

#### `def get_usuario()`

#### `def leer_config()`

#### `def escribir_json(ruta, data)`

#### `def guardar_config(data)`

#### `def abrir_en_explorador(ruta, es_archivo=False)`

#### `def leer_traspaso(argv)`

El dict del traspaso, o None. Nunca lanza: si el JSON esta roto se cae al
modo manual.

**— UTILIDADES —**

#### `def normalizar(t)`

#### `def fmt_tiempo(seg)`

#### `def fmt_n(n)`

#### `def driver_access()`

#### `def tablas_access(ruta_mdb)`

Nombres de las tablas de usuario del .mdb.

#### `def elegir_origen(disponibles, candidatos)`

El primer candidato que exista, comparando sin mayusculas ni tildes.

#### `def objeto_existe(cn, nombre, text, tipos=('U',))`

True si el objeto existe en la base conectada.

OJO con el tipo: en OBJECT_ID(nombre, tipo) la 'U' es TABLA DE USUARIO y la
'V' es VISTA. Preguntando por 'U' una vista devuelve NULL aunque exista,
que es justo lo que pasaba con 10_Pago_retiros.

Se consulta sys.objects, que no obliga a elegir un tipo y ademas permite
preguntar por varios de una.

#### `def tabla_existe(cn, nombre, text)`

Solo tablas: es lo que se puede borrar con DROP TABLE.

#### `def vista_o_tabla_existe(cn, nombre, text)`

Vista o tabla. La 10_Pago_retiros es una VISTA, pero si algun dia fuera
una tabla el paso final funcionaria igual, asi que se aceptan las dos.

#### `def contar(cn, nombre, text)`

#### `def tipos_sql(df)`

(tipos, avisos) para crear la tabla en SQL Server.

Hace falta porque to_sql, sin decirle nada, crea las columnas de texto como
NVARCHAR(MAX). Y NVARCHAR(MAX) **no se puede indexar**: un join sobre esa
columna obliga al servidor a recorrer la tabla entera. El asistente
"Import Data" de Management Studio crea longitudes concretas, y por eso las
tablas que hizo el asistente andan mas rapido que las que haria to_sql a
secas.

El texto va en NVARCHAR(255), igual que lo que crea el asistente y que lo
que tienen hoy las tablas. Solo se agranda si algun dato no entra.

#### `def periodos_access(ruta_mdb, tabla, columna)`

Los Clave Año_Mes que hay en una tabla del .mdb.

#### `def periodos_retiros(base_retiros, text, create_engine)`

Los Clave_Anio_Mes que hay hoy en la tabla Retiros de esa base.

#### `def ejecutar(ruta_mdb, escenario, solo_mirar, log, progreso)`

Devuelve (ok, resumen).

**— VENTANA —**

#### `def main()`


---

## Constantes definidas en más de un archivo

Cada una es un punto donde un cambio hay que hacerlo en varios lados a la vez. Candidatas a mudarse a `comun/`.

| Constante | Archivos |
|---|---|
| `CENTRALES_EMBALSE` | `Revisor Reliquidación/Revisor_Reliquidacion.py`, `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py` |
| `CHUNK` | `Revisor Reliquidación/actualizadores/Carga_Retiros.py`, `Revisor Reliquidación/actualizadores/Prorratear.py` |
| `CONFIG_PATH` | `Revisor Reliquidación/Reemplazos REUC/ActualizaRemplazos.py`, `Revisor Reliquidación/Revisor_Reliquidacion.py`, `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`, `Revisor Reliquidación/actualizadores/Actualiza_Cuadro0.py`, `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py`, `Revisor Reliquidación/actualizadores/Actualiza_Energia.py`, `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py`, `Revisor Reliquidación/actualizadores/Actualiza_datos.py`, `Revisor Reliquidación/actualizadores/Carga_Retiros.py`, `Revisor Reliquidación/actualizadores/Prorratear.py` |
| `DIR_SCRIPT` | `Revisor Reliquidación/Revisor_Reliquidacion.py`, `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`, `Revisor Reliquidación/actualizadores/Actualiza_Cuadro0.py`, `Revisor Reliquidación/actualizadores/Actualiza_Energia.py`, `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py`, `Revisor Reliquidación/actualizadores/Carga_Retiros.py`, `Revisor Reliquidación/actualizadores/Prorratear.py` |
| `FUENTES` | `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py`, `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py` |
| `LARGO_TEXTO` | `Revisor Reliquidación/actualizadores/Carga_Retiros.py`, `Revisor Reliquidación/actualizadores/Prorratear.py` |
| `NS_REL` | `Revisor Reliquidación/Revisor_Reliquidacion.py`, `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py` |
| `NS_XL` | `Revisor Reliquidación/Revisor_Reliquidacion.py`, `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py` |
| `SERVER` | `Revisor Reliquidación/actualizadores/Carga_Retiros.py`, `Revisor Reliquidación/actualizadores/Prorratear.py` |
| `TRASPASO_ORIGEN` | `Revisor Reliquidación/Reemplazos REUC/ActualizaRemplazos.py`, `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`, `Revisor Reliquidación/actualizadores/Actualiza_Cuadro0.py`, `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py`, `Revisor Reliquidación/actualizadores/Actualiza_Energia.py`, `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py`, `Revisor Reliquidación/actualizadores/Actualiza_datos.py`, `Revisor Reliquidación/actualizadores/Carga_Retiros.py`, `Revisor Reliquidación/actualizadores/Prorratear.py` |
| `TRASPASO_VERSION_MAX` | `Revisor Reliquidación/Reemplazos REUC/ActualizaRemplazos.py`, `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`, `Revisor Reliquidación/actualizadores/Actualiza_Cuadro0.py`, `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py`, `Revisor Reliquidación/actualizadores/Actualiza_Energia.py`, `Revisor Reliquidación/actualizadores/Actualiza_SC_CO.py`, `Revisor Reliquidación/actualizadores/Actualiza_datos.py`, `Revisor Reliquidación/actualizadores/Carga_Retiros.py`, `Revisor Reliquidación/actualizadores/Prorratear.py` |
| `_ENT_XML` | `Revisor Reliquidación/Revisor_Reliquidacion.py`, `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py` |
| `_NECESITA` | `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`, `Revisor Reliquidación/actualizadores/Actualiza_Energia.py` |
| `_RE_ENT` | `Revisor Reliquidación/Revisor_Reliquidacion.py`, `Revisor Reliquidación/actualizadores/Actualiza_Data_Access.py` |
| `_TIENE` | `Revisor Reliquidación/actualizadores/Actualiza_Access_P9.py`, `Revisor Reliquidación/actualizadores/Actualiza_Energia.py` |
