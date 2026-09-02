# Estructura y detalle de archivos — CASO RELIQUIDACIÓN

Reemplaza a `02 CASO RELIQUIDACION.txt`. Además del árbol de carpetas incluye la
ubicación exacta de cada dato (hoja, celda, columna, tabla) que usan el revisor y
los scripts de actualización.

Convenciones: `AAMM` es el año-mes reliquidado (ej. `2407`), `AAAAMMDD` una fecha
completa, y `*` un sufijo variable de revisión (ej. `R01D`, `R01P`).

Lo marcado con **(verificado)** se comprobó contra
`0_CUADROS_RELIQUIDACIÓN_SSCC_2312_R01D_v2.xlsm` y `CPRT_2312_R01D.csv`. Lo que
sigue siendo supuesto está en la sección 10.

---

## 1. Árbol de carpetas

```
<carpeta del mes>/
│
├── FD/                                    ← FUERA de 02 CASO RELIQUIDACION
│   └── SSCC_Desempeno*.xlsx|xlsm             origen de todas las hojas FD
│
└── 02 CASO RELIQUIDACION/
    │
    ├── 00 Entregables/
    │   ├── 01 Sobrecostos/
    │   │   ├── Detalles diarios/
    │   │   │   └── Detalle Sobrecostos AAAAMMDD.xlsx      (uno por día)
    │   │   └── Cálculo_SobrecostosSSCC_AAMM_*.xlsm        (copia)
    │   │
    │   ├── 02 Costo de Oportunidad/
    │   │   ├── Detalle diario/
    │   │   │   └── Detalle Costo de Oportunidad AAAAMMDD.xlsx
    │   │   └── Cálculo_CO_AAMM_*.xlsm
    │   │
    │   ├── 03 Costo de Combustible Adicional/
    │   │   ├── Detalle diario/
    │   │   │   └── Detalle CCA AAAAMMDD.xlsx
    │   │   └── Consolidado_CCA_AAMM_*.xlsm
    │   │
    │   ├── 0_CUADROS_RELIQUIDACIÓN SSCC_AAMM_*.xlsm       « EL QUE SE VA A PAGO »
    │   ├── 1_CUADROS_PAGO_SSCC_AAMM_*.xlsm
    │   ├── 3_REMUNERACIÓN_SUBASTAS_E_ID_AAMM_*.xlsm       (copia)
    │   ├── 4_REMUNERACIÓN_SC_CO_CCA_Y_Pagos_Retiros_AAMM_*.xlsx
    │   ├── 5_REMUNERACIÓN_CRA_AAMM_*.xlsx                 (copia)
    │   ├── 6_REMUNERACIÓN_REA_Y_CO_ERNC_AAMM_*.xlsx       (copia)
    │   └── 9_Pagos_Retiros_CRA_REA_CO_ERNC_Subastas_AAMM_*.xlsx
    │
    ├── 01 Sobrecostos/                                    « MAESTRO »
    │   ├── Detalles diarios/
    │   │   ├── Detalle Sobrecostos AAAAMMDD.xlsx          (maestro de los diarios)
    │   │   └── 02 Consolidado_Tabulado_AAMM_*.xlsm
    │   ├── 03b ENTRADA_SOB_SSCC_AAMM_*.mdb
    │   └── Cálculo_SobrecostosSSCC_AAMM_*.xlsm            (maestro)
    │
    ├── 01.a Sobrecostos de Energia/
    │   ├── Detalles diarios/
    │   │   └── Detalle Sobrecostos AAAAMMDD.xlsx          (copia)
    │   ├── 03b ENTRADA_SOB_AAMM_*.mdb
    │   ├── Consolidado_AAMM_*.xlsm                        (a veces .xlsx)
    │   └── Pago_Sobrecostos_AAMM_*.xlsx
    │
    └── 04 Planilla 9/                     « MAESTRO de 3_ / 5_ / 6_ »
        ├── 3_REMUNERACIÓN_SUBASTAS_E_ID_AAMM_*.xlsm       (maestro)
        ├── 5_REMUNERACIÓN_CRA_AAMM_*.xlsx                 (maestro)
        ├── 6_REMUNERACIÓN_REA_Y_CO_ERNC_AAMM_*.xlsx       (maestro)
        ├── Ocupar_este_para_Reliquidacion_AAMM_*.mdb
        ├── Prorrata_Retiros_AAMM_*.xlsx
        └── Retiros_h.parquet                              (se carga a SQL Server)
```

Junto a los `.py`, fuera del árbol del mes:

```
<carpeta de los scripts>/
├── Revisor_Reliquidacion.py
├── Actualiza_datos.py
├── Actualiza_Data_Access.py
├── Actualiza_Energia.py              ← importa el motor de Actualiza_Data_Access
├── Actualiza_Cuadro0.py
├── Actualiza_SC_CO.py
├── Carga_Retiros.py
├── Prorratear.py
├── Actualiza_Access_P9.py
├── config.json                       ← compartido por estos cuatro
├── Salidas/AAMM/                     ← estado, caché y JSON de traspaso, por mes
└── Reemplazos REUC/
    ├── ActualizaRemplazos.py
    └── Auxiliares/config.json        ← el suyo, aparte del compartido
```

### Notas del árbol

- **`Detalles diarios` vs `Detalle diario`**: el nombre varía. En Sobrecostos y en
  Sobrecostos de Energía es plural; en Costo de Oportunidad y en Costo de
  Combustible Adicional es singular. Aceptar las cuatro variantes.
- **`02 Consolidado_Tabulado`** está *dentro* de `01 Sobrecostos/Detalles diarios`,
  no en la raíz de `01 Sobrecostos`.
- **Copias que se propagan por copiado**, no por script:

  | Maestro | Copias |
  |---|---|
  | `01 Sobrecostos/Detalles diarios/Detalle Sobrecostos AAAAMMDD` | `00 Entregables/01 Sobrecostos/Detalles diarios/`, `01.a Sobrecostos de Energia/Detalles diarios/` |
  | `01 Sobrecostos/Cálculo_SobrecostosSSCC` | `00 Entregables/01 Sobrecostos/` |
  | `04 Planilla 9/3_`, `5_`, `6_` | `00 Entregables/` |

- **Sufijo de revisión**: puede no coincidir entre archivos (visto un `.mdb` en
  `R01P` con todo el resto en `R01D`). No es error del árbol pero suele explicar
  descuadres de monto.
- **Copias de Windows**: al buscar por patrón hay que descartar los nombres que
  terminen en `- copia`, `- copia (2)`, `- Copy` o `(2)`, o una copia más nueva le
  gana al original.
- **Reliquidar es reciclar**: se parte de un cuadro de un mes pasado y se pisa. Los
  archivos llegan dimensionados para el mes anterior, y a veces con valores donde
  debería haber fórmulas. Eso explica casi todas las verificaciones de largo y todo
  el "limpiar antes de escribir" de la sección 7.

---

## 2. Bases Access

Las tres tienen la **misma estructura**:

| Archivo | Ubicación |
|---|---|
| `03b ENTRADA_SOB_SSCC_AAMM_*.mdb` | `01 Sobrecostos` |
| `03b ENTRADA_SOB_AAMM_*.mdb` | `01.a Sobrecostos de Energia` |
| `Ocupar_este_para_Reliquidacion_AAMM_*.mdb` | `04 Planilla 9` |

Tabla **`Sobrecostos`**:

| Columna | Contenido |
|---|---|
| `Clave Año_Mes` | AAMM |
| `Tipo_sobrecosto` | SCMT, SCPC, SCAGC, … |
| `Central` | nombre de la central |
| `Hora Mensual` | hora del mes |
| `Sobrecosto` | monto |

El total de una base es `SELECT SUM([Sobrecosto]) FROM [Sobrecostos]`.

Otras tablas presentes en el `.mdb` de SSCC, no usadas por el revisor:
`Central_Empresa` (Central, Empresa), `Central_Empresa Feb26` (Central, Empresa,
F3), `Central_Empresa_Actualizada` (Central, Empresa), `TIPOS` (Tipo_Pago,
Concepto), `Verificacion propietarios actualizado` (Central,
Central_Empresa_Empresa, `Central_Empresa Feb26_Empresa`).

Requiere el driver `Microsoft Access Driver (*.mdb, *.accdb)` de la **misma
arquitectura** (32/64 bits) que el Python que ejecuta.

---

## 3. Dónde está cada dato en las planillas

### Totales que son una columna completa bajo un encabezado

| Archivo | Hoja | Columna | Encabezado | Datos desde | Filtro |
|---|---|---|---|---|---|
| `Cálculo_SobrecostosSSCC` (01 Sobrecostos) | `SOBRECOSTOS TOTAL` | F | F4 | F5 | — |
| `Cálculo_CO` | `CO TOTAL` | F | F4 | F5 | — |
| `Consolidado_CCA` | `CCA` | BC | BC2 | BC3 | — |
| `02 Consolidado_Tabulado` | `Sobrecostos` | AE | AE2 | AE3 | col **AB** = tipo |
| `Consolidado_AAMM` | `Sobrecostos` | E | E1 | E2 | — |

### Columnas conocidas del `Consolidado_AAMM`, hoja `Sobrecostos`

Se pega del tabulado con un **corrimiento**, así que no coinciden las letras:

| Origen (tabulado, desde fila 3) | Destino (Consolidado, desde fila 2) | Qué es |
|---|---|---|
| `A`…`G` | `A`…`G` | `G` = **Generación** |
| `I` | `H` | **CV** |
| `J` | `I` | **CMg** |
| `H` | `J` | (queda al final) |
| `E` | `E` | el monto que suma V6 |

El destino **no tiene columna de tipo**, porque la `AB` del origen no se copia: el
Consolidado lleva solo SCMT y SCPC.

En `02 Consolidado_Tabulado` la columna **AB** trae el `Tipo sobrecosto` (AB2 es el
encabezado). De ahí salen los totales por tipo **SCMT**, **SCPC** y **SCAGC**.

### Columnas del `02 Consolidado_Tabulado`, hoja `Sobrecostos`

Encabezado en la **fila 2**, datos desde la **fila 3**.

| Columna | Contenido |
|---|---|
| `E` | **Sobrecosto**, ya calculado |
| `G` | **Generación** |
| `I` | **CV** (costo variable) |
| `J` | **CMg** (costo marginal) |
| `W` | **USD** (tipo de cambio) |
| `AB` | Tipo sobrecosto (SCMT, SCPC, SCAGC, …) |
| `AE` | Sobrecosto (el que suman V5, V6 y `Actualiza_Energia`) |
| `AA`, `AC`, `AD` | Clave Año_Mes, Central, Hora Mensual |

El sobrecosto sale de sus componentes:

```
Sobrecosto = (CV − CMg) × Generación × USD
```

o sea `E = (I − J) × G × W`, fila por fila. Eso es lo que comprueba **V17**.

### Celdas puntuales

| Archivo | Hoja | Celda | Qué es |
|---|---|---|---|
| `Cálculo_SobrecostosSSCC` | `SOBRECOSTOS` | EE6 | debe igualar el total SCAGC |
| `Cálculo_SobrecostosSSCC` | `SOBRECOSTOS TOTAL` | H1 | debe ser 0 |
| `Pago_Sobrecostos` | `VERIFICADORES` | L5 | total, vs suma del `ENTRADA_SOB.mdb` |
| `Pago_Sobrecostos` | `VERIFICADORES` | M5 | debe ser 0 |
| `3_REMUNERACIÓN_SUBASTAS` | `RESUMEN` | D7, E7, F7, L7, M7, N7, U7, V7, W7 | U7 = total Subastas |
| `3_REMUNERACIÓN_SUBASTAS` | `RESUMEN_DESGLOSADO` | D7, E7, F7, K7, L7, M7 | |
| `5_REMUNERACIÓN_CRA` | `RESUMEN` | D7, E7, F7 | D7 = total CRA |
| `6_REMUNERACIÓN_REA_Y_CO_ERNC` | `RESUMEN` | E7, F7, G7, L7, M7, N7 | E7 = total REA, L7 = total CO-ERNC |
| `4_REMUNERACIÓN_SC_CO_CCA` | `VERIFICADORES` | K6, L6 | K6 = −L6; L6 cuadra con el `.mdb` SSCC |
| `1_CUADROS_PAGO_SSCC` | `RESUMEN` | C10, C11, C12, C13, C14, C15, C17:C25 | |
| `0_CUADROS_RELIQUIDACIÓN` | `CPRT` | I3 | descuadre del cuadro de pago **(verificado)** |

### Tabla de largo variable

`9_Pagos_Retiros`, hoja `VERIFICADORES`: la tabla **cambia de cantidad de filas
entre meses**, así que el total **no** está en una celda fija. Hay que buscar la
fila rotulada **`Total general`**:

- rótulo en columna **A** → el total está en la columna **B** de esa fila
- rótulo en columna **D** → el total está en la columna **E** de esa fila

Comparar el rótulo normalizado (sin tildes, sin mayúsculas, sin espacios
sobrantes). Si aparece más de una vez, usar la última.

### `1_CUADROS_PAGO_SSCC`, hoja `01.SSCC_Recurso_Técnico`

| Rango | Contenido |
|---|---|
| `I9` hacia abajo | empresa |
| `J9` hacia abajo | lo que **paga** |
| `K9` hacia abajo | lo que **recibe** |

Esta tabla `I9:K` se **copia y pega** en `A5:C` de la hoja #3 del cuadro cero. Al
ser copiar y pegar tienen que ser idénticas, pero **pueden quedar en orden
distinto**, así que hay que comparar indexando por empresa.

Ninguna empresa se llama `""` ni `0`: si aparece uno de esos, se descarta. Y no
puede haber empresas repetidas — una repetida significa que a alguien se le paga o
se le cobra dos veces.

---

## 4. El cuadro cero — `0_CUADROS_RELIQUIDACIÓN SSCC` **(todo verificado)**

Es el archivo que **se va a pago**. Si sale mal hay que refacturarle a todas las
empresas, así que es el más detallado.

### 4.1 Hojas, por posición

La posición importa porque el revisor pide la hoja como `#3`, y **la cuenta incluye
las hojas ocultas**:

| # | Hoja | |
|---|---|---|
| 1 | `RESULTADO` | oculta |
| 2 | `Paso` | oculta |
| **3** | **`Dic23`** — el nombre cambia cada mes (`Ene24`, …) | ← la primera visible |
| 4 | `01.SSCC_Recurso_Técnico` | |
| 5 | `EMPRESAS` | |
| 6 | `Data` | oculta |
| 7 | `CPRT` | |
| 8 | `Retenciones Empresas FMCP` | |
| 9 | `DATA_CP` | oculta |
| 10 | `DATA_CP2` | oculta |

> La hoja #3 se pide por posición y no por nombre porque el nombre cambia todos los
> meses. Y `01.SSCC_Recurso_Técnico` se llama **igual** que una hoja del
> `1_CUADROS_PAGO`: son de archivos distintos, ojo al leer.

### 4.2 Hoja #3 — el cálculo por empresa

| Rango | Qué es |
|---|---|
| `A5` hacia abajo | empresa (pegada del `1_CUADROS`) |
| `B5` hacia abajo | lo que **paga** esa empresa |
| `C5` hacia abajo | lo que **recibe** esa empresa |
| `D5` hacia abajo | `=B5+C5`, el **neto** por empresa. Es fórmula, y su largo lo manda la tabla `A:C` |
| `F5` hacia abajo | el otro listado de empresas |
| `K5` hacia abajo | lista consolidada de empresas, **fórmula desbordada** |
| `O1` | rótulo "Tasa interés periodo" |
| `O2` | **la tasa del período** (valor duro, ej. `0,168`) |

Fórmula de `K5`, la que se pega a mano cada mes:

```
=UNICOS(APILARV(F5:F175;A5:A167))
```

Los rangos son **fijos**, así que hay que verificar que alcancen para todas las
filas de A y de F. Y como abarcan celdas vacías, `UNICOS` devuelve un `0`: ése es
el cero que hay que descartar en todas las listas de empresas. La versión que evita
las dos cosas:

```
=LET(x;UNICOS(APILARV(F5:F1000;A5:A1000));FILTRAR(x;(x<>0)*(x<>"")))
```

Antes de escribirla hay que **limpiar lo que haya debajo**, o tira
`#¡DESBORDAMIENTO!` contra los restos del mes anterior.

Fórmulas de la fila 5, que se extienden hasta la última empresa de `K`:

| Celda | Fórmula | Qué es |
|---|---|---|
| `L5` | `=SUMAR.SI(A:A;K5;D:D)` | lo que paga |
| `M5` | `=SUMAR.SI(F:F;K5;I:I)` | lo que recibe |
| `N5` | `=L5-M5` | neto |
| `O5` | `=N5*$O$2` | **el interés — acá entra la tasa** |
| `P5` | `=N5+O5` | neto con interés |
| `Q` | — | **VACÍA a propósito**: separa las dos tablas, no se verifica |
| `R5` | `=+K5` | la empresa otra vez |
| `S5` | `=SI(P5>0;P5;0)` | positivo |
| `T5` | `=SI(P5<0;-P5;0)` | negativo |
| `U5` | `=S5-T5` | el neto que se lleva a `01.SSCC` |

> **La tasa multiplica cada fila.** Si `O2` está mal, **todos** los montos están
> mal, y ninguna verificación posterior lo caza: el descuadre de `CPRT!I3` sigue
> saliendo chico porque la matriz reparte proporcionalmente. Si la tasa llega al
> final del mes, es mejor dejar la del Excel que poner una a medias.

`L:P` y `R:U` son 9 columnas (sin la `Q`) y tienen que llegar **justo** hasta la
última empresa de `K`: si faltan, hay empresas sin calcular; si sobran, arrastran
`0` y ensucian los totales.

### 4.3 Hoja `01.SSCC_Recurso_Técnico` del cuadro cero

Fórmulas de la fila 9:

| Celda | Fórmula | Qué es |
|---|---|---|
| `A9` | `=+'<hoja #3>'!R5` | empresa (desfase de **4 filas**: K5 ↔ A9) |
| `B9` | `=+$B$1` | |
| `C9` | `=SI.ND(BUSCARV(A9;EMPRESAS!$H:$I;2;0);A9)` | **empresa DESPUÉS de los reemplazos** |
| `D9` | `=+'<hoja #3>'!S5` | |
| `E9` | `=+'<hoja #3>'!T5` | |
| `F9` | `=+'<hoja #3>'!U5` | |
| `G9` | `=+F9` | **total neto de la fila**: puede ser −, 0 o + |
| `I9` | lista única de `C` | ver abajo |
| `J9` | `=SI(SUMAR.SI($C$9:$C$4345;I9;$G$9:$G$4345)<-0,00001; SUMAR.SI(…); "")` | lo que **paga** |
| `K9` | `=SI(SUMAR.SI($C$9:$C$4345;I9;$G$9:$G$4345)>0,00001; SUMAR.SI(…); "")` | lo que **recibe** |

Reglas de largo:

- `A9:G` tiene **tantas filas como empresas haya en `K5:K` de la hoja #3**.
- `I9:K` es **después de los reemplazos**, así que su largo es **menor o igual** que
  el de `A9:G`. En `C9:C` las empresas vienen **repetidas** (una fila por
  reemplazo).
- Al final puede haber `0` o vacíos: se descartan, no son error.

**La `I9` de hoy es la fórmula vieja de únicos**, rellenada celda por celda:

```
=SI.ERROR(INDICE($C$9:$C$4341; COINCIDIR(0; INDICE(CONTAR.SI($I$8:I8;$C$9:$C$4341);0;0);0)); "")
```

Devuelve `""` cuando se le acaban las empresas — de ahí salen las cadenas vacías de
la columna. Y es **cuadrática**: cada celda hace un `CONTAR.SI` sobre 4.341 filas,
o sea millones de comparaciones en cada recálculo. Conviene cambiarla por:

```
=LET(x;UNICOS(C9:C4345);FILTRAR(x;(x<>0)*(x<>"")))
```

> `J9`/`K9` salen del **mismo** `SUMAR.SI` que suma `G` por empresa. O sea que
> comparar "suma de G por empresa" contra "J+K" es casi tautológico: sirve para
> cazar fórmulas que no llegan o un rango corto, **no** para validar los montos de
> forma independiente.

### 4.4 La matriz del cuadro de pagos

La arma la macro `CuadroPago`, en la misma hoja `01.SSCC_Recurso_Técnico`:

| Rango | Qué es |
|---|---|
| `N5` | el título (queda **aislada**: su `CurrentRegion` es 1×1, no sirve para medir) |
| `N6`, `N7` | vacías |
| `N8` | el rótulo `Pagan` |
| `O8` hacia la derecha | las empresas que **reciben** |
| `N9` hacia abajo | las empresas que **pagan** |
| interior | el reparto: `total_del_que_paga × total_del_que_recibe / gran_total` |

**Cambia de tamaño todos los meses.** En 2312: 81 pagan × 6 reciben → `N8:T89`. En
2401: 76 × 8 → `N8:V84`.

Al final de cada lado hay una fila y una columna de **totales**, que no son pares de
pago y hay que descartar al leerla.

### 4.5 Hoja `CPRT` — el cuadro de pago

| Rango | Qué es |
|---|---|
| filas 1–5 | cabecera: `Coordinador:`, `Concepto:`, `Tipo:`, `Fecha Generación:`, `Fecha Emisión:` |
| fila 6 | encabezado de columnas |
| `A7:G` hacia abajo | los datos — **salida de la tabla dinámica** (`TablaDinámica1` ocupa `A6:G487`) |
| `A` | la palabra `Fila` — sirve para cortar: debajo hay fórmulas que arrastran 0 |
| `B` / `C` | nemotécnico y RUT del **deudor** |
| `D` | la palabra `Columna` |
| `E` / `F` | nemotécnico y RUT del **acreedor** |
| `G` | **Monto** |
| `H` | **Monto retenido** = `=SI(CONTAR.SI('Retenciones Empresas FMCP'!I:I;E7)<>0;0;1)*G7` |
| `M` | `=CONCAT(A7;",";B7;…;H7;",")` — lo que exporta la macro `CPRT_csv` |
| `I1` | `=SUMA(G:G)` — total del cuadro de pago |
| `I2` | `=SUMA('01.SSCC_Recurso_Técnico'!K:K)` — total de origen |
| `I3` | `=I2-I1` — **el descuadre** |

`H` es `G` puesto en **0** cuando el acreedor está en la lista de retenciones; o
sea, `H` es el monto *neto de retenciones*.

`I3` nunca da 0 exacto porque el reparto proporcional redondea; lo que importa es
que sea **chico**. En 2312 dio **31,30**. El revisor lo compara contra
`UMBRAL_DESCUADRE_CPRT` (500 por omisión).

### 4.6 Nombres definidos

Un nombre definido es un apodo para un rango (Fórmulas → Administrador de nombres).
El que importa:

| Nombre | Valor en 2312 | |
|---|---|---|
| `CPTEE` | `='01.SSCC_Recurso_Técnico'!$N$8:$T$89` | **la matriz**; de acá lee la Power Query |
| `CPBP` | `#REF!` | roto |
| `CPITE` | `='[1]02.IT ENERGIA Oct-22 R01D'!$B$10:$CX$238` | apunta al vínculo externo roto |

`CPTEE` hay que reapuntarlo cada mes porque la matriz cambia de tamaño. Eso es lo
que hace el botón **Actualiza Rango**, con una sola línea:

```vba
ThisWorkbook.Names("CPTEE").RefersTo = "='01.SSCC_Recurso_Técnico'!$N$8:" & ran
```

Puede quedar guardado en A1 (`$N$8:$T$89`) o en R1C1 localizado
(`F8C14:F84C22` = Fila 8 Columna 14 hasta Fila 84 Columna 22). Son la misma cosa.

### 4.7 Macros del libro

| Macro | Qué hace |
|---|---|
| `CuadroPago` | arma la matriz. **Valida signos**: aborta si un monto de "Pagan" o "Reciben" tiene el signo al revés. Empieza con `Celda.Offset(1,0).CurrentRegion.Clear`, o sea **borra antes de escribir** |
| `Actualiza_rango_1` | reapunta `CPTEE`. Termina con un `MsgBox "Rangos Actualizados"` |
| `Actualiza_rango_2` | apunta a una hoja `02.SSCC_Infraestruct` que **no existe** en este libro, y su nombre `CPBP` está en `#REF!`. Es basura |
| `CPRT_csv` | exporta la columna `M` a csv |

> **`CuadroPago` atrapa su propio error.** Si los signos no cuadran, muestra el
> `MsgBox` y sale con `Exit Sub`, **sin propagar nada**. Llamándola desde código
> "termina bien" aunque no haya escrito la matriz, y el descuadre de `I3` tampoco lo
> delata, porque la matriz vieja sigue intacta. Por eso conviene revisar los signos
> de `J:K` antes: en la tabla, todo número de `J` tiene que ser **negativo** y todo
> número de `K` **positivo**.

### 4.8 El csv del CPRT

`CPRT_AAMM_*.csv`, junto al cuadro cero. Formato **verificado byte a byte**:

| | |
|---|---|
| Codificación | **cp1252** (latin-1), NO utf-8 |
| Fin de línea | **CRLF** |
| Separador | coma, comillas solo si hacen falta |
| Campos | **7**: A:F más **la `H`** (Monto retenido), no la `G` |
| Rótulo del monto | sigue diciendo **"Monto"**, tomado de `G6` (ver abajo) |
| Líneas | 5 de cabecera + 1 de encabezado + N de datos |
| Ancho | todas con 7 campos (las de cabecera se rellenan con comas vacías) |
| Corte | por la columna `A`: si no dice `Fila`, la fila no va |
| Montos | enteros, sin separador de miles |

Un detalle que hay que copiar **tal cual y sin arreglar**: la cabecera trae
`Fecha Generaci„n:` con el carácter roto. Está así **en la celda del Excel**, y el
archivo original lo escribió igual. Si se "corrige" a `Generación`, el archivo deja
de ser idéntico al que espera el destinatario. Si algún día quieren que salga bien
escrito, se arregla en la celda, no en el script.

El csv paga el monto **neto de retenciones**, igual que la macro `CPRT_csv` del libro
(su columna `M` concatena hasta la `H`). Pero el **rótulo** de la columna se sigue
tomando de `G6`, así que el encabezado dice `Monto` y no `Monto retenido`: el único
csv que sabemos que el destinatario aceptó dice `Monto`, y si el sistema que lo lee
mira el nombre de la columna, cambiarlo lo rompe. La constante es
`CPRT_ROTULO_DESDE_G`; con `CPRT_COL_MONTO = 7` se vuelve a la `G`.

En 2312 el archivo sale **byte a byte idéntico** al de referencia, porque `G = H` en
las 481 filas. El mes que haya retenciones de verdad, el csv va a pagar menos: por eso
el exportador avisa cuántas filas van en 0 y cuánto se retiene en total, y la
comprobación 9 de V16 lo repite al verificar.

> **La `H` solo puede ser igual a la `G` o exactamente 0.** Cualquier otro valor
> significa que la fórmula de la `H` está mal, y eso sí es un error, porque el csv se
> arma con la `H`. Que haya filas en 0 **no** es error: es una retención real, y se
> informa con el monto. Y si a alguna fila le falta el valor de `H`, el csv saldría
> con el monto vacío, así que también falla.

### 4.9 Reconstruir el CPRT sin la dinámica — lo que falta

Desarmando la matriz `CPTEE` en pares pagador–receptor y filtrando `|monto| ≥ 10`
salen **exactamente las 481 filas** del csv de 2312. Pero **los RUT no calzan**: 60
de 481 difieren de los de la hoja `EMPRESAS` (ej. COCHRANE: en `EMPRESAS`
`76.085.254-6`, en el csv `77.285.492-7`).

Los RUT del CPRT **no salen de `EMPRESAS`**: salen de la conexión
`WorksheetConnection_Cuadros de Pago_Balances_SEN_Ago18_def.xlsm!RUT`, o sea de un
rango en **otro libro**. Falta identificar la fuente correcta antes de poder generar
el csv sin la dinámica.

El corte por monto tampoco está documentado: en 2312 los excluidos llegaban a
**8,74** y el menor incluido era **22,33**, así que el umbral real está entre esos
dos.

---

## 5. El vínculo externo roto

El cuadro cero tiene un **vínculo a un archivo de 2022**:

```
../../../../../../../2022/10 Octubre/03 Reliquidación/02 2508 (R01D)/
   Entregables/01 Resultados/Cuadros de Pago_Balances_SEN_Oct22_R01D.xlsx
```

La ruta es **relativa y sube siete niveles**, así que depende de dónde esté parado el
archivo. Movido de carpeta deja de resolver.

Consecuencias observadas:

- `Consulta - TEE` (una **Power Query**, `type=100`) falla con
  `[Expression.Error] No se pudieron actualizar los datos`, **también refrescando a
  mano en Excel**. No es un problema de la automatización.
- `CPBP` está en `#REF!` y `CPITE` apunta a `[1]`, que es ese vínculo.
- Sin refrescar la dinámica, `CPRT!I3` muestra el residuo del **refresco anterior**:
  parece información y no lo es.

Salidas posibles: reponer el archivo externo en la ruta que espera (arregla hoy, se
rompe el próximo mes que se mueva algo), o traer la tabla de RUT adentro del libro y
generar el csv desde la matriz (ver 4.9).

---

## 6. Verificaciones del revisor

### Por fecha de modificación (botón ACTUALIZAR)

Las copias deben tener la misma fecha y hora que su maestro; si no, se pintan de
amarillo. Ver la tabla de maestros y copias de la sección 1.

### De valores (botón Verificar de cada archivo)

| # | Archivo | Qué comprueba | Previas |
|---|---|---|---|
| V4 | `03b ENTRADA_SOB_SSCC.mdb` | suma = SSCC + CO + CCA | V8 |
| V5 | `03b ENTRADA_SOB.mdb` | suma = SCMT + SCPC | V17 |
| V6 | `Consolidado_AAMM` | total = SCMT + SCPC; y las 3 columnas de control (ver abajo) | V17 |
| V7 | `Pago_Sobrecostos` | L5 = suma `ENTRADA_SOB.mdb`; M5 = 0 | V5 |
| V8 | `Cálculo_SobrecostosSSCC` | EE6 = total SCAGC; H1 = 0; sin `#N/D` ni "REVISAR" | V17 |
| V9 | `3_REMUNERACIÓN_SUBASTAS` (+ la prorrata) | D7=E7, L7=M7, U7=V7, U7=D7+L7, V7=E7+M7, F7=N7=W7=0; desglosado D7=E7, K7=L7, F7=M7=0 | — |
| V10 | `5_REMUNERACIÓN_CRA` | D7 = E7; F7 = 0; y las 5 de la hoja "SC y CO" | — |
| V11 | `6_REMUNERACIÓN_REA` (+ la prorrata) | E7=F7, L7=M7, G7=N7=0 | — |
| V12 | `Ocupar_este.mdb` | suma = U7 + D7 + E7 + L7; y que ninguna central con monto quede sin dueño | V9, V10, V11 |
| V13 | `9_Pagos_Retiros` | total general A/B = total general D/E = suma `Ocupar_este.mdb` | V12 |
| V14 | `4_REMUNERACIÓN_SC_CO_CCA` | K6 = −L6; L6 = suma `ENTRADA_SOB_SSCC.mdb` (en valor absoluto) | V4 |
| V15 | `1_CUADROS_PAGO` | los totales contra las planillas 4 y 9, y el pago por empresa y concepto | V14, V13 |
| V16 | `0_CUADROS_RELIQUIDACIÓN` | las 8 comprobaciones de abajo | V15 |
| V17 | `02 Consolidado_Tabulado` | `(I − J) × G × W` = `E`, fila por fila | — |

Cadena de dependencias:

```
V16 → V15 → V14 → V4 → V8 → V17
          → V13 → V12 → {V9, V10, V11}
V7  → V5  → V17
V6  → V17
```

**V17 es la raíz de casi todo**, porque el `02 Consolidado_Tabulado` es el origen: de
ahí sale la hoja `SOBRECOSTOS` del `Cálculo_SobrecostosSSCC` (V8), y contra ahí se
comparan V5 y V6. Si el tabulado tiene el sobrecosto mal calculado, todas las
comparaciones de totales cuadran igual y el error se propaga hasta el cuadro de pago.
Por eso verificar el cuadro cero ahora pide V17 primero, por el camino
`V16 → V15 → V14 → V4 → V8 → V17`.

### V16 — las 8 comprobaciones del cuadro cero

| # | Qué comprueba | Qué caza si falla |
|---|---|---|
| 1 | tabla `I9:K` del `1_CUADROS` = tabla `A5:C` de la hoja #3, por empresa, sin repetidas | mal pegado, o una empresa dos veces |
| 2 | `K5:K` contiene todas las de `A5:A` y `F5:F`, sin repetir | los rangos fijos de la fórmula de `K5` no alcanzaron |
| 3 | las fórmulas de `L:P` y `R:U` llegan justo hasta la última empresa de `K` | restos del mes anterior, o empresas sin calcular |
| 3.b | la fórmula de `D` llega justo hasta el final de la tabla `A:C` | idem, pero contra la tabla pegada |
| 4 | las empresas de `K5:K` (#3) son **las mismas** que las de `A9:A` (01.SSCC) | ni más ni menos; los 0 y vacíos del final se descartan |
| 5 | empresas de `C9:C` = empresas de `I9:I`, y suma de `G9:G` por empresa = `J+K` | fórmulas que no llegan, rango `$C$9:$C$4345` corto |
| 6 | `CPRT!I3` ≤ `UMBRAL_DESCUADRE_CPRT` (500) | descuadre grande del cuadro de pago |
| 7 | la matriz está armada con las empresas de ahora, y `CPTEE` la cubre justo | **falta correr `Cuadro de pagos`**, o **falta `Actualiza Rango`** |
| 8 | las fórmulas de la `H` del CPRT llegan a todas las filas | el csv saldría con el monto vacío |
| 9 | la `H` es igual a la `G` o es exactamente 0 | la fórmula de la `H` está mal |
| 10 | los pares del `CPRT` calzan con la matriz, y los montos coinciden redondeados | **falta refrescar la tabla dinámica** |

En la 8 la exigencia es asimétrica a propósito: **todo** par del CPRT tiene que
existir en la matriz, pero al revés sólo se exigen los pares de `UMBRAL_PAR_SEGURO`
(100) o más, porque el CPRT descarta los montos chicos y el corte exacto no está
documentado (ver 4.9).

### V9, V10 y V11 — la prorrata está al día

Las tres planillas llevan la prorrata pegada en `PRORRATA_RETIROS`, y **nada más la
delataba**: los totales del `RESUMEN` cuadran igual con la prorrata del mes pasado,
porque son otra cuenta.

| Lado | Archivo | Hoja | Layout |
|---|---|---|---|
| origen | `Prorrata_Retiros` | `PRORRATA_HORARIA_TABULAR`, desde la fila 2 | `A` Hora, `B` Suministrador, `C` valor |
| destino | la planilla 3, 5 o 6 | `PRORRATA_RETIROS` | `B8` = "Hora", `C8`… suministradores, `B9`… las horas |

El destino es el **pivote** del origen: los mismos números en otra forma. Así que se
compara la **suma por suministrador**, que tiene que dar exactamente igual
(`TOL_PRORRATA_SUMA`, 0,0001).

**La exigencia es asimétrica**: todo suministrador del **origen** tiene que estar en
la planilla con la misma suma. Lo que sobra en la planilla sólo se informa, porque hay
dos motivos legítimos:

| Qué sobra | Dónde |
|---|---|
| centrales pegadas a mano que quedan en **0%** | planilla 3 |
| columnas de totales (`Total CO Horario [$]`, `Total REA Horario [$]`) | planillas 5 y 6 |
| un **segundo bloque** con los mismos suministradores pero con **montos** | planillas 5 y 6 |

> Ese segundo bloque obliga a tomar **sólo la primera columna de cada nombre**. El
> bloque pegado empieza en `B8`, así que es el de más a la izquierda; los montos van a
> la derecha. Sumando los dos daban millones contra prorratas de dos dígitos.

Falla sólo si:

- la suma de algún suministrador del origen no coincide → quedó la prorrata de otro mes
- un suministrador del origen **no está** en la planilla
- la fila 8 está vacía → la prorrata **nunca se pegó**

La cantidad de horas se informa si no coincide, pero no hace fallar: la hoja puede
tener filas más abajo que no son parte de la matriz.

Cuando falla, el log dice qué hacer: correr «Actualizar data» y marcar Prorrata.

Las tres pasan a depender de `a_prorrata`, así que **tocar el `Prorrata_Retiros`
vence las tres** — y con eso también V12, V13 y V15, que las arrastran.

### V10 — la hoja "SC y CO" de la planilla 5

| # | Qué comprueba | Qué caza |
|---|---|---|
| 3 | **SC**: filas y suma de `G` del destino = origen filtrado por embalses | los SC no son los de este mes |
| 4 | **CO**: idem contra `Calculo_CO` | los CO no son los de este mes |
| 5 | cada fila de `I9:X` suma el 100% | prorrata pegada a medias o corrida de columna |
| 6 | en `E9:E` solo hay centrales de la lista de embalses | se colaron centrales que no son embalse |
| 7 | las fórmulas de `Y:AB` y `AD:AF` cubren exactamente las filas que hay | restos del mes anterior, o filas sin calcular |
| 8 | en el **origen SC** (`U7:U`) no quedó fuera ninguna central «-número» | apareció una unidad de embalse nueva |
| 9 | lo mismo en el **origen CO** (`D7:D`) | idem |

Las 3 y 4 son las que dicen si los datos son **los de este mes**: todo el resto solo
comprueba que el destino sea coherente consigo mismo. Van **por bloque separado** a
propósito, y cuando fallan el log dice cuál hay que volver a traer:

```
>> CO: filas y monto de la planilla 5 = origen (embalses)
      destino (CO): 3 fila(s), 3.000,00   [SC y CO col G]
      origen  : 3 fila(s), 10.777,00   [Calculo_CO..., PRORRATA CO col G]
      diferencia de monto: -7.777,00
      >>> HAY QUE VOLVER A TRAER LOS CO <<<
          (no hace falta tocar el otro bloque)
```

La 5 también dice a qué bloque pertenece cada fila mala (por la columna `D`), y si
todas son del mismo, lo señala igual:

```
      fila 15 (CO): suma 0.9375
      las filas malas son: 1 de CO
      >>> HAY QUE VOLVER A TRAER LOS CO <<<
```

La 5 acepta que la fila sume **1 o 100**, porque las dos formas se usan, y también
**0**: una fila que no reparte nada es válida. Las filas en 0 se cuentan y se
muestran en el log, pero no hacen fallar, y no entran en la detección de escala
porque el 0 no dice si la planilla trabaja en 1 o en 100.

Lo que sí falla es que un mismo archivo **mezcle** filas en 1 con filas en 100: una
de las dos está mal. O que una fila sume cualquier otra cosa (por ejemplo 0,9375, que
es una prorrata a la que le falta una columna).

Las 6 y 7 son el chequeo de que la lista no se quedó vieja: el sufijo «-número»
indica unidad, y las unidades son lo que tienen los embalses. Una central así que no
esté en `CENTRALES_EMBALSE` es probablemente una unidad nueva. Las centrales sin
sufijo se descartan sin avisar, porque no son embalses.

### V12 — centrales con monto y sin dueño

Cruza las dos tablas del `.mdb` de la planilla 9: toda central con **monto** en
`Sobrecostos` tiene que tener dueño en `Central_Empresa`.

Falla si una central con monto tiene la empresa **vacía o en 0**, o si **no está** en
`Central_Empresa`. Una central sin dueño pero **sin monto** no falla: en
`CONSUMOS_PROPIOS` las hay a propósito.

Los nombres se comparan normalizados, así que `El Toro-1` y `ELTORO-1` cuentan como
la misma central.

### V15 — el pago por empresa contra las planillas 9 y 4

La planilla 1 y las planillas 9 y 4 son **dos cálculos paralelos del mismo pago**: la
1 lo saca por empresa y concepto directo, las otras lo prorratean por retiros y
después lo agrupan. Tienen que dar lo mismo y hasta ahora nadie lo comparaba.

Va en **V15** (la planilla 1) y no en V13/V14 porque la planilla 1 se arma **después**
de la 9: al verificar la 1 tiene sentido compararla contra lo que ya está hecho. Al
revés, verificar la 9 obligaría a tener lista la 1, que todavía no existe.

V15 ya dependía de `a_4_rem` y `a_9_pagos` y arrastraba V13 y V14 como previas, así
que no hizo falta cambiar la cadena.

| Lado | Archivo | Hoja | Columnas |
|---|---|---|---|
| detalle | `9_Pagos_Retiros` / `4_REMUNERACIÓN` | `PAGO_RETIRO` desde la fila 2 | `A` concepto, `G` suministrador, `H` monto |
| resumen | `1_CUADROS_PAGO_SSCC` | `01.SSCC_Recurso_Técnico` desde la fila 9 | `B` concepto, `C` empresa, `E` **PAGA** |

Se agrupa el detalle por `(concepto, empresa)` y se compara con la planilla 1, con un
margen de **100 pesos** (`TOL_PAGO_EMPRESA`). En 2409: los **664 pares** de la
planilla 9 y los **255** de la planilla 4 cuadran, con la peor diferencia en 58 pesos.

Cuatro cosas que hay que respetar y que no son obvias:

1. **Los nombres de concepto no se escriben igual.** La planilla 1 dice `CO ERNC` con
   espacio y la 9 dice `CO_ERNC` con guion bajo. Por eso `clave_concepto()`, que saca
   espacios y guiones bajos pero **conserva los paréntesis y el signo**, porque
   `CSF(+)` y `CSF(-)` son conceptos distintos.
2. **El monto corresponde a la columna `E` (PAGA), no a la `D` (RECIBE).** Verificado
   en 2409: contra PAGA cuadran los 664 pares; contra RECIBE, sólo 41.
3. **El signo de la columna `E` no es consistente**: es positivo para los conceptos de
   la planilla 9 y **negativo** para `CCA`, `CO` y `SC_SSCC`, los de la planilla 4 —
   en la misma columna de la misma hoja. Como la convención cambia dentro del propio
   archivo, no se puede distinguir convención de error, así que se compara la
   **magnitud** y se informa aparte cuántos pares venían con signo opuesto (en 2409
   son 249 en la planilla 4).
4. **Sólo se comparan los conceptos que están en los dos lados.** La planilla 1 tiene
   además los cinco `... ID` (instrucción directa) y los tres de la planilla 4; los
   que sobran de un lado se listan en el log pero no hacen fallar. Y un par que existe
   en un lado con monto **0** y no en el otro tampoco falla: la planilla 1 lista la
   empresa con 0 y la 9 simplemente no tiene filas para ella.

### V17 — el sobrecosto recalculado desde sus componentes

Es la **única verificación de toda la cadena que no depende de que un archivo se
copie bien**: comprueba que el número sea *correcto*, no que coincida con otra copia
de sí mismo. Todas las demás comparan un total contra otro total.

Recalcula `(I − J) × G × W` y lo compara con `E`, y falla si:

- la **suma** difiere más de la tolerancia, o
- **alguna fila** difiere más de `TOL_SOBRECOSTO_FILA` (1 peso), aunque el total
  cuadre, o
- una fila trae sobrecosto pero le **falta un componente**, o
- una fila tiene los componentes y la `E` **vacía**.

Lo de "aunque el total cuadre" es a propósito. Si una fila se equivoca en +500.000 y
otra en −500.000, la suma da bien y el archivo está mal igual. Verificado con un caso
de prueba: diferencia total de 0,00 y estado `NO CUADRA`, con las dos filas
señaladas.

El log muestra la fila, los dos montos, la diferencia y **los componentes de esa
fila**, así se ve al toque si el problema es el CV, el CMg, la generación o el
cambio.

Una fila con `CV = CMg` (sobrecosto 0) a la que le falte un componente **no** hace
fallar: no hay plata en juego. La tolerancia de 1 peso por fila cubre que la `E`
venga redondeada al peso, que es el caso normal.

### V6 — las 4 comprobaciones del `Consolidado_AAMM`

| # | Qué comprueba | Qué caza |
|---|---|---|
| 1 | suma de `E` = SCMT + SCPC del tabulado (col. `AE`, filtro `AB`) | filas viejas con monto, filtro no aplicado |
| 2 | suma de `G` = suma de `G` del tabulado | |
| 3 | suma de `H` = suma de `I` del tabulado | **el pegado se corrió de columna** |
| 4 | suma de `I` = suma de `J` del tabulado | idem |

El lado del **origen** se filtra por `SCMT` + `SCPC` (columna `AB`); el del
**destino** no, porque no tiene columna de tipo.

Las tres últimas existen porque el total de la `E` solo **no** detecta un pegado
corrido: la `E` puede seguir sumando bien mientras CV y CMg están intercambiadas.
Verificado con un caso de prueba: cruzando `H` e `I` en el destino, la
comprobación 1 sigue en OK y las comprobaciones 3 y 4 fallan.

> Sumar el CMg es raro, porque es un **precio** y no un monto. Como suma de control
> sirve igual: si las dos sumas coinciden, la columna es la misma. No hay que leer
> ese número como si significara algo por sí solo.

### En el chequeo de marcas de V8

Hoja `SOBRECOSTOS`, desde la fila 7:

- columnas **CF:CI** → no debe haber errores de fórmula (`#N/D`, `#REF!`, …) ni el
  texto `REVISAR`
- columnas **CD**, **CE**, **CS** → no debe haber el texto `REVISAR`

Los errores se guardan en el archivo en inglés (`#N/A`) aunque Excel los muestre como
`#N/D`. Y hay que leer **solo el resultado** de la celda, nunca el texto de la
fórmula: muchas fórmulas son `IFERROR(...,"REVISAR")` y contienen la palabra sin que
el resultado esté mal.

### Detalle de lectura que afecta a toda comparación de empresas

Los nombres pueden venir guardados con **entidades XML numéricas**
(`Enel Generaci&#243;n`). Hay que desescaparlas o ese texto y `Enel Generación` no
coinciden al comparar, y el cuadro de pago reporta un descuadre que no existe. Y hay
que hacerlo en **una pasada**: reemplazar `&amp;` antes que `&lt;` convierte
`&amp;lt;` (un literal) en `<`.

---

## 7. Actualizaciones automáticas: origen → destino

### `Actualiza_datos.py`

Origen de todas las hojas FD: **`FD/SSCC_Desempeno*.xlsx|xlsm`**, en la carpeta
hermana de `02 CASO RELIQUIDACION`. Hojas de origen: `CT Diario`, `CPF Horario`,
`CSF Horario`, `CTF Horario`, con datos desde la fila 12.

**A `Cálculo_SobrecostosSSCC` (opción `sc`, "Actualizar FD"):**

| Hoja origen | Cols origen | Hoja destino | Cols destino | Desde fila | Fórmulas a extender |
|---|---|---|---|---|---|
| CT Diario | D:I | `FD_CT` | C:H | 12 | B |
| CPF Horario | B:J | `FD_CPF` | C:K | 12 | B, L:P |
| CSF Horario | B:H | `FD_CSF` | C:I | 12 | A:B, J:M |
| CTF Horario | B:I | `FD_CTF` | C:J | 12 | A:B, K:N |
| CSF Horario | B:D + F | `FD_CSF_Disponibilidad` | B:E | 12 | A, F:H |
| CTF Horario | B:D + G | `FD_CTF_Disponibilidad` | B:E | 12 | A, F:H |

**A `Cálculo_SobrecostosSSCC` (opción `sc`, "Traer Consolidado"):**

origen `02 Consolidado_Tabulado`, hoja `Sobrecostos`, columnas A:G + I:J + Q:W desde
la fila 3, **solo las filas donde la columna C = `C.Frec`** → destino hoja
`SOBRECOSTOS`, columnas A:G + I:J + K:Q desde la fila 7, extendiendo las fórmulas de
**R:EB**.

> Ese rango R:EB es el que contiene las fórmulas de la hoja `SOBRECOSTOS`, y por eso
> las columnas de control CD, CE, CF:CI y CS caen ahí.

**A las planillas 3, 5 y 6 ("Actualizar FD"):**

| Planilla | Hoja origen | Cols origen | Hoja destino | Cols destino | Desde fila | Fórmulas |
|---|---|---|---|---|---|---|
| 3 | CPF Horario | B:J | `CPF_FD` | D:L | 9 | A:B, N:P |
| 3 | CSF Horario | B:H | `CSF_FD` | E:K | 9 | A:D, M:N, P:X |
| 3 | CTF Horario | B:I | `CTF_FD` | E:L | 9 | A:D, N:P, R:Y |
| 5 | CPF Horario | B:J | `FD_CPF` | B:J | 7 | A |
| 5 | CSF Horario | B:H | `FD_CSF` | B:H | 7 | A |
| 5 | CTF Horario | B:I | `FD_CTF` | B:I | 7 | A |
| 6 | CT Diario | B + D:I | `FD` | B:H | 7 | — |

**Prorrata (planillas 3, 5 y 6, opción "Actualizar Prorrata"):**

origen `04 Planilla 9/Prorrata_Retiros_AAMM_*.xlsx`, hoja
`PRORRATA_HORARIA_TABULAR` desde la fila 2, con columna A = Hora, B = Suministrador,
C = Valor. Se transforma de tabla larga a matriz pivote y se pega en la hoja destino
`PRORRATA_RETIROS` a partir de **B8**, borrando antes el rango **B8:EC756**. Los
suministradores quedan ordenados alfabéticamente y los faltantes se llenan con 0.

Requisito: los `.xlsm` destino deben estar **cerrados** antes de correr. Al terminar
el archivo queda guardado y **abierto en Excel** a propósito.

### `Actualiza_Data_Access.py`

Consolida tres Excel en la tabla `Sobrecostos` del `.mdb` de `01 Sobrecostos`
(`03b ENTRADA_SOB_SSCC`). Los Excel se abren **solo en lectura** y nunca se guardan.

| Fuente | Archivo | Hoja | Columnas que lee | Encabezado | Datos desde | Filtra ceros |
|---|---|---|---|---|---|---|
| SSCC | `01 Sobrecostos/Cálculo_SobrecostosSSCC*.xlsm` **(el maestro)** | `SOBRECOSTOS TOTAL` | B:F | fila 4 | 5 | no |
| CO | `00 Entregables/02 Costo de Oportunidad/Cálculo_CO*.xlsm` | `CO TOTAL` | B:C, G, E:F | fila 4 | 5 | sí |
| CCA | `00 Entregables/03 Costo de Combustible Adicional/Consolidado_CCA*.xlsm` | `CCA` | I:L, BC | fila 2 | 3 | sí |

Las columnas quedan en el orden
`Clave Año_Mes, Tipo_sobrecosto, Central, Hora Mensual, Sobrecosto`. La columna de
monto de cada fuente es la última de su bloque: **F** para SSCC, **F** para CO y
**BC** para CCA, que es exactamente lo que el revisor suma para V4.

En el Access se borran solo las filas cuyo `Tipo_sobrecosto` viene en los Excel de
las fuentes seleccionadas y se insertan las nuevas; los demás tipos quedan intactos.

El SSCC se resuelve desde la **carpeta del propio `.mdb`** (que *es*
`01 Sobrecostos`), no desde `00 Entregables`. Antes usaba la copia, que es un archivo
distinto del que edita el usuario y del que suma el revisor en V4.

Declara qué sabe hacer, para que los scripts que lo importan puedan comprobar que
la versión que tienen al lado no es antigua y abortar con un mensaje claro:

```python
CAPACIDADES = {"fuentes_externas", "filtro_por_valores",
               "borrar_todo", "cols_no_cero"}
```

| Capacidad | Qué habilita | Quién la necesita |
|---|---|---|
| `fuentes_externas` | `proceso(..., fuentes=...)` con un dict propio | Energía, P9 |
| `filtro_por_valores` | `leer_fuente` respeta `cfg["filtro"]` | Energía, P9 |
| `borrar_todo` | `proceso(..., borrar_todo=True)` vacía la tabla | P9 |
| `cols_no_cero` | `leer_fuente` respeta `cfg["cols_no_cero"]` | P9 |

> **Cada capacidad nueva hay que sumarla a esa lista.** Si no, el script que la use
> se cae con un `TypeError` raro en vez de decir "copiá el archivo actualizado". Y
> peor cuando la capacidad es un **filtro**: no falla nada y entran datos de más.
> Pasó con `cols_no_cero`: sin él, las filas con `Central = 0` habrían entrado en
> silencio.

### `Actualiza_Energia.py`

Los dos entregables de `01.a Sobrecostos de Energia`, los dos desde el
**`02 Consolidado_Tabulado`**, hoja `Sobrecostos`, **datos desde la fila 3** (la 2 es
encabezado). En los dos casos se toman **solo las filas de tipo SCMT y SCPC**
(columna `AB`), que es lo que comprueban V5 y V6.

**Al `.mdb` (`03b ENTRADA_SOB`), tabla `Sobrecostos`:**

| Origen | Destino |
|---|---|
| `AA` | `Clave Año_Mes` |
| `AB` | `Tipo_sobrecosto` |
| `AC` | `Central` |
| `AD` | `Hora Mensual` |
| `AE` | `Sobrecosto` |

Ya vienen en el mismo orden que la tabla. Reutiliza el motor de
`Actualiza_Data_Access.py` (mismo destino, mismas 5 columnas, misma regla de
borrado), así que los dos `.py` tienen que estar en la misma carpeta.

**Al `Consolidado_AAMM`, hoja `Sobrecostos`:**

origen `A:G` + `I:J` + `H` **en ese orden** desde la fila 3 → destino `A:J` desde la
**fila 2** (la 1 es encabezado). Ojo con el orden: la `H` del origen queda al final,
en la columna `J` del destino.

Se limpia `A2:J` de lo anterior antes de pegar, midiendo el largo con el
`used_range` de la hoja: si el mes nuevo trae menos filas, sin limpiar quedarían las
viejas abajo y el total de la columna `E` (lo que suma V6) saldría inflado.

> **Caso no cubierto:** el `DELETE` del Access borra sólo los tipos que **vienen en
> el Excel**. Si un mes desaparece SCPC del tabulado, las filas viejas de SCPC
> **quedan** en el `.mdb`. V5 lo caza, pero se ve como descuadre de monto, no como
> "sobró un tipo".

### `Actualiza_Cuadro0.py`

Deja los **datos y las fórmulas** del cuadro cero al día. **No** llama macros ni
refresca la dinámica: eso se hace a mano (ver sección 5).

Orden de los pasos — **es la cadena de dependencias, no una preferencia**:

```
A:C  →  K  →  L:P y R:U  →  A:G de 01.SSCC  →  C  →  I  →  J:K
```

| Paso | Qué hace |
|---|---|
| 1 | pega `I9:K` del `1_CUADROS` en `A5:C` de la hoja #3 |
| 2 | la tasa en `O2`, si se le da una |
| 3 | ajusta la `D` al largo de la tabla que acaba de pegar |
| 4 | reescribe `K5` y ajusta `L:P` y `R:U` |
| 5 | ajusta `A:G` de `01.SSCC`, que es lo que calcula `C` |
| 6 | reescribe `I9` y ajusta `J:K` |
| 7 | avisa si hay signos invertidos en `J:K` |
| 8 | guarda y deja el libro abierto |

> **La `D` va contra la tabla `A:C`, no contra `K`.** `K` junta las empresas de `A`
> **y** de `F`, así que casi siempre es más larga: en 2312, `K` llegaba a la fila 121
> y la tabla `A:C` a la 110. Ajustar la `D` hasta `K` le agregaría 11 filas de
> `=B+C` sobre celdas vacías.

`I` va **después** de que `C` esté calculada, porque `C` es el `BUSCARV` que aplica
los reemplazos. Correr los pasos por separado deja el libro a medio calcular sin
ninguna señal, y por eso la ventana no tiene casilleros por paso.

Detalles de implementación que importan:

- **Cálculo en manual** de punta a punta, con recálculo explícito sólo en los cuatro
  puntos donde el resultado se necesita para seguir. Con automático, cada AutoFill
  dispara un recálculo completo, y `L5`/`M5` son `SUMAR.SI` de **columna entera**
  repetidos por empresa, y `J9`/`K9` evalúan su `SUMAR.SI` **dos veces** cada uno.
  Cada recálculo queda cronometrado en el log.
- **Por COM las fórmulas se escriben en inglés y con coma**
  (`=LET(x,UNIQUE(VSTACK(...)),FILTER(...))`), aunque en pantalla se vean en español
  con punto y coma. Escribirlas en español por esa vía falla.
- Cada bloque se **limpia hasta su última celda con algo** antes de escribirse.
- La tasa se guarda en `config.json` **por mes** (`tasa_interes_por_mes`), con la
  fecha. Guardada suelta, el mes siguiente arrastraría la del anterior sin que se
  note.

### `Actualiza_SC_CO.py`

La hoja **"SC y CO"** de la planilla 5_, con los SC y los CO de los **embalses** y su
prorrata de instrucción directa. Los dos bloques van en la misma hoja: **SC arriba,
CO abajo**.

Encabezados aparte: los datos arrancan en la fila **7** en los orígenes y en la fila
**9** en el destino.

**SC — `Calculo_SobrecostosSSCC`, hoja `SOBRECOSTOS`:**

| Origen | Destino | Qué es |
|---|---|---|
| `S` | `C` | Clave Año_Mes |
| `T` | `D` | tipo (`SCCF`) |
| `U` | `E` | central ← **por acá se filtra** |
| `V` | `F` | hora mensual |
| `W` | `G` | monto |
| `AU:BJ` | `I:X` | prorrata (16 columnas) |

**CO — `Calculo_CO`, hoja `PRORRATA CO`:**

| Origen | Destino | Qué es |
|---|---|---|
| — | `C` | Clave Año_Mes: **no viene en el origen**, se copia la de los SC |
| `B` | `D` | tipo (`CO`) |
| `D` | `E` | central ← **por acá se filtra** |
| `F` | `G`… ver abajo | |
| `F:G` | `F:G` | hora mensual y monto |
| `BM:CB` | `I:X` | prorrata (16 columnas) |

Las fórmulas del destino son **`Y:AB` y `AD:AF`** (7 columnas): se estiran o se
cortan para cubrir exactamente las filas que quedaron.

> **La `AC` queda afuera**, igual que la `Q` de la hoja #3 del cuadro cero: no lleva
> fórmula. Hay que tratarla en dos bloques y no como un `Y:AF` corrido, porque un
> AutoFill de `Y:AF` de una sola vez rellenaría la `AC` con lo que tenga en la fila 9.
> El actualizador y el verificador usan la misma partición.

Solo se copian las centrales de la lista `CENTRALES_EMBALSE` del propio script (ver
8.b), 27 unidades. El filtro cae en la columna `E` del destino en los dos casos, que es la central.

> **Los dos bloques se reescriben siempre, aunque se elija solo uno.** El bloque CO va
> *debajo* del de SC, así que no son independientes: si los SC de este mes traen más o
> menos embalses que los del mes pasado, el CO tiene que correrse. El bloque que no se
> eligió se lee del propio destino (por la columna `D`: `SCCF` o `CO`) y se vuelve a
> escribir en su lugar nuevo. Sin eso quedaría un hueco o una fila pisada.

La Clave Año_Mes de los CO sale de los SC. Si no hay filas SC de dónde copiarla, usa
la que ya tenían los CO en el destino, y si tampoco, el mes de la ventana.

### `Actualiza_Access_P9.py`

El `.mdb` de la planilla 9 (`Ocupar_este_para_Reliquidacion_AAMM_*.mdb`). Reemplaza
a `para ricardo.py` + `archivo_de_configuracion.yaml`.

**Una casilla por planilla.** Al marcar una se actualizan sus bloques de
`Sobrecostos` **y** sus propietarios: salen del mismo archivo, no tiene sentido
separarlos.

| Planilla | Sobrecostos | Propietarios |
|---|---|---|
| **3** | `DB`, `BC:BG`, desde la fila 3 | `DB`, `M`/`L`, fila 3 |
| **5** | `CÁLCULO_CRA`, `AQ:AU`, desde la 9 | `EMPRESAS`, `B`/`C`, fila 9 |
| **6** | `ENERGIA_Y_CALCULO_CO_ERNC`, `AY:BC`, desde la 9 **y** `CALCULO_REA_CENTRAL`, `AT:AX`, desde la 10 | `CONSUMOS_PROPIOS`, `B`/`H`, fila 9 |

Las filas de inicio salen del `header=` del script viejo (`header=1` → fila 3,
`header=7` → 9, `header=8` → 10). Los bloques van en el orden de la tabla:
`Clave Año_Mes`, `Tipo_sobrecosto`, `Central`, `Hora Mensual`, `Sobrecosto`.

El bloque **BESS** del script viejo (planilla 11, `PRORRATA_RETIROS`, `IV:IZ`) **no
se incluye**: `incluir_bess` estaba en `False`.

Se descartan las filas con **monto 0** o **Central 0/vacía**, que es lo que hacía
`df[df['Pago']!=0]` y `df[df['Central']!=0]`.

> **La `Clave Año_Mes` del origen NO se usa.** Viene mal: siempre trae `23xx` aunque
> el mes sea otro (2405 llega como 2305). Se pisa con el mes sacado del **nombre de
> los archivos** (`..._2502_R01P.xlsm` → 2502), que se muestra en la ventana y se
> puede corregir a mano. Si los tres archivos son de meses distintos, corta y lo dice.

Las dos tablas se **vacían completas** antes de cargar, porque se arman enteras
desde las planillas marcadas. El vaciado ocurre **una sola vez**, antes de insertar:
hacerlo dentro del bucle por fuente borraría lo que insertaron las anteriores. Si se
deja una planilla sin marcar, lo que ella aporta **no queda** en el Access, y la
ventana lo avisa antes de escribir.

En la planilla 3 los propietarios salen de la **misma hoja `DB`**: `K` es la
configuración, `L` el propietario y `M` la unidad infotécnica. La "central" que se
lleva al Access es la **unidad infotécnica**, así que el propietario tiene que salir
de esa misma tabla o los nombres no calzarían con la columna `Central` del bloque de
sobrecostos. Ojo que ahí la empresa (`L`) está **antes** que la central (`M`).

Como esa tabla tiene una fila por registro, la misma unidad aparece repetida cientos
de veces; se deja una sola fila por central (ver abajo).

**Una sola fila por central.** La tabla tiene índice único en `Central`: dos filas
con la misma central hacen fallar el `INSERT` entero con *"crearían valores
duplicados en el índice"*. No alcanza con quitar los pares `(Central, Empresa)`
repetidos, porque la misma central puede venir de dos planillas, o repetida con
empresas distintas.

| Situación | Qué pasa |
|---|---|
| la misma central con la misma empresa, en dos planillas | se deja una |
| en una planilla sin dueño y en otra con dueño | **gana la que tiene dueño** |
| dos empresas **distintas** | queda la primera (orden 3 → 5 → 6) y se avisa con las dos y de qué planilla vino cada una |
| escrita distinto (`EL TORO-1` / `ELTORO-1`) | cuentan como la misma |

La comparación es sobre el nombre normalizado (sin tildes, espacios ni guiones
bajos), que es **más estricto** que Access: así no se mandan dos filas que el índice
consideraría iguales.

> **Las centrales sin propietario se conservan.** En `CONSUMOS_PROPIOS` hay centrales
> con la empresa vacía o en 0, y son datos válidos. Por eso el corte de esa hoja es la
> primera celda vacía de la **central** (columna `B`) y no de la empresa: cortando por
> la `H` se perderían. Se cargan con la empresa en `NULL`.
>
> Una central sin dueño **no es un error por sí misma**. Lo que sí lo es: que una
> central **con monto** no tenga dueño. Eso lo comprueba V12 en el Revisor.

**Qué se cambió respecto del script viejo, y por qué:**

1. **Los bloques se leen por posición, no por nombre de columna.** El original hacía
   `pd.concat`, que alinea por **nombre**: si una hoja tenía el encabezado escrito
   distinto (un `Pago ` con espacio), aparecían columnas extra con NaN y el `INSERT`
   de 5 parámetros se desalineaba o fallaba. Igual se lee la fila de encabezado y se
   **avisa** si no se parece a lo esperado.
2. **Transacción única con rollback.** El original usaba `autocommit=True`: si el
   `INSERT` fallaba después del `DELETE *`, la tabla quedaba vacía o a medias sin
   aviso.
3. **Verificación después de cargar**, la del motor de `Actualiza_Data_Access.py`,
   que se reutiliza en vez de repetir la parte de pyodbc.
4. **Modo SOLO MIRAR**, marcado por defecto.
5. El `df_ricardo_salida.xlsx` que el original escribía siempre ahora es opcional, y
   queda junto al `.mdb`.

### Lectura de planillas sin abrir Excel

`Actualiza_Data_Access.py`, `Actualiza_Energia.py` y `Actualiza_Access_P9.py` leen
las planillas abriendo el `.xlsx`/`.xlsm` como **ZIP** y escaneando el XML, sin
levantar Excel. Es el mismo lector que usa el Revisor.

Importa porque esos scripts **sólo leen** esas planillas: abrirlas con xlwings
levanta Excel entero por COM, que en la T: son varios minutos por archivo. El
escaneo del XML no depende de Excel y no carga el archivo entero en memoria.

Si el archivo no es un `.xlsx`/`.xlsm` (un `.xls` viejo, por ejemplo) o el escaneo
falla, **cae automáticamente a xlwings**, que es el camino de antes. Con
`cfg["lectura_rapida"] = False` se fuerza el camino viejo.

Verificado que los dos caminos dan **exactamente las mismas filas** sobre la misma
planilla.

> **Limitación**: se lee el resultado guardado de cada celda, nunca la fórmula. Si
> el archivo nunca se recalculó, los resultados pueden ser viejos. Excel guarda el
> resultado cada vez que se guarda el archivo, así que en la práctica no molesta.
> Los archivos que se **escriben** (los destinos) siguen yendo por xlwings.

### `Prorratear.py`

Automatiza lo que se hacía a mano en SQL Server Management Studio. Botón
**Prorratear** en los **tres** `.mdb`, porque sirve cualquiera de ellos.

**Las bases van de a pares.** La de sobrecostos depende de dónde están los retiros,
y elegir mal significa prorratear contra los retiros equivocados. Por eso en la
ventana se elige el **escenario**, no las dos bases por separado:

| Escenario | Retiros | Sobrecostos |
|---|---|---|
| Normal | `02_RETIROS` | `05_SOBRECOSTOS` |
| Reliquidación | `14_RETIROS_RELIQUIDACION` | `16_SOBRECOSTOS_RELIQUIDACION` |

Qué hace, en orden:

1. **Borra** de la base de sobrecostos, *sólo si existen*: `Central_Empresa`,
   `Pago_Retiro_reporte_tabla`, `Sobrecostos`, `TIPOS`.
2. **Copia del Access** (el `Tasks → Import Data`):
   `Central_Empresa_Actualizada` → `Central_Empresa`, `Sobrecostos`, `TIPOS`.
   El `.mdb` de la planilla 9 no tiene la `_Actualizada`: ahí la tabla ya se llama
   `Central_Empresa` y se copia tal cual. El script detecta cuál hay.
3. **Arma** `Pago_Retiro_reporte_tabla` con el `SELECT … INTO … GROUP BY` desde la
   vista `dbo.[10_Pago_retiros]`, que ya existe en la base y cruza lo recién
   importado con los retiros.

El Access **sólo se lee**.

**Lo primero que comprueba es el mes.** Prorratear los sobrecostos de un mes contra
los retiros de otro no da error: da un resultado **mal**, y como el número sale igual
de plausible, no se nota después. Así que antes de nada compara el `Clave Año_Mes` de
la tabla `Sobrecostos` del Access contra los períodos que hay en la tabla `Retiros` de
la base que corresponde al escenario. Corta si:

- el mes del Access **no está** entre los retiros cargados
- el Access tiene **más de un mes** en `Sobrecostos`
- la tabla `Retiros` está **vacía**

Si coincide pero la base tiene además otros períodos, lo avisa: eso es lo que hace más
lenta la prorrata cuando la vista no filtra por mes.

**Antes de borrar nada** se comprueba también que la vista `10_Pago_retiros` exista y
que el Access tenga las tres tablas. Si falta algo, corta sin tocar la base — si no, se
borrarían las tablas para descubrir después que no se puede terminar.

Hay un modo **SOLO MIRAR**, marcado por defecto, que cuenta todo y dice qué haría sin
escribir.

> Si falla **después** del borrado, la base queda incompleta. El log lo dice claro:
> hay que volver a correrlo, que borra lo que haya y carga de nuevo.

Al terminar informa cuántas filas quedaron en el reporte y la suma de `Pago`. Si el
reporte queda **vacío**, suele significar que los retiros de esa base no son de este
mes, o que la vista no encontró con qué cruzar.

### `Carga_Retiros.py`

Carga `04 Planilla 9/Retiros_h.parquet` a **SQL Server**. Es el único script que
escribe en una base de datos y no en un Excel.

| | |
|---|---|
| Servidor | `SRV-DTE` |
| Tabla | `Retiros` |
| Base | a elegir: `02_RETIROS` o `14_RETIROS_RELIQUIDACIÓN` |
| Driver | `ODBC Driver 17 for SQL Server`, conexión de confianza |

Qué hace: lee el parquet → **borra** del servidor → carga por trozos de 50.000 →
verifica que la cuenta cuadre.

**Siempre vacía la tabla** (`TRUNCATE`) y carga de cero. El borrado por período y
el modo "solo mirar" se sacaron porque no se usaban, y complicaban la verificación
del final.

### El cambio de hora

En el mes del cambio de horario de primavera hay una hora **que no existe** (la 145,
por ejemplo). El archivo suele venir con las horas **corridas**: llega hasta la 719 en
vez de la 720, y lo que está guardado como 145 es en realidad la 146.

Con la casilla **«Cambio de hora (−)»** y la hora del mes, al cargar se empuja una
hora todo lo que esté desde ahí. Queda 1…144 y 146…720, con la 145 vacía, que es lo
correcto. **El parquet no se modifica**: el desplazamiento es sólo al cargar.

| Situación | Qué hace |
|---|---|
| el archivo trae la hora del cambio (1…719) | la aplica |
| el archivo ya viene corrido (1…720 sin la 145) | **no** la aplica, y lo dice |
| el mes está completo (1…744) | **no** la aplica: empujar dejaría 745 |

> La detección de "ya viene aplicado" es mirar si **existe** la hora del cambio en los
> datos. Es la comprobación correcta porque el desplazamiento es justamente lo que
> deja ese hueco: no se puede aplicar dos veces sin que se note. Verificado corriéndolo
> dos veces seguidas.

La ruta del parquet sale del Revisor, o se elige la carpeta del caso, o se elige el
`.parquet` directo con «Examinar archivo». La base elegida se recuerda en
`config.json` (`retiros_base`).

**Los nombres de columna no se escriben igual en todos lados.** El parquet a veces
trae `Clave Año_Mes` (con espacio y ñ) y a veces `Clave_Anio_Mes` o
`Clave_anio_mes`, y la tabla de SQL Server tiene el suyo. Se resuelven comparando
sin tildes, espacios ni guiones bajos, y tratando `ANIO` como `ANO` — la ñ no se
arregla quitando tildes: descomponer `Año` da `ANO` y `Anio` da `ANIO`.

Y no alcanza con encontrarlos: `to_sql` mapea por **nombre**, así que si el parquet y
la tabla los escriben distinto, la carga tampoco calza. Antes de cargar se renombran
las columnas del parquet a como se llaman en la tabla, y el log lo dice:

```
Columnas de la tabla: Clave_Anio_Mes, Suministrador, Medida_kWh
Renombrando 1 columna(s) para que calcen con la tabla:
   'Clave Año_Mes'  ->  'Clave_Anio_Mes'
```

Si el parquet trae una columna que la tabla **no tiene**, se aborta antes de tocar la
base. Si la tabla tiene una que el parquet no trae, se avisa (queda en `NULL`).

> **Si falla a mitad de la carga**, el borrado ya se hizo y la tabla queda a medias.
> No es grave: volver a correrlo borra ese período otra vez y recarga. Lo que no hay
> que hacer es dejarla así. El script lo dice en el log cuando falla.

Diferencias con el script original de una sola corrida:

- El `BORRAR_PERIODO = 2412` del original **no se usaba**: el borrado siempre iba por
  los períodos que trae el parquet. Se quitó para que no confunda.
- La alerta de "residuo de otras cargas" por cantidad de suministradores solo salta
  con TRUNCATE. Con borrado por período es **normal** que el servidor tenga más
  suministradores que el parquet, porque los otros períodos aportan los suyos; antes
  eso alertaba siempre.
- Se agregó una comprobación de que los períodos del parquet **quedaron** en el
  servidor, y que el parquet traiga las columnas que se usan.

El nodo del árbol es informativo: no se verifica su contenido (es un parquet, no un
Excel). Lo que valida la carga es el propio script, contra el servidor.

### `Reemplazos REUC/ActualizaRemplazos.py`

Genera la tabla de empresas con RUT y los reemplazos, y los pega en el cuadro cero:
`B:C` las empresas, `H:I` los reemplazos (`Reemplazada` / `Reemplazante`), fila 1 de
encabezado en los dos bloques. Usa su **propio** `config.json`, en
`Reemplazos REUC/Auxiliares/`.

> **Orden corregido.** Los reemplazos forzados (`Reemplazos forzados*.xlsx`, hoja
> `Reemplazos forzados`, columnas `Reemplazada` / `Reemplazante`) se leen y se
> aplican **antes** del cruce con el registro de empresas. Antes se aplicaban
> después, y como las filas sin RUT ya se habían descartado, un forzado nunca podía
> arreglar el nombre que no coincidía: la alerta volvía todos los meses **y la
> empresa desaparecía del entregable en silencio**. Verificado: 3 empresas por
> $6.000 salían como $3.000.
>
> El mapa se aplica dos veces (antes del cruce y después de los reemplazos del
> REUC), así que si un `Reemplazante` es a la vez `Reemplazada`, el nombre salta dos
> escalones. El script lo avisa en el log.

---

## 8. Integración con el revisor

Cada fila maestra del árbol tiene su botón; las copias de `00 Entregables` no, porque
se actualizan copiando el maestro.

| Fila (`id`) | Botón | Script |
|---|---|---|
| `a_calc_sscc_01` | Actualizar data | `Actualiza_datos.py` (`planilla="sc"`) |
| `a_3_p9` | Actualizar data | `Actualiza_datos.py` (`p3`) |
| `a_ocupar` | Actualizar data | `Actualiza_Access_P9.py` |
| `a_retiros_parq` | Cargar retiros | `Carga_Retiros.py` |
| `a_mdb_sscc`, `a_mdb_sob`, `a_ocupar` | Prorratear | `Prorratear.py` |
| `a_5_p9` | Actualizar data | `Actualiza_datos.py` (`p5`) |
| `a_5_p9` | Actualizar "SC y CO" | `Actualiza_SC_CO.py` |
| `a_6_p9` | Actualizar data | `Actualiza_datos.py` (`p6`) |
| `a_mdb_sscc` | Actualizar data | `Actualiza_Data_Access.py` |
| `a_mdb_sob` | Actualizar data | `Actualiza_Energia.py` |
| `a_consolidado` | Actualizar data | `Actualiza_Energia.py` |
| `a_0_cuadros` | Actualizar reemplazos | `Reemplazos REUC/ActualizaRemplazos.py` |
| `a_0_cuadros` | Actualizar cuadro 0 | `Actualiza_Cuadro0.py` |
| `a_0_cuadros` | Exportar CPRT | interno, no lanza proceso |

Además, **toda fila que declare `espejo`** (o sea, que sea copia de otra) lleva un
botón **"Traer maestro"**, que rehace la copia desde su maestro. Son 6:
`d_ent_sob`, `a_calc_sscc_ent`, `a_3_ent`, `a_5_ent`, `a_6_ent` y `d_sobe`. No hay
que listarlas a mano: el `espejo` ya dice cuáles son.

### "Traer maestro"

Reemplaza el copiar y pegar a mano. Antes de tocar nada arma el plan y lo muestra:

| Situación | Qué hace |
|---|---|
| la copia coincide con el maestro | nada, avisa que ya está al día |
| la copia está vieja | la reemplaza |
| la copia tiene **otro nombre** (revisión vieja) | **borra** la vieja y copia la del maestro |
| no hay copia | la crea |
| en los diarios, un archivo que **sobra** | lo **borra** |

Detalles que importan:

- Se copia con `shutil.copy2`, que **conserva la fecha de modificación**. Con un
  `copy()` a secas la copia quedaría con la fecha de ahora y el revisor la seguiría
  marcando en amarillo, porque compara justamente por fecha.
- Antes de copiar se comprueba que **ningún destino esté abierto en Excel**. Si uno
  está tomado no se hace nada: a medio camino quedaría la carpeta con una parte
  vieja y otra nueva.
- Es **idempotente**: correrlo dos veces seguidas, la segunda no hace nada.
- El plan se arma en `plan_traer_maestro()`, separado del ejecutor, para poder
  mostrarlo antes y probarlo sin borrar nada.

Los dos de Energía abren la **misma** ventana; desde ahí se elige el Access, el
Consolidado o los dos. Los tres del cuadro cero van apilados en una columna, porque
en una sola línea la fila se sale de la pantalla.

### El JSON de traspaso

El revisor escribe `Salidas/AAMM/_traspaso_actualizador.json` y le pasa esa ruta como
único argumento. **Sin argumento, cada script funciona como siempre** y busca los
archivos por su cuenta: es la vía de escape si el revisor no está.

```json
{
  "origen": "Revisor_Reliquidacion",
  "version": 1,
  "aamm": "2407",
  "carpeta_reliq": "T:\\...\\02 CASO RELIQUIDACION",
  "planilla": "p3",
  "rutas": {
    "sscc_desempeno":       "...\\FD\\SSCC_Desempeno_2407.xlsx",
    "consolidado_tabulado": "...\\01 Sobrecostos\\Detalles diarios\\02 Consolidado_Tabulado_...",
    "prorrata_retiros":     "...\\04 Planilla 9\\Prorrata_Retiros_...",
    "calculo_sscc_maestro": "...\\01 Sobrecostos\\Cálculo_SobrecostosSSCC_...",
    "calculo_co":           "...\\00 Entregables\\02 Costo de Oportunidad\\Cálculo_CO_...",
    "consolidado_cca":      "...\\00 Entregables\\03 Costo de Combustible Adicional\\...",
    "p3": "...", "p5": "...", "p6": "...",
    "mdb_sscc":             "...\\01 Sobrecostos\\03b ENTRADA_SOB_SSCC_...",
    "mdb_sob":              "...\\01.a Sobrecostos de Energia\\03b ENTRADA_SOB_...",
    "consolidado_energia":  "...\\01.a Sobrecostos de Energia\\Consolidado_...",
    "cuadro_0":             "...\\00 Entregables\\0_CUADROS_RELIQUIDACIÓN...",
    "cuadro_1":             "...\\00 Entregables\\1_CUADROS_PAGO_SSCC_..."
  }
}
```

Además de `rutas`, el JSON trae **de qué fila se apretó el botón**:

```json
"nodo": "a_ocupar",
"clave_nodo": "mdb_ocupar",
"ruta_nodo": "...\\04 Planilla 9\\Ocupar_este_para_Reliquidacion_2502_R01P.mdb"
```

Hace falta cuando el mismo script cuelga de **varias filas**: `Prorratear.py` está
en los tres `.mdb`, y sin esto no puede saber a cuál le dieron. Antes usaba siempre
el de SSCC aunque se hubiera apretado el de la planilla 9.

Un script que reciba un JSON sin `ruta_nodo` (de un Revisor anterior) sigue
funcionando: cae al orden de siempre y lo avisa en el log.

Al recibirlo, el actualizador rellena sus rutas con esos valores, **no vuelve a
buscar**, deja los casilleros múltiples desmarcados y muestra de qué mes viene. Si el
usuario elige una carpeta a mano, se sale del modo traspaso y vuelve a buscar solo.

Se escriben solo las claves que el revisor pudo resolver; el actualizador tolera que
falte alguna y la muestra en rojo.

### Comportamiento del botón

- Deshabilitado si falta el archivo destino, y mientras corren verificaciones.
- Antes de lanzar se **releen del disco** el destino y las rutas del traspaso, por si
  apareció una revisión nueva (un `R01E` donde había `R01D`).
- Si el destino es un Excel, se comprueba que se pueda **escribir** (abrirlo en
  `r+b`), no que exista el `~$`: Excel deja ese archivo huérfano cuando se cae, y
  entonces existe para siempre aunque el libro esté cerrado. El `~$` sólo se usa como
  pista de quién lo tiene.
- Todo queda en la bitácora: script, mes, destino y rutas enviadas.
- No se vigila el proceso ni se refresca solo. Al terminar hay que apretar
  ACTUALIZAR.

### Caché de directorios

Una relectura completa hacía **68 recorridos de carpeta para 13 carpetas
distintas**: cada nodo del árbol recorría su carpeta de nuevo, y `resolver_carpeta`
recorría la raíz una vez por nodo. En un disco local no se nota; en la T: cada
recorrido es un viaje de red.

Ahora el listado de cada carpeta se guarda **mientras dura la relectura**, así cada
carpeta se recorre una vez: de **70 recorridos a 15**. Y se usa `os.scandir` en vez
de `iterdir`, que trae la fecha y el tamaño en el mismo recorrido — antes cada
archivo candidato costaba un `stat()` aparte.

Medido con latencia simulada por carpeta:

| Latencia | Antes | Ahora | |
|---|---|---|---|
| 5 ms | 0,37 s | 0,08 s | 4,4× |
| 20 ms | 1,44 s | 0,31 s | 4,6× |
| 50 ms | 3,54 s | 0,77 s | 4,6× |

El caché está **apagado por omisión** y sólo se enciende dentro de
`with cache_directorios():`, que envuelve el trabajo de disco de `actualizar()`.
Fuera de ahí todo lee del disco como siempre, y eso importa: `mtime()` se usa
después para decidir si una verificación venció, y con datos viejos daría mal. Se
apaga también si algo falla dentro del bloque.

Verificado que las rutas y los diarios salen **idénticos** con y sin caché, que
`mtime` y `tamaño` del caché coinciden con los del disco, y que se siguen
descartando las copias y los `~$`.

El log dice cuánto se ahorró: `15 carpeta(s) recorrida(s), 86 lectura(s) servida(s)
del caché (antes serían 101)`.

### Relectura parcial

Antes de verificar o de lanzar un actualizador, el revisor relee **solo los nodos que
hacen falta**, no el árbol completo. Ninguna verificación necesita las 5 carpetas de
`Detalles diarios`, que es lo caro en un disco de red. Verificar V9 relee 1 nodo de
27; V16 relee 13 (arrastra la cadena de previas); el botón ACTUALIZAR sigue releyendo
todo.

La primera vez relee todo aunque se pida parcial: sin nada en memoria no se puede
leer a medias. En modo parcial no se re-detecta el mes ni se recarga el estado.

---

## 8.b La lista de centrales de embalse

`CENTRALES_EMBALSE`, las 27 unidades (con ANGOSTURA-1/2/3). Está **en dos lugares**: en
`Actualiza_SC_CO.py` (para filtrar el origen) y en `Revisor_Reliquidacion.py` (para
verificar). Al agregar o quitar una central hay que cambiarla **en los dos**, con el
nombre **exacto** como viene en el origen.

Si se desincronizan, **V10 lo caza en los dos sentidos**, así que la lista vieja no
pasa en silencio:

| Desincronización | Cómo se nota |
|---|---|
| el actualizador tiene una central que el revisor no | se pega, y la comprobación 4 la marca como "no es embalse" |
| el revisor tiene una que el actualizador no | no se pega, y la 6/7 la marca como «-número» fuera de la lista |

Verificado con los dos casos simulados.

Los nombres se comparan normalizados (sin tildes, sin espacios ni guiones bajos, en
mayúsculas), así que `El Toro-1`, `EL_TORO-1` y `ELTORO-1` son la misma central.

---

## 9. `config.json`

Compartido por el revisor y los tres actualizadores que viven junto a él (no por
Reemplazos REUC, que tiene el suyo). Indexado por `<host>_<usuario>`.

Se escribe de forma **atómica** (`.tmp` + `os.replace`) y **no se escribe** si el
archivo existe pero no se puede interpretar: mejor perder un ajuste que el archivo
entero. Sólo se agregan o actualizan claves propias, nunca se borra nada.

Claves conocidas: `carpeta_base`, `carpeta_reliq`, `mdb`, `ultimo_mes`,
`carpetas_por_mes`, `sel_SSCC`/`sel_CO`/`sel_CCA`, `ene_sel_mdb`/`ene_sel_cons`,
`tasa_interes_por_mes`, `_valores`.

> `carpeta_base` (revisor) y `carpeta_reliq` (`Actualiza_datos.py`) son **dos claves
> distintas para lo mismo**. Al apretar un botón de actualización, el revisor escribe
> `carpeta_reliq` y `mdb` del mes en curso, así que si después se abre un
> actualizador a mano ya arranca en el mes correcto.

---

## 10. Cosas por confirmar

- **De dónde salen los RUT del CPRT.** No son los de la hoja `EMPRESAS`: 60 de 481
  difieren. Vienen de la conexión al libro externo (ver 4.9). Es lo que bloquea
  generar el csv sin la dinámica.
- **El umbral de monto del CPRT**: está entre 8,74 y 22,33 (ver 4.9).
- **La tasa de `O2`**: se escribe a mano; no está documentado de dónde viene el
  número.
- **`Tipo:` de la cabecera del CPRT** dice `R02D` mientras el archivo se llama
  `R01D`. Puede ser un descuido de ese mes.
- **El `1_CUADROS_PAGO` que busca `ActualizaRemplazos.py`** sale del árbol de
  `T:\Facturacion\<mes>\<versión>`, no del `00 Entregables` que tiene el revisor.
  Falta confirmar si es el mismo archivo; por eso el revisor todavía no le pasa esa
  ruta (la clave `cuadro_1` va en el JSON pero el script no la usa).
- **Si `K` e `I` se dejan como fórmula o se pegan como valores.** El archivo de 2312
  las tenía como valores, "porque así se hacía antes". Con fórmula viva hay que
  cuidar el desbordamiento; con valores, hay que reescribirlas cada mes.

### Corregido respecto de versiones anteriores de este documento

- La hoja del `0_CUADROS` para las tablas A:C es la **#3**, no `VERIFICADORES`. Y las
  filas de inicio son **5** y **9**, no 4 y 8.
- Las columnas `F`, `K`, `R` de la hoja #3 no son "listados de empresas para
  control": `K` es la lista consolidada y `R` es `=+K5`.
- La columna `C` de `01.SSCC_Recurso_Técnico` sí contiene nombres de empresa: es el
  `BUSCARV` que aplica los reemplazos. **Confirmado.**
- La carpeta `FD` no forma parte de `02 CASO RELIQUIDACION`; está en el árbol del
  revisor como nodo informativo (solo nombre y fecha, sin verificaciones).
