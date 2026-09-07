# BITÁCORA — registro de sesiones

> Se agrega, nunca se edita ni se borra una entrada vieja — ni siquiera la
> propia (ver `REGLAS.md`). Entradas más nuevas arriba. Esto es lo que un
> `git log` no cuenta: qué quedó pendiente, qué se probó y qué no, qué se
> decidió y por qué.
>
> La única excepción es la sección **"Pendientes abiertos"**, de acá abajo:
> esa sí se edita — se tacha o se saca el ítem cuando se resuelve, porque es
> un estado actual, no un historial. Quien resuelve un pendiente edita esta
> lista **y** agrega la entrada correspondiente en la bitácora.
>
> Formato de cada entrada: fecha · quién · qué se hizo · qué queda
> pendiente. El hash entre paréntesis ubica el commit exacto en `git log`.

---

## Pendientes abiertos ahora mismo

- [ ] **Validar el Revisor modularizado de punta a punta en Windows con un mes
      real:** arranque de la ventana, árbol completo, V4…V17, OOXML y fallback
      Excel, lectura Access, estado/caché, lanzamiento de un actualizador e
      igualdad de `_traspaso_actualizador.json`. La extracción de la Fase 4 está
      completa; este control requiere Excel, Access y los archivos reales.
- [ ] **La usuaria tiene que armar `__config__/` a mano** (no se sube al
      repo, está gitignoreada), con la misma estructura `AAAA/MM Mes` que
      `00_Salidas`, migrando ahí sus `config.json`,
      `_revisor_verificaciones.json`, `_revisor_cache_valores.json`,
      `_traspaso_actualizador.json`, `estado.json`, `rutas.json`, `parquet/`
      y `vistas/` existentes — el código no los mueve solo. El propio
      `config.json` compartido va en `__config__/config.json` y el de
      `ActualizaRemplazos.py` en `__config__/reemplazos_reuc.json`.
- [ ] Queda abierta una pregunta chica, no decidida con la usuaria: los dos
      comparadores guardan sus respaldos de los últimos 5 Excel anuales en
      `00_Salidas/AAAA/respaldos/` (los dos comparadores comparten esa misma
      carpeta, sin colisión porque el nombre del archivo ya distingue cuál es
      cuál). No es exactamente un "resultado", pero tampoco es state/caché —
      se dejó ahí a criterio propio al corregir la Tarea 2. Confirmar con la
      usuaria si eso también debería vivir en `__config__/AAAA/_comparador*/`
      en vez de en `00_Salidas/`.
- [ ] `docs/ESTRUCTURA_CASO_RELIQUIDACION.md` tiene 5 diferencias conocidas
      contra el código real, listadas en `MAPA.md` → "Diferencias con el
      documento de dominio". El documento de dominio todavía no se corrigió.
- [ ] Confirmar si el `1_CUADROS_PAGO` que busca `ActualizaRemplazos.py` en
      `T:\Facturacion\<mes>\<versión>` es el mismo archivo que el
      `00 Entregables` que usa el Revisor (documento de dominio, sección 10).
- [ ] Probar visualmente en Windows los temas claro y oscuro de los dos
      comparadores. La verificación automatizada corrió con `tkinter` real
      (instalado en este entorno) y `ttk.Style` simulado, pero sin pantalla no
      hay forma de ver si el resultado es realmente legible/prolijo.
---

## 2026-09-07 — Claude — un ayudante que se mudó y su llamador que se quedó

Segundo hallazgo de la corrida en Windows. En el log del Revisor:

```
1_CUADROS_PAGO  suma C17:C25  <-  RESUMEN!C17:C25
    · openpyxl no pudo leer (name '_suma_rango_openpyxl' is not defined);
      reintentando con Excel...
      = 400.338.715,23
```

**No era un número mal calculado.** El total salía bien — por eso el error se
veía como una línea más del log y no como una falla. Lo que estaba roto era el
**camino rápido**: `_suma_rango_openpyxl` se mudó a `revisor/lectores.py` al
completar la Fase 4, pero `leer_valor_excel` se quedó en el punto de entrada y
lo sigue usando, y el import de `revisor.lectores` no lo trae por ser privado.
Como el uso está dentro de un `try/except`, el `NameError` se tragaba y caía al
reintento por COM.

**Por qué importa igual:** ese camino existe para no levantar Excel en la unidad
de red T:, donde cada archivo cuesta minutos (está en las convenciones de
`AGENTS.md`). Con esto, **cada rango que se lee abría Excel**. Resultado
correcto, proceso mucho más lento.

**Arreglo:** el punto de entrada importa `_suma_rango_openpyxl` junto a los
demás nombres de `lectores`, con un comentario de por qué un nombre privado
cruza el límite del módulo.

**Por qué no lo cacé.** En la revisión del cierre comprobé que los tres módulos
resolvieran los nombres que reciben inyectados — pero no la dirección inversa:
que el punto de entrada siguiera resolviendo los suyos después de la mudanza. Es
el mismo error de encuadre que con los wrappers: miré el código que se movía y
no el que quedaba. Van dos veces; queda como trampa en `AGENTS.md`.

**Prueba nueva** en `test_modulos_fase4.py`: recorre por AST todos los nombres
que usa `Revisor_Reliquidacion.py` y falla si alguno no está definido ni
importado. Comprobado que falla al quitar el import y que pasa con él. Se
verificó además que los tres módulos siguen sin nombres sin resolver.

**Pendiente:** quedan del checklist las V4…V17 contra los valores de antes, que
el actualizador lanzado reciba su JSON, y el reúso del caché entre corridas.

---

## 2026-09-07 — Claude — la corrida en Windows encontró dos wrappers rotos

La usuaria probó el Revisor ya partido y avisó: **el botón "Prorratear" abría la
ventana de Actualiza_Data_Access.** Es una regresión real de la Fase 4, y de las
que no avisan.

**Causa.** Al extraer los motores, los métodos que quedaron delegando repetían
el valor por omisión en vez de reenviar el argumento:

```
def _lanzar_actualizador(self, nid, indice=0):
    return _lanzamiento.lanzar_actualizador(self, nid, indice=0)   # <- indice
def _armar_traspaso(self, aamm, planilla, nid=None):
    return _lanzamiento.armar_traspaso(self, aamm, planilla, nid=None)  # <- nid
```

El botón sí manda bien su índice (`command=lambda i=nodo["id"], j=k: ...`); se
perdía una capa más abajo.

**Era más ancho de lo que se vio.** No es solo Prorratear: el **segundo botón de
las cinco filas que tienen dos** lanzaba el script de la primera.

| Fila | Segundo botón | Lanzaba en realidad |
|---|---|---|
| `a_ocupar` | Prorratear | Actualiza_Access_P9 |
| `a_5_p9` | Actualizar "SC y CO" | Actualiza_datos |
| `a_mdb_sscc` | Prorratear | Actualiza_Data_Access |
| `a_mdb_sob` | Prorratear | Actualiza_Energia |
| `a_0_cuadros` | Actualizar cuadro 0 | ActualizaRemplazos |

**El segundo, más silencioso.** `_armar_traspaso` perdiendo `nid` dejaba el JSON
sin `nodo`, `ruta_nodo` ni `clave_nodo`. Nada falla: el actualizador arranca y
tiene que adivinar desde qué fila lo llamaron — exactamente el caso que
`AGENTS.md` documenta para Prorratear, que cuelga de los tres `.mdb`.

**Por qué no lo cacé en la revisión anterior.** Comparé por AST las funciones
*movidas* y salieron idénticas, y probé `armar_traspaso` llamando al módulo
directo. Los wrappers de dos líneas que quedaron en el punto de entrada no
entraron en ninguna de las dos comprobaciones. La lección: al extraer, el
riesgo no está solo en el código que se mueve sino en el pegamento que queda.

**Arreglo:** los dos wrappers reenvían (`indice=indice`, `nid=nid`), con un
comentario que explica por qué. Verificado en caliente: los 10 botones de las
cinco filas dobles lanzan ahora el script que dice su texto, y el traspaso
vuelve a llevar `nodo`/`clave_nodo`/`ruta_nodo` cuando corresponde.

**Prueba que cubre toda la clase**, no solo estos dos: `test_modulos_fase4.py`
recorre por AST *todos* los wrappers que delegan en `revisor/*.py` y falla si
alguno no reenvía un argumento. Comprobado que falla con cada uno de los dos
bugs reintroducido por separado, y que pasa con los dos arreglados. (Al
agregarla quedó primero debajo de `unittest.main()` y no corría; se detectó
justamente por exigirle que fallara.)

**Pendiente:** repetir en Windows los otros puntos del checklist — que las
V4…V17 den lo mismo que antes, que el actualizador lanzado reciba su JSON y que
el caché de valores se reuse entre corridas.

---

## 2026-09-07 — Claude — revisa el cierre de la Fase 4: sin correcciones

Revisión de `codex/ejecutar-script-de-sincronizacion` (el commit que completa la
extracción: `verificaciones.py`, `lanzamiento.py` y la parte MDB de
`lectores.py`; 2.048 líneas menos en el punto de entrada). **Segunda vez
seguida sin nada que corregir.**

**Es un movimiento puro, comprobado por AST, no leyendo el diff:**

- `_comprobar` → `verificaciones.comprobar`: las **21 ramas** de tipo de
  verificación son idénticas árbol contra árbol. La única diferencia en todo el
  cuerpo es la sangría de un docstring dentro de `pertenencia`, por el
  desangrado de método a función. Sin efecto.
- Las 11 funciones movidas a `lanzamiento.py` y a la parte MDB de
  `lectores.py`: idénticas (salvo esa misma sangría de docstrings).
- Los **97 nombres de nivel superior** que el Revisor tenía siguen accesibles.
  `NODOS`=36, `VERIFICADORES`=14 (V4…V17), `CENTRALES_EMBALSE`=27, tolerancias
  y umbrales intactos.

**El mecanismo nuevo — `configurar(globals())` — está bien armado.** Es el punto
que más riesgo tenía, porque copia un *snapshot* del espacio del Revisor: todo
lo que se definiera DESPUÉS de esa llamada quedaría fuera y reventaría con un
`NameError` recién al correr esa verificación. Comprobado:

- los **49 nombres** que los tres módulos toman inyectados están todos definidos
  ANTES de la llamada a `configurar` (línea 2061); ninguno llega tarde y ninguno
  falta;
- **ninguno se reasigna después**, ni a nivel de módulo ni vía `global`, así que
  el snapshot no puede quedar viejo;
- en caliente, cada uno de esos nombres apunta **al mismo objeto** que en el
  Revisor, `CACHE` incluido.

**El contrato del traspaso, que es lo que van a leer los actualizadores:**
`armar_traspaso` da un JSON byte a byte igual al de antes en 6 escenarios (sin
rutas, con una, con las 16, con y sin `planilla`, con `nodo` inexistente), y los
6 pasan `__comun__/traspaso.validar()`. `ARCHIVO_TRASPASO` y
`TRASPASO_VERSION`=1 sin cambios.

**Lo demás:** guardas de import puestas y verificadas (sin `verificaciones.py`
el Revisor aborta con el mensaje correcto); ningún nombre definido y a la vez
importado; las 12 suites pasan; `INTERFACES.md` al día; y los comentarios
viajaron con el código — 178 comentarios en las 1.663 líneas de
`verificaciones.py`, incluido el de "parece un problema de signo".

**Estado real:** la *separación de código* de la Fase 4 está completa, y el plan
lo dice sin sobreafirmar: "pendiente validación integral en Windows". Eso es
correcto. Nada de lo verificado acá reemplaza abrir el Revisor con un mes real:
desde Linux no hay Excel, Access ni SQL Server, y lo que se comprobó es que el
código es **el mismo**, no que el proceso completo corra.

**Pendiente y bloqueante para cerrar:** la corrida real en Windows. Checklist
mínimo — que la ventana abra y arme el árbol del mes; que corran las V4…V17 y
den lo mismo que antes de partir el archivo; que un actualizador lanzado desde
una fila reciba su `_traspaso_actualizador.json`; y que el caché de valores se
reuse entre dos corridas seguidas.

---

## 2026-09-07 — ChatGPT — completa la extracción de la Fase 4

Se completaron los tres bloques que quedaban de la división interna del
Revisor. `revisor/lectores.py` absorbe ahora los lectores MDB y la resolución
cacheada de valores; `revisor/verificaciones.py` contiene sin simplificaciones
el motor de comprobaciones V4…V17; y `revisor/lanzamiento.py` contiene las
pruebas de bloqueo de Excel, el armado del traspaso y el lanzamiento de los
actualizadores. `Revisor_Reliquidacion.py` sigue siendo el único punto de
entrada y conserva wrappers con todos los nombres históricos.

Los módulos nuevos no importan el punto de entrada: reciben sus dependencias una
vez al cargarlo. La inyección evita expresamente pisar funciones propias si hay
un nombre homónimo en el espacio histórico, caso cubierto por una prueba de
regresión. También se prueban el sobre del traspaso, la comprobación directa de
escritura y la salida defensiva para un tipo de verificación desconocido.

**Pendiente:** la validación integral en Windows con un mes real sigue siendo
obligatoria y se agregó a "Pendientes abiertos". El contenedor Linux no dispone
de Excel/COM, Access ni los archivos de trabajo, por lo que no puede reemplazar
esa corrida.

## 2026-09-07 — Claude — el entorno de Codex queda configurado del todo

El usuario puso `bash scripts/sincronizar.sh || true` como script de
mantenimiento del entorno. Con eso la configuración de Codex queda completa:
base `main`, acceso a internet activado con "Todos los métodos",
`codex_setup.sh` como script de configuración y `sincronizar.sh` como
mantenimiento (corre al reanudar un contenedor desde la caché, que es cuando
el clon está viejo).

No quedan pendientes del flujo Claude/Codex. Los que siguen en la lista son
todos anteriores, del sistema de reliquidación.

## 2026-09-07 — Claude — probado el caso real: una tarea abierta levanta trabajo posterior

Queda cerrado lo que la entrada de abajo dejaba pendiente. Se subió `ef4d28e`
a `main` **después** de que una tarea de Codex ya estuviera abierta y
sincronizada, y a esa misma tarea —sin abrirla de nuevo— se le pidió correr
`scripts/sincronizar.sh`. Trajo el commit en fast-forward, actualizó
`BITACORA.md` y avisó sola de ir a leer "Pendientes abiertos".

Ese era el problema con el que empezó todo esto: Codex no veía lo que Claude
subía sin abrir un chat nuevo. Con el entorno configurado (internet activado,
"Todos los métodos", `codex_setup.sh` como script de configuración) y la regla
3b de `REGLAS.md`, el circuito funciona de punta a punta.

Único pendiente del flujo: el script de mantenimiento del entorno.

## 2026-09-07 — Claude — `sincronizar.sh` confirmado dentro de una tarea de Codex

Corrida real, en una tarea nueva de Codex, después del arreglo del remoto:
`bash scripts/sincronizar.sh` terminó limpio y reportó que la copia ya estaba
al día con `main`. El circuito completo —acceso a internet activado, "Todos
los métodos", remoto configurado por el propio script, fetch anónimo por
HTTPS— funciona.

Queda por probar el caso que de verdad importa, que es el otro: que una tarea
**ya abierta** levante un commit subido *después* de que esa tarea arrancó.
Esta corrida no lo demuestra, porque no había nada nuevo que traer. Se prueba
subiendo algo a `main` desde otro lado y pidiéndole a la misma tarea de Codex
—sin abrirla de nuevo— que corra el script.

Sigue pendiente cambiar el script de mantenimiento del entorno, que hoy tiene
`codex_setup.sh` repetido.

## 2026-09-07 — Claude — el contenedor de Codex clona sin remoto `origin`

Primera corrida real de `scripts/sincronizar.sh` dentro de una tarea de Codex:
falló con `fatal: 'origin' does not appear to be a git repository`. **El
contenedor de Codex clona el repositorio sin dejar configurado ningún
remoto.** El script daba por sentado que `origin` existía; era un supuesto mío,
mal puesto, no un problema del entorno del usuario ni de la red.

Arreglo: `sincronizar.sh` y `codex_setup.sh` ahora resuelven el remoto solos.
Si `origin` no existe, lo configuran apuntando a
`https://github.com/italocarreran/Abby.git`. **El repositorio es público**
(verificado contra la API), así que el fetch anónimo por HTTPS alcanza y no
hacen falta tokens ni secretos en el entorno. Si algún día pasa a privado esto
deja de funcionar, y el mensaje de error del script lo dice como cuarta causa
posible.

`sincronizar.sh` además dejó de usar `origin/<rama>` y ahora trabaja contra
`FETCH_HEAD`: en un clon sin remoto esa referencia no existe, así que era la
segunda mitad del mismo bug.

**Probado de punta a punta** reproduciendo el caso: clon sin remoto, dos
commits atrás de `main`. El script avisa que no había remoto, lo configura,
lista los dos commits con autor y fecha, fusiona y muestra el `diff --stat`.

También se descubrió que la pantalla del entorno tiene **dos** campos de
script. El de mantenimiento corre al reanudar un contenedor desde la caché,
"después de comprobar la rama" — que es exactamente cuando el clon está viejo.
Ahí va `bash scripts/sincronizar.sh || true` (con el `|| true` para que un
conflicto quede en el log pero no tumbe el arranque). Queda anotado como
pendiente porque hoy tiene `codex_setup.sh` repetido.

Y el desplegable de métodos HTTP solo ofrece "Todos los métodos" o
"GET, HEAD y OPTIONS": hay que dejar el primero, porque git negocia el fetch
con un POST. `docs/FLUJO_CLAUDE_CODEX.md` se corrigió con todo esto, incluida
una nota honesta sobre el cartel de RIESGO ELEVADO de esa pantalla.

## 2026-09-07 — Claude — `main` queda como rama base y recibe los scripts

El usuario creó `main` desde la rama por defecto anterior (`claude/eso-uozpi4`),
la dejó como rama por defecto del repositorio y borró las once ramas viejas.
Verificado contra el remoto: `HEAD -> refs/heads/main`, y de las `codex/*` no
queda ninguna.

Se fusionó `claude/codex-shared-repo-config-3kho6z` en `main` en fast-forward
(sin commit de merge): `main` pasó de `9871577` a `8b0d15c` y ahora sí tiene
`scripts/` y `docs/FLUJO_CLAUDE_CODEX.md`. Esto era condición para que el
setup script del entorno de Codex funcione — Codex clona `main`, y hasta este
commit el archivo que le pide correr no existía ahí.

`bash scripts/sincronizar.sh` sin argumentos ya resuelve contra `main` como
corresponde; queda anulada la aclaración de la entrada anterior sobre tener
que pasarle la rama a mano. Ese pendiente sale de la lista.

Sigue abierto el paso 3 del documento: el internet access del entorno de
Codex. *(8b0d15c en `main`)*

## 2026-09-07 — Claude — flujo Claude/Codex: sincronización dentro de la tarea

Nada de código del sistema: es configuración de cómo trabajan los dos
asistentes sobre el repositorio.

**El problema que se resolvió.** Una tarea de Codex en la web clona el
repositorio una sola vez, al crearse, y ese clon no se actualiza nunca solo.
Todo lo que Claude o el usuario suban después es invisible para esa tarea, y
por eso el usuario tenía que abrir un chat nuevo cada vez que algo cambiaba.
No hay ninguna opción que haga que una tarea ya empezada se sincronice sola —
lo que sí se puede es que haga `fetch` en cada turno, y eso es lo que se
montó.

**Lo que se agregó:**

- `docs/FLUJO_CLAUDE_CODEX.md` — la configuración paso a paso, para el
  usuario: rama `main` por defecto, entorno de Codex, acceso a internet con
  `POST` habilitado, y el atajo de `@codex` sobre un PR (que arranca desde el
  HEAD actual del PR y por eso nunca queda viejo).
- `scripts/sincronizar.sh` — `fetch` de `main`, resumen de qué llegó y de
  quién, merge, y aviso de que conviene releer los pendientes.
- `scripts/verificar.sh` — `generar_interfaces.py --check` más los 11 tests
  del repositorio de una sola vez. `test_tema.py` se omite solo si no hay
  `tkinter` (contenedores headless); eso es del entorno, no del cambio.
- `scripts/codex_setup.sh` — setup script del entorno de Codex: instala
  `python3-tk`, le pone identidad de git al agente y corre la verificación.

**Lo que se cambió:** `REGLAS.md` punto 3 ahora manda correr
`scripts/sincronizar.sh` y suma el 3b (repetirlo en cada turno, no solo al
abrir la sesión) y el 3c (`main` es la base). El punto 1 avisa que los commits
de Codex pueden salir con el nombre del usuario y por eso llevan prefijo
`codex:`. El punto 5 apunta a `scripts/verificar.sh`. `AGENTS.md` §3 lista los
scripts y explica en un párrafo qué tiene que hacer Codex al empezar.

**Estado del repositorio al hacer esto:** no existe rama `main`; la rama por
defecto es `claude/eso-uozpi4` y hay once ramas `codex/*` y `claude/*`
colgando, **todas fusionadas** (verificado con `git rev-list --count` contra
la rama por defecto: cero commits adelante en las once). Nada que rescatar
antes de borrarlas.

**Qué queda pendiente:** los dos ítems nuevos de "Pendientes abiertos" — son
acciones en la interfaz de GitHub y de ChatGPT que ningún asistente puede
hacer por el usuario. Mientras `main` no exista, `scripts/sincronizar.sh` hay
que llamarlo con la rama explícita.

## 2026-09-07 — Claude — revisa los lectores Excel: sin correcciones

Revisión de `codex/continua-con-fases-pendientes-de-la-aplicacion` (rama nueva
porque el chat de Codex no vuelve a la anterior; el contenido sí sale de la
principal al día, `014ccf1`, y conserva los arreglos previos).

**Primera vez que no hubo nada que corregir.** Codex aplicó las tres lecciones
de las revisiones anteriores, sin que hiciera falta pedírselo de nuevo:

- Los `from revisor.lectores import ...` van **dentro de `try/except
  ImportError`** con `_morir_import()` y título propio. Verificado: sin el
  archivo, el Revisor aborta con el mensaje correcto.
- **No quedó ningún nombre definido y a la vez importado** (la trampa que se
  comió a `TOL_MTIME`). Lo comprobé por AST sobre todo el punto de entrada.
- **Las pruebas corren** desde cualquier carpeta, y los comentarios que explican
  el porqué viajaron con el código: 16 comentarios y 21 docstrings en 604
  líneas, contra el único comentario que traían los módulos de la etapa pasada.

**Verificación del comportamiento**, comparando función vieja contra nueva:

- Con `openpyxl` instalado en este entorno se ejercitó el **camino real** de
  `leer_columna_excel` sobre un libro de verdad: 13 casos (sin filtro, con
  filtro por tipo, filtro vacío, `fila_inicio` corrido, hoja por nombre / por
  nombre normalizado / por `#N` / inexistente, columna de texto, archivo
  ausente) — 0 diferencias en resultado, en logs y en el contenido del caché.
  El reúso del caché entre dos filtros distintos (SCMT y luego SCPC) da los
  mismos totales con una sola entrada, igual que antes.
- `col_letra`, `col_letra_a_num`, `_es_num`, `es_significativo`,
  `resolver_hoja`, `leer_celdas_rapido`, `leer_formulas_rapido`,
  `diagnosticar_celda`, `buscar_marcas_rapido` y `armar_tabla` sobre libros
  OOXML sintéticos: 0 diferencias.
- `CACHE_COLUMNAS` sigue siendo **el mismo objeto** a ambos lados del import, así
  que los dos `.clear()` del Revisor siguen vaciando el caché real. Era el
  riesgo principal de mover un diccionario mutable de módulo.
- Los **117 nombres de nivel superior** que el Revisor tenía en la principal
  siguen todos accesibles. `CENTRALES_EMBALSE`=27, `NODOS`=36,
  `VERIFICADORES`=14, tolerancias intactas.

**Ojo con el estado real de la Fase 4: no está terminada.** El propio plan la
marca como "en curso: puntos 1 y 2 completos, 3 iniciado". Faltan los lectores
MDB y los adaptadores ligados al cuadro de pago (resto del punto 3), los motores
V4…V17 (punto 4) y el traspaso/lanzamiento (punto 5).

**Pendiente, y ya es el más importante:** el Revisor no se ha corrido de punta a
punta en Windows desde que empezó a partirse. Los actualizadores sí se probaron.
Antes de tocar los motores V4…V17 conviene abrir el Revisor con un mes real y
pasar las verificaciones una vez.

---

## 2026-09-07 — ChatGPT — continúa la Fase 4 con los lectores Excel

Se extrajo a `Revisor_Relq/revisor/lectores.py` el primer bloque cohesivo de
adaptadores Excel del Revisor: suma de columnas con fallback openpyxl/xlwings,
escaneo de marcas y fórmulas, lectura y diagnóstico OOXML, resolución de hojas y
armado de tablas. `Revisor_Reliquidacion.py` conserva los nombres históricos
mediante imports guardados con `_morir_import()`; el módulo nuevo no importa la
ventana ni las reglas V4…V17.

La extracción conserva el cuerpo de las funciones sin simplificarlo y mantiene
el mismo `CACHE_COLUMNAS`. Se agregaron pruebas stdlib para conversiones de
columnas, nombres de hojas, casos significativos y el contrato histórico de
`armar_tabla`. Se regeneró `INTERFACES.md` y se actualizaron el mapa y el plan.

**Pendiente de la Fase 4:** completar los lectores MDB y los adaptadores Excel
que siguen ligados al cuadro de pago; después extraer motores V4…V17 y finalmente
traspaso/lanzamiento. Sigue pendiente la validación integral del Revisor en
Windows con Excel, Access y archivos reales.

## 2026-09-07 — Claude — revisa la Fase 4 (parcial) sobre la rama de Codex

La usuaria probó los actualizadores contra archivos reales y **funcionan**, y le
dio luz verde a Codex para la Fase 4. Revisión de
`codex/dividir-revisor_reliquidacion.py`. Los arreglos se hicieron **sobre esa
misma rama**, a pedido de la usuaria, para que le siga sirviendo en su chat de
Codex; recién después se fusionó.

**Buena noticia de proceso:** esta vez la rama salió de la principal al día
(`03e66bd`), no de un commit viejo. Eso solo ya evitó revivir lo corregido antes.

**Alcance:** solo los puntos 1 y 2 de los cinco de la Fase 4 — `revisor/estado.py`
y `revisor/archivos.py`. Codex frenó a propósito antes de los lectores
Excel/MDB, los motores V4…V17 y el traspaso. Coincido: es el corte correcto.

**Las extracciones son fieles**, verificado comparando vieja contra nueva:

- `revisor/archivos.py`: sobre un árbol de prueba con copias de Windows,
  temporales, tildes y diarios — `buscar_carpeta`, `resolver_carpeta`,
  `buscar_archivo`, `listar_diarios`, `mtime`, `tamano`, `fmt_fecha`,
  `fmt_monto`, `iguales_mtime`: 0 diferencias. El caché encendido da además los
  **mismos contadores** de scans/hits que antes.
- `revisor/estado.py`: 32 comprobaciones sobre `Estado` y `CacheValores`,
  incluidos JSON roto, `aamm` vacío, firma cambiada y las tres invalidaciones
  del caché (mtime, tamaño, valor no numérico): 0 diferencias.
- El Revisor sigue exponiendo los 31 nombres que declara `MAPA.md`, con
  `CENTRALES_EMBALSE`=27, `NODOS`=36, `VERIFICADORES`=14 y las tolerancias
  intactas.

**Cuatro correcciones, ninguna de comportamiento:**

1. **`from revisor.archivos import ...` y `from revisor.estado import ...` sin
   guarda.** Tercera vez que aparece el mismo agujero, ahora con el paquete
   nuevo. Quedaron dentro de `try/except ImportError` con `_morir_import()`, que
   además acepta título propio para decir que lo que falta es `revisor/` y no
   `__comun__/`. Verificado: sin la carpeta, el Revisor aborta con `rc=1` y el
   mensaje correcto en vez de morir callado bajo `pythonw`.

2. **`TOL_MTIME` quedó definido dos veces y la buena se pisaba.** Estaba en el
   bloque de tolerancias del Revisor (línea 141, con su comentario) y otra vez
   en `revisor/archivos.py`, y el `import` de más abajo pisaba la primera. Mismo
   valor, así que hoy no cambiaba nada — pero editar la línea 141 no habría
   hecho **nada**, en silencio. Es la trampa de `CENTRALES_EMBALSE` otra vez. Se
   dejó una sola definición, en `revisor/archivos.py`, y un comentario en el
   bloque de tolerancias que dice dónde vive.

3. **Las pruebas nuevas no se podían correr.** Ni `python revisor/test_x.py`, ni
   `python -m revisor.test_x`, ni desde la raíz: las tres fallaban con
   `ModuleNotFoundError`. Ahora resuelven su propio `sys.path` como las de
   `__comun__` y corren desde cualquier carpeta. Quedó documentado en AGENTS.md.

4. **Se perdieron los comentarios que explican el porqué.** Los dos módulos
   nuevos absorbieron ~380 líneas y entre ambos tenían **un** comentario. Se
   restauraron los que importan: el porqué del caché de directorios (68
   recorridos, los viajes de red en la T:, por qué está apagado por omisión), el
   docstring de `vigente()` sobre por qué NO se mira la firma, y sobre todo el
   `OJO: aca la comparacion es EXACTA, sin la tolerancia de TOL_MTIME` de
   `CacheValores.obtener` — que es justamente lo que evita que alguien
   "unifique" las dos comparaciones de fecha y reviva un bug de valor viejo.

**Verificado** (Linux, con `tkinter` y las libs de Windows simuladas): el Revisor
importa y expone todo; aborta con el mensaje correcto sin `revisor/` y sin
`__comun__/`; las 10 suites pasan; `generar_interfaces.py --check` al día.

**Pendiente:** los tres bloques que faltan de la Fase 4 (lectores Excel/MDB,
motores V4…V17, traspaso y lanzamiento). Y aunque los actualizadores ya se
probaron de verdad, **el Revisor mismo no se corrió de punta a punta en Windows
después de esta división** — es lo próximo antes de seguir partiéndolo.

---

## 2026-09-07 — ChatGPT — Fase 4 parcial: estado y búsqueda del Revisor

Se trabajó desde `03e66bd`, punta de `claude/eso-uozpi4` y fuente de verdad
posterior a la revisión de las Fases 1–3. La rama local no tenía remoto
configurado: `git pull` no pudo determinar upstream y tampoco existe un destino
de push en este clon.

La Fase 4 se inició con dos extracciones completas y coherentes.
`Revisor_Relq/revisor/estado.py` contiene el estado mensual y el caché persistente
de valores; recibe explícitamente la resolución de rutas, escritura atómica,
firma y metadatos de archivos. `revisor/archivos.py` contiene búsqueda tolerante,
filtro de copias, metadatos y el caché de directorios acotado al context manager.
`Revisor_Reliquidacion.py` sigue siendo el único punto de entrada y conserva todos
los nombres históricos por imports o wrappers breves. No cambiaron `VALORES`,
`VERIFICADORES`, `NODOS`, `ACTUALIZADORES`, `CLAVES_TRASPASO`, los IDs V4…V17 ni
sus dependencias. El sobre del traspaso y el código que lanza procesos no se
tocaron.

Se agregaron pruebas stdlib para persistencia, JSON roto, firmas, invalidación del
caché por huella/fecha/tamaño, búsqueda normalizada, preferencia por originales y
encendido/apagado del caché de directorios. La comparación AST contra el commit
base confirmó que no falta ningún nombre público y que las siete constantes
estructurales permanecen idénticas. Los módulos importan aislados sin tkinter y
no importan el punto de entrada, por lo que la dirección de dependencias no forma
ciclos.

**Fase 4 todavía incompleta:** quedan lectores Excel/MDB, motores V4…V17 y
traspaso/lanzamiento. Claude debe revisar en Windows: arranque real de la ventana;
árbol completo de un mes; V4…V17 con caso real; persistencia y recuperación de
estado/caché; OOXML y fallback Excel; Access; lanzamiento de un actualizador;
igualdad de `_traspaso_actualizador.json`; y conservación de configuración al
cerrar y reabrir. La usuaria confirmó en esta solicitud que ya probó los
actualizadores reales de las fases anteriores, por lo que se retiró ese pendiente
abierto antiguo.

## 2026-09-07 — Claude — revisa las Fases 2 y 3 y corrige tres regresiones

Revisión de `codex/reorganizar-codigo-para-modularidad-hljpmw` y fusión a
`claude/eso-uozpi4`. **Ojo con cómo venía la rama:** un único commit cortado de
`5a5ba6d`, o sea de ANTES del arreglo de la Fase 1. Rehace la Fase 1 desde cero
y le agrega las Fases 2 y 3. No es una rama encima de la principal.

**El trabajo está bien hecho y las extracciones son fieles**, verificado
comparando función vieja contra nueva, entrada por entrada:

- `texto.py`: 425 comparaciones contra las 8 variantes originales de
  `normalizar`, 0 diferencias. Conserva a propósito las tres variantes que
  diferían en `None` (`suave`, `suave_textual`, `suave_requerido`) en vez de
  unificarlas, que era lo correcto.
- `excel_xml.py`: 18 libros OOXML sintéticos (cadenas compartidas, `inlineStr`,
  booleanos, errores, entidades dec/hex, celdas vacías, hoja por nombre y por
  `#N`, target con `/` y con `xl/`, un libro de 9.000 filas que cruza el corte
  de 1 MiB, no-zip). 0 diferencias contra las dos copias originales.
- `archivos.py`: 23 nombres contra los dos `es_copia` originales, 0 diferencias.
- `comparadores.py`: `ColaTk`, `respaldar`, `subcarpeta`, `buscar_mdb`,
  `es_hoja_propia` y el caché son traducciones fieles.

**Tres regresiones corregidas.**

1. **Los imports sin guarda, otra vez** (el mismo agujero de la Fase 1, que
   volvió porque la rama se cortó antes del arreglo). Ahora eran **39 imports
   sueltos en 12 archivos**, más que antes, porque las Fases 2 y 3 agregaron
   cuatro módulos compartidos. Se consolidaron: cada ejecutable tiene un único
   `try/except ImportError` detrás de su `_morir()`.

2. **`color_de` perdió dos de sus cinco casos.** El original mapeaba
   `desactualizado` → amarillo y cualquier estado desconocido → gris. La versión
   extraída dejaba solo `ok`/`pendiente` y mandaba **todo lo demás a rojo**.
   `estado_etapa()` sigue devolviendo `desactualizado`, así que un mes que solo
   había que rehacer se veía igual que uno al que le falta el archivo. En
   `Comparador_Tabulado` es color de fondo del chip: bien visible.

3. **`mes_incluido`/`fijar_incluido` cambiaron el formato en disco.** El
   original guarda la marca dentro del registro del mes
   (`estado[aamm]["incluir"]`); la versión extraída la movió a un diccionario
   aparte (`estado["_incluidos"][aamm]`). Nada falla: simplemente **deja de
   encontrar las exclusiones ya escritas** en los `estado.json` del equipo de la
   usuaria, y todos los meses que ella había sacado a mano vuelven a entrar al
   consolidado anual sin ningún aviso. Se volvió al formato de siempre, que
   además no necesita migrar nada.

Las 2 y 3 pasaban las pruebas de la rama porque las pruebas se escribieron
contra el código nuevo, no contra el comportamiento viejo. Se agregaron casos
que fijan los cinco estados de color y el formato en disco de la inclusión.

**También:** `excel_xml.py` no anotaba cuál de las dos variantes de
`col_letra_a_num` se quedó (la defensiva del Revisor, no la de Data Access).
AGENTS.md lo pide explícitamente al juntar copias; queda en su docstring.

**Verificado** (Linux, sin Windows ni Excel, con `tkinter` y las libs simuladas):
los 10 ejecutables abortan con `rc=1` y el mensaje de `__comun__` si falta la
carpeta, y con ella pasan el import; los envoltorios de config y traspaso de los
10 siguen cumpliendo las tres reglas del archivo compartido; las 8 suites de
`__comun__` pasan; `generar_interfaces.py --check` al día.

**Pendiente:** la Fase 4 no se empezó, y coincido con ChatGPT en que es la
frontera de riesgo alto. Sigue sin probarse ningún actualizador real de punta a
punta en Windows — y ahora hay bastante más código compartido en juego.

---

## 2026-09-07 — ChatGPT — Fases 2 y 3; se detiene antes del riesgo alto

Después de que Claude revisó y fusionó la Fase 1, la usuaria pidió continuar
hasta llegar a las fases de riesgo alto. Se completaron la Fase 2 (lectura y
utilidades comunes) y la Fase 3 (infraestructura de comparadores), y se dejó sin
iniciar la Fase 4 porque divide internamente las 6.000+ líneas del Revisor y es
la frontera de riesgo alto.

Fase 2: `__comun__/excel_xml.py` reúne el lector OOXML que era idéntico en el
Revisor y Data Access; ambos conservan aliases y el mismo fallback. Se crearon
`texto.py` y `archivos.py` para las normalizaciones y los filtros de temporales
y copias. Antes de migrar se inventariaron las variantes: no todas trataban
igual `None` y `0`, por lo que el módulo conserva contratos explícitos
(`suave`, `suave_textual`, `suave_requerido`, `clave*`) y se comprobó su paridad
contra las funciones de `HEAD`.

Fase 3: `__comun__/comparadores.py` comparte por composición `ColaTk`,
caché de una consulta `scandir`, estado e inclusión
mensual, firmas, respaldos, hojas ajenas y utilidades de rutas/Access. Los dos
comparadores mantienen wrappers históricos y siguen separados en todo lo de
dominio (Access vs. Tabulado, SQL, vistas, columnas y Excel). La prueba de la
cola usa un hilo real y confirma que los widgets no cambian hasta bombear desde
el hilo de UI.

Se añadieron suites stdlib para los cuatro módulos nuevos, se regeneraron mapa
e interfaces y se compiló todo el árbol. También se comparó por AST/paridad el
comportamiento histórico. No se ejecutaron Excel, Access, SQL Server ni ventanas
reales por no ser Windows. Claude debe revisar especialmente el OOXML sintético,
el puente de cola y que los respaldos sigan en su ruta actual antes de fusionar.

**Pendiente:** la Fase 4 está descrita en
`docs/PLAN_MODULARIZACION_TOKENS.md`, marcada expresamente como riesgo alto y no
se tocó.

---

## 2026-09-07 — Claude — revisa la Fase 1 de ChatGPT y le tapa un agujero real

Revisión de `codex/reorganizar-codigo-para-modularidad` (1 commit sobre la rama
principal) y fusión a `claude/eso-uozpi4`.

**El trabajo de ChatGPT está bien hecho.** `__comun__/traspaso.py` es una
traducción fiel de las nueve copias de `leer_traspaso()`, los envoltorios
conservan los nombres históricos y ningún punto de llamada cambió. La mejora de
`config.leer()` (bloque propio que no es dict → `{}`) reemplaza correctamente
al `or {}` que tenían los comparadores.

**Lo que había que corregir — una regresión real, no cosmética.** El commit
agregó 20 `from __comun__ import ...` **sin `try/except ImportError`**, en 8
scripts que hasta ese commit **no dependían de `__comun__` para nada** (tenían
su propia copia del config). En `Actualiza_Energia.py` y `Actualiza_Access_P9.py`
el import quedó además *arriba* de la definición de `_morir()`, así que ni
siquiera se podía guardar sin reordenar.

Por qué importa: el Revisor lanza estos scripts con `pythonw`, sin consola. Un
`ImportError` suelto los mata en silencio — la usuaria solo ve que la ventana
nunca aparece. Es exactamente el caso que ya defendían con `_morir()` el
Revisor, los dos comparadores y `Actualiza_SC_CO.py`, y el que AGENTS.md nombra
como "baja el repositorio completo, no los .py sueltos".

Arreglo, con el idioma que ya usaba el repo:

- 4 archivos que ya tenían guarda (Revisor, los 2 comparadores,
  `Actualiza_SC_CO.py`): los imports nuevos se movieron adentro del `try`.
- 2 que tenían `_morir()` pero definido después (`Actualiza_Energia.py`,
  `Actualiza_Access_P9.py`): el bloque se bajó detrás de `_morir()`, ya con guarda.
- 6 que no tenían `_morir()` (`Actualiza_datos`, `Actualiza_Cuadro0`,
  `Actualiza_Data_Access`, `Carga_Retiros`, `Prorratear`, `ActualizaRemplazos`):
  se les agregó. **`_morir()` no puede vivir en `__comun__/`**: es justamente lo
  que avisa cuando `__comun__/` es lo que falta. Queda anotado en MAPA.md y como
  trampa en AGENTS.md.

**Verificado en este entorno** (Linux, sin Windows ni Excel), con `tkinter` y las
libs de Windows simuladas:

- Los 10 ejecutables abortan con `rc=1` y el mensaje de `__comun__` cuando se
  corren desde una copia sin la carpeta. Los 2 comparadores no llegan a ese punto
  acá porque antes cortan por librerías faltantes (`pyarrow`, `duckdb`…); en
  esos dos el arreglo fue mover una línea adentro de un `try` que ya existía.
- Con `__comun__` presente, los 10 pasan el import y llegan a construir la
  ventana.
- Los envoltorios de los 10, contra un `config.json` temporal: round-trip
  guardar/leer, un config roto **no se pisa**, y las claves de otro equipo
  **no se borran**. Más `leer_traspaso()`: sin argumento → modo manual; con un
  JSON válido → lo lee.
- `CONFIG_PATH` idéntico a la rama principal en los 12. `ORIGEN` y
  `VERSION_ACTUAL` = 1, y el Revisor sigue escribiendo versión 1.
- `test_config.py` (14), `test_salidas.py` (13) y `test_traspaso.py` (8): OK.
  `test_tema.py` no corre acá porque este contenedor no tiene `tkinter` — ya
  pasaba antes de este cambio.
- `generar_interfaces.py --check`: al día.

**Pendiente:** nada de este cambio. Sigue en pie que nadie probó un actualizador
real de punta a punta en Windows. La Fase 2 (`excel_xml.py`, `texto.py`,
`archivos.py`) está planificada en `docs/PLAN_MODULARIZACION_TOKENS.md` y **no**
se empezó.

---

## 2026-09-07 — ChatGPT — Fase 1 de modularización: config y traspaso

La usuaria pidió ejecutar la primera fase del plan orientado a reducir tokens y
dejarle a Claude el plan completo para revisar y fusionar. Los doce programas
que manejan configuración delegan ahora en `__comun__/config.py`: Revisor, nueve
actualizadores y dos comparadores. Se conservaron los nombres históricos como
aliases o wrappers para no cambiar puntos de llamada, y
`ActualizaRemplazos.py` sigue apuntando a su JSON propio. La migración corrigió
además el riesgo pendiente de ese script: ya no puede pisar un config roto con
`{}` y ahora escribe mediante `.tmp` + `os.replace`, igual que los demás.

Se agregó `__comun__/traspaso.py` como contrato único del argumento opcional.
Centraliza origen, versión, lectura tolerante y normalización de `rutas`; los
nueve actualizadores mantienen `leer_traspaso()` como wrapper y el Revisor usa
las mismas constantes al producir el JSON. Sin argumento o ante un JSON
inválido se conserva exactamente la vía manual. También se corrigieron los tres
comentarios viejos que todavía ubicaban ese JSON bajo `Salidas/AAMM`.

El plan y los límites de las fases 2 a 4 quedaron en
`docs/PLAN_MODULARIZACION_TOKENS.md`: OOXML/texto/archivos por piezas;
infraestructura compartida de comparadores mediante composición; y, al final,
división interna del Revisor sin mover su entry point. Se agregaron 8 pruebas
del traspaso y una prueba defensiva nueva a config (14 en total). Verificado con
los tests de los cuatro módulos comunes, compilación de todos los Python,
chequeo del generador, comparación de los doce `CONFIG_PATH` contra `HEAD` y
auditoría AST de wrappers/consumidores. No se pudo ejecutar Excel, Access, SQL
Server ni las ventanas reales porque este entorno no es Windows.

**Pendiente para Claude:** revisar esta rama contra la lista específica del
plan, ejecutar al menos un actualizador real en Windows cuando haya archivos de
trabajo y fusionar si coincide. Las fases 2, 3 y 4 no se empezaron.

## 2026-09-03 — Claude — plan para sacar `comun/` y todos los `.json` de `Revisor_Relq/`

La usuaria vio el warning del editor sobre `sys.path.insert` + `from comun
import ...` en los comparadores y pidió una reorganización más de fondo, no
un parche: que `comun/` y absolutamente todos los `.json` del repo queden
afuera de `Revisor_Relq/`, en dos carpetas hermanas nuevas —
`__comun__/` (código, se sube al repo) y `__config__/` (datos, NO se sube —
la usuaria la arma y ordena a mano, con la misma estructura `AAAA/MM Mes`
que ya usa `00_Salidas`). Además: `00_Salidas/` debe terminar conteniendo
solo resultados (`.xlsx`), nada de estado ni caché.

Antes de escribir el plan, hubo una pregunta pendiente: los comparadores
guardan, junto a `estado.json`/`rutas.json`, unas carpetas `parquet/` y
`vistas/` (datos de los `.mdb` ya leídos, para no releer Access cada vez) —
ni son `.json` ni son el resultado final. Se le preguntó a la usuaria dónde
debían ir; eligió la opción recomendada: junto con el estado, en
`__config__/AAAA/_comparador/` (y `_comparador_tabulado/`). Con esa
respuesta, `00_Salidas/AAAA/` queda con el Excel anual directo (sin la
subcarpeta `_comparador*`) y `00_Salidas/AAAA/MM Mes/` con los dos Excel
mensuales — nada más.

Se auditó el repo entero para inventariar cada `.json` real que existe hoy
(no de memoria): el `config.json` compartido de `Revisor_Relq/` (Revisor + 8
actualizadores + 2 comparadores), el `config.json` propio de
`Reemplazos REUC/Auxiliares/` (de `ActualizaRemplazos.py`, que además guarda
ahí archivos `.xlsx` descargados que NO se mudan — solo su `.json`), y por
mes/año: `_revisor_verificaciones.json`, `_revisor_cache_valores.json`,
`_traspaso_actualizador.json` (que además leen los dos comparadores, no solo
el Revisor — hay que mover las dos puntas juntas o uno lee donde el otro ya
no escribe) y, por comparador, `estado.json`/`rutas.json` (Tabulado además
lee el `rutas.json` de Etapas, para reusar sus rutas de `.mdb`). De paso
apareció una carpeta huérfana en la raíz del repo, `comun/README.md` —un
placeholder de una etapa anterior del proyecto, nunca usado por ningún
script real (el `comun/` que sí se usa siempre vivió dentro de
`Revisor_Relq/`) — el plan la borra en la Tarea 1 para no terminar con dos
`comun` en el árbol.

El plan completo, en 3 tareas (mover `comun/` → `__comun__/` y generalizar
`_hallar_revisor()` a `_hallar_workroot()` — busca la carpeta que tiene
`__comun__` al lado, en vez de buscar un archivo puntual del Revisor, porque
después de este cambio los comparadores ya no necesitan encontrar
`Revisor_Relq/` para nada; mover todos los `.json` a `__config__/` y
reordenar `00_Salidas/`; actualizar toda la documentación), quedó en
`docs/PLAN_reorganizacion_config.md`, con instrucciones línea por línea para
cada uno de los 12 scripts que tocan config y con su propia verificación por
tarea. Se lo confirma como listo para que lo aplique Codex.

Nada de código se tocó en esta sesión — es 100% planificación, como pidió la
usuaria ("planifica y piensa, el ejecuta"). El plan mismo dice que se borra
una vez aplicado y verificado (mismo ciclo que tuvo
`docs/PLAN_comparadores.md`).

**Pendiente**: aplicar el plan (Codex), verificar cada tarea con pruebas de
comportamiento reales (no solo lectura de código — ya fue el estándar en
`PLAN_comparadores.md` y encontró bugs reales que una lectura no habría
visto), y avisarle a la usuaria que después de aplicado tiene que armar
`__config__/` a mano con sus archivos existentes.

---

## 2026-09-03 — Claude — el warning del editor: causa real y arreglo verificado

La usuaria avisó que el subrayado amarillo seguía apareciendo **después** de
toda la reorganización. Tenía razón, y la suposición con la que se cerró la
sesión anterior (que mover `comun/` a `__comun__/` como carpeta de primer
nivel iba a hacerlo desaparecer solo) era incorrecta.

Se reprodujo de verdad, no por lectura: se instaló `pyright` (el mismo motor
que usa Pylance dentro de VS Code) y se corrió sobre una copia del árbol real.
El diagnóstico exacto es
`error | reportMissingImports | Import "__comun__" could not be resolved`, y
aparece **según qué carpeta esté abierta en el editor**:

- raíz = la carpeta de trabajo (la que contiene `__comun__/`) → 0 errores.
- raíz = `Comparadores/` → 2 errores (las dos líneas del import).
- raíz = `Revisor_Relq/` → 1 error.
- raíz = `Revisor_Relq/actualizadores/` → 1 error.

O sea: el analizador resuelve los imports mirando la carpeta abierta, no
ejecutando el `sys.path.insert` — que es justamente lo que le dice a Python
dónde está `__comun__/` en tiempo de ejecución. Por eso el programa corre
perfecto y el editor igual protesta. Y por eso mover la carpeta **no podía**
arreglarlo: al quedar `__comun__/` un nivel ARRIBA de los programas, abrir la
carpeta de un programa nunca la va a ver.

El arreglo es de editor, no de código: un `.vscode/settings.json` con
`python.analysis.extraPaths` (y su gemelo `python.autoComplete.extraPaths`)
en `[".", "..", "../.."]`, repetido en las tres carpetas —raíz,
`Revisor_Relq/` y `Comparadores/`— porque VS Code solo lee el de la carpeta
que abriste. Las tres rutas cubren los cuatro casos de arriba, incluido abrir
`actualizadores/` suelta. Verificado con el mismo pyright: los cuatro casos
pasan a **0 errores**. Lo que no se pudo probar acá es VS Code en sí (no hay
entorno gráfico): se verificó el motor y los valores de ruta, que es donde
estaba el problema; `python.analysis.extraPaths` es exactamente cómo Pylance
consume esa misma opción.

Se agregó la fila correspondiente en `AGENTS.md` → "Trampas conocidas",
porque el reflejo equivocado acá es caro: alguien "arregla" el subrayado
copiando `__comun__/` adentro de cada programa o volviendo a duplicar el
código, y se pierde la única regla que sostiene toda la reorganización (lo
compartido vive en un solo lugar). En `README.md` quedó la explicación en
criollo, con las tres causas por las que podría seguir apareciendo: que
`__comun__/` no esté todavía bajada al lado de los programas, que la carpeta
abierta no traiga su `.vscode/`, o que VS Code tenga el análisis viejo en
memoria (*Developer: Reload Window*).

---

## 2026-09-03 — ChatGPT — separa configuración, código común y resultados

A pedido del usuario se establece una separación estricta en la raíz: `__comun__/`
contiene el paquete Python compartido; `__config__/` concentra absolutamente todos
los JSON y datos intermedios; `00_Salidas/` queda solo para resultados Excel. El
config compartido pasa a `__config__/config.json` y el propio de REUC a
`__config__/reemplazos_reuc.json`. Estado, caché y traspaso mensual del Revisor
replican `AAAA/MM Mes` bajo `__config__/`; estado, rutas, parquet y vistas de los
comparadores pasan a `__config__/AAAA/_comparador*`. Los Excel mensuales, anuales y
sus respaldos permanecen en `00_Salidas/`.

`Revisor_Relq/comun/` se movió a `__comun__/`, como hermana de los programas. Los
imports ahora nombran directamente `__comun__`, eliminando la referencia amarilla
a un paquete `comun` que el analizador no encontraba en esa ubicación. Las rutas
crean `__config__/` automáticamente al primer guardado; la carpeta completa está
ignorada por Git. Se actualizaron mapa, interfaces, instrucciones y documentación.
No se migran archivos locales viejos automáticamente: el usuario indicó que va a
ordenarlos. Verificado con compilación de todos los Python, 13+13+3 pruebas de los
módulos comunes, chequeo del generador y una prueba estructural de las doce rutas
de configuración. No queda pendiente de esta reorganización.

---

## 2026-09-03 — Claude — verifica la reorganización de ChatGPT y corrige dos regresiones reales

ChatGPT aplicó la reorganización sin ver `docs/PLAN_reorganizacion_config.md`
(bifurcó de un commit anterior a que ese plan se subiera) — hizo su propia
lectura del pedido de la usuaria, en una rama aparte
(`codex/reorganizar-estructura-de-carpetas-y-archivos`). Se revisó línea por
línea contra los dos, el plan y el pedido original, y se corrió cada prueba
de comportamiento real posible en este entorno (no solo lectura de código),
antes de mergear a esta rama.

**Lo que estaba bien, verificado de verdad:** los 10 `CONFIG_PATH` del
repo apuntan sin excepción dentro de `__config__/`; el traspaso
(`_traspaso_actualizador.json`) lo escribe el Revisor y lo leen los dos
comparadores en la MISMA ruta — se armó un árbol de prueba real (con
`__comun__`, `Revisor_Relq` y `Comparadores` de verdad, tkinter simulado) y
se confirmó que las tres puntas coinciden; las 26 pruebas de
`__comun__/test_salidas.py` + `test_config.py` pasan igual que antes de
mover el módulo; `.gitignore` ignora `__config__/` entera sin tocar
`__comun__/`; `__pycache__/` sigue cubierto (no hacía falta agregar nada,
la usuaria solo avisó "por si"); se borró la carpeta huérfana `comun/` de
la raíz; `Reemplazos REUC/Auxiliares/` conserva sus `.xlsx` descargados,
solo se le fue el `.json`; `generar_interfaces.py --check` queda limpio.

**Dos regresiones reales, corregidas en esta rama antes de avisar:**

1. Los dos comparadores seguían guardando el Excel **anual** dentro de una
   subcarpeta `_comparador`/`_comparador_tabulado` bajo `00_Salidas/AAAA/`
   (`00_Salidas/2024/_comparador/Comparacion_Etapas_2024.xlsx`), en vez de
   directo bajo el año
   (`00_Salidas/2024/Comparacion_Etapas_2024.xlsx`) — que es exactamente lo
   que la usuaria confirmó en la pregunta que se le hizo antes de escribir
   el plan (`00_Salidas` "solo resultados, nada más", con el Excel anual
   como único archivo del año). Confirmado con una prueba real (se importó
   el módulo con tkinter simulado y se comparó la ruta devuelta contra la
   esperada) antes y después de la corrección. `dir_resultados_anuales()`
   ahora arma `00_Salidas/AAAA` directo, sin la subcarpeta — en los dos
   comparadores.
2. `Actualiza_SC_CO.py` tenía, antes de esta reorganización, un
   `try/except ImportError` alrededor de `from comun import config` que
   mostraba un diálogo de error y salía limpio si faltaba la carpeta —
   ChatGPT lo sacó al cambiar a `from __comun__ import config`, dejando un
   `ModuleNotFoundError` sin capturar (invisible si el script corre con
   `pythonw`, sin consola). Se repuso el mismo patrón `_morir()` que ya usa
   el resto del archivo. Se agregó el mismo resguardo, que antes no hacía
   falta porque `comun/` viajaba SIEMPRE adentro de `Revisor_Relq/`, en dos
   lugares nuevos donde ahora aplica el mismo riesgo (bajar el repo a medias
   ahora puede dejar afuera `__comun__/`, que es una carpeta hermana
   aparte): `Revisor_Reliquidacion.py` (nunca tuvo un `_morir`, se le agregó
   uno chico) y el `from __comun__ import salidas/tema` de los dos
   comparadores (que ya tenían `_morir` para el caso de no encontrar
   `Revisor_Relq/`, pero no para este import).

Quedó una decisión sin confirmar con la usuaria, anotada en "Pendientes
abiertos": los respaldos de los últimos 5 Excel anuales de los comparadores
quedaron en `00_Salidas/AAAA/respaldos/` (compartida entre los dos, sin
colisión de nombres) — no es un resultado final ni es state/caché, se dejó
ahí a criterio propio, pero podría discutirse si no debería ir a
`__config__/AAAA/_comparador*/` también.

`docs/PLAN_reorganizacion_config.md` se borra en este mismo commit: ya está
aplicado y verificado (mismo ciclo que tuvo `docs/PLAN_comparadores.md`).

---

## 2026-09-03 — Claude — verifica la Tarea 4 con `tkinter` real y agrega la trampa que faltaba

Codex hizo un trabajo sólido: `comun/tema.py` resuelve bien los tres problemas
que el plan anticipaba (widgets `tk` clásicos no siguen a ttk, colores de
estado no legibles en oscuro, `SystemButtonFace`), con un truco elegante que
el plan no pedía — `pintar_tk` reconoce el color **semántico** que ya tiene un
widget (comparándolo contra los valores de las dos paletas) y lo remapea, en
vez de que cada punto de llamada tenga que saber en qué paleta está.

Este entorno no tenía `tkinter` de verdad (se venía simulando todo con
`MagicMock`). Se instaló (`apt-get install python3-tk`), pero resultó ser para
otro intérprete — el `python3` real de esta sesión es un `pyenv` aparte, así
que `test_tema.py` de Codex no se pudo correr tal cual. Se verificó igual con
`tkinter` real, simulando solo `ttk.Style` (el mismo patrón que ya usa su
propio test): la paleta clara da carácter por carácter los 4 colores
históricos; un modo desconocido cae a claro; `aplicar()` con un `root` real
configura `bg`; y lo más delicado — un botón pintado en rojo-claro migra a
rojo-oscuro al llamar `pintar_tk` con la paleta oscura, confirmando que el
remapeo semántico funciona con el objeto real, no con un doble. Se forzó
además una excepción real en `_tema.aplicar()` (`RuntimeError`) y se confirmó
que `_aplicar_tema()` no la deja propagar: devuelve `False` y avisa en el log,
sin tumbar la ventana — la regla del plan de "nunca impedir que arranque".

Confirmado por diff de contenido (no de números de línea) que **ningún**
control existente cambió de posición en las dos ventanas — la única línea
`.pack()` nueva es la casilla "Tema oscuro". Y que no se agregó ninguna
librería externa de temas.

**Antes de aceptar el borrado de `docs/PLAN_comparadores.md`** (738 líneas,
que el propio plan decía que se eliminaría al terminar) se revisó que lo que
valía la pena sobreviviera en otro lado. Encontrado un hueco: la razón central
de todo el trabajo — que el Revisor y los dos comparadores arman la carpeta
del mes de una única manera, vía `comun/salidas.py`, para no repetir el error
de `CENTRALES_EMBALSE` — nunca había llegado a la tabla "Trampas conocidas" de
`AGENTS.md`, que es el lugar pensado justo para eso. Agregada esa fila antes
de dar la Tarea 4 por cerrada.

Con esto, **las cuatro tareas del plan quedan aplicadas y verificadas**, cada
una con pruebas de comportamiento reales y no solo lectura de código.

## 2026-09-03 — ChatGPT — Tarea 4: tema oscuro experimental

Se agregó `comun/tema.py`, sin dependencias externas, con paletas clara y
oscura, estilos `ttk` y pintado recursivo de los widgets `tk` clásicos. La
paleta oscura incluye colores de estado legibles; la clara conserva exactamente
los cuatro colores de estado y el color de enlace que usaban los comparadores.

Los dos comparadores tienen ahora una casilla «Tema oscuro». Lee y escribe la
clave compartida `tema` en `config.json`, aplica el cambio en vivo y arranca en
claro cuando la clave no existe. Toda aplicación está protegida: si falla, se
registra el problema y la herramienta conserva el aspecto anterior. El Revisor
y los actualizadores no se tocaron, porque el plan define este cambio como un
piloto limitado a los dos comparadores.

`test_tema.py` verifica sin pantalla el modo por omisión, los colores históricos
y `aplicar()` con un root/estilo simulados. Pasaron también las pruebas de
configuración y salidas, sintaxis de ambos comparadores y el generador. Falta la
prueba visual real en Windows, que queda en Pendientes abiertos. Con las cuatro
tareas completas se elimina el documento temporal `docs/PLAN_comparadores.md`.

## 2026-09-03 — Claude — verifica la Tarea 3 con pruebas de comportamiento reales

Codex resolvió los tres bugs de mayor gravedad del plan (3.1, 3.2, 3.3) y dejó
anotado el 3.4 (duplicación) sin migrar, tal como pedía la sección 9. No se
tocó nada de `Revisor_Relq/`, dentro de alcance.

**No alcanzaba con leer el código** — esta vez se instaló `pandas` (no estaba
en este entorno) y se corrieron pruebas de comportamiento reales, con hilos de
verdad y datos sintéticos, no solo AST:

- **3.1 (threading):** una `App` mínima simulada, con `queue.Queue` y
  `_bombear_cola` reales. Un hilo real llamó a `log()` 50 veces; confirmado que
  la cola acumuló los 50 mensajes **sin tocar el widget ni una vez** desde el
  hilo, y que `self.txt.insert` recién se llamó al vaciar la cola desde el
  hilo principal.
- **3.2 (Excel):** confirmado por código que ningún bloque de continuación de
  hoja tiene `break`/`continue` que descarte filas — los dos motores
  (`xlsxwriter` 0-indexado y `openpyxl` 1-indexado) abren hoja nueva y siguen,
  cada uno respetando su propio encabezado (`fila = 1` vs `fila = 2`, correcto
  por la diferencia de indexado, no un descuido).
- **3.3 (`hora_mes`):** `acumulado_por_dia` con `pandas` de verdad, tres casos:
  mes completo de 744 horas (no avisa), mes con el día 15 completo faltante
  (avisa "dias no contiguos (faltan: [15])"), y un día con la hora inicial
  corrida a 2 (avisa "hora inicial distinta de 1"). Los tres dieron exactamente
  lo esperado.

Verificación mecánica también: sintaxis de los 14 `.py`, 13+13 pruebas de
`comun/`, `generar_interfaces.py --check`.

Con esto, **Tareas 1, 2 y 3 del plan quedan aplicadas y verificadas**. Falta
solo la Tarea 4 (tema oscuro, experimental) para poder borrar
`docs/PLAN_comparadores.md`.

## 2026-09-03 — ChatGPT — Tarea 3: corrige los bugs de los comparadores

Los dos comparadores pasan ahora todos los cambios de log, estado, progreso y
repintado por una `queue.Queue` que vacía exclusivamente el hilo principal de
tkinter; los workers ya no tocan widgets ni llaman `after()`.

`Comparador_Tabulado` continúa en hojas `AAMM_2`, `AAMM_3`, etc. al alcanzar el
límite de Excel, tanto al crear un libro como al preservar hojas ajenas, sin
descartar filas. También valida antes de calcular `hora_mes` que días y horas
sean contiguos, que cada día empiece en 1 y que el total mensual sea uno de los
largos admisibles con tolerancia de una hora. Una anomalía genera una advertencia
fuerte con archivo y mes, pero no corta el proceso, tal como pide el plan.

La duplicación de 283 líneas queda documentada pero no se migró: el punto 3.4
indica explícitamente hacerlo por piezas y fuera de esta tarea. Pasaron el
checklist completo de la sección 9, las pruebas de `comun/`, el generador y
pruebas aisladas de la validación horaria, la cola UI y la continuación de hojas.
No se pudo hacer una prueba real con Access/Excel por ser un entorno Linux.
Queda pendiente la Tarea 4 experimental y la prueba visual/real en Windows.

## 2026-09-03 — Claude — verifica la Tarea 2 y corrige una regresión real

Codex hizo la Tarea 2 bien en lo estructural: partió del commit al día
(`fdaec6d`), localiza al Revisor por archivo (no por nombre de carpeta),
`CONFIG_PATH` y `SALIDAS` resuelven al mismo lugar que el Revisor, y
`Comparador_Tabulado` conservó `parquet_variables`/`vistas_variables` (el
comportamiento real, no las constantes muertas que tenía el original).
Verificado con AST que no quedó ninguna referencia colgante a las constantes
viejas (`CDIR`, `DIR_PARQUET`, etc.) — la migración a funciones por año fue
completa en los dos archivos.

**Encontrada una regresión real al leer el código, antes de correr nada
pesado.** Antes de esta tarea, `ACTUAL_PARQUET` (Etapas) y las rutas de
`Comparador_Tabulado` eran constantes de módulo **independientes del año** —
un bug/simplificación preexistente, pero que nunca lanzaba una excepción.
Ahora que correctamente dependen del año (`carpeta_comparador` →
`normalizar_anio`, que desde la corrección de la Tarea 1 lanza `ValueError`
con un año vacío o irreconocible), tres puntos podían reventar la ventana:

- `App.consolidar()` en **los dos** comparadores llamaba a `cargar_estado()`
  (y en Etapas también a `actual_parquet()`) **antes** de comprobar si había
  algo que consolidar. Con el año vacío o inválido — posible al apretar
  "Reconsolidar TODO" sin haber escrito un año — la ventana caía con
  `ValueError` sin ningún mensaje útil.
- `App.pintar_actual()` (Etapas) tenía el mismo riesgo si `self.est` traía
  `_actual.archivo` cargado de una sesión anterior pero `self.var_anio` ya no
  coincidía (por ejemplo, si el usuario borra el año a mano).

Se agregó una guarda al principio de los tres métodos, con el mismo estilo que
ya usaba el propio `__init__` (`if meses_del_anio(...) else ...`): si el año no
es válido, se avisa en el log y se corta antes de tocar el disco, en vez de
reventar. Reproducido el crash y confirmado el arreglo con un `App` simulado
(sin abrir ventana real) antes y después del parche.

Verificación completa de la sección 6.6 del plan, hecha por cuenta propia:
sintaxis de los 14 `.py`, las 13 pruebas de `comun/config.py`, las 13 de
`comun/salidas.py`, `generar_interfaces.py --check`, y — la parte que más
importaba — que `Comparador_Etapas` y `Comparador_Tabulado` resuelven **el
mismo** `CONFIG_PATH` que el Revisor, la misma `ruta_json_mes("2407")`, que
`Comparador_Tabulado.dir_parquet("2024")` sigue terminando en
`parquet_variables` (no en `parquet`), y que `Tabulado.cdir_mdb("2024")` ==
`Etapas.cdir("2024")` (la dependencia cruzada entre los dos).

Quedan la Tarea 3 (los 4 bugs ya documentados: threading, límite de Excel,
`hora_mes`, código duplicado) y la Tarea 4 (tema oscuro, experimental).

## 2026-09-03 — ChatGPT — Tarea 2: cablea los dos comparadores

Los dos comparadores localizan ahora la carpeta hermana que contiene
`Revisor_Reliquidacion.py`, usan exactamente su `config.json` y comparten con el
Revisor la lógica de `comun/salidas.py`. Todo su estado es anual:
`_comparador`, `_comparador_tabulado`, sus `estado.json`, `rutas.json`, parquet,
vistas y respaldos viven bajo `00_Salidas/AAAA/`; los JSON y Excel mensuales
usan `AAAA/MM Mes`. Tabulado conserva las carpetas reales
`parquet_variables`/`vistas_variables` y lee el `rutas.json` del comparador de
Access del mismo año.

Se registraron ambos scripts en el generador, `MAPA.md`, `README.md`, `AGENTS.md`
y `INTERFACES.md`. La verificación de la sección 6.6 confirmó sintaxis, pruebas
del módulo común, rutas esperadas y que ambos comparadores apuntan al mismo
`config.json` del Revisor. Quedan fuera de alcance las Tareas 3 y 4 y la futura
migración de sus copias de configuración a `comun/config.py`.

## 2026-09-03 — Claude — verifica la Tarea 1 y corrige dos cosas

Codex hizo la Tarea 1 bien: partió del commit al día (`3842447`), tocó solo lo
que correspondía y **no tocó los comparadores**, como pedía el plan. Verificado
por cuenta propia y no solo por su reporte: sus 9 pruebas, las 13 de config, el
chequeo del generador, la sintaxis del Revisor, y que `dir_mes("2407")` y
`dir_mes("sin_mes")` den las rutas exactas que pedía la sección 5.6.

Al leer `comun/salidas.py` aparecieron dos cosas que las pruebas no cubrían:

1. **`crear=True` partía el mes en dos carpetas.** Si el usuario ya había
   escrito `2024/7 Julio` a mano, la lectura iba a esa carpeta pero la primera
   escritura creaba `07 Julio` al lado: el estado se leía de una y se escribía
   en la otra, sin ningún aviso. **La culpa es de la especificación, no de
   Codex** — el plan decía textual "crear=True crea siempre la canónica".
   Corregido: si ya hay una variante, se usa esa también para escribir.
   Reemplazada la prueba `test_crear_usa_siempre_la_canonica`, que afirmaba el
   comportamiento viejo, dejando anotado por qué cambió.
2. **`carpeta_comparador` no normalizaba el año.** La ventana de los
   comparadores acepta el año con 2 o 4 dígitos (`meses_del_anio` ya lo hace),
   así que escribir "25" habría armado `00_Salidas/25/_comparador`. Habría
   aparecido recién en la Tarea 2, con los parquet de un año en una carpeta con
   otro nombre. Agregado `normalizar_anio`, y `carpeta_comparador` ahora lanza
   `ValueError` con un año irreconocible en vez de inventar una carpeta.

Cinco pruebas nuevas para los dos casos. `comun/salidas.py` queda en 13 pruebas.

## 2026-09-03 — ChatGPT — Tarea 1: `00_Salidas` por año y mes

Se creó `Revisor_Relq/comun/salidas.py` como única fuente para convertir AAMM
a `00_Salidas/AAAA/MM Mes`, con reconocimiento tolerante de variantes ya
existentes (`7 Julio`, `07 julio`) y detección separada de las carpetas planas
antiguas. El Revisor usa ahora ese módulo desde su único `dir_mes()`, sin cambiar
sus puntos de llamada, y al arrancar avisa en la bitácora de pantalla cuáles
carpetas antiguas debe mover el usuario a mano. `sin_mes` conserva el
comportamiento plano anterior.

Se agregaron 9 pruebas stdlib en `comun/test_salidas.py` y se actualizaron
`MAPA.md`, `AGENTS.md`, `README.md` e `INTERFACES.md`. Pasaron las 9 pruebas
nuevas, las 13 de config, el chequeo del generador, la sintaxis del Revisor y
la importación sin ventana que confirma las dos rutas pedidas en el plan. Los
comparadores no se tocaron. Quedan pendientes las Tareas 2 a 4 y que el usuario
mueva manualmente las carpetas antiguas informadas por el nuevo aviso.

## 2026-09-03 — Claude — la copia del repo que ve Codex puede estar vieja

**Lección de flujo de trabajo, para que no vuelva a pasar.** El usuario le pidió
a Codex aplicar la Tarea 1 del plan; Codex dijo "listo" y no llegó nada a
GitHub. Al pedirle la salida cruda de `git log`, `git status` y `ls`, quedó
claro qué pasó:

- Su commit de arriba (`8afb0ec`) colgaba de `e2443c8`, **3 commits atrás** del
  repo real. Le faltaban `d2b13f8`, `2955326` y `19d323f`.
- `2955326` es justamente el que trae `docs/PLAN_comparadores.md` **y** la
  carpeta `Comparadores/`. O sea: **el plan no existía en su copia**. No podía
  aplicarlo aunque quisiera, y terminó rehaciendo la tarea anterior (mismo
  cambio que el ya mergeado `1f16007`, con otro hash).
- `git status` decía `## work` sin upstream: esa rama no tiene remoto, así que
  los push no llegan a ningún lado.

**Regla que sale de esto:** antes de pedirle a un asistente externo que aplique
algo del repo, **pedirle que muestre `git log --oneline -3` y confirme que ve
los archivos que va a necesitar**. Si su copia no está en el commit de arriba,
no tiene sentido que empiece — va a trabajar sobre otra cosa y el resultado, si
llega, va a chocar. Cuesta cinco segundos y evita rehacer todo.

Nada que revertir: el trabajo duplicado quedó en el sandbox de Codex, nunca
llegó al repositorio.

## 2026-09-02 — Claude — revisión de bugs de los comparadores y plan del tema oscuro

El usuario se quedó sin acceso a Codex y pidió aprovechar para buscar bugs y
para pensar un tema oscuro. Revisión hecha **sin correr nada** (no hay Windows,
Excel ni Access acá): lectura del código más análisis por AST. Todo quedó en
`docs/PLAN_comparadores.md`, Tareas 3 y 4.

Cuatro hallazgos, en orden de gravedad:

1. 🔴 **Tkinter desde el hilo de trabajo, en los dos comparadores.** `App.log()`
   hace `self.txt.insert(...)` y `self.root.update_idletasks()` directo, y se lo
   llama desde `correr` y `buscar`, que corren en `threading.Thread`. Tkinter no
   es thread-safe: da cuelgues y crashes intermitentes, justo en las corridas
   largas. Los autores conocían la regla (usaron `after(0, ...)` para `pintar` y
   `botones`), pero el log se les escapó. Ninguno de los dos importa `queue`; el
   Revisor sí, y su patrón (`cola` + `_bombear_cola` cada 300 ms) es el arreglo
   a copiar.
2. 🔴 **`Comparador_Tabulado` corta el Excel en silencio** al pasar el límite de
   filas: hace `break` y descarta el resto, mientras `Comparador_Etapas` abre
   una hoja nueva y sigue. En una herramienta que existe para encontrar
   diferencias, un Excel incompleto que parece completo es lo peor que puede
   pasar. Además ese `break` no sale del `while`, así que vacía el cursor sin
   escribir.
3. 🟠 **`hora_mes` se calcula sin validar que el mes venga completo.** El uso de
   `max(hora_dia)` por día está bien pensado para el cambio de hora, pero si
   falta un día entero todas las horas siguientes se corren — y `hora_mes` es la
   clave con la que se cruzan las etapas. El resultado serían diferencias
   inventadas o diferencias reales tapadas, sin ningún error visible.
4. 🟡 **283 líneas idénticas duplicadas entre los dos comparadores** (33
   funciones, medidas por hash). Es `CENTRALES_EMBALSE` a mayor escala. No se
   migra ahora; queda el número y el orden sugerido en el plan.

Se revisó y **está bien**: el cierre de conexiones Access y libros Excel (todos
con `try/finally`), el marshalling de `pintar`/`botones`/`set_estado` con
`after(0, ...)`, el contrato del JSON de traspaso, y que
`rutas_desde_json_mes` se comporte igual en los dos archivos (difiere solo en el
nombre de una variable).

Sobre el tema oscuro: se plantea como piloto **solo en los comparadores** (los
que menos se usan, así un experimento fallido no molesta en el trabajo diario),
con `comun/tema.py`, apagado por omisión y conmutable desde `config.json`. Se
anotan las tres cosas que hacen que un tema oscuro en tkinter quede a medias:
los widgets `tk.` clásicos no siguen a ttk, los colores de estado están
calibrados para fondo claro, y `SystemButtonFace` queda como un parche gris.

## 2026-09-02 — Claude — plan de los comparadores y de `00_Salidas` por año

El usuario pidió: los dos comparadores a una carpeta hermana `Comparadores/`,
tomando las rutas del mismo origen que el Revisor, y `00_Salidas` reorganizada
en `AAAA/MM Mes` con `_comparador` y `_comparador_tabulado` dentro de su año.
Pidió explícitamente planificar, no ejecutar: lo aplica Codex.

Se leyeron los dos comparadores y el código del Revisor antes de planificar.
Tres hallazgos que cambian el plan:

1. **`dir_mes()` del Revisor es el único punto por donde pasa el armado de la
   carpeta del mes** (~10 llamadas). Cambiar la estructura es cambiar una sola
   función, no diez. Y ningún actualizador usa `DIR_SALIDAS`: reciben la ruta
   del traspaso por `argv[1]`.
2. **Los comparadores ya leen el `config.json` compartido y ya leen/escriben
   los `_traspaso_actualizador.json` del Revisor** (Etapas escribe bajo su
   propia clave `comparador_etapas`, sin tocar nada ajeno). O sea que el pedido
   de "las rutas del mismo origen" ya está casi resuelto en el código: solo
   falta que `CONFIG_PATH` siga apuntando al del Revisor después de la mudanza.
3. **Bug en `Comparador_Tabulado.py`:** `DIR_PARQUET` y `DIR_VISTAS` están
   definidos dos veces (159-160 y 827-828). Todos los usos vienen después de
   la 827, así que las carpetas reales son `parquet_variables` /
   `vistas_variables`; las líneas de arriba son código muerto que engaña a
   quien lee el bloque de constantes. Queda anotado en el plan para que el
   refactor preserve el comportamiento real y borre las muertas.

Decisión de diseño: la lógica AAMM → carpeta va en `comun/salidas.py`, una sola
vez, importada por el Revisor y por los dos comparadores. Si se duplicara, uno
leería donde el otro no escribe **sin ningún error visible** — el mismo tipo de
bug que ya documentó `CENTRALES_EMBALSE`. Y como `Comparadores/` es hermana de
`Revisor_Relq/`, los comparadores localizan al Revisor buscando el archivo
`Revisor_Reliquidacion.py` entre las carpetas hermanas, no por el nombre de la
carpeta — que ya cambió tres veces.

Los dos `.py` se subieron sin modificar a `Comparadores/` para que Codex pueda
trabajarlos (solo existían como adjuntos en el chat). **Quedan rotos en esa
ubicación hasta la Tarea 2**, y así lo dice `Comparadores/README.md`.

## 2026-09-02 — Claude — mergea la rama de Codex, verifica el cambio de `00_Salidas`

El usuario le pidió a **Codex** (no a ChatGPT) mover `DIR_SALIDAS` fuera de
`Revisor_Relq/`. Codex no tenía push configurado en su sandbox y avisó que el
commit quedó solo local — pero apareció igual en GitHub, en la rama
`codex/modificar-carpeta-de-salida-del-revisor` (`1f16007`), aparentemente
subida por el usuario. Esa rama partía exactamente del commit anterior de
esta bitácora (`e2443c8`), sin divergencia, así que se pudo traer con un
`git merge --ff-only` — nada que resolver.

Se verificó de nuevo por cuenta propia, sin confiar solo en lo que reportó
Codex: `generar_interfaces.py --check` (al día), sintaxis del Revisor,
`DIR_SALIDAS` leído por AST (`DIR_SCRIPT.parent / '00_Salidas'`, exacto), y
las 13 pruebas de `comun/test_config.py` (no tocadas por este cambio, pero
confirman que sigue sano). Todo coincidió con lo reportado.

Dos correcciones sobre la entrada de acá abajo, sin editarla (así lo pide
`REGLAS.md`): la firma dice "ChatGPT" pero fue Codex quien lo hizo; y donde
dice "no queda nada pendiente" en realidad sí quedó algo menor — 3
comentarios y el documento de dominio con la ruta vieja sin actualizar (ver
"Pendientes abiertos", arriba). No bloquea nada, es cosmético.

## 2026-09-02 — ChatGPT — mueve las salidas fuera de `Revisor_Relq`

El Revisor ahora resuelve `DIR_SALIDAS` como
`DIR_SCRIPT.parent / "00_Salidas"`: la carpeta de estado, caché y traspaso queda
como hermana de `Revisor_Relq/`, no adentro. Se actualizaron los textos del
Revisor, el mapa, las instrucciones, el README y la exclusión de Git. No queda
nada pendiente de este cambio.

## 2026-09-02 — Claude — reglas obligatorias y esta bitácora

El usuario va a darle a ChatGPT el mismo acceso de escritura por git que
tiene este asistente, así que ninguno de los dos se entera de lo que hace el
otro salvo que lo lea. Se crean `REGLAS.md` (checklist obligatorio, se lee
primero, antes que `AGENTS.md`) y este archivo. `AGENTS.md` sección 3 se
recorta para apuntar a `REGLAS.md` en vez de duplicar la lista. *(08ba96a y
el commit que sigue)*

## 2026-09-02 — Claude — nombre final de la carpeta de trabajo

Tres vueltas en la misma sesión, a pedido del usuario: `scripts` →
`Revisor Reliquidación` → `Revisor` → **`Revisor_Relq`** (nombre
definitivo). Cada vuelta fue puro renombre — 0 líneas de código tocadas en
cada `.py`, confirmado con `git diff --stat` antes de subir — más la
actualización de `generar_interfaces.py` (la lista `CARPETAS`) y los tres
`.md` que mencionan la ruta. Un bug real apareció en la primera vuelta: el
generador de anclas de `INTERFACES.md` le borraba las tildes a los nombres
con un regex ASCII (`[^a-z0-9]`); no se notaba antes porque ningún nombre de
carpeta tenía tilde. Corregido a filtrar por `c.isalnum()`.
*(362b1c1, 87ae9a2, a8122f2)*

## 2026-09-02 — Claude — separa el Revisor de los actualizadores en carpetas

A pedido del usuario, por prolijidad: `Revisor_Reliquidacion.py` queda solo
en la raíz; los 8 que lanza por botón pasan a `actualizadores/`;
`Reemplazos REUC/` sigue aparte (tiene su propio `config.json`, no el
compartido). Lo delicado: `config.json` es compartido entre el Revisor y los
8 actualizadores, y cada uno lo resolvía relativo a su **propia** carpeta —
hubo que cambiar `DIR_SCRIPT / "config.json"` a
`DIR_SCRIPT.parent / "config.json"` en los 8 para que siguieran mirando el
mismo archivo. Verificado de punta a punta con una corrida simulada (un
script escribe una ruta, otro la lee, mismo archivo, confirmado). *(1d7ad05)*

## 2026-09-02 — Claude — módulo `comun/`, primera pieza: `config.py`

El manejo del `config.json` estaba copiado en los 10 scripts, con 4
variantes de `_modificar_config` y 4 de `get_usuario`. Dos diferencias no
eran cosméticas: `ActualizaRemplazos.py` podía borrarle los ajustes a los
otros nueve si el archivo estaba roto (no pasaba por un `.tmp` antes de
escribir), y `get_usuario` tenía versiones con y sin `try/except`. Quedó la
versión defensiva de cada una en `comun/config.py`, con 13 pruebas en
`comun/test_config.py`. Se migró un solo script (`Actualiza_SC_CO.py`) para
no tocar los 10 de una — quedan 7 pendientes (ver arriba). *(6f7796c)*

## 2026-09-02 — Claude — generador ajustado al estilo real; MAPA.md validado

Con los 10 `.py` ya subidos por el usuario, se ajustó `generar_interfaces.py`
para reconocer el estilo real de comentarios del repo (encabezado como
docstring o bloque `#`, banners de sección en dos formatos) en vez de
tomarlos como descripción de la función de al lado. Se validó cada bloque de
`MAPA.md` contra el código y aparecieron 5 diferencias con
`docs/ESTRUCTURA_CASO_RELIQUIDACION.md` (`UMBRAL_DESCUADRE_CPRT` real es
1000 no 500, `CAPACIDADES` son 5 no 4, etc.) — quedaron anotadas en
`MAPA.md` → "Diferencias con el documento de dominio". *(4ecb99b)*

## 2026-09-02 — Claude — repositorio inicial (estructura de control)

Se armó `AGENTS.md`, `MAPA.md` e `INTERFACES.md` (generado por
`generar_interfaces.py`) **antes** de que llegara ningún `.py`, con los
bloques de `MAPA.md` deducidos del documento de dominio y marcados para
validar cuando llegara el código real. `docs/ESTRUCTURA_CASO_RELIQUIDACION.md`
entró tal cual, sin tocar. *(0996254)*

## 2026-09-02 — usuario — sube los 10 scripts

Los 9 scripts principales y `ActualizaRemplazos.py`, subidos por la web del
repositorio (todavía sin acceso de escritura por git para ningún asistente
en ese momento). *(7a5fc7a, d53251e)*
