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

- [ ] **Aplicar `docs/PLAN_reorganizacion_config.md`** — mover `comun/` a
      `__comun__/` (hermana, se sube al repo) y sacar todos los `.json` del
      repo a `__config__/` (hermana, NO se sube — la usuaria la arma a mano
      con la misma estructura `AAAA/MM Mes` que `00_Salidas`). `00_Salidas/`
      queda solo con los `.xlsx` finales. Plan en 3 tareas, cada una con su
      verificación — ver el archivo para el detalle exacto. Mientras esto no
      esté aplicado, varios de los pendientes de abajo (el módulo se llama
      `comun/config.py`, rutas de `00_Salidas` para estado) van a quedar
      obsoletos apenas se aplique: no corregirlos dos veces, corregirlos ya
      con el nombre y la ruta nueva cuando se toquen esos archivos.
- [ ] El usuario no probó todavía ningún actualizador real de punta a punta
      (solo los verificadores del Revisor, que funcionan bien). Falta correr
      al menos uno contra archivos reales.
- [ ] Migrar los 7 actualizadores que faltan a `comun/config.py` — solo
      `Actualiza_SC_CO.py` está migrado. Uno a la vez, verificando que sigue
      corriendo antes de seguir con el próximo (ver `MAPA.md` → "El módulo
      común").
- [ ] `docs/ESTRUCTURA_CASO_RELIQUIDACION.md` tiene 5 diferencias conocidas
      contra el código real, listadas en `MAPA.md` → "Diferencias con el
      documento de dominio". El documento de dominio todavía no se corrigió.
- [ ] Confirmar si el `1_CUADROS_PAGO` que busca `ActualizaRemplazos.py` en
      `T:\Facturacion\<mes>\<versión>` es el mismo archivo que el
      `00 Entregables` que usa el Revisor (documento de dominio, sección 10).
- [ ] Quedaron 3 comentarios `# ... Salidas/AAMM/ ...` (sin el `00_` nuevo) en
      `Actualiza_datos.py`, `Actualiza_Data_Access.py` y
      `ActualizaRemplazos.py` — scripts que todavía no leen ni escriben
      `DIR_SALIDAS` directamente, así que no son un bug funcional, pero van a
      quedar mal si algún día alguno de esos scripts empieza a usar esa ruta.
      Corregirlos de paso la próxima vez que se toque cualquiera de esos tres
      archivos. `docs/ESTRUCTURA_CASO_RELIQUIDACION.md` tiene las mismas 2
      menciones viejas — se suma a la fila ya abierta de diferencias con el
      código real (arriba).

- [ ] Probar visualmente en Windows los temas claro y oscuro de los dos
      comparadores. La verificación automatizada corrió con `tkinter` real
      (instalado en este entorno) y `ttk.Style` simulado, pero sin pantalla no
      hay forma de ver si el resultado es realmente legible/prolijo.
- [ ] El usuario tiene que mover a mano las carpetas de `00_Salidas` al formato
      `AAAA/MM Mes`. La Tarea 1 agrega un aviso que dice cuáles faltan.
- [ ] Los comparadores tienen su propia copia de `leer_config` /
      `guardar_config` / `escribir_json_atomico`. Migrarlas a `comun/config.py`
      como ya se hizo con `Actualiza_SC_CO.py` — **no** dentro de las tareas
      del plan, después y por separado.

---

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
