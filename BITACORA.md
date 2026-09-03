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

- [ ] **En curso — `docs/PLAN_comparadores.md`.** Tarea 1 (`comun/salidas.py` +
      `dir_mes` del Revisor) terminada; faltan Tarea 2 (cablear los dos
      comparadores), Tarea 3
      (4 bugs encontrados) y Tarea 4 (tema oscuro, experimental). Los dos `.py`
      ya están en `Comparadores/` pero **sin cablear: hoy no corren en esa
      ubicación**. Borrar ese plan cuando esté todo aplicado.
- [ ] **Bug 🔴 sin arreglar:** los dos comparadores escriben en la ventana desde
      el hilo de trabajo (`App.log`). Da cuelgues intermitentes en corridas
      largas. Plan, 3.1.
- [ ] **Bug 🔴 sin arreglar:** `Comparador_Tabulado` descarta filas en silencio
      al pasar el límite de Excel, en vez de seguir en una hoja nueva como hace
      `Comparador_Etapas`. Plan, 3.2.
- [ ] **Bug 🟠 sin arreglar:** `Comparador_Tabulado` calcula `hora_mes` sin
      comprobar que el mes venga completo; un día faltante corre todas las horas
      siguientes y cruza etapas mal. Plan, 3.3.
- [ ] El usuario tiene que mover a mano las carpetas de `00_Salidas` al formato
      `AAAA/MM Mes`. La Tarea 1 agrega un aviso que dice cuáles faltan.
- [ ] Los comparadores tienen su propia copia de `leer_config` /
      `guardar_config` / `escribir_json_atomico`. Migrarlas a `comun/config.py`
      como ya se hizo con `Actualiza_SC_CO.py` — **no** dentro de las tareas
      del plan, después y por separado.

---

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
