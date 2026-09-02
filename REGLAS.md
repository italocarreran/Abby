# REGLAS.md — obligatorio, sin excepciones

> Se lee ENTERO, primero, en cualquier sesión, antes de tocar nada — antes
> incluso que `AGENTS.md`. Es corto a propósito: si algún día deja de serlo,
> hay que acortarlo, no saltearlo.
>
> Este repositorio lo edita más de un asistente de IA con acceso de
> escritura por git (Claude, ChatGPT), sin que ninguno vea lo que hace el
> otro en tiempo real. No hay notificaciones, no hay bloqueo, no hay nada
> que avise si el otro tocó algo. Estas reglas son la única red de
> seguridad que existe. No son sugerencias.

---

## Al empezar CUALQUIER sesión, en este orden

1. **`git log -n 15`** (o más si hace falta). El autor de cada commit dice
   quién lo hizo — los de Claude quedan firmados `Claude
   <noreply@anthropic.com>`, los que suba el usuario a mano quedan a su
   nombre, un commit de ChatGPT va a tener su propia firma. Si hay uno
   reciente que no es propio, **leer el mensaje completo** antes de asumir
   en qué estado está el repo.
2. **Leer `BITACORA.md` entero**, empezando por "Pendientes abiertos". Es lo
   que `git log` no cuenta: qué falta probar, qué quedó a medias, qué se
   decidió y por qué.
3. **`git pull`** (traer la rama al día) antes del primer cambio. Trabajar
   sobre una copia vieja es la forma más común de pisar el trabajo del otro
   asistente sin que nadie se entere hasta después.

## Mientras se trabaja

4. **No dejar un cambio a medias.** Si toca varios archivos (una
   reorganización, un módulo compartido), terminarlo y verificarlo en la
   misma sesión. Si de verdad no se puede terminar, dejarlo dicho explícito
   en `BITACORA.md`: qué falta, por qué se cortó, cuál es el paso siguiente.
   Nunca dejarlo mudo — "la próxima sesión" puede ser el otro asistente, sin
   el contexto de esta.
5. **Antes de dar un cambio por terminado:** correr
   `python generar_interfaces.py --check` y, si existe, el test suite del
   módulo tocado (por ejemplo `python Revisor_Relq/comun/test_config.py`).
   Si el cambio toca algo que lee o escribe `config.json`, verificar de
   punta a punta que sigue siendo el mismo archivo compartido — no alcanza
   con que el `.py` compile.
6. **Nunca reescribir el historial compartido.** Nada de `push --force`,
   `commit --amend` sobre algo ya subido, ni `reset --hard` contra la rama
   remota. Otro asistente puede estar construyendo sobre esos commits.

## Antes de cerrar CUALQUIER sesión que haya cambiado algo

7. **Agregar una entrada en `BITACORA.md`** — qué se hizo, qué quedó
   pendiente, qué haría falta que sepa el próximo que llegue.
8. **Actualizar `MAPA.md` y/o `AGENTS.md`** si cambió algo estructural, una
   firma, o se tomó una decisión de diseño.
9. **Regenerar `INTERFACES.md`** (`python generar_interfaces.py`) si cambió
   la firma de algo.
10. **Commitear y pushear.** Un cambio que queda solo en el working tree de
    esta sesión no existe para el otro asistente.

## `BITACORA.md` es de solo agregar

Nunca se edita ni se borra una entrada vieja — ni siquiera la propia. Si algo
quedó mal anotado, se agrega una entrada nueva que lo corrige. Es un log, no
un documento a pulir. La única sección que sí se edita es "Pendientes
abiertos": ahí se tacha o se saca el ítem cuando se resuelve, porque es un
estado, no un historial (ver el encabezado de ese archivo).

## Si dos asistentes están por tocar lo mismo

No hay forma de bloquear esto desde acá — lo único que existe es lo que se
lea en `git log` y `BITACORA.md`. Si hay señales de que el otro asistente
empezó algo relacionado hace muy poco (mismo día, mismo tema), preguntarle
al usuario antes de seguir, en vez de asumir que el camino está libre.
