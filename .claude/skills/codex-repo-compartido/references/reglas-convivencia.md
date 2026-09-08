# Reglas de convivencia entre dos asistentes

Plantilla para el archivo de reglas del repositorio. Adaptar los nombres de
archivos y comandos al proyecto; lo que no conviene tocar es el **por qué** de
cada regla, que es lo que hace que un agente las respete en vez de saltearlas.

Dónde ponerlas: en el archivo que los dos asistentes leen al arrancar. Claude
Code lee `CLAUDE.md` y también `AGENTS.md`; Codex lee `AGENTS.md`. Si se quiere
un solo lugar, `AGENTS.md` es el mínimo común denominador. Un archivo corto y
obligatorio aparte (`REGLAS.md`) funciona bien cuando `AGENTS.md` ya es largo:
lo corto se lee entero, lo largo se hojea.

---

## Por qué hacen falta

Dos asistentes con acceso de escritura por git, sin que ninguno vea lo que hace
el otro en tiempo real. No hay notificaciones, no hay bloqueo, no hay nada que
avise si el otro tocó algo. Estas reglas son la única red de seguridad que
existe. Conviene decirlo así de explícito en el archivo: un agente que entiende
que es la única red de seguridad las trata distinto que una lista de estilo.

---

## Plantilla

### Al empezar cualquier sesión, en este orden

1. **`git log -n 15`.** El autor de cada commit dice quién lo hizo. Si hay uno
   reciente que no es propio, leer el mensaje completo antes de asumir en qué
   estado está el repositorio.
2. **Leer la bitácora entera**, empezando por los pendientes abiertos. Es lo
   que `git log` no cuenta: qué falta probar, qué quedó a medias, qué se
   decidió y por qué.
3. **Sincronizar**: `bash scripts/sincronizar.sh`. Trabajar sobre una copia
   vieja es la forma más común de pisar al otro sin que nadie se entere hasta
   después.
4. **Y de nuevo al empezar cada turno, no solo al abrir la sesión.** Esto es
   sobre todo para Codex: su contenedor clona el repositorio una sola vez y no
   se actualiza nunca solo. Si el fetch falla, decírselo al usuario en vez de
   seguir a ciegas.
5. **La rama base es una sola.** Todo sale de ahí y todo vuelve ahí. Una rama
   de trabajo se fusiona y se borra; no se acumulan ramas vivas.

### Mientras se trabaja

6. **No dejar un cambio a medias.** Si de verdad no se puede terminar, dejarlo
   dicho explícito en la bitácora: qué falta, por qué se cortó, cuál es el paso
   siguiente. "La próxima sesión" puede ser el otro asistente, sin el contexto
   de esta.
7. **Antes de dar un cambio por terminado**, correr la verificación del
   repositorio y dejarla en verde.
8. **Nunca reescribir el historial compartido.** Nada de `push --force`,
   `commit --amend` sobre algo ya subido, ni `reset --hard` contra el remoto.
   El otro asistente puede estar construyendo sobre esos commits.

### Antes de cerrar cualquier sesión que haya cambiado algo

9. **Agregar una entrada en la bitácora**: qué se hizo, qué quedó pendiente,
   qué haría falta que sepa el próximo que llegue.
10. **Commitear y pushear.** Un cambio que queda solo en el working tree de
    esta sesión no existe para el otro asistente.

### Si dos asistentes están por tocar lo mismo

No hay forma de bloquear esto. Lo único que existe es lo que se lea en
`git log` y en la bitácora. Ante señales de que el otro empezó algo relacionado
hace muy poco, preguntarle al usuario antes de seguir.

---

## La bitácora

Un archivo de solo agregar, con las entradas nuevas arriba. Nunca se edita ni
se borra una entrada vieja: si algo quedó mal anotado, se agrega una entrada
nueva que lo corrige. Es un log, no un documento a pulir.

La única sección que sí se edita es la de pendientes abiertos, porque es un
estado y no un historial.

Por qué importa más de lo que parece: es el único canal entre los dos
asistentes. Todo lo que no quede escrito ahí, el otro no lo va a saber nunca.

---

## La firma de los commits

Los commits que Codex sube por el botón de PR **pueden salir con el nombre del
dueño de la cuenta**, porque quien abre el PR es esa cuenta. Eso rompe la regla
1: mirar el autor deja de servir para saber quién tocó qué.

Dos mitigaciones, se usan juntas:

- El setup script configura `user.name` y `user.email` del agente en el
  contenedor.
- Por convención, los mensajes de commit del agente llevan un prefijo (`codex:`
  o el que se elija). Es lo que sobrevive aunque la atribución se pierda.
