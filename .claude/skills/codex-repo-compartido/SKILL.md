---
name: codex-repo-compartido
description: >
  Configura un repositorio de GitHub para que Codex (el de ChatGPT, en la web) y Claude
  trabajen sobre el mismo código sin pisarse, y sobre todo para que una tarea de Codex ya
  abierta vea los commits subidos después de que arrancó, sin abrir un chat nuevo cada vez.
  Usá esta skill siempre que alguien cuente que Codex no ve los cambios, que trabaja sobre
  una copia vieja o desactualizada, que tiene que abrir otro chat para que se entere de algo,
  que no sabe cómo seguir después del botón "Create PR", o que dos asistentes de IA (Codex,
  ChatGPT, Claude, Copilot) editan el mismo repositorio y se pisan entre ellos. También
  cuando pidan configurar el environment de Codex, su setup script o script de mantenimiento,
  el acceso a internet del agente, o cuando aparezca el error "'origin' does not appear to be
  a git repository" dentro de una tarea de Codex. Sirve igual si el repositorio es público o
  privado, y no requiere tener Codex instalado localmente.
---

# Repositorio compartido entre Codex y Claude

## Qué está pasando en realidad

Casi siempre la persona llega diciendo alguna versión de *"Codex no ve lo que
subo"*. Antes de tocar nada conviene explicar el mecanismo, porque cambia la
expectativa y evita que busquen una opción que no existe:

1. Codex crea una tarea y levanta un contenedor aislado.
2. **Clona el repositorio una sola vez**, en ese momento, desde la rama base.
3. Todo lo que se hable después en esa misma tarea corre en ese contenedor, con
   **ese clon congelado**.
4. El botón de PR empuja el resultado a una rama en GitHub.

El punto 3 es el problema entero. Nadie le avisa a la tarea de los commits
nuevos y nada la actualiza sola. Por eso abrir un chat nuevo "arregla" el
asunto: un chat nuevo es un clon nuevo.

**No existe una opción que haga que una tarea ya empezada se sincronice sola.**
Decirlo explícitamente. Lo que sí se puede es que haga `fetch` cuando se lo
piden, en cinco segundos, y que una regla escrita en el repositorio la obligue
a hacerlo sin que nadie se lo recuerde. Eso es lo que instala esta skill.

## Antes de empezar: cuatro datos

Preguntar lo que no se pueda averiguar solo. Cada uno cambia decisiones
concretas más adelante, así que no son trámite:

1. **Qué repositorio** (`owner/repo`) y si es **público o privado**. Público →
   `fetch` anónimo, nada que configurar. Privado → hace falta un token como
   secreto del environment.
2. **Cuál es la rama base estable**, si es que existe. Muy seguido no existe:
   ver el paso 1.
3. **Quién más edita el repositorio** — otro asistente, la persona a mano, un
   equipo. Define si hacen falta las reglas de convivencia o alcanza con la
   sincronización.
4. **Qué comando verifica que el proyecto está sano** (tests, linter, build).
   Se necesita para el setup script y para las reglas.

Muchos de estos se pueden mirar en vez de preguntar: `git remote -v`,
`git ls-remote --heads origin`, el `README`, los archivos de test.

---

## Paso 1 — Una rama base estable

Mirar qué hay antes de asumir:

```bash
git ls-remote --symref origin HEAD    # cuál es la rama por defecto
git ls-remote --heads origin          # todas las ramas
```

Dos patologías frecuentes, y son la causa raíz de bastante desorden:

- **No hay `main`**, y la rama por defecto es una rama de sesión de un agente
  (`claude/algo-xyz`, `codex/otra-cosa`). Cada tarea nueva arranca desde una
  rama cuyo nombre cambia: no hay punto fijo contra el cual sincronizar.
- **Se acumulan ramas** de tareas viejas, una por cada PR que se abrió.

Antes de proponer borrar nada, verificar qué está fusionado — nunca de
memoria ni por el nombre:

```bash
for b in $(git ls-remote --heads origin | sed 's#.*refs/heads/##'); do
  git fetch origin "$b" --quiet
  echo "$b  adelante_de_base=$(git rev-list --count origin/<base>..FETCH_HEAD)"
done
```

Las que dan `0` están fusionadas y se pueden borrar sin perder nada. Con
cualquier otra, avisar y no tocar.

Crear la rama estable y ponerla por defecto **es acción del usuario en la
interfaz de GitHub** (Settings → General → Default branch): no hay forma de
hacerlo desde git. Dar los pasos concretos y esperar a que confirme.

## Paso 2 — Los scripts, al repositorio

Copiar `scripts/sincronizar.sh` y `scripts/setup_codex.sh` desde `scripts/` de
esta skill, y **completar el bloque de configuración** que está arriba de todo
en cada uno: la URL del repositorio, la rama base, y las dependencias y la
verificación del proyecto en el setup.

Estos dos scripts existen porque el contenedor de Codex tiene una particularidad
que no está documentada en ningún lado y cuesta horas descubrir: **clona el
repositorio sin dejar configurado el remoto `origin`**. Un `git fetch origin`
pelado falla con `'origin' does not appear to be a git repository` aunque la red
esté perfecta. Los dos scripts lo configuran solos. Por lo mismo
`sincronizar.sh` trabaja contra `FETCH_HEAD` y no contra `origin/<rama>`: en un
clon sin remoto esa referencia tampoco existe.

Si el proyecto ya tiene un script de verificación, apuntar el setup a ese en vez
de inventar uno nuevo.

**Fusionar esto a la rama base antes de seguir.** Codex clona la rama base y
recién ahí corre el setup script: si los archivos no están ahí, falla.

## Paso 3 — Las reglas en el repositorio

Solo si hay más de un asistente editando. Leer
`references/reglas-convivencia.md` y adaptarlo al proyecto.

La regla que resuelve el problema original es una sola: **sincronizar al empezar
cada turno, no solo al abrir la sesión.** Como queda escrita en el archivo que
los agentes leen al arrancar (`AGENTS.md`, que leen tanto Codex como Claude), se
cumple sin que nadie la recuerde. Sin esa regla, la persona vuelve a tener que
pedirlo a mano cada vez, que es el 80% de la molestia original.

## Paso 4 — El environment de Codex

Leer `references/entorno-codex.md`, que tiene la pantalla campo por campo. Lo
que hay que transmitir sí o sí, porque son las tres trampas que hacen fallar
todo lo demás:

- Hay **dos** campos de script y hacen cosas distintas: el de configuración
  corre al crear el contenedor, el de mantenimiento al reanudarlo desde caché.
  La sincronización va en el de mantenimiento — es justo cuando el clon quedó
  viejo.
- En métodos HTTP hay que dejar **"Todos los métodos"**. Git negocia el `fetch`
  con un `POST`, así que "GET, HEAD y OPTIONS" no sirve.
- El cartel *"el acceso a la red siempre está habilitado en este paso"* que
  aparece bajo los scripts vale solo para esos pasos, no para cuando el agente
  trabaja. Es distinto del ajuste de acceso a internet del agente, que hay que
  activar aparte.

Sobre el aviso de riesgo que muestra esa pantalla: no minimizarlo. Explicar qué
riesgo aplica en el caso concreto (`references/entorno-codex.md` tiene el
desglose) y dejar que la persona decida.

## Paso 5 — Probar el caso que importa

Acá es donde es fácil declarar victoria antes de tiempo. Que
`scripts/sincronizar.sh` corra limpio **no prueba nada** si no había commits
nuevos que traer: prueba que el script no explota, no que resuelve el problema.

La prueba real tiene tres partes y hay que hacerla en ese orden:

1. Abrir una tarea en Codex y pedirle que corra `bash scripts/sincronizar.sh`.
2. **Con esa tarea todavía abierta**, subir un commit a la rama base desde otro
   lado.
3. Pedirle **a esa misma tarea**, sin reabrirla, que corra el script otra vez.

Si en el segundo intento lista el commit nuevo con su autor y lo fusiona, está
resuelto. Ese es el momento de decir que funciona, y no antes.

---

## Trampas conocidas

Las que cuestan tiempo si no se saben de antemano:

| Síntoma | Causa |
|---|---|
| `'origin' does not appear to be a git repository` | El contenedor clona sin remoto. Lo resuelven los scripts |
| El fetch falla con la red activada | Métodos HTTP en "GET, HEAD y OPTIONS"; git necesita `POST` |
| El setup script falla con `No such file or directory` | Los scripts no están todavía en la rama base |
| Codex no ve un commit reciente | La tarea arrancó antes y nadie sincronizó ese turno |
| Codex arranca de una base rara | La rama por defecto del repositorio no es la estable |
| Los commits de Codex salen con el nombre del usuario | Quien abre el PR es esa cuenta. Se mitiga con prefijo en el mensaje |
| `@codex` no responde en un PR | La GitHub App de Codex no está instalada en ese repositorio |

## Lo que sigue sin poder hacerse

Conviene decirlo al cerrar, para que nadie persiga una opción que no existe:

- Una tarea de Codex **no se entera sola de nada**. La actualización siempre la
  dispara alguien: la persona pidiéndola, o el agente por la regla escrita.
- Los dos asistentes **no se ven en tiempo real**. El único canal entre ellos es
  lo que quede escrito en la rama base: los commits y la bitácora.
- Codex **no fusiona**. Entrega un PR; fusionar lo hace una persona o el otro
  asistente.
- Codex **no puede hacer `git push`** desde el contenedor. La entrega sale por
  el botón de la tarea.

## Al cerrar

Dejar la configuración escrita **en el repositorio**, no solo dicha en el chat:
un documento con los pasos, la tabla de síntomas y causas, y qué quedó
pendiente. La próxima vez que algo falle, la persona no va a tener a mano esta
conversación.
