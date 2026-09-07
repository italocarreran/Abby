# Flujo de trabajo Claude + Codex sobre este repositorio

> Para el usuario, no para los asistentes. Esto es la configuración: qué hay
> que apretar una sola vez para que Codex vea lo que Claude sube sin tener que
> abrir un chat nuevo cada vez.
>
> Las reglas que siguen los asistentes están en `REGLAS.md` y `AGENTS.md`.

---

## 1. Por qué pasa lo que pasa

Codex en la web no trabaja "conectado" al repositorio. Funciona así:

1. Creás una tarea.
2. Codex levanta un contenedor aislado y **clona el repositorio una sola vez**,
   en ese momento, desde la rama base que tenga configurada.
3. Todo lo que hablás después dentro de esa misma tarea corre en ese mismo
   contenedor, con **ese clon congelado**.
4. El botón "Create PR" empuja el resultado a una rama nueva en GitHub.

El punto 3 es el problema entero. Si Claude sube un commit después de que la
tarea arrancó, la tarea no se entera: nadie le avisa y nada la actualiza sola.
Por eso abrir un chat nuevo "arregla" el asunto — un chat nuevo es un clon
nuevo. No es que Codex esté mal configurado; es cómo está construido.

**No existe una opción que haga que una tarea ya empezada se sincronice sola.**
Lo que sí existe es que la tarea haga `git fetch` cuando se lo pedís, y eso
tarda cinco segundos. Para que eso funcione hacen falta tres cosas, y las tres
son los pasos de abajo:

- una **rama base estable** contra la cual sincronizar (paso 2),
- **acceso a internet** en el entorno, o el `fetch` falla (paso 4),
- una **regla escrita en el repo** que obligue a Codex a hacerlo en cada turno
  sin que se lo tengas que recordar (paso 5 — ya está hecha).

---

## 2. Paso 0 — crear `main` y dejarla como rama por defecto

> **Hecho el 2026-09-07.** `main` existe, es la rama por defecto y las once
> ramas viejas están borradas. Se deja escrito el paso por si alguna vez hay
> que rehacerlo o entender por qué el repositorio está armado así.

**Este era el paso más importante y el que faltaba.** Hasta ese día el
repositorio **no tenía una rama `main`**: la rama por defecto era
`claude/eso-uozpi4`, el nombre de una sesión de Claude, y había once ramas
`codex/*` colgando de tareas viejas.

Eso significa que cada tarea de Codex arranca desde una rama cuyo nombre cambia
cada tanto. No hay ningún punto fijo contra el cual sincronizar, y por eso
tampoco sirve de nada configurar el entorno: la base se mueve.

Las once ramas ya están todas fusionadas en la rama por defecto — no se pierde
nada al borrarlas.

En GitHub, en `italocarreran/Abby`:

1. **Branches** (arriba, junto a la cantidad de ramas) → **New branch**.
   Nombre: `main`. Source: `claude/eso-uozpi4`. **Create**.
2. **Settings** → **General** → sección **Default branch** → el ícono de las
   dos flechas → elegir `main` → **Update** → confirmar.
3. Volver a **Branches** y borrar las ramas `codex/*` y `claude/*` viejas con
   el tacho. Todas están fusionadas; si alguna te da aviso de "no fusionada",
   dejala y avisá.

De acá en adelante: **`main` es la única verdad.** Claude fusiona a `main`,
Codex arranca de `main`, vos sincronizás contra `main`.

---

## 3. Paso 1 — conectar el repositorio a Codex

En `chatgpt.com/codex`, si no lo hiciste ya:

1. **Connect to GitHub** → autorizar.
2. Instalar la GitHub App de Codex **en `italocarreran/Abby`**. Si te deja
   elegir, dale acceso a ese repositorio solamente.

Esto además habilita el paso 7, que es el atajo que más te va a servir.

---

## 4. Paso 2 — crear el entorno (Environment)

**Settings → Codex → Environments → New environment** (o editar el que ya
tengas para este repositorio):

| Campo | Qué poner |
|---|---|
| Repository | `italocarreran/Abby` |
| Base branch | `main` |
| Script de configuración | **Manual** → `bash scripts/codex_setup.sh` |
| Script de mantenimiento | **Manual** → `bash scripts/sincronizar.sh \|\| true` |
| Container image | la que venga por defecto (universal) |

**Script de configuración** (`scripts/codex_setup.sh`): corre una sola vez, al
armar el contenedor. Instala `tkinter` (que no viene y lo necesita un test),
**configura el remoto `origin`** (ver el aviso de abajo), le pone identidad de
git al agente para que sus commits no salgan con tu nombre, y deja corrida una
verificación inicial.

**Script de mantenimiento** (`scripts/sincronizar.sh`): corre cada vez que se
reanuda un contenedor desde la caché, después de comprobar la rama. Es
exactamente el momento en que el clon está viejo, así que poner acá la
sincronización hace que buena parte del trabajo se haga sola. El `|| true` es
a propósito: si hay un conflicto, que quede anotado en el log pero que no
tumbe el arranque del contenedor — el agente lo resuelve después.

> **Aviso importante: el contenedor de Codex clona el repositorio sin dejar
> configurado el remoto `origin`.** Por eso un `git fetch origin` pelado
> falla con `'origin' does not appear to be a git repository`, aunque el
> acceso a internet esté bien. Los dos scripts lo configuran solos apuntando
> a `https://github.com/italocarreran/Abby.git`. Como el repositorio es
> **público**, el fetch anónimo alcanza: no hacen falta tokens ni
> credenciales en el entorno. Si algún día pasa a privado, esto deja de
> funcionar y hay que agregar un token como secreto del entorno.

---

## 5. Paso 3 — acceso a internet: esta es la llave del asunto

En la misma pantalla del entorno, sección **Internet access** (o "Advanced
settings" según la versión):

| Campo | Qué poner |
|---|---|
| Acceso a Internet del agente | **Activado** |
| Lista de dominios permitidos | `Dependencias comunes` (o la más chica que te deje) |
| Dominios permitidos adicionales | `github.com, codeload.github.com, objects.githubusercontent.com` |
| Métodos HTTP permitidos | **Todos los métodos** |

**Lo de los métodos no es opcional y el desplegable solo ofrece dos opciones:**
"Todos los métodos" o "GET, HEAD y OPTIONS". Git sobre HTTPS negocia el `fetch`
con un `POST`, así que la segunda opción no sirve — hay que dejar **Todos los
métodos**.

Sobre el cartel de **RIESGO ELEVADO** que muestra esa pantalla: es real y vale
la pena entenderlo, no ignorarlo. Los riesgos que enumera son inyección de
instrucciones desde contenido que el agente traiga de la red, y filtración de
código o secretos hacia afuera. En este repositorio en particular el segundo
pesa poco, porque **el repositorio ya es público** y `__config__/` (donde
viven las rutas y datos reales) está gitignoreado y nunca se sube. Lo que sí
conviene: mantener la lista de dominios lo más corta que la interfaz permita,
y mirar el registro de actividad de Codex si algo sale raro.

Si el acceso a internet queda en **Off**, nada de lo de abajo funciona: el
contenedor queda incomunicado y la única forma de traer novedades vuelve a ser
abrir una tarea nueva.

---

## 6. Paso 4 — la regla que hace el trabajo (ya está en el repo)

`REGLAS.md` ahora dice, en el punto 3, que **al empezar cada turno dentro de una
tarea ya abierta** —no solo al abrirla— hay que correr:

```
bash scripts/sincronizar.sh
```

Ese script hace `fetch` de `main`, te muestra qué commits llegaron y de quién,
fusiona, y avisa si llegó trabajo del otro asistente. Como está escrito en
`REGLAS.md` y `AGENTS.md`, Codex lo lee al empezar y no se lo tenés que
recordar.

Si alguna vez ves que no lo corrió, alcanza con escribirle:

> Corré `bash scripts/sincronizar.sh` antes de seguir.

---

## 7. Paso 5 — arrancar cada tarea sobre la rama correcta

En el cuadro donde escribís la tarea, al lado del nombre del repositorio hay un
selector de rama. Por defecto usa la base del entorno (`main`).

- Trabajo nuevo, desde cero → dejalo en `main`.
- **Continuar algo que Claude dejó en una rama** → elegí esa rama ahí. Es la
  diferencia entre que Codex vea el trabajo de Claude o no.

---

## 8. Paso 6 — cómo entrega Codex

El agente no puede hacer `git push` desde el contenedor: la entrega sale por el
botón de la tarea. Cuando termina, usá **Create PR** (o **Push to branch** si
la versión te lo ofrece) apuntando a `main`.

De ahí en adelante el PR es el lugar compartido: Claude lo lee, lo revisa y lo
fusiona. No hace falta que le pases nada por chat.

---

## 9. Paso 7 — el atajo: `@codex` en el PR, sin chat

Con la GitHub App instalada (paso 1), podés comentar dentro de un PR o un issue:

```
@codex revisá este PR
@codex corregí el manejo de rutas en revisor/lectores.py
```

Codex arranca una tarea nueva **desde el estado actual de ese PR**, no desde un
clon viejo. Es decir: siempre al día, por definición, sin abrir ningún chat.

Para el ciclo que querés —Claude planifica, Codex ejecuta, Claude revisa— este
es el flujo más limpio, porque el PR es el único lugar donde los dos miran lo
mismo.

---

## 10. Cómo queda el ciclo

| Momento | Quién | Qué pasa |
|---|---|---|
| Planificar | Claude | Deja el plan en `BITACORA.md`, en `main` |
| Ejecutar | Codex | Tarea nueva desde `main` → trabaja → **Create PR** hacia `main` |
| Revisar | Claude | Lee el PR, corrige si hace falta, fusiona a `main` |
| Corregir | Codex | `@codex ...` en el PR, o en la tarea abierta corriendo `scripts/sincronizar.sh` primero |
| Cerrar | quien haya tocado | Entrada en `BITACORA.md`, `bash scripts/verificar.sh` en verde |

El chat de Codex se abre una vez por *cambio*, no una vez por *actualización*.
Esa es la diferencia entre el flujo de ahora y el de después.

---

## 11. Lo que sigue sin poder hacerse

Vale la pena tenerlo claro para no perseguir una opción que no existe:

- **Una tarea de Codex no se entera sola de nada.** No hay notificaciones ni
  sincronización automática. La actualización siempre la dispara alguien: vos
  pidiéndola, o Codex por la regla de `REGLAS.md`.
- **Los dos asistentes no se ven en tiempo real.** El único canal entre ellos
  es lo que quede escrito en `main`: los commits y `BITACORA.md`. Por eso
  `REGLAS.md` insiste tanto en dejar todo anotado antes de cerrar.
- **Codex no fusiona a `main`.** Entrega un PR; fusionar lo hacés vos o Claude.
- **Los commits de Codex pueden salir con tu nombre.** El setup script intenta
  arreglarlo, pero si igual pasa, el mensaje del commit empieza con `codex:`
  (`REGLAS.md`, punto 1) para que se sepa quién lo hizo.

---

## 12. Si algo no funciona

| Síntoma | Causa casi segura |
|---|---|
| `'origin' does not appear to be a git repository` | El contenedor clona sin remoto. Los scripts ya lo configuran solos; si ves esto, el entorno está usando una versión vieja de `main` — sincronizá |
| `scripts/sincronizar.sh` falla en el `fetch` | Acceso a internet desactivado, o los métodos HTTP en "GET, HEAD y OPTIONS" en vez de "Todos los métodos" (paso 3) |
| Codex no ve un commit reciente | La tarea arrancó antes del commit y nadie corrió `sincronizar.sh` |
| Codex arranca de una base rara | La rama por defecto del repositorio no es `main` (paso 0) |
| El PR sale contra la rama equivocada | El entorno tiene otra Base branch configurada (paso 2) |
| `@codex` no responde en un PR | La GitHub App no está instalada en el repositorio (paso 1) |
