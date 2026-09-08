# El environment de Codex, campo por campo

Todo esto vive en **Settings → Codex → Environments → \<repo\>** dentro de
ChatGPT. Los nombres de los campos cambian entre versiones y entre idiomas;
lo que sigue describe qué hace cada uno para que se reconozca aunque el
rótulo no coincida exacto.

---

## Repositorio y rama base

| Campo | Qué poner |
|---|---|
| Repository | el repositorio |
| Base branch | la rama estable (`main`) |

La base branch es de dónde clona **cada tarea nueva**. Si acá hay una rama que
cambia de nombre cada tanto —una rama de sesión de un agente, por ejemplo— no
hay punto fijo contra el cual sincronizar y el resto de la configuración no
sirve de nada. Por eso el paso de la rama estable va primero.

---

## Los dos campos de script

La pantalla tiene **dos** campos, y hacen cosas distintas. Es el error más
fácil de cometer: poner el mismo script en los dos.

### Script de configuración (setup script)

> *"Se ejecuta después de crear nuevos contenedores, una vez clonado el
> repositorio."*

Corre **una sola vez**, al armar el contenedor. Es donde van las dependencias
pesadas y la configuración de git. Poner, en modo **Manual**:

```
bash scripts/setup_codex.sh
```

### Script de mantenimiento (maintenance script)

> *"Se ejecuta en contenedores que se reanudaron desde la caché, después de
> comprobar la rama."*

Corre cada vez que se **reanuda** un contenedor cacheado — que es exactamente
el momento en que el clon quedó viejo. Es el mejor lugar para la
sincronización. Poner, en modo **Manual**:

```
bash scripts/sincronizar.sh || true
```

El `|| true` es a propósito: si hay un conflicto de merge, que quede anotado
en el log pero que no tumbe el arranque del contenedor. El agente lo resuelve
después, con contexto.

### La trampa del cartel

Los dos campos muestran abajo: **"El acceso a la red siempre está habilitado
en este paso"**. Es verdad y es engañoso: vale para *esos* pasos. Cuando el
agente está trabajando es otra fase, gobernada por el ajuste de más abajo. Ver
ese cartel y saltear la configuración de internet es un error común.

### La otra trampa: el script tiene que existir en la rama base

Codex clona la base branch y recién ahí corre el script. Si los scripts todavía
no están fusionados a esa rama, el setup falla con `No such file or directory`.
Los archivos van al repositorio **antes** de configurar el environment.

---

## Acceso a Internet del agente

| Campo | Qué poner |
|---|---|
| Acceso a Internet del agente | **Activado** |
| Lista de dominios permitidos | el preset más chico que alcance |
| Dominios permitidos adicionales | `github.com, codeload.github.com, objects.githubusercontent.com` |
| Métodos HTTP permitidos | **Todos los métodos** |

**Lo de los métodos no es negociable.** El desplegable ofrece solo dos
opciones: "Todos los métodos" o "GET, HEAD y OPTIONS". Git sobre HTTPS negocia
el `fetch` con un `POST`, así que la segunda no sirve: el fetch falla y se
vuelve al problema original.

### Sobre el cartel de RIESGO ELEVADO

La pantalla avisa que habilitar internet expone el entorno a inyección de
instrucciones desde contenido traído de la red, y a filtración de código o
secretos. Es real y conviene decirlo en vez de minimizarlo.

Qué pesa cuánto, en la práctica:

- **Filtración**: pesa poco si el repositorio ya es público y los datos
  sensibles (rutas reales, credenciales, planillas) están gitignoreados y
  nunca se suben. Pesa mucho si el repositorio es privado y tiene secretos.
- **Inyección**: pesa según qué trae el agente de la red. Con la lista de
  dominios acotada a GitHub, la superficie es chica.

Mitigación concreta: dominios lo más acotados que la interfaz permita, y
revisar el registro de actividad de Codex si algo sale raro. Si el repositorio
es privado y sensible, vale la pena que el usuario tome la decisión a
conciencia, no por inercia.

---

## Repositorios privados

Con repositorio público el `fetch` anónimo por HTTPS alcanza y no hace falta
nada más. Con repositorio privado, no: hay que cargar un token de GitHub como
**secreto del environment** y decirle al script cómo se llama la variable.

1. Crear un Personal Access Token en GitHub con permiso de lectura sobre el
   repositorio (`repo` en los clásicos; `Contents: Read` en los de grano fino).
2. Cargarlo como secreto del environment de Codex, por ejemplo `GITHUB_TOKEN`.
3. En `scripts/sincronizar.sh`, poner ese nombre en `VAR_TOKEN`.

El script lo inyecta solo para la llamada al `fetch` y redacta el token de
todo lo que imprime.

**Por qué `VAR_TOKEN` viene vacío por defecto:** si el contenedor tiene alguna
variable de token suelta que no sirve para ese repositorio, inyectarla rompe un
fetch anónimo que andaba bien. Se activa a propósito, nunca por detección
automática.

---

## Cómo entrega Codex

El agente **no puede hacer `git push`** desde el contenedor. La entrega sale por
el botón de la tarea (Create PR, o Push to branch según la versión). Eso no se
configura: es así.

## El atajo que evita el chat

Con la GitHub App de Codex instalada en el repositorio, se puede comentar dentro
de un PR o un issue:

```
@codex revisá este PR
@codex corregí el manejo de rutas en <archivo>
```

Codex arranca una tarea **desde el estado actual de ese PR**, no desde un clon
viejo. Es decir: siempre al día por definición, sin abrir ningún chat. Para el
ciclo "un asistente planifica, otro ejecuta, el primero revisa" es el flujo más
limpio, porque el PR es el único lugar donde los dos miran lo mismo.
