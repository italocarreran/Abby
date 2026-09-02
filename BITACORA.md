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

---

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
