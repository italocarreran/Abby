#!/usr/bin/env bash
# Trae la rama base al día DENTRO de una tarea ya empezada.
#
#   bash scripts/sincronizar.sh          # contra la rama base
#   bash scripts/sincronizar.sh otra     # contra otra rama
#
# Para qué existe: una tarea de Codex clona el repositorio una sola vez, al
# crearse. Todo lo que se suba después es invisible para esa tarea hasta que
# alguien haga fetch. Este script es ese fetch, más el merge, más el resumen
# de qué llegó y de quién. Se corre al empezar CADA turno, no solo al abrir
# la tarea.

set -uo pipefail
cd "$(dirname "$0")/.."

# ---------------------------------------------------------------------------
# CONFIGURACIÓN — ajustar estas dos líneas al instalar en un repositorio
# ---------------------------------------------------------------------------
URL_REPO="https://github.com/OWNER/REPO.git"
RAMA_BASE_POR_DEFECTO="main"
# SOLO para repositorios privados: nombre de la variable de entorno donde el
# entorno de Codex guarda el token de GitHub (Secrets del environment).
# Si el repositorio es público hay que DEJAR ESTO VACÍO: un token suelto en el
# contenedor puede no servir para este repositorio y romper un fetch anónimo
# que funcionaba bien. Se activa a propósito, nunca por detección automática.
VAR_TOKEN=""
# ---------------------------------------------------------------------------

BASE="${1:-$RAMA_BASE_POR_DEFECTO}"

# El contenedor de Codex clona el repositorio SIN dejar configurado el remoto
# 'origin'. Por eso no se puede dar por sentado: se resuelve acá.
if REMOTO=$(git remote get-url origin 2>/dev/null) && [ -n "$REMOTO" ]; then
    :
else
    REMOTO="$URL_REPO"
    git remote add origin "$URL_REPO" 2>/dev/null
    echo "aviso: no había remoto 'origin' configurado (es lo normal en un"
    echo "       contenedor de Codex). Se configuró $URL_REPO"
    echo
fi

# Si hay token en el entorno, se inyecta solo para esta llamada. Nunca se
# imprime: lo que se muestra siempre es la URL sin credenciales.
REMOTO_VISIBLE="$REMOTO"
TOKEN=""
if [ -n "$VAR_TOKEN" ]; then TOKEN="${!VAR_TOKEN:-}"; fi
if [ -n "$TOKEN" ]; then
    REMOTO="${REMOTO/https:\/\//https://x-access-token:${TOKEN}@}"
    REMOTO_VISIBLE="$REMOTO_VISIBLE (con token de \$$VAR_TOKEN)"
fi

echo "== fetch $BASE desde $REMOTO_VISIBLE =="
if ! git fetch "$REMOTO" "$BASE" 2>&1 | sed "s/${TOKEN:-__nada__}/***/g"; then
    echo
    echo "No se pudo hacer fetch. Causas probables, en orden:"
    echo "  1. El entorno de Codex tiene el acceso a internet del agente"
    echo "     desactivado, o los métodos HTTP en 'GET, HEAD y OPTIONS':"
    echo "     git sobre HTTPS necesita POST para negociar el fetch."
    echo "  2. github.com no está en la lista de dominios permitidos."
    echo "  3. La rama base '$BASE' no existe en el remoto."
    echo "  4. El repositorio es privado y no hay token: hay que cargarlo"
    echo "     como secreto del entorno en la variable \$$VAR_TOKEN."
    exit 1
fi

# git deja lo recién traído en FETCH_HEAD. Se usa eso y no 'origin/<rama>'
# porque en un clon sin remoto esa referencia no existe.
echo
echo "== Commits nuevos en $BASE que esta copia no tenía =="
nuevos=$(git rev-list --count HEAD..FETCH_HEAD)
if [ "$nuevos" -eq 0 ]; then
    echo "  ninguno: esta copia ya estaba al día"
else
    git log --format='  %h | %an | %ad | %s' --date=short HEAD..FETCH_HEAD
fi

echo
echo "== Integrando =="
if git merge --no-edit FETCH_HEAD; then
    echo "  ok"
else
    echo
    echo "CONFLICTO. No resolverlo a ciegas: mirar quién tocó el archivo"
    echo "  git log --format='%h | %an | %s' HEAD..FETCH_HEAD -- <archivo>"
    echo "y leer la bitácora del repositorio antes de decidir."
    exit 1
fi

if [ "$nuevos" -ne 0 ]; then
    echo
    echo "== Archivos que cambiaron =="
    git diff --stat ORIG_HEAD..HEAD | sed 's/^/  /'
    echo
    echo "Llegó trabajo del otro asistente. Antes de seguir, leer la bitácora"
    echo "del repositorio: puede que lo que ibas a hacer ya esté hecho."
fi
