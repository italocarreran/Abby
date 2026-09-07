#!/usr/bin/env bash
# Trae la rama base al día DENTRO de una tarea ya empezada.
#
#   bash scripts/sincronizar.sh          # contra main
#   bash scripts/sincronizar.sh otra     # contra otra rama base
#
# Para qué existe: una tarea de Codex clona el repositorio una sola vez, al
# crearse. Todo lo que Claude o el usuario suban después es invisible para esa
# tarea hasta que alguien haga fetch. Este script es ese fetch, más el merge,
# más el resumen de qué llegó. Se corre al empezar CADA turno (REGLAS.md 3b),
# no solo al abrir la tarea.

set -uo pipefail
cd "$(dirname "$0")/.."

BASE="${1:-main}"

# El repositorio es público, así que el fetch anónimo por HTTPS alcanza: no
# hacen falta credenciales ni tokens.
URL_PUBLICA="https://github.com/italocarreran/Abby.git"

# El contenedor de Codex clona el repositorio SIN dejar configurado el remoto
# 'origin'. Por eso no se puede dar por sentado: se resuelve acá.
if REMOTO=$(git remote get-url origin 2>/dev/null) && [ -n "$REMOTO" ]; then
    :
else
    REMOTO="$URL_PUBLICA"
    git remote add origin "$URL_PUBLICA" 2>/dev/null
    echo "aviso: no había remoto 'origin' configurado (es lo normal en un"
    echo "       contenedor de Codex). Se configuró $URL_PUBLICA"
    echo
fi

echo "== fetch $BASE desde $REMOTO =="
if ! git fetch "$REMOTO" "$BASE"; then
    echo
    echo "No se pudo hacer fetch. Causas probables, en orden:"
    echo "  1. El entorno de Codex tiene el acceso a internet del agente"
    echo "     apagado, o no permite el método POST: git sobre HTTPS lo"
    echo "     necesita para negociar el fetch. Ver docs/FLUJO_CLAUDE_CODEX.md."
    echo "  2. github.com no está en la lista de dominios permitidos."
    echo "  3. La rama base '$BASE' no existe en el remoto."
    echo "  4. El repositorio dejó de ser público: entonces el fetch anónimo"
    echo "     ya no alcanza y hace falta un token en el entorno."
    exit 1
fi

# git deja lo recién traído en FETCH_HEAD. Se usa eso y no 'origin/$BASE'
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
    echo "y leer la entrada correspondiente de BITACORA.md antes de decidir."
    exit 1
fi

if [ "$nuevos" -ne 0 ]; then
    echo
    echo "== Archivos que cambiaron =="
    git diff --stat ORIG_HEAD..HEAD | sed 's/^/  /'
    echo
    echo "Llegó trabajo del otro asistente. Antes de seguir: leer 'Pendientes"
    echo "abiertos' de BITACORA.md, por si lo que ibas a hacer ya está hecho."
fi
