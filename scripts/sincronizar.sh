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
#
# Si el fetch falla, casi siempre es que el entorno de Codex tiene el acceso a
# internet apagado: Settings -> Codex -> Environments -> este repo -> Internet
# access. Ver docs/FLUJO_CLAUDE_CODEX.md.

set -uo pipefail
cd "$(dirname "$0")/.."

BASE="${1:-main}"

echo "== fetch origin/$BASE =="
if ! git fetch origin "$BASE"; then
    echo
    echo "No se pudo hacer fetch. Causas probables, en orden:"
    echo "  1. El entorno de Codex tiene el acceso a internet apagado."
    echo "     Hay que encenderlo y permitir github.com (métodos GET y POST:"
    echo "     git sobre HTTPS necesita POST para negociar el fetch)."
    echo "  2. La rama base '$BASE' no existe en el remoto."
    echo "     Ver: git ls-remote --heads origin"
    exit 1
fi

echo
echo "== Commits nuevos en origin/$BASE que esta copia no tenía =="
nuevos=$(git rev-list --count "HEAD..origin/$BASE")
if [ "$nuevos" -eq 0 ]; then
    echo "  ninguno: esta copia ya estaba al día"
else
    git log --format='  %h | %an | %ad | %s' --date=short "HEAD..origin/$BASE"
fi

echo
echo "== Integrando =="
if git merge --no-edit "origin/$BASE"; then
    echo "  ok"
else
    echo
    echo "CONFLICTO. No resolverlo a ciegas: mirar quién tocó el archivo"
    echo "  git log --format='%h | %an | %s' HEAD..origin/$BASE -- <archivo>"
    echo "y leer la entrada correspondiente de BITACORA.md antes de decidir."
    exit 1
fi

if [ "$nuevos" -ne 0 ]; then
    echo
    echo "== Archivos que cambiaron =="
    git diff --stat "ORIG_HEAD..HEAD" | sed 's/^/  /'
    echo
    echo "Llegó trabajo del otro asistente. Antes de seguir: leer 'Pendientes"
    echo "abiertos' de BITACORA.md, por si lo que ibas a hacer ya está hecho."
fi
