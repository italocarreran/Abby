#!/usr/bin/env bash
# Setup script del entorno de Codex. Se pega tal cual en
#   Settings -> Codex -> Environments -> <repo> -> Script de configuración
# en modo Manual:
#
#     bash scripts/setup_codex.sh
#
# Corre UNA vez, al armar el contenedor. Ahí siempre hay red, aunque el
# acceso a internet del agente esté desactivado — son dos cosas distintas.

set -uo pipefail
cd "$(dirname "$0")/.."

# ---------------------------------------------------------------------------
# CONFIGURACIÓN — ajustar al instalar en un repositorio
# ---------------------------------------------------------------------------
URL_REPO="https://github.com/OWNER/REPO.git"
NOMBRE_AGENTE="Codex"
MAIL_AGENTE="codex@openai.com"
# ---------------------------------------------------------------------------

echo "== git: remoto 'origin' =="
# El contenedor de Codex clona el repositorio SIN dejar remoto configurado, y
# entonces cualquier 'git fetch origin' falla con "does not appear to be a git
# repository" aunque la red esté perfecta. Esto lo deja resuelto de entrada.
if git remote get-url origin >/dev/null 2>&1; then
    git remote get-url origin | sed 's/^/  ya estaba: /'
else
    git remote add origin "$URL_REPO" \
        && echo "  configurado: $URL_REPO" \
        || echo "  no se pudo; scripts/sincronizar.sh lo reintenta solo"
fi

echo
echo "== git: identidad del agente =="
# Sin esto los commits de Codex salen con el nombre del dueño de la cuenta, y
# mirar el autor del commit deja de servir para saber quién tocó qué.
git config user.name  "$NOMBRE_AGENTE"
git config user.email "$MAIL_AGENTE"
echo "  $(git config user.name) <$(git config user.email)>"

# --- Dependencias del proyecto -------------------------------------------
# Agregar acá lo que el repositorio necesite y el contenedor no traiga. Es el
# único momento con red garantizada, así que todo lo pesado va acá y no en el
# script de mantenimiento. Ejemplos:
#   pip install -r requirements.txt
#   npm ci
#   apt-get update -qq && apt-get install -y -qq python3-tk
# -------------------------------------------------------------------------

# --- Verificación inicial ------------------------------------------------
# Dejar corriendo el check propio del repositorio, si existe, para que una
# tarea nueva arranque sabiendo si la base está sana. Ejemplos:
#   bash scripts/verificar.sh
#   npm test
# -------------------------------------------------------------------------
