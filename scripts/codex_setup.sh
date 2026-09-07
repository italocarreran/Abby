#!/usr/bin/env bash
# Setup script del entorno de Codex (Settings -> Codex -> Environments ->
# este repo -> Setup script). Se corre UNA vez, al armar el contenedor,
# cuando todavía hay red aunque después se apague.
#
# Pegar en ese campo, literal:
#     bash scripts/codex_setup.sh
#
# No instala dependencias del proyecto porque no hay: los scripts usan solo
# la biblioteca estándar. Lo único que falta en un contenedor headless es
# tkinter, que test_tema.py necesita.

set -uo pipefail
cd "$(dirname "$0")/.."

echo "== Python =="
python3 --version

echo
echo "== tkinter (para __comun__/test_tema.py) =="
if python3 -c 'import tkinter' 2>/dev/null; then
    echo "  ya estaba"
elif command -v apt-get >/dev/null 2>&1; then
    apt-get update -qq && apt-get install -y -qq python3-tk \
        && echo "  instalado" \
        || echo "  no se pudo instalar; test_tema.py se va a omitir (no es bloqueante)"
else
    echo "  sin apt-get; test_tema.py se va a omitir (no es bloqueante)"
fi

echo
echo "== git: identidad para los commits del agente =="
# Sin esto los commits de Codex quedan con el nombre del usuario y REGLAS.md
# punto 1 (mirar el autor del commit para saber quién tocó qué) deja de servir.
git config user.name  "Codex"
git config user.email "codex@openai.com"
git config --get user.name | sed 's/^/  user.name  = /'
git config --get user.email | sed 's/^/  user.email = /'

echo
echo "== Verificación inicial =="
bash scripts/verificar.sh
