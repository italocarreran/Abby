#!/usr/bin/env bash
# Corre las verificaciones obligatorias de REGLAS.md (punto 5) de una sola vez.
#
#   bash scripts/verificar.sh
#
# Devuelve 0 si todo pasa. Cualquier asistente (Claude o Codex) tiene que
# poder correr esto antes de dar un cambio por terminado.

set -uo pipefail
cd "$(dirname "$0")/.."

fallas=0

echo "== INTERFACES.md al día =="
if python3 generar_interfaces.py --check; then
    echo "   ok"
else
    echo "   FALLA: correr 'python3 generar_interfaces.py' y commitear el resultado"
    fallas=$((fallas + 1))
fi

echo
echo "== Tests =="
while IFS= read -r archivo; do
    printf '%-45s ' "$archivo"
    salida=$(python3 "$archivo" 2>&1)
    if [ $? -eq 0 ]; then
        echo "ok"
    elif printf '%s' "$salida" | grep -q "No module named 'tkinter'"; then
        # tkinter no está en los contenedores headless (Codex, Claude web).
        # No es una falla del cambio: es el entorno. Se corre en el PC.
        echo "omitido (sin tkinter en este entorno)"
    else
        echo "FALLA"
        printf '%s\n' "$salida" | tail -20 | sed 's/^/      /'
        fallas=$((fallas + 1))
    fi
done < <(find . -name 'test_*.py' -not -path './.git/*' | sort)

echo
if [ "$fallas" -eq 0 ]; then
    echo "TODO OK"
else
    echo "$fallas verificación(es) fallaron"
fi
exit "$fallas"
