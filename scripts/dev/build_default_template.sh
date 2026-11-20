#!/usr/bin/env bash
# Rebuild templates/default/bg/*.pdf from templates/default/src/backgrounds.tex.
set -euo pipefail
cd -- "$(dirname -- "$0")/../../templates/default"
mkdir -p bg src/.build
xelatex -interaction=nonstopmode -halt-on-error -output-directory=src/.build src/backgrounds.tex >/dev/null
i=1
for role in cover agenda section body closing; do
    pdfseparate -f $i -l $i src/.build/backgrounds.pdf "bg/$role.pdf"
    i=$((i + 1))
done
rm -rf src/.build
echo "wrote $(ls bg | tr '\n' ' ')"
