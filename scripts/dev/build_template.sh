#!/usr/bin/env bash
# Rebuild templates/<name>/bg/*.pdf from templates/<name>/src/backgrounds.tex.
#   scripts/dev/build_template.sh [default|paper]
set -euo pipefail
name="${1:-default}"
cd -- "$(dirname -- "$0")/../../templates/$name"
mkdir -p bg src/.build
xelatex -interaction=nonstopmode -halt-on-error -output-directory=src/.build src/backgrounds.tex >/dev/null
i=1
for role in cover agenda section body closing; do
    pdfseparate -f $i -l $i src/.build/backgrounds.pdf "bg/$role.pdf"
    i=$((i + 1))
done
rm -rf src/.build
echo "$name: wrote $(ls bg | tr '\n' ' ')"
