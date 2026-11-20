#!/usr/bin/env python3
"""Check that everything latex-slides needs is installed. Exit 1 if something is missing."""
from __future__ import annotations

import importlib
import shutil
import subprocess
import sys

TOOLS = {
    "xelatex": "TeX Live / MacTeX / MiKTeX",
    "latexmk": "TeX Live (latexmk)",
    "kpsewhich": "TeX Live",
    "pdftoppm": "poppler-utils (apt: poppler-utils, brew: poppler)",
    "pdftocairo": "poppler-utils",
    "pdftotext": "poppler-utils",
    "pdfseparate": "poppler-utils",
}
OPTIONAL_TOOLS = {"soffice": "LibreOffice (PPTX template import, render check)"}
TEX_PACKAGES = ["beamer.cls", "textpos.sty", "pgfplots.sty", "xeCJK.sty", "fontspec.sty", "tabularx.sty",
                "booktabs.sty", "etoolbox.sty"]
PY_MODULES = {"pptx": "python-pptx", "lxml": "lxml", "yaml": "PyYAML", "pypdf": "pypdf", "PIL": "Pillow",
              "numpy": "numpy"}


def main() -> None:
    bad = False
    print("command-line tools")
    for tool, hint in {**TOOLS, **OPTIONAL_TOOLS}.items():
        found = shutil.which(tool) or (tool == "soffice" and shutil.which("libreoffice"))
        optional = tool in OPTIONAL_TOOLS
        print(f"  {'ok ' if found else ('-- ' if optional else 'MISSING')} {tool:<12} {'' if found else hint}")
        bad |= not found and not optional
    print("TeX packages")
    if shutil.which("kpsewhich"):
        for pkg in TEX_PACKAGES:
            found = subprocess.run(["kpsewhich", pkg], capture_output=True, text=True).stdout.strip()
            print(f"  {'ok ' if found else 'MISSING'} {pkg}")
            bad |= not found
    print("Python modules")
    for mod, pip in PY_MODULES.items():
        try:
            importlib.import_module(mod)
            print(f"  ok  {pip}")
        except ImportError:
            print(f"  MISSING {pip}  (pip install {pip})")
            bad = True
    print("\nall required dependencies found" if not bad else "\nsome required dependencies are missing")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
