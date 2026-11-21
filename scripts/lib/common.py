"""Shared helpers: template/deck loading, ls-config generation, metadata parsing."""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import yaml

SKILL_ROOT = Path(__file__).resolve().parents[2]
ROLES = ("cover", "agenda", "section", "body", "closing")
CJK_LANGS = ("zh", "ja", "ko")
EMU_PER_BP = 12700


def die(msg: str) -> None:
    print(f"error: {msg}", file=sys.stderr)
    sys.exit(1)


def warn(msg: str) -> None:
    print(f"warning: {msg}", file=sys.stderr)


def need(*tools: str) -> None:
    missing = [t for t in tools if shutil.which(t) is None]
    if missing:
        die("missing required tools: " + ", ".join(missing))


def run(cmd: list[str], cwd: Path | None = None, env: dict | None = None, quiet: bool = True) -> subprocess.CompletedProcess:
    res = subprocess.run(cmd, cwd=cwd, env=env, text=True,
                         stdout=subprocess.PIPE if quiet else None,
                         stderr=subprocess.STDOUT if quiet else None)
    if res.returncode != 0:
        tail = (res.stdout or "")[-3000:]
        die(f"command failed ({res.returncode}): {' '.join(cmd)}\n{tail}")
    return res


def load_yaml(path: Path) -> dict:
    if not path.exists():
        die(f"{path} not found")
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


# ------------------------------------------------------------------ template
DEFAULT_TEMPLATE = {
    "page_size": {"width_bp": 960, "height_bp": 540},
    "backgrounds": {r: None for r in ROLES},
    "safe_area": {"x": 56, "y": 96, "w": 848, "h": 392},
    "title": {"x": 56, "y": 30, "w": 848, "h": 52},
    "page_number": {"x": 860, "y": 508, "w": 60, "h": 20},
    "palette": {"primary": "#123A5A", "accent": "#3F86D8", "muted": "#4F5964",
                "text": "#111111", "title": "#111111", "rule": "#BFCAD5"},
    "fonts": {
        "latin": {"heading": ["Arial", "TeX Gyre Heros"], "body": ["Times New Roman", "TeX Gyre Termes"]},
        "cjk": {"heading": ["Microsoft YaHei", "Noto Sans CJK SC"], "body": ["SimSun", "Noto Serif CJK SC"]},
        "math": ["TeX Gyre Termes Math", "Latin Modern Math"],
    },
    "sizes": {"title": 24, "heading": 19, "body": 17, "detail": 15, "caption": 12, "number": 12},
}

# CJK font chains for languages other than Simplified Chinese; a template can override any of
# them with fonts.cjk_<key> (e.g. fonts.cjk_ja: {heading: [...], body: [...]}).
CJK_LANG_FONTS = {
    "ja": {"heading": ["Yu Gothic", "Hiragino Sans", "Noto Sans CJK JP", "Source Han Sans JP"],
           "body": ["Yu Mincho", "Hiragino Mincho ProN", "Noto Serif CJK JP", "Noto Sans CJK JP"]},
    "ko": {"heading": ["Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans CJK KR", "Source Han Sans KR"],
           "body": ["Batang", "Noto Serif CJK KR", "Noto Sans CJK KR"]},
    "hant": {"heading": ["Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", "Noto Sans CJK HK"],
             "body": ["PMingLiU", "Songti TC", "Noto Serif CJK TC", "Noto Sans CJK TC"]},
}


def lang_key(language: str) -> str:
    """'en', 'zh' (Simplified), 'hant' (Traditional), 'ja' or 'ko'."""
    low = (language or "en").lower().replace("_", "-")
    base = low.split("-")[0]
    if base == "zh":
        return "hant" if any(t in low for t in ("-tw", "-hk", "-mo", "hant")) else "zh"
    return base if base in ("ja", "ko") else "en"


def resolve_template(deck_dir: Path, value: str | None) -> Path:
    """Template dir from deck.yaml: a path relative to the deck, or the name of a built-in template."""
    value = value or "template"
    cand = (deck_dir / value)
    if (cand / "template.yaml").exists():
        return cand.resolve()
    builtin = SKILL_ROOT / "templates" / value
    if (builtin / "template.yaml").exists():
        return builtin
    die(f"template '{value}' not found (looked in {cand} and {builtin})")


def merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = merge(base[k], v) if isinstance(v, dict) and isinstance(base.get(k), dict) else v
    return out


def load_template(template_dir: Path) -> dict:
    t = merge(DEFAULT_TEMPLATE, load_yaml(template_dir / "template.yaml"))
    bgs = t["backgrounds"]
    if not bgs.get("body"):
        warn("template has no body background; pages will be white")
    for role in ROLES:
        if not bgs.get(role) and role != "body":
            bgs[role] = bgs.get("body")
        if bgs.get(role) and not (template_dir / bgs[role]).exists():
            die(f"template background for '{role}' not found: {template_dir / bgs[role]}")
    return t


def _font_list(v) -> list[str]:
    return [v] if isinstance(v, str) else list(v or [])


def _font_chain(cmd: str, fonts: list[str], opts: str = "") -> str:
    """Nested \\IfFontExistsTF: first installed font wins."""
    if not fonts:
        return ""
    head, rest = fonts[0], fonts[1:]
    use = f"{cmd}{opts}{{{head}}}"
    if not rest:
        return use
    return f"\\IfFontExistsTF{{{head}}}{{{use}}}{{{_font_chain(cmd, rest, opts)}}}"


def is_cjk(language: str) -> bool:
    return (language or "").lower().split("-")[0] in CJK_LANGS


def write_ls_config(deck_dir: Path, template_rel: str, language: str) -> Path:
    tdir = resolve_template(deck_dir, template_rel)
    template_rel = Path(os.path.relpath(tdir, deck_dir)).as_posix()
    t = load_template(tdir)
    for key in ("page_size", "safe_area", "title", "page_number"):   # TeX \numexpr needs integers
        t[key] = {k: int(round(float(v))) for k, v in t[key].items()}
    p, s, ti, pn, sz, pal = (t["page_size"], t["safe_area"], t["title"], t["page_number"], t["sizes"], t["palette"])
    lines = ["% Generated by latex-slides build.py from template.yaml -- do not edit.",
             f"\\def\\lsPageW{{{p['width_bp']}}}\\def\\lsPageH{{{p['height_bp']}}}",
             f"\\def\\lsSafeX{{{s['x']}}}\\def\\lsSafeY{{{s['y']}}}\\def\\lsSafeW{{{s['w']}}}\\def\\lsSafeH{{{s['h']}}}",
             f"\\def\\lsTitleX{{{ti['x']}}}\\def\\lsTitleY{{{ti['y']}}}\\def\\lsTitleW{{{ti['w']}}}\\def\\lsTitleH{{{ti['h']}}}",
             f"\\def\\lsNumX{{{pn['x']}}}\\def\\lsNumY{{{pn['y']}}}\\def\\lsNumW{{{pn['w']}}}"]
    for key in ("title", "heading", "body", "detail", "caption", "number"):
        name = key.capitalize()
        lines.append(f"\\def\\lsSize{name}{{{sz[key]}}}\\def\\lsLead{name}{{{round(sz[key] * 1.42, 1)}}}")
    for key, col in pal.items():
        lines.append(f"\\definecolor{{ls{key.capitalize()}}}{{HTML}}{{{str(col).lstrip('#').upper()}}}")
    for role in ROLES:
        f = t["backgrounds"].get(role)
        path = f"{template_rel}/{f}" if f else ""
        lines.append(f"\\expandafter\\def\\csname lsBg@{role}\\endcsname{{{path}}}")
    latin, cjk = t["fonts"]["latin"], t["fonts"]["cjk"]
    body, heading = _font_list(latin.get("body")) + ["Latin Modern Roman"], _font_list(latin.get("heading")) + ["Latin Modern Sans"]
    lines.append(_font_chain("\\setmainfont", body))
    lines.append(_font_chain("\\setsansfont", body))
    lines.append(_font_chain("\\newfontfamily\\lsHeadingLatin", heading))
    math = _font_list(t["fonts"].get("math")) or ["Latin Modern Math"]
    lines.append("\\RequirePackage{unicode-math}")
    lines.append(_font_chain("\\setmathfont", math + (["Latin Modern Math"] if "Latin Modern Math" not in math else [])))
    if is_cjk(language):
        lines.append("\\RequirePackage{xeCJK}\\xeCJKsetup{CJKspace=true}")
        key = lang_key(language)
        if key != "zh":
            cjk = t["fonts"].get(f"cjk_{key}") or CJK_LANG_FONTS[key]
        cb, ch = _font_list(cjk.get("body")), _font_list(cjk.get("heading"))
        lines.append(_font_chain("\\setCJKmainfont", cb, "[AutoFakeBold=1.6]"))
        lines.append(_font_chain("\\setCJKsansfont", cb, "[AutoFakeBold=1.6]"))
        lines.append(_font_chain("\\newCJKfontfamily\\lsHeadingCJK", ch, "[AutoFakeBold=1.6]"))
        lines.append("\\newcommand{\\lsHeadingFamily}{\\lsHeadingLatin\\lsHeadingCJK}")
    else:
        lines.append("\\newcommand{\\lsHeadingFamily}{\\lsHeadingLatin}")
    out = deck_dir / "ls-config.tex"
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return out


# ------------------------------------------------------------------ lsmeta
PT_TO_BP = 72 / 72.27


def parse_lsmeta(path: Path) -> dict[int, dict]:
    """page -> {role, ids, visuals, videos}"""
    pages: dict[int, dict] = {}

    def page(n: int) -> dict:
        return pages.setdefault(n, {"role": "body", "id": None, "visuals": 0, "videos": []})

    if not path.exists():
        die(f"{path} not found (does deck.tex load latex-slides?)")
    for raw in path.read_text(encoding="utf-8").splitlines():
        parts = raw.split()
        if len(parts) < 2 or not parts[1].isdigit():
            continue
        kind, n = parts[0], int(parts[1])
        if kind == "page":
            page(n)["role"] = parts[2]
        elif kind == "id":
            page(n)["id"] = parts[2]
        elif kind == "visual":
            page(n)["visuals"] += 1
        elif kind == "video" and len(parts) >= 8:
            x, y, w = (float(v) for v in parts[2:5])
            h = float(parts[5].removesuffix("pt")) * PT_TO_BP
            page(n)["videos"].append({"x": x, "y": y, "w": w, "h": h, "poster": parts[6], "file": parts[7]})
    return pages


def next_version(out_dir: Path, stem: str) -> str:
    """'stem', then 'stem_v2', 'stem_v3', ... -- never overwrite a delivered file."""
    name, v = stem, 1
    while any((out_dir / f"{name}{ext}").exists() for ext in (".pptx", ".pdf")):
        v += 1
        name = f"{stem}_v{v}"
    return name


def dump_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def slugify(text: str) -> str:
    return re.sub(r"[^\w\-]+", "_", text, flags=re.UNICODE).strip("_") or "deck"
