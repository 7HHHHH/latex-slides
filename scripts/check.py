#!/usr/bin/env python3
"""Verify the latest build of a latex-slides deck.

  check.py <deck_dir> [--no-render]

Runs every mandatory check, writes verify/<version>.json and prints a summary.
Exit code 1 if any check fails. Warnings do not fail the build but must be
reported to the user.
"""
from __future__ import annotations

import argparse
import re
import unicodedata
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation

from lib.common import die, dump_json, load_template, load_yaml, need, resolve_template, run

CJK_RE = re.compile(r"[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af\uf900-\ufaff]")
LATIN_WORD_RE = re.compile(r"[A-Za-z][A-Za-z\-']*")
CONTENT_ROLES = {"body"}


class Report:
    def __init__(self):
        self.items = []

    def add(self, name: str, ok: bool, detail: str = "", level: str = "error"):
        self.items.append({"check": name, "status": "pass" if ok else ("fail" if level == "error" else "warn"),
                           "detail": detail})

    @property
    def failed(self):
        return [i for i in self.items if i["status"] == "fail"]


# ------------------------------------------------------------------ helpers
def words_with_boxes(pdf: Path, page: int) -> list[dict]:
    """Words and their boxes (bp, top-left origin) via pdftotext -bbox."""
    res = run(["pdftotext", "-f", str(page), "-l", str(page), "-bbox", "-enc", "UTF-8", str(pdf), "-"])
    out = []
    for m in re.finditer(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>',
                         res.stdout):
        x0, y0, x1, y1 = (float(v) for v in m.groups()[:4])
        out.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1, "text": m.group(5)})
    return out


def inside(w: dict, r: dict, tol: float = 3.0) -> bool:
    return (w["x0"] >= r["x"] - tol and w["y0"] >= r["y"] - tol and
            w["x1"] <= r["x"] + r["w"] + tol and w["y1"] <= r["y"] + r["h"] + tol)


def _punct_only(text: str) -> bool:
    return bool(text) and all(unicodedata.category(ch).startswith("P") for ch in text)


def overlapping(words: list[dict], ratio: float = 0.3) -> tuple[str, str] | None:
    """First pair of words whose boxes overlap by more than `ratio` of the smaller box.
    Punctuation-only words are skipped: CJK punctuation compression legitimately overlaps glyph boxes."""
    words = [w for w in words if not _punct_only(w["text"])]
    for i, a in enumerate(words):
        for b in words[i + 1:]:
            w = min(a["x1"], b["x1"]) - max(a["x0"], b["x0"])
            h = min(a["y1"], b["y1"]) - max(a["y0"], b["y0"])
            if w <= 0 or h <= 0:
                continue
            small = min((a["x1"] - a["x0"]) * (a["y1"] - a["y0"]), (b["x1"] - b["x0"]) * (b["y1"] - b["y0"]))
            if small > 0 and w * h > ratio * small:
                return a["text"], b["text"]
    return None


def text_units(words: list[str]) -> float:
    text = " ".join(words)
    return len(CJK_RE.findall(text)) + 1.6 * len(LATIN_WORD_RE.findall(text))


def render_pptx(pptx: Path, workdir: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    profile = workdir / "lo-profile"
    cmd = [soffice, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf",
           "--outdir", str(workdir), str(pptx)]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600, check=True)
    except (subprocess.SubprocessError, OSError):
        return None
    out = workdir / (pptx.stem + ".pdf")
    return out if out.exists() else None


def raster(pdf: Path, prefix: Path, width: int = 480) -> list[Path]:
    run(["pdftoppm", "-png", "-scale-to-x", str(width), "-scale-to-y", "-1", str(pdf), str(prefix)])
    return sorted(prefix.parent.glob(prefix.name + "-*.png"))


def mean_abs_diff(a: Path, b: Path) -> float:
    import numpy as np
    from PIL import Image
    ia = Image.open(a).convert("RGB")
    ib = Image.open(b).convert("RGB").resize(ia.size)
    return float(np.abs(np.asarray(ia, dtype=np.int16) - np.asarray(ib, dtype=np.int16)).mean())


# ------------------------------------------------------------------ main
def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck_dir", type=Path)
    ap.add_argument("--no-render", action="store_true", help="skip the LibreOffice PPTX-vs-PDF comparison")
    a = ap.parse_args()
    need("pdftotext", "pdftoppm")

    deck = a.deck_dir.resolve()
    cfg = load_yaml(deck / "deck.yaml")
    man_path = deck / "build" / "manifest.json"
    if not man_path.exists():
        die("build/manifest.json not found; run scripts/build.py first")
    import json
    man = json.loads(man_path.read_text(encoding="utf-8"))
    tpl = load_template(resolve_template(deck, cfg.get("template", "template")))
    checks = {"max_text_units": 150, "min_text_units": 25, "duration_tolerance": 0.10,
              "render_diff_threshold": 12, "max_svg_mb": 5, **(cfg.get("checks") or {})}
    slides = {s["id"]: s for s in cfg.get("slides") or []}
    pdf, pptx = Path(man["pdf"]), Path(man["pptx"])
    pages = man["pages"]
    r = Report()

    # 1 LaTeX log
    log = (deck / "build" / "deck.log").read_text(encoding="utf-8", errors="replace")
    missing = sorted(set(re.findall(r"Missing character: There is no (.+?) in font", log)))
    r.add("latex.missing_glyphs", not missing, ", ".join(missing[:20]))
    overfull = re.findall(r"^(Overfull \\[hv]box \(([\d.]+)pt too \w+\).*?)$", log, flags=re.M)
    big = [m[0] for m in overfull if float(m[1]) > 2.0]
    small = [m[0] for m in overfull if float(m[1]) <= 2.0]
    r.add("latex.overfull_boxes", not big, "; ".join(big[:10]))
    if small:
        r.add("latex.overfull_minor", False, f"{len(small)} box(es) overfull by <= 2pt", level="warn")
    undefined = re.findall(r"(?:Reference|Citation) `([^']+)' .*undefined", log)
    r.add("latex.undefined_refs", not undefined, ", ".join(sorted(set(undefined))[:20]))

    # 2 structure & timing
    r.add("structure.page_count", len(pages) == len(Presentation(str(pptx)).slides),
          f"{len(pages)} PDF pages")
    total = sum(int(slides[p["id"]].get("seconds") or 0) for p in pages if not p["continued"])
    target = int(cfg.get("duration_seconds") or 0)
    tol = float(checks["duration_tolerance"])
    ok = target == 0 or abs(total - target) <= tol * target
    r.add("structure.duration", ok, f"slides sum to {total}s, target {target}s (±{int(tol * 100)}%)")
    no_script = [p["id"] for p in pages if not p["continued"] and not (slides[p["id"]].get("script") or "").strip()]
    r.add("notes.scripts_present", not no_script, "missing script: " + ", ".join(no_script), level="warn")
    sourcing = cfg.get("sourcing", "normal")
    if sourcing != "off":
        no_src = [p["id"] for p in pages if p["role"] in CONTENT_ROLES and not p["continued"]
                  and not slides[p["id"]].get("sources")]
        r.add("sources.content_slides", not no_src, "no sources: " + ", ".join(no_src),
              level="error" if sourcing == "strict" else "warn")

    # 3 layout: text must stay inside safe area / title / page-number boxes
    safe, title, num = tpl["safe_area"], tpl["title"], tpl["page_number"]
    page_w, page_h = man["page_size_bp"]
    full = {"x": 0, "y": 0, "w": page_w, "h": page_h}
    outside, dense, sparse, no_visual, overlaps = [], [], [], [], []
    caption_cut = (tpl["sizes"]["caption"] + tpl["sizes"]["detail"]) / 2
    for p in pages:
        words = words_with_boxes(pdf, p["page"])
        if p["role"] in CONTENT_ROLES:
            bad = [w["text"] for w in words if not any(inside(w, b) for b in (safe, title, num))]
        else:
            bad = [w["text"] for w in words if not inside(w, full, tol=0)]
        if bad:
            outside.append(f"p{p['page']}: {' '.join(bad[:8])}")
        hit = overlapping(words)
        if hit:
            overlaps.append(f"p{p['page']}: '{hit[0]}' / '{hit[1]}'")
        if p["role"] in CONTENT_ROLES:
            body = [w for w in words if inside(w, safe) and (w["y1"] - w["y0"]) >= caption_cut * 0.9]
            units = text_units([w["text"] for w in body])
            # density ignores the \lead line and the \takeaway band, which every slide has anyway
            band = (safe["y"] + 48, safe["y"] + safe["h"] - 44)
            inner = text_units([w["text"] for w in body if band[0] <= w["y0"] and w["y1"] <= band[1]])
            p["text_units"] = round(units, 1)
            if units > checks["max_text_units"]:
                dense.append(f"p{p['page']} ({units:.0f} > {checks['max_text_units']})")
            elif inner < checks["min_text_units"] and p["visuals"] < 2 and not p["continued"]:
                sparse.append(f"p{p['page']} ({inner:.0f} < {checks['min_text_units']})")
            if p["visuals"] == 0 and not slides[p["id"]].get("text_only_reason"):
                no_visual.append(f"p{p['page']} ({p['id']})")
    r.add("layout.text_inside_safe_area", not outside, "; ".join(outside[:10]))
    r.add("layout.text_overlap", not overlaps, "; ".join(overlaps[:10]))
    r.add("visual.text_budget", not dense, "too much body text: " + ", ".join(dense))
    r.add("visual.information_density", not sparse,
          "sparse slides (little explanation and a single visual; add numbers, labels or a second visual): "
          + ", ".join(sparse), level="warn")
    r.add("visual.every_content_slide_has_visual", not no_visual,
          "no figure/diagram/chart/table: " + ", ".join(no_visual))

    # 4 SVG hygiene
    if man["svg"]:
        masks, big, texty = [], [], []
        for p in pages:
            f = deck / "build" / "svg" / f"slide-{p['page']:02d}.svg"
            data = f.read_text(encoding="utf-8", errors="replace")
            if "<mask" in data:
                masks.append(p["page"])
            if "<text" in data:
                texty.append(p["page"])
            if f.stat().st_size > checks["max_svg_mb"] * 1e6:
                big.append(f"p{p['page']} {f.stat().st_size / 1e6:.1f}MB")
        r.add("svg.no_live_text", not texty, f"pages with <text>: {texty}")
        r.add("svg.size", not big, ", ".join(big), level="warn")
        r.add("svg.masks", not masks,
              f"soft masks on pages {masks}: fine in PowerPoint 365, may render badly in WPS/older viewers "
              "(flatten transparent images or rebuild with --raster-only)", level="warn")

    # 5 PPTX package
    prs = Presentation(str(pptx))
    with zipfile.ZipFile(pptx) as z:
        names = set(z.namelist())
        bad_zip = z.testzip()
    problems = []
    if bad_zip:
        problems.append(f"corrupt member {bad_zip}")
    n_videos = n_links = 0
    for i, (slide, p) in enumerate(zip(prs.slides, pages), 1):
        xml = etree.tostring(slide._element).decode()
        if man["svg"] and "svgBlip" not in xml:
            problems.append(f"slide {i}: no svgBlip")
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        if not notes.strip():
            problems.append(f"slide {i}: empty notes")
        n_videos += xml.count("<p:video")
        n_links += xml.count("hlinkClick")
        if len(p["videos"]) != xml.count("<p:video"):
            problems.append(f"slide {i}: expected {len(p['videos'])} video(s)")
    with zipfile.ZipFile(pptx) as z:
        rel_files = {n: z.read(n) for n in names if n.startswith("ppt/slides/_rels/")}
    for rel_file, blob in rel_files.items():
        for rel in etree.fromstring(blob):
            if rel.get("TargetMode") == "External":
                continue
            tgt = rel.get("Target")
            norm = _norm(tgt[1:] if tgt.startswith("/") else f"{Path(rel_file).parent.parent.as_posix()}/{tgt}")
            if norm not in names:
                problems.append(f"{rel_file}: missing target {tgt}")
    r.add("pptx.integrity", not problems, "; ".join(problems[:10]))
    expected_links = sum(len(p.get("links", [])) for p in pages)
    r.add("pptx.links", n_links >= expected_links, f"{n_links} hyperlinks, {expected_links} expected")
    for p in pages:
        for v in p["videos"]:
            src = deck / v["file"]
            r.add(f"media.video_p{p['page']}", src.exists() and src.stat().st_size > 0, v["file"])

    # 6 render PPTX with LibreOffice and compare with the PDF
    diffs = []
    if not a.no_render:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            rendered = render_pptx(pptx, td)
            if rendered is None:
                r.add("render.pptx_vs_pdf", False, "skipped: LibreOffice not available or failed", level="warn")
            else:
                pa, pb = raster(pdf, td / "pdf"), raster(rendered, td / "ppt")
                if len(pa) != len(pb):
                    r.add("render.pptx_vs_pdf", False, f"page count {len(pa)} vs rendered {len(pb)}")
                else:
                    diffs = [round(mean_abs_diff(x, y), 2) for x, y in zip(pa, pb)]
                    worst = [f"p{i}={d}" for i, d in enumerate(diffs, 1) if d > checks["render_diff_threshold"]]
                    r.add("render.pptx_vs_pdf", not worst,
                          f"mean abs diff per page (0-255), max {max(diffs)}; over threshold: {', '.join(worst)}")

    report = {"version": man["version"], "pptx": str(pptx), "pdf": str(pdf), "checks": r.items,
              "render_diff": diffs, "pages": [{k: p.get(k) for k in ("page", "id", "role", "visuals", "text_units")}
                                              for p in pages],
              "visual_review": "PENDING: inspect build/png/slide-*.png page by page and record findings here"}
    out = deck / "verify" / f"{man['version']}.json"
    dump_json(out, report)

    width = max(len(i["check"]) for i in r.items)
    for i in r.items:
        mark = {"pass": "PASS", "warn": "WARN", "fail": "FAIL"}[i["status"]]
        print(f"{mark}  {i['check']:<{width}}  {i['detail'] if i['status'] != 'pass' else ''}".rstrip())
    print(f"\nreport: {out}")
    if r.failed:
        print(f"{len(r.failed)} check(s) failed")
        raise SystemExit(1)


def _norm(path: str) -> str:
    parts = []
    for seg in path.split("/"):
        if seg == "..":
            if parts:
                parts.pop()
        elif seg and seg != ".":
            parts.append(seg)
    return "/".join(parts)


if __name__ == "__main__":
    main()
