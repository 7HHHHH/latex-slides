#!/usr/bin/env python3
"""Build a latex-slides deck: deck.tex -> PDF -> per-page SVG + PNG -> PPTX.

  build.py <deck_dir> [--pdf-only] [--raster-only] [--png-width 1920]

Outputs go to <deck_dir>/output/<output_name>[_vN].{pdf,pptx,notes.md}; existing
files are never overwritten. build/manifest.json describes the build for check.py.
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pypdf import PdfReader

from lib.common import (EMU_PER_BP, SKILL_ROOT, die, dump_json, lang_key, load_yaml, need, next_version,
                        parse_lsmeta, run, write_ls_config)

SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
LABELS = {"en": ("Slide", "Suggested time", "s", "Sources", "(continued)"),
          "zh": ("第 {n} 页", "建议用时", " 秒", "资料依据", "（续）"),
          "hant": ("第 {n} 頁", "建議用時", " 秒", "資料依據", "（續）"),
          "ja": ("スライド {n}", "目安時間", " 秒", "出典", "（続き）"),
          "ko": ("슬라이드 {n}", "권장 시간", "초", "출처", "(계속)")}
FULLWIDTH_COLON = ("zh", "hant", "ja")


def compile_pdf(deck: Path) -> Path:
    env = dict(os.environ)
    env["TEXINPUTS"] = f".:{SKILL_ROOT / 'latex'}//:" + env.get("TEXINPUTS", "")
    res = subprocess.run(["latexmk", "-xelatex", "-interaction=nonstopmode", "-halt-on-error", "-outdir=build",
                          "deck.tex"], cwd=deck, env=env, text=True, capture_output=True)
    if res.returncode != 0:
        log = deck / "build" / "deck.log"
        lines = log.read_text(encoding="utf-8", errors="replace").splitlines() if log.exists() else []
        errs = [("\n".join(lines[i:i + 4])) for i, ln in enumerate(lines) if ln.startswith("!")]
        die("LaTeX failed:\n" + ("\n\n".join(errs[:5]) or (res.stdout + res.stderr)[-3000:])
            + f"\n(full log: {log})")
    return deck / "build" / "deck.pdf"


def render_pages(pdf: Path, build: Path, n: int, png_width: int, svg: bool) -> None:
    png_dir, svg_dir = build / "png", build / "svg"
    for d in (png_dir, svg_dir):
        shutil.rmtree(d, ignore_errors=True)
        d.mkdir(parents=True)
    run(["pdftoppm", "-png", "-scale-to-x", str(png_width), "-scale-to-y", "-1", str(pdf), str(png_dir / "p")])
    for i, f in enumerate(sorted(png_dir.glob("p-*.png")), 1):
        f.rename(png_dir / f"slide-{i:02d}.png")
    if svg:
        for i in range(1, n + 1):
            run(["pdftocairo", "-svg", "-f", str(i), "-l", str(i), str(pdf), str(svg_dir / f"slide-{i:02d}.svg")])


def page_text(pdf: Path, i: int) -> str:
    res = run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", "-enc", "UTF-8", str(pdf), "-"])
    return " ".join(res.stdout.split())


def page_links(reader: PdfReader, i: int, height: float) -> list[dict]:
    links = []
    for annot in reader.pages[i - 1].get("/Annots") or []:
        a = annot.get_object()
        if a.get("/Subtype") != "/Link":
            continue
        uri = (a.get("/A") or {}).get("/URI")
        if not uri or not str(uri).startswith(("http://", "https://", "mailto:")):
            continue
        x0, y0, x1, y1 = (float(v) for v in a["/Rect"])
        links.append({"uri": str(uri), "x": min(x0, x1), "y": height - max(y0, y1),
                      "w": abs(x1 - x0), "h": abs(y1 - y0)})
    return links


def notes_for(slide: dict, n: int, continued: bool, lang: str) -> str:
    page_lbl, time_lbl, unit, src_lbl, cont = LABELS[lang]
    head = page_lbl.format(n=f"{n:02d}") if "{n}" in page_lbl else f"{page_lbl} {n:02d}"
    if continued:
        return f"{head}  {slide.get('title', '')} {cont}"
    colon = "：" if lang in FULLWIDTH_COLON else ": "
    parts = [f"{head}  {slide.get('title', '')}", f"{time_lbl}{colon}{slide.get('seconds', 0)}{unit}",
             "", (slide.get("script") or "").strip()]
    if slide.get("sources"):
        parts += ["", f"{src_lbl}{colon.strip()}"] + [f"- {s}" for s in slide["sources"]]
    return "\n".join(parts).strip()


def add_svg(slide, pic, svg_path: Path, idx: int) -> None:
    pkg = slide.part.package
    part = Part(PackURI(f"/ppt/media/ls-slide-{idx:02d}.svg"), "image/svg+xml", pkg, svg_path.read_bytes())
    rid = slide.part.relate_to(part, RT.IMAGE)
    blip = pic._element.find(".//" + qn("a:blip"))
    ext_lst = blip.find(qn("a:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(blip, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"), uri=SVG_EXT_URI)
    svg_blip = etree.SubElement(ext, f"{{{ASVG_NS}}}svgBlip", nsmap={"asvg": ASVG_NS})
    svg_blip.set(qn("r:embed"), rid)


def add_link(slide, link: dict) -> None:
    e = lambda v: int(round(v * EMU_PER_BP))
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, e(link["x"]), e(link["y"]), e(link["w"]), e(link["h"]))
    shp.name = f"Link: {link['uri']}"
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    clr = shp._element.spPr.find(qn("a:solidFill"))[0]
    etree.SubElement(clr, qn("a:alpha"), val="0")       # invisible but clickable
    shp.line.fill.background()
    shp.shadow.inherit = False
    shp.click_action.hyperlink.address = link["uri"]


def resolve(deck: Path, name: str) -> Path:
    for cand in (deck / name, deck / "assets" / name):
        if cand.exists():
            return cand
    die(f"file referenced in deck.tex not found: {name}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck_dir", type=Path)
    ap.add_argument("--pdf-only", action="store_true", help="compile the PDF and stop")
    ap.add_argument("--raster-only", action="store_true", help="PNG slides only (no SVG)")
    ap.add_argument("--png-width", type=int, help="width of PNG fallback images in pixels")
    a = ap.parse_args()

    need("latexmk", "xelatex", "pdftoppm", "pdftocairo", "pdftotext")
    deck = a.deck_dir.resolve()
    cfg = load_yaml(deck / "deck.yaml")
    lang = lang_key(cfg.get("language", "en"))
    bcfg = cfg.get("build") or {}
    use_svg = bool(bcfg.get("svg", True)) and not a.raster_only
    png_width = a.png_width or int(bcfg.get("png_width", 1920))

    write_ls_config(deck, cfg.get("template", "template"), cfg.get("language", "en"))
    pdf = compile_pdf(deck)
    build = deck / "build"
    print(f"compiled {pdf.relative_to(deck)}")
    if a.pdf_only:
        return

    reader = PdfReader(str(pdf))
    n = len(reader.pages)
    box = reader.pages[0].mediabox
    width_bp, height_bp = float(box.width), float(box.height)
    meta = parse_lsmeta(build / "deck.lsmeta")

    by_id = {}
    for s in cfg.get("slides") or []:
        if s.get("id") in by_id:
            die(f"duplicate slide id in deck.yaml: {s.get('id')}")
        by_id[s.get("id")] = s
    tex_ids, pages = [], []
    for i in range(1, n + 1):
        m = meta.get(i, {"role": "body", "id": None, "visuals": 0, "videos": []})
        if not m["id"]:
            die(f"page {i} has no \\slideid{{...}}; every frame needs one")
        if m["id"] not in by_id:
            die(f"slide id '{m['id']}' (page {i}) is missing from deck.yaml")
        continued = bool(tex_ids) and tex_ids[-1] == m["id"]
        if not continued and m["id"] in tex_ids:
            die(f"slide id '{m['id']}' is used by two different frames")
        tex_ids.append(m["id"])
        pages.append({"page": i, **m, "continued": continued})
    unused = [k for k in by_id if k not in tex_ids]
    if unused:
        die("deck.yaml slides not found in deck.tex: " + ", ".join(unused))

    render_pages(pdf, build, n, png_width, use_svg)

    prs = Presentation()
    prs.slide_width, prs.slide_height = int(width_bp * EMU_PER_BP), int(height_bp * EMU_PER_BP)
    blank = prs.slide_layouts[6]
    notes_md = [f"# {cfg.get('title', '')}\n"]
    for p in pages:
        i, s = p["page"], by_id[p["id"]]
        slide = prs.slides.add_slide(blank)
        pic = slide.shapes.add_picture(str(build / "png" / f"slide-{i:02d}.png"), 0, 0,
                                       prs.slide_width, prs.slide_height)
        p["text"] = page_text(pdf, i)
        pic.name = f"Slide {i:02d}"
        pic._element.nvPicPr.cNvPr.set("descr", f"{s.get('title', '')}. {p['text']}"[:1500])
        if use_svg:
            add_svg(slide, pic, build / "svg" / f"slide-{i:02d}.svg", i)
        p["links"] = page_links(reader, i, height_bp)
        for link in p["links"]:
            add_link(slide, link)
        for v in p["videos"]:
            e = lambda val: int(round(val * EMU_PER_BP))
            slide.shapes.add_movie(str(resolve(deck, v["file"])), e(v["x"]), e(v["y"]), e(v["w"]), e(v["h"]),
                                   poster_frame_image=str(resolve(deck, v["poster"])), mime_type="video/mp4")
        notes = notes_for(s, i, p["continued"], lang)
        slide.notes_slide.notes_text_frame.text = notes
        notes_md.append(notes.replace("\n", "\n\n", 1) + "\n")

    out_dir = deck / "output"
    out_dir.mkdir(exist_ok=True)
    name = next_version(out_dir, cfg.get("output_name") or "deck")
    pptx_path, pdf_out = out_dir / f"{name}.pptx", out_dir / f"{name}.pdf"
    prs.save(str(pptx_path))
    shutil.copy2(pdf, pdf_out)
    (out_dir / f"{name}.notes.md").write_text("\n---\n\n".join(notes_md), encoding="utf-8")

    dump_json(build / "manifest.json", {
        "version": name, "pptx": str(pptx_path), "pdf": str(pdf_out), "svg": use_svg,
        "page_size_bp": [width_bp, height_bp], "language": cfg.get("language", "en"), "pages": pages})
    print(f"wrote {pptx_path.relative_to(deck)} ({n} slides, {'SVG+PNG' if use_svg else 'PNG'})")
    print(f"wrote {pdf_out.relative_to(deck)}, {name}.notes.md")
    print("next: python3 scripts/check.py", deck)


if __name__ == "__main__":
    main()
