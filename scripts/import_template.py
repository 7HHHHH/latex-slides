#!/usr/bin/env python3
"""Import a user template into a latex-slides template directory.

Sources (pick one style):
  --pptx T.pptx                 render every slide layout of a PowerPoint template
  --pptx T.pptx --mode slides   render the template's existing slides (use --map)
  --pdf  T.pdf  --map cover=1,body=2,closing=3
  --cover a.pdf --body b.png ... one file per role (PDF page 1, PNG or JPG)

Then:
  out_dir/template.yaml   geometry, palette and fonts (review and edit!)
  out_dir/bg/<role>.pdf|png
  out_dir/preview.png     backgrounds with safe area (green), title (blue), page number (red)

Example:
  import_template.py my_template --pptx lab.pptx
  import_template.py my_template --pdf backgrounds.pdf --map cover=1,agenda=2,body=3,closing=4 \
      --safe 56,96,848,392 --title 56,30,848,52
"""
from __future__ import annotations

import argparse
import copy
import shutil
import subprocess
import tempfile
from pathlib import Path

import yaml
from lxml import etree
from PIL import Image, ImageDraw
from pypdf import PdfReader

from lib.common import DEFAULT_TEMPLATE, EMU_PER_BP, ROLES, die, need, run, warn

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NAME_HINTS = {
    "cover": ["title slide", "标题幻灯片", "封面", "cover", "title page"],
    "section": ["section", "节标题", "章节", "过渡"],
    "agenda": ["agenda", "目录", "contents", "outline"],
    "closing": ["closing", "end", "thank", "结束", "致谢", "谢谢"],
    "body": ["title and content", "标题和内容", "title only", "仅标题", "content", "正文", "内容"],
}


def parse_box(s: str | None) -> dict | None:
    if not s:
        return None
    x, y, w, h = (round(float(v)) for v in s.split(","))
    return {"x": x, "y": y, "w": w, "h": h}


def parse_map(s: str | None) -> dict[str, int]:
    out = {}
    for item in filter(None, (s or "").split(",")):
        role, _, num = item.partition("=")
        if role not in ROLES:
            die(f"unknown role '{role}' (use {', '.join(ROLES)})")
        out[role] = int(num)
    return out


def soffice_to_pdf(src: Path, outdir: Path) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        die("LibreOffice (soffice) is required to import PPTX templates")
    profile = outdir / "lo-profile"
    subprocess.run([soffice, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf",
                    "--outdir", str(outdir), str(src)], check=True, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=600)
    pdf = outdir / (src.stem + ".pdf")
    if not pdf.exists():
        die(f"LibreOffice failed to convert {src}")
    return pdf


def extract_page(pdf: Path, page: int, dest: Path) -> None:
    n = len(PdfReader(str(pdf)).pages)
    if not 1 <= page <= n:
        die(f"{pdf} has {n} pages; page {page} requested")
    run(["pdfseparate", "-f", str(page), "-l", str(page), str(pdf), str(dest)])


def guess_role(name: str) -> str | None:
    low = name.lower()
    for role in ("cover", "section", "agenda", "closing", "body"):
        if any(h in low for h in NAME_HINTS[role]):
            return role
    return None


# ------------------------------------------------------------------ PPTX
def pptx_layout_deck(src: Path, tmp: Path):
    """Copy of the template with all slides removed and one empty slide per layout."""
    from pptx import Presentation
    prs = Presentation(str(src))
    sld_ids = prs.slides._sldIdLst
    for sld in list(sld_ids):
        prs.part.drop_rel(sld.rId)
        sld_ids.remove(sld)
    layouts = []
    for m_i, master in enumerate(prs.slide_masters):
        for layout in master.slide_layouts:
            prs.slides.add_slide(layout)
            layouts.append({"index": len(layouts) + 1, "name": layout.name, "master": m_i, "layout": layout})
    out = tmp / "layouts.pptx"
    prs.save(str(out))
    return prs, layouts, out


def ph_box(shapes, types) -> dict | None:
    from pptx.enum.shapes import PP_PLACEHOLDER
    wanted = {getattr(PP_PLACEHOLDER, t) for t in types}
    for sh in shapes:
        if sh.is_placeholder and sh.placeholder_format.type in wanted and sh.width:
            return {"x": round(sh.left / EMU_PER_BP), "y": round(sh.top / EMU_PER_BP),
                    "w": round(sh.width / EMU_PER_BP), "h": round(sh.height / EMU_PER_BP)}
    return None


def theme_info(prs) -> tuple[dict, dict]:
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    try:
        theme = etree.fromstring(prs.slide_masters[0].part.part_related_by(RT.THEME).blob)
    except Exception:  # noqa: BLE001 -- template without a readable theme
        return {}, {}
    ns = {"a": A_NS}

    def color(tag):
        el = theme.find(f".//a:clrScheme/a:{tag}", ns)
        if el is None or not len(el):
            return None
        c = el[0]
        val = c.get("val") if c.tag.endswith("srgbClr") else c.get("lastClr")
        return f"#{val}" if val else None

    def font(kind, script):
        el = theme.find(f".//a:fontScheme/a:{kind}/a:{script}", ns)
        face = el.get("typeface") if el is not None else ""
        return face if face and not face.startswith("+") else None

    palette = {k: v for k, v in {"primary": color("accent1"), "accent": color("accent2"), "text": color("dk1"),
                                 "title": color("dk1"), "inverse": color("lt1")}.items() if v}
    fonts = {"latin_heading": font("majorFont", "latin"), "latin_body": font("minorFont", "latin"),
             "cjk_heading": font("majorFont", "ea"), "cjk_body": font("minorFont", "ea")}
    return palette, fonts


def import_pptx(src: Path, mode: str, mapping: dict, out: Path, tmp: Path) -> tuple[dict, dict, dict]:
    from pptx import Presentation
    orig = Presentation(str(src))
    page = {"width_bp": round(orig.slide_width / EMU_PER_BP), "height_bp": round(orig.slide_height / EMU_PER_BP)}
    palette, fonts = theme_info(orig)
    geometry = {}
    if mode == "slides":
        if not mapping:
            die("--mode slides needs --map role=slideNumber,...")
        pdf = soffice_to_pdf(src, tmp)
        files = {}
        for role, num in mapping.items():
            extract_page(pdf, num, out / "bg" / f"{role}.pdf")
            files[role] = f"bg/{role}.pdf"
        warn("slide mode copies everything on those slides; remove sample text from the template first")
        return page, files, {"palette": palette, "fonts": fonts, **geometry}

    prs, layouts, deck = pptx_layout_deck(src, tmp)
    pdf = soffice_to_pdf(deck, tmp)
    sheet = layout_sheet(pdf, layouts, out, tmp)
    print(f"wrote {sheet}  <- numbered thumbnails of every layout")
    print("layouts found:")
    guessed = {}
    for lay in layouts:
        role = guess_role(lay["name"])
        if role and role not in guessed:
            guessed[role] = lay["index"]
        print(f"  {lay['index']:>2}  {lay['name']:<32} {('-> ' + role) if role else ''}")
    chosen = {**guessed, **mapping}
    if "body" not in chosen:
        die("could not identify a body layout; look at layouts.png and pass --map body=N[,cover=N,...]"
            " (rerun with --force)")
    files = {}
    for role, num in chosen.items():
        extract_page(pdf, num, out / "bg" / f"{role}.pdf")
        files[role] = f"bg/{role}.pdf"
    body_layout = layouts[chosen["body"] - 1]["layout"]
    master = body_layout.slide_master
    geometry["title"] = ph_box(body_layout.placeholders, ["TITLE", "CENTER_TITLE"]) or \
        ph_box(master.placeholders, ["TITLE"])
    geometry["safe_area"] = ph_box(body_layout.placeholders, ["BODY", "OBJECT"]) or \
        ph_box(master.placeholders, ["BODY", "OBJECT"])
    geometry["page_number"] = ph_box(body_layout.placeholders, ["SLIDE_NUMBER"]) or \
        ph_box(master.placeholders, ["SLIDE_NUMBER"])
    return page, files, {"palette": palette, "fonts": fonts, **{k: v for k, v in geometry.items() if v}}


def layout_sheet(pdf: Path, layouts: list, out: Path, tmp: Path) -> Path:
    run(["pdftoppm", "-png", "-scale-to-x", "400", "-scale-to-y", "-1", str(pdf), str(tmp / "lay")])
    ims = [Image.open(f).convert("RGB") for f in sorted(tmp.glob("lay-*.png"))]
    w, h, cols = ims[0].width, ims[0].height, 3
    sheet = Image.new("RGB", (cols * (w + 10), ((len(ims) + cols - 1) // cols) * (h + 10)), "#888888")
    for i, (im, lay) in enumerate(zip(ims, layouts)):
        d = ImageDraw.Draw(im)
        d.rectangle([0, 0, w, 20], fill="white")
        d.text((6, 4), f"{lay['index']}: {lay['name']}"[:60], fill="black")
        sheet.paste(im, ((i % cols) * (w + 10), (i // cols) * (h + 10)))
    dest = out / "layouts.png"
    sheet.save(dest)
    return dest


# ------------------------------------------------------------------ PDF / images
def page_size_of(path: Path) -> dict:
    if path.suffix.lower() == ".pdf":
        box = PdfReader(str(path)).pages[0].mediabox
        return {"width_bp": round(float(box.width)), "height_bp": round(float(box.height))}
    w, h = Image.open(path).size
    return {"width_bp": 960, "height_bp": round(960 * h / w)}


def import_files(args, mapping: dict, out: Path) -> tuple[dict, dict]:
    files = {}
    if args.pdf:
        if not mapping:
            die("--pdf needs --map role=page,... (e.g. cover=1,body=2,closing=3)")
        for role, num in mapping.items():
            extract_page(args.pdf, num, out / "bg" / f"{role}.pdf")
            files[role] = f"bg/{role}.pdf"
    for role in ROLES:
        src = getattr(args, role)
        if not src:
            continue
        ext = src.suffix.lower()
        if ext not in (".pdf", ".png", ".jpg", ".jpeg"):
            die(f"{src}: backgrounds must be PDF, PNG or JPG")
        dest = out / "bg" / f"{role}{'.pdf' if ext == '.pdf' else ext}"
        if ext == ".pdf":
            extract_page(src, 1, dest)
        else:
            shutil.copy2(src, dest)
        files[role] = f"bg/{dest.name}"
    if not files:
        die("no template source given (see --help)")
    first = out / next(iter(files.values()))
    return page_size_of(first), files


# ------------------------------------------------------------------ output
def scale_box(box: dict, sx: float, sy: float) -> dict:
    return {"x": round(box["x"] * sx), "y": round(box["y"] * sy), "w": round(box["w"] * sx), "h": round(box["h"] * sy)}


def preview(out: Path, t: dict) -> Path:
    tiles = []
    W = 640
    s = W / t["page_size"]["width_bp"]
    with tempfile.TemporaryDirectory() as td:
        for role in ROLES:
            f = t["backgrounds"].get(role)
            if not f:
                continue
            path = out / f
            if path.suffix == ".pdf":
                run(["pdftoppm", "-png", "-singlefile", "-scale-to-x", str(W), "-scale-to-y", "-1", str(path),
                     f"{td}/{role}"])
                im = Image.open(f"{td}/{role}.png").convert("RGB")
            else:
                im = Image.open(path).convert("RGB").resize((W, round(W * t["page_size"]["height_bp"] /
                                                                       t["page_size"]["width_bp"])))
            d = ImageDraw.Draw(im)
            if role == "body":
                for key, col in (("safe_area", "#18a058"), ("title", "#2f6fdb"), ("page_number", "#d03050")):
                    b = t[key]
                    d.rectangle([b["x"] * s, b["y"] * s, (b["x"] + b["w"]) * s, (b["y"] + b["h"]) * s],
                                outline=col, width=2)
            d.rectangle([0, 0, 120, 22], fill="white")
            d.text((6, 4), role, fill="black")
            tiles.append(im)
    cols = 2
    h = tiles[0].height
    sheet = Image.new("RGB", (cols * W + 10, ((len(tiles) + 1) // cols) * (h + 10)), "#888888")
    for i, im in enumerate(tiles):
        sheet.paste(im, ((i % cols) * (W + 10), (i // cols) * (h + 10)))
    dest = out / "preview.png"
    sheet.save(dest)
    return dest


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out_dir", type=Path)
    ap.add_argument("--pptx", type=Path)
    ap.add_argument("--mode", choices=["layouts", "slides"], default="layouts")
    ap.add_argument("--pdf", type=Path)
    for role in ROLES:
        ap.add_argument(f"--{role}", type=Path, help=f"background for {role} pages (PDF/PNG/JPG)")
    ap.add_argument("--map", help="role=number,... (PDF page, layout number or slide number)")
    ap.add_argument("--safe", help="content box x,y,w,h in bp")
    ap.add_argument("--title", help="title box x,y,w,h in bp")
    ap.add_argument("--number", help="page-number box x,y,w,h in bp")
    ap.add_argument("--name", help="template name")
    ap.add_argument("--force", action="store_true", help="overwrite an existing out_dir")
    a = ap.parse_args()
    need("pdfseparate", "pdftoppm")

    out = a.out_dir.resolve()
    if out.exists():
        if not a.force:
            die(f"{out} exists (use --force to overwrite)")
        shutil.rmtree(out)
    (out / "bg").mkdir(parents=True)
    mapping = parse_map(a.map)

    extra: dict = {}
    with tempfile.TemporaryDirectory() as td:
        if a.pptx:
            page, files, extra = import_pptx(a.pptx, a.mode, mapping, out, Path(td))
        else:
            page, files = import_files(a, mapping, out)

    t = copy.deepcopy(DEFAULT_TEMPLATE)
    sx, sy = page["width_bp"] / 960, page["height_bp"] / 540
    for key in ("safe_area", "title", "page_number"):
        t[key] = scale_box(t[key], sx, sy)
    t["page_size"] = page
    t["name"] = a.name or out.name
    t["backgrounds"] = {r: files.get(r) for r in ROLES}
    t["palette"] = {**t["palette"], "inverse": "#FFFFFF", **extra.get("palette", {})}
    for key in ("safe_area", "title", "page_number"):
        if extra.get(key):
            t[key] = extra[key]
    f = extra.get("fonts") or {}
    for script, kind, key in (("latin", "heading", "latin_heading"), ("latin", "body", "latin_body"),
                              ("cjk", "heading", "cjk_heading"), ("cjk", "body", "cjk_body")):
        if f.get(key):
            t["fonts"][script][kind] = [f[key]] + [x for x in t["fonts"][script][kind] if x != f[key]]
    for key, flag in (("safe_area", a.safe), ("title", a.title), ("page_number", a.number)):
        if flag:
            t[key] = parse_box(flag)
    missing = [r for r in ROLES if not files.get(r)]
    if missing:
        warn(f"no background for {', '.join(missing)}: body (or cover for closing) will be reused")
        if not files.get("closing") and files.get("cover"):
            t["backgrounds"]["closing"] = files["cover"]

    order = ["name", "page_size", "backgrounds", "safe_area", "title", "page_number", "palette", "fonts", "sizes"]
    with (out / "template.yaml").open("w", encoding="utf-8") as fh:
        fh.write("# Generated by latex-slides import_template.py. Review preview.png and adjust boxes,\n"
                 "# palette and fonts so they match the template. Coordinates: bp from the top-left corner.\n")
        yaml.safe_dump({k: t[k] for k in order}, fh, allow_unicode=True, sort_keys=False, width=100)
    print(f"wrote {out / 'template.yaml'}")
    print(f"wrote {preview(out, t)}  <- check the boxes before using the template")


if __name__ == "__main__":
    main()
