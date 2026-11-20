# template.yaml

A template directory contains `template.yaml` and background files (`bg/*.pdf|png|jpg`).
`new_deck.py --template DIR` copies it into the deck as `template/`.

```yaml
name: my-lab
page_size: {width_bp: 960, height_bp: 540}      # 16:9; 4:3 would be 720 x 540
backgrounds:                                    # PDF (vector, preferred), PNG or JPG
  cover: bg/cover.pdf
  agenda: bg/agenda.pdf                         # optional
  section: null                                 # optional; null -> body background
  body: bg/body.pdf                             # required for a non-white deck
  closing: bg/closing.pdf                       # optional; importer falls back to cover
safe_area: {x: 56, y: 96, w: 848, h: 392}       # free content region of the body page
title: {x: 56, y: 30, w: 848, h: 52}            # frame title box (text vertically centred)
page_number: {x: 844, y: 508, w: 60, h: 20}     # right-aligned frame number
palette:                                        # names become LaTeX colours lsPrimary, ...
  primary: "#123A5A"     # headings, emphasis
  accent: "#3F86D8"      # diagrams, markers
  muted: "#4F5964"       # captions, details
  text: "#111111"
  title: "#111111"
  rule: "#BFCAD5"        # thin lines
  inverse: "#FFFFFF"     # text on dark areas
fonts:                                          # first installed font wins
  latin: {heading: [Arial, TeX Gyre Heros], body: [Times New Roman, TeX Gyre Termes]}
  cjk:   {heading: [Microsoft YaHei, Noto Sans CJK SC], body: [SimSun, Noto Serif CJK SC]}
sizes: {title: 24, heading: 21, body: 20, detail: 18, caption: 13, number: 14}   # pt
```

Coordinates are bp from the top-left corner. All boxes must be integers.

## Importing

| Source | Command | Notes |
|---|---|---|
| PowerPoint template | `import_template.py OUT --pptx t.pptx` | Renders every layout (empty placeholders are invisible). Reads theme colours, theme fonts, and title/body/slide-number placeholder boxes of the body layout. Check `layouts.png`, then fix roles with `--map body=4,cover=1,... --force`. |
| PowerPoint slides | `--pptx t.pptx --mode slides --map cover=1,body=2` | For templates whose design lives on slides, not layouts. Remove sample text first. |
| One PDF | `--pdf bg.pdf --map cover=1,body=2,closing=3` | Page numbers, 1-based. |
| Separate files | `--cover a.pdf --body b.png --closing c.jpg` | PDF page 1 or images. |

Override boxes on the command line with `--safe x,y,w,h`, `--title x,y,w,h`,
`--number x,y,w,h`, or edit `template.yaml` afterwards.

## Fitting a template (what the agent must do)

1. Open `preview.png`. The **green** safe area must not overlap logos, bands or
   decorations of the body background; the **blue** title box must sit where the
   template expects the title; the **red** box where the page number goes.
2. Adjust the boxes, re-run the import with overrides (or edit YAML and render one test
   deck), and compare again.
3. Pick palette colours from the template (theme colours are imported for PPTX; for
   PDF/images sample them from the background).
4. Set fonts the template uses, followed by free fallbacks that exist on most systems.
5. Show the preview and one rendered body slide to the user for approval.

Templates usually carry copyright or institutional branding: keep user templates out
of public repositories unless the user says otherwise.
