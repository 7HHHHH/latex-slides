# Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `ls-config.tex not found` | Compile through `build.py` (it regenerates the file from template.yaml). |
| `Illegal unit of measure` | A coordinate is not an integer (e.g. `56.0` in template.yaml) or contains an expression pgfmath cannot parse. |
| Boxes (tofu) / `Missing character` | Font lacks glyphs. Add a font with that script to `fonts` in template.yaml; for CJK set `language` to zh/ja/ko. |
| Fonts differ between machines | First installed font in each list wins. Put the template's font first and free fallbacks after it. Fonts become paths in SVG/PNG, so the PPTX looks the same everywhere once built. |
| Text outside safe area | Shorten, split the slide, or move the block. Do not shrink below template sizes. |
| `page N has no \slideid` | Add `\slideid{id}` inside every frame. |
| Black patches in WPS / older Office | SVG soft masks. Flatten transparent images to opaque PNG, or rebuild with `--raster-only`. PowerPoint 365 is the supported target. |
| Slide looks blurry in Keynote/Google Slides | They show the PNG fallback; raise `build.png_width` (e.g. 3840). |
| Video does not play in the PDF | Many viewers block `run:` links. Ship `media/` next to the PDF or use the PPTX. |
| Video does not play in PowerPoint | Use H.264/AAC MP4. |
| `render.pptx_vs_pdf` fails | Usually LibreOffice rendering differences with very complex SVG; inspect both images. If the PPTX is right in PowerPoint, raise `render_diff_threshold` and note it in the report. |
| PPTX import picks wrong layouts | Look at `layouts.png`, pass `--map body=N,cover=N,... --force`. |
| Template has decorations in the content area | Shrink `safe_area` so content avoids them. |
| PPTX too large | Downsample bitmaps in `assets/`, lower `png_width`, compress videos. |
