"""End-to-end test: scaffold -> build -> check for an English and a Chinese deck.

Run: python3 -m unittest discover -s tests
Skips automatically when TeX or poppler is missing.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
TOOLS = ("xelatex", "latexmk", "pdftoppm", "pdftocairo", "pdftotext")


def py(*args, check=True):
    return subprocess.run([sys.executable, *map(str, args)], cwd=SCRIPTS, text=True,
                          capture_output=True, check=check)


@unittest.skipIf(any(shutil.which(t) is None for t in TOOLS), "TeX Live / poppler not installed")
class PipelineTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="latex-slides-test-"))

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def build_deck(self, lang: str) -> Path:
        deck = self.tmp / lang
        py(SCRIPTS / "new_deck.py", deck, "--title", "Test deck", "--presenter", "Tester",
           "--language", lang, "--duration", "105")
        cfg = (deck / "deck.yaml").read_text(encoding="utf-8")
        (deck / "deck.yaml").write_text(cfg.replace("script: ''", "script: hello"), encoding="utf-8")
        res = py(SCRIPTS / "build.py", deck, check=False)
        self.assertEqual(res.returncode, 0, res.stdout + res.stderr)
        return deck

    def check_outputs(self, deck: Path):
        man = json.loads((deck / "build" / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(len(man["pages"]), 5)
        self.assertEqual([p["role"] for p in man["pages"]], ["cover", "agenda", "section", "body", "closing"])
        with zipfile.ZipFile(man["pptx"]) as z:
            names = z.namelist()
            self.assertEqual(sum(n.endswith(".svg") for n in names), 5)
            self.assertIn("svgBlip", z.read("ppt/slides/slide4.xml").decode())
        res = py(SCRIPTS / "check.py", deck, "--no-render", check=False)
        self.assertEqual(res.returncode, 0, res.stdout + res.stderr)
        # rebuilding must never overwrite
        py(SCRIPTS / "build.py", deck)
        self.assertEqual(len(list((deck / "output").glob("*.pptx"))), 2)

    def test_english(self):
        self.check_outputs(self.build_deck("en"))

    def test_chinese(self):
        self.check_outputs(self.build_deck("zh-CN"))

    def test_text_only_slide_fails(self):
        deck = self.build_deck("en")
        tex = (deck / "deck.tex").read_text(encoding="utf-8")
        start = tex.index("\\begin{tikzpicture}")
        end = tex.index("\\end{tikzpicture}") + len("\\end{tikzpicture}")
        (deck / "deck.tex").write_text(tex[:start] + "no picture" + tex[end:], encoding="utf-8")
        py(SCRIPTS / "build.py", deck)
        res = py(SCRIPTS / "check.py", deck, "--no-render", check=False)
        self.assertNotEqual(res.returncode, 0)
        self.assertIn("every_content_slide_has_visual", res.stdout)

    def test_raw_includegraphics_counts_as_visual(self):
        deck = self.build_deck("en")
        from PIL import Image
        Image.new("RGB", (400, 300), "#3F86D8").save(deck / "assets" / "fig.png")
        tex = (deck / "deck.tex").read_text(encoding="utf-8")
        start = tex.index("\\begin{tikzpicture}")
        end = tex.index("\\end{tikzpicture}") + len("\\end{tikzpicture}")
        (deck / "deck.tex").write_text(tex[:start] + "\\includegraphics[width=300bp]{fig.png}" + tex[end:],
                                       encoding="utf-8")
        py(SCRIPTS / "build.py", deck)
        res = py(SCRIPTS / "check.py", deck, "--no-render", check=False)
        self.assertEqual(res.returncode, 0, res.stdout)
        man = json.loads((deck / "build" / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual([p["visuals"] for p in man["pages"]], [0, 0, 0, 1, 0])   # backgrounds don't count

    def test_scaffold_duration_matches(self):
        deck = self.tmp / "dur"
        py(SCRIPTS / "new_deck.py", deck, "--title", "T", "--duration", "600", "--template", "paper")
        import yaml
        cfg = yaml.safe_load((deck / "deck.yaml").read_text(encoding="utf-8"))
        self.assertEqual(sum(s["seconds"] for s in cfg["slides"]), 600)
        self.assertTrue((deck / "template" / "bg" / "body.pdf").exists())

    def test_japanese_labels_and_fonts(self):
        deck = self.build_deck("ja")
        cfg = (deck / "ls-config.tex").read_text(encoding="utf-8")
        self.assertIn("Noto Sans CJK JP", cfg)
        self.assertNotIn("SimSun", cfg)
        notes = next((deck / "output").glob("*.notes.md")).read_text(encoding="utf-8")
        self.assertIn("目安時間", notes)

    def test_builtin_template_by_name(self):
        deck = self.build_deck("en")
        cfg = (deck / "deck.yaml").read_text(encoding="utf-8")
        (deck / "deck.yaml").write_text(cfg.replace("template: template", "template: paper"), encoding="utf-8")
        shutil.rmtree(deck / "template")
        py(SCRIPTS / "build.py", deck)
        self.assertIn("templates/paper/bg/body.pdf", (deck / "ls-config.tex").read_text(encoding="utf-8"))

    def test_sparse_slide_warns(self):
        deck = self.build_deck("en")
        tex = (deck / "deck.tex").read_text(encoding="utf-8")
        start = tex.index("\\at{56}{\\lsContentTop}{380}")
        end = tex.index("\\at{470}")
        (deck / "deck.tex").write_text(tex[:start] + tex[end:], encoding="utf-8")
        py(SCRIPTS / "build.py", deck)
        res = py(SCRIPTS / "check.py", deck, "--no-render", check=False)
        self.assertIn("WARN  visual.information_density", res.stdout)

    def test_import_missing_file(self):
        res = py(SCRIPTS / "import_template.py", self.tmp / "t", "--pptx", self.tmp / "nope.pptx", check=False)
        self.assertNotEqual(res.returncode, 0)
        self.assertIn("file not found", res.stderr)
        self.assertNotIn("Traceback", res.stderr)


if __name__ == "__main__":
    unittest.main()
