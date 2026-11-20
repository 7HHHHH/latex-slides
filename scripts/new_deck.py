#!/usr/bin/env python3
"""Scaffold a latex-slides deck project.

Example:
  new_deck.py decks/grpo_20261120 --title "GRPO in practice" --presenter "Ada" \
      --date 2026-11-20 --language en --duration 900 [--template path/to/template]
"""
from __future__ import annotations

import argparse
import datetime as dt
import shutil
from pathlib import Path

import yaml

from lib.common import SKILL_ROOT, die, is_cjk, write_ls_config

TEXT = {
    "en": {"agenda": "Agenda", "part1": "Background", "part2": "Method and results", "thanks": "Thank you",
           "questions": "Questions and discussion", "body_title": "Replace with a claim-style title",
           "lead": "One sentence: the question this slide answers.",
           "p1": "First point", "p1t": "Explain what the figure shows.",
           "p2": "Second point", "p2t": "Why it matters for the audience.",
           "take": "Takeaway: one sentence that leads to the next slide.", "cap": "Figure: schematic (illustrative).",
           "a": "Input", "b": "Method", "c": "Result", "section": "Part 1", "presenter": "Presenter"},
    "zh": {"agenda": "目\\;录", "part1": "研究背景", "part2": "方法与结果", "thanks": "THANKS",
           "questions": "请老师同学批评指正", "body_title": "用一句结论作为标题",
           "lead": "一句话引导：本页要回答的问题。",
           "p1": "要点一", "p1t": "解释图中展示了什么。",
           "p2": "要点二", "p2t": "说明它对听众为什么重要。",
           "take": "本页结论：一句话收束，并引出下一页。", "cap": "图：示意图（非真实数据）。",
           "a": "输入", "b": "方法", "c": "结果", "section": "第一部分", "presenter": "汇报人"},
}

DECK_TEX = r"""% latex-slides deck. Edit content here; rebuild with scripts/build.py.
\documentclass[aspectratio=169,t]{beamer}
\usepackage{latex-slides}
\graphicspath{{assets/}}
\hypersetup{colorlinks=true,urlcolor=lsPrimary,linkcolor=lsPrimary,pdftitle={@TITLE@},pdfauthor={@PRESENTER@}}
\begin{document}

% ---------------------------------------------------------------- cover
{\lsrole{cover}
\begin{frame}[plain]\slideid{cover}
\at{88}{104}{780}{{\headingfont\fontsize{40}{52}\selectfont\bfseries\color{lsPrimary} @TITLE@\par}}
\at{88}{430}{780}{{\color{lsInverse}\fontsize{20}{30}\selectfont @PRESENTER_LABEL@\par @DATE@}}
\end{frame}}

% ---------------------------------------------------------------- agenda
{\lsrole{agenda}
\begin{frame}[plain]\slideid{agenda}
\at{40}{230}{220}{{\headingfont\fontsize{36}{46}\selectfont\bfseries\color{lsInverse} @AGENDA@}}
\at{370}{170}{520}{{\headingfont\fontsize{28}{40}\selectfont\color{lsAccent} 01\quad\color{lsText}\bfseries @PART1@\par}}
\at{370}{290}{520}{{\headingfont\fontsize{28}{40}\selectfont\color{lsAccent} 02\quad\color{lsText}\bfseries @PART2@\par}}
\end{frame}}

% ---------------------------------------------------------------- section
{\lsrole{section}
\begin{frame}[plain]\slideid{part1}
\at{96}{220}{760}{{\headingfont\fontsize{36}{46}\selectfont\bfseries\color{lsPrimary} 01\quad @PART1@}}
\end{frame}}

% ---------------------------------------------------------------- body: text left, diagram right
\lsChapter{01}
\begin{frame}{@BODY_TITLE@}\slideid{example}
\lead{@LEAD@}
\at{56}{\lsContentTop}{380}{\point{@P1@}{@P1T@}\vspace{14pt}\point{@P2@}{@P2T@}}
\at{470}{\lsContentTop}{434}{\centering
\begin{tikzpicture}[x=1bp,y=1bp,node distance=36bp,
    box/.style={draw=lsAccent,line width=1.2bp,rounded corners=4bp,minimum width=112bp,minimum height=56bp,
                font=\headingfont\fontsize{18}{22}\selectfont},
    arr/.style={-latex,line width=1.4bp,color=lsPrimary}]
  \node[box] (a) {@A@};
  \node[box,right=of a,fill=lsAccent!12] (b) {@B@};
  \node[box,below=48bp of b] (c) {@C@};
  \draw[arr] (a) -- (b); \draw[arr] (b) -- (c);
\end{tikzpicture}\par\vspace{6pt}
{\fontsize{\lsSizeCaption}{\lsLeadCaption}\selectfont\color{lsMuted}@CAP@}}
\takeaway{@TAKE@}
\end{frame}

% ---------------------------------------------------------------- closing
{\lsrole{closing}
\begin{frame}[plain]\slideid{closing}
\at{80}{230}{800}{\centering{\headingfont\fontsize{40}{50}\selectfont\bfseries\color{lsInverse} @THANKS@\par}}
\at{80}{340}{800}{\centering{\fontsize{22}{30}\selectfont\color{lsInverse} @QUESTIONS@}}
\end{frame}}

\end{document}
"""


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck_dir", type=Path)
    ap.add_argument("--title", required=True)
    ap.add_argument("--presenter", default="")
    ap.add_argument("--date", default=dt.date.today().isoformat(), help="YYYY-MM-DD")
    ap.add_argument("--language", default="en", help="BCP-47 tag of the slide language, e.g. en, zh-CN, ja")
    ap.add_argument("--duration", type=int, default=600, help="talk length in seconds")
    ap.add_argument("--sourcing", choices=["strict", "normal", "off"], default="normal")
    ap.add_argument("--template", type=Path, default=SKILL_ROOT / "templates" / "default",
                    help="template directory containing template.yaml (see import_template.py)")
    a = ap.parse_args()

    deck = a.deck_dir.resolve()
    if deck.exists() and any(deck.iterdir()):
        die(f"{deck} exists and is not empty; refusing to overwrite")
    if not (a.template / "template.yaml").exists():
        die(f"{a.template}/template.yaml not found")
    for sub in ("assets", "media", "build", "output", "verify"):
        (deck / sub).mkdir(parents=True, exist_ok=True)
    shutil.copytree(a.template, deck / "template", ignore=shutil.ignore_patterns("src", "preview*"))
    shutil.copy2(SKILL_ROOT / "latex" / "latex-slides.sty", deck / "latex-slides.sty")

    lang = "zh" if is_cjk(a.language) else "en"
    t = TEXT[lang]
    date_text = a.date
    if lang == "zh":
        y, m, d = a.date.split("-")
        date_text = f"{int(y)} 年 {int(m)} 月 {int(d)} 日"
    presenter_label = f"{t['presenter']}{'：' if lang == 'zh' else ': '}{a.presenter}" if a.presenter else ""
    rep = {"TITLE": a.title, "PRESENTER": a.presenter, "PRESENTER_LABEL": presenter_label, "DATE": date_text,
           "AGENDA": t["agenda"], "PART1": t["part1"], "PART2": t["part2"], "BODY_TITLE": t["body_title"],
           "LEAD": t["lead"], "P1": t["p1"], "P1T": t["p1t"], "P2": t["p2"], "P2T": t["p2t"],
           "A": t["a"], "B": t["b"], "C": t["c"], "CAP": t["cap"], "TAKE": t["take"],
           "THANKS": t["thanks"], "QUESTIONS": t["questions"]}
    tex = DECK_TEX
    for k, v in rep.items():
        tex = tex.replace(f"@{k}@", v)
    (deck / "deck.tex").write_text(tex, encoding="utf-8")

    slides = [
        {"id": "cover", "title": a.title, "seconds": 15, "script": "", "sources": []},
        {"id": "agenda", "title": t["agenda"].replace("\\;", ""), "seconds": 15, "script": "", "sources": []},
        {"id": "part1", "title": t["part1"], "seconds": 5, "script": "", "sources": []},
        {"id": "example", "title": t["body_title"], "seconds": 60, "script": "", "sources": []},
        {"id": "closing", "title": t["questions"], "seconds": 10, "script": "", "sources": []},
    ]
    meta = {"title": a.title, "presenter": a.presenter, "date": a.date, "language": a.language,
            "duration_seconds": a.duration, "sourcing": a.sourcing,
            "output_name": f"{a.title}_{a.date}".replace("/", "-"),
            "template": "template",
            "build": {"svg": True, "png_width": 1920},
            "checks": {"max_text_units": 80, "duration_tolerance": 0.10, "render_diff_threshold": 12},
            "slides": slides}
    with (deck / "deck.yaml").open("w", encoding="utf-8") as f:
        f.write("# Single source for speaker notes, timing and sources. Slide ids match \\slideid{...} in deck.tex.\n")
        yaml.safe_dump(meta, f, allow_unicode=True, sort_keys=False, width=100)
    (deck / "references.md").write_text("# References\n\n<!-- One entry per source used in the deck. -->\n", encoding="utf-8")
    (deck / "assets" / "SOURCES.md").write_text(
        "# Asset provenance\n\n| file | origin (URL / paper + figure / self-made) | license or permission | notes |\n|---|---|---|---|\n",
        encoding="utf-8")
    (deck / ".gitignore").write_text("build/\nls-config.tex\n", encoding="utf-8")
    write_ls_config(deck, "template", a.language)
    print(deck)


if __name__ == "__main__":
    main()
