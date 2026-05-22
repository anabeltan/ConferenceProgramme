#!/usr/bin/env python3
"""Generate a CASCA-2026-style LaTeX conference programme for SMBH 2026.

Reads participant data from the XLSX workbook (Sheet2) and writes:
  - sections/generated_smbh_schedule.tex   — schedule tables, grouped by day/theme
  - sections/generated_smbh_abstracts.tex  — abstract sections, grouped by day/theme
  - smbh_main.tex                          — standalone compilable main document

Compile with:  lualatex smbh_main.tex   (twice for TOC/hyperref)
"""

from __future__ import annotations

import argparse
import re
import sys
import textwrap
import unicodedata
from collections import defaultdict
from pathlib import Path

from generate_programme_tex import Entry, build_entries, latex_escape, latex_paragraphs


# ---------------------------------------------------------------------------
# Day palette definitions
# ---------------------------------------------------------------------------

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# (background, accent, glow, date-line)
DAY_PALETTES: dict[str, tuple[str, str, str, str]] = {
    "Monday":    ("mondaybg",    "mondayaccent",    "mondayglow",    "June 29, 2026"),
    "Tuesday":   ("tuesdaybg",   "tuesdayaccent",   "tuesdayglow",   "June 30, 2026"),
    "Wednesday": ("wednesdaybg", "wednesdayaccent", "wednesdayglow", "July 1, 2026"),
    "Thursday":  ("thursdaybg",  "thursdayaccent",  "thursdayglow",  "July 2, 2026"),
    "Friday":    ("fridaybg",    "fridayaccent",    "fridayglow",    "July 3, 2026"),
}

# Theme labels (themes in the XLSX are stored as single digits)
THEME_LABELS: dict[str, str] = {
    "1": "Black Hole Demographics and Scaling Relations",
    "2": "Black Hole Growth and AGN Feedback",
    "3": "Accretion, Jets and Multi-messenger Signatures",
    "4": "First Black Holes, Seeds and High-redshift",
}


def primary_theme(theme_raw: str) -> str:
    """Return the primary (first) theme number from values like '1 or 2'."""
    m = re.search(r"\d", theme_raw)
    return m.group(0) if m else theme_raw


def theme_label(theme_raw: str) -> str:
    key = primary_theme(theme_raw)
    return THEME_LABELS.get(key, f"Theme {key}")


def day_name_from_label(day_label: str) -> str:
    """Extract the weekday name from 'Monday, June 29, 2026'."""
    return day_label.split(",")[0].strip() if "," in day_label else day_label


TEXT_REPLACEMENTS = {
    "\u2061": "",
    "⍺": "α",
    "⁻": "-",
    "☉": "sun",
}


def normalize_programme_text(value: str) -> str:
    text = unicodedata.normalize("NFKC", value or "")
    for source, target in TEXT_REPLACEMENTS.items():
        text = text.replace(source, target)
    return text


def programme_tex(value: str) -> str:
    return latex_escape(normalize_programme_text(value))


def programme_paragraphs(value: str) -> str:
    normalized = normalize_programme_text(value)
    parts = [programme_tex(part.strip()) for part in normalized.split("\n\n") if part.strip()]
    return "\n\n".join(parts) if parts else "Not provided."


def is_poster_entry(entry: Entry) -> bool:
    return "Poster" in entry.presentation_type


# ---------------------------------------------------------------------------
# LaTeX preamble / main document
# ---------------------------------------------------------------------------

PREAMBLE = r"""% !TEX TS-program = lualatex
\documentclass[12pt,letterpaper]{report}

\usepackage[margin=1in]{geometry}
\usepackage{fontspec}
\usepackage{microtype}
\usepackage{setspace}
\usepackage{parskip}
\usepackage[table,dvipsnames]{xcolor}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{longtable}
\usepackage{tabularx}
\usepackage{array}
\usepackage{multirow}
\usepackage{hhline}
\IfFileExists{ragged2e.sty}{\usepackage{ragged2e}}{}
\IfFileExists{needspace.sty}{\usepackage{needspace}}{}
\usepackage{hyperref}
\usepackage{fancyhdr}
\IfFileExists{eso-pic.sty}{\usepackage{eso-pic}}{}

\IfFontExistsTF{TeX Gyre Pagella}{
  \setmainfont{TeX Gyre Pagella}
}{
  \setmainfont{Latin Modern Roman}
}
\IfFontExistsTF{Avenir Next}{
  \setsansfont{Avenir Next}
}{\IfFontExistsTF{TeX Gyre Heros}{
  \setsansfont{TeX Gyre Heros}
}{
  \setsansfont{Latin Modern Sans}
}}
\IfFontExistsTF{DejaVu Sans}{
  \newfontfamily\smbhnotefont{DejaVu Sans}
}{
  \newfontfamily\smbhnotefont{Latin Modern Sans}
}

% ---- Colours ---------------------------------------------------------------
\definecolor{smbhblue}{HTML}{1C4E80}
\definecolor{smbhteal}{HTML}{2E7D8A}
\definecolor{smbhgold}{HTML}{C8922A}
\definecolor{smbhdeep}{HTML}{16324A}
\definecolor{paperwarm}{HTML}{FBF8F1}
\definecolor{lightrule}{HTML}{C9DCE8}
\definecolor{softpanel}{HTML}{EFF6FB}
\definecolor{nightbg}{HTML}{EAF5FB}
\definecolor{nightpanel}{HTML}{D7E9F4}
\definecolor{starlight}{HTML}{FFFDF8}
\definecolor{nebulablue}{HTML}{6FAFD6}
\definecolor{nebulateal}{HTML}{81C7CF}
\definecolor{auroraglow}{HTML}{4E8EBC}
\definecolor{deepviolet}{HTML}{406C9A}
\definecolor{nebulaedge}{HTML}{F4FAFD}
\definecolor{noteblue}{HTML}{356FA8}
\definecolor{noteink}{HTML}{2A5076}
\definecolor{titlewash}{HTML}{F1F8FD}
\definecolor{talkaccent}{HTML}{356FA8}
\definecolor{talkwash}{HTML}{EAF4FB}
\definecolor{posteraccent}{HTML}{C87363}
\definecolor{posterwash}{HTML}{FBEEE9}
\definecolor{abstractaccent}{HTML}{3A8B7B}
\definecolor{abstractwash}{HTML}{EAF7F3}
\definecolor{posterabstractaccent}{HTML}{8F6AAE}
\definecolor{posterabstractwash}{HTML}{F3ECF8}
\definecolor{headerink}{HTML}{5A4A3D}
\definecolor{headerline}{HTML}{CBB8A4}
% Day palettes
\definecolor{mondaybg}{HTML}{E8F4FB}
\definecolor{mondayaccent}{HTML}{4C9EC5}
\definecolor{mondayglow}{HTML}{8FD0E5}
\definecolor{tuesdaybg}{HTML}{F8EEE8}
\definecolor{tuesdayaccent}{HTML}{D98858}
\definecolor{tuesdayglow}{HTML}{EBC78C}
\definecolor{wednesdaybg}{HTML}{EAF7F3}
\definecolor{wednesdayaccent}{HTML}{4FA99D}
\definecolor{wednesdayglow}{HTML}{9ED9D0}
\definecolor{thursdaybg}{HTML}{EEF1FB}
\definecolor{thursdayaccent}{HTML}{6E8EDC}
\definecolor{thursdayglow}{HTML}{B9C7F4}
\definecolor{fridaybg}{HTML}{FAEEE8}
\definecolor{fridayaccent}{HTML}{C87363}
\definecolor{fridayglow}{HTML}{E7B59E}

\hypersetup{
  colorlinks=true,
  linkcolor=smbhblue,
  urlcolor=smbhteal,
  pdftitle={Supermassive Black Holes and Blue Notes},
  pdfauthor={SMBH 2026 Organizing Team}
}

\setstretch{1.03}
\setlength{\parindent}{0pt}
\setlength{\LTpre}{0.8em}
\setlength{\LTpost}{0.4em}
\renewcommand{\arraystretch}{1.15}
\arrayrulecolor{lightrule}
\rowcolors{2}{softpanel!75!white}{white}
\setcounter{tocdepth}{1}

\newcommand{\TableRaggedRight}{\RaggedRight}
\newcommand{\TopicNeedSpace}[1]{\Needspace{#1}}
\newcommand{\currentchapteraccent}{talkaccent}
\newcommand{\currentchapterwash}{talkwash}
\newcommand{\setchapterpalette}[2]{%
  \renewcommand{\currentchapteraccent}{#1}%
  \renewcommand{\currentchapterwash}{#2}%
}

% ---- Section formatting ----------------------------------------------------
\makeatletter
\def\chaptermark#1{\markboth{#1}{}}
\def\sectionmark#1{\markright{#1}{}}
\renewcommand\thesection{\arabic{section}}
\renewcommand\thesubsection{\thesection.\arabic{subsection}}
\renewcommand\section{\@startsection{section}{1}{\z@}{-3.25ex \@plus -1ex \@minus -.2ex}{1.1ex \@plus .2ex}{\normalfont\Large\sffamily\bfseries\color{smbhblue}}}
\renewcommand\subsection{\@startsection{subsection}{2}{\z@}{-2.4ex \@plus -1ex \@minus -.2ex}{0.7ex \@plus .15ex}{\normalfont\large\sffamily\bfseries\color{smbhteal}}}

\newcommand{\chaptertitlepanel}[1]{%
  \noindent\fcolorbox{\currentchapteraccent}{\currentchapterwash}{%
    \begin{minipage}{0.95\textwidth}
      \vspace{0.25em}%
      {\sffamily\bfseries\Huge\color{smbhdeep} #1\par}
      \vspace{0.35em}%
      {\color{\currentchapteraccent}\rule{0.28\textwidth}{1.2pt}}%
      \vspace{0.25em}%
    \end{minipage}%
  }%
}

\newcommand{\sectiondividerartleft}{%
  \IfFileExists{music/musical-notes-colorful-clipart-lg.png}{%
    \includegraphics[width=0.16\textwidth]{music/musical-notes-colorful-clipart-lg.png}%
  }{%
    \IfFileExists{music/colorful-music-notes-picture-2.png}{%
      \includegraphics[width=0.16\textwidth]{music/colorful-music-notes-picture-2.png}%
    }{}%
  }%
}

\newcommand{\titlecornerart}{%
  \IfFileExists{music/pngtree-colorful-music-notes-png-image_17160794.png}{%
    \includegraphics[width=0.15\textwidth]{music/pngtree-colorful-music-notes-png-image_17160794.png}%
  }{%
    \sectiondividerartleft
  }%
}

\newcommand{\sectiondividerartright}{%
  \IfFileExists{music/9448086.png}{%
    \includegraphics[width=0.11\textwidth]{music/9448086.png}%
  }{%
    \IfFileExists{music/images.png}{%
      \includegraphics[width=0.11\textwidth]{music/images.png}%
    }{%
      \IfFileExists{music/colorful-music-notes-picture-2.png}{%
        \includegraphics[width=0.15\textwidth]{music/colorful-music-notes-picture-2.png}%
      }{}%
    }%
  }%
}

\newcommand{\sectiondividerartbottom}{%
  \IfFileExists{music/colorful-music-notes-picture-2.png}{%
    \includegraphics[width=0.17\textwidth]{music/colorful-music-notes-picture-2.png}%
  }{%
    \IfFileExists{music/images.png}{%
      \includegraphics[width=0.15\textwidth]{music/images.png}%
    }{}%
  }%
}

\newcommand{\sectiondividerpage}[4]{%
  \clearpage
  \thispagestyle{empty}%
  \vspace*{0.08\textheight}%
  \noindent\hfill\sectiondividerartright\par
  \vspace{0.35em}%
  \begin{center}
    \fcolorbox{#2}{#3}{%
      \begin{minipage}{0.82\textwidth}
        \centering
        \vspace{1.1em}%
        {\sffamily\bfseries\Huge\color{#2} #1\par}
        \vspace{0.5em}%
        {\color{#2}\rule{0.42\textwidth}{1.25pt}\par}
        \vspace{0.7em}%
        {\sffamily\large\color{smbhdeep} #4\par}
        \vspace{1.05em}%
      \end{minipage}%
    }%
  \end{center}
  \vspace{0.35em}%
  \noindent\sectiondividerartbottom\par
  \clearpage
}

\def\@makechapterhead#1{%
  \vspace*{1.2em}%
  {\parindent\z@\raggedright\normalfont
    {\color{smbhgold}\sffamily\bfseries\large \@chapapp\space \thechapter\par}
    \vspace{0.4em}%
    {\color{deepviolet}\rule{\textwidth}{1.2pt}\par}
    \vspace{0.85em}%
    {\chaptertitlepanel{#1}\par}
    \vspace{1.35em}%
  }}
\def\@makeschapterhead#1{%
  \vspace*{1.2em}%
  {\parindent\z@\raggedright\normalfont
    {\color{deepviolet}\rule{\textwidth}{1.2pt}\par}
    \vspace{0.85em}%
    {\chaptertitlepanel{#1}\par}
    \vspace{1.15em}%
  }}
\makeatother

% ---- Page decorations ------------------------------------------------------
\newcommand{\smbhstarfield}{%
  {
  \setlength{\unitlength}{1pt}%
  \begin{picture}(0,0)
    \put(20,-22){\color{starlight}\circle*{2.9}}
    \put(74,-56){\color{smbhgold}\circle*{2.2}}
    \put(136,-34){\color{starlight}\circle*{2.5}}
    \put(212,-70){\color{nebulateal}\circle*{3.3}}
    \put(308,-24){\color{starlight}\circle*{2.9}}
    \put(398,-62){\color{smbhgold}\circle*{2.2}}
    \put(464,-32){\color{starlight}\circle*{2.5}}
    \put(38,-144){\color{nebulablue}\circle*{36}}
    \put(454,-170){\color{nebulateal}\circle*{28}}
    \put(46,-508){\color{nebulateal}\circle*{20}}
    \put(470,-538){\color{nebulablue}\circle*{32}}
    \put(84,-674){\color{starlight}\circle*{2.9}}
    \put(240,-706){\color{smbhgold}\circle*{2.9}}
    \put(430,-690){\color{starlight}\circle*{2.5}}
    \put(160,-320){\color{auroraglow!50}\circle*{14}}
    \put(510,-290){\color{nebulablue!60}\circle*{10}}
    \put(340,-480){\color{nebulateal!50}\circle*{16}}
    \put(430,-58){\makebox(0,0)[lt]{\titlecornerart}}
    \put(44,-650){\makebox(0,0)[lb]{\sectiondividerartleft}}
  \end{picture}%
  }
}
\newcommand{\smbhastrocorner}{%
  {
  \setlength{\unitlength}{1pt}%
  \begin{picture}(0,0)
    \put(476,-16){\color{nebulaedge!55}\circle*{108}}
    \put(532,-30){\color{auroraglow!38}\circle*{172}}
    \put(596,-104){\color{nebulateal!28}\circle*{240}}
    \put(556,-158){\color{nebulablue!20}\circle*{300}}
    \put(440,-22){\color{smbhblue!25}\circle*{1.0}}
    \put(450,-38){\color{starlight!88}\circle*{2.1}}
    \put(468,-86){\color{starlight!72}\circle*{1.4}}
    \put(504,-44){\color{smbhgold!90}\circle*{2.3}}
    \put(530,-118){\color{auroraglow!84}\circle*{1.8}}
    \put(558,-62){\color{starlight!78}\circle*{2.0}}
    \put(572,-98){\color{starlight!60}\circle*{1.1}}
    \put(588,-136){\color{smbhgold!74}\circle*{1.3}}
    \put(608,-50){\color{starlight!74}\circle*{1.6}}
    \put(622,-82){\color{auroraglow!70}\circle*{1.0}}
    \put(630,-110){\color{auroraglow!68}\circle*{1.2}}
    \put(490,-152){\color{starlight!62}\circle*{1.5}}
    \put(516,-178){\color{smbhgold!66}\circle*{1.1}}
  \end{picture}%
  }
}
\newcommand{\enablesmbhastrobackground}{%
  \AddToShipoutPictureBG{%
    \AtPageUpperLeft{\smbhastrocorner}%
  }%
}

% ---- Headers/footers -------------------------------------------------------
\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\sffamily\small\color{headerink}\bfseries SMBH and Blue Notes}
\fancyhead[R]{\sffamily\small\color{smbhdeep}\nouppercase{\leftmark}}
\fancyfoot[L]{\sffamily\small\color{smbhteal} \href{https://sites.google.com/view/smbh2026/home}{sites.google.com/view/smbh2026/home}}
\fancyfoot[C]{\IfFileExists{assets/Symbole_carré-UdeM.png}{\includegraphics[height=16pt]{assets/Symbole_carré-UdeM.png}}{}}
\fancyfoot[R]{\sffamily\small\color{smbhdeep}\thepage}
\renewcommand{\headrulewidth}{0.5pt}
\renewcommand{\footrulewidth}{0pt}
\setlength{\headheight}{25pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{headerline}\leaders\hrule height \headrulewidth\hfill}}

% ---- Day banner command ----------------------------------------------------
\newcommand{\daypalettebanner}[4]{%
  % #1=day name  #2=bg colour  #3=accent colour  #4=date line
  \noindent\fcolorbox{#3}{#2}{%
    \begin{minipage}{0.96\textwidth}
      \vspace{0.2em}%
      {\sffamily\bfseries\Large\color{smbhdeep} #1\par}
      \vspace{0.15em}%
      {\sffamily\itshape\small\color{#3} #4\par}
      \vspace{0.35em}%
      {\color{#3}\rule{0.28\textwidth}{1.15pt}\par}
      \vspace{0.1em}%
    \end{minipage}%
  }%
  \par\vspace{0.95em}%
}

% ---- Invited talk badge ----------------------------------------------------
\newcommand{\typepill}[3]{%
  \fcolorbox{#2}{#3}{\sffamily\scriptsize\bfseries\color{smbhdeep} #1}%
}
\newcommand{\invitedbadge}{%
  \typepill{Invited}{smbhgold}{smbhgold!20}%
}
\newcommand{\flashbadge}{\typepill{Flash}{auroraglow}{talkwash}}
\newcommand{\talkbadge}{\typepill{Talk}{talkaccent}{talkwash}}
\newcommand{\publicbadge}{\typepill{Public}{posteraccent}{posterwash}}
"""


DOCUMENT_BODY = r"""
\begin{document}

\hypersetup{pageanchor=false}
\begin{titlepage}
  \thispagestyle{empty}
  \pagecolor{titlewash}
  \color{smbhdeep}
  \smbhstarfield
  \vspace*{0.23\textheight}
  \begin{center}
    {\sffamily\bfseries\large\color{auroraglow} SMBH 2026\par}
    \vspace{0.7em}
    {\sffamily\bfseries\Huge\color{smbhdeep} Supermassive Black Holes\par}
    \vspace{0.2em}
    {\sffamily\bfseries\Huge\color{noteblue} and Blue Notes\par}
    \vspace{0.45em}
    {\sffamily\itshape\large\color{smbhteal} Conference Programme\par}
    \vspace{1.2em}
    {\color{auroraglow}\rule{0.48\textwidth}{1.6pt}\par}
    \vspace{1.2em}
    {\large\color{smbhdeep} Campus MIL, Universit\'e de Montr\'eal\par}
    \vspace{0.4em}
    {\normalsize\itshape\color{noteink} Montr\'eal, Qu\'ebec, Canada\par}
    \vspace{0.45em}
    {\large\color{noteblue} June 29 -- July 3, 2026\par}
    \vspace{0.8em}
    {\normalsize\color{smbhteal}\href{https://sites.google.com/view/smbh2026/home}{sites.google.com/view/smbh2026/home}\par}
    \vspace{1.3em}
    \IfFileExists{assets/Symbole_carré-UdeM.png}{%
      \includegraphics[width=3.0cm]{assets/Symbole_carré-UdeM.png}%
    }{}
  \end{center}
  \vfill
\end{titlepage}
\nopagecolor
\hypersetup{pageanchor=true}
\enablesmbhastrobackground

\pagenumbering{roman}
\tableofcontents
\clearpage

\pagenumbering{arabic}

\setchapterpalette{talkaccent}{talkwash}
\sectiondividerpage{Talk Schedule}{talkaccent}{talkwash}{Scheduled oral and flash presentations}
\chapter*{Talk Schedule}
\addcontentsline{toc}{chapter}{Talk Schedule}
\markboth{Talk Schedule}{}

\input{sections/generated_smbh_talk_schedule}

\clearpage
\setchapterpalette{posteraccent}{posterwash}
\sectiondividerpage{Poster Session}{posteraccent}{posterwash}{Poster presentations and contributed poster titles}
\chapter*{Poster Session}
\addcontentsline{toc}{chapter}{Poster Session}
\markboth{Poster Session}{}

\input{sections/generated_smbh_poster_session}

\clearpage
\setchapterpalette{abstractaccent}{abstractwash}
\sectiondividerpage{Talk Abstracts}{abstractaccent}{abstractwash}{Abstracts for scheduled and schedule-TBD talks}
\chapter*{Talk Abstracts}
\addcontentsline{toc}{chapter}{Talk Abstracts}
\markboth{Talk Abstracts}{}

\input{sections/generated_smbh_talk_abstracts}

\clearpage
\setchapterpalette{posterabstractaccent}{posterabstractwash}
\sectiondividerpage{Poster Abstracts}{posterabstractaccent}{posterabstractwash}{Poster abstracts grouped by theme}
\chapter*{Poster Abstracts}
\addcontentsline{toc}{chapter}{Poster Abstracts}
\markboth{Poster Abstracts}{}

\input{sections/generated_smbh_poster_abstracts}

\end{document}
"""


# ---------------------------------------------------------------------------
# Renderers
# ---------------------------------------------------------------------------

def render_talk_schedule(entries: list[Entry], unscheduled: list[Entry]) -> str:
    grouped = _group_by_day_theme(entries)
    parts: list[str] = []

    for day_label in _day_order(entries):
        parts.append(_day_banner(day_label, add_to_toc=True, anchor_prefix="sched-day"))
        for theme_key in _theme_order(grouped[day_label]):
            parts.append(_theme_heading(theme_key, "talkaccent"))
            parts.append(_talk_table(grouped[day_label][theme_key]))

    if unscheduled:
        parts.append(
            "\\daypalettebanner{Schedule TBD}{softpanel}{abstractaccent}{Awaiting final placement in the programme}\n"
        )
        grouped_tbd = _group_by_theme(unscheduled)
        for theme_key in _theme_order(grouped_tbd):
            parts.append(_theme_heading(theme_key, "abstractaccent"))
            parts.append(_talk_table(grouped_tbd[theme_key]))

    return "\n".join(parts) if parts else "No talks were found."


def render_poster_session(entries: list[Entry]) -> str:
    parts: list[str] = []
    grouped_by_day = _group_by_day_theme(entries)

    for day_label in _poster_day_order(entries):
        parts.append(_poster_day_banner(day_label))
        grouped = grouped_by_day[day_label]
        for theme_key in _theme_order(grouped):
            parts.append(_theme_heading(theme_key, "posteraccent"))
            parts.append(_poster_table(grouped[theme_key]))
    return "\n".join(parts) if parts else "No posters were found."


def render_talk_abstracts(entries: list[Entry], unscheduled: list[Entry]) -> str:
    grouped = _group_by_day_theme(entries)
    parts: list[str] = []

    for day_label in _day_order(entries):
        parts.append(_day_banner(day_label, add_to_toc=False, anchor_prefix="abs-day"))
        for theme_key in _theme_order(grouped[day_label]):
            parts.append(_theme_heading(theme_key, "abstractaccent"))
            parts.extend(_abstract_entry(e) for e in grouped[day_label][theme_key])

    if unscheduled:
        parts.append(
            "\\daypalettebanner{Schedule TBD}{softpanel}{abstractaccent}{Talk abstracts without a final day assignment}\n"
        )
        grouped_tbd = _group_by_theme(unscheduled)
        for theme_key in _theme_order(grouped_tbd):
            parts.append(_theme_heading(theme_key, "abstractaccent"))
            parts.extend(_abstract_entry(e) for e in grouped_tbd[theme_key])

    return "\n".join(parts) if parts else "No talk abstracts were found."


def render_poster_abstracts(entries: list[Entry]) -> str:
    grouped = _group_by_theme(entries)
    parts: list[str] = []
    for theme_key in _theme_order(grouped):
        parts.append(_theme_heading(theme_key, "posterabstractaccent"))
        parts.extend(_abstract_entry(e) for e in grouped[theme_key])
    return "\n".join(parts) if parts else "No poster abstracts were found."


def _day_banner(day_label: str, add_to_toc: bool, anchor_prefix: str) -> str:
    day_name = day_name_from_label(day_label)
    bg, accent, glow, date_str = DAY_PALETTES.get(
        day_name, ("softpanel", "auroraglow", "smbhgold", day_label)
    )
    toc_line = (
        f"\\addcontentsline{{toc}}{{section}}{{{programme_tex(day_label)}}}\n"
        if add_to_toc
        else ""
    )
    return (
        f"\\hypertarget{{{anchor_prefix}-{day_name.lower()}}}{{}}\n"
        f"{toc_line}"
        f"\\daypalettebanner{{{programme_tex(day_name)}}}{{{bg}}}{{{accent}}}{{{programme_tex(date_str)}}}\n"
    )


def _poster_day_banner(day_label: str) -> str:
    if day_label == "Unscheduled":
        return (
            "\\hypertarget{poster-day-schedule-tbd}{}\n"
            "\\addcontentsline{toc}{section}{Poster Session: Schedule TBD}\n"
            "\\daypalettebanner{Poster Session}{posterwash}{posteraccent}{Schedule TBD in workbook}\n"
        )
    return _day_banner(day_label, add_to_toc=True, anchor_prefix="poster-day")


def _theme_heading(theme_key: str, accent_color: str) -> str:
    label = theme_label(theme_key)
    return (
        f"\\subsection*{{Theme {programme_tex(theme_key)}: {programme_tex(label)}}}\n"
        f"{{\\color{{{accent_color}}}\\rule{{0.2\\textwidth}}{{1.1pt}}}}\\par\\vspace{{0.45em}}\n"
    )


def _talk_table(entries: list[Entry]) -> str:
    rows: list[str] = []
    for e in entries:
        speaker = programme_tex(e.full_name)
        title = f"\\hyperlink{{abs-{e.row_number}}}{{{programme_tex(e.title)}}}"
        rows.append(
            f"\\hypertarget{{sched-{e.row_number}}}{{}}{speaker} & {_type_badge(e)} & {title} \\\\"
        )

    table_rows = "\n".join(rows)
    col_spec = (
        r"@{}>{\TableRaggedRight\arraybackslash\hspace{0pt}}p{0.24\textwidth}"
        r" >{\centering\arraybackslash}p{0.12\textwidth}"
        r" >{\TableRaggedRight\arraybackslash\hspace{0pt}}p{0.58\textwidth}@{}"
    )
    return (
        "{\n"
        "\\small\n"
        "\\hbadness=10000\n"
        "\\sloppy\n"
        "\\setlength{\\emergencystretch}{1em}\n"
        f"\\begin{{longtable}}{{{col_spec}}}\n"
        "\\toprule\n"
        "Speaker & Type & Title \\\\\n"
        "\\midrule\n"
        "\\endfirsthead\n"
        "\\toprule\n"
        "Speaker & Type & Title \\\\\n"
        "\\midrule\n"
        "\\endhead\n"
        f"{table_rows}\n"
        "\\bottomrule\n"
        "\\end{longtable}\n"
        "}"
    )


def _poster_table(entries: list[Entry]) -> str:
    rows: list[str] = []
    for e in entries:
        presenter = programme_tex(e.full_name)
        title = f"\\hyperlink{{abs-{e.row_number}}}{{{programme_tex(e.title)}}}"
        rows.append(f"\\hypertarget{{sched-{e.row_number}}}{{}}{presenter} & {title} \\\\")

    table_rows = "\n".join(rows)
    col_spec = (
        r"@{}>{\TableRaggedRight\arraybackslash\hspace{0pt}}p{0.26\textwidth}"
        r" >{\TableRaggedRight\arraybackslash\hspace{0pt}}p{0.68\textwidth}@{}"
    )
    return (
        "{\n"
        "\\small\n"
        "\\hbadness=10000\n"
        "\\sloppy\n"
        "\\setlength{\\emergencystretch}{1em}\n"
        f"\\begin{{longtable}}{{{col_spec}}}\n"
        "\\toprule\n"
        "Presenter & Poster Title \\\\\n"
        "\\midrule\n"
        "\\endfirsthead\n"
        "\\toprule\n"
        "Presenter & Poster Title \\\\\n"
        "\\midrule\n"
        "\\endhead\n"
        f"{table_rows}\n"
        "\\bottomrule\n"
        "\\end{longtable}\n"
        "}"
    )


def _abstract_entry(e: Entry) -> str:
    type_tag = _type_badge(e)
    affil = programme_tex(e.affiliation) if e.affiliation else ""
    career = programme_tex(e.career_stage) if e.career_stage else ""
    meta_parts = [p for p in [affil, career] if p]
    meta_line = " \\textperiodcentered{} ".join(meta_parts) if meta_parts else ""
    meta_block = f"\\textbf{{Affiliation:}} {meta_line}\\\\" if meta_line else ""

    return textwrap.dedent(f"""
\\TopicNeedSpace{{10\\baselineskip}}
\\hypertarget{{abs-{e.row_number}}}{{}}
\\subsubsection*{{{programme_tex(e.title)}}}
\\textbf{{Speaker:}} {programme_tex(e.full_name)} {type_tag}\\\\
{meta_block}
{programme_paragraphs(e.abstract)}
""").strip() + "\n"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _type_badge(entry: Entry) -> str:
    if entry.presentation_type == "Invited Speaker":
        return r"\invitedbadge{}"
    if entry.presentation_type == "Invited Speaker (Public Talk)":
        return r"\publicbadge{}"
    if entry.presentation_type == "Flash Talk":
        return r"\flashbadge{}"
    if entry.presentation_type == "Contributed Talk":
        return r"\talkbadge{}"
    if entry.presentation_type == "Poster":
        return r"\typepill{Poster}{posteraccent}{posterwash}"
    label = programme_tex(entry.presentation_type or "Talk")
    return rf"\typepill{{{label}}}{{talkaccent}}{{talkwash}}"


def _group_by_day_theme(entries: list[Entry]) -> dict[str, dict[str, list[Entry]]]:
    result: dict[str, dict[str, list[Entry]]] = defaultdict(lambda: defaultdict(list))
    for entry in entries:
        result[entry.day_label][primary_theme(entry.theme)].append(entry)
    return result


def _group_by_theme(entries: list[Entry]) -> dict[str, list[Entry]]:
    result: dict[str, list[Entry]] = defaultdict(list)
    for entry in entries:
        result[primary_theme(entry.theme)].append(entry)
    return result


def _day_order(entries: list[Entry]) -> list[str]:
    seen: set[str] = set()
    order: list[str] = []
    for entry in entries:
        if entry.day_label not in seen:
            seen.add(entry.day_label)
            order.append(entry.day_label)
    return order


def _poster_day_order(entries: list[Entry]) -> list[str]:
    real_days = [day for day in _day_order(entries) if day != "Unscheduled"]
    if real_days:
        if any(entry.day_label == "Unscheduled" for entry in entries):
            real_days.append("Unscheduled")
        return real_days
    return ["Unscheduled"] if entries else []


def _theme_order(theme_dict: dict[str, list[Entry]] | dict[str, dict[str, list[Entry]]]) -> list[str]:
    def sort_key(theme_key: str) -> tuple[int, str]:
        try:
            return (int(theme_key), "")
        except ValueError:
            return (999, theme_key)

    return sorted(theme_dict.keys(), key=sort_key)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate the styled LaTeX programme for SMBH 2026."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="SMBH 2026 Participant Tracking - MASTER.xlsx",
        help="Path to the XLSX workbook.",
    )
    parser.add_argument(
        "-s", "--sheet",
        default=None,
        help="Worksheet name (default: first sheet).",
    )
    parser.add_argument(
        "--title",
        default="Supermassive Black Holes and Blue Notes",
        help="Conference title for the document.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Input workbook not found: {input_path}", file=sys.stderr)
        return 1

    try:
        entries = build_entries(input_path, args.sheet)
    except Exception as exc:
        print(f"Error reading workbook: {exc}", file=sys.stderr)
        return 1

    talk_entries = [entry for entry in entries if not is_poster_entry(entry)]
    poster_entries = [entry for entry in entries if is_poster_entry(entry)]
    scheduled_talks = [entry for entry in talk_entries if entry.day_label != "Unscheduled"]
    unscheduled_talks = [entry for entry in talk_entries if entry.day_label == "Unscheduled"]

    # Write section files
    sections_dir = Path("sections")
    sections_dir.mkdir(exist_ok=True)

    rendered_sections = {
        sections_dir / "generated_smbh_talk_schedule.tex": render_talk_schedule(
            scheduled_talks, unscheduled_talks
        ),
        sections_dir / "generated_smbh_poster_session.tex": render_poster_session(
            poster_entries
        ),
        sections_dir / "generated_smbh_talk_abstracts.tex": render_talk_abstracts(
            scheduled_talks, unscheduled_talks
        ),
        sections_dir / "generated_smbh_poster_abstracts.tex": render_poster_abstracts(
            poster_entries
        ),
    }
    for path, content in rendered_sections.items():
        path.write_text(content, encoding="utf-8")

    # Write main document
    main_path = Path("smbh_main.tex")
    main_path.write_text(PREAMBLE + DOCUMENT_BODY, encoding="utf-8")

    print(
        "Included "
        f"{len(scheduled_talks)} scheduled talks, "
        f"{len(unscheduled_talks)} talks with TBD placement, "
        f"and {len(poster_entries)} posters."
    )
    for path in rendered_sections:
        print(f"  {path}")
    print(f"  {main_path}")
    print("Compile with:  lualatex smbh_main.tex  (run twice for TOC)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
