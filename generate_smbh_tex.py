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

\IfFontExistsTF{Palatino}{
  \setmainfont{Palatino}
}{\IfFontExistsTF{TeX Gyre Pagella}{
  \setmainfont{TeX Gyre Pagella}
}{
  \setmainfont{Latin Modern Roman}
}}
\IfFontExistsTF{Avenir Next}{
  \setsansfont{Avenir Next}
}{\IfFontExistsTF{TeX Gyre Heros}{
  \setsansfont{TeX Gyre Heros}
}{
  \setsansfont{Latin Modern Sans}
}}

% ---- Colours ---------------------------------------------------------------
\definecolor{smbhblue}{HTML}{0E2A45}
\definecolor{smbhteal}{HTML}{1E6878}
\definecolor{smbhgold}{HTML}{C8922A}
\definecolor{smbhdeep}{HTML}{07111F}
\definecolor{paperwarm}{HTML}{F7F4EC}
\definecolor{lightrule}{HTML}{D7DEE7}
\definecolor{softpanel}{HTML}{EEF4F5}
\definecolor{nightbg}{HTML}{07111F}
\definecolor{nightpanel}{HTML}{10243A}
\definecolor{starlight}{HTML}{F8F4E8}
\definecolor{nebulablue}{HTML}{1E476E}
\definecolor{nebulateal}{HTML}{2D7C8B}
\definecolor{auroraglow}{HTML}{5FA8B8}
\definecolor{deepviolet}{HTML}{1A2741}
\definecolor{nebulaedge}{HTML}{DDECF2}
% Day palettes
\definecolor{mondaybg}{HTML}{1A2F3A}
\definecolor{mondayaccent}{HTML}{5AB4CC}
\definecolor{mondayglow}{HTML}{A8DDE8}
\definecolor{tuesdaybg}{HTML}{311A2B}
\definecolor{tuesdayaccent}{HTML}{F28A64}
\definecolor{tuesdayglow}{HTML}{F6C177}
\definecolor{wednesdaybg}{HTML}{102E35}
\definecolor{wednesdayaccent}{HTML}{63C7BE}
\definecolor{wednesdayglow}{HTML}{C6F3EE}
\definecolor{thursdaybg}{HTML}{1B2346}
\definecolor{thursdayaccent}{HTML}{90A7FF}
\definecolor{thursdayglow}{HTML}{F0D08F}
\definecolor{fridaybg}{HTML}{2D1A1A}
\definecolor{fridayaccent}{HTML}{E07070}
\definecolor{fridayglow}{HTML}{F5C8A8}

\hypersetup{
  colorlinks=true,
  linkcolor=smbhblue,
  urlcolor=smbhteal,
  pdftitle={SMBH 2026 Conference Programme},
  pdfauthor={SMBH 2026 Organizing Team}
}

\setstretch{1.03}
\setlength{\parindent}{0pt}
\setlength{\LTpre}{0.8em}
\setlength{\LTpost}{0.4em}
\renewcommand{\arraystretch}{1.15}
\arrayrulecolor{lightrule}
\rowcolors{2}{softpanel!75!white}{white}

\IfFileExists{ragged2e.sty}{%
  \newcommand{\TableRaggedRight}{\RaggedRight}%
}{%
  \newcommand{\TableRaggedRight}{\raggedright}%
}
\IfFileExists{needspace.sty}{%
  \newcommand{\TopicNeedSpace}[1]{\Needspace{#1}}%
}{%
  \newcommand{\TopicNeedSpace}[1]{}%
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
  \noindent\fcolorbox{nebulablue}{nightpanel}{%
    \begin{minipage}{0.95\textwidth}
      \vspace{0.25em}%
      {\sffamily\bfseries\Huge\color{starlight} #1\par}
      \vspace{0.35em}%
      {\color{auroraglow}\rule{0.28\textwidth}{1.2pt}}%
      \vspace{0.25em}%
    \end{minipage}%
  }%
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

% ---- Star / nebula background decorations --------------------------------
% Upper-right corner: overlapping translucent nebula circles + small stars
\newcommand{\smbhastrocorner}{%
  \begingroup
  \setlength{\unitlength}{1pt}%
  \begin{picture}(0,0)
    \put(476,-16){\color{nebulaedge!55}\circle*{108}}
    \put(532,-30){\color{auroraglow!38}\circle*{172}}
    \put(596,-104){\color{nebulateal!28}\circle*{240}}
    \put(556,-158){\color{nebulablue!20}\circle*{300}}
    \put(440,-22){\color{starlight!88}\circle*{2.1}}
    \put(468,-86){\color{starlight!72}\circle*{1.4}}
    \put(504,-44){\color{smbhgold!90}\circle*{2.3}}
    \put(530,-118){\color{auroraglow!84}\circle*{1.8}}
    \put(558,-62){\color{starlight!78}\circle*{2.0}}
    \put(588,-136){\color{smbhgold!74}\circle*{1.3}}
    \put(608,-50){\color{starlight!74}\circle*{1.6}}
    \put(630,-110){\color{auroraglow!68}\circle*{1.2}}
    \put(490,-152){\color{starlight!62}\circle*{1.5}}
    \put(516,-178){\color{smbhgold!66}\circle*{1.1}}
  \end{picture}%
  \endgroup
}
% Lower-left corner: matching nebula glow
\newcommand{\smbhastrolowerleft}{%
  \begingroup
  \setlength{\unitlength}{1pt}%
  \begin{picture}(0,0)
    \put(8,6){\color{nebulablue!20}\circle*{260}}
    \put(50,26){\color{nebulateal!26}\circle*{184}}
    \put(18,64){\color{auroraglow!34}\circle*{124}}
    \put(84,14){\color{nebulaedge!32}\circle*{94}}
    \put(24,20){\color{starlight!84}\circle*{1.9}}
    \put(46,72){\color{smbhgold!80}\circle*{1.6}}
    \put(70,36){\color{starlight!72}\circle*{1.4}}
    \put(94,90){\color{auroraglow!76}\circle*{1.8}}
    \put(116,26){\color{starlight!76}\circle*{1.5}}
    \put(138,64){\color{smbhgold!84}\circle*{2.0}}
    \put(150,12){\color{starlight!68}\circle*{1.2}}
    \put(62,108){\color{auroraglow!60}\circle*{1.3}}
  \end{picture}%
  \endgroup
}
% Enable corner decorations on interior pages (call once after \begin{document})
\newcommand{\enablesmbhastrobackground}{%
  \AddToShipoutPictureBG{%
    \AtPageUpperLeft{\smbhastrocorner}%
    \AtPageLowerLeft{\smbhastrolowerleft}%
  }%
}
% Title-page full starfield
\newcommand{\smbhstarfield}{%
  \begingroup
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
  \end{picture}%
  \endgroup
}

% ---- Star / nebula decorations --------------------------------------------
% Full starfield for dark title page
\newcommand{\smbhstarfield}{%
  \begingroup
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
  \end{picture}%
  \endgroup
}
% Upper-right corner: translucent nebula blobs + small stars
\newcommand{\smbhastrocorner}{%
  \begingroup
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
  \endgroup
}
% Lower-left corner: matching nebula glow + small stars
\newcommand{\smbhastrolowerleft}{%
  \begingroup
  \setlength{\unitlength}{1pt}%
  \begin{picture}(0,0)
    \put(8,6){\color{nebulablue!20}\circle*{260}}
    \put(50,26){\color{nebulateal!26}\circle*{184}}
    \put(18,64){\color{auroraglow!34}\circle*{124}}
    \put(84,14){\color{nebulaedge!32}\circle*{94}}
    \put(24,20){\color{starlight!84}\circle*{1.9}}
    \put(38,52){\color{starlight!62}\circle*{1.1}}
    \put(46,72){\color{smbhgold!80}\circle*{1.6}}
    \put(70,36){\color{starlight!72}\circle*{1.4}}
    \put(86,58){\color{auroraglow!66}\circle*{1.0}}
    \put(94,90){\color{auroraglow!76}\circle*{1.8}}
    \put(116,26){\color{starlight!76}\circle*{1.5}}
    \put(130,48){\color{smbhgold!70}\circle*{1.2}}
    \put(138,64){\color{smbhgold!84}\circle*{2.0}}
    \put(150,12){\color{starlight!68}\circle*{1.2}}
    \put(62,108){\color{auroraglow!60}\circle*{1.3}}
  \end{picture}%
  \endgroup
}
% Enable corner decorations on all interior pages
\newcommand{\enablesmbhastrobackground}{%
  \AddToShipoutPictureBG{%
    \AtPageUpperLeft{\smbhastrocorner}%
    \AtPageLowerLeft{\smbhastrolowerleft}%
  }%
}

% ---- Headers/footers -------------------------------------------------------
\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\sffamily\small\color{smbhblue}\bfseries SMBH 2026}
\fancyhead[R]{\sffamily\small\color{smbhdeep}\nouppercase{\leftmark}}
\fancyfoot[L]{\sffamily\small\color{smbhteal} Conference Programme \textcolor{auroraglow}{.}}
\fancyfoot[C]{\IfFileExists{assets/Symbole_carré-UdeM.png}{\includegraphics[height=16pt]{assets/Symbole_carré-UdeM.png}}{}}
\fancyfoot[R]{\sffamily\small\color{smbhdeep}\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\renewcommand{\footrulewidth}{0pt}
\setlength{\headheight}{25pt}

% ---- Day banner command ----------------------------------------------------
\newcommand{\daypalettebanner}[4]{%
  % #1=day name  #2=bg colour  #3=accent colour  #4=date line
  \noindent\fcolorbox{#3}{#2}{%
    \begin{minipage}{0.96\textwidth}
      \vspace{0.2em}%
      {\sffamily\bfseries\Large\color{starlight} #1\par}
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
\newcommand{\invitedbadge}{%
  \fcolorbox{smbhgold}{smbhgold!20}{\sffamily\footnotesize\bfseries\color{smbhgold!70!black} Invited}%
}
"""


DOCUMENT_BODY = r"""
\begin{document}

\hypersetup{pageanchor=false}
\begin{titlepage}
  \thispagestyle{empty}
  \pagecolor{nightbg}
  \color{starlight}
  \smbhstarfield
  \vspace*{0.28\textheight}
  \begin{center}
    {\sffamily\bfseries\Huge\color{starlight} SMBH 2026\par}
    \vspace{0.5em}
    {\sffamily\bfseries\Large\color{auroraglow} Supermassive Black Holes:\par}
    \vspace{0.2em}
    {\sffamily\itshape\large\color{smbhgold} From Seeds to Giants\par}
    \vspace{1.5em}
    {\color{auroraglow}\rule{0.4\textwidth}{1.6pt}\par}
    \vspace{1.5em}
    {\large\color{starlight} June 29 -- July 3, 2026\par}
    \vspace{0.4em}
    {\normalsize\itshape\color{smbhgold!80} Montr\'eal, Qu\'ebec\par}
    \vspace{1.8em}
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

\chapter*{Conference Schedule}
\addcontentsline{toc}{chapter}{Conference Schedule}
\markboth{Conference Schedule}{}

\input{sections/generated_smbh_schedule}

\clearpage
\chapter*{Abstracts}
\addcontentsline{toc}{chapter}{Abstracts}
\markboth{Abstracts}{}

\input{sections/generated_smbh_abstracts}

\end{document}
"""


# ---------------------------------------------------------------------------
# Schedule renderer
# ---------------------------------------------------------------------------

def render_schedule(entries: list[Entry]) -> str:
    """Return LaTeX for the schedule section (all days)."""
    grouped = _group_by_day_theme(entries)
    day_order = _day_order(entries)
    parts: list[str] = []

    for day_label in day_order:
        day_name = day_name_from_label(day_label)
        bg, accent, glow, date_str = DAY_PALETTES.get(
            day_name, ("nightpanel", "auroraglow", "smbhgold", day_label)
        )
        parts.append(
            f"\\hypertarget{{sched-day-{day_name.lower()}}}{{}}\n"
            f"\\daypalettebanner{{{latex_escape(day_name)}}}{{{bg}}}{{{accent}}}{{{date_str}}}\n"
        )

        theme_entries = grouped[day_label]
        theme_order = _theme_order(theme_entries)
        for theme_key in theme_order:
            label = theme_label(theme_key)
            section_id = f"sched-{day_name.lower()}-{theme_key}"
            parts.append(
                f"\\hypertarget{{{section_id}}}{{}}\n"
                f"\\section*{{Theme {theme_key}: {latex_escape(label)}}}\n"
                f"\\addcontentsline{{toc}}{{section}}{{Theme {theme_key}: {latex_escape(label)}}}\n"
            )
            parts.append(_schedule_table(theme_entries[theme_key]))
        parts.append("")

    return "\n".join(parts)


def _schedule_table(entries: list[Entry]) -> str:
    rows: list[str] = []
    for e in entries:
        type_cell = r"\invitedbadge{}" if e.presentation_type == "Invited Speaker" else ""
        speaker = latex_escape(e.full_name)
        title = (
            f"\\hyperlink{{abs-{e.row_number}}}{{{latex_escape(e.title)}}}"
        )
        rows.append(f"\\hypertarget{{sched-{e.row_number}}}{{}}{speaker} & {type_cell} & {title} \\\\")

    table_rows = "\n".join(rows)
    col_spec = (
        r"@{}>{\TableRaggedRight\arraybackslash\hspace{0pt}}p{0.22\textwidth}"
        r" >{\centering\arraybackslash}p{0.09\textwidth}"
        r" >{\TableRaggedRight\arraybackslash\hspace{0pt}}p{0.65\textwidth}@{}"
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


# ---------------------------------------------------------------------------
# Abstracts renderer
# ---------------------------------------------------------------------------

def render_abstracts(entries: list[Entry]) -> str:
    """Return LaTeX for the abstracts section (all days)."""
    grouped = _group_by_day_theme(entries)
    day_order = _day_order(entries)
    parts: list[str] = []

    for day_label in day_order:
        day_name = day_name_from_label(day_label)
        bg, accent, glow, date_str = DAY_PALETTES.get(
            day_name, ("nightpanel", "auroraglow", "smbhgold", day_label)
        )
        parts.append(
            f"\\hypertarget{{abs-day-{day_name.lower()}}}{{}}\n"
            f"\\daypalettebanner{{{latex_escape(day_name)}}}{{{bg}}}{{{accent}}}{{{date_str}}}\n"
        )

        theme_entries = grouped[day_label]
        theme_order = _theme_order(theme_entries)
        for theme_key in theme_order:
            label = theme_label(theme_key)
            parts.append(
                f"\\section*{{Theme {theme_key}: {latex_escape(label)}}}\n"
                f"\\addcontentsline{{toc}}{{section}}{{Theme {theme_key}: {latex_escape(label)}}}\n"
            )
            for e in theme_entries[theme_key]:
                parts.append(_abstract_entry(e))
        parts.append("")

    return "\n".join(parts)


def _abstract_entry(e: Entry) -> str:
    type_tag = r"\invitedbadge{}" if e.presentation_type == "Invited Speaker" else ""
    affil = latex_escape(e.affiliation) if e.affiliation else ""
    career = latex_escape(e.career_stage) if e.career_stage else ""
    meta_parts = [p for p in [affil, career] if p]
    meta_line = " \\textperiodcentered{} ".join(meta_parts) if meta_parts else ""

    return textwrap.dedent(f"""
\\TopicNeedSpace{{10\\baselineskip}}
\\hypertarget{{abs-{e.row_number}}}{{}}
\\subsubsection*{{{latex_escape(e.title)}}}
\\addcontentsline{{toc}}{{subsubsection}}{{{latex_escape(e.title)}}}
\\textbf{{Speaker:}} {latex_escape(e.full_name)} {type_tag}\\\\
{'\\textbf{Affiliation:} ' + meta_line + '\\\\' if meta_line else ''}
{latex_paragraphs(e.abstract)}
""").strip() + "\n"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _group_by_day_theme(entries: list[Entry]) -> dict[str, dict[str, list[Entry]]]:
    result: dict[str, dict[str, list[Entry]]] = defaultdict(lambda: defaultdict(list))
    for e in entries:
        key = primary_theme(e.theme)
        result[e.day_label][key].append(e)
    return result


def _day_order(entries: list[Entry]) -> list[str]:
    seen: set[str] = set()
    order: list[str] = []
    for e in entries:
        if e.day_label not in seen:
            seen.add(e.day_label)
            order.append(e.day_label)
    return order


def _theme_order(theme_dict: dict[str, list[Entry]]) -> list[str]:
    def sort_key(k: str) -> tuple[int, str]:
        try:
            return (int(k), "")
        except ValueError:
            return (999, k)
    return sorted(theme_dict.keys(), key=sort_key)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate CASCA-style LaTeX programme for SMBH 2026."
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
        default="SMBH 2026",
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

    # Separate scheduled from unscheduled
    scheduled = [e for e in entries if e.day_label != "Unscheduled"]
    unscheduled = [e for e in entries if e.day_label == "Unscheduled"]
    if unscheduled:
        print(f"Note: {len(unscheduled)} unscheduled entries will be omitted from the programme.")

    # Write section files
    sections_dir = Path("sections")
    sections_dir.mkdir(exist_ok=True)

    schedule_tex = render_schedule(scheduled)
    abstracts_tex = render_abstracts(scheduled)

    schedule_path = sections_dir / "generated_smbh_schedule.tex"
    abstracts_path = sections_dir / "generated_smbh_abstracts.tex"
    schedule_path.write_text(schedule_tex, encoding="utf-8")
    abstracts_path.write_text(abstracts_tex, encoding="utf-8")

    # Write main document
    main_path = Path("smbh_main.tex")
    main_path.write_text(PREAMBLE + DOCUMENT_BODY, encoding="utf-8")

    print(f"Wrote {len(scheduled)} entries.")
    print(f"  {schedule_path}")
    print(f"  {abstracts_path}")
    print(f"  {main_path}")
    print("Compile with:  lualatex smbh_main.tex  (run twice for TOC)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
