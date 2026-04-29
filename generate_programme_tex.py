#!/usr/bin/env python3
"""Generate a LaTeX conference programme from an XLSX workbook.

This script uses only the Python standard library. It reads the first worksheet
by default, extracts the requested conference columns, and writes a standalone
LaTeX document suitable for PDF generation with pdflatex/xelatex.
"""

from __future__ import annotations

import argparse
import re
import sys
import textwrap
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ET


NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

REQUIRED_HEADERS = [
    "Name",
    "Surname",
    "Email",
    "Career Stage",
    "Affiliation",
    "Presentation Type",
    "Title",
    "Abstract",
    "Day",
    "Theme",
]


@dataclass
class Entry:
    row_number: int
    full_name: str
    email: str
    career_stage: str
    affiliation: str
    presentation_type: str
    title: str
    abstract: str
    day_raw: str
    day_label: str
    day_sort_key: tuple[int, str]
    theme: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a LaTeX conference programme from an XLSX file."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="SMBH 2026 Participant Tracking - MASTER.xlsx",
        help="Path to the XLSX workbook.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="conference_programme.tex",
        help="Path to the LaTeX output file.",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        default=None,
        help="Worksheet name to read. Defaults to the first worksheet.",
    )
    parser.add_argument(
        "--programme-title",
        default="Conference Programme",
        help="Title shown on the front page.",
    )
    return parser.parse_args()


def excel_column_to_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for char in letters:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1


def normalize_header(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def load_workbook_parts(xlsx_path: Path) -> tuple[list[str], dict[str, str], zipfile.ZipFile]:
    archive = zipfile.ZipFile(xlsx_path)
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    sheet_names: list[str] = []
    sheet_targets: dict[str, str] = {}
    sheets = workbook.find("main:sheets", NS)
    if sheets is None:
        archive.close()
        raise ValueError("Workbook does not contain any worksheets.")
    for sheet in sheets:
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        name = sheet.attrib["name"]
        target = rel_map[rel_id]
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        sheet_names.append(name)
        sheet_targets[name] = target
    return sheet_names, sheet_targets, archive


def load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values = []
    for item in root.findall("main:si", NS):
        values.append("".join(node.text or "" for node in item.iterfind(".//main:t", NS)))
    return values


def read_cell(cell: ET.Element, shared_strings: list[str]) -> str:
    inline = cell.find("main:is", NS)
    if inline is not None:
        return "".join(node.text or "" for node in inline.iterfind(".//main:t", NS))

    value = cell.find("main:v", NS)
    if value is None or value.text is None:
        return ""

    raw = value.text
    if cell.attrib.get("t") == "s":
        return shared_strings[int(raw)]
    return raw


def sheet_rows(
    archive: zipfile.ZipFile, sheet_target: str, shared_strings: list[str]
) -> Iterable[tuple[int, dict[int, str]]]:
    root = ET.fromstring(archive.read(sheet_target))
    rows = root.findall(".//main:sheetData/main:row", NS)
    for row in rows:
        row_index = int(row.attrib["r"])
        values: dict[int, str] = {}
        for cell in row.findall("main:c", NS):
            cell_ref = cell.attrib.get("r", "")
            col_index = excel_column_to_index(cell_ref)
            values[col_index] = read_cell(cell, shared_strings)
        yield row_index, values


def excel_serial_to_label(raw: str) -> tuple[tuple[int, str], str]:
    text = normalize_header(raw)
    try:
        serial = int(float(text))
    except ValueError:
        return (10**9, text), text or "Unscheduled"
    date_value = datetime(1899, 12, 30) + timedelta(days=serial)
    label = date_value.strftime("%A, %B %d, %Y")
    return (serial, label), label


def clean_value(value: str) -> str:
    text = value or ""
    text = text.replace("\u200b", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def latex_escape(value: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(replacements.get(char, char) for char in value)


def latex_paragraphs(value: str) -> str:
    parts = [latex_escape(part.strip()) for part in value.split("\n\n") if part.strip()]
    return "\n\n".join(parts) if parts else "Not provided."


def build_entries(xlsx_path: Path, sheet_name: str | None) -> list[Entry]:
    sheet_names, sheet_targets, archive = load_workbook_parts(xlsx_path)
    try:
        selected_sheet = sheet_name or sheet_names[0]
        if selected_sheet not in sheet_targets:
            raise ValueError(
                f"Worksheet '{selected_sheet}' not found. Available sheets: {', '.join(sheet_names)}"
            )

        shared_strings = load_shared_strings(archive)
        rows = list(sheet_rows(archive, sheet_targets[selected_sheet], shared_strings))
    finally:
        archive.close()

    if not rows:
        raise ValueError("Selected worksheet is empty.")

    header_row_number, header_cells = rows[0]
    header_map = {
        normalize_header(value): index
        for index, value in header_cells.items()
        if normalize_header(value)
    }

    missing = [header for header in REQUIRED_HEADERS if header not in header_map]
    if missing:
        raise ValueError(
            "Missing required columns: " + ", ".join(missing)
        )

    entries: list[Entry] = []
    for row_number, row in rows[1:]:
        values = {
            header: clean_value(row.get(index, ""))
            for header, index in header_map.items()
        }

        title = values["Title"]
        abstract = values["Abstract"]
        if not title and not abstract:
            continue

        full_name = " ".join(part for part in [values["Name"], values["Surname"]] if part).strip()
        day_sort_key, day_label = excel_serial_to_label(values["Day"])
        entries.append(
            Entry(
                row_number=row_number,
                full_name=full_name or "Unknown Speaker",
                email=values["Email"],
                career_stage=values["Career Stage"],
                affiliation=values["Affiliation"],
                presentation_type=values["Presentation Type"] or "Unspecified",
                title=title or "Untitled Presentation",
                abstract=abstract or "Abstract not provided.",
                day_raw=values["Day"],
                day_label=day_label,
                day_sort_key=day_sort_key,
                theme=values["Theme"] or "Unspecified",
            )
        )

    entries.sort(key=lambda entry: (entry.day_sort_key[0], entry.row_number))
    return entries


def render_entry(entry: Entry) -> str:
    meta_parts = []
    if entry.affiliation:
        meta_parts.append(entry.affiliation)
    if entry.career_stage:
        meta_parts.append(entry.career_stage)
    if entry.email:
        meta_parts.append(entry.email)
    meta_line = r" \\ ".join(latex_escape(part) for part in meta_parts) or "Not provided."

    return textwrap.dedent(
        f"""
        \\needspace{{16\\baselineskip}}
        \\begin{{samepage}}
        \\begin{{tcolorbox}}[
          enhanced,
          colback=white,
          colframe=black!18,
          boxrule=0.5pt,
          arc=2pt,
          left=6pt,
          right=6pt,
          top=6pt,
          bottom=6pt
        ]
        {{\\large\\bfseries {latex_escape(entry.full_name)}}}\\hfill{{\\small\\textsc{{{latex_escape(entry.presentation_type)}}}}}

        \\vspace{{0.35em}}
        {{\\bfseries Theme:}} {latex_escape(entry.theme)}

        \\vspace{{0.25em}}
        {{\\bfseries Title:}} {latex_escape(entry.title)}

        \\vspace{{0.25em}}
        {{\\bfseries Affiliation / Stage / Email:}} {meta_line}

        \\vspace{{0.5em}}
        {{\\bfseries Abstract}}

        {latex_paragraphs(entry.abstract)}
        \\end{{tcolorbox}}
        \\end{{samepage}}
        """
    ).strip()


def render_document(entries: list[Entry], programme_title: str) -> str:
    grouped: dict[str, list[Entry]] = defaultdict(list)
    day_order: list[str] = []
    seen_days: set[str] = set()
    for entry in entries:
        grouped[entry.day_label].append(entry)
        if entry.day_label not in seen_days:
            seen_days.add(entry.day_label)
            day_order.append(entry.day_label)

    day_sections = []
    for day_label in day_order:
        blocks = "\n\n".join(render_entry(entry) for entry in grouped[day_label])
        day_sections.append(
            textwrap.dedent(
                f"""
                \\section*{{{latex_escape(day_label)}}}
                {blocks}
                """
            ).strip()
        )

    body = "\n\n".join(day_sections) if day_sections else "No programme entries were found."

    return textwrap.dedent(
        f"""
        \\documentclass[11pt]{{article}}
        \\usepackage[a4paper,margin=1.8cm]{{geometry}}
        \\usepackage[T1]{{fontenc}}
        \\usepackage[utf8]{{inputenc}}
        \\usepackage{{lmodern}}
        \\usepackage{{hyperref}}
        \\usepackage{{xcolor}}
        \\usepackage[most]{{tcolorbox}}
        \\usepackage{{needspace}}
        \\usepackage{{parskip}}
        \\usepackage{{titlesec}}

        \\definecolor{{programmeblue}}{{HTML}}{{17324D}}
        \\definecolor{{programmeaccent}}{{HTML}}{{D9E7F2}}

        \\hypersetup{{
          colorlinks=true,
          urlcolor=programmeblue,
          linkcolor=programmeblue
        }}

        \\titleformat{{\\section}}{{\\Large\\bfseries\\color{{programmeblue}}}}{{}}{{0pt}}{{}}
        \\setlength{{\\parindent}}{{0pt}}

        \\begin{{document}}

        \\begin{{center}}
        {{\\Huge\\bfseries\\color{{programmeblue}} {latex_escape(programme_title)}}} \\\\[0.75em]
        {{\\large Conference schedule extracted from the participant workbook}}
        \\end{{center}}

        \\vspace{{1em}}
        {body}

        \\end{{document}}
        """
    ).strip() + "\n"


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Input workbook not found: {input_path}", file=sys.stderr)
        return 1

    try:
        entries = build_entries(input_path, args.sheet)
        document = render_document(entries, args.programme_title)
    except Exception as exc:  # pragma: no cover - surfaced to CLI
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    output_path = Path(args.output)
    output_path.write_text(document, encoding="utf-8")
    print(f"Wrote {len(entries)} entries to {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
