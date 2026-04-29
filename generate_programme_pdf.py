#!/usr/bin/env python3
"""Generate a simple conference programme PDF from the XLSX workbook.

This script depends only on the Python standard library and the local
`generate_programme_tex.py` module for workbook parsing.
"""

from __future__ import annotations

import argparse
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from generate_programme_tex import Entry, build_entries


PAGE_WIDTH = 595.0
PAGE_HEIGHT = 842.0
MARGIN = 42.0
CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN)
TOP_Y = PAGE_HEIGHT - MARGIN
BOTTOM_Y = MARGIN


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a simple PDF conference programme from an XLSX file."
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
        default="conference_programme.pdf",
        help="Path to the PDF output file.",
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
        help="Title shown on the first page.",
    )
    return parser.parse_args()


def pdf_escape(text: str) -> str:
    safe = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    return safe.encode("cp1252", errors="replace").decode("cp1252")


def estimate_chars(width: float, font_size: float, factor: float = 0.52) -> int:
    return max(24, int(width / (font_size * factor)))


def wrap_text(text: str, width: float, font_size: float) -> list[str]:
    import textwrap

    chars = estimate_chars(width, font_size)
    paragraphs = [part.strip() for part in text.replace("\r", "").split("\n\n")]
    lines: list[str] = []
    for paragraph in paragraphs:
        if not paragraph:
            continue
        for raw_line in paragraph.splitlines():
            raw_line = " ".join(raw_line.split())
            if not raw_line:
                continue
            lines.extend(textwrap.wrap(raw_line, width=chars, break_long_words=False))
        lines.append("")
    if lines and lines[-1] == "":
        lines.pop()
    return lines or [""]


@dataclass
class TextLine:
    text: str
    font: str
    size: float
    indent: float = 0.0
    color: tuple[float, float, float] = (0, 0, 0)


def section_heading(text: str) -> list[TextLine]:
    return [
        TextLine("", "Helvetica", 6),
        TextLine(text, "Helvetica-Bold", 18, color=(0.10, 0.20, 0.32)),
        TextLine("", "Helvetica", 4),
    ]


def render_entry_lines(entry: Entry) -> list[TextLine]:
    lines: list[TextLine] = []
    lines.append(TextLine(entry.full_name, "Helvetica-Bold", 13))
    lines.append(TextLine(f"{entry.presentation_type}  |  Theme {entry.theme}", "Helvetica", 10))
    if entry.title:
        lines.append(TextLine(f"Title: {entry.title}", "Helvetica-Bold", 10))

    meta_parts = []
    if entry.affiliation:
        meta_parts.append(entry.affiliation)
    if entry.career_stage:
        meta_parts.append(entry.career_stage)
    if entry.email:
        meta_parts.append(entry.email)
    if meta_parts:
        lines.append(TextLine(" | ".join(meta_parts), "Helvetica", 9))

    lines.append(TextLine("Abstract", "Helvetica-Bold", 10))
    for wrapped in wrap_text(entry.abstract, CONTENT_WIDTH - 18, 9):
        if wrapped:
            lines.append(TextLine(wrapped, "Helvetica", 9, indent=10))
        else:
            lines.append(TextLine("", "Helvetica", 5))
    lines.append(TextLine("", "Helvetica", 6))
    return lines


def line_height(line: TextLine) -> float:
    if not line.text:
        return max(6.0, line.size)
    return line.size * 1.35


def entry_height(lines: Iterable[TextLine]) -> float:
    return sum(line_height(line) for line in lines) + 10.0


class SimplePDF:
    def __init__(self) -> None:
        self.objects: list[bytes] = []

    def add_object(self, data: bytes) -> int:
        self.objects.append(data)
        return len(self.objects)

    def build(self, pages: list[bytes]) -> bytes:
        font_regular_id = self.add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        font_bold_id = self.add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")

        content_ids: list[int] = []

        for page_stream in pages:
            stream = (
                f"<< /Length {len(page_stream)} >>\nstream\n".encode("latin-1")
                + page_stream
                + b"\nendstream"
            )
            content_ids.append(self.add_object(stream))
        page_ids = [len(self.objects) + index + 2 for index in range(len(content_ids))]
        pages_id = len(self.objects) + 1
        kids = [f"{page_id} 0 R" for page_id in page_ids]

        self.add_object(
            f"<< /Type /Pages /Count {len(page_ids)} /Kids [{' '.join(kids)}] >>".encode("latin-1")
        )
        for page_id, content_id in zip(page_ids, content_ids):
            page_obj = (
                f"<< /Type /Page /Parent {pages_id} 0 R /MediaBox [0 0 {PAGE_WIDTH:.0f} {PAGE_HEIGHT:.0f}] "
                f"/Resources << /Font << /F1 {font_regular_id} 0 R /F2 {font_bold_id} 0 R >> >> "
                f"/Contents {content_id} 0 R >>".encode("latin-1")
            )
            actual_page_id = self.add_object(page_obj)
            if actual_page_id != page_id:
                raise RuntimeError("PDF object numbering drifted while building page tree.")

        catalog_id = self.add_object(f"<< /Type /Catalog /Pages {pages_id} 0 R >>".encode("latin-1"))

        output = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for index, obj in enumerate(self.objects, start=1):
            offsets.append(len(output))
            output.extend(f"{index} 0 obj\n".encode("latin-1"))
            output.extend(obj)
            output.extend(b"\nendobj\n")

        xref_offset = len(output)
        output.extend(f"xref\n0 {len(self.objects) + 1}\n".encode("latin-1"))
        output.extend(b"0000000000 65535 f \n")
        for offset in offsets[1:]:
            output.extend(f"{offset:010d} 00000 n \n".encode("latin-1"))
        output.extend(
            f"trailer\n<< /Size {len(self.objects) + 1} /Root {catalog_id} 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode(
                "latin-1"
            )
        )
        return bytes(output)


def write_text(stream: list[str], x: float, y: float, line: TextLine) -> None:
    font_name = "F2" if "Bold" in line.font else "F1"
    r, g, b = line.color
    safe = pdf_escape(line.text)
    stream.append("BT")
    stream.append(f"/{font_name} {line.size:.2f} Tf")
    stream.append(f"{r:.3f} {g:.3f} {b:.3f} rg")
    stream.append(f"1 0 0 1 {x + line.indent:.2f} {y:.2f} Tm")
    stream.append(f"({safe}) Tj")
    stream.append("ET")


def make_pages(entries: list[Entry], programme_title: str) -> list[bytes]:
    pages: list[bytes] = []
    stream: list[str] = []
    y = TOP_Y
    current_day = None

    def flush_page() -> None:
        nonlocal stream, y
        pages.append("\n".join(stream).encode("latin-1", errors="replace"))
        stream = []
        y = TOP_Y

    title_block = [
        TextLine(programme_title, "Helvetica-Bold", 22, color=(0.10, 0.20, 0.32)),
        TextLine("Conference schedule extracted from the participant workbook", "Helvetica", 11),
        TextLine("", "Helvetica", 10),
    ]
    for line in title_block:
        if line.text:
            write_text(stream, MARGIN, y, line)
        y -= line_height(line)

    for entry in entries:
        if entry.day_label != current_day:
            heading = section_heading(entry.day_label)
            heading_height = sum(line_height(line) for line in heading)
            if y - heading_height < BOTTOM_Y:
                flush_page()
            for line in heading:
                if line.text:
                    write_text(stream, MARGIN, y, line)
                y -= line_height(line)
            current_day = entry.day_label

        block = render_entry_lines(entry)
        needed = entry_height(block)
        if y - needed < BOTTOM_Y:
            flush_page()
            heading = section_heading(entry.day_label)
            for line in heading:
                if line.text:
                    write_text(stream, MARGIN, y, line)
                y -= line_height(line)

        for line in block:
            if line.text:
                write_text(stream, MARGIN, y, line)
            y -= line_height(line)

        stream.append(f"0.85 0.89 0.94 RG {MARGIN:.2f} {y + 4:.2f} m {PAGE_WIDTH - MARGIN:.2f} {y + 4:.2f} l S")
        y -= 8

    if stream:
        flush_page()
    return pages


def main() -> int:
    args = parse_args()
    entries = build_entries(Path(args.input), args.sheet)
    pdf = SimplePDF()
    output = pdf.build(make_pages(entries, args.programme_title))
    Path(args.output).write_bytes(output)
    print(f"Wrote {len(entries)} entries to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
