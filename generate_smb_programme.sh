#!/usr/bin/env bash
set -euo pipefail

WORKBOOK="${1:-SMBH 2026 Participant Tracking - MASTER.xlsx}"
TITLE="SMBH 2026 Conference Programme"

# --- Simple card-style LaTeX + PDF (original generators) ---
python generate_programme_tex.py \
  "$WORKBOOK" \
  -o smb_programme.tex \
  --programme-title "$TITLE"

python generate_programme_pdf.py \
  "$WORKBOOK" \
  -o smb_programme.pdf \
  --programme-title "$TITLE"

printf 'Generated %s and %s from %s\n' \
  "smb_programme.tex" \
  "smb_programme.pdf" \
  "$WORKBOOK"

# --- CASCA-2026-style LaTeX (lualatex-ready) ---
python generate_smbh_tex.py "$WORKBOOK"

if command -v lualatex &>/dev/null; then
  printf 'Compiling smbh_main.tex with lualatex...\n'
  lualatex -interaction=nonstopmode smbh_main.tex
  lualatex -interaction=nonstopmode smbh_main.tex   # second pass for TOC/links
  printf 'Generated smbh_main.pdf\n'
else
  printf 'lualatex not found — skipping PDF compilation of smbh_main.tex\n'
fi

# --- Mobile web schedule + QR codes ---
python generate_smbh_web.py "$WORKBOOK"
printf 'Web pages in web/  |  QR codes and door signs in qr/\n'
