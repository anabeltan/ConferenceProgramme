#!/usr/bin/env bash
set -euo pipefail

WORKBOOK="${1:-SMBH 2026 Participant Tracking - MASTER.xlsx}"
TITLE="SMBH 2026 Conference Programme"

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
