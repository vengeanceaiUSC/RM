#!/usr/bin/env bash
# Idempotent Cloud Agent setup for the RM investment-research toolkit.
#
# Provides:
#   - LibreOffice (Calc + Impress) and the python3-uno bridge, used to
#     recalculate the live Excel models (LULU/scripts/recalc_workbook.py) and
#     to export decks/workbooks to PDF (soffice --headless --convert-to pdf).
#   - Python libraries (python-pptx, openpyxl, reportlab, PyMuPDF).
#
# Safe to run repeatedly and against a snapshot that already has these deps.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

# --- System dependencies (LibreOffice + UNO) --------------------------------
# Only attempt apt when soffice is missing and sudo is usable, so the script
# stays a no-op on snapshots that already bundle LibreOffice.
if ! command -v soffice >/dev/null 2>&1; then
  if command -v sudo >/dev/null 2>&1 && sudo -n true >/dev/null 2>&1; then
    export DEBIAN_FRONTEND=noninteractive
    sudo apt-get update -y
    sudo apt-get install -y --no-install-recommends \
      libreoffice-calc libreoffice-impress python3-uno fonts-liberation
  else
    echo "WARNING: soffice not found and sudo unavailable; skipping LibreOffice install." >&2
    echo "         Workbook recalculation and PDF export will not work." >&2
  fi
fi

# --- Python libraries -------------------------------------------------------
# Installed into the user site (--break-system-packages) so the system python3
# — which also owns the python3-uno bridge — can import everything.
python3 -m pip install --break-system-packages --no-cache-dir \
  -r "${REPO_ROOT}/requirements.txt"

echo "RM environment ready:"
soffice --version 2>/dev/null || true
python3 -c "import pptx, openpyxl, reportlab, pymupdf; print('python libs OK')"
python3 -c "import uno; print('uno bridge OK')" 2>/dev/null || \
  echo "NOTE: uno bridge unavailable (LibreOffice not installed)." >&2
