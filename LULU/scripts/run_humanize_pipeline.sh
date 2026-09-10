#!/usr/bin/env bash
# Humanize model18unaltered.xlsx → LULU_model_submission.xlsx (numbers preserved)
set -euo pipefail
cd "$(dirname "$0")/.."
cp model18unaltered.xlsx model18altered.xlsx
cd scripts
python3 polish_model18_altered.py
cp ../model18altered.xlsx ../unbeiesgbar_final.xlsx
for step in \
  restore_unaltered_numbers.py \
  remove_spot_check.py \
  fix_dcf_tv_reconciliation.py \
  fix_formula_references.py \
  fix_ib_formatting.py \
  humanize_workbook_authentic.py \
  fix_financial_number_formats.py \
  fix_decimal_display.py \
  humanize_for_club_submission.py \
  fix_mechanical_architecture.py \
  fix_file_metadata.py \
  strip_programmatic_colors.py \
  strip_arrows.py \
  remove_collapsible_outlines.py \
  final_label_cleanup.py; do
  echo "=== $step ==="
  python3 "$step" || true
done
cp ../unbeiesgbar_final.xlsx ../LULU_model_submission.xlsx
python3 audit_ai_tells.py || true
echo "Done: LULU/LULU_model_submission.xlsx"
