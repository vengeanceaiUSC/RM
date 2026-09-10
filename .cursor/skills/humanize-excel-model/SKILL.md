---
name: humanize-excel-model
description: Humanize Excel financial models (remove AI tells) without deleting hardcoded assumptions, Source links, or breaking cross-sheet formulas. Use when cleaning LULU/unbeiesgbar workbooks, fixing model errors after column shifts, or when the user asks to humanize an Excel model.
---

# Humanize Excel Model (Safe)

## Golden rule

**Never delete hardcoded assumptions, Source links, or cross-sheet formulas.** Rewrite in place; do not blank cells.

## Sourcing rule (verbatim)

> Every time you introduce a hardcoded number or assumption, state the source in the adjacent **Source** column (e.g. `FY25 10-K`, `Q2 FY26 guidance`, `FRED DGS10`). No raw URLs in the Notes column; links live in Source only.

Notes column: **≤ 8 words**, analyst shorthand.

---

## Safe pipeline (LULU/unbeiesgbar)

Run from `LULU/` in this order:

```bash
python3 scripts/restore_source_columns.py   # if Notes/Source cols missing
python3 scripts/repair_scenarios_refs.py    # fix $H$→$F$ base-case locks
python3 scripts/remove_comment_artifacts.py  # drop red-triangle comments; reset tall rows
python3 scripts/humanize_workbook_authentic.py
python3 scripts/strip_arrows.py
python3 scripts/audit_ai_tells.py           # read-only — fix flagged items by hand
```

**Do not run** on finished files:
- `strip_model_manual.py` — deletes doc columns
- `humanize_final_model.py` — can strip col C hyperlinks

Read `LULU/AI_SELF_AUDIT.md` before editing.

---

## Column layout after restore

| Sheet | A | B | C | D | E | F | G |
|-------|---|---|---|---|---|---|---|
| WACC / DCF / etc. | Label | Notes | Source | Values | — | — | — |
| Scenarios | Label | Notes | Source | — | Bear | **Base** | Bull |

**Base-case locked assumptions** (rows 4–22) must reference `Scenarios!$F$…`, not `$D$` (stripped) or `$H$` (empty/wrong).

After inserting Notes+Source at col B, verify:
- DCF `E15` = `=Scenarios!$F$12` (D&A %)
- DCF `E17` = `=Scenarios!$F$13` (Capex %)
- Scenarios `F12`=0.045, `F13`=0.055
- Schedule rows `F55`=D&A, `F60`=Capex (formulas, not blank)

---

## What to humanize (safe)

| Fix | Script / action |
|-----|-----------------|
| Metadata `openpyxl` → `Microsoft Excel` | `humanize_workbook_authentic.py` |
| Prompt headers (`Justification (~20 words)`) | same |
| Robotic 15-digit floats | same (round days to 1 dp, rates to 2–3 dp) |
| Blue inputs / black formulas / green cross-sheet | same |
| Unicode `→` arrows | `strip_arrows.py` |
| AI phrasing in Notes | hand-edit; keep Source link |
| Duplicate Analyst comments (>3 identical) | tailor first sentence; keep URL |

## What NOT to touch

- Hardcoded blue inputs (rates, margins, days, shares)
- Source column hyperlinks (FRED, 10-K, NASDAQ, Damodaran)
- Cross-sheet formulas (`=WACC!…`, `=Scenarios!…`)
- Analyst hover comments containing real URLs

---

## Pre-save checklist

1. [ ] No `Scenarios!$H$` refs on assumption rows (run `repair_scenarios_refs.py`)
2. [ ] D&A % and Capex % pull from col F, not empty col H
3. [ ] Every blue hardcode has Source label or link
4. [ ] Notes ≤ 8 words
5. [ ] Creator ≠ `openpyxl`
6. [ ] Spot-check: WACC chain, Scenarios F9→WACC, DCF Gordon TV (D45)

---

## If numbers look missing

1. Run read-only audit: `python3 scripts/audit_ai_tells.py`
2. Compare formula refs to `unbeiesgbar2model.xlsx` (stripped) — base `$D$` maps to restored `$F$`
3. Run `repair_scenarios_refs.py` — fixes the common post-restore column bug
4. Confirm values exist on Scenarios col F before blaming deletion
