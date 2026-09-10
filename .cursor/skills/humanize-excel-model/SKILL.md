---
name: humanize-excel-model
description: Humanize Excel financial models (remove AI tells) without changing hardcoded assumptions, Source links, or DCF output numbers. Use when cleaning LULU/unbeiesgbar workbooks, fixing model errors, or when the user asks to humanize an Excel model. Never round or alter values from model18unaltered (12).xlsx.
---

# Humanize Excel Model (Safe)

## Golden rule

**Never delete, round, or alter hardcoded assumptions, Source links, or cross-sheet formulas.**

Numeric source of truth: **`model18unaltered (12).xlsx`**. Base-case DCF implied value must stay **~$133.64/sh**.

## Sourcing rule (verbatim)

> Every time you introduce a hardcoded number or assumption, state the source in the adjacent **Source** column (e.g. `FY25 10-K`, `Q2 FY26 guidance`, `FRED DGS10`). No raw URLs in the Notes column; links live in Source only.

Notes column: **≤ 8 words**, analyst shorthand.

---

## Safe pipeline (LULU/unbeiesgbar)

Run from `LULU/` in this order:

```bash
python3 scripts/restore_source_columns.py      # if Notes/Source cols missing
python3 scripts/repair_scenarios_refs.py       # fix $H$→$F$ base-case locks
python3 scripts/restore_unaltered_numbers.py   # restore exact unaltered hardcodes
python3 scripts/remove_comment_artifacts.py    # red triangles + tall rows only
python3 scripts/humanize_workbook_authentic.py # metadata/colors — NO rounding
python3 scripts/strip_arrows.py
python3 scripts/audit_ai_tells.py              # read-only
```

**Do not run** on finished files:
- `strip_model_manual.py` — deletes doc columns
- `humanize_final_model.py` — can strip col C hyperlinks

Read `LULU/AI_SELF_AUDIT.md` before editing.

---

## Never change these numbers (examples)

| Item | Unaltered value | Common bad edit |
|------|-----------------|-----------------|
| Terminal growth | **0.0225** (2.25%) | Rounded to 0.022 |
| DSO | **6.267883648875038** | Rounded to 6.3 |
| DIO anchor | **128.8324100108167** | Rounded to 128.8 |
| DPO | **25.10521290169407** | Rounded to 25.1 |
| Share price input | **$100** | Changed to $100.61 |
| Capex % | **0.055** | Must not delete or relink to empty col |

If DCF implied price drifts from **$133.64**, run `restore_unaltered_numbers.py` immediately.

---

## Column layout after restore

| Sheet | A | B | C | D | E | F | G |
|-------|---|---|---|---|---|---|---|
| WACC / DCF / etc. | Label | Notes | Source | Values | — | — | — |
| Scenarios | Label | Notes | Source | — | Bear | **Base** | Bull |

Base-case locked assumptions (rows 4–22): `Scenarios!$F$…`

---

## What to humanize (safe)

| Fix | Script / action |
|-----|-----------------|
| Metadata `openpyxl` → `Microsoft Excel` | `humanize_workbook_authentic.py` |
| Prompt headers | same |
| Blue/black/green font colors | same |
| Unicode `→` arrows | `strip_arrows.py` |
| Red-triangle hover comments (Source in col C exists) | `remove_comment_artifacts.py` |
| AI phrasing in Notes | hand-edit; keep Source link |

## What NOT to touch

- Hardcoded blue inputs (exact floats from unaltered)
- Source column hyperlinks
- Cross-sheet formulas
- DCF valuation output chain (rows 42–56)

---

## Pre-save checklist

1. [ ] `restore_unaltered_numbers.py` passes (hardcodes match unaltered)
2. [ ] Terminal growth = **0.0225**, share price input = **$100**
3. [ ] No `Scenarios!$H$` refs on assumption rows
4. [ ] DCF E15/E17 link to `$F$12` / `$F$13`
5. [ ] Every blue hardcode has Source label or link
6. [ ] Creator ≠ `openpyxl`
7. [ ] Spot-check implied value ≈ **$133.64** in Excel (F9 recalc)
