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
python3 scripts/audit_valuation_numbers.py     # verify 0 diffs vs unaltered
python3 scripts/remove_comment_artifacts.py    # red triangles + tall rows only
python3 scripts/fix_notes_commentary.py        # Notes match model values (8–15 words)
python3 scripts/fix_truncated_notes.py         # complete vs. minus. as. endings
python3 scripts/clear_formula_sources.py       # blank Source on internal formula rows
python3 scripts/fix_dcf_tv_reconciliation.py   # DCF Gordon vs Exit recon column refs
python3 scripts/fix_nopat_label_formulas.py     # NOPAT labels (fixes Excel repair error)
python3 scripts/remove_spot_check.py           # clear Scenarios col I artifact
python3 scripts/fix_formula_references.py      # link FY25 rev in formulas (same values)
python3 scripts/humanize_workbook_authentic.py # metadata/colors/numfmt — NO value rounding
python3 scripts/scrub_hover_comments.py        # shorten AI hover comment boilerplate
python3 scripts/add_analyst_touch.py           # Scratch tab, yellow highlights (optional)
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
| Notes that cite wrong % vs model cell | `fix_notes_commentary.py` |
| AI phrasing in Notes | hand-edit; keep Source link |
| Robotic 15-decimal **display** in grid | `humanize_workbook_authentic.py` numfmt (`0.0`, `0.0%`) |
| Truncated Notes (`vs.`, `minus.`, `..`) | `fix_truncated_notes.py` |
| Robotic Source on formula rows | `clear_formula_sources.py` |
| AI hover comment boilerplate | `scrub_hover_comments.py` |
| Too-clean template feel | `add_analyst_touch.py` (Scratch tab, yellow highlights) |

## Formula bar vs display (critical)

Reviewers who **click into cells** see full float precision in the formula bar (e.g.
`6.267883648875038`). **Do not script-round stored values** — that broke DCF ($133.64 → ~$132).

| Layer | Safe fix |
|-------|----------|
| Sheet grid | Number formats: days `0.0`, rates `0.0%`, $ `#,##0` |
| Formula bar | Re-type rounded inputs manually in Excel **after** F9 confirms ~$133.64 |
| File metadata | Save As from native Excel so Last Modified By is your profile |

## Notes column rules (verbatim)

- **8–15 words max** on assumption rows.
- State the **model value actually used** (e.g. **5.5%** Capex), not only the FY26 guide (**7%**).
- If the guide differs from the plug, one line: *"5.5% model rate; FY26 7% capex guide fades down."*
- No truncated sentences (`not.`, `for.`, `÷.`), no `overlay`, `read-through`, `→`, or double periods.
- Write like a bank analyst note, not AI boilerplate.

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
