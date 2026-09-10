# AI self-audit — unbeiesgbar_final.xlsx

**Read this before editing the model.** Automated scripts catch metadata, prompt headers, arrow
characters, and long floats. Many AI tells survive those passes. Your job is to find and fix
what scripts miss — **without deleting hardcoded assumptions or Source links.**

Run the read-only scanner first:

```bash
cd LULU && python3 scripts/audit_ai_tells.py
```

---

## Never delete or weaken

| Keep | Why |
|------|-----|
| **Hardcoded assumption values** (blue inputs: rates, margins, days, shares, prices) | These are the model; removing them breaks the DCF |
| **Source column (C / J) hyperlinks** | FRED, 10-K, NASDAQ, Damodaran, PitchBook, earnings release |
| **Source labels** even if text-only (e.g. "Revenue Drivers tab") | Shows provenance; replace wording, do not clear |
| **Cross-sheet formulas** (`=WACC!…`, `=Scenarios!…`) | Model logic |
| **Analyst hover comments that contain a real URL** | Shorten prose around the URL; keep the URL |

If unsure, **rewrite in place** — do not blank the cell.

---

## Sourcing rule (use on every new hardcode)

> Every time you introduce a hardcoded number or assumption, state the source in the adjacent
> **Source** column (e.g. `FY25 10-K`, `Q2 FY26 guidance`, `FRED DGS10`). No raw URLs in the
> Notes column; links live in Source only.

Notes column: **≤ 8 words**, analyst shorthand (`Current 10Y UST`, not `FRED DGS10 anchor (4.8%).`).

---

## What scripts already fix

- `humanize_workbook_authentic.py` — metadata, prompt headers, rounding, blue/black/green fonts
- `strip_arrows.py` — Unicode `→`
- `restore_source_columns.py` — visible Notes + Source columns (+ auto `repair_scenarios_refs`)
- `restore_unaltered_numbers.py` — copies exact hardcodes from `model18unaltered (12).xlsx` (never round)
- `repair_scenarios_refs.py` — fixes Scenarios `$H$` → `$F$` base-case locks (D&A, Capex, WACC)
- `humanize_final_model.py` — AI tutorial rows, mangled in-cell URLs (do **not** re-run if it
  would strip col C links)

---

## Subtle tells scripts miss (you must eyeball these)

### 1. Voice and phrasing

Scan **Notes (B)**, **Source (C)**, **labels (A)**, and **Analyst comments** for:

- Template openers: `anchor.`, `overlay`, `read-through`, `high-conviction`, `durable`,
  `margin-of-safety`, `operating assumption with`, `model cross-reference`
- Pipeline jargon: `Phase 1–5`, `Convention A`, `5-phase pipeline`, `agent workflow`
- Prompt debris: `~20 words`, `click +`, `Ctrl+F`, `cols B–D`, `Source (click)`, `Equation`
- Consulting tone: `underwrite`, `hangs together`, `valuation anchor`, `sanity check`,
  `not the exit`, `pitch deck reads`

**Fix:** One-line analyst shorthand. Keep the Source link; trim the comment to source + one fact.

### 2. Duplicate comment templates

If the same comment body appears on **>3 cells** (e.g. every 10-K row starts with
`FY25 10-K anchor. … on reported baseline. Source: https://…`), it reads like a batch script.

**Fix:** Tailor the first sentence to the line item (`Lease debt FY25 10-K`, `Share count FY25
10-K`). Keep URL.

### 3. Comments vs visible Source column

Prefer **clickable Source in col C**. Long URLs only in hover comments look like a hidden audit
trail.

**Fix:** Ensure col C has the link; shorten comment to one line or remove duplicate URL from
comment (keep comment if it adds line-specific detail).

### 4. Label polish

- `Phase 3:` / `(Phase 1)` in row labels → plain English (`Store EBIT margin %`)
- ALL CAPS section headers unless the rest of the book uses them
- Slash-notes: `/ Alo not a peer pick` → delete or fold into Notes

### 5. Numeric presentation

- Floats with **>3 decimal places** visible in the **sheet grid** (General format)
- Same assumption copied to Bear/Base/Bull with **identical 15-digit float** in all three columns
- Hardcodes that should be **formulas** tied to FY25 (GM %, DSO) but sit as pasted floats

**Fix (safe):** Apply display formats via `humanize_workbook_authentic.py` — days `0.0`, rates
`0.0%`. **Do not round stored values**; that breaks DCF (~$133.64 → ~$132).

**Formula bar tell:** Selecting a cell still shows full precision (e.g. `6.267883648875038`).
That is expected until you manually re-type `6.3` in native Excel **after** F9 confirms the
implied price. Never batch-round via script.

**Submission metadata:** Creator shows `Microsoft Excel` until you Save As from your Office profile.

### 6. Formatting consistency

- Inputs not blue, formulas not black/green
- Source links not blue underlined
- Mixed fonts (Calibri vs Garamond) on one tab
- Percent stored as `0.059714…` with **General** format instead of `0.0%`

**Fix:** Re-run `humanize_workbook_authentic.py` for colors/numfmt; do not clear values.

### 7. Structural AI tells

- Collapsible outline groups (`+` on row headers)
- Extra check columns (`IF(MAX(ABS…))` guardrails)
- Hidden memo rows, beta walkthrough blocks, Alo build essays in column A
- Tab order or Cover tab listing `[Tab 1]`, `[Tab 2]`

**Fix:** Delete **instructional rows only**, not assumption rows.

### 8. File metadata (submission)

Script sets Creator to `Microsoft Excel`. For final submission, **paste all sheets into a new
workbook saved from native Excel on your machine** so Last Modified By is your user profile.

### 9. Cover / README drift

Cover tab should not reference `LULU_Assumptions_Memo.pdf`, GIS assignment boilerplate, or
font-convention essays. Keep the four live source links (EDGAR, 10-K, earnings, NASDAQ).

---

## Safe edit checklist (before you save)

1. [ ] No prompt strings in any header cell
2. [ ] Every blue hardcode on WACC / Scenarios has a Source label or link
3. [ ] No URL as the **visible** cell value (URLs in Source col or comment only)
4. [ ] Notes ≤ 8 words per line (except Comps football-field lo/hi pairs)
5. [ ] DSO / DIO / DPO **display** one decimal via number format; stored values match unaltered
6. [ ] No identical Analyst comment pasted on >3 unrelated rows
7. [ ] Creator ≠ `openpyxl`
8. [ ] Spot-check: WACC chain, Scenarios F9 → WACC, DCF → Scenarios base column

---

## Safe pipeline order

```bash
cd LULU
python3 scripts/restore_source_columns.py   # if Notes/Source cols missing
python3 scripts/repair_scenarios_refs.py    # fix D&A/Capex/WACC col refs
python3 scripts/restore_unaltered_numbers.py  # exact unaltered hardcodes — run if DCF ≠ $133.64
python3 scripts/remove_comment_artifacts.py  # red triangles + inflated row heights
python3 scripts/fix_notes_commentary.py        # Notes match model values (8–15 words)
python3 scripts/humanize_workbook_authentic.py  # NO rounding of assumptions
python3 scripts/scrub_hover_comments.py        # shorten AI hover comment boilerplate
python3 scripts/add_analyst_touch.py           # Scratch tab, yellow highlights (optional)
python3 scripts/strip_arrows.py
python3 scripts/audit_ai_tells.py           # read-only report — fix flagged items by hand
```

Do **not** run `strip_model_manual.py` or `humanize_final_model.py` on the finished file unless
you confirm they will not strip col C links.
