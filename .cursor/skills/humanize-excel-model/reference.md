# AI Detection Rubric (Excel Valuation Models)

Use this when executing Step 0 of the humanize-excel-model skill. **Ignore cell color and Notes column content** when scoring.

---

## Scoring method

Start at **0 points**. Add weights for each hit visible in labels, headers, structure, formulas, metadata, or Source column (not Notes).

| Weight | Category | Examples |
|--------|----------|----------|
| +25 | **Metadata smoking gun** | `Creator: openpyxl`, `Python`, generator comment |
| +20 | **Audit-trail columns** | Ctrl+F, Justification (~20 words), Alt. Ctrl+F, Equation |
| +15 | **Prompt headers** | `cols B–D: click + to expand`, `Source (click)` |
| +12 | **Pipeline jargon in labels** | Phase 1–5, Convention A, agent workflow |
| +10 | **Identical long floats** across scenario columns | Same 15-digit DSO/DIO in Bear/Base/Bull |
| +10 | **Hardcoded literals in formulas** | `=11102600*0.055` instead of cell refs |
| +10 | **Circular references** | Row references itself via another sheet loop |
| +8 | **Template layout clone** | Same doc-row pattern on ≥4 tabs |
| +8 | **Cover tab tutorial** | Font convention block + tab list + memo links |
| +6 | **Consulting boilerplate in labels** | underwrite, sanity check, valuation anchor, high-conviction |
| +6 | **Guard/check columns** | `IF(MAX(ABS(` variance spot-check |
| +5 | **Collapsible outline groups** on assumption blocks |
| +5 | **Comps narrative essay** | Multi-row peer build story |
| +3 | **ALL CAPS section headers** everywhere |
| +3 | **Unicode arrows** in labels |

### Catch probability tiers

| Points | Tier | Meaning |
|--------|------|---------|
| 0–15 | **LOW** (<25%) | Plausible hand-built; minor polish left |
| 16–35 | **MEDIUM** (25–60%) | Mixed signals; a sharp reviewer asks questions |
| 36–55 | **HIGH** (60–85%) | Multiple ban-level tells; likely flagged |
| 56+ | **VERY HIGH** (>85%) | Near-certain AI/script origin |

**Confident they won't catch it:** only when total ≤15 **and** no single +20 or +25 hit remains.

---

## Template vs custom

| Looks third-party (CFI/WSP) | Looks AI-generated internal template | Looks hand-built |
|------------------------------|--------------------------------------|------------------|
| Tab names: Inputs, WSP, Dashboard | Tab names OK but identical doc cols everywhere | Varied tab layouts |
| Generic ticker-agnostic labels | Company-specific but robotic headers | Messy but consistent voice |
| LBO/debt schedule not needed for DCF | Hamada walkthrough + Gordon identity spelled out | CAPM 5-line build |
| Protected hidden sheets | Cover font-convention tutorial | Cover = name + rec + date |

---

## Phrase patterns (scan labels, headers, Source, comments — not Notes)

```text
prompt debris:      justification | ~20 word | click + | ctrl+f | source (click) | equation | cols B
pipeline jargon:    convention a | phase [1-5] | 5-phase | agent workflow | pipeline
comment boilerplate: high-conviction | read-through | margin-of-safety | operating assumption with
                     model cross-reference | on reported baseline | durable … anchor
consulting tone:    underwrite | hangs together | valuation anchor | sanity check | not the exit
                     pitch deck reads | bottom-up path | stays viable
template anchor:    anchor. | overlay
meta language:      formula | cell ref | hardcod
```

---

## Numeric presentation (not Notes)

Reviewers click cells. These show in the **formula bar** and grid:

| Tell | Why it screams AI/script |
|------|--------------------------|
| `6.267883648875038` stored as General | Computed ratio pasted at full float precision |
| Same float in C, D, E scenario cols | Batch copy, not hand entry |
| `0.0597142857142857` without `%` format | Script output |
| Revenue literal inside `=11102600*…` | Builder hardcoded, not linked |

**Safe fix:** number formats on grid only; re-type rounded inputs manually in Excel after F9 confirms valuation unchanged.

---

## LULU baseline (model18unaltered.xlsx — unaltered branch)

Assessment with color + Notes ignored:

| Hit | Pts |
|-----|-----|
| openpyxl metadata (when not re-saved) | +25 |
| Ctrl+F + Justification cols every tab | +20 |
| Prompt headers (~20 words, click +) | +15 |
| Phase 5 / NOPAT pipeline labels | +12 |
| Identical 15-digit NWC days across scenarios | +10 |
| Cover tutorial + GIS branding block | +8 |
| Hamada beta walkthrough rows | +6 |
| Gordon identity / TV reconciliation prose headers | +6 |
| Collapsible doc column groups | +5 |
| Comps Alo narrative | +5 |

**Estimated total: ~112 → VERY HIGH (>85%)**

**Template?** Partial — not CFI, but a **cloned internal AI layout** across tabs reads as automated template.

**Confident they won't catch it?** **No.**

After `polish_model18_altered.py` (structural strip only, numbers unchanged): ~45–55 → still **HIGH** until metadata fixed, floats display-formatted, formula architecture cleaned, Sources verified.

---

## Source column standard (humanization, not detection)

Every hardcoded input row:

| Col A | Col B (optional) | Col C Source | Col D+ values |
|-------|------------------|--------------|---------------|
| Label | (skip for detection) | `FY25 10-K` + link | input |

Do **not** put audit prose in Source. One label + one hyperlink beats a paragraph.

---

## Pre-submit verification

1. F9 recalc in native Excel — implied price unchanged vs unaltered baseline
2. `Creator` ≠ openpyxl
3. Zero Ctrl+F / Justification columns
4. Every hardcoded input has Source
5. Re-run detection score — target **LOW** or defensible **MEDIUM**
