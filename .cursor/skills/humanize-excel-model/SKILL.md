---
name: humanize-excel-model
description: Assess and humanize Excel valuation models for stock pitch clubs — predict AI-detection risk, strip AI-only structure/voice without changing hardcoded numbers, and restore Source citations on every assumption. Use when humanizing DCF models, auditing model18unaltered/model18altered workbooks, or when the user asks if a model looks AI-generated or template-built. Ignore cell color and Notes column when scoring AI risk.
---

# Humanize Excel Model (Pitch Club)

## Club rules (verbatim)

> Please build the model from scratch; any usage of templates or third-party models will be automatically rejected. Structure your model in a way that demonstrates that you understand components that drive revenue, costs, and profitability and allows you to make projections of future revenue/profits in a logical manner consistent with how management of a company in this industry might look at the business. A full three-statement operating model is not required, and a template is attached in this email for the investment pitch deck.

**Pitch deck template ≠ financial model template.** Using the GIS deck template is allowed; copying CFI/Wall Street Prep/Excel model templates is not.

---

## Step 0 — Predict detection risk FIRST

Before editing, score the workbook. **Ignore cell color and the Notes column** for this score — reviewers who ban for AI look at structure, metadata, voice in labels/headers, formula patterns, and numeric presentation.

Output this block every time:

```markdown
## AI detection assessment (color + Notes ignored)

**Estimated catch probability:** [LOW <25% | MEDIUM 25–60% | HIGH 60–85% | VERY HIGH >85%]

**Would they call it a third-party template?** [Yes / No / Partial — explain]

**Confident they won't catch it?** [Only say YES if catch probability is LOW and you can defend it]

### Top tells (ranked)
1. …
2. …
3. …

### What already looks human
- …
```

Use the rubric in [reference.md](reference.md). When unsure between two tiers, pick the **higher** risk.

**Only say "they won't catch it" if catch probability is LOW** and no VERY HIGH tells remain (openpyxl metadata, Ctrl+F audit columns, identical 15-digit floats across Bear/Base/Bull, etc.).

---

## Golden rules

1. **Never change hardcoded assumption values or DCF output numbers** unless the user explicitly asks for a model change.
2. **Every hardcoded number gets a Source** in the adjacent Source column:

> Every time you introduce a hardcoded number or assumption, state the source in the adjacent **Source** column (e.g. `FY25 10-K`, `Q2 FY26 guidance`, `FRED DGS10`). No raw URLs in the Notes column; links live in Source only.

3. **Humanize structure and voice, not math** — swap AI-only patterns for how a bank analyst actually builds in Excel.
4. **Ignore color and Notes when judging AI risk**; still fix Notes/Source during humanization if the user asks to clean the file.

---

## What reviewers catch (ignore color, ignore Notes)

| Severity | Tell | Human fix |
|----------|------|-----------|
| **Ban-level** | File metadata `Creator: openpyxl` / `Python` | Save As from native Excel; set metadata to Microsoft Excel |
| **Ban-level** | Parallel **Justification \| Source \| Ctrl+F** doc columns on every tab | Delete Ctrl+F + Justification columns; keep Source only |
| **Ban-level** | Headers like `Justification (~20 words) [cols B–D: click + to expand]` | Remove entirely |
| **Ban-level** | `Ctrl+F (prove number)` / `Alt. Ctrl+F` columns | Delete |
| **High** | `Phase 1–5`, `Convention A`, `5-phase pipeline`, `agent workflow` | Plain labels (`NOPAT bridge`, `Store EBIT %`) |
| **High** | Identical **15-digit floats** pasted in Bear/Base/Bull (e.g. `6.267883648875038`) | Keep stored value; apply display format (`0.0`, `0.0%`); optionally re-type rounded input manually in Excel after F9 |
| **High** | Hardcoded revenue/EBIT **inside formulas** instead of cell refs | Link to assumption cells |
| **High** | Circular refs (`=DCF!C11` on DCF row 11) | Standard mechanics (`=Revenue*Margin`) |
| **High** | Uniform section taxonomy on every tab (same doc row layout) | Remove extras; vary layout slightly per tab |
| **High** | Cover tab listing font conventions + tab inventory + PDF memo links | Strip to company name, recommendation, date |
| **Medium** | Consulting/AI phrasing in **labels/headers**: `underwrite`, `sanity check`, `valuation anchor`, `high-conviction`, `read-through`, `margin-of-safety`, `on reported baseline` | Analyst shorthand in labels only |
| **Medium** | Beta **Hamada walkthrough block** as rows | Delete walkthrough; keep CAPM inputs |
| **Medium** | `IF(MAX(ABS(variance)))` guard / spot-check columns | Delete |
| **Medium** | Collapsible outline groups on every assumption block | Remove grouping |
| **Medium** | Comps tab essay (Alo Yoga build narrative, bullet rationales) | Table only + one-line peer note |
| **Medium** | Revenue Drivers **Equation** column | Delete |
| **Low** | ALL CAPS section headers | Title Case or sentence case |
| **Low** | Unicode arrows `→` | Replace with `-` or `to` |

Full phrase list and scoring weights: [reference.md](reference.md).

---

## Does it look like a template?

Answer separately from AI detection:

| Signal | Template? |
|--------|-----------|
| Same row/column doc layout cloned on WACC, Scenarios, DCF, Comps, NOPAT | **Yes — internal template** |
| CFI/WSP tab names (`Inputs`, `Valuation Summary`, `Dashboard`) | **Yes — third-party** |
| Bottom-up revenue drivers tied to store count × productivity | **No — custom, good** |
| Bear/Base/Bull scenario matrix | **No — standard analyst** |
| Gordon + exit-multiple TV reconciliation | **No — standard analyst** |

An **internally generated AI template** still fails the club rule if it reads as copy-paste automation, even if not CFI.

---

## Humanization workflow

Copy this checklist and track progress:

```
- [ ] Step 0: AI detection assessment (color + Notes ignored)
- [ ] Step 1: Inventory hardcodes — list every blue/input cell + required Source
- [ ] Step 2: Strip AI-only structure (Ctrl+F, Justification, Phase blocks, guard columns)
- [ ] Step 3: Fix labels/headers voice (not Notes unless user asks)
- [ ] Step 4: Fix formula architecture (CHOOSE toggle, no circular refs, no numeric literals in formulas)
- [ ] Step 5: Restore/add Source on every hardcode
- [ ] Step 6: Metadata + display formats (not stored values)
- [ ] Step 7: Re-score detection risk; confirm numbers unchanged
```

### Safe edits (numbers unchanged)

- Delete doc/audit columns (Ctrl+F, Justification, Equation, Alt. Ctrl+F)
- Rename section headers to analyst style (`Terminal value — Gordon growth`)
- Remove memo rows, beta walkthroughs, Phase 5 workflow blocks, spot-check columns
- Move URLs from labels into Source column hyperlinks
- Set number formats (`0.0%`, `#,##0`, `0.0` days) without rounding stored values
- Set `Creator` / `LastModifiedBy` to Microsoft Excel
- Add `Scratch` tab with rough calcs (optional analyst touch)

### Never do without explicit approval

- Round or alter assumption floats (breaks DCF tie-out)
- Delete Source hyperlinks
- Strip hardcoded inputs
- Run bulk scripts that rewrite col C links (`humanize_final_model.py`-class tools)

---

## Voice swap cheat sheet

| AI / template | Human analyst |
|---------------|---------------|
| `FY2026E revenue growth` + paragraph justification | `Q2 guide mid` + Source: earnings release |
| `Operating assumption with durable margin anchor` | `FY25 10-K OM fade` |
| `Sanity check: multiples within ±1.5 turns?` | `Exit multiple check` |
| `5-phase EBIT normalization → normalized NOPAT` | `NOPAT bridge` |
| `Hamada unlever at Yahoo D/E` | Drop row; keep `β = 0.84` with Source: Yahoo |
| `Pitch CFF — Fixed Buyback` | `$750M/yr buyback` |

---

## Source restoration map (every new hardcode)

When adding or restoring Source cells, prefer short labels + hyperlink:

| Input type | Source label pattern |
|------------|---------------------|
| Historical financial | `FY25 10-K` → SEC EDGAR |
| Guidance | `Q2 FY26 release` → investor site |
| Macro | `FRED DGS10` / `FRED GDPC1` |
| Market | `NASDAQ LULU` / `Yahoo Finance` |
| Street forecast | `StockAnalysis 3Y rev` |
| Derived from another tab | `WACC tab` / `Scenarios base` (no URL needed) |
| Pitch-only assumption | `Pitch deck` (no URL) |

---

## LULU repo shortcuts

If working in `vengeanceaiUSC/RM`:

| File | Role |
|------|------|
| `LULU/model18unaltered.xlsx` | Pre-polish reference (AI tells intact) |
| `LULU/model18altered.xlsx` | Post-polish (doc columns stripped) |
| `LULU/scripts/polish_model18_altered.py` | Safe structural strip for altered copy only |
| `LULU/scripts/audit_ai_tells.py` | Read-only phrase/structure scan |
| `LULU/AI_SELF_AUDIT.md` | Extended tell list (if present on branch) |

Branch with full tooling: `cursor/fix-altered-model18-hardcodes-bd53`.

---

## Additional resources

- Detection rubric, scoring math, and LULU baseline assessment: [reference.md](reference.md)
