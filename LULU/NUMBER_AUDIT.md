# LULU — Number Audit (what still doesn't tie)

*I ran the pitch against the live DCF + 3-statement workbooks after the $134 / 9% WACC refresh. Here's what I would flag in diligence — fixed items struck through, open items still need a call.*

---

## TL;DR — what I'd say in the room

We fixed the headline valuation drift ($124 → **$134**, bear/bull, WACC). **Good.** Round 3 found **structural** issues: **3-statement vs DCF forecast drift**, **$500M vs 75%-UFCF buybacks**, a **bogus $167 "DCF exit" football-field row**, **two EV/EBITDA tapes (3.5x vs 4.7x)**, and a **broken earnings scrape** in `ingested_kpis.json`. See items 11–22 below.

---

## Fixed this pass (were wrong, now match model)

| Item | Was | Now |
|------|-----|-----|
| Base DCF | ~$124 | **~$134** |
| Bear / bull | $66 / $227 | **$56 / $202** |
| Prob-weighted | ~$155 | **~$131** |
| WACC (base) | ~10.5% | **~9.0%** (β ~0.84, lease-adjusted) |
| Football field DCF | $91–$239 | **$56–$202** |
| FY2028E / FY2030E revenue | $10,845 / $11,730 | **$10,942 / $11,452** |

Pitch deck + thesis outline + README now pull from `pitch_values.json`. Rebuild: `python3 pitch_values.py` → `python3 build_pitch.py`.

---

## Still open — fix or footnote before submission

### 1. Three different "clean" margins (we use all three; we don't label them)

This is the biggest narrative risk. We have **three** defensible EBIT margin numbers for FY26 and we swap labels:

| Label | Value | What it is |
|-------|-------|------------|
| **Q2 run-rate clean** | **13.2%** | Q2 18.8% OM minus 560 bps tariff boost — **this is what Scenarios / DCF underwrites as the trough** |
| **FY26 ex-refund (model)** | **~13.9%** | 3-statement EBIT $1,587M − $134.5M refund ÷ $10,425M sales |
| **FY26 reported (model)** | **15.2%** | Full IS line — includes the one-time refund |

**Problem:** Thesis summary and README still say **"~14.5% reported"** in places. The model prints **15.2%**. That's not rounding — it's the old shorthand before we reconciled the IS.

**What I'd say:** *"DCF trough is 13.2% run-rate. Reported FY26 OM in our model is 15.2% because we flow the $134.5M refund through COGS once. Don't mix 14.5% into the script."*

---

### 2. Bear case does **not** bracket $100

Pitch Thesis I still says bear "brackets the current price." **It doesn't.** Bear implied is **~$56** (−44% vs $100). Only base ($134) and bull ($202) sit above.

**What I'd say:** *"Downside is cushioned by net cash and FCF yield — not because bear DCF is near $100. Bear is a real impairment case."*

---

### 3. Gordon exit multiple: deck says ~6.0x, model says **~7.4x**

README, comps slide, and thesis appendix still cite **~6.0x** exit EV/EBITDA. Live DCF (Gordon TV ÷ FY30 EBITDA) = **7.36x** at base WACC 9.0% and g 2.25%.

`build_dcf.py` Comps tab commentary still says "At base WACC 10.5% … ~6.0x" — that's **double stale** (wrong WACC *and* wrong multiple).

**What I'd say:** *"Selected terminal value is the Gordon identity — today that's ~7.4x on FY30 EBITDA, not 6.0x and not PitchBook NKE 12.7x."*

---

### 4. Sensitivity grid center ≠ scenario base DCF

| Method | Implied $/sh @ ~9% WACC, g 2.25% |
|--------|-----------------------------------|
| Scenarios tab (base) | **$134** |
| DCF sensitivity (9.5% row, 2.25% col) | **~$124** |
| DCF sensitivity (10.0% row, 2.25% col) | **~$116** |

Same workbook, two answers. Likely path/timing differences between the scenario engine and the static sensitivity table — **not a pitch typo**, but a **Q&A landmine**.

**What I'd say:** *"Base case is the Scenarios column G output ($134). The sensitivity grid is a WACC×g bracket — use it for direction, not as the headline value."*

---

### 5. P/E 10.4x is on **guide** EPS, not model EPS

Comps tab uses **$9.61** guide midpoint ($9.48–$9.73). At $100 → **10.4x**. Our 3-statement prints **$10.02** EPS → **~10.0x**.

Pitch mixes **"~10x"** and **"10.4x"** in the same deck. Both are defensible; pick one convention and footnote.

**What I'd say:** *"Trading multiple slide uses company guide EPS ($9.61). Our model EPS is $10.02 because we don't plug other income to hit the guide."*

---

### 6. Repurchase schedule still uses **10.5%** CoE growth — WACC CoE is **~9.9%**

DCF repurchase block: `Cost of equity — share-price growth rate = 10.5%`. WACC tab CAPM CoE = **9.86%**. Thesis already says ~9.9% for repurchase — **that's wrong vs the sheet**.

Convention A is intact (IV on **111.4M basic**, repurchase doesn't move IV/share). The **price path** for accretion math is just on a stale CoE assumption.

---

### 7. WACC slide EV line — formula wording is backwards

Still says *"Enterprise value ≈ equity value less cash; ~$9.3B EV at ~$100."*

**~$9.3B market EV is right** (Comps memo ~3.45x × FY25 EBITDA). The formula is sloppy — for a net-cash name it's **market cap minus cash** (≈ $11.1B − $1.8B), not "equity less cash" without defining terms.

**Intrinsic DCF EV is ~$14.9B** — don't confuse market EV at $100 with DCF EV.

---

### 8. Equity bridge wording in thesis

Slide 20 says *"EV ~$14.9B + cash → equity ~$14.9B."* The model actually does **EV + cash − lease debt ($1.8B) ≈ equity**. Leases net out — that's why equity ≈ EV, but the **mechanics** matter if someone asks about ASC 842.

---

### 9. Stale strings still in model builder (flow into Assumptions Memo on rebuild)

| File | Stale text |
|------|------------|
| `build_dcf.py` | "base WACC 10.5% … ~6.0x" |
| `data.py` | `sens_wacc` Ctrl+F proof references "10.5% base"; `sens_wacc` doc says "10.0% base" |
| `repurchase_schedule.py` | `COE_GROWTH = 0.105` hardcoded |

Rebuild `build_dcf.py` after fixing or the PDF memo reintroduces wrong language.

---

### 10. FCF yield "8–9%" — specify the year

| Basis | FCF yield @ $100 mkt cap |
|-------|--------------------------|
| FY2025 actual (~$922M FCF) | **~8.3%** |
| FY2026E model (~$1,067M FCF) | **~9.6%** |

**"8–9%"** is fine as a range if we mean FY25–FY26. If someone uses FY26 only, we're understating.

---

## What ties cleanly (don't re-litigate)

- FY26 revenue **$10,425M**, cash **$2,375M → $5,648M** by FY30 (3-statement)
- Trading **~3.5x EV/EBITDA** and **~10.4x FY26E P/E** (Comps memo rows 11–12)
- **$140 target** = partial re-rating above DCF — intentional, not model output
- Convention A: SBC in EBIT, IV on **111.4M basic shares**
- Channel margins (18.6% / 23.6% / 9.1%) = EBIT ÷ **that channel's revenue**

---

## My rebuild order before GIS

```bash
cd LULU/scripts
python3 build_3statement.py
python3 build_dcf.py
python3 pitch_values.py
python3 build_pitch.py
```

Then Ctrl+F the deck for: **124, 66, 227, 155, 10.5%, 6.0x, 14.5% reported, bear brackets**.

---

*Last checked against recalculated workbooks — Sep 2026.*

---

## Round 3 — deeper cuts (the stuff that only shows up if you open both workbooks)

### 11. **3-statement and DCF diverge after FY26 — and the pitch uses both**

FY26 ties (`$10,425M`). After that, the **Income Statement** and **DCF / Scenarios** paths **do not match**:

| Year | 3-statement revenue | DCF / Scenarios revenue | Gap |
|------|----------------------:|------------------------:|----:|
| FY2027E | $10,696M | $10,665M | ~$31M |
| FY2028E | $10,942M | $10,909M | ~$32M |
| FY2029E | $11,194M | $11,161M | ~$33M |
| FY2030E | $11,452M | $11,418M | ~$33M |

**Why it matters:** `pitch_values.json` pulls **financial tables from the 3-statement** but **valuation from DCF / Scenarios**. A PM comparing slide 13 revenue to slide 18 FCF build is comparing **two different forecast engines**.

**What I'd say:** *"IS tables are the operating model. DCF revenue is the Scenarios column G path — they share FY26, then drift ~30M/yr. Either link them or footnote the split."*

---

### 12. **FY27 growth: 2.6% in the 3-statement, 2.3% in Scenarios**

- **3-statement:** FY27 revenue growth = **+2.6%** (tracks StockAnalysis next-year **+2.64%**)
- **Scenarios / DCF FY27–30:** flat **+2.3%** every year (3Y forecast **2.26%**)

We say "2.3% FY27–30" in the pitch. The **IS table on slide 13 contradicts that** in the first recovery year.

---

### 13. **EBIT FY26: $1,587M (3-statement) vs $1,511M (DCF)**

Same year, two EBIT numbers (~$76M gap). DCF uses the Scenarios EBIT build (clean margin path + refund); the 3-statement has its own GM / SG&A stack.

Pitch **IS slide shows $1,587M**. DCF **values off $1,511M**. Nobody will catch this unless they reconcile tabs — but it's real.

---

### 14. **Buybacks: $500M (3-statement) vs ~$796M (DCF repurchase schedule)**

| Source | FY26 buyback assumption |
|--------|-------------------------|
| **3-statement CF** | **$500M** fixed annual |
| **DCF repurchase block** | **75% × UFCF ≈ $796M** |

Thesis says *"$500M/yr buybacks in projections."* DCF repurchase schedule says **75% of ~$1.06B UFCF**. Those are **different capital-return stories**.

Convention A (IV on 111.4M basic) is fine — but **don't cite both numbers without explaining which model you're in**.

---

### 15. **Football field has a hidden $167 "DCF exit" row that is NOT the DCF**

Comps tab includes **"DCF exit method (Gordon-implied on FY2030E EBITDA)" → ~$167/sh**.

That row uses `(FY30 EBITDA × 7.4x + cash) ÷ shares` — **no discounting, no explicit-period FCF, no lease bridge**. It's a **terminal capitalization shortcut**, not the unlevered DCF output (**$134**).

**We don't show $167 on the pitch football field** (good), but it's in the workbook. If a judge opens Comps, they'll ask why DCF says $134 and "DCF exit" says $167.

**What I'd say:** *"Full DCF is $134. The $167 row is an undiscounted terminal check — ignore it or relabel it."*

---

### 16. **Two "current" EV/EBITDA prints: 3.5x vs 4.7x**

| Source | Multiple | Denominator |
|--------|----------|-------------|
| **Comps memo** (trading @ $100) | **~3.45x** | FY25 EBITDA ($2.71B) |
| **PitchBook LULU comp row** | **~4.7x** | TTM EBITDA (daily tape) |

Pitch uses **~3.5x** on valuation slides. Comps slide shows **LULU at 4.7x**. Both are in the model — **different EBITDA bases**.

**What I'd say:** *"3.5x is our FY25 EBITDA at $100. 4.7x is PitchBook TTM. Use 3.5x for the trough narrative; cite 4.7x only when talking about the pubcomp tape."*

---

### 17. **Americas is 70.7%, not 71%**

10-K segment: **$7,847M / $11,103M = 70.7%**. Pitch and thesis say **~71%**. Revenue Drivers comp formula hardcodes **71% / 16% / 13%** — weights don't match geo (**70.7% / 15.8% / 13.5%**).

Rounding is fine in the room; **don't defend 71% as sourced** when the model says 70.7%.

---

### 18. **Channel mix ≠ geo mix (easy to talk past each other)**

| Lens | FY25 split |
|------|------------|
| **Channel** | Stores 45.5% · E-comm 44.3% · Other 10.2% |
| **Geography** | Americas 70.7% · China 15.8% · RoW 13.5% |

Thesis channel table ($5.05B stores / $4.92B e-comm) is **not** Americas 71%. Macro overlap slide is **geo**. Fine — but **don't imply Americas = stores**.

---

### 19. **`ingested_kpis.json` has a bad earnings scrape**

```json
"fy2026_rev_guide": { "low": 2290000, "high": 2320000 }
```

That's **~$2.3B** — looks like **Q3 quarterly revenue** misparsed as FY guide. The `-10% to -11%` block next to it is **Q3**, not FY26.

FY guide is **$10.35–$10.50B**. Fix `parse_earnings()` in `ingest_kpis.py` or this JSON will poison any automation that reads it.

---

### 20. **"Mid-single-digit" recovery vs model +2.3%**

Pitch Thesis II says international restores **"mid-single-digit"** growth. Model is **+2.3%** FY27–30 — that's **low-single-digit**, not mid (4–6%).

Industry TAM at mid-single-digit is fine. **LULU revenue recovery is not.**

---

### 21. **Live `LULU_DCF_Valuation_Model.xlsx` still has stale Comps commentary**

Rows 49 and 53 on the Comps tab still say **"WACC 10.5% … ~6.0x"** even though `build_dcf.py` was patched. **Workbook not rebuilt.**

Run `python3 build_dcf.py` or the Assumptions Memo PDF will keep printing wrong exit language.

---

### 22. **DCF reported EBIT margin row shows 14.5%; 3-statement shows 15.2%**

DCF tab "Reported EBIT margin % (incl. FY26 refund)" = **14.49%**. 3-statement OM row = **15.22%**. Different numerators (DCF EBIT build vs IS COGS/SG&A stack).

Another instance of **same label, different number**.

---

## Updated Ctrl+F before GIS

`124, 66, 227, 155, 10.5%, 6.0x, 14.5% reported, bear brackets, mid-single-digit, 71%, 500M buyback, 167`

---

*Round 3 added Sep 2026.*
