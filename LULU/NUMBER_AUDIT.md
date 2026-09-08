# LULU — Number Audit (what still doesn't tie)

*I ran the pitch against the live DCF + 3-statement workbooks after the $134 / 9% WACC refresh. Here's what I would flag in diligence — fixed items struck through, open items still need a call.*

---

## TL;DR — what I'd say in the room

We fixed the headline valuation drift ($124 → **$134**, bear/bull, WACC). **Good.** But if a PM Ctrl+F's the deck against the model, they'll still catch **margin definition mixing**, **a bear case that doesn't bracket $100**, **Gordon exit still printed at ~6x when the model is ~7.4x**, and **P/E on guide EPS vs our own EPS**. None of these kill the thesis — they're credibility leaks. Clean them before GIS.

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
