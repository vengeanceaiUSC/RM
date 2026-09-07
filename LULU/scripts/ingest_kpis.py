"""Scrape LULU operational KPIs via Firecrawl and write ingested_kpis.json.

Run before build_dcf.py / build_3statement.py:
    python3 ingest_kpis.py

Uses the Firecrawl scrape API (no key required in this environment). Falls back to
the last committed JSON if the scrape fails.
"""
import json
import os
import re
import urllib.request
from datetime import datetime, timezone

import data as D

ROOT = os.path.dirname(__file__)
OUT = os.path.join(ROOT, "..", "data", "ingested_kpis.json")
FALLBACK = OUT

URLS = {
    "10k_fy2025": D.filing_url("FY2025"),
    "earnings_sep2026": D.SOURCES["earnings_sep2026"],
}


def _scrape_markdown(url: str) -> str:
    payload = json.dumps({"url": url, "formats": ["markdown"], "onlyMainContent": True}).encode()
    req = urllib.request.Request(
        "https://api.firecrawl.dev/v1/scrape",
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        body = json.load(resp)
    if not body.get("success"):
        raise RuntimeError(f"Firecrawl failed for {url}: {body}")
    return body["data"]["markdown"]


def _num(s: str) -> int:
    return int(re.sub(r"[^\d]", "", s))


def _parse_table_row(md: str, label: str) -> list[int]:
    """Return numeric cells from a markdown table row (columns are newest FY first)."""
    pat = re.compile(
        rf"\|\s*{re.escape(label)}\s*\|([^\n]+)",
        re.IGNORECASE,
    )
    m = pat.search(md)
    if not m:
        return []
    return [int(x) for x in re.findall(r"([\d,]+)", m.group(1))]


def _assign_years(nums: list[int], labels: list[str]) -> dict:
    """Map extracted numbers to fiscal years (10-K tables list newest year first)."""
    out = {}
    for i, yr in enumerate(labels[: len(nums)]):
        out[yr] = nums[i]
    return out


def _parse_revenue_row(md: str, label: str) -> list[int]:
    pat = re.compile(
        rf"\|\s*{re.escape(label)}\s*\|([^\n]+)",
        re.IGNORECASE,
    )
    m = pat.search(md)
    if not m:
        return []
    nums = re.findall(r"\$\s*\|\s*([\d,]+)", m.group(1))
    if not nums:
        nums = re.findall(r"([\d,]{4,})", m.group(1))
    return [int(x.replace(",", "")) for x in nums]


def parse_10k(md: str) -> dict:
    out = {}
    year_pairs = ["FY2025", "FY2024"]
    year_triple = ["FY2025", "FY2024", "FY2023"]

    for key, label in [
        ("stores_americas", "Americas"),
        ("stores_china", "China Mainland"),
        ("stores_row", "Rest of World"),
        ("stores_total", "Total company-operated stores"),
    ]:
        nums = _parse_table_row(md, label)
        if nums:
            out[key] = _assign_years(nums, year_pairs)

    for key, label in [
        ("revenue_stores", "Company-operated stores"),
        ("revenue_ecomm", "E-commerce"),
    ]:
        nums = _parse_revenue_row(md, label)
        if nums:
            out[key] = _assign_years(nums, year_triple)

    for key, label in [
        ("revenue_geo_americas", "Americas"),
        ("revenue_geo_china", "China Mainland"),
        ("revenue_geo_row", "Rest of World"),
    ]:
        nums = _parse_revenue_row(md, label)
        if nums:
            out[key] = _assign_years(nums, year_triple[: len(nums)])

    m = re.search(r"sales per square foot were \$\s*([\d,]+)", md, re.I)
    if m:
        out["sales_per_sqft"] = {"FY2025": _num(m.group(1))}

    m = re.search(
        r"women's, men's, and accessories.*?(\d+)%.*?(\d+)%.*?(\d+)%",
        md,
        re.I | re.S,
    )
    if m:
        out["mix_women"] = {"FY2025": int(m.group(1)) / 100}
        out["mix_men"] = {"FY2025": int(m.group(2)) / 100}
        out["mix_accessories"] = {"FY2025": int(m.group(3)) / 100}

    comps = {}
    m = re.search(r"Americas comparable sales decreased (\d+)%", md, re.I)
    if m:
        comps["americas"] = {"FY2025": -int(m.group(1)) / 100}
    m = re.search(r"China Mainland comparable sales increased (\d+)%", md, re.I)
    if m:
        comps["china"] = {"FY2025": int(m.group(1)) / 100}
    m = re.search(r"Rest of World comparable sales increased (\d+)%", md, re.I)
    if m:
        comps["row"] = {"FY2025": int(m.group(1)) / 100}
    m = re.search(r"Comparable sales increased (\d+)%\.", md, re.I)
    if m:
        comps["total"] = {"FY2025": int(m.group(1)) / 100}
    if comps:
        out["comp_sales"] = comps

    if "revenue_stores" in out and "sales_per_sqft" in out:
        rev = out["revenue_stores"].get("FY2025")
        spsf = out["sales_per_sqft"]["FY2025"]
        if rev and spsf:
            sqft = round(rev * 1000 / spsf)
            out["total_sqft"] = {"FY2025": sqft}
            n = out.get("stores_total", {}).get("FY2025")
            if n:
                out["avg_sqft_per_store"] = {"FY2025": round(sqft / n)}

    out["total_revenue"] = {"FY2025": D.IS["revenue"]["FY2025"]}
    if "revenue_stores" in out and "revenue_ecomm" in out:
        rs = out["revenue_stores"].get("FY2025")
        ec = out["revenue_ecomm"].get("FY2025")
        if rs and ec:
            out["revenue_other"] = {"FY2025": out["total_revenue"]["FY2025"] - rs - ec}

    # Sanity: channel + other should tie to 10-K total within 1%
    total = out["total_revenue"]["FY2025"]
    parts = sum(
        out.get(k, {}).get("FY2025", 0)
        for k in ("revenue_stores", "revenue_ecomm", "revenue_other")
    )
    if parts and abs(parts - total) / total > 0.02:
        out["_parse_warning"] = f"channel sum {parts} vs IS total {total}"

    return out


def parse_earnings(md: str) -> dict:
    out = {}
    m = re.search(r"decline of (\d+)% to (\d+)%", md, re.I)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        out["fy2026_rev_growth_guide"] = {"low": -hi / 100, "high": -lo / 100, "mid": -(lo + hi) / 200}
    m = re.search(r"\$\s*([\d.]+)\s*billion.*?\$\s*([\d.]+)\s*billion", md, re.I)
    if m:
        out["fy2026_rev_guide"] = {
            "low": int(float(m.group(1)) * 1_000_000),
            "high": int(float(m.group(2)) * 1_000_000),
        }
    return out


def ingest() -> dict:
    result = {
        "ingested_at": datetime.now(timezone.utc).isoformat(),
        "sources": URLS,
        "kpis": {},
        "earnings": {},
    }
    md_10k = _scrape_markdown(URLS["10k_fy2025"])
    result["kpis"] = parse_10k(md_10k)
    try:
        md_er = _scrape_markdown(URLS["earnings_sep2026"])
        result["earnings"] = parse_earnings(md_er)
    except Exception as exc:
        result["earnings_error"] = str(exc)
    return result


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    try:
        data = ingest()
    except Exception as exc:
        if os.path.exists(FALLBACK):
            print(f"Scrape failed ({exc}); keeping existing {FALLBACK}")
            return
        raise
    with open(OUT, "w") as f:
        json.dump(data, f, indent=2)
    print(f"Wrote {os.path.abspath(OUT)}")
    print(f"  stores FY2025: {data['kpis'].get('stores_total', {})}")
    print(f"  e-comm FY2025: {data['kpis'].get('revenue_ecomm', {})}")


if __name__ == "__main__":
    main()
