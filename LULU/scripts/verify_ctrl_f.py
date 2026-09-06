"""Verify every quoted Ctrl+F phrase appears in its linked source."""
import re
import sys
import urllib.request
import data as D

UA = "GIS-Research verify@example.com"
FY25_10K = (
    "https://www.sec.gov/Archives/edgar/data/1397187/"
    "000139718726000020/lulu-20260201.htm"
)
CACHE = {}


def fetch(url):
    if url not in CACHE:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=45) as resp:
            CACHE[url] = resp.read().decode("utf-8", errors="replace")
    return CACHE[url]


def plain(html):
    t = re.sub(r"<[^>]+>", " ", html)
    t = re.sub(r"&#160;", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t


def quotes(hint):
    return re.findall(r'"([^"]+)"', hint or "")


def resolve_url(url):
    if not url:
        return None
    if url == D.filing_url("FY2025"):
        return FY25_10K
    return url


def check_hint(name, hint, url):
    if not hint:
        return None
    if not url:
        if name in ("sc_wacc", "sens_axes"):
            return None
        return (name, "no_url", [])
    body = plain(fetch(url))
    missing = [q for q in quotes(hint) if q.lower() not in body.lower()]
    if missing:
        return (name, "MISSING", missing)
    return (name, "OK", quotes(hint))


def main():
    fails = []
    seen = set()

    for mapping, label in ((D.SOURCE_HINT, "SOURCE"), (D.REPORTED_HINTS, "REPORTED"), (D.COVER_HINTS, "COVER")):
        for key, hint in mapping.items():
            url = None
            if key in D.ASSUMPTION_SRC:
                url = resolve_url(D.ASSUMPTION_SRC[key][1])
            elif label == "REPORTED" and key.startswith("10k"):
                url = FY25_10K
            elif key == "earnings_sep2026" or key.startswith("earnings"):
                url = D.SOURCES["earnings_sep2026"]
            elif key == "nasdaq" or key == "nasdaq_quote":
                url = D.SOURCES["nasdaq_quote"]
            elif key == "edgar_xbrl":
                url = D.SOURCES["edgar_xbrl"]
            elif key == "filing_fy2025":
                url = FY25_10K
            elif key in ("wacc_rf", "sc_g", "sens_g"):
                url = D.SOURCES.get("fred_dgs10" if key == "wacc_rf" else "fred_gdpc1")
            elif key in ("wacc_erp", "wacc_tax", "sc_tax", "sens_wacc"):
                url = D.SOURCES.get("damodaran_erp" if "erp" in key or key == "sens_wacc" else "damodaran_tax")
            elif key == "wacc_beta" or key.startswith("comps_") or key == "dcf_exitm":
                src = D.ASSUMPTION_SRC.get(key)
                url = resolve_url(src[1]) if src else D.SOURCES["lulu_stats"]

            sig = (key, hint)
            if sig in seen:
                continue
            seen.add(sig)

            result = check_hint(f"{label}:{key}", hint, url)
            if not result:
                continue
            name, status, detail = result
            print(f"{name}: {status} {detail}")
            if status == "MISSING":
                fails.append(name)

    print(f"\n{len(fails)} failures")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
