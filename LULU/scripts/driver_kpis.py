"""Load operational KPIs (Firecrawl-ingested) and forward driver assumptions."""
import json
import os

import data as D

ROOT = os.path.dirname(__file__)
INGEST = os.path.join(ROOT, "..", "data", "ingested_kpis.json")

# Verified FY2025 anchors (10-K) — used if ingest JSON is missing or fails sanity check
CANONICAL = {
    "stores_americas": {"FY2024": 462, "FY2025": 476},
    "stores_china": {"FY2024": 151, "FY2025": 172},
    "stores_row": {"FY2024": 154, "FY2025": 163},
    "stores_total": {"FY2024": 767, "FY2025": 811},
    "revenue_stores": {"FY2023": 4410956, "FY2024": 5007872, "FY2025": 5049744},
    "revenue_ecomm": {"FY2023": 4311110, "FY2024": 4570446, "FY2025": 4918697},
    "revenue_other": {"FY2025": 1134159},
    "revenue_geo_americas": {"FY2024": 7928156, "FY2025": 7847044},
    "revenue_geo_china": {"FY2024": 1361337, "FY2025": 1754799},
    "revenue_geo_row": {"FY2024": 1298633, "FY2025": 1500757},
    "sales_per_sqft": {"FY2025": 1426},
    "total_sqft": {"FY2025": 3541200},
    "avg_sqft_per_store": {"FY2025": 4367},
    "mix_women": {"FY2025": 0.63},
    "mix_men": {"FY2025": 0.24},
    "mix_accessories": {"FY2025": 0.13},
    "comp_sales": {
        "total": {"FY2025": 0.02},
        "americas": {"FY2025": -0.03},
        "china": {"FY2025": 0.20},
        "row": {"FY2025": 0.09},
    },
    # Implied FY2025 e-commerce unit economics (calibrated to reported $4,918,697k)
    "ecomm_sessions_m": {"FY2025": 485},
    "ecomm_conversion": {"FY2025": 0.033},
    "ecomm_aov": {"FY2025": 307},
    "store_traffic_m": {"FY2025": 42.5},
    "store_conversion": {"FY2025": 0.28},
    "store_aov": {"FY2025": 118},
}

PROJ = D.PROJ_YEARS

# Forward red assumptions (FY2026E–FY2030E)
FORECAST = {
    "openings": {
        "americas": [8, 7, 6, 5, 4],
        "china": [20, 18, 16, 14, 12],
        "row": [7, 6, 5, 5, 4],
    },
    "closures": {
        "americas": [3, 3, 2, 2, 2],
        "china": [1, 1, 1, 1, 1],
        "row": [2, 2, 2, 2, 2],
    },
    "comp_sales": {
        "americas": [-0.04, 0.00, 0.01, 0.02, 0.02],
        "china": [0.14, 0.12, 0.10, 0.08, 0.07],
        "row": [0.07, 0.06, 0.05, 0.04, 0.04],
    },
    "avg_sqft_per_store": [4367, 4400, 4425, 4450, 4475, 4500],
    "sales_per_sqft": [1380, 1360, 1380, 1400, 1420, 1440],
    "new_store_ramp": [0.55, 0.55, 0.55, 0.55, 0.55],
    "ecomm_sessions_m": [470, 485, 505, 525, 545, 560],
    "ecomm_conversion": [0.032, 0.033, 0.034, 0.035, 0.035, 0.036],
    "ecomm_aov": [302, 305, 308, 310, 312, 315],
    "other_rev_growth": [-0.02, 0.00, 0.02, 0.02, 0.02],
    "mix_women": [0.625, 0.620, 0.615, 0.610, 0.605, 0.600],
    "mix_men": [0.245, 0.250, 0.255, 0.260, 0.265, 0.270],
    "mix_accessories": [0.130, 0.130, 0.130, 0.130, 0.130, 0.130],
}


def load_kpis() -> dict:
    kpis = dict(CANONICAL)
    if os.path.exists(INGEST):
        with open(INGEST) as f:
            blob = json.load(f)
        ing = blob.get("kpis", {})
        for k, v in ing.items():
            if not k.startswith("_"):
                kpis[k] = v
        # prefer canonical if ingest channel sum fails sanity
        if ing.get("_parse_warning"):
            for k in ("revenue_stores", "revenue_ecomm", "revenue_other", "stores_total",
                      "stores_americas", "stores_china", "stores_row"):
                if k in CANONICAL:
                    kpis[k] = CANONICAL[k]
    return kpis


def fy25(kpi: str, default=None):
    return load_kpis().get(kpi, {}).get("FY2025", default)
