"""Sample company data for the GIS pitch template (Summit Outdoor Co.).

Fictional but realistic DTC outdoor apparel — use as a structural example.
Replace ticker, numbers, and narrative when building your own pitch.
"""

TICK = "SUMM"
EXCHANGE = "NYSE"
COMPANY = "Summit Outdoor Co."
COMPANY_INC = "Summit Outdoor Co. inc."
HEADER = f"{COMPANY} ({EXCHANGE}: {TICK})"
CIK = "0001234567"

PRICE = 42.00
TARGET = 58.00
UPSIDE = "+38%"
RECOMMENDATION = "OVERWEIGHT / LONG"
DESCRIPTOR = (
    "A profitable DTC brand trading at trough multiples after a one-year reset — "
    "we see stabilization, not structural decline"
)

# Historical FY (US$ M)
HY = ["FY2022", "FY2023", "FY2024", "FY2025"]
REV = [2840, 3125, 3380, 3512]
GP = [1562, 1750, 1926, 1989]
OI = [398, 512, 574, 492]
NI = [248, 338, 382, 318]
EPS = [2.14, 2.89, 3.24, 2.68]
OM = [oi / rev * 100 for oi, rev in zip(OI, REV)]

# Forecast columns
FC = ["FY2026E", "FY2028E", "FY2030E"]
FC_REV = [3298, 3584, 4012]
FC_OI = [412, 502, 602]
FC_OM = ["12.5%", "14.0%", "15.0%"]
FC_NI = [278, 348, 428]
FC_EPS = ["2.35", "2.94", "3.62"]

CASH = [420, 512, 468, 445]
INV = [312, 298, 325, 358]
TA = [2105, 2288, 2412, 2520]
TL = [890, 912, 945, 968]
TE = [1215, 1376, 1467, 1552]

CFO = [385, 445, 462, 398]
CAPEX = [118, 125, 132, 128]
FCF = [c - x for c, x in zip(CFO, CAPEX)]
BUYBACKS = [85, 120, 185, 142]

# Charts
REV_CHART_LABELS = HY + FC
REV_CHART_VALUES = REV + FC_REV
OM_CHART_LABELS = HY + ["FY26E", "FY28E", "FY30E"]
OM_CHART_VALUES = [round(x, 1) for x in OM] + [12.5, 14.0, 15.0]
GEO_MIX = [("Americas", 68), ("EMEA", 18), ("APAC", 14)]

DCF_BASE = 52
DCF_BEAR = 34
DCF_BULL = 78
