"""
Chewy Sales Dashboard ETL
=========================
Reads the Chewy "Catalyst_Feline Fresh SKU L52W Report" Excel, builds a JSON
payload, and injects it into chewy_dashboard_template.html -> chewy_dashboard.html.

Data sources:
  - Sheet "FY25_FY26 YTD Data" : flat monthly table (16 mo), source of the time
    series -> Units, Retail $ (Net Sales), Autoship %, PDP OOS%  (per Chewy SKU).
  - Sheet "Summary"            : Chewy's own FY26-YTD vs FY25-YTD retail rollup
    (authoritative), plus the internal Lignetics SKU codes (CT01.. / FW01).
  - Brand Snapshot PDFs        : avg customer rating + top-10 states (latest month).
    These are the only source of brand-level wholesale ("Chewy's Historical
    Purchases"); we keep them as reference only. Per-SKU wholesale $ is computed
    as WHOLESALE_PRICES[part] * units once prices are supplied.

Wholesale prices: fill WHOLESALE_PRICES below (keyed by Chewy part number) and
re-run. Until then wholesale $ renders as pending.
"""
import json
import os
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Chewy Brand Snapshots",
                    "Catalyst_Feline Fresh SKU L52W Report 5.4 (1).xlsx")
TEMPLATE = os.path.join(HERE, "chewy_dashboard_template.html")
OUTPUT = os.path.join(HERE, "chewy_dashboard.html")

# ---------------------------------------------------------------------------
# Wholesale unit prices, keyed by Chewy part number.  wholesale $ = price *
# units_sold.  Each entry: {"base": <through Sep 2026>, "new": <from Oct 2026>}.
# Total changeover to "new" pricing at WHOLESALE_CHANGEOVER (Sep is a blended
# month but uses Aug/base pricing per Lignetics).  None = no price -> pending.
# Prices per Lignetics wholesale sheet (Catalyst only; FF prices TBD).
# ---------------------------------------------------------------------------
WHOLESALE_CHANGEOVER = "2026-10"  # first month billed at "new" pricing
WHOLESALE_PRICES = {
    "241757": {"base": 9.80, "new": 10.29},    # CT01 Healthy 10-lb
    "241758": {"base": 15.75, "new": 16.54},   # CT02 Healthy 20-lb
    "965502": {"base": 23.25, "new": 24.37},   # CT08 Healthy 30-lb
    "241760": {"base": 9.80, "new": 10.29},    # CT03 Unscented 10-lb
    "241761": {"base": 15.75, "new": 16.49},   # CT04 Unscented 20-lb
    "241763": {"base": 9.80, "new": 10.29},    # CT05 Multi-Cat 10-lb
    "241764": {"base": 15.75, "new": 16.52},   # CT06 Multi-Cat 20-lb
    "1633142": {"base": 13.29, "new": 13.95},  # CT19 Pine Pellet 20-lb
    "1665670": {"base": 24.49, "new": 24.49},  # CT09 Sisal Mat
    "1674830": {"base": 16.79, "new": 16.79},  # FW01 Litter Scoop
    "1685430": {"base": 6.19, "new": 6.19},    # CT18 Poop Bags
    "1932182": None,   # Feline Fresh Pine 10-lb (TBD)
    "1932190": None,   # Feline Fresh Pine 20-lb (TBD)
    "1932198": None,   # Feline Fresh Pine 40-lb (TBD)
}


def wholesale_price(part, month):
    """Per-unit wholesale price for a part in a given 'YYYY-MM' month."""
    p = WHOLESALE_PRICES.get(part)
    if not p:
        return None
    return p["new"] if month >= WHOLESALE_CHANGEOVER else p["base"]

# Latest-month brand snapshot (from the June 2026 Brand Snapshot PDFs).
SNAPSHOT = {
    "Catalyst": {
        "month": "June 2026",
        "rating": 4.19,
        "top_states": [["CA", 1293], ["PA", 762], ["NY", 761], ["TX", 689],
                       ["FL", 679], ["OH", 578], ["NC", 457], ["MI", 450],
                       ["IL", 414], ["VA", 404]],
    },
    "Feline Fresh": {
        "month": "June 2026",
        "rating": 4.02,
        "top_states": [["VA", 99], ["CA", 98], ["OH", 63], ["IL", 50],
                       ["NY", 50], ["PA", 49], ["FL", 42], ["TX", 38],
                       ["NC", 36], ["MI", 24]],
    },
}


def ym(dt):
    """datetime -> 'YYYY-MM' string."""
    return f"{dt.year:04d}-{dt.month:02d}"


def short_name(product_name):
    """Compact chart label from the full Chewy product name."""
    n = product_name
    brand = "FF " if n.startswith("Feline Fresh") else ""
    formula = ""
    for key in ["Healthy", "Multi-Cat", "Pine Pellet", "Unscented",
                "Sisal", "Wide Slatted", "Poop Bags"]:
        if key in n:
            formula = key
            break
    size = ""
    for tok in n.replace(",", " ").split():
        if tok.endswith("-lb") or tok.endswith("lb"):
            size = tok
            break
    if "Poop Bags" in n:
        return "Poop Bags"
    if "Sisal" in n:
        return "Sisal Mat"
    if "Wide Slatted" in n or "Scoop" in n:
        return "Litter Scoop"
    if formula == "Pine Pellet":
        formula = "Pine"
    label = f"{brand}{formula}"
    if size:
        label += f" {size}"
    return label.strip()


def load_flat(wb):
    """Return time series keyed by part number, plus product metadata."""
    ws = wb["FY25_FY26 YTD Data"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c).strip() if c is not None else "" for c in rows[0]]
    idx = {name: i for i, name in enumerate(hdr)}

    months = set()
    products = {}
    series = defaultdict(lambda: {"retail": {}, "units": {},
                                  "autoship": {}, "oos": {}})
    for r in rows[1:]:
        month = r[idx["MONTH"]]
        if month is None:
            continue
        m = ym(month)
        months.add(m)
        part = str(r[idx["PRODUCT_PART_NUMBER"]])
        name = r[idx["PRODUCT_NAME"]]
        brand = "Feline Fresh" if str(r[idx["BRAND"]]).startswith("Feline") \
            else "Catalyst"
        if part not in products:
            products[part] = {"name": name, "brand": brand,
                              "short": short_name(name), "ct": None}
        net = r[idx["Net Sales"]] or 0
        units = r[idx["Units Sold"]] or 0
        auto = r[idx["Autoship % Units Sold"]]
        oos = r[idx["PDP OOS%"]]
        series[part]["retail"][m] = round(float(net), 2)
        series[part]["units"][m] = int(units)
        if auto is not None:
            series[part]["autoship"][m] = round(float(auto) * 100, 1)
        if oos is not None:
            series[part]["oos"][m] = round(float(oos) * 100, 1)
    return sorted(months), products, series


def load_summary(wb, products, series, months):
    """Attach CT codes + Chewy's FY-YTD retail rollup (authoritative)."""
    ws = wb["Summary"]
    rows = list(ws.iter_rows(values_only=True))
    # Monthly column headers live in row index 2, starting at col 12.
    month_cols = {}
    for ci, cell in enumerate(rows[2]):
        if hasattr(cell, "year"):
            month_cols[ym(cell)] = ci

    fy_summary = {}
    used = set()
    for r in rows[3:]:
        ct = r[1]
        desc = r[3]
        if not ct or desc in (None, "") or str(desc).startswith("Total"):
            continue
        # Summary row's monthly retail vector.
        svec = {}
        for m, ci in month_cols.items():
            val = r[ci]
            if val is not None and not isinstance(val, str):
                svec[m] = round(float(val), 2)
        # Best flat part = minimal total abs diff over shared months (unused).
        best, best_score = None, None
        for part, s in series.items():
            if part in used:
                continue
            shared = [m for m in svec if m in s["retail"]]
            if not shared:
                continue
            score = sum(abs(svec[m] - s["retail"][m]) for m in shared) / len(shared)
            if best_score is None or score < best_score:
                best, best_score = part, score
        if best is None:
            continue
        used.add(best)
        products[best]["ct"] = ct
        fy_summary[best] = {
            "fy25": _num(r[4]), "fy26_ytd": _num(r[6]),
            "fy25_ytd": _num(r[7]), "chg_d": _num(r[9]),
            "chg_p": _num(r[10]) if not isinstance(r[10], str) else None,
        }
    return fy_summary


def _num(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def main():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    months, products, series = load_flat(wb)
    fy_summary = load_summary(wb, products, series, months)

    # Detect a partial trailing month (report pulled mid-month): its total
    # units fall far below the recent run-rate. Excluded from charts/KPIs.
    partial_month = None
    if len(months) >= 4:
        last = months[-1]
        lu = sum(series[p]["units"].get(last, 0) for p in products)
        prior = sorted(sum(series[p]["units"].get(m, 0) for p in products)
                       for m in months[-4:-1])
        med = prior[len(prior) // 2]
        if med and lu < 0.5 * med:
            partial_month = last

    # Compute wholesale $ series (price schedule * units) where priced.
    wholesale = {}
    for part, s in series.items():
        if not WHOLESALE_PRICES.get(part):
            continue
        wholesale[part] = {m: round(wholesale_price(part, m) * u, 2)
                           for m, u in s["units"].items()}
    have_wholesale = bool(wholesale)

    payload = {
        "months": months,
        "products": products,
        "series": series,
        "wholesale": wholesale,
        "have_wholesale": have_wholesale,
        "fy_summary": fy_summary,
        "snapshot": SNAPSHOT,
        "partial_month": partial_month,
        "generated": months[-1],
    }

    with open(TEMPLATE, encoding="utf-8") as f:
        tpl = f.read()
    html = tpl.replace("/*DATA_PLACEHOLDER*/",
                       json.dumps(payload, separators=(",", ":")))
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)

    cat = sum(1 for p in products.values() if p["brand"] == "Catalyst")
    ff = sum(1 for p in products.values() if p["brand"] == "Feline Fresh")
    print(f"[ok] {len(months)} months ({months[0]}..{months[-1]}), "
          f"{len(products)} SKUs ({cat} Catalyst / {ff} Feline Fresh)")
    print(f"[ok] CT codes mapped: "
          f"{sum(1 for p in products.values() if p['ct'])}/{len(products)}")
    print(f"[ok] wholesale prices set: "
          f"{sum(1 for v in WHOLESALE_PRICES.values() if v)}/{len(WHOLESALE_PRICES)}")
    if partial_month:
        print(f"[ok] partial trailing month excluded from charts: {partial_month}")
    print(f"[ok] wrote {OUTPUT}")


if __name__ == "__main__":
    main()
