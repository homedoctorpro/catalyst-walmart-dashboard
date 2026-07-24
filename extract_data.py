"""
extract_data.py — Catalyst Pet Walmart Sales Dashboard Generator
Reads 6 weekly Excel files, geocodes stores, computes metrics, writes dashboard.html
"""

import json
import os
import sys
import glob
import re
import math
from datetime import date, timedelta
import pandas as pd
import pgeocode

# Windows consoles default to cp1252, which can't encode the Unicode glyphs
# (→ ✓ …) this script prints — that raised UnicodeEncodeError mid-run and
# aborted before the dashboard was written. Force UTF-8 stdout/stderr so manual
# runs work regardless of console codepage or PYTHONIOENCODING.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

# ─── Configuration ────────────────────────────────────────────────────────────

SHEET_MAP = {
    "202601": {"instore": "202601 In Store Total Catalyst", "bystore": None,
               "ecomm_l52": "L52Wk Ecomm",                 "ecomm_lw": None},
    "202602": {"instore": "Catalyst LW Sales Total",        "bystore": "202602 Sales by Store",
               "ecomm_l52": None,                          "ecomm_lw": None},
    "202603": {"instore": "LW Instore Sales Total ",        "bystore": "Sales by Store",
               "ecomm_l52": None,                          "ecomm_lw": None},
    "202604": {"instore": "202604 InStore Sales",           "bystore": "202604 Sales by Store",
               "ecomm_l52": None,                          "ecomm_lw": None},
    "202605": {"instore": "Catalyst LW Sales",              "bystore": "Catalyst Sales by Store",
               "ecomm_l52": None,                          "ecomm_lw": "LW Ecomm Sales Total"},
    "202606": {"instore": "Catalyst LW Sales Total",        "bystore": "Catalyst Sales by Stores",
               "ecomm_l52": "LIGNETICS Ecomm L52WK",       "ecomm_lw": None},
    "202607": {"instore": "Catalyst LW Instore Sales",      "bystore": "Catalyst Sales by Store",
               "ecomm_l52": "Lignetics Ecomm L52WK",       "ecomm_lw": None},
    "202608": {"instore": "Catalyst LW Instore Sales",      "bystore": "Catalyst Sales by Store",
               "ecomm_l52": "LIGNETICS L52WK Ecomm",       "ecomm_lw": None},
    "202609": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": "Lignetics Total L52Wk Ecomm ", "ecomm_lw": None},
    "202610": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": "LIGNETICS L52Wk Ecomm ",       "ecomm_lw": "LW Ecomm Catalyst "},
    "202611": {"instore": "Catalyst LW Sales",              "bystore": "Sales by Store",
               "ecomm_l52": "Lignetics L52Wk Ecomm",        "ecomm_lw": "LW Catalyst Ecomm"},
    "202612": {"instore": "Catalyst LW Sales",              "bystore": "Sales by Store",
               "ecomm_l52": "LIGNETICS L52WK Ecomm",        "ecomm_lw": "Catalyst LW Ecomm"},
    "202613": {"instore": "Catalyst LW Sales ",             "bystore": "LW Sales by Store",
               "ecomm_l52": "Lignetics L52Wk Ecomm Total",  "ecomm_lw": "Catalyst LW Ecomm"},
    "202614": {"instore": "Catalyst LW Sales",              "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm"},
    "202615": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": "LIGNETICS L52Wk Ecomm ",       "ecomm_lw": "Catalyst- LW Ecomm"},
    "202616": {"instore": "Catalyst LW Sales",              "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm"},
    "202617": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm"},
    "202618": {"instore": "LW Catalyst Sales",              "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "LW Catalyst Ecomm"},
    "202619": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm "},
    "202620": {"instore": "Catalyst LW Sales",              "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm"},
    "202621": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm "},
    "202622": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm "},
    "202623": {"instore": "Catalyst LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "Catalyst LW Ecomm"},
    "202624": {"instore": "CATALYST LW Sales ",             "bystore": "Sales by Store",
               "ecomm_l52": None,                           "ecomm_lw": "CATALYST LW Ecomm"},
}

SKUS = [
    "CATALYST15ORIG",
    "CATALYST15UNSCEN",
    "CATALYST34LBORIGINAL",
    "CATALYSTPET34LBUNSCE",
]

WHOLESALE_PRICE = {
    "CATALYST15ORIG":        9.18,
    "CATALYST15UNSCEN":      9.18,
    "CATALYST34LBORIGINAL":  12.98,
    "CATALYSTPET34LBUNSCE":  12.98,
}

# ── Price rollback ───────────────────────────────────────────────────────────
# Walmart rolled BOTH 15 lb Catalyst SKUs (Original + Unscented) from $18.24 to
# $15.97 starting 2026-07-13 — mid fiscal week 202624. The weekly feed has no
# per-day price, so "units sold at the rollback price" is IMPUTED from each
# week's blended implied price (pos_dollars / pos_qty) via a two-price back-out:
#     units@rollback = qty * (P_pre - P_blended) / (P_pre - P_rollback)
# where P_pre is the SKU's trailing pre-rollback implied price. Fully-post
# weeks count all units; fully-pre weeks count zero. build_rollback() below.
ROLLBACK = {
    "date":      "2026-07-13",                          # first day at rollback price
    "pre_price": 18.24,                                 # headline pre-rollback shelf price
    "price":     15.97,                                 # rollback price (both 15 lb scents)
    "skus":      ["CATALYST15ORIG", "CATALYST15UNSCEN"],
}


# ── Rollback co-op tracker ───────────────────────────────────────────────────
# We pay a fixed $150k co-op fee that funds the $2.27/unit rollback discount for
# 150000/2.27 = 66,079 units. While that fund is being consumed our per-unit
# profit is $0.99 (we absorb the $2.27); once it's exhausted profit returns to
# $3.26. The rollback program ends 2026-10-31. build_coop() projects cumulative
# units, the co-op exhaustion date, and total profit two ways — flat at the
# latest weekly run rate, and +3%/week growth.
COOP = {
    "fee":           150000.0,
    "end_date":      "2026-10-31",    # rollback program hard stop
    "profit_coop":   0.99,            # $/unit while co-op fund is being consumed
    "profit_post":   3.26,            # $/unit after co-op fund is exhausted
    "growth_weekly": 0.03,            # +3%/week scenario
}

# Endcap program goes live 2026-08-01. "Stocked before" is fixed at the last
# weekly report BEFORE this date; cohorts are then tracked forward.
ENDCAP_LIVE_DATE = "2026-08-01"


def build_endcap_cohorts(all_store_weeks, traited, endcap, week_dates):
    """Track weekly U/S/W for three endcap cohorts (fixed at the pre-Aug-1 baseline):
      A = endcap + stocked before      B = no endcap + stocked before
      C = endcap + NOT stocked before
    "Stocked before" = traited AND present in the baseline week's feed. U/S/W =
    cohort units that week / fixed cohort store count (includes zero-sellers)."""
    if not endcap or not endcap.get("rows"):
        return None
    live  = date.fromisoformat(ENDCAP_LIVE_DATE)
    weeks = sorted(w for w in all_store_weeks if w in week_dates)
    if not weeks:
        return None
    pre      = [w for w in weeks if date.fromisoformat(week_dates[w]) < live]
    baseline = pre[-1] if pre else None

    tset = set()
    if traited and traited.get("by_sku"):
        for a in traited["by_sku"].values():
            tset.update(str(x) for x in a)
    endcap_set = {str(r["store_number"]) for r in endcap["rows"]}

    if baseline:
        feed_b = {str(s) for s in all_store_weeks.get(baseline, {})}
        stocked_before = tset & feed_b
    else:
        stocked_before = set()

    cohorts = {
        "A": endcap_set & stocked_before,     # endcap + stocked before
        "B": stocked_before - endcap_set,     # no endcap + stocked before
        "C": endcap_set - stocked_before,     # endcap + NOT stocked before
    }
    sizes = {k: len(v) for k, v in cohorts.items()}

    series = {}
    for w in weeks:
        wd = all_store_weeks.get(w, {})
        row = {}
        for k, sset in cohorts.items():
            units = sum((wd.get(sn) or {}).get("total_qty") or 0 for sn in sset)
            n = sizes[k]
            row[k + "_units"] = units
            row[k + "_usw"]   = round(units / n, 3) if n else None
        series[w] = row

    return {
        "live_date":     ENDCAP_LIVE_DATE,
        "baseline_week": baseline,
        "weeks":         weeks,
        "sizes":         sizes,
        "series":        series,
        "labels": {
            "A": "Endcap + stocked before",
            "B": "No endcap + stocked before",
            "C": "Endcap + not stocked before",
        },
    }


def build_coop(rollback, metrics, week_dates):
    """Project the $150k rollback co-op: exhaustion date + total profit, two ways."""
    if not rollback or not rollback.get("start_week") or not rollback.get("by_week"):
        return None
    fee      = COOP["fee"]
    discount = round(rollback["pre_price"] - rollback["price"], 2)   # $2.27
    if discount <= 0:
        return None
    units_covered = int(fee // discount)
    end   = date.fromisoformat(COOP["end_date"])
    skus  = rollback["skus"]
    p_coop, p_post = COOP["profit_coop"], COOP["profit_post"]

    as_of_wk = sorted(rollback["by_week"])[-1]
    as_of    = date.fromisoformat(week_dates[as_of_wk])
    # Run rate = latest week's full 15 lb volume (both scents) — go-forward every
    # unit sells at the rollback price.
    run = sum(((metrics.get(as_of_wk, {}).get(s) or {}).get("pos_qty") or 0) for s in skus)

    units_to_date = rollback["units_total"]

    fut = []
    d = as_of + timedelta(days=7)
    while d <= end:
        fut.append(d)
        d += timedelta(days=7)

    def project(growth):
        cum, profit, exhaust = units_to_date, units_to_date * p_coop, None
        for i, wend in enumerate(fut, start=1):
            if growth:
                wk = run * ((1 + COOP["growth_weekly"]) ** i)   # +3% compounded each week
            else:
                wk = float(run)
            start, end_c = cum, cum + wk
            coop_u = max(0.0, min(end_c, units_covered) - start)
            profit += coop_u * p_coop + (wk - coop_u) * p_post
            if exhaust is None and end_c >= units_covered and wk > 0:
                frac = (units_covered - start) / wk
                exhaust = (wend - timedelta(days=7)) + timedelta(days=round(frac * 7))
            cum = end_c
        total_disc = round(cum * discount, 2)          # total rollback discount $
        our_share  = round(min(fee, total_disc), 2)    # we pay $150k (fixed) first
        wm_share   = round(total_disc - our_share, 2)   # Walmart funds the remainder
        return {
            "units_total":  int(round(cum)),
            "units_coop":   int(min(round(cum), units_covered)),
            "units_post":   int(max(0, round(cum) - units_covered)),
            "profit_total": round(profit, 2),
            "exhaust_date": exhaust.isoformat() if exhaust else None,
            "total_discount": total_disc,
            "our_share":    our_share,
            "wm_share":     wm_share,
            "our_pct":      round(our_share / total_disc * 100, 1) if total_disc else None,
            "wm_pct":       round(wm_share / total_disc * 100, 1) if total_disc else None,
        }

    return {
        "fee": fee, "discount": discount, "units_covered": units_covered,
        "end_date": COOP["end_date"], "growth_weekly": COOP["growth_weekly"],
        "profit_coop": p_coop, "profit_post": p_post,
        "as_of_week": as_of_wk, "as_of_date": as_of.isoformat(),
        "run_rate": int(run),
        "units_to_date": units_to_date,
        "spent": round(units_to_date * discount, 2),
        "remaining": round(fee - units_to_date * discount, 2),
        "units_remaining": max(0, units_covered - units_to_date),
        "profit_to_date": round(units_to_date * p_coop, 2),
        "future_weeks": len(fut),
        "flat":   project(False),
        "growth": project(True),
    }


def build_rollback(metrics, week_dates):
    """Impute cumulative 15 lb units sold at the rollback price. See ROLLBACK.

    Returns a dict for the dashboard/email, or None if no rollback-era week
    has data yet.
    """
    rb_date = date.fromisoformat(ROLLBACK["date"])
    price   = ROLLBACK["price"]
    skus    = ROLLBACK["skus"]
    weeks   = sorted(w for w in week_dates)

    def implied(w, sku):
        m = metrics.get(w, {}).get(sku, {})
        d, q = m.get("pos_dollars"), m.get("pos_qty")
        return (d / q) if d and q else None

    # Per-SKU pre-rollback baseline = mean implied price over up to 4 fully-pre
    # weeks (week END before the rollback date). Captures each scent's own price.
    baseline = {}
    for sku in skus:
        pre = [implied(w, sku) for w in weeks
               if date.fromisoformat(week_dates[w]) < rb_date]
        pre = [p for p in pre if p]
        baseline[sku] = round(sum(pre[-4:]) / len(pre[-4:]), 4) if pre else ROLLBACK["pre_price"]

    by_week = {}
    units_by_sku = {sku: 0 for sku in skus}
    start_week = None
    partial = None
    for w in weeks:
        end   = date.fromisoformat(week_dates[w])
        start = end - timedelta(days=6)
        if end < rb_date:
            continue                                    # fully pre-rollback
        if start_week is None:
            start_week = w
        row = {}
        for sku in skus:
            q = metrics.get(w, {}).get(sku, {}).get("pos_qty")
            if not q:
                continue
            if start >= rb_date:
                u = q                                   # entire week at rollback price
            else:
                imp   = implied(w, sku)
                denom = baseline[sku] - price
                frac  = ((baseline[sku] - imp) / denom) if (imp is not None and denom) else 0.0
                frac  = min(1.0, max(0.0, frac))
                u = int(round(q * frac))
            row[sku] = u
            units_by_sku[sku] += u
        by_week[w] = row
        if start < rb_date <= end and partial is None:
            partial = {"week": w, "days_pre": (rb_date - start).days,
                       "days_roll": (end - rb_date).days + 1}

    if start_week is None:
        return None

    units_total = sum(units_by_sku.values())
    return {
        "date":           ROLLBACK["date"],
        "pre_price":      ROLLBACK["pre_price"],
        "price":          price,
        "skus":           skus,
        "start_week":     start_week,
        "baseline_price": baseline,
        "by_week":        by_week,
        "units_by_sku":   units_by_sku,
        "units_total":    units_total,
        "dollars_total":  round(units_total * price, 2),
        "partial":        partial,
    }

# ── Ecomm product matching ───────────────────────────────────────────────────
# Catalyst products are matched by tokens (brand + size + scent) so format
# wobble (extra commas, " Bag" suffix, spacing) doesn't break the join.
# Feline Fresh uses exact-match since the product taxonomy is varied.
FELINE_FRESH_MAP = {
    "feline fresh non-clumping natural pine pellet cat litter, unscented, 20 lb bag":
        ("FF NonClump 20lb", "Feline Fresh"),
    "feline fresh non-clumping natural pine pellet cat litter, unscented, 40 lb bag":
        ("FF NonClump 40lb", "Feline Fresh"),
    "feline fresh non-clumping natural pine pellet cat litter, unscented, 10 lb bag":
        ("FF NonClump 10lb", "Feline Fresh"),
    "feline fresh natural clumping softwood cat litter, unscented, 10 lb bag":
        ("FF Clump SW 10lb", "Feline Fresh"),
    "feline fresh natural pine cat litter, 7-lb":
        ("FF Pine 7lb", "Feline Fresh"),
    "feline fresh natural pine cat litter, 20 lb.":
        ("FF Pine 20lb", "Feline Fresh"),
    "feline fresh pine pellet cat litter, 20lb":
        ("FF Pellet 20lb", "Feline Fresh"),
}


def parse_ecomm_product(name_raw):
    """Walmart ecomm product name → (short_label, brand) or None."""
    n = name_raw.lower().strip().replace("non- clumping", "non-clumping")

    if "catalyst" in n:
        if re.search(r"\b34\s*lb\b", n):
            size = "34lb"
        elif re.search(r"\b15\s*lb\b", n):
            size = "15lb"
        else:
            return None
        if "unscent" in n:
            scent = "Unsc"
        elif "original" in n:
            scent = "Orig"
        else:
            return None
        return (f"Catalyst {size} {scent}", "Catalyst")

    return FELINE_FRESH_MAP.get(n)


# Walmart fiscal calendar anchor: Friday end of fiscal week 1, by year.
# Add a new entry when calendar year rolls over (FY27, FY28, ...).
FISCAL_YEAR_WEEK1_FRIDAY = {
    "2026": date(2026, 2, 6),
}


def compute_week_label(week_code):
    """'202614' → 'Week 14 (5/8/26)'. Falls back to the raw code if unknown year."""
    if len(week_code) != 6 or not week_code.isdigit():
        return week_code
    year = week_code[:4]
    wk = int(week_code[4:])
    anchor = FISCAL_YEAR_WEEK1_FRIDAY.get(year)
    if not anchor:
        return week_code
    friday = anchor + timedelta(days=7 * (wk - 1))
    return f"Week {wk} ({friday.month}/{friday.day}/{friday.year % 100})"


def compute_week_date(week_code):
    """'202614' → '2026-05-08' (Friday end-of-week ISO date). Returns None if year unknown."""
    if len(week_code) != 6 or not week_code.isdigit():
        return None
    year = week_code[:4]
    wk = int(week_code[4:])
    anchor = FISCAL_YEAR_WEEK1_FRIDAY.get(year)
    if not anchor:
        return None
    return (anchor + timedelta(days=7 * (wk - 1))).isoformat()

GEO_CACHE_FILE = "stores_geo.json"
TEMPLATE_FILE  = "dashboard_template.html"
OUTPUT_FILE    = "dashboard.html"
STORE_MAP_TEMPLATE_FILE = "store_map_template.html"
STORE_MAP_OUTPUT_FILE   = "store_map.html"

STATE_ABBR_TO_NAME = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho",
    "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah",
    "VT": "Vermont", "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming", "DC": "District of Columbia",
}

# ─── Helpers ──────────────────────────────────────────────────────────────────

def normalize_zip(z):
    """Return 5-digit zero-padded zip string."""
    return str(z).split("-")[0].split(".")[0].strip().zfill(5)


def find_excel_files():
    """Auto-detect all '2026XX Weekly Sales Report Catalyst' files (.xlsx or .xlsb).

    Hailey's source workbooks arrive as .xlsb; converted copies are .xlsx. If both
    exist for a week, the .xlsx wins (globbed last).
    """
    base = os.path.dirname(__file__)
    result = {}
    for ext in ("xlsb", "xlsx"):
        for f in glob.glob(os.path.join(base, f"2026?? Weekly Sales Report Catalyst.{ext}")):
            m = re.search(r"(2026\d{2})", os.path.basename(f))
            if m:
                result[m.group(1)] = f
    return result


def read_excel(filepath, **kwargs):
    """pd.read_excel that selects the pyxlsb engine for .xlsb workbooks."""
    if filepath.lower().endswith(".xlsb"):
        kwargs.setdefault("engine", "pyxlsb")
    return pd.read_excel(filepath, **kwargs)


def _sheet_names(filepath):
    """List a workbook's sheet names (.xlsx or .xlsb)."""
    engine = "pyxlsb" if filepath.lower().endswith(".xlsb") else None
    xl = pd.ExcelFile(filepath, engine=engine)
    try:
        return list(xl.sheet_names)
    finally:
        xl.close()


def detect_sheets(filepath):
    """Infer {instore, bystore, ecomm_l52, ecomm_lw} sheet names by keyword.

    Used ONLY for weeks absent from SHEET_MAP (new auto-ingested weeks); the 21
    historical weeks keep their exact SHEET_MAP entries, so nothing regresses.
    Returns names exactly as they appear in the workbook (trailing spaces intact).
    """
    names = _sheet_names(filepath)
    low = {n: n.lower().strip() for n in names}

    def pick(pred, prefer_catalyst=True):
        cands = [n for n in names if pred(low[n])]
        if not cands:
            return None
        if prefer_catalyst:
            cat = [n for n in cands if "catalyst" in low[n]]
            if cat:
                cands = cat
        return cands[0]

    is_ecomm = lambda s: "ecomm" in s
    is_l52   = lambda s: "l52" in s or "52wk" in s or "52 wk" in s

    def is_instore(s):
        if is_ecomm(s) or "by store" in s:
            return False
        if any(k in s for k in ("inventory", "forecast", "supply", "demand",
                                "modular", "order", "plan")):
            return False
        return "sales" in s or "instore" in s or "in store" in s

    return {
        "instore":   pick(is_instore),
        "bystore":   pick(lambda s: "by store" in s),
        "ecomm_l52": pick(lambda s: is_ecomm(s) and is_l52(s), prefer_catalyst=False),
        "ecomm_lw":  pick(lambda s: is_ecomm(s) and not is_l52(s)
                                    and ("lw" in s or "last week" in s)),
    }


# ─── Extraction ───────────────────────────────────────────────────────────────

def extract_instore_metrics(week, df):
    """
    Parse the InStore Sales Total sheet.
    Row 0: time period metadata (skip)
    Row 1: headers (skip — use column indices)
    Rows 2–5: SKUs; Row 6: Total
    Col 0: item name | Col 1: POS $ | Col 7: POS Qty | Col 9: Instock % | Col 33: U/S/W
    """
    result = {}

    # Map row index 2–5 to SKU labels by reading col 0, not assuming fixed order
    sku_rows = {}
    for row_idx in [2, 3, 4, 5]:
        sku_name = str(df.iloc[row_idx, 0]).strip()
        if sku_name in SKUS:
            sku_rows[row_idx] = sku_name

    def safe_float(val, multiply=1):
        try:
            v = float(val)
            if math.isnan(v):
                return None
            return round(v * multiply, 6)
        except (ValueError, TypeError):
            return None

    for row_idx, sku in sku_rows.items():
        row = df.iloc[row_idx]
        pos_dollars = safe_float(row.iloc[1])
        pos_qty_raw = safe_float(row.iloc[7])
        pos_qty = int(pos_qty_raw) if pos_qty_raw is not None else None
        instock_pct = safe_float(row.iloc[9], 100)   # stored as fraction → multiply × 100
        usw = safe_float(row.iloc[33])
        markdown_pct = safe_float(row.iloc[41], 100)  # col 41: Markdown % Sales TY, fraction → %

        wholesale = None
        if pos_qty is not None:
            wholesale = round(pos_qty * WHOLESALE_PRICE[sku], 2)

        result[sku] = {
            "pos_dollars":      pos_dollars,
            "pos_qty":          pos_qty,
            "instock_pct":      instock_pct,
            "usw":              usw,
            "wholesale_dollars": wholesale,
            "markdown_pct":     markdown_pct,
        }

    # Total row (row 6)
    row = df.iloc[6]
    pos_dollars_t = safe_float(row.iloc[1])
    pos_qty_raw_t = safe_float(row.iloc[7])
    pos_qty_t = int(pos_qty_raw_t) if pos_qty_raw_t is not None else None
    instock_pct_t = safe_float(row.iloc[9], 100)
    usw_t = safe_float(row.iloc[33])
    markdown_pct_t = safe_float(row.iloc[41], 100)

    # Wholesale Total = sum of 4 SKU values
    wholesale_t = sum(
        v["wholesale_dollars"] for v in result.values() if v["wholesale_dollars"] is not None
    )
    wholesale_t = round(wholesale_t, 2) if wholesale_t else None

    result["Total"] = {
        "pos_dollars":       pos_dollars_t,
        "pos_qty":           pos_qty_t,
        "instock_pct":       instock_pct_t,
        "usw":               usw_t,
        "wholesale_dollars": wholesale_t,
        "markdown_pct":      markdown_pct_t,
    }

    return result


def extract_store_data(week, df):
    """
    Parse Sales by Store sheet.
    Col 0: item_name | Col 1: store_number | Col 2: street | Col 3: state
    Col 4: city | Col 5: zip
    POS Qty and On Hand Qty columns are detected by header name (column layout
    varies by week — e.g. 202607 added a POS Sales $ column before POS Qty).
    """
    # Detect inventory column indices from header row
    header = [str(df.iloc[0, c]).replace("\n", " ") for c in range(df.shape[1])]
    pos_qty_col    = next((i for i, h in enumerate(header) if "POS Quantity" in h), 7)
    on_hand_col    = next((i for i, h in enumerate(header) if "On Hand Quantity" in h), 8)
    in_transit_col = next((i for i, h in enumerate(header) if "In Transit Quantity" in h), None)
    on_order_col   = next((i for i, h in enumerate(header) if "On Order Quantity" in h), None)
    pipeline_col   = next((i for i, h in enumerate(header) if "Total Pipeline Quantity" in h), None)

    def to_int(v):
        try:
            f = float(v)
            return int(f) if not math.isnan(f) else 0
        except (ValueError, TypeError):
            return 0

    rows = []
    for i in range(len(df)):
        row = df.iloc[i]
        try:
            store_num = str(row.iloc[1]).strip()
            # Skip header-ish rows
            if not store_num.isdigit():
                continue
            item_name = str(row.iloc[0]).strip().upper()
            street    = str(row.iloc[2]).strip()
            state_raw = str(row.iloc[3]).strip()
            city      = str(row.iloc[4]).strip()
            zip_raw   = row.iloc[5]
            zip5 = normalize_zip(zip_raw)

            rows.append({
                "item_name":    item_name,
                "store_num":    store_num,
                "street":       street,
                "state":        state_raw,
                "city":         city,
                "zip5":         zip5,
                "pos_qty":      to_int(row.iloc[pos_qty_col]),
                "on_hand":      to_int(row.iloc[on_hand_col]),
                "in_transit":   to_int(row.iloc[in_transit_col]) if in_transit_col is not None else 0,
                "on_order":     to_int(row.iloc[on_order_col])   if on_order_col   is not None else 0,
                "total_pipeline": to_int(row.iloc[pipeline_col]) if pipeline_col   is not None else 0,
            })
        except Exception:
            continue
    return rows


# ─── Ecomm Extraction ─────────────────────────────────────────────────────────

def _safe_num(val):
    """Return float or None; convert int-like floats."""
    try:
        v = float(val)
        return None if math.isnan(v) else v
    except (ValueError, TypeError):
        return None


def extract_ecomm_data(df):
    """
    Parse a Lignetics L52WK Ecomm (or LW Ecomm) sheet.
    Row 0: headers; Rows 1+: products (skip 'Total' rows and unknowns).
    Cols: 0=Product Name, 1=Net Retail Sales, 3=Net Unit Sales.
    Returns {short_name: {"brand": str, "r": float|None, "u": int|None}}.
    Logs a [WARN] for any row containing "catalyst" that fails to parse —
    surfaces silent format drift instead of dropping our SKUs.
    """
    result = {}
    unmapped_catalyst = []
    for i in range(1, len(df)):
        row = df.iloc[i]
        try:
            name_raw = str(row.iloc[0]).strip()
            if not name_raw or name_raw.lower() in ("total", "nan", "product name", ""):
                continue
            mapping = parse_ecomm_product(name_raw)
            if not mapping:
                if "catalyst" in name_raw.lower():
                    unmapped_catalyst.append(name_raw)
                continue
            short, brand = mapping
            retail = _safe_num(row.iloc[1])
            units  = _safe_num(row.iloc[3])
            result[short] = {
                "brand": brand,
                "r": round(retail, 2) if retail is not None else None,
                "u": int(units) if units is not None else None,
            }
        except Exception:
            continue
    for n in unmapped_catalyst:
        print(f"  [WARN] Unmapped Catalyst ecomm row: {n!r}")
    return result


def compute_ecomm_weekly(ecomm_l52_raw, ecomm_lw_raw):
    """
    Build weekly ecomm data, preferring LW sheets per-product over L52WK deltas:
      - L52WK sheets contribute week-over-week deltas vs the immediately preceding
        L52WK week.  The 'span' field records how many calendar weeks the
        delta covers (>1 when there's a gap in L52WK coverage).
      - LW sheets are direct 1-week observations and take precedence per-product —
        they overlay on top of L52WK deltas so any product present in both uses
        the LW value (span=1).  L52WK fills in any products LW didn't report
        (e.g. Feline Fresh when the LW sheet is Catalyst-only).

    Returns {week: {short_name: {"brand", "r", "u", "span"}}}
    """
    weekly = {}

    # L52WK deltas first (lower precedence)
    l52_weeks = sorted(ecomm_l52_raw.keys())
    for i in range(1, len(l52_weeks)):
        week     = l52_weeks[i]
        prev_wk  = l52_weeks[i - 1]
        span     = int(week) - int(prev_wk)   # e.g. 202606-202601 = 5

        cur  = ecomm_l52_raw[week]
        prev = ecomm_l52_raw[prev_wk]

        all_prods = set(cur) | set(prev)
        week_out = {}
        for short in all_prods:
            c = cur.get(short, {})
            p = prev.get(short, {})
            brand = c.get("brand") or p.get("brand")
            r_c, r_p = c.get("r"), p.get("r")
            u_c, u_p = c.get("u"), p.get("u")
            r_delta = round(r_c - r_p, 2) if (r_c is not None and r_p is not None) else r_c
            u_delta = (u_c - u_p) if (u_c is not None and u_p is not None) else u_c
            week_out[short] = {"brand": brand, "r": r_delta, "u": u_delta, "span": span}
        weekly[week] = week_out

    # LW observations overlay on top — LW always wins per-product (span=1)
    for week, prods in ecomm_lw_raw.items():
        existing = weekly.get(week, {})
        for short, d in prods.items():
            existing[short] = {**d, "span": 1}
        weekly[week] = existing

    return weekly


def build_supply_plan(files):
    """
    Parse the dedicated 'Supply Plan' sheet from each weekly report.

    The sheet name varies week to week ('Supply Plan', 'Supply Plan Export',
    'Supply Plan ', '202603 Supply Plan'…), so it's auto-detected by matching
    'supply plan' in the sheet name (the 'Forecast & Supply' sheet is a separate,
    differently-shaped sheet and is intentionally NOT used here).

    Each row is a scheduled inbound order: item desc (= our SKU code),
    scheduled arrival date, and Order Each Qty. Each report holds a rolling
    ~25-week (≈6-month) forward plan, so each weekly file is one snapshot.

    Returns {
      "snapshots": [report_week, ...],                       # weeks with a plan, sorted
      "by_sku":   {report_week: {sku: total_planned_units}}, # whole forward book per snapshot
      "arrival":  {report_week: {week_start_iso: {sku: units}}}, # forward schedule by arrival week
    }
    """
    from datetime import timedelta
    sku_set    = set(SKUS)
    snapshots  = []
    by_sku     = {}
    arrival    = {}

    for week in sorted(files.keys()):
        fp = files[week]
        try:
            xls = pd.ExcelFile(fp)
        except Exception as e:
            print(f"  [WARN] Supply Plan: cannot open {week}: {e}")
            continue
        cand = [s for s in xls.sheet_names if "supply plan" in s.lower()]
        if not cand:
            continue
        sheet = cand[0]
        try:
            df = xls.parse(sheet, header=0)
        except Exception as e:
            print(f"  [ERROR] Supply Plan sheet '{sheet}' ({week}): {e}")
            continue

        desc_c = next((c for c in df.columns if "Desc" in str(c)), None)
        qty_c  = next((c for c in df.columns if "Order Each" in str(c) or "Each Qty" in str(c)), None)
        arr_c  = next((c for c in df.columns if "Arrival" in str(c)), None)
        if not (desc_c and qty_c and arr_c):
            print(f"  [WARN] Supply Plan {week}: missing columns (cols={list(df.columns)})")
            continue

        df["_sku"] = df[desc_c].astype(str).str.strip().str.upper()
        df["_qty"] = pd.to_numeric(df[qty_c], errors="coerce").fillna(0)
        df["_arr"] = pd.to_datetime(df[arr_c], errors="coerce")
        sub = df[df["_sku"].isin(sku_set)]
        if sub.empty:
            print(f"  [WARN] Supply Plan {week}: no Catalyst rows matched (skus seen: "
                  f"{sorted(df['_sku'].unique())[:6]}…)")
            continue

        snapshots.append(week)
        by_sku[week] = {s: int(sub.loc[sub["_sku"] == s, "_qty"].sum()) for s in SKUS}

        wk_arr = {}
        for _, r in sub.iterrows():
            d = r["_arr"]
            if pd.isna(d):
                continue
            wstart = (d - timedelta(days=int(d.weekday()))).strftime("%Y-%m-%d")  # Monday of arrival wk
            bucket = wk_arr.setdefault(wstart, {})
            bucket[r["_sku"]] = bucket.get(r["_sku"], 0) + int(r["_qty"])
        arrival[week] = wk_arr

        print(f"  Supply Plan {week} [{sheet.strip()}]: {len(sub)} Catalyst order rows, "
              f"{sum(by_sku[week].values()):,} units over {len(wk_arr)} arrival weeks")

    return {"snapshots": snapshots, "by_sku": by_sku, "arrival": arrival}


# ─── Geocoding ────────────────────────────────────────────────────────────────

def load_geo_cache():
    if os.path.exists(GEO_CACHE_FILE):
        with open(GEO_CACHE_FILE, "r") as f:
            return json.load(f)
    return {}


def save_geo_cache(cache):
    with open(GEO_CACHE_FILE, "w") as f:
        json.dump(cache, f, indent=2)


def geocode_stores(store_meta, cache):
    """
    Bulk geocode unique zip codes using pgeocode.
    store_meta: {store_num: {zip5, state, city, street}}
    Returns updated cache and fills lat/lon into store_meta.
    """
    nomi = pgeocode.Nominatim("us")

    # Collect uncached zips
    all_zips = set(v["zip5"] for v in store_meta.values())
    new_zips = [z for z in all_zips if z not in cache]

    if new_zips:
        print(f"  Geocoding {len(new_zips)} new zip codes...")
        result = nomi.query_postal_code(new_zips)
        for i, zip5 in enumerate(new_zips):
            row = result.iloc[i]
            lat = row.get("latitude", None)
            lon = row.get("longitude", None)
            try:
                lat = round(float(lat), 4) if lat is not None and not math.isnan(float(lat)) else None
                lon = round(float(lon), 4) if lon is not None and not math.isnan(float(lon)) else None
            except (ValueError, TypeError):
                lat = lon = None
            cache[zip5] = {"lat": lat, "lon": lon}
        save_geo_cache(cache)
        print(f"  Geocoded and cached {len(new_zips)} zips.")
    else:
        print("  All zips already cached.")

    return cache


# ─── OOS exclusions for de-listed (non-traited) stores ────────────────────────

def mark_oos_exclusions(all_store_weeks, store_weeks, traited_by_sku):
    """For NON-traited (store, sku) pairs, mark the *terminal* OOS tail — the
    run of on_hand==0 weeks at the very end of their history (after the last
    week they held inventory) — with excl_oos=True on the slot dict.

    These trailing stockouts are the store dropping the listing, not a real
    out-of-stock, so OOS computations skip flagged slots (numerator AND
    denominator). Earlier weeks (when the store was active — in-stock, or a
    temporary mid-history stockout) are left untouched.

    traited_by_sku: {sku: set(store_num)} of traited & valid stores.
    Returns the number of slots flagged.
    """
    # Per (store, sku): ordered list of slot dicts across weeks
    timeline = {}  # (sku, store_num) -> [(week, slot_dict), ...] in week order
    for week in store_weeks:
        for store_num, sdata in all_store_weeks.get(week, {}).items():
            for sku, skudata in sdata.get("skus", {}).items():
                timeline.setdefault((sku, store_num), []).append(skudata)

    n = 0
    for (sku, store_num), seq in timeline.items():
        if store_num in traited_by_sku.get(sku, set()):
            continue  # traited → real distribution, keep all OOS
        for skudata in reversed(seq):       # walk from latest week backward
            if skudata.get("on_hand") == 0:
                skudata["excl_oos"] = True
                n += 1
            else:
                break                        # first in-stock week ends the tail
    return n


# ─── Consecutive OOS computation ──────────────────────────────────────────────

def compute_consecutive_oos_by_week(all_store_weeks, store_weeks):
    """
    For each week W (in store_weeks order), for each (sku, store) pair,
    count trailing consecutive OOS weeks ending at W.
    OOS = on_hand == 0. Slots flagged excl_oos (de-listed non-traited tails)
    are skipped — they neither count as OOS nor get recorded.

    Returns: {week: {sku: {store_num: consecutive_count}}}
    """
    result = {w: {sku: {} for sku in SKUS} for w in store_weeks}

    # Build timeline: for each (sku, store) — {week: (on_hand, excluded)}
    timeline = {}  # (sku, store_num) -> {week: (on_hand, excl)}
    for week in store_weeks:
        week_data = all_store_weeks.get(week, {})
        for store_num, sdata in week_data.items():
            for sku, skudata in sdata.get("skus", {}).items():
                key = (sku, store_num)
                if key not in timeline:
                    timeline[key] = {}
                timeline[key][week] = (skudata.get("on_hand", None), skudata.get("excl_oos", False))

    for (sku, store_num), week_map in timeline.items():
        for w_idx, week in enumerate(store_weeks):
            if week not in week_map:
                continue
            if week_map[week][1]:        # excluded slot — don't record OOS here
                continue
            # Count trailing OOS ending at this week
            count = 0
            for past_week in reversed(store_weeks[:w_idx + 1]):
                entry = week_map.get(past_week)
                if entry is None:
                    break   # store didn't report that week — stop chain
                oh, excl = entry
                if oh is None or excl:
                    break   # de-listed tail or missing — chain ends
                if oh == 0:
                    count += 1
                else:
                    break
            if count > 0:
                result[week][sku][store_num] = count

    return result


# ─── Aggregation ──────────────────────────────────────────────────────────────

def compute_state_sales(all_store_weeks, stores):
    """Returns {week: {state_name: total_pos_qty}}"""
    out = {}
    for week, week_data in all_store_weeks.items():
        state_totals = {}
        for store_num, sdata in week_data.items():
            state_abbr = stores.get(store_num, {}).get("state", "")
            state_name = STATE_ABBR_TO_NAME.get(state_abbr, state_abbr)
            total_qty = sdata.get("total_qty", 0)
            state_totals[state_name] = state_totals.get(state_name, 0) + total_qty
        out[week] = state_totals
    return out


def compute_state_oos(all_store_weeks, stores):
    """Returns {week: {state_name: oos_fraction (0.0–1.0)}}"""
    out = {}
    for week, week_data in all_store_weeks.items():
        state_counts = {}   # state -> [total_sku_slots, oos_slots]
        for store_num, sdata in week_data.items():
            state_abbr = stores.get(store_num, {}).get("state", "")
            state_name = STATE_ABBR_TO_NAME.get(state_abbr, state_abbr)
            if state_name not in state_counts:
                state_counts[state_name] = [0, 0]
            for sku, skudata in sdata.get("skus", {}).items():
                if skudata.get("excl_oos"):
                    continue   # de-listed non-traited tail — out of OOS scope
                state_counts[state_name][0] += 1
                if skudata.get("on_hand", 1) == 0:
                    state_counts[state_name][1] += 1
        out[week] = {
            st: round(counts[1] / counts[0], 4) if counts[0] > 0 else 0.0
            for st, counts in state_counts.items()
        }
    return out


def build_weekly_store_summary(raw_store_rows_by_week):
    """
    Collapse per-SKU rows into per-store dicts.
    Returns {week: {store_num: {total_qty, skus: {sku: {qty, on_hand}}}}}
    """
    out = {}
    for week, rows in raw_store_rows_by_week.items():
        week_stores = {}
        for r in rows:
            sn = r["store_num"]
            sku = r["item_name"]
            if sn not in week_stores:
                week_stores[sn] = {"total_qty": 0, "skus": {}}
            week_stores[sn]["skus"][sku] = {
                "qty":          r["pos_qty"],
                "on_hand":      r["on_hand"],
                "in_transit":   r["in_transit"],
                "on_order":     r["on_order"],
                "total_pipeline": r["total_pipeline"],
            }
            week_stores[sn]["total_qty"] += r["pos_qty"]
        out[week] = week_stores
    return out


def build_stores_dict(raw_store_rows_by_week, geo_cache):
    """
    Build master stores dict: {store_num: {zip5, state, city, street, lat, lon}}
    Use state abbreviation from data (first occurrence wins).
    """
    stores = {}
    for week, rows in raw_store_rows_by_week.items():
        for r in rows:
            sn = r["store_num"]
            if sn not in stores:
                zip5 = r["zip5"]
                geo = geo_cache.get(zip5, {})
                stores[sn] = {
                    "zip5":   zip5,
                    "state":  r["state"],
                    "city":   r["city"],
                    "street": r["street"],
                    "lat":    geo.get("lat"),
                    "lon":    geo.get("lon"),
                }
    return stores


def compute_weekly_inventory(all_store_weeks):
    """
    Aggregate on_hand, in_transit, on_order, total_pipeline by week and SKU.
    Returns {week: {sku: {on_hand, in_transit, on_order, total_pipeline}, "Total": {...}}}
    """
    out = {}
    inv_keys = ("on_hand", "in_transit", "on_order", "total_pipeline")
    for week, week_data in all_store_weeks.items():
        sku_totals = {}  # sku -> {key: sum}
        for store_num, sdata in week_data.items():
            for sku, skudata in sdata.get("skus", {}).items():
                if sku not in sku_totals:
                    sku_totals[sku] = {k: 0 for k in inv_keys}
                for k in inv_keys:
                    sku_totals[sku][k] += skudata.get(k, 0)
        # Compute Total across all SKUs
        total = {k: sum(sku_totals[s][k] for s in sku_totals) for k in inv_keys}
        out[week] = {**sku_totals, "Total": total}
    return out


def compute_state_inventory(all_store_weeks, stores):
    """
    Aggregate on_hand, in_transit, on_order, total_pipeline by week and state.
    Returns {week: {state_name: {on_hand, in_transit, on_order, total_pipeline}}}
    """
    out = {}
    inv_keys = ("on_hand", "in_transit", "on_order", "total_pipeline")
    for week, week_data in all_store_weeks.items():
        state_totals = {}
        for store_num, sdata in week_data.items():
            state_abbr = stores.get(store_num, {}).get("state", "")
            state_name = STATE_ABBR_TO_NAME.get(state_abbr, state_abbr)
            if state_name not in state_totals:
                state_totals[state_name] = {k: 0 for k in inv_keys}
            for sku, skudata in sdata.get("skus", {}).items():
                for k in inv_keys:
                    state_totals[state_name][k] += skudata.get(k, 0)
        out[week] = state_totals
    return out


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    # 1. Find Excel files
    files = find_excel_files()
    print(f"Found {len(files)} Excel files: {sorted(files.keys())}")

    # 2. Load geo cache
    geo_cache = load_geo_cache()
    print(f"Loaded geo cache with {len(geo_cache)} entries.")

    # 3. Read data
    metrics = {}
    raw_store_rows_by_week = {}
    store_weeks_list = []
    ecomm_l52_raw = {}   # {week: {short: {brand,r,u}}}
    ecomm_lw_raw  = {}   # {week: {short: {brand,r,u}}}

    for week in sorted(files.keys()):
        filepath = files[week]
        sheet_info = SHEET_MAP.get(week)
        if not sheet_info:
            try:
                sheet_info = detect_sheets(filepath)
                print(f"  [auto] Detected sheets for {week}: {sheet_info}")
            except Exception as e:
                print(f"  [WARN] Sheet auto-detection failed for {week}: {e}; skipping.")
                continue

        print(f"\nProcessing {week}...")

        # InStore metrics
        instore_sheet = sheet_info["instore"]
        try:
            df_in = read_excel(filepath, sheet_name=instore_sheet, header=None)
            metrics[week] = extract_instore_metrics(week, df_in)
            print(f"  InStore metrics: OK ({len(metrics[week])} entries)")
        except Exception as e:
            print(f"  [ERROR] InStore sheet '{instore_sheet}': {e}")

        # Sales by Store
        bystore_sheet = sheet_info["bystore"]
        if bystore_sheet:
            try:
                df_bs = read_excel(filepath, sheet_name=bystore_sheet, header=None)
                rows = extract_store_data(week, df_bs)
                raw_store_rows_by_week[week] = rows
                store_weeks_list.append(week)
                print(f"  Sales by Store: {len(rows)} rows")
            except Exception as e:
                print(f"  [ERROR] ByStore sheet '{bystore_sheet}': {e}")

        # L52WK Ecomm
        l52_sheet = sheet_info.get("ecomm_l52")
        if l52_sheet:
            try:
                df_e = read_excel(filepath, sheet_name=l52_sheet, header=None)
                prods = extract_ecomm_data(df_e)
                ecomm_l52_raw[week] = prods
                print(f"  L52WK Ecomm: {len(prods)} products")
            except Exception as e:
                print(f"  [ERROR] L52WK Ecomm sheet '{l52_sheet}': {e}")

        # LW Ecomm (direct weekly)
        lw_sheet = sheet_info.get("ecomm_lw")
        if lw_sheet:
            try:
                df_e = read_excel(filepath, sheet_name=lw_sheet, header=None)
                prods = extract_ecomm_data(df_e)
                ecomm_lw_raw[week] = prods
                print(f"  LW Ecomm: {len(prods)} products")
            except Exception as e:
                print(f"  [ERROR] LW Ecomm sheet '{lw_sheet}': {e}")

    # 4. Geocode
    print("\nGeocoding store zip codes...")
    # Collect all store meta from raw rows
    temp_store_meta = {}
    for week, rows in raw_store_rows_by_week.items():
        for r in rows:
            sn = r["store_num"]
            if sn not in temp_store_meta:
                temp_store_meta[sn] = {"zip5": r["zip5"]}

    geo_cache = geocode_stores(temp_store_meta, geo_cache)

    # 5. Build structures
    print("\nBuilding data structures...")
    all_store_weeks = build_weekly_store_summary(raw_store_rows_by_week)
    stores = build_stores_dict(raw_store_rows_by_week, geo_cache)

    # Traited/valid snapshot first — needed to strip de-listed OOS tails before
    # the OOS computations run.
    print("Building Traited/Valid authorization data...")
    try:
        from traited_status import build_traited_data
        traited = build_traited_data(verbose=True)
    except Exception as e:
        print(f"  [WARN] Traited/valid data unavailable: {e}")
        traited = None
    if traited and traited.get("by_sku"):
        traited_by_sku = {sku: set(lst) for sku, lst in traited["by_sku"].items()}
        n_excl = mark_oos_exclusions(all_store_weeks, store_weeks_list, traited_by_sku)
        print(f"  [OOS] Excluded {n_excl:,} de-listed non-traited OOS slots (terminal stockout tails)")

    state_sales = compute_state_sales(all_store_weeks, stores)
    state_oos   = compute_state_oos(all_store_weeks, stores)
    consecutive_oos = compute_consecutive_oos_by_week(all_store_weeks, store_weeks_list)
    weekly_inventory = compute_weekly_inventory(all_store_weeks)
    state_inventory  = compute_state_inventory(all_store_weeks, stores)

    print(f"  Stores: {len(stores)}")
    print(f"  Store-weeks: {sorted(all_store_weeks.keys())}")

    # 5b. Ecomm weekly deltas
    print("\nComputing ecomm weekly data...")
    ecomm_weekly = compute_ecomm_weekly(ecomm_l52_raw, ecomm_lw_raw)
    print(f"  L52WK weeks: {sorted(ecomm_l52_raw.keys())}")
    print(f"  Weekly ecomm weeks: {sorted(ecomm_weekly.keys())}")

    # 6. Assemble JSON
    all_week_codes = sorted(set(files.keys()) | set(store_weeks_list) | set(ecomm_weekly.keys()))
    week_labels = {w: compute_week_label(w) for w in all_week_codes}
    week_dates  = {w: d for w in all_week_codes if (d := compute_week_date(w))}

    # Price rollback: impute cumulative 15 lb units sold at the $15.97 rollback price.
    rollback = build_rollback(metrics, week_dates)
    if rollback:
        p = rollback.get("partial")
        pnote = (f", partial start week {p['week']}: {p['days_roll']}/7 days"
                 if p else "")
        print(f"\nRollback: 15 lb ${ROLLBACK['pre_price']:.2f}→${ROLLBACK['price']:.2f} "
              f"from {ROLLBACK['date']} (start {rollback['start_week']}{pnote}); "
              f"imputed {rollback['units_total']:,} units @ rollback "
              f"(≈${rollback['dollars_total']:,.0f})")

    # Co-op tracker: project the $150k fund's exhaustion + total profit (flat & +3%/mo).
    coop = build_coop(rollback, metrics, week_dates)
    if coop:
        print(f"  Co-op $150k: {coop['units_covered']:,} units covered; "
              f"{coop['units_to_date']:,} used (${coop['spent']:,.0f}). "
              f"Exhausts flat {coop['flat']['exhaust_date']} / +3% {coop['growth']['exhaust_date']}; "
              f"total profit by {coop['end_date']}: flat ${coop['flat']['profit_total']:,.0f} / "
              f"+3% ${coop['growth']['profit_total']:,.0f}")

    # 5c. Endcap data (optional - requires EndcapStoreList.xlsx + Store & DC Addresses.xlsx)
    print("\nBuilding endcap store data...")
    try:
        from endcap_export import build_endcap_rows
        endcap_rows, endcap_summary = build_endcap_rows(verbose=False)
        print(f"  Endcap stores: {endcap_summary['total']}, mapped: {endcap_summary['mapped']}, "
              f"no Catalyst: {endcap_summary['no_catalyst']}")
    except Exception as e:
        print(f"  [WARN] Endcap data unavailable: {e}")
        endcap_rows, endcap_summary = [], {"total": 0, "mapped": 0, "addressed": 0, "no_catalyst": 0}

    # 5d. Trial & Repeat data (optional - requires Catalyst Trial and Repeat Report*.xlsx)
    print("\nBuilding Trial & Repeat data...")
    try:
        from trial_repeat import build_trial_repeat_data
        trial_repeat = build_trial_repeat_data(verbose=True)
    except Exception as e:
        print(f"  [WARN] Trial & Repeat data unavailable: {e}")
        trial_repeat = None

    # 5e. Supply Plan snapshots (dedicated 'Supply Plan' sheet, auto-detected)
    print("\nBuilding Supply Plan data...")
    supply_plan = build_supply_plan(files)
    print(f"  Supply Plan snapshots: {supply_plan['snapshots']}")

    # 5f. Frozen forecast baseline (plan-of-record) — load if present, never regenerate
    #     here. Re-baseline only by running gen_forecast_baseline.js intentionally.
    forecast_baseline = None
    fb_path = os.path.join(os.path.dirname(__file__), "forecast_baseline.json")
    if os.path.exists(fb_path):
        with open(fb_path, "r", encoding="utf-8") as f:
            forecast_baseline = json.load(f)
        print(f"\nForecast baseline: frozen {forecast_baseline['meta'].get('frozen_as_of')} "
              f"(from week {forecast_baseline['meta'].get('generated_from_week')})")

    # Endcap cohort U/S/W tracker (fixed at pre-Aug-1 baseline, tracked forward).
    endcap_data = {"rows": endcap_rows, "summary": endcap_summary}
    endcap_cohorts = build_endcap_cohorts(all_store_weeks, traited, endcap_data, week_dates)
    if endcap_cohorts:
        s = endcap_cohorts["sizes"]
        print(f"\nEndcap cohorts (baseline {endcap_cohorts['baseline_week']}, live {ENDCAP_LIVE_DATE}): "
              f"A endcap+stocked={s['A']:,}, B no-endcap+stocked={s['B']:,}, C endcap+new={s['C']:,}")

    data = {
        "weeks":       sorted(files.keys()),
        "store_weeks": store_weeks_list,
        "skus":        SKUS,
        "metrics":     metrics,
        "stores":      stores,
        "weekly_stores": all_store_weeks,
        "state_sales": state_sales,
        "state_oos":   state_oos,
        "consecutive_oos_by_week": consecutive_oos,
        "ecomm_l52":   ecomm_l52_raw,
        "ecomm_weekly": ecomm_weekly,
        "weekly_inventory": weekly_inventory,
        "state_inventory":  state_inventory,
        "week_labels":  week_labels,
        "week_dates":   week_dates,
        "endcap":      endcap_data,
        "endcap_cohorts": endcap_cohorts,
        "trial_repeat": trial_repeat,
        "traited": traited,
        "supply_plan": supply_plan,
        "forecast_baseline": forecast_baseline,
        "rollback": rollback,
        "coop": coop,
    }

    # 7. Read template and embed JSON
    template_path = os.path.join(os.path.dirname(__file__), TEMPLATE_FILE)
    if not os.path.exists(template_path):
        print(f"\n[ERROR] Template file '{TEMPLATE_FILE}' not found. Please create it first.")
        return

    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    json_str = json.dumps(data, separators=(",", ":"))
    html = html.replace("/*DATA_PLACEHOLDER*/", f"const DATA = {json_str};")

    output_path = os.path.join(os.path.dirname(__file__), OUTPUT_FILE)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    json_mb = len(json_str) / 1024 / 1024
    print(f"\nWrote {output_path} ({json_mb:.2f} MB JSON embedded)")

    # 7b. Standalone public store map (no password gate) — shareable link
    write_store_map(stores, all_store_weeks, store_weeks_list, week_labels)
    print("Done.")

    # ── Send weekly report email ──────────────────────────────────────────────
    # Mondays (first run): full recipient list
    # All other runs: pross@lignetics.com only
    today = date.today()
    sent_flag = os.path.join(os.path.dirname(__file__), ".email_sent_date")
    already_sent = os.path.exists(sent_flag) and open(sent_flag).read().strip() == str(today)
    is_monday = today.weekday() == 0
    # Cloud auto-ingest runs set FORCE_FULL_DISTRO=1 so any day's new data goes
    # to the full recipient list (bypassing the Monday-only / already-sent gate).
    force_full = os.environ.get("FORCE_FULL_DISTRO") == "1"
    dev_only  = False if force_full else (not is_monday or already_sent)
    # DEV_ONLY_EMAIL=1 forces a dev-only send (test/example emails), overriding
    # the full-distro flag.
    if os.environ.get("DEV_ONLY_EMAIL") == "1":
        dev_only = True
    if os.environ.get("SKIP_EMAIL") == "1":
        print("  [Email] Skipped (SKIP_EMAIL=1)")
    else:
        try:
            from email_report import send_report
            send_report(data, dev_only=dev_only)
            if is_monday and not already_sent:
                open(sent_flag, "w").write(str(today))
        except Exception as e:
            print(f"  [Email] Error: {e}")


def write_store_map(stores, all_store_weeks, store_weeks_list, week_labels):
    """Write store_map.html — standalone, ungated store map for public sharing.

    Embeds only what the map needs: per-store location + the set of SKUs the
    store has carried across all weeks (mirrors the dashboard's Store Map tab).
    """
    template_path = os.path.join(os.path.dirname(__file__), STORE_MAP_TEMPLATE_FILE)
    if not os.path.exists(template_path):
        print(f"  [WARN] '{STORE_MAP_TEMPLATE_FILE}' not found — skipping standalone store map.")
        return

    # SKUs carried per store, across all weeks (same logic as dashboard initStoreMap)
    store_sku_sets = {}
    for week in store_weeks_list:
        for sn, sdata in all_store_weeks.get(week, {}).items():
            store_sku_sets.setdefault(sn, set()).update((sdata.get("skus") or {}).keys())

    map_stores = {}
    for sn, info in stores.items():
        skus = sorted(store_sku_sets.get(sn, ()))
        if not skus or not info.get("lat") or not info.get("lon"):
            continue
        map_stores[sn] = {
            "lat": info["lat"], "lon": info["lon"],
            "city": info.get("city", ""), "state": info.get("state", ""),
            "street": info.get("street", ""), "skus": skus,
        }

    latest_week = store_weeks_list[-1] if store_weeks_list else None
    as_of = week_labels.get(latest_week, latest_week) if latest_week else None

    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()
    payload = json.dumps({"stores": map_stores, "as_of": as_of}, separators=(",", ":"))
    html = html.replace("/*DATA_PLACEHOLDER*/", f"const MAP_DATA = {payload};")

    output_path = os.path.join(os.path.dirname(__file__), STORE_MAP_OUTPUT_FILE)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {output_path} ({len(map_stores):,} stores)")


if __name__ == "__main__":
    main()
