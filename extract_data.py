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
from collections import Counter
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
# $15.97 starting 2026-07-13 — mid fiscal week 202624.
#
# The feed has no per-day price, but it DOES have per-store POS $ and POS qty,
# so each store-week's AUR (POS $ / qty) is its actual realized shelf price.
# Those AURs are sharply discrete and persistent week to week — in 202628, 5
# price points cover 95% of units and 92-96% of stores hold the same point as
# the prior week — so AUR is a real price, not a statistical blur:
#     $15.97  ~80% of units   the rollback price
#     $16.97  ~7%             higher-zone rollback price (still discounted)
#     $19.97  ~7%             ABOVE the old $18.24 shelf price (never rolled
#                             back, or already reverted)
# The one genuine blend is the transition week 202624 (5 of 7 days at the new
# price), where intermediate AURs are real pre/post mixes.
#
# build_rollback() therefore reports TWO distinct measures instead of one, so
# neither gets read as the other:
#   units_discounted  — units at stores whose week AUR is strictly below the
#                       $18.24 pre-rollback price. This is the literal "units
#                       sold below the old price" count.
#   units_equiv       — markdown dollars / $2.27. Every store's own discount
#                       (pre - AUR, capped at the $2.27 program discount) is
#                       what actually draws against the co-op fund, so a store
#                       at $16.97 contributes 1.27/2.27 = 0.56 of a unit. This
#                       is the number the $150k co-op burn runs on, and it is
#                       also exact for the blended transition week.
# Stores at or above $18.24 contribute zero to both, so the count self-corrects
# as stores revert to $19.97. Weeks with no per-store POS $ column fall back to
# the same back-out on the week's blended implied price.
ROLLBACK = {
    "date":      "2026-07-13",                          # first day at rollback price
    "pre_price": 18.24,                                 # headline pre-rollback shelf price
    "price":     15.97,                                 # rollback price (both 15 lb scents)
    "skus":      ["CATALYST15ORIG", "CATALYST15UNSCEN"],
    "band_eps":  0.005,   # cents tolerance when testing AUR against a price point
    "top_bands": 6,       # distinct price points listed per week (rest -> "other")
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

# Endcap rollout status (build_endcap_status). The 36-bag endcap allocation
# started shipping in week 202626, so "has the allocation landed?" tests on-hand
# plus every week of sell-through from that week forward. The merchandiser field
# survey is the authoritative set / not-set record; it is a point-in-time visit,
# so the set counts stay pinned to the survey week while sales roll forward.
ENDCAP_ARRIVAL_WEEK = "202626"
ENDCAP_UNITS        = 36
# Weeks of pre-set context on the endcap U/S/W charts. The program has ~25 weeks
# of history in front of it that says nothing about the endcap; plotting all of it
# compresses the part that matters into the last inch of the axis. Two pre-set
# weeks is enough to show the lines were flat before they diverged.
ENDCAP_CHART_PRE_WEEKS = 2
ENDCAP_SURVEY_FILE  = "(Walmart) Lignetics Inc. Cat Litter Endcap Set WK27.xlsx"
ENDCAP_SURVEY_WEEK  = "202627"
# The set list grows in waves: the original WK27 sweep plus every follow-up
# re-visit Anderson sends ("* Endcap Update.xlsx", overlaid in load_survey).
# A follow-up store had no display up in the weeks before its own visit, so it
# is broken out of the set cohort everywhere sales are compared.
ENDCAP_FOLLOWUP_WEEK = "202628"
ENDCAP_SKU          = "CATALYST15ORIG"   # the SKU the endcap feature is built on
ENDCAP_REASONS = {  # survey answer -> (key, label)
    "No Available space":                    ("space",     "No available space"),
    "Not enough inventory to build feature": ("inventory", "Not enough inventory"),
    "Product not located":                   ("located",   "Product not located"),
    "Store Refusal":                         ("refusal",   "Store refusal"),
}


def _load_endcap_survey():
    """store_num (str) -> survey record from the merchandiser field visit.

    Reuses build_endcap_report.load_survey so the set/not-set definition and the
    reason taxonomy can't drift between the standalone report and the dashboard.
    Returns {} when the survey workbook isn't present.
    """
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ENDCAP_SURVEY_FILE)
    if not os.path.exists(path):
        return {}
    try:
        from build_endcap_report import load_survey
        return {str(s): v for s, v in load_survey().items()}
    except Exception as e:
        print(f"  [WARN] Endcap survey unreadable ({e}) — rollout status skipped.")
        return {}


def build_endcap_status(all_store_weeks, endcap, week_dates):
    """Endcap rollout scorecard: set rate, 36-bag arrival, and sales lift.

    Answers, refreshed every week off the live feed:
      - how many of the endcap stores the merchandisers confirmed set, and the
        filed reason for each store that isn't
      - how many have the 36-bag allocation at store, inbound, or short
      - 15 lb Original lift for the set stores this week vs last week, and vs
        the last week before the program went live, each against a control of
        every non-endcap Catalyst store (so market-wide moves net out)

    Returns None when the endcap roster or survey isn't available.
    """
    if not endcap or not endcap.get("rows"):
        return None
    survey = _load_endcap_survey()
    if not survey:
        return None

    weeks = sorted(w for w in all_store_weeks if w in week_dates)
    if len(weeks) < 2:
        return None
    live      = date.fromisoformat(ENDCAP_LIVE_DATE)
    this_week = weeks[-1]
    last_week = weeks[-2]
    pre       = [w for w in weeks if date.fromisoformat(week_dates[w]) < live]
    base_week = pre[-1] if pre else None
    arrival   = [w for w in weeks if w >= ENDCAP_ARRIVAL_WEEK]

    # per-store 15 lb Original units by week, plus latest on-hand / pipeline
    def q15(w):
        return {sn: ((d.get("skus") or {}).get(ENDCAP_SKU) or {}).get("qty") or 0
                for sn, d in (all_store_weeks.get(w) or {}).items()}
    qty    = {w: q15(w) for w in
              set(weeks[-2:] + ([base_week] if base_week else []) + arrival)}
    latest = all_store_weeks.get(this_week) or {}

    def sku_at(sn, key):
        return ((latest.get(sn, {}).get("skus") or {}).get(ENDCAP_SKU) or {}).get(key) or 0

    roster  = {str(r["store_number"]): r for r in endcap["rows"]}
    seg_of  = {sn: (survey.get(sn) or {}).get("seg", "unvisited") for sn in roster}
    set_ok  = {sn for sn in roster if seg_of[sn] == "set"}
    wave_of = {sn: (survey.get(sn) or {}).get("wave", "") for sn in roster}
    set_w27 = {sn for sn in set_ok if wave_of[sn] != "followup"}
    set_fu  = {sn for sn in set_ok if wave_of[sn] == "followup"}
    unvis   = {sn for sn in roster if seg_of[sn] == "unvisited"}
    notset  = set(roster) - set_ok - unvis
    visited = set_ok | notset

    # 36-bag arrival: on-hand now + everything sold since the allocation shipped.
    # A store that already sold through its 36 still counts as received. There is
    # no DC-warehouse column in this feed, so stock not yet on order reads "short".
    def inv_status(sn):
        if sn not in latest:
            return "short"
        have = sku_at(sn, "on_hand") + sum(qty[w].get(sn, 0) for w in arrival)
        transit, onorder = sku_at(sn, "in_transit"), sku_at(sn, "on_order")
        if have >= ENDCAP_UNITS:
            return "received"
        if have + transit + onorder >= ENDCAP_UNITS:
            return "transit" if transit > 0 else "onorder"
        return "short"

    inv_of   = {sn: inv_status(sn) for sn in roster}
    INV_KEYS = ["received", "transit", "onorder", "short"]

    # Lift: units this week vs each comparison week, per segment plus control.
    # The control is non-endcap stores that CARRIED the endcap SKU in the baseline
    # week — not merely present in the feed. Including stores that stock other
    # Catalyst sizes but not this one would pad the denominator with structural
    # zeros and understate the control's U/S/W, flattering the endcap comparison.
    ref_week = base_week or last_week
    control = [sn for sn, d in (all_store_weeks.get(ref_week) or {}).items()
               if sn not in roster and ENDCAP_SKU in (d.get("skus") or {})]

    # ── Per-SKU U/S/W by segment (like-for-like cohorts) ─────────────────────
    # The 15 lb lift table below divides by every roster store in the segment,
    # so an endcap store that wasn't carrying the SKU before the program counts
    # as a zero in the baseline and its ramp reads as growth. That is the right
    # read for "what did the program do", but it cannot answer "what was this
    # store's U/S/W before?" — a store with no prior distribution has none.
    #
    # So this table fixes each cohort to the stores that CARRIED the SKU in the
    # pre-endcap week and holds that cohort constant across all three weeks.
    # Same stores, same denominator, every week: the U/S/W move is real rate of
    # sale, not distribution arriving. Segments are the three the field asks
    # about — endcap set, endcap not set, and everything off the endcap list.
    notset_all = set(roster) - set_ok          # incl. the not-visited stores
    non_endcap = [sn for sn in (all_store_weeks.get(ref_week) or {}) if sn not in roster]
    sku_ref    = base_week or last_week

    def _skus_at(sn, wk):
        return ((all_store_weeks.get(wk) or {}).get(sn) or {}).get("skus") or {}

    def _units(stores, sku, wk):
        if sku == "ALL":
            return sum(sum((v or {}).get("qty") or 0 for k, v in _skus_at(sn, wk).items()
                           if k in SKUS) for sn in stores)
        return sum((_skus_at(sn, wk).get(sku) or {}).get("qty") or 0 for sn in stores)

    def sku_row(label, key, stores, sku):
        # Cohort = segment stores carrying this SKU in the pre-endcap week.
        cohort = [sn for sn in stores
                  if (any(k in SKUS for k in _skus_at(sn, sku_ref))
                      if sku == "ALL" else sku in _skus_at(sn, sku_ref))]
        n = len(cohort)
        out = {"key": key, "label": label, "n": n}
        for tag, wk in (("base", base_week), ("last", last_week), ("cur", this_week)):
            if not wk:
                out[tag + "_units"], out[tag + "_usw"] = None, None
                continue
            u = _units(cohort, sku, wk)
            out[tag + "_units"] = round(u)
            out[tag + "_usw"]   = round(u / n, 3) if n else None
        b, l, c = out["base_usw"], out["last_usw"], out["cur_usw"]
        out["wow_pct"]     = round((c - l) / l * 100, 1) if (l and c is not None) else None
        out["vs_base_pct"] = round((c - b) / b * 100, 1) if (b and c is not None) else None
        return out

    SKU_SEGS = [("set",     "Endcap — confirmed set",     set_ok),
                ("notset",  "Endcap — not confirmed set", notset_all),
                ("control", "Non-endcap stores",          non_endcap)]
    by_sku = {}
    for sku in ["ALL"] + list(SKUS):
        rows_ = [sku_row(lab, k, st, sku) for k, lab, st in SKU_SEGS]
        r_of  = {r["key"]: r for r in rows_}
        def _gap(a, b, f):
            x, y = r_of[a].get(f), r_of[b].get(f)
            return None if (x is None or y is None) else round(x - y, 1)
        by_sku[sku] = {
            "sku": sku,
            "rows": rows_,
            "net_vs_control": _gap("set", "control", "vs_base_pct"),
            "net_vs_notset":  _gap("set", "notset",  "vs_base_pct"),
        }

    def lift_row(label, stores, key=None):
        stores = list(stores)
        n = len(stores)
        out = {"label": label, "key": key, "n": n,
               "units": round(sum(qty[this_week].get(sn, 0) for sn in stores)),
               "usw": None}
        if n:
            out["usw"] = round(out["units"] / n, 2)
        for tag, wk in (("wow", last_week), ("vs_base", base_week)):
            if not wk:
                continue
            prev = sum(qty[wk].get(sn, 0) for sn in stores)
            out[tag + "_units"] = round(prev)
            out[tag + "_usw"]   = round(prev / n, 2) if n else None
            out[tag + "_pct"]   = (round((out["units"] - prev) / prev * 100, 1)
                                   if prev else None)
        return out

    lift = [lift_row("Endcap confirmed set", set_ok, "set")]
    if set_fu:
        lift += [lift_row("· set on the WK27 visit", set_w27, "set_w27"),
                 lift_row("· set on the follow-up sweep", set_fu, "set_fu")]
    lift.append(lift_row("Not set — all reasons", notset, "notset"))
    for _a, (key, label) in ENDCAP_REASONS.items():
        lift.append(lift_row("· " + label, {sn for sn in notset if seg_of[sn] == key}, key))
    if unvis:
        lift.append(lift_row("Not visited yet", unvis, "unvisited"))
    lift.append(lift_row("Control: all non-endcap stores", control, "control"))

    # Weekly set-vs-control U/S/W series, windowed the same way as the cohort
    # chart (ENDCAP_CHART_PRE_WEEKS of pre-set context, then everything after) so
    # the two charts share an x-axis and the lift reads as a trend rather than a
    # single pair of weeks.
    series_start = pre[-ENDCAP_CHART_PRE_WEEKS:][0] if pre else weeks[0]
    series = []
    for w in [w for w in weeks if w >= series_start]:
        if w not in qty:
            qty[w] = q15(w)
        row = {"week": w}
        for key, sset in (("set", set_ok), ("set_w27", set_w27), ("set_fu", set_fu),
                          ("notset", notset), ("control", control)):
            if not sset:
                continue
            n = len(sset)
            u = sum(qty[w].get(sn, 0) for sn in sset)
            row[key + "_units"] = round(u)
            row[key + "_usw"]   = round(u / n, 3) if n else None
        series.append(row)

    inv_rows = []
    seg_order = [("set", "Confirmed set")] + \
        [(k, l) for _a, (k, l) in ENDCAP_REASONS.items()] + [("unvisited", "Not visited yet")]
    for key, label in seg_order:
        stores = [sn for sn in roster if seg_of[sn] == key]
        if not stores:
            continue
        c = Counter(inv_of[sn] for sn in stores)
        inv_rows.append({"key": key, "label": label, "n": len(stores),
                         **{k: c.get(k, 0) for k in INV_KEYS}})
    inv_tot = Counter(inv_of.values())

    # Two different "lift vs control" readings, because the endcap list got a
    # 36-bag allocation whether or not the display was ever built:
    #   net_vs_control = set-store growth minus control growth. This is the whole
    #     program effect — allocation AND display — not the display alone.
    #   net_vs_notset  = set-store growth minus NOT-set endcap-store growth. Both
    #     groups got the inventory, so the gap isolates building the display.
    row_of = {r["key"]: r for r in lift}
    set_row, ns_row, ctl_row = row_of["set"], row_of["notset"], row_of["control"]
    def gap(a, b):
        if a.get("vs_base_pct") is None or b.get("vs_base_pct") is None:
            return None
        return round(a["vs_base_pct"] - b["vs_base_pct"], 1)

    return {
        "week":         this_week,
        "last_week":    last_week,
        "base_week":    base_week,
        "survey_week":  ENDCAP_SURVEY_WEEK,
        "live_date":    ENDCAP_LIVE_DATE,
        "sku":          ENDCAP_SKU,
        "units_target": ENDCAP_UNITS,
        "n_endcap":     len(roster),
        "n_set":        len(set_ok),
        "n_set_w27":    len(set_w27),
        "n_set_fu":     len(set_fu),
        "followup_week": ENDCAP_FOLLOWUP_WEEK if set_fu else None,
        "n_notset":     len(notset),
        "n_unvisited":  len(unvis),
        "n_visited":    len(visited),
        "set_pct":      round(len(set_ok) / len(visited) * 100, 1) if visited else 0,
        "set_pct_all":  round(len(set_ok) / len(roster) * 100, 1) if roster else 0,
        "reasons":      [{"key": k, "label": l,
                          "n": sum(1 for sn in notset if seg_of[sn] == k),
                          "received": sum(1 for sn in notset if seg_of[sn] == k
                                          and inv_of[sn] == "received")}
                         for _a, (k, l) in ENDCAP_REASONS.items()],
        "inv": {
            "counts": {k: inv_tot.get(k, 0) for k in INV_KEYS},
            "pct":    {k: round(inv_tot.get(k, 0) / len(roster) * 100, 1) for k in INV_KEYS},
            "rows":   inv_rows,
            "set_received":     sum(1 for sn in set_ok if inv_of[sn] == "received"),
            "set_short":        sum(1 for sn in set_ok if inv_of[sn] != "received"),
            "received_selling": sum(1 for sn in roster if inv_of[sn] == "received"
                                    and qty[this_week].get(sn, 0) > 0),
        },
        "lift":           lift,
        "by_sku":         by_sku,
        "sku_order":      list(SKUS),
        "series":         series,
        "pre_weeks":      ENDCAP_CHART_PRE_WEEKS,
        "net_vs_control": gap(set_row, ctl_row),
        "net_vs_notset":  gap(set_row, ns_row),
        "inc_units":      round(set_row["units"] - set_row.get("vs_base_units", set_row["units"])),
        "set_zero":       sum(1 for sn in set_ok if qty[this_week].get(sn, 0) == 0),
    }


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

    # Chart window: the last ENDCAP_CHART_PRE_WEEKS weeks before the endcap was
    # set, then everything after. The table below the chart keeps full history.
    chart_weeks = weeks[weeks.index(pre[-ENDCAP_CHART_PRE_WEEKS:][0]):] if pre else weeks

    return {
        "live_date":     ENDCAP_LIVE_DATE,
        "baseline_week": baseline,
        "weeks":         weeks,
        "chart_weeks":   chart_weeks,
        "pre_weeks":     ENDCAP_CHART_PRE_WEEKS,
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
    # Run rate = latest week's full-discount-EQUIVALENT units (both scents),
    # i.e. that week's markdown dollars / $2.27 — go-forward the same price mix
    # keeps drawing on the fund at the same rate. Falls back to full 15 lb
    # volume if the classification returned nothing for the latest week.
    run = sum((rollback["by_week"].get(as_of_wk) or {}).values())
    if not run:
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


def build_rollback(metrics, week_dates, store_rows_by_week=None):
    """Measure the 15 lb rollback from per-store realized prices. See ROLLBACK.

    For every rollback-era store-week the AUR (POS $ / POS qty) is that store's
    actual shelf price, so each store-week is classified against the $18.24
    pre-rollback price and its markdown dollars are accumulated:

      units_discounted  units at stores priced strictly below $18.24
      units_at_rollback units at stores priced at $15.97 (the headline price)
      units_full_price  units at stores priced at or above $18.24
      markdown_dollars  sum of qty * min(pre - AUR, $2.27) over discounted
                        stores — the actual draw against the co-op fund
      units_equiv       markdown_dollars / $2.27, i.e. full-discount-equivalent
                        units. This is what `units_total` carries, because the
                        $150k co-op converts to units only at $2.27 apiece.

    Also returns per-week price bands (units and share by distinct price point)
    so the price mix is visible rather than inferred.

    Weeks without per-store POS $ fall back to the week-level blended AUR.
    Returns None if no rollback-era week has data yet.
    """
    rb_date  = date.fromisoformat(ROLLBACK["date"])
    price    = ROLLBACK["price"]
    pre_hdr  = ROLLBACK["pre_price"]
    skus     = ROLLBACK["skus"]
    eps      = ROLLBACK["band_eps"]
    discount = round(pre_hdr - price, 2)                 # $2.27 program discount
    weeks    = sorted(w for w in week_dates)
    store_rows_by_week = store_rows_by_week or {}

    def implied(w, sku):
        m = metrics.get(w, {}).get(sku, {})
        d, q = m.get("pos_dollars"), m.get("pos_qty")
        return (d / q) if d and q else None

    # Per-SKU pre-rollback baseline = mean implied price over up to 4 fully-pre
    # weeks (week END before the rollback date). Reported for reference; the
    # classification below uses the headline $18.24 so the threshold is the one
    # the co-op deal was written against, not a drifting blend.
    baseline = {}
    for sku in skus:
        pre = [implied(w, sku) for w in weeks
               if date.fromisoformat(week_dates[w]) < rb_date]
        pre = [p for p in pre if p]
        baseline[sku] = round(sum(pre[-4:]) / len(pre[-4:]), 4) if pre else pre_hdr

    tot = {k: 0.0 for k in ("units", "disc_u", "roll_u", "at_pre_u",
                            "above_u", "markdown", "disc_dol")}
    by_week          = {}      # co-op-equivalent units per sku (drives build_coop)
    by_week_disc     = {}      # strict units-below-$18.24 per week
    units_by_sku     = {sku: 0 for sku in skus}
    disc_by_sku      = {sku: 0 for sku in skus}
    price_bands      = {}
    start_week, partial = None, None

    for w in weeks:
        end   = date.fromisoformat(week_dates[w])
        start = end - timedelta(days=6)
        if end < rb_date:
            continue                                    # fully pre-rollback
        if start_week is None:
            start_week = w
        srows = [r for r in (store_rows_by_week.get(w) or [])
                 if r["item_name"] in skus and r.get("pos_dollars") is not None
                 and (r.get("pos_qty") or 0) > 0]

        wk = {k: 0.0 for k in ("units", "disc_u", "roll_u", "at_pre_u",
                               "above_u", "markdown", "disc_dol")}
        band_units = {}          # price point (rounded cents) -> units
        row, row_disc = {}, {}

        for sku in skus:
            mine = [r for r in srows if r["item_name"] == sku]
            if not mine:
                # Fallback: one synthetic "store" at the week's blended AUR.
                q = metrics.get(w, {}).get(sku, {}).get("pos_qty")
                a = implied(w, sku)
                if not q or a is None:
                    continue
                mine = [{"pos_qty": q, "pos_dollars": q * a}]

            s_md, s_disc, s_dol = 0.0, 0.0, 0.0
            for r in mine:
                q   = r["pos_qty"]
                aur = round(r["pos_dollars"] / q, 2)
                wk["units"] += q
                band_units[aur] = band_units.get(aur, 0) + q
                if aur < pre_hdr - eps:
                    s_disc += q
                    s_dol  += r["pos_dollars"]
                    # Cap at the program discount so deep clearance ($9.97 etc.)
                    # can't over-draw a fund that only pays $2.27 a bag.
                    s_md += q * min(pre_hdr - aur, discount)
                    if abs(aur - price) <= 0.01:
                        wk["roll_u"] += q
                elif aur <= pre_hdr + eps:
                    wk["at_pre_u"] += q
                else:
                    wk["above_u"] += q

            wk["disc_u"]   += s_disc
            wk["markdown"] += s_md
            wk["disc_dol"] += s_dol
            equiv = int(round(s_md / discount)) if discount else 0
            if equiv:
                row[sku] = equiv
                units_by_sku[sku] += equiv
            if s_disc:
                row_disc[sku] = int(round(s_disc))
                disc_by_sku[sku] += int(round(s_disc))

        by_week[w]      = row
        by_week_disc[w] = row_disc
        for k in tot:
            tot[k] += wk[k]

        # Price bands: keep the biggest distinct points, fold the rest into
        # "other" so the mix reads at a glance without 60 rows of noise. The kept
        # points are then re-sorted by PRICE so a stacked bar reads left-to-right
        # as a price ramp, with the above-pre-price tail at the right end.
        u = wk["units"]
        pts = sorted(band_units.items(), key=lambda kv: -kv[1])
        keep, rest = pts[:ROLLBACK["top_bands"]], pts[ROLLBACK["top_bands"]:]
        keep.sort(key=lambda kv: kv[0])
        bands = [{"price": p, "units": int(round(q)),
                  "pct": round(q / u * 100, 1) if u else 0,
                  "vs_pre": ("below" if p < pre_hdr - eps else
                             "at" if p <= pre_hdr + eps else "above")}
                 for p, q in keep]
        if rest:
            rq = sum(q for _p, q in rest)
            bands.append({"price": None, "n_points": len(rest),
                          "units": int(round(rq)),
                          "pct": round(rq / u * 100, 1) if u else 0,
                          "vs_pre": "mixed"})
        price_bands[w] = {
            "units":       int(round(u)),
            "bands":       bands,
            "below_pct":   round(wk["disc_u"]   / u * 100, 1) if u else 0,
            "at_pct":      round(wk["at_pre_u"] / u * 100, 1) if u else 0,
            "above_pct":   round(wk["above_u"]  / u * 100, 1) if u else 0,
            "below_units": int(round(wk["disc_u"])),
            "at_units":    int(round(wk["at_pre_u"])),
            "above_units": int(round(wk["above_u"])),
            "roll_units":  int(round(wk["roll_u"])),
            "roll_pct":    round(wk["roll_u"] / u * 100, 1) if u else 0,
            "markdown":    round(wk["markdown"], 2),
            "blended":     w == start_week,   # transition week: AURs are real mixes
        }

        if start < rb_date <= end and partial is None:
            partial = {"week": w, "days_pre": (rb_date - start).days,
                       "days_roll": (end - rb_date).days + 1}

    if start_week is None:
        return None

    units_equiv = sum(units_by_sku.values())
    u_all       = tot["units"]
    share = lambda v: round(v / u_all * 100, 1) if u_all else 0
    return {
        "date":            ROLLBACK["date"],
        "pre_price":       pre_hdr,
        "price":           price,
        "discount":        discount,
        "skus":            skus,
        "start_week":      start_week,
        "baseline_price":  baseline,
        "by_week":         by_week,           # co-op-equivalent units per sku
        "by_week_disc":    by_week_disc,      # strict units below $18.24
        "units_by_sku":    units_by_sku,
        "units_total":     units_equiv,       # == units_equiv (co-op burn basis)
        "units_equiv":     units_equiv,
        "units_discounted":   int(round(tot["disc_u"])),
        "disc_by_sku":        disc_by_sku,
        "units_at_rollback":  int(round(tot["roll_u"])),
        "units_at_pre":       int(round(tot["at_pre_u"])),
        "units_above_pre":    int(round(tot["above_u"])),
        "units_full_price":   int(round(tot["at_pre_u"] + tot["above_u"])),
        "units_all":          int(round(u_all)),
        "pct_discounted":     share(tot["disc_u"]),
        "pct_at_rollback":    share(tot["roll_u"]),
        "pct_above_pre":      share(tot["above_u"]),
        "markdown_dollars":   round(tot["markdown"], 2),
        # Back-compat: this key's label has always been "units at/above the
        # pre-rollback price", which is now what it actually holds.
        "units_excluded":  int(round(tot["at_pre_u"] + tot["above_u"])),
        # Actual POS dollars rung on the discounted units (not units x $15.97 —
        # the $16.97 tier makes that estimate low).
        "dollars_total":   round(tot["disc_dol"], 2),
        "dollars_discounted": round(tot["disc_dol"], 2),
        "price_bands":     price_bands,
        "partial":         partial,
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
    # Some weeks ship a second, endcap-only cut of the same sheet ("CATALYST
    # Endcap Sales by Store", wk202629). It matches every keyword the real
    # store feed does but covers only the ~1,880 endcap stores, so picking it
    # silently drops 55% of the chain. Never let an endcap cut win either slot.
    is_endcap = lambda s: "endcap" in s

    def is_instore(s):
        if is_ecomm(s) or "by store" in s or is_endcap(s):
            return False
        if any(k in s for k in ("inventory", "forecast", "supply", "demand",
                                "modular", "order", "plan")):
            return False
        return "sales" in s or "instore" in s or "in store" in s

    return {
        "instore":   pick(is_instore),
        "bystore":   pick(lambda s: "by store" in s and not is_endcap(s)),
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
    pos_sales_col  = next((i for i, h in enumerate(header) if "POS Sales" in h), None)
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

    def to_float(v):
        try:
            f = float(v)
            return None if math.isnan(f) else f
        except (ValueError, TypeError):
            return None

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
                "pos_dollars":  to_float(row.iloc[pos_sales_col]) if pos_sales_col is not None else None,
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


# Brands carried through to the dashboard's Ecomm tab. Feline Fresh is still
# PARSED (so its rows don't trip the unmapped-row warning and FELINE_FRESH_MAP
# keeps documenting the taxonomy) but is dropped here — this dashboard is the
# Catalyst report, and the Ecomm tab is the only consumer of the ecomm payload.
# Add a brand back to this set to surface it again.
ECOMM_BRANDS = {"Catalyst"}


def extract_ecomm_data(df):
    """
    Parse a Lignetics L52WK Ecomm (or LW Ecomm) sheet.
    Row 0: headers; Rows 1+: products (skip 'Total' rows and unknowns).
    Cols: 0=Product Name, 1=Net Retail Sales, 3=Net Unit Sales.
    Returns {short_name: {"brand": str, "r": float|None, "u": int|None}},
    restricted to ECOMM_BRANDS.
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
            if brand not in ECOMM_BRANDS:
                continue
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
    # 0. Sync any missing weekly workbooks from the reports inbox before
    # building (auto-ingested weeks otherwise live only in the private data
    # repo, and the week-regression guard would abort the build). Non-fatal:
    # cloud runs have no credentials module and already have the files.
    try:
        from _download_missing_weeks import download_missing_weeks
        download_missing_weeks()
    except Exception as e:
        print(f"[WARN] inbox sync skipped ({e}) — using workbooks on disk")

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
                n_stores = len({r["store_num"] for r in rows})
                print(f"  Sales by Store: {len(rows)} rows ({n_stores} stores)")
                # Coverage guard. The store feed grows slowly week to week, so a
                # double-digit drop means the wrong sheet was picked (an endcap-
                # only or single-SKU cut), not a real distribution change.
                if store_weeks_list[:-1]:
                    prev_w = store_weeks_list[-2]
                    prev_n = len({r["store_num"] for r in raw_store_rows_by_week[prev_w]})
                    if prev_n and n_stores < prev_n * 0.9:
                        print(f"  [WARN] Store coverage fell {1 - n_stores / prev_n:.0%} vs "
                              f"{prev_w} ({prev_n} → {n_stores} stores) from sheet "
                              f"'{bystore_sheet}' — check that this is the full-chain "
                              f"feed and not a partial cut.")
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

    # Price rollback: classify every rollback-era store-week by its realized
    # price (AUR) against the $18.24 pre-rollback price.
    rollback = build_rollback(metrics, week_dates, raw_store_rows_by_week)
    if rollback:
        p = rollback.get("partial")
        pnote = (f", partial start week {p['week']}: {p['days_roll']}/7 days"
                 if p else "")
        print(f"\nRollback: 15 lb ${ROLLBACK['pre_price']:.2f}→${ROLLBACK['price']:.2f} "
              f"from {ROLLBACK['date']} (start {rollback['start_week']}{pnote})")
        print(f"  Units below ${ROLLBACK['pre_price']:.2f}: "
              f"{rollback['units_discounted']:,} of {rollback['units_all']:,} "
              f"({rollback['pct_discounted']:.1f}%)  |  at ${ROLLBACK['price']:.2f} exactly: "
              f"{rollback['units_at_rollback']:,} ({rollback['pct_at_rollback']:.1f}%)  |  "
              f"at/above pre price: {rollback['units_full_price']:,} "
              f"(of which {rollback['units_above_pre']:,} ABOVE, "
              f"{rollback['pct_above_pre']:.1f}%)")
        print(f"  Markdown ${rollback['markdown_dollars']:,.0f} "
              f"= {rollback['units_equiv']:,} full-discount-equivalent units "
              f"(co-op burn basis)")
        lw = sorted(rollback["price_bands"])[-1]
        pb = rollback["price_bands"][lw]
        mix = "  ".join(
            (f"${b['price']:.2f}:{b['pct']:.0f}%" if b["price"] is not None
             else f"other({b['n_points']}):{b['pct']:.0f}%")
            for b in pb["bands"])
        print(f"  {lw} price mix ({pb['units']:,} units): {mix}")

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

    # Endcap rollout scorecard: set rate (field survey), 36-bag arrival, lift.
    endcap_status = build_endcap_status(all_store_weeks, endcap_data, week_dates)
    if endcap_status:
        es, iv = endcap_status, endcap_status["inv"]
        sr = es["lift"][0]
        print(f"\nEndcap rollout (survey wk{es['survey_week']}, sales wk{es['week']}): "
              f"{es['n_set']:,} set of {es['n_visited']:,} visited ({es['set_pct']:.0f}%), "
              f"{es['n_notset']:,} not set, {es['n_unvisited']:,} not visited")
        print(f"  36-bag allocation: {iv['counts']['received']:,} received "
              f"({iv['pct']['received']:.0f}%), {iv['counts']['transit']:,} in transit, "
              f"{iv['counts']['onorder']:,} on order, {iv['counts']['short']:,} short; "
              f"set stores without product: {iv['set_short']:,}")
        wow = sr.get("wow_pct")
        vb  = sr.get("vs_base_pct")
        print(f"  Set-store 15O units {sr['units']:,} ({sr['usw']} U/S/W): "
              f"WoW {wow if wow is None else f'{wow:+.1f}%'} vs {es['last_week']}, "
              f"{vb if vb is None else f'{vb:+.1f}%'} vs pre-endcap {es['base_week']}; "
              f"control {es['lift'][-1].get('vs_base_pct')}% vs base")
        print(f"  Per-SKU U/S/W, like-for-like cohorts (base {es['base_week']} → "
              f"{es['last_week']} → {es['week']}):")
        for sku in ["ALL"] + es["sku_order"]:
            blk = es["by_sku"][sku]
            print(f"    {('All 4 SKUs' if sku == 'ALL' else sku):<22}"
                  f" net vs control {blk['net_vs_control']}pp | "
                  f"net vs not-set {blk['net_vs_notset']}pp")
            for r in blk["rows"]:
                f2 = lambda v: "—" if v is None else f"{v:.3f}"
                fp = lambda v: "—" if v is None else f"{v:+.1f}%"
                print(f"      {r['label']:<30} n={r['n']:<5} "
                      f"base {f2(r['base_usw'])} → last {f2(r['last_usw'])} → "
                      f"now {f2(r['cur_usw'])}  WoW {fp(r['wow_pct']):<8} "
                      f"vs pre-endcap {fp(r['vs_base_pct'])}")

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
        "endcap_status":  endcap_status,
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

    # Regression guard: never overwrite the published dashboard with one that
    # DROPS weeks. This catches the failure mode where a local rebuild runs
    # against a folder missing auto-ingested source files (weeks live only in
    # the private data repo) and would silently clobber the cloud's newer build
    # on push. Compare the new week set against the weeks embedded in the
    # currently-committed dashboard.html. Override with ALLOW_WEEK_REGRESSION=1
    # for a deliberate removal.
    new_weeks = set(data["weeks"])
    if os.path.exists(output_path):
        try:
            with open(output_path, "r", encoding="utf-8") as f:
                prev_html = f.read()
            m = re.search(r'"weeks":\s*\[([^\]]*)\]', prev_html)
            prev_weeks = set(re.findall(r"20\d{4}", m.group(1))) if m else set()
        except Exception:
            prev_weeks = set()
        dropped = prev_weeks - new_weeks
        if dropped and os.environ.get("ALLOW_WEEK_REGRESSION") != "1":
            print(f"\n[ABORT] Refusing to overwrite dashboard.html — this build DROPS "
                  f"already-published week(s): {sorted(dropped)}")
            print( "        The local folder is likely missing auto-ingested source "
                   "files (they live in the private data repo / reports inbox).")
            print( "        Fix: run  python _download_missing_weeks.py  to pull the "
                   "missing weekly workbooks, then rebuild.")
            print( "        To drop weeks on purpose, set ALLOW_WEEK_REGRESSION=1.")
            sys.exit(1)

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
    elif os.environ.get("EMAIL_DRY_RUN") == "1":
        # Render email_preview.html from the real build data (the standalone
        # `python email_report.py --dry-run` path re-runs a reduced ETL that has
        # no rollback or endcap data, so its preview is missing those sections).
        try:
            from email_report import send_report
            send_report(data, dry_run=True)
        except Exception as e:
            print(f"  [Email] Dry-run error: {e}")
    else:
        try:
            from email_report import send_report
            send_report(data, dev_only=dev_only)
            # Mark the day on ANY full-distro send, not just the Monday one. A
            # FORCE_FULL_DISTRO cloud run used to leave the flag unwritten, so a
            # second ingest of the same week sailed straight through this gate.
            # send_report() also refuses same-week duplicates on its own; this
            # keeps later runs on the dev list rather than relying on that.
            if not dev_only:
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
