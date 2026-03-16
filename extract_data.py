"""
extract_data.py — Catalyst Pet Walmart Sales Dashboard Generator
Reads 6 weekly Excel files, geocodes stores, computes metrics, writes dashboard.html
"""

import json
import os
import glob
import re
import math
import pandas as pd
import pgeocode

# ─── Configuration ────────────────────────────────────────────────────────────

SHEET_MAP = {
    "202601": {"instore": "202601 In Store Total Catalyst", "bystore": None},
    "202602": {"instore": "Catalyst LW Sales Total",        "bystore": "202602 Sales by Store"},
    "202603": {"instore": "LW Instore Sales Total ",        "bystore": "Sales by Store"},
    "202604": {"instore": "202604 InStore Sales",           "bystore": "202604 Sales by Store"},
    "202605": {"instore": "Catalyst LW Sales",              "bystore": "Catalyst Sales by Store"},
    "202606": {"instore": "Catalyst LW Sales Total",        "bystore": "Catalyst Sales by Stores"},
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

GEO_CACHE_FILE = "stores_geo.json"
TEMPLATE_FILE  = "dashboard_template.html"
OUTPUT_FILE    = "dashboard.html"

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
    """Auto-detect all 2026XX Weekly Sales Report Catalyst.xlsx files."""
    pattern = os.path.join(os.path.dirname(__file__), "2026?? Weekly Sales Report Catalyst.xlsx")
    files = glob.glob(pattern)
    result = {}
    for f in files:
        m = re.search(r"(2026\d{2})", os.path.basename(f))
        if m:
            result[m.group(1)] = f
    return result


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

    # Map row index 2–5 to SKU labels
    sku_rows = {2: SKUS[0], 3: SKUS[1], 4: SKUS[2], 5: SKUS[3]}

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

        wholesale = None
        if pos_qty is not None:
            wholesale = round(pos_qty * WHOLESALE_PRICE[sku], 2)

        result[sku] = {
            "pos_dollars":      pos_dollars,
            "pos_qty":          pos_qty,
            "instock_pct":      instock_pct,
            "usw":              usw,
            "wholesale_dollars": wholesale,
        }

    # Total row (row 6)
    row = df.iloc[6]
    pos_dollars_t = safe_float(row.iloc[1])
    pos_qty_raw_t = safe_float(row.iloc[7])
    pos_qty_t = int(pos_qty_raw_t) if pos_qty_raw_t is not None else None
    instock_pct_t = safe_float(row.iloc[9], 100)
    usw_t = safe_float(row.iloc[33])

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
    }

    return result


def extract_store_data(week, df):
    """
    Parse Sales by Store sheet.
    Col 0: item_name | Col 1: store_number | Col 2: street | Col 3: state
    Col 4: city | Col 5: zip | Col 7: POS Qty | Col 8: On Hand Qty
    Returns list of dicts.
    """
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
            pos_qty_raw  = row.iloc[7]
            on_hand_raw  = row.iloc[8]

            zip5 = normalize_zip(zip_raw)

            def to_int(v):
                try:
                    f = float(v)
                    return int(f) if not math.isnan(f) else 0
                except (ValueError, TypeError):
                    return 0

            rows.append({
                "item_name":  item_name,
                "store_num":  store_num,
                "street":     street,
                "state":      state_raw,
                "city":       city,
                "zip5":       zip5,
                "pos_qty":    to_int(pos_qty_raw),
                "on_hand":    to_int(on_hand_raw),
            })
        except Exception:
            continue
    return rows


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


# ─── Consecutive OOS computation ──────────────────────────────────────────────

def compute_consecutive_oos_by_week(all_store_weeks, store_weeks):
    """
    For each week W (in store_weeks order), for each (sku, store) pair,
    count trailing consecutive OOS weeks ending at W.
    OOS = on_hand == 0

    Returns: {week: {sku: {store_num: consecutive_count}}}
    """
    result = {w: {sku: {} for sku in SKUS} for w in store_weeks}

    # Build timeline: for each (sku, store) — list of (week, on_hand)
    timeline = {}  # (sku, store_num) -> {week: on_hand}
    for week in store_weeks:
        week_data = all_store_weeks.get(week, {})
        for store_num, sdata in week_data.items():
            for sku, skudata in sdata.get("skus", {}).items():
                key = (sku, store_num)
                if key not in timeline:
                    timeline[key] = {}
                timeline[key][week] = skudata.get("on_hand", None)

    for (sku, store_num), week_map in timeline.items():
        for w_idx, week in enumerate(store_weeks):
            if week not in week_map:
                continue
            # Count trailing OOS ending at this week
            count = 0
            for past_week in reversed(store_weeks[:w_idx + 1]):
                oh = week_map.get(past_week)
                if oh is None:
                    break   # store didn't report that week — stop chain
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
                "qty":     r["pos_qty"],
                "on_hand": r["on_hand"],
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

    for week in sorted(files.keys()):
        filepath = files[week]
        sheet_info = SHEET_MAP.get(week)
        if not sheet_info:
            print(f"  [WARN] No sheet map for {week}, skipping.")
            continue

        print(f"\nProcessing {week}...")

        # InStore metrics
        instore_sheet = sheet_info["instore"]
        try:
            df_in = pd.read_excel(filepath, sheet_name=instore_sheet, header=None)
            metrics[week] = extract_instore_metrics(week, df_in)
            print(f"  InStore metrics: OK ({len(metrics[week])} entries)")
        except Exception as e:
            print(f"  [ERROR] InStore sheet '{instore_sheet}': {e}")

        # Sales by Store
        bystore_sheet = sheet_info["bystore"]
        if bystore_sheet:
            try:
                df_bs = pd.read_excel(filepath, sheet_name=bystore_sheet, header=None)
                rows = extract_store_data(week, df_bs)
                raw_store_rows_by_week[week] = rows
                store_weeks_list.append(week)
                print(f"  Sales by Store: {len(rows)} rows")
            except Exception as e:
                print(f"  [ERROR] ByStore sheet '{bystore_sheet}': {e}")

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
    state_sales = compute_state_sales(all_store_weeks, stores)
    state_oos   = compute_state_oos(all_store_weeks, stores)
    consecutive_oos = compute_consecutive_oos_by_week(all_store_weeks, store_weeks_list)

    print(f"  Stores: {len(stores)}")
    print(f"  Store-weeks: {sorted(all_store_weeks.keys())}")

    # 6. Assemble JSON
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
    print("Done.")


if __name__ == "__main__":
    main()
