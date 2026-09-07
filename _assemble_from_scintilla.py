#!/usr/bin/env python3
"""Assemble '<WEEK> Weekly Sales Report Catalyst.xlsx' in Hailey's layout from our
own Scintilla Report Builder pulls — the backup path when the rep firm doesn't
send the weekly workbook (first used wk202631, Labor Day 2026).

Inputs (three saved reports in the Lignetics Scintilla account, each run once
with Walmart Calendar Week = "WMT Last week", downloaded as XLSX):

  Catalyst Weekly Sales by Store   -> "CATALYST Sales by Store"   (store x SKU)
  Catalyst Weekly Item Summary     -> "CATALYST- LW Sales"        (per-SKU ratios)
  Lignetics Weekly Ecomm           -> "LIGNETICS LW Ecomm" + "CATALYST LW Ecomm"

Output sheets use Hailey's names and column positions so extract_data.py's
detect_sheets() and the position-based parsers work unchanged.

Usage:
  python _assemble_from_scintilla.py 202631
  python _assemble_from_scintilla.py 202631 --bystore X.xlsx --items Y.xlsx --ecomm Z.xlsx
  python _assemble_from_scintilla.py 202631 --ecomm-basis shipped   (default: auth)

By default the three files are looked up in ~/Downloads by their Scintilla
report names ("... wk<WEEK>.xlsx").
"""
import argparse
import os
import sys

import openpyxl
from openpyxl import Workbook

ITEMS = {  # prime item nbr -> canonical item name (Hailey's row order)
    680268871: "CATALYST15ORIG",
    680065761: "CATALYST34LBORIGINAL",
    680268872: "CATALYST15UNSCEN",
    680065800: "CATALYSTPET34LBUNSCE",
}

# Hailey-format instore positions (0-based): parser reads rows 2-5 SKUs + row 6 Total
H_POS, H_QTY, H_INSTOCK, H_USW, H_MD = 1, 7, 9, 33, 41
H_WIDTH = 44

DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")


def load_sheet(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h or "") for h in rows[0]]

    def col(sub):
        for i, h in enumerate(hdr):
            if sub in h:
                return i
        raise SystemExit(f"column {sub!r} missing from {os.path.basename(path)}: {hdr}")

    return rows[1:], col


def check_week(rows, wcol, week, label):
    weeks = {r[wcol] for r in rows if r[wcol] is not None}
    if weeks != {int(week)}:
        raise SystemExit(f"{label} covers weeks {sorted(weeks)}, expected {week}")


# ── Sales by Store ────────────────────────────────────────────────────────────

def write_bystore(out_wb, path, week):
    rows, col = load_sheet(path)
    c = {
        "week":    col("walmart_calendar_week"),
        "desc":    col("prime_item_description"),
        "store":   col("store_number"),
        "street":  col("street_address_line_1"),
        "state":   col("state_or_province_name"),
        "city":    col("city_name"),
        "zip":     col("zip_code_or_postal_code"),
        "dc":      col("distribution_center_number"),
        "pos":     col("pos_sales_this_year"),
        "qty":     col("pos_quantity_this_year"),
        "onhand":  col("store_on_hand_quantity_this_year"),
        "transit": col("store_in_transit_quantity_this_year"),
        "onorder": col("store_on_order_quantity_this_year"),
        "traited": col("traited_store_count_this_year"),
    }
    check_week(rows, c["week"], week, "by-store export")

    bysums = {}
    tr = f"Time Range 1\nMIN:{week}\nMAX:{week}\n"
    ws = out_wb.create_sheet("CATALYST Sales by Store")
    ws.append([
        "item_name", "store_number", "street_address_line_1",
        "state_or_province_name", "city_name", "zip_code_or_postal_code",
        "distribution_center_number",
        tr + "POS Sales", tr + "POS Quantity", tr + "On Hand Quantity",
        tr + "In Transit Quantity", tr + "On Order Quantity",
        tr + "Total Pipeline Quantity", tr + "Traited Store Count",
        tr + "U/S/W Traited",
    ])
    n = 0
    for r in rows:
        if r[c["store"]] is None:
            continue
        name    = str(r[c["desc"]]).strip()
        pos     = float(r[c["pos"]] or 0)
        qty     = int(r[c["qty"]] or 0)
        onhand  = int(r[c["onhand"]] or 0)
        transit = int(r[c["transit"]] or 0)
        onorder = int(r[c["onorder"]] or 0)
        traited = int(r[c["traited"]] or 0)
        s, q = bysums.get(name, (0.0, 0))
        bysums[name] = (s + pos, q + qty)
        ws.append([
            name, int(r[c["store"]]),
            r[c["street"]], r[c["state"]], r[c["city"]], str(r[c["zip"]]),
            r[c["dc"]],
            pos, qty, onhand, transit, onorder, onhand + transit + onorder,
            traited, (qty if traited else None),
        ])
        n += 1
    bysums = {k: (round(v[0], 2), v[1]) for k, v in bysums.items()}
    return n, bysums


# ── In-store summary ──────────────────────────────────────────────────────────

def load_items(path, week):
    rows, col = load_sheet(path)
    c = {
        "week":    col("walmart_calendar_week"),
        "nbr":     col("prime_item_number"),
        "pos":     col("pos_sales_this_year"),
        "qty":     col("pos_quantity_this_year"),
        "instock": col("instock_percentage_this_year"),
        "traited": col("traited_store_count_this_year"),
        "mumd":    col("si_total_mumd_amount_this_year"),
    }
    check_week(rows, c["week"], week, "item summary export")
    by_item = {}
    for r in rows:
        nbr = r[c["nbr"]]
        if nbr in ITEMS:
            by_item[nbr] = {
                "pos":     float(r[c["pos"]] or 0),
                "qty":     int(r[c["qty"]] or 0),
                "instock": float(r[c["instock"]] or 0),
                "traited": int(r[c["traited"]] or 0),
                "mumd":    float(r[c["mumd"]] or 0),
            }
    if len(by_item) != 4:
        raise SystemExit(f"expected 4 catalyst item rows, found {sorted(by_item)}")
    return by_item


def write_instore(out_wb, by_item, bysums):
    """POS $/qty come from the by-store sums so the two sheets agree exactly
    (as Hailey's workbooks always do). Ratios come from the item summary:
    instock % as reported; U/S/W = qty / traited stores (that is what Hailey's
    'U/S/W TY' column is — 202629: 9268 / 3607 = 2.5694); Markdown % Sales =
    SI MUMD $ / POS $.
    """
    ws = out_wb.create_sheet("CATALYST- LW Sales")
    meta = [""] * H_WIDTH
    meta[0] = "Time Range Name"
    for cidx in (H_POS, H_QTY, H_INSTOCK, H_USW, H_MD):
        meta[cidx] = "LW"
    hdr = [""] * H_WIDTH
    hdr[0] = "All Links Item Desc"
    hdr[H_POS], hdr[H_QTY] = "POS $ TY", "POS Qty TY"
    hdr[H_INSTOCK], hdr[H_USW], hdr[H_MD] = "Instock % TY", "U/S/W TY", "Markdown % Sales TY"
    ws.append(meta)
    ws.append(hdr)

    def metric_row(name, pos, qty, instock, usw, md):
        row = [""] * H_WIDTH
        row[0] = name
        row[H_POS], row[H_QTY] = pos, qty
        row[H_INSTOCK], row[H_USW], row[H_MD] = instock, usw, md
        return row

    tot_pos = tot_mumd = instock_w = 0.0
    tot_qty = tot_traited = 0
    for nbr, name in ITEMS.items():
        it = by_item[nbr]
        pos, qty = bysums.get(name, (0.0, 0))
        usw = qty / it["traited"] if it["traited"] else None
        md  = it["mumd"] / pos if pos else None
        ws.append(metric_row(name, pos, qty, it["instock"], usw, md))
        tot_pos += pos
        tot_qty += qty
        tot_traited += it["traited"]
        tot_mumd += it["mumd"]
        instock_w += it["instock"] * it["traited"]
    tot_pos = round(tot_pos, 2)
    ws.append(metric_row(
        "Total", tot_pos, tot_qty,
        instock_w / tot_traited if tot_traited else None,
        tot_qty / tot_traited if tot_traited else None,
        tot_mumd / tot_pos if tot_pos else None,
    ))
    return tot_pos, tot_qty


# ── Ecomm ─────────────────────────────────────────────────────────────────────

def write_ecomm(out_wb, path, week, basis):
    rows, col = load_sheet(path)
    c = {
        "week":  col("walmart_calendar_week"),
        "name":  col("product_name"),
        "brand": col("brand_name"),
        "sales": col("auth_based_net_sales_amount_this_year" if basis == "auth"
                     else "shipped_based_net_sales_amount_this_year"),
        "units": col("auth_based_item_quantity_this_year" if basis == "auth"
                     else "shipped_based_quantity_this_year"),
    }
    check_week(rows, c["week"], week, "ecomm export")
    hdr = ["Product Name", "Net Retail Sales", "Net Retail Sales % Chg",
           "Net Unit Sales", "Net Unit Sales % Chg", "Net AUR", "LY Net AUR"]

    def product_rows(pred):
        out = []
        for r in rows:
            name = r[c["name"]]
            if not name or not pred(r):
                continue
            sales = float(r[c["sales"]] or 0)
            units = int(r[c["units"]] or 0)
            out.append([name, round(sales, 2), None, units, None,
                        (sales / units if units else None), None])
        out.sort(key=lambda x: -x[1])
        return out

    def write(sheet_name, pred):
        ws = out_wb.create_sheet(sheet_name)
        ws.append(hdr)
        prs = product_rows(pred)
        for pr in prs:
            ws.append(pr)
        ts = round(sum(p[1] for p in prs), 2)
        tu = sum(p[3] for p in prs)
        ws.append(["Total", ts, None, tu, None, (ts / tu if tu else None), None])
        return ts, tu

    lig = write("LIGNETICS LW Ecomm",
                lambda r: "bear mountain" not in str(r[c["brand"]]).lower())
    cat = write("CATALYST LW Ecomm",
                lambda r: str(r[c["brand"]]).strip().lower() == "catalyst pet")
    return lig, cat


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("week")
    ap.add_argument("--bystore")
    ap.add_argument("--items")
    ap.add_argument("--ecomm")
    ap.add_argument("--ecomm-basis", choices=["auth", "shipped"], default="auth")
    ap.add_argument("--out")
    a = ap.parse_args()
    week = a.week

    bystore = a.bystore or os.path.join(DOWNLOADS, f"Catalyst Weekly Sales by Store wk{week}.xlsx")
    items   = a.items   or os.path.join(DOWNLOADS, f"Catalyst Weekly Item Summary wk{week}.xlsx")
    ecomm   = a.ecomm   or os.path.join(DOWNLOADS, f"Lignetics Weekly Ecomm wk{week}.xlsx")
    out     = a.out     or f"{week} Weekly Sales Report Catalyst.xlsx"
    for p in (bystore, items, ecomm):
        if not os.path.exists(p):
            raise SystemExit(f"missing input: {p}")

    wb = Workbook()
    wb.remove(wb.active)
    by_item = load_items(items, week)
    # Hailey's order: LW Sales, LW Ecomm, ..., Sales by Store. Build by-store
    # first to get the sums, then move the instore sheet to the front.
    nrows, bysums = write_bystore(wb, bystore, week)
    tot_pos, tot_qty = write_instore(wb, by_item, bysums)
    lig, cat = write_ecomm(wb, ecomm, week, a.ecomm_basis)
    wb.move_sheet("CATALYST- LW Sales", offset=-3)
    wb.save(out)

    print(f"[OK] wrote {out}")
    print(f"  instore/bystore total: ${tot_pos:,.2f} / {tot_qty:,} units ({nrows} store rows)")
    print(f"  item summary POS     : ${sum(v['pos'] for v in by_item.values()):,.2f} / "
          f"{sum(v['qty'] for v in by_item.values()):,} units")
    for nbr, name in ITEMS.items():
        it = by_item[nbr]
        pos, qty = bysums.get(name, (0.0, 0))
        print(f"    {name:22s} ${pos:>12,.2f} {qty:>7,}  instock {it['instock']:.4f}  "
              f"usw {qty / it['traited'] if it['traited'] else 0:.3f}  traited {it['traited']}")
    print(f"  ecomm ({a.ecomm_basis}): Lignetics ${lig[0]:,.2f} / {lig[1]} units; "
          f"Catalyst ${cat[0]:,.2f} / {cat[1]} units")


if __name__ == "__main__":
    main()
