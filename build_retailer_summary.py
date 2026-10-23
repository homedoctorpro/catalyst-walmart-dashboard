"""Retailer pipeline summary workbook for investors — a slice of the dashboard's Retailers tab.

Reads the retailer list (name, channel, doors) from dashboard_template.html and the
live status (plus any U/S/W override) from the shared Google Sheet via the
dashboard's sync URL.

Sizing matches the dashboard exactly:
    annual units  = doors × U/S/W × 52
    wholesale $   = annual units × $10
U/S/W is the row override if one is set, else the channel default. Units and $ are
Excel formulas, so editing a U/S/W or door count in the workbook re-sizes the row.

This file goes outside the company, so it carries no internal fields (priority, rep
firm, deadlines, contacts) and the commentary is the short, reviewed line in
COMMENTARY below, never the sheet's working notes. A retailer with no entry there
gets a blank comment. Retailers we already sell to and Non-Target rows are left
out. So are Declined rows, except those in KEEP_DECLINED (Target), which are listed
but kept out of the totals, same as the dashboard.

    python build_retailer_summary.py      → Reports/Catalyst_Retailer_Summary.xlsx
"""
import json
import os
import re
import sys
import urllib.request
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Reports", "Catalyst_Retailer_Summary.xlsx")

STATUS_LABEL = {"in": "Currently In", "pitched": "Pitched", "sent": "Sent Materials",
                "requested": "Requested Materials", "target": "Target",
                "non-target": "Non-Target", "declined": "Declined", "": "Prospect"}
PARKED = {"in", "non-target", "declined"}
STATUS_ORDER = ["pitched", "sent", "requested", "target", "", "non-target", "declined"]
# Declined retailers still worth showing investors (listed, kept out of the totals).
KEEP_DECLINED = {"target"}

DIST_NOTE = "Distribution route; doors counted under independents"
# Investor-facing one-liners. Update these by hand when a retailer's story changes.
COMMENTARY = {
    "petsmart": "Pitched; follow-up in progress",
    "petco": "Pitched; follow-up in progress",
    "food-lion": "Buyer meeting set for the 2027 cat litter cycle",
    "bradley-caldwell": "Samples and distributor terms sent",
    "publix": "Buyer requested materials",
    "tractor-supply": "Catalyst and private label discussion",
    "pet-supplies-plus": "Teaser deck sent; awaiting response",
    "pet-supermarket": "New buyer; awaiting response",
    "dollar-general": "Existing Lignetics supplier (BBQ pellets)",
    "aldi": "Private label opportunity",
    "trader-joes": "Private label opportunity",
    "walmart-canada": "Feline Fresh approved online; store distribution in review",
    "heb": "Meeting being arranged",
    "bomgaars": "Carries Feline Fresh; price sensitive",
    "rural-king": "Price sensitive",
    "kahoots": "Needs a distributor",
    "kriser": "Needs a distributor",
    "phillips-pet": "New vendors paused during Central integration; revisit Dec",
    "ingles": "Slotting ~$8K/SKU; launch with 15 lb",
    "petsense": "~75-store closure announced in 2026",
    "southern-states": "~1,200 dealer outlets",
    "indie-via-distributors": "Reached through the distributors listed; doors counted once",
    "target": "Revisit next year",
}


def load_template():
    src = open(os.path.join(HERE, "dashboard_template.html"), encoding="utf-8").read()
    js_num = lambda name: float(re.search(rf"const {name} = ([\d.]+);", src).group(1))
    consts = {k: js_num(k) for k in ("RT_WHOLESALE_PRICE", "RT_WEEKS_PER_YEAR")}

    ch_block = src[src.index("const RT_CHANNELS"):src.index("const RT_CHANNEL_ORDER")]
    channels = {k: {"name": n, "usw": float(u)} for k, n, u in re.findall(
        r'(\w+):\s*\{\s*name:\s*"([^"]+)",\s*color:\s*"[^"]+",\s*defaultUSW:\s*([\d.]+)', ch_block)}
    order = re.findall(r'"(\w+)"', re.search(r"const RT_CHANNEL_ORDER = \[([^\]]+)\]", src).group(1))

    block = src[src.index("const RT_RETAILERS"):src.index("const RT_STATUS_OPTIONS")]
    rows = []
    for obj in re.findall(r"\{\s*id:[^{}]*\}", block):
        f = lambda k: (re.search(rf'\b{k}:\s*"([^"]*)"', obj) or [None, ""])[1]
        rows.append({"id": f("id"), "name": f("name"), "channel": f("channel"),
                     "stores": int(re.search(r"\bstores:\s*(\d+)", obj).group(1)),
                     "defaultStatus": f("defaultStatus")})
    sync_url = re.search(r'RT_DEFAULT_SYNC_URL\s*=\s*"([^"]+)"', src).group(1)
    return consts, channels, order, rows, sync_url


def fetch_live(url):
    with urllib.request.urlopen(url + "?action=get", timeout=120) as r:
        payload = json.load(r)
    if not payload.get("ok"):
        raise RuntimeError(payload.get("error") or "sync endpoint returned ok:false")
    return payload.get("data") or {}


def main():
    consts, channels, order, rows, sync_url = load_template()
    try:
        live = fetch_live(sync_url)
    except Exception as e:  # still useful offline, but say so
        print(f"  [WARN] couldn't read the shared sheet ({e}); using built-in statuses only")
        live = {}

    price, weeks = consts["RT_WHOLESALE_PRICE"], consts["RT_WEEKS_PER_YEAR"]
    recs = []
    for r in rows:
        o = live.get(r["id"], {})
        status = o["status"] if o.get("status") is not None else r["defaultStatus"]
        if (status or "") not in STATUS_LABEL:
            print(f"  [WARN] {r['name']}: unknown status {status!r}")
        usw = o.get("usw")
        usw = float(usw) if usw not in (None, "") else channels[r["channel"]]["usw"]
        note = COMMENTARY.get(r["id"]) or (DIST_NOTE if r["channel"] == "distributor"
                                           and r["stores"] == 0 else "")
        recs.append({**r, "status": status or "", "usw": usw, "note": note})

    # Retailers we already sell to, ones ruled out (Non-Target) and ones that declined are
    # left out, except the declined ones in KEEP_DECLINED. Sort by status (furthest along
    # first, Declined last), then by wholesale $ potential.
    recs = [x for x in recs if x["status"] not in ("in", "non-target", "declined")
            or (x["status"] == "declined" and x["id"] in KEEP_DECLINED)]
    rank = lambda s: STATUS_ORDER.index(s) if s in STATUS_ORDER else len(STATUS_ORDER)
    recs.sort(key=lambda x: (rank(x["status"]), -x["stores"] * x["usw"], x["name"]))

    wb = Workbook()
    ws = wb.active
    ws.title = "Retailers"
    hdr = ["Retailer", "Channel", "Doors", "U/S/W", "Annual Units",
           "Wholesale $ Potential", "Status", "Commentary"]
    widths = [34, 17, 9, 8, 13, 15, 20, 58]
    ws.append(hdr)
    for i, rec in enumerate(recs, start=2):
        ws.append([rec["name"], channels[rec["channel"]]["name"], rec["stores"], rec["usw"],
                   f"=C{i}*D{i}*Assumptions!$B$3", f"=E{i}*Assumptions!$B$2",
                   STATUS_LABEL.get(rec["status"], rec["status"]), rec["note"]])
    last = len(recs) + 1

    thin = Side(style="thin", color="D0D7E2")
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    for row, rec in zip(ws.iter_rows(min_row=2, max_row=last), recs):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column == 8)
            cell.border = Border(bottom=thin)
        row[2].number_format = "#,##0"
        row[3].number_format = "0.00"
        row[4].number_format = "#,##0"
        row[5].number_format = "$#,##0"
        if rec["status"] in PARKED:
            for cell in row:
                cell.font = Font(color="8A8A8A")
    tbl = Table(displayName="Retailers", ref=f"A1:H{last}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tbl)
    ws.freeze_panes = "B2"

    # ── By Channel ──
    sc = wb.create_sheet("By Channel")
    order = [ch for ch in order if any(x["channel"] == ch for x in recs)]
    shown = ["pitched", "sent", "requested", "target", ""]
    sc.append(["Channel", "Retailers", "Doors", "Annual Units", "Wholesale $ Potential"]
              + [STATUS_LABEL[s] for s in shown] + ["Declined"])
    rng = lambda col: f"Retailers!${col}$2:${col}${last}"
    for i, ch in enumerate(order, start=2):
        nm = channels[ch]["name"]
        inc = f'{rng("B")},A{i},{rng("G")},"<>Declined"'
        st = lambda s: f'COUNTIFS({rng("B")},A{i},{rng("G")},"{s}")'
        sc.append([nm, f"=COUNTIFS({inc})", f"=SUMIFS({rng('C')},{inc})",
                   f"=SUMIFS({rng('E')},{inc})", f"=SUMIFS({rng('F')},{inc})"]
                  + ["=" + st(STATUS_LABEL[s]) for s in shown]
                  + ["=" + st("Declined")])
    tot = len(order) + 2
    ncol = 5 + len(shown) + 1
    sc.append(["TOTAL (working list)"] + [f"=SUM({get_column_letter(c)}2:{get_column_letter(c)}{tot-1})"
                                          for c in range(2, ncol + 1)])
    for c, w in enumerate([22, 10, 11, 14, 18, 9, 10, 11, 9, 10, 14], start=1):
        sc.column_dimensions[get_column_letter(c)].width = w
    for cell in sc[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0057E7")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in sc.iter_rows(min_row=2, max_row=tot):
        row[2].number_format = row[3].number_format = "#,##0"
        row[4].number_format = "$#,##0"
    for cell in sc[tot]:
        cell.font = Font(bold=True)
        cell.border = Border(top=Side(style="thin", color="1A1A2E"))
    sc.cell(tot + 2, 1, "Retailers, doors, units and $ count the working list only; Declined "
            "is excluded. Status columns count every row. Retailers Catalyst already sells "
            "to, and ones ruled out or declined (other than Target), are not included.").font = Font(italic=True, color="666666")
    sc.freeze_panes = "B2"

    # ── Assumptions ──
    sa = wb.create_sheet("Assumptions")
    sa.append(["Assumption", "Value", "Note"])
    sa.append(["Wholesale $ per unit", price, "Blended 15 lb / 34 lb"])
    sa.append(["Weeks per year", weeks, "Annual units = Doors × U/S/W × weeks"])
    sa.append([])
    sa.append(["Channel U/S/W", "", "Units per store per week assumed for each channel"])
    for ch in order:
        sa.append([channels[ch]["name"], channels[ch]["usw"]])
    sa.append([])
    sa.append(["As of", date.today().strftime("%m/%d/%y")])
    sa["B2"].number_format = "$#,##0.00"
    for c, w in enumerate([24, 12, 60], start=1):
        sa.column_dimensions[get_column_letter(c)].width = w
    for cell in sa[1] + sa[5]:
        cell.font = Font(bold=True)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)

    work = [x for x in recs if x["status"] not in PARKED]
    units = sum(x["stores"] * x["usw"] * weeks for x in work)
    print(f"  {len(recs)} retailers ({len(work)} working, {len(recs)-len(work)} parked), "
          f"{sum(x['stores'] for x in work):,} doors, {units:,.0f} units, ${units*price:,.0f} wholesale")
    print(f"  -> {OUT}")


if __name__ == "__main__":
    sys.exit(main())
