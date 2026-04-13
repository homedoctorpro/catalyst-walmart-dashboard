"""
email_report.py — Catalyst Pet Walmart Weekly Sales Report Email
Called automatically at the end of extract_data.py.
Usage standalone: python email_report.py --dry-run
"""

import os, sys, smtplib, math, argparse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── Credentials ───────────────────────────────────────────────────────────────
try:
    from credentials import EMAIL_USER, EMAIL_APP_PASSWORD
except ImportError:
    EMAIL_USER         = os.environ.get("EMAIL_USER", "ligneticsdata@gmail.com")
    EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")

EMAIL_FROM    = f"Catalyst Pet Reports <{EMAIL_USER}>"
EMAIL_TO      = ["pross@lignetics.com"]
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 587
DASHBOARD_URL = "https://homedoctorpro.github.io/catalyst-walmart-dashboard/"

# ── Static maps (kept in sync with dashboard_template.html) ───────────────────
WEEK_LABELS = {
    "202601": "Week 1 (2/6/26)",   "202602": "Week 2 (2/13/26)",
    "202603": "Week 3 (2/20/26)",  "202604": "Week 4 (2/27/26)",
    "202605": "Week 5 (3/6/26)",   "202606": "Week 6 (3/13/26)",
    "202607": "Week 7 (3/20/26)",  "202608": "Week 8 (3/27/26)",
    "202609": "Week 9 (4/3/26)",   "202610": "Week 10 (4/10/26)",
}

SKU_LABELS = {
    "CATALYST15ORIG":       "15lb Original",
    "CATALYST15UNSCEN":     "15lb Unscented",
    "CATALYST34LBORIGINAL": "34lb Original",
    "CATALYSTPET34LBUNSCE": "34lb Unscented",
}

GEOLIFT_GROUPS = {
    "Meta":    ["TENNESSEE","KENTUCKY","UTAH","OKLAHOMA","WASHINGTON","GEORGIA","ARKANSAS","LOUISIANA"],
    "YouTube": ["NEVADA","OREGON","MISSOURI","MICHIGAN","ARIZONA","ALABAMA","INDIANA","NEBRASKA"],
    "Control": ["VIRGINIA","OHIO","MINNESOTA","KANSAS","IOWA","NEW MEXICO","COLORADO","WISCONSIN"],
}
GEOLIFT_COLORS = {"Meta": "#0057e7", "YouTube": "#e63221", "Control": "#555555"}
GEOLIFT_BASELINE = "202609"

# ── Formatters ────────────────────────────────────────────────────────────────
def wl(week):
    return WEEK_LABELS.get(week, week)

def fmt_usd(v):
    if v is None: return "—"
    if v >= 1_000_000: return f"${v/1_000_000:.2f}M"
    if v >= 1_000:     return f"${v/1_000:.1f}K"
    return f"${v:,.0f}"

def fmt_num(v):
    if v is None: return "—"
    return f"{int(v):,}"

def fmt_pct(v):
    if v is None: return "—"
    return f"{v:.1f}%"

def fmt_usw(v):
    if v is None: return "—"
    return f"{v:.3f}"

def delta_str(curr, prev, higher_better=True, is_pct=False):
    """Return (text, color) for a week-over-week change."""
    if curr is None or prev is None or prev == 0:
        return "", "#888"
    chg = curr - prev
    pct = chg / abs(prev) * 100
    up  = chg > 0
    good = up if higher_better else not up
    color = "#00a86b" if good else "#d32f2f"
    arrow = "▲" if up else "▼"
    if is_pct:
        return f"{arrow} {abs(chg):.1f}pp", color
    return f"{arrow} {abs(pct):.1f}%", color

# ── GeoLift helpers ───────────────────────────────────────────────────────────
def compute_geolift(data):
    """Return {week: {group: {usw, index, stores, units}}} for weeks >= baseline."""
    store_weeks = data["store_weeks"]

    # Per-week, per-state store counts
    ws = {}
    for week in store_weeks:
        ws[week] = {}
        for sn, sdata in data["weekly_stores"].get(week, {}).items():
            info = data["stores"].get(sn, {})
            st = info.get("state", "")
            if not st: continue
            ws[week][st] = ws[week].get(st, 0) + 1

    def w_usw(states, week):
        q = sum(data["state_sales"].get(week, {}).get(st, 0) for st in states)
        s = sum(ws.get(week, {}).get(st, 0) for st in states)
        return q / s if s > 0 else None

    def uw_usw(states, week):
        vals = []
        for st in states:
            q = data["state_sales"].get(week, {}).get(st, 0)
            s = ws.get(week, {}).get(st, 0)
            if s > 0: vals.append(q / s)
        return sum(vals) / len(vals) if vals else None

    # Baseline
    baseline_w  = {g: w_usw(states,  GEOLIFT_BASELINE) for g, states in GEOLIFT_GROUPS.items()}
    baseline_uw = {g: uw_usw(states, GEOLIFT_BASELINE) for g, states in GEOLIFT_GROUPS.items()}

    gl_weeks = [w for w in store_weeks if w >= GEOLIFT_BASELINE]
    result = {}
    for week in gl_weeks:
        result[week] = {}
        for g, states in GEOLIFT_GROUPS.items():
            wv  = w_usw(states, week)
            uwv = uw_usw(states, week)
            bw  = baseline_w.get(g)
            buw = baseline_uw.get(g)
            result[week][g] = {
                "usw_w":   wv,
                "usw_uw":  uwv,
                "idx_w":   round(wv  / bw  * 100, 1) if (wv  and bw)  else None,
                "idx_uw":  round(uwv / buw * 100, 1) if (uwv and buw) else None,
                "stores":  sum(ws.get(week, {}).get(st, 0) for st in states),
                "units":   sum(data["state_sales"].get(week, {}).get(st, 0) for st in states),
            }
    return gl_weeks, result

# ── OOS helpers ───────────────────────────────────────────────────────────────
def compute_oos_summary(data):
    """Return total OOS stores and 2+ week OOS count for the latest week."""
    store_weeks = data["store_weeks"]
    if not store_weeks: return 0, 0, 0
    latest = store_weeks[-1]
    consec = data.get("consecutive_oos_by_week", {}).get(latest, {})

    oos_stores = set()
    multi_oos  = set()
    for sku, store_map in consec.items():
        for sn, cnt in store_map.items():
            oos_stores.add(sn)
            if cnt >= 2:
                multi_oos.add(sn)

    total_stores = len(data["weekly_stores"].get(latest, {}))
    return len(oos_stores), len(multi_oos), total_stores

# ── HTML builder ──────────────────────────────────────────────────────────────
def build_html(data):
    weeks       = data["weeks"]
    store_weeks = data["store_weeks"]
    metrics     = data["metrics"]
    skus        = data["skus"]

    cur_week  = store_weeks[-1] if store_weeks else weeks[-1]
    prev_week = store_weeks[-2] if len(store_weeks) >= 2 else None

    cur_m  = metrics.get(cur_week,  {})
    prev_m = metrics.get(prev_week, {}) if prev_week else {}
    total  = cur_m.get("Total",  {})
    p_tot  = prev_m.get("Total", {})

    # GeoLift
    gl_weeks, gl_data = compute_geolift(data)
    show_geolift = len(gl_weeks) >= 1

    # OOS
    oos_total, oos_multi, total_stores = compute_oos_summary(data)

    # ── Inline CSS ────────────────────────────────────────────────────────────
    css = """
    body,table,td,th{font-family:-apple-system,Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a2e;}
    body{background:#f0f2f5;margin:0;padding:0;}
    a{color:#0057e7;text-decoration:none;}
    .wrap{max-width:640px;margin:0 auto;background:#f0f2f5;padding:20px 0;}
    .header{background:linear-gradient(90deg,#1a1a2e 0%,#16213e 100%);border-radius:12px 12px 0 0;padding:28px 32px;}
    .header h1{color:#fff;font-size:20px;margin:0 0 4px;}
    .header p{color:rgba(255,255,255,0.55);font-size:13px;margin:0;}
    .body-card{background:#fff;padding:28px 32px;}
    .section-title{font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.07em;margin:0 0 12px;}
    .kpi-table{width:100%;border-collapse:collapse;margin-bottom:24px;}
    .kpi-table td{padding:0;width:16.6%;}
    .kpi-box{background:#f8f9ff;border-radius:10px;padding:14px 10px;text-align:center;margin:3px;}
    .kpi-label{font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px;}
    .kpi-value{font-size:20px;font-weight:800;color:#1a1a2e;line-height:1;}
    .kpi-delta{font-size:11px;font-weight:600;margin-top:4px;}
    .sku-table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:24px;}
    .sku-table th{background:#f5f5f5;padding:8px 10px;text-align:right;font-weight:700;color:#555;border-bottom:2px solid #e0e0e0;white-space:nowrap;}
    .sku-table th:first-child{text-align:left;}
    .sku-table td{padding:7px 10px;border-bottom:1px solid #f0f0f0;text-align:right;}
    .sku-table td:first-child{text-align:left;font-weight:600;}
    .sku-table tr.total-row td{font-weight:700;background:#f0f4ff;border-top:2px solid #d0d8f0;}
    .gl-table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px;}
    .gl-table th{background:#f5f5f5;padding:7px 10px;text-align:right;font-weight:700;color:#555;border-bottom:2px solid #e0e0e0;}
    .gl-table th:first-child,.gl-table th:nth-child(2){text-align:left;}
    .gl-table td{padding:6px 10px;border-bottom:1px solid #f0f0f0;text-align:right;}
    .gl-table td:first-child,.gl-table td:nth-child(2){text-align:left;}
    .badge{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;border-radius:10px;}
    .badge-meta{background:#e8f0fe;color:#0057e7;}
    .badge-youtube{background:#fce8e6;color:#c62828;}
    .badge-control{background:#f0f0f0;color:#444;}
    .idx-chip{font-weight:700;}
    .oos-row{display:flex;gap:16px;}
    .oos-box{flex:1;background:#fff3f3;border-radius:8px;padding:12px 16px;text-align:center;}
    .oos-num{font-size:24px;font-weight:800;color:#c62828;}
    .oos-lbl{font-size:11px;color:#888;margin-top:3px;}
    .footer{background:#f8f9ff;border-radius:0 0 12px 12px;padding:20px 32px;text-align:center;}
    .cta-btn{display:inline-block;background:#0057e7;color:#fff;font-weight:700;font-size:14px;
             padding:12px 28px;border-radius:8px;text-decoration:none;}
    .divider{border:none;border-top:1px solid #f0f0f0;margin:20px 0;}
    """

    # ── KPI row ───────────────────────────────────────────────────────────────
    def kpi_box(label, value_str, curr, prev, higher_better=True, is_pct=False):
        d_str, d_color = delta_str(curr, prev, higher_better, is_pct)
        delta_html = f'<div class="kpi-delta" style="color:{d_color}">{d_str}</div>' if d_str else '<div class="kpi-delta">&nbsp;</div>'
        return f"""
        <td><div class="kpi-box">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value_str}</div>
          {delta_html}
        </div></td>"""

    kpi_html = f"""
    <p class="section-title">Overall — {wl(cur_week)}</p>
    <table class="kpi-table"><tr>
      {kpi_box("U/S/W",        fmt_usw(total.get("usw")),           total.get("usw"),            p_tot.get("usw"))}
      {kpi_box("Adj. U/S/W",   fmt_usw(total.get("usw") / (total.get("instock_pct",100)/100) if total.get("usw") and total.get("instock_pct") else None),
                                total.get("usw") / (total.get("instock_pct",100)/100) if total.get("usw") and total.get("instock_pct") else None,
                                p_tot.get("usw")  / (p_tot.get("instock_pct",100)/100) if p_tot.get("usw")  and p_tot.get("instock_pct")  else None)}
      {kpi_box("Instock %",    fmt_pct(total.get("instock_pct")),   total.get("instock_pct"),    p_tot.get("instock_pct"),    is_pct=True)}
      {kpi_box("Units Sold",   fmt_num(total.get("pos_qty")),       total.get("pos_qty"),        p_tot.get("pos_qty"))}
      {kpi_box("Retail $",     fmt_usd(total.get("pos_dollars")),   total.get("pos_dollars"),    p_tot.get("pos_dollars"))}
      {kpi_box("Wholesale $",  fmt_usd(total.get("wholesale_dollars")), total.get("wholesale_dollars"), p_tot.get("wholesale_dollars"))}
    </tr></table>"""

    # ── SKU breakdown ─────────────────────────────────────────────────────────
    def sku_row(sku, label, m):
        d = m.get(sku, {})
        adj_usw = None
        if d.get("usw") and d.get("instock_pct"):
            adj_usw = d["usw"] / (d["instock_pct"] / 100)
        return f"""<tr>
          <td>{label}</td>
          <td>{fmt_usw(d.get("usw"))}</td>
          <td>{fmt_usw(adj_usw)}</td>
          <td>{fmt_pct(d.get("instock_pct"))}</td>
          <td>{fmt_num(d.get("pos_qty"))}</td>
          <td>{fmt_usd(d.get("pos_dollars"))}</td>
          <td>{fmt_usd(d.get("wholesale_dollars"))}</td>
        </tr>"""

    tot = cur_m.get("Total", {})
    tot_adj = None
    if tot.get("usw") and tot.get("instock_pct"):
        tot_adj = tot["usw"] / (tot["instock_pct"] / 100)

    sku_rows = "\n".join(sku_row(sku, SKU_LABELS.get(sku, sku), cur_m) for sku in skus)

    sku_html = f"""
    <hr class="divider">
    <p class="section-title">SKU Breakdown — {wl(cur_week)}</p>
    <table class="sku-table">
      <thead><tr>
        <th>SKU</th><th>U/S/W</th><th>Adj U/S/W</th><th>Instock %</th>
        <th>Units</th><th>Retail $</th><th>Wholesale $</th>
      </tr></thead>
      <tbody>
        {sku_rows}
        <tr class="total-row">
          <td>TOTAL</td>
          <td>{fmt_usw(tot.get("usw"))}</td>
          <td>{fmt_usw(tot_adj)}</td>
          <td>{fmt_pct(tot.get("instock_pct"))}</td>
          <td>{fmt_num(tot.get("pos_qty"))}</td>
          <td>{fmt_usd(tot.get("pos_dollars"))}</td>
          <td>{fmt_usd(tot.get("wholesale_dollars"))}</td>
        </tr>
      </tbody>
    </table>"""

    # ── OOS ───────────────────────────────────────────────────────────────────
    oos_pct     = f"{oos_total/total_stores*100:.1f}%" if total_stores else "—"
    multi_pct   = f"{oos_multi/total_stores*100:.1f}%" if total_stores else "—"

    oos_html = f"""
    <hr class="divider">
    <p class="section-title">Out-of-Stock Snapshot — {wl(cur_week)}</p>
    <table width="100%"><tr>
      <td width="50%" style="padding:4px;">
        <div style="background:#fff3f3;border-radius:8px;padding:14px 16px;text-align:center;">
          <div style="font-size:28px;font-weight:800;color:#c62828;">{oos_total}</div>
          <div style="font-size:11px;color:#888;margin-top:3px;">Stores OOS &nbsp;·&nbsp; {oos_pct} of {total_stores:,}</div>
        </div>
      </td>
      <td width="50%" style="padding:4px;">
        <div style="background:#fff8f0;border-radius:8px;padding:14px 16px;text-align:center;">
          <div style="font-size:28px;font-weight:800;color:#e67300;">{oos_multi}</div>
          <div style="font-size:11px;color:#888;margin-top:3px;">OOS 2+ Consecutive Weeks &nbsp;·&nbsp; {multi_pct}</div>
        </div>
      </td>
    </tr></table>"""

    # ── GeoLift ───────────────────────────────────────────────────────────────
    geolift_html = ""
    if show_geolift:
        latest_gl = gl_weeks[-1]
        is_baseline = latest_gl == GEOLIFT_BASELINE
        gl_note = "Wk 9 = baseline (100)" if is_baseline else f"indexed to {wl(GEOLIFT_BASELINE)} = 100"

        def idx_chip(v, group):
            if v is None: return "—"
            color = GEOLIFT_COLORS[group]
            bg    = "#e8f0fe" if group=="Meta" else "#fce8e6" if group=="YouTube" else "#f0f0f0"
            return f'<span style="background:{bg};color:{color};font-weight:700;padding:2px 7px;border-radius:8px;">{v:.1f}</span>'

        rows_html = ""
        for week in gl_weeks:
            for g, gd in gl_data[week].items():
                key = g.lower()
                badge = f'<span class="badge badge-{key}">{g}</span>'
                rows_html += f"""<tr>
                  <td>{wl(week)}</td>
                  <td>{badge}</td>
                  <td>{gd["stores"]}</td>
                  <td>{fmt_num(gd["units"])}</td>
                  <td>{fmt_usw(gd["usw_w"])}</td>
                  <td>{idx_chip(gd["idx_w"],  g)}</td>
                  <td>{fmt_usw(gd["usw_uw"])}</td>
                  <td>{idx_chip(gd["idx_uw"], g)}</td>
                </tr>"""

        geolift_html = f"""
    <hr class="divider">
    <p class="section-title">GeoLift Test — {gl_note}</p>
    <table class="gl-table">
      <thead><tr>
        <th>Week</th><th>Group</th><th>Stores</th><th>Units</th>
        <th>Wtd U/S/W</th><th>Wtd Idx</th>
        <th>Unwtd U/S/W</th><th>Unwtd Idx</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>"""

    # ── WoW comparison note ───────────────────────────────────────────────────
    wow_note = f"vs {wl(prev_week)}" if prev_week else "no prior week"

    # ── Assemble ──────────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Catalyst Pet Walmart Report — {wl(cur_week)}</title>
<style>{css}</style>
</head>
<body>
<div class="wrap">

  <!-- Header -->
  <div class="header">
    <h1>⚡ Catalyst Pet — Walmart Sales Report</h1>
    <p>{wl(cur_week)} &nbsp;·&nbsp; WoW deltas: {wow_note} &nbsp;·&nbsp; {len(store_weeks)} weeks of data</p>
  </div>

  <!-- Body -->
  <div class="body-card">
    {kpi_html}
    {sku_html}
    {oos_html}
    {geolift_html}
  </div>

  <!-- Footer -->
  <div class="footer">
    <a href="{DASHBOARD_URL}" class="cta-btn">Open Full Dashboard →</a>
    <p style="font-size:11px;color:#aaa;margin-top:14px;">
      Sent from ligneticsdata@gmail.com &nbsp;·&nbsp; Catalyst Pet internal use only
    </p>
  </div>

</div>
</body>
</html>"""

    return html


# ── Send ──────────────────────────────────────────────────────────────────────
def send_report(data, dry_run=False):
    html = build_html(data)

    cur_week = (data["store_weeks"] or data["weeks"])[-1]
    subject  = f"Catalyst Pet Walmart Report — {wl(cur_week)}"

    if dry_run:
        out = "email_preview.html"
        with open(out, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  [Email] Dry-run: preview written to {out}")
        return

    if not EMAIL_APP_PASSWORD or EMAIL_APP_PASSWORD == "PASTE_APP_PASSWORD_HERE":
        print("  [Email] Skipped — no app password configured in credentials.py")
        return

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = EMAIL_FROM
    msg["To"]      = ", ".join(EMAIL_TO)
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(EMAIL_USER, EMAIL_APP_PASSWORD)
            server.sendmail(EMAIL_USER, EMAIL_TO, msg.as_string())
        print(f"  [Email] Sent to {', '.join(EMAIL_TO)} — {subject}")
    except Exception as e:
        print(f"  [Email] FAILED: {e}")


# ── Standalone entry point ────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true",
                        help="Write email_preview.html instead of sending")
    args = parser.parse_args()

    # Re-run the full ETL to get data, then send
    sys.path.insert(0, os.path.dirname(__file__))
    import extract_data as etl

    print("Loading data for email...")
    files    = etl.find_excel_files()
    geo      = etl.load_geo_cache()

    import pandas as pd
    metrics, raw_by_week = {}, {}
    ecomm_l52_raw, ecomm_lw_raw = {}, {}

    for week in sorted(files.keys()):
        si = etl.SHEET_MAP.get(week)
        if not si: continue
        fp = files[week]

        try:
            df = pd.read_excel(fp, sheet_name=si["instore"], header=None)
            metrics[week] = etl.extract_instore_metrics(week, df)
        except Exception as e:
            print(f"  [WARN] InStore {week}: {e}")

        if si["bystore"]:
            try:
                df = pd.read_excel(fp, sheet_name=si["bystore"], header=None)
                raw_by_week[week] = etl.extract_store_data(week, df)
            except Exception as e:
                print(f"  [WARN] ByStore {week}: {e}")

    geo = etl.geocode_stores({sn: {"zip5": r["zip5"]} for wk, rows in raw_by_week.items()
                               for sn, r in [(r["store_num"], r) for r in rows]}, geo)

    all_sw     = etl.build_weekly_store_summary(raw_by_week)
    stores     = etl.build_stores_dict(raw_by_week, geo)
    state_sales = etl.compute_state_sales(all_sw, stores)
    consec_oos  = etl.compute_consecutive_oos_by_week(all_sw, sorted(raw_by_week.keys()))

    data = {
        "weeks":        sorted(files.keys()),
        "store_weeks":  sorted(raw_by_week.keys()),
        "skus":         etl.SKUS,
        "metrics":      metrics,
        "stores":       stores,
        "weekly_stores": all_sw,
        "state_sales":  state_sales,
        "consecutive_oos_by_week": consec_oos,
    }

    send_report(data, dry_run=args.dry_run)
