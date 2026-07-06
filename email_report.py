"""
email_report.py — Catalyst Pet Walmart Weekly Sales Report Email
Called automatically at the end of extract_data.py.
Usage standalone: python email_report.py --dry-run
"""

import os, sys, smtplib, math, argparse, io, base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── Credentials ───────────────────────────────────────────────────────────────
try:
    from credentials import EMAIL_USER, EMAIL_APP_PASSWORD
except ImportError:
    EMAIL_USER         = os.environ.get("EMAIL_USER", "ligneticsdata@gmail.com")
    EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")

EMAIL_FROM    = f"Catalyst Pet Reports <{EMAIL_USER}>"
EMAIL_TO      = [
    "pross@lignetics.com",
    "jlevine@lignetics.com",
    "hailey@storandco.com",
    "ckohagen@lignetics.com",
    "mscanlon@lignetics.com",
    "ckaminski@lignetics.com",
    "JGallman@Lignetics.com",
    "asnider@lignetics.com",
    "SBendezu@lignetics.com",
    "mweaver@lignetics.com",
    "glenda.vale@quartile.com",
    "Amanda.Lourenco@quartile.com",
    "bjordan@lignetics.com",
    "jlatham@storandco.com",
    "jdingus@lignetics.com",
    "rdiefert@lignetics.com",
    "card@lignetics.com",
    "bking@lignetics.com",
]
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 465   # SSL; falls back to 587/STARTTLS if needed
DASHBOARD_URL = "https://homedoctorpro.github.io/catalyst-walmart-dashboard/dashboard.html"

# Populated from data["week_labels"] at the start of send_report() —
# single source of truth lives in extract_data.compute_week_label().
WEEK_LABELS = {}

SKU_LABELS = {
    "CATALYST15ORIG":       "15lb Original",
    "CATALYST15UNSCEN":     "15lb Unscented",
    "CATALYST34LBORIGINAL": "34lb Original",
    "CATALYSTPET34LBUNSCE": "34lb Unscented",
}

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
    return f"{v:.2f}"

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

def count_oos_stores(data, week):
    """Distinct stores OOS (any SKU) in a given week."""
    consec = data.get("consecutive_oos_by_week", {}).get(week, {})
    s = set()
    for store_map in consec.values():
        s.update(store_map.keys())
    return len(s)

# ── Highlights / auto-commentary ──────────────────────────────────────────────
def build_highlights(data, cur_week, prev_week, metrics, skus, total, p_tot,
                     oos_total, oos_multi, total_stores):
    """Return HTML <li> bullets summarizing the week's most important movements."""
    bullets = []  # (emoji, html_text)

    def signed_pct(curr, prev):
        if curr is None or prev is None or prev == 0: return None
        return (curr - prev) / abs(prev) * 100

    def col(txt, c): return f'<span style="color:{c};font-weight:700;">{txt}</span>'

    # 1. Sales momentum (U/S/W)
    usw_c, usw_p = total.get("usw"), p_tot.get("usw")
    p = signed_pct(usw_c, usw_p)
    if usw_c is not None:
        if p is None:
            bullets.append(("📈", f"Total <b>U/S/W is {fmt_usw(usw_c)}</b> this week."))
        else:
            c = "#00a86b" if p >= 0 else "#d32f2f"
            word = "up" if p >= 0 else "down"
            bullets.append(("📈", f"Total <b>U/S/W {fmt_usw(usw_c)}</b> — {col(('▲' if p>=0 else '▼')+f' {abs(p):.1f}% WoW', c)} ({word} from {fmt_usw(usw_p)})."))

    # 2. Demand vs availability (adjusted U/S/W)
    ins = total.get("instock_pct")
    if usw_c and ins:
        adj = usw_c / (ins / 100)
        gap = adj - usw_c
        if ins < 95 and gap > 0:
            bullets.append(("🔓", f"At {fmt_pct(ins)} instock, demand-adjusted U/S/W is <b>{fmt_usw(adj)}</b> — roughly {col(f'+{gap/usw_c*100:.0f}% upside', '#0057e7')} if fully in stock."))

    # 3. Best / worst SKU WoW
    sku_moves = []
    for sku in skus:
        c = metrics.get(cur_week, {}).get(sku, {}).get("usw")
        pv = metrics.get(prev_week, {}).get(sku, {}).get("usw") if prev_week else None
        mp = signed_pct(c, pv)
        if mp is not None:
            sku_moves.append((sku, mp))
    if sku_moves:
        sku_moves.sort(key=lambda x: x[1])
        worst, best = sku_moves[0], sku_moves[-1]
        if best[1] > 0:
            bullets.append(("🏆", f"Best mover: <b>{SKU_LABELS.get(best[0], best[0])}</b> {col(f'▲ {best[1]:.1f}%', '#00a86b')} WoW."))
        if worst[1] < 0 and worst[0] != best[0]:
            bullets.append(("⚠️", f"Lagging: <b>{SKU_LABELS.get(worst[0], worst[0])}</b> {col(f'▼ {abs(worst[1]):.1f}%', '#d32f2f')} WoW."))

    # 4. OOS trend
    if total_stores:
        oos_pct = oos_total / total_stores * 100
        prev_oos = count_oos_stores(data, prev_week) if prev_week else None
        trend = ""
        if prev_oos is not None and prev_oos != oos_total:
            dlt = oos_total - prev_oos
            tc  = "#d32f2f" if dlt > 0 else "#00a86b"
            trend = f" — {col(('▲' if dlt>0 else '▼')+f' {abs(dlt)} vs last week', tc)}"
        bullets.append(("📦", f"<b>{oos_total} stores out of stock</b> ({oos_pct:.1f}% of {total_stores:,}){trend}."))
        if oos_multi:
            bullets.append(("🚚", f"{col(str(oos_multi)+' stores', '#e67300')} have been OOS <b>2+ consecutive weeks</b> — replenishment priority."))

    if not bullets:
        return ""

    lis = "\n".join(
        f'<li style="margin-bottom:6px;"><span style="margin-right:6px;">{e}</span>{t}</li>'
        for e, t in bullets
    )
    return f"""
    <div style="background:#f4f7ff;border-left:4px solid #0057e7;border-radius:8px;padding:16px 20px 14px;margin-bottom:22px;">
      <p style="font-size:11px;font-weight:700;color:#0057e7;text-transform:uppercase;letter-spacing:.07em;margin:0 0 10px;">📌 Key Takeaways</p>
      <ul style="margin:0;padding-left:20px;line-height:1.55;font-size:13px;color:#1a1a2e;">{lis}</ul>
    </div>"""

# ── Chart generators (matplotlib → base64 PNG) ───────────────────────────────
def _mpl():
    """Return (plt, io) or (None, None) if matplotlib unavailable."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        return plt
    except ImportError:
        return None

SKU_COLORS_MPL = {
    "CATALYST15ORIG":       "#0057e7",
    "CATALYST15UNSCEN":     "#00a86b",
    "CATALYST34LBORIGINAL": "#e67300",
    "CATALYSTPET34LBUNSCE": "#9c27b0",
}
SKU_SHORT_MPL = {
    "CATALYST15ORIG": "15lb Orig", "CATALYST15UNSCEN": "15lb Unsc",
    "CATALYST34LBORIGINAL": "34lb Orig", "CATALYSTPET34LBUNSCE": "34lb Unsc",
}

def _fig_to_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight")
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()

def make_usw_chart(metrics, all_weeks, skus):
    """Overall U/S/W trend from inception — all SKUs + Total bold."""
    plt = _mpl()
    if not plt: return None

    weeks_with_data = [w for w in all_weeks if metrics.get(w, {}).get("Total", {}).get("usw") is not None]
    if not weeks_with_data: return None
    labels = [wl(w) for w in weeks_with_data]

    fig, ax = plt.subplots(figsize=(7, 3.2))
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#f8f9ff")

    for sku in skus:
        vals = [metrics.get(w, {}).get(sku, {}).get("usw") for w in weeks_with_data]
        if any(v is not None for v in vals):
            ax.plot(labels, vals, color=SKU_COLORS_MPL[sku], linewidth=1.5,
                    marker="o", markersize=4, label=SKU_SHORT_MPL.get(sku, sku), alpha=0.75)

    total_vals = [metrics.get(w, {}).get("Total", {}).get("usw") for w in weeks_with_data]
    ax.plot(labels, total_vals, color="#1a1a2e", linewidth=2.8,
            marker="o", markersize=6, label="Total", zorder=5)

    ax.set_ylabel("U/S/W", fontsize=9, color="#555")
    ax.tick_params(axis="x", rotation=35, labelsize=7.5)
    ax.tick_params(axis="y", labelsize=8)
    for spine in ["top", "right"]: ax.spines[spine].set_visible(False)
    for spine in ["left", "bottom"]: ax.spines[spine].set_color("#ddd")
    ax.grid(axis="y", alpha=0.3, color="#ccc")
    ax.legend(loc="upper left", fontsize=7.5, framealpha=0.9, ncol=3)
    plt.tight_layout(pad=1.0)

    b64 = _fig_to_b64(fig)
    plt.close(fig)
    return b64

def make_oos_map(data):
    """Scatter map colored by consecutive weeks OOS (matches dashboard scale)."""
    plt = _mpl()
    if not plt: return None

    store_weeks = data["store_weeks"]
    if not store_weeks: return None
    latest  = store_weeks[-1]
    wk_data = data["weekly_stores"].get(latest, {})
    consec  = data.get("consecutive_oos_by_week", {}).get(latest, {})

    # Max consecutive OOS weeks per store across all SKUs
    store_max_consec = {}
    for sku_map in consec.values():
        for sn, cnt in sku_map.items():
            store_max_consec[sn] = max(store_max_consec.get(sn, 0), cnt)

    # Color buckets matching the dashboard feel
    OOS_BUCKETS = [
        (1,  "#FFB300", "1 week"),
        (2,  "#F57C00", "2 weeks"),
        (3,  "#E53935", "3 weeks"),
        (4,  "#C62828", "4 weeks"),
        (5,  "#7B1FA2", "5+ weeks"),
    ]
    def oos_color(cnt):
        if cnt >= 5: return "#7B1FA2"
        if cnt == 4: return "#C62828"
        if cnt == 3: return "#E53935"
        if cnt == 2: return "#F57C00"
        return "#FFB300"

    # Separate in-stock vs OOS stores
    in_lons, in_lats = [], []
    buckets = {b[0]: ([], []) for b in OOS_BUCKETS}  # week_count -> (lons, lats)

    for sn in wk_data:
        info = data["stores"].get(sn, {})
        lat, lon = info.get("lat"), info.get("lon")
        if lat is None or lon is None: continue
        cnt = store_max_consec.get(sn, 0)
        if cnt == 0:
            in_lons.append(lon); in_lats.append(lat)
        else:
            bucket_key = min(cnt, 5)
            buckets[bucket_key][0].append(lon)
            buckets[bucket_key][1].append(lat)

    fig, ax = plt.subplots(figsize=(7, 4.2))
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#e8f0fe")

    # In-stock (small, low opacity)
    ax.scatter(in_lons, in_lats, s=2.5, color="#0057e7", alpha=0.25,
               linewidths=0, label=f"In Stock ({len(in_lats):,})", zorder=2)

    # OOS by bucket — draw lowest severity first so worst shows on top
    for cnt, color, label in OOS_BUCKETS:
        lons, lats = buckets[cnt]
        if lons:
            total = len(lons)
            ax.scatter(lons, lats, s=10, color=color, alpha=0.85,
                       linewidths=0, label=f"{label} OOS ({total:,})", zorder=3 + cnt)

    ax.set_xlim(-125, -65)
    ax.set_ylim(24, 50)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.legend(loc="lower left", fontsize=7.5, framealpha=0.95,
              markerscale=2, edgecolor="#ddd", ncol=2)
    ax.set_title(f"Store OOS Map — {wl(latest)}", fontsize=10, color="#1a1a2e", pad=8)

    plt.tight_layout(pad=0.5)
    b64 = _fig_to_b64(fig)
    plt.close(fig)
    return b64

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

    # Charts
    usw_chart_b64       = make_usw_chart(metrics, sorted(metrics.keys()), skus)
    oos_map_b64         = make_oos_map(data)

    def img_tag(b64, alt=""):
        if not b64: return ""
        return (f'<img src="data:image/png;base64,{b64}" '
                f'alt="{alt}" width="560" style="display:block;max-width:100%;'
                f'border-radius:8px;margin:0 auto;" />')

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
    def kpi_box(label, value_str, curr, prev, higher_better=True, is_pct=False, sub=None):
        d_str, d_color = delta_str(curr, prev, higher_better, is_pct)
        delta_html = f'<div class="kpi-delta" style="color:{d_color}">{d_str}</div>' if d_str else '<div class="kpi-delta">&nbsp;</div>'
        sub_html = f'<div style="font-size:9px;color:#999;margin-top:3px;">{sub}</div>' if sub else ''
        return f"""
        <td><div class="kpi-box">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value_str}</div>
          {delta_html}
          {sub_html}
        </div></td>"""

    usw_cur  = total.get("usw")
    usw_prev = p_tot.get("usw")
    adj_usw_cur  = usw_cur  / (total.get("instock_pct", 100) / 100) if usw_cur  and total.get("instock_pct")  else None
    adj_usw_prev = usw_prev / (p_tot.get("instock_pct",  100) / 100) if usw_prev and p_tot.get("instock_pct")  else None

    usw_chart_html = (f'<div style="margin:14px 0 4px;">{img_tag(usw_chart_b64, "U/S/W Trend")}</div>'
                      if usw_chart_b64 else "")

    kpi_html = f"""
    <p class="section-title">Overall — {wl(cur_week)}</p>
    <table class="kpi-table"><tr>
      {kpi_box("U/S/W",        fmt_usw(usw_cur),             usw_cur,             usw_prev)}
      {kpi_box("Adj. U/S/W",   fmt_usw(adj_usw_cur),         adj_usw_cur,         adj_usw_prev)}
      {kpi_box("Instock %",    fmt_pct(total.get("instock_pct")),   total.get("instock_pct"),    p_tot.get("instock_pct"),    is_pct=True)}
      {kpi_box("Units Sold",   fmt_num(total.get("pos_qty")),       total.get("pos_qty"),        p_tot.get("pos_qty"))}
      {kpi_box("Retail $",     fmt_usd(total.get("pos_dollars")),   total.get("pos_dollars"),    p_tot.get("pos_dollars"))}
      {kpi_box("Wholesale $",  fmt_usd(total.get("wholesale_dollars")), total.get("wholesale_dollars"), p_tot.get("wholesale_dollars"), sub=(f"≈ {fmt_usd(total.get('wholesale_dollars') * 52 / 12)}/mo run-rate" if total.get("wholesale_dollars") else None))}
    </tr></table>
    {usw_chart_html}"""

    # ── Highlights / auto-commentary ────────────────────────────────────────────
    highlights_html = build_highlights(
        data, cur_week, prev_week, metrics, skus, total, p_tot,
        oos_total, oos_multi, total_stores)

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

    # U/S/W trend — last 3 weeks per SKU
    all_metric_weeks_usw = sorted(w for w in metrics if any(
        metrics[w].get(sku, {}).get("usw") is not None for sku in skus + ["Total"]
    ))
    usw_trend_weeks = all_metric_weeks_usw[-3:]
    usw_trend_headers = "".join(f"<th>{wl(w)}</th>" for w in usw_trend_weeks) + "<th>WoW</th>"
    usw_trend_rows = ""
    for sku in skus + ["Total"]:
        label = "TOTAL" if sku == "Total" else SKU_LABELS.get(sku, sku)
        vals  = [metrics.get(w, {}).get(sku, {}).get("usw") for w in usw_trend_weeks]
        cur_v, prev_v = vals[-1], vals[-2] if len(vals) >= 2 else None
        d_str, d_color = delta_str(cur_v, prev_v)
        delta_td = f'<span style="color:{d_color};font-weight:600;white-space:nowrap;">{d_str}</span>' if d_str else "—"
        cells = "".join(f"<td>{fmt_usw(v)}</td>" for v in vals)
        row_style = ' class="total-row"' if sku == "Total" else ""
        usw_trend_rows += f"<tr{row_style}><td>{label}</td>{cells}<td>{delta_td}</td></tr>"

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
    </table>
    <p style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.07em;margin:18px 0 8px;">U/S/W Trend by SKU</p>
    <table class="sku-table">
      <thead><tr><th>SKU</th>{usw_trend_headers}</tr></thead>
      <tbody>{usw_trend_rows}</tbody>
    </table>"""

    # ── OOS ───────────────────────────────────────────────────────────────────
    oos_pct     = f"{oos_total/total_stores*100:.1f}%" if total_stores else "—"
    multi_pct   = f"{oos_multi/total_stores*100:.1f}%" if total_stores else "—"
    oos_map_html = (f'<div style="margin:12px 0 0;">{img_tag(oos_map_b64, "OOS Store Map")}</div>'
                    if oos_map_b64 else "")

    # Per-SKU OOS table
    wk_data_oos    = data["weekly_stores"].get(cur_week, {})
    consec_latest  = data.get("consecutive_oos_by_week", {}).get(cur_week, {})
    sku_oos_rows   = ""
    grand_total_pl = grand_oos_pl = 0
    for sku in skus:
        total_pl = sum(1 for sd in wk_data_oos.values() if sku in sd.get("skus", {}))
        oos_pl   = len(consec_latest.get(sku, {}))
        grand_total_pl += total_pl
        grand_oos_pl   += oos_pl
        oos_pct_sku = f"{oos_pl/total_pl*100:.1f}%" if total_pl else "—"
        sku_oos_rows += f"""<tr>
          <td>{SKU_LABELS.get(sku, sku)}</td>
          <td>{total_pl:,}</td>
          <td>{oos_pl:,}</td>
          <td>{oos_pct_sku}</td>
        </tr>"""
    grand_oos_pct = f"{grand_oos_pl/grand_total_pl*100:.1f}%" if grand_total_pl else "—"
    sku_oos_rows += f"""<tr class="total-row">
      <td>TOTAL</td>
      <td>{grand_total_pl:,}</td>
      <td>{grand_oos_pl:,}</td>
      <td>{grand_oos_pct}</td>
    </tr>"""

    oos_sku_table = f"""
    <table class="sku-table" style="margin-top:14px;">
      <thead><tr>
        <th style="text-align:left;">SKU</th>
        <th>Total Placements</th>
        <th>OOS Placements</th>
        <th>OOS %</th>
      </tr></thead>
      <tbody>{sku_oos_rows}</tbody>
    </table>"""

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
    </tr></table>
    {oos_sku_table}
    {oos_map_html}"""

    # ── Markdowns trend ───────────────────────────────────────────────────────
    all_metric_weeks = sorted(metrics.keys())
    trend_weeks = all_metric_weeks[-3:]  # last 3 weeks with instore data

    def md_cell(week, sku):
        v = metrics.get(week, {}).get(sku, {}).get("markdown_pct")
        return fmt_pct(v) if v is not None else "—"

    md_rows = ""
    for sku in skus:
        vals = [metrics.get(w, {}).get(sku, {}).get("markdown_pct") for w in trend_weeks]
        cur_v  = vals[-1]
        prev_v = vals[-2] if len(vals) >= 2 else None
        d_str, d_color = delta_str(cur_v, prev_v, higher_better=False, is_pct=True)
        delta_td = f'<span style="color:{d_color};font-weight:600;">{d_str}</span>' if d_str else "—"
        cells = "".join(f"<td>{md_cell(w, sku)}</td>" for w in trend_weeks)
        md_rows += f"<tr><td>{SKU_LABELS.get(sku, sku)}</td>{cells}<td>{delta_td}</td></tr>"

    # Total row
    tot_vals = [metrics.get(w, {}).get("Total", {}).get("markdown_pct") for w in trend_weeks]
    tot_d_str, tot_d_color = delta_str(tot_vals[-1], tot_vals[-2] if len(tot_vals)>=2 else None, higher_better=False, is_pct=True)
    tot_delta_td = f'<span style="color:{tot_d_color};font-weight:600;">{tot_d_str}</span>' if tot_d_str else "—"
    tot_cells = "".join(f"<td>{md_cell(w, 'Total')}</td>" for w in trend_weeks)
    md_rows += f'<tr class="total-row"><td>TOTAL</td>{tot_cells}<td>{tot_delta_td}</td></tr>'

    trend_headers = "".join(f"<th>{wl(w)}</th>" for w in trend_weeks)
    markdowns_html = f"""
    <hr class="divider">
    <p class="section-title">Markdown % Trend — Last {len(trend_weeks)} Weeks</p>
    <table class="sku-table">
      <thead><tr><th>SKU</th>{trend_headers}<th>WoW</th></tr></thead>
      <tbody>{md_rows}</tbody>
    </table>
    <p style="font-size:11px;color:#aaa;margin:6px 0 0;">Lower markdown % = less discounting. ▼ = improvement.</p>"""

    # ── Catalyst ecomm trend ──────────────────────────────────────────────────
    ecomm_weekly = data.get("ecomm_weekly", {})
    catalyst_ecomm_html = ""
    if ecomm_weekly:
        # Find weeks that have at least one Catalyst product
        cat_weeks = sorted(
            w for w, prods in ecomm_weekly.items()
            if any(v.get("brand") == "Catalyst" for v in prods.values())
        )
        if cat_weeks:
            latest_ew  = cat_weeks[-1]
            prev_ew    = cat_weeks[-2] if len(cat_weeks) >= 2 else None
            cur_prods  = {k: v for k, v in ecomm_weekly[latest_ew].items() if v.get("brand") == "Catalyst"}
            prev_prods = ({k: v for k, v in ecomm_weekly[prev_ew].items()  if v.get("brand") == "Catalyst"}
                          if prev_ew else {})

            span = max((v.get("span", 1) for v in cur_prods.values()), default=1)
            span_note = f"1-week delta" if span == 1 else f"{span}-week cumulative delta"

            # Sort consistently
            cat_order = ["Catalyst 15lb Orig","Catalyst 15lb Unsc","Catalyst 34lb Orig","Catalyst 34lb Unsc"]
            all_cat_prods = [p for p in cat_order if p in cur_prods] + \
                            [p for p in cur_prods if p not in cat_order]

            ecomm_rows = ""
            tot_u, tot_r = 0, 0.0
            prev_tot_u, prev_tot_r = 0, 0.0
            for prod in all_cat_prods:
                cd  = cur_prods.get(prod, {})
                pd_ = prev_prods.get(prod, {})
                u, r = cd.get("u"), cd.get("r")
                pu, pr = pd_.get("u"), pd_.get("r")
                if u: tot_u += u
                if r: tot_r += r
                if pu: prev_tot_u += pu
                if pr: prev_tot_r += pr
                u_d, u_c = delta_str(u, pu)
                r_d, r_c = delta_str(r, pr)
                u_delta = f'<span style="color:{u_c};font-weight:600;">{u_d}</span>' if u_d else "—"
                r_delta = f'<span style="color:{r_c};font-weight:600;">{r_d}</span>' if r_d else "—"
                ecomm_rows += f"""<tr>
                  <td>{prod}</td>
                  <td>{fmt_num(u)}</td><td>{u_delta}</td>
                  <td>{fmt_usd(r)}</td><td>{r_delta}</td>
                </tr>"""

            # Totals
            tot_u_d, tot_u_c = delta_str(tot_u or None, prev_tot_u or None)
            tot_r_d, tot_r_c = delta_str(tot_r or None, prev_tot_r or None)
            tot_u_delta = f'<span style="color:{tot_u_c};font-weight:600;">{tot_u_d}</span>' if tot_u_d else "—"
            tot_r_delta = f'<span style="color:{tot_r_c};font-weight:600;">{tot_r_d}</span>' if tot_r_d else "—"
            ecomm_rows += f"""<tr class="total-row">
              <td>TOTAL</td>
              <td>{fmt_num(tot_u)}</td><td>{tot_u_delta}</td>
              <td>{fmt_usd(tot_r)}</td><td>{tot_r_delta}</td>
            </tr>"""

            prev_label = f" vs {wl(prev_ew)}" if prev_ew else ""
            catalyst_ecomm_html = f"""
    <hr class="divider">
    <p class="section-title">Catalyst Ecomm — {wl(latest_ew)}{prev_label}</p>
    <table class="sku-table">
      <thead><tr>
        <th>Product</th>
        <th>Units</th><th>WoW Units</th>
        <th>Retail $</th><th>WoW $</th>
      </tr></thead>
      <tbody>{ecomm_rows}</tbody>
    </table>
    <p style="font-size:11px;color:#aaa;margin:6px 0 0;">{span_note} · Walmart.com only</p>
    <p style="font-size:11px;color:#e67300;margin:4px 0 0;">⚠️ These figures cover <strong>Ship from FC only</strong> and do not account for ship-from-store or pickup-from-store orders, which represent the majority of actual Walmart ecomm sales.</p>"""

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
  <div style="background-color:#1a1a2e;border-radius:12px 12px 0 0;padding:28px 32px;">
    <h1 style="color:#ffffff;font-size:20px;margin:0 0 4px;font-family:Arial,Helvetica,sans-serif;">&#9889; Catalyst Pet &#8212; Walmart Sales Report</h1>
    <p style="color:#aabbcc;font-size:13px;margin:0;font-family:Arial,Helvetica,sans-serif;">{wl(cur_week)} &nbsp;&middot;&nbsp; WoW deltas: {wow_note} &nbsp;&middot;&nbsp; {len(store_weeks)} weeks of data</p>
  </div>

  <!-- Body -->
  <div class="body-card">
    {highlights_html}
    {kpi_html}
    {sku_html}
    {oos_html}
    {markdowns_html}
    {catalyst_ecomm_html}
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
DEV_RECIPIENTS = ["pross@lignetics.com"]

def send_report(data, dry_run=False, dev_only=False):
    WEEK_LABELS.clear()
    WEEK_LABELS.update(data.get("week_labels", {}))
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

    recipients = DEV_RECIPIENTS if dev_only else EMAIL_TO
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = EMAIL_FROM
    msg["To"]      = ", ".join(recipients)
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(EMAIL_USER, EMAIL_APP_PASSWORD)
            server.sendmail(EMAIL_USER, recipients, msg.as_string())
        print(f"  [Email] Sent to {', '.join(recipients)} — {subject}")
    except Exception as e:
        # Fall back to STARTTLS on port 587
        try:
            with smtplib.SMTP(SMTP_HOST, 587) as server:
                server.starttls()
                server.login(EMAIL_USER, EMAIL_APP_PASSWORD)
                server.sendmail(EMAIL_USER, recipients, msg.as_string())
            print(f"  [Email] Sent (587/TLS) to {', '.join(recipients)} — {subject}")
        except Exception as e2:
            print(f"  [Email] FAILED (465: {e}) (587: {e2})")


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

        if si.get("ecomm_l52"):
            try:
                df = pd.read_excel(fp, sheet_name=si["ecomm_l52"], header=None)
                ecomm_l52_raw[week] = etl.extract_ecomm_data(df)
            except Exception: pass

        if si.get("ecomm_lw"):
            try:
                df = pd.read_excel(fp, sheet_name=si["ecomm_lw"], header=None)
                ecomm_lw_raw[week] = etl.extract_ecomm_data(df)
            except Exception: pass

    geo = etl.geocode_stores({sn: {"zip5": r["zip5"]} for wk, rows in raw_by_week.items()
                               for sn, r in [(r["store_num"], r) for r in rows]}, geo)

    all_sw     = etl.build_weekly_store_summary(raw_by_week)
    stores     = etl.build_stores_dict(raw_by_week, geo)
    state_sales = etl.compute_state_sales(all_sw, stores)
    consec_oos  = etl.compute_consecutive_oos_by_week(all_sw, sorted(raw_by_week.keys()))

    ecomm_weekly = etl.compute_ecomm_weekly(ecomm_l52_raw, ecomm_lw_raw)

    data = {
        "weeks":        sorted(files.keys()),
        "store_weeks":  sorted(raw_by_week.keys()),
        "skus":         etl.SKUS,
        "metrics":      metrics,
        "stores":       stores,
        "weekly_stores": all_sw,
        "state_sales":  state_sales,
        "consecutive_oos_by_week": consec_oos,
        "ecomm_weekly": ecomm_weekly,
    }

    send_report(data, dry_run=args.dry_run)
