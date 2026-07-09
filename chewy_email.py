"""
Chewy monthly report email
==========================
Builds a summary email from the generated chewy_dashboard.html payload and sends
it to the Chewy distro via Gmail SMTP. Designed to run in the chewy-report
GitHub Action after chewy_extract.py rebuilds the dashboard.

Guard: only sends when a NEW latest month appears (tracked in a state file), so
re-runs / partial snapshot arrivals don't re-send. Override with --force.

Env: EMAIL_USER, EMAIL_APP_PASSWORD (Gmail app password — same as the Walmart
pipeline). Modes: --dry-run (write preview, no send), --dev-only (send to Phil
only), default (full distro).

Usage:
    python chewy_email.py --dry-run
    python chewy_email.py --dev-only
    python chewy_email.py                 # full distro, once per new month
"""
import os
import re
import io
import ssl
import json
import base64
import smtplib
import argparse
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
DASHBOARD = os.path.join(HERE, "chewy_dashboard.html")
DASHBOARD_URL = "https://homedoctorpro.github.io/catalyst-walmart-dashboard/chewy_dashboard.html"
STATE_FILE = os.path.join(HERE, ".chewy_last_emailed_month")
PREVIEW = os.path.join(HERE, "chewy_email_preview.html")

EMAIL_TO = [
    "jeff@alignconsulting.com",
    "pross@lignetics.com",
    "jgallman@lignetics.com",
    "ckaminski@lignetics.com",
    "jschmidt@lignetics.com",
]
DEV_RECIPIENTS = ["pross@lignetics.com"]
def _cred(name, default=None):
    v = os.environ.get(name)
    if v:
        return v
    try:
        import credentials
        return getattr(credentials, name, default)
    except Exception:
        return default


SENDER = _cred("EMAIL_USER", "ligneticsdata@gmail.com")
SMTP_HOST, SMTP_PORT = "smtp.gmail.com", 465

CAT, FF = "#12805c", "#d98324"
_MONTHS = ["", "January", "February", "March", "April", "May", "June", "July",
           "August", "September", "October", "November", "December"]


def load_payload():
    html = open(DASHBOARD, encoding="utf-8").read()
    m = re.search(r"const DATA = (\{.*?\});\nconst PW", html, re.S)
    return json.loads(m.group(1))


def money(v):
    if v is None:
        return "—"
    if abs(v) >= 1e6:
        return f"${v/1e6:.2f}M"
    if abs(v) >= 1e3:
        return f"${v/1e3:.1f}K"
    return f"${v:,.0f}"


def num(v):
    return "—" if v is None else f"{round(v):,}"


def pct(v):
    return "—" if v is None else f"{v:+.0f}%"


def mlabel(m):
    y, mo = m.split("-")
    return f"{_MONTHS[int(mo)]} {y}"


METRICS = [("units", "Units"), ("wholesale", "Wholesale $"),
           ("retail", "Retail $")]


def _val(d, parts, metric, month):
    """Sum of a metric over parts for a month, or None if none is known
    (so a genuinely-unknown value renders blank, not $0)."""
    tot, seen = 0, False
    for p in parts:
        if metric == "wholesale":
            src = d["wholesale"].get(p, {})
        elif metric == "implied":
            src = d["implied"].get(p, {})
        else:
            src = d["series"][p][metric]
        if month in src:
            tot += src[month]
            seen = True
    return tot if seen else None


def _growth(a, b):
    if a is None or not b:
        return None
    return (a - b) / b * 100


def brand_parts(d, brand):
    return [p for p, v in d["products"].items() if v["brand"] == brand]


def make_units_png(d):
    """Last-12-month unit-sales line chart -> PNG bytes."""
    months = d["months"][-12:]
    xs = range(len(months))
    total = [_val(d, list(d["products"].keys()), "units", m) for m in months]
    labels = [f"{_MONTHS[int(m[5:])][:3]} {m[2:4]}" for m in months]

    fig, ax = plt.subplots(figsize=(5.6, 2.5), dpi=150)
    ax.plot(xs, total, color="#12805c", lw=2.6, marker="o", ms=4,
            markerfacecolor="#12805c")
    ax.fill_between(xs, total, color="#12805c", alpha=0.08)
    ax.set_xticks(list(xs))
    ax.set_xticklabels(labels, fontsize=7.5, rotation=45, ha="right",
                       color="#666")
    ax.set_ylim(bottom=0)
    ax.tick_params(axis="y", labelsize=7.5, colors="#666", length=0)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v/1000:.0f}K" if v else "0")
    for sp in ("top", "right", "left"):
        ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color("#ddd")
    ax.grid(axis="y", color="#eee", lw=0.8)
    ax.set_axisbelow(True)
    fig.tight_layout(pad=0.4)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def build_summary(d):
    months = d["months"]
    allp = list(d["products"].keys())
    u_last, u_prev = months[-1], months[-2]        # units/wholesale (snapshot)
    # net sales $ is only known where retail exists (from the L52W Excel)
    r_months = [m for m in months if _val(d, allp, "retail", m) is not None]
    n_last = r_months[-1] if r_months else None
    n_prev = (months[months.index(n_last) - 1]
              if n_last and months.index(n_last) > 0 else None)

    def cell(parts, metric, last, prev):
        cur = _val(d, parts, metric, last)
        return {"cur": cur, "mom": _growth(cur, _val(d, parts, metric, prev))}

    def uw(parts):
        return {"units": cell(parts, "units", u_last, u_prev),
                "wholesale": cell(parts, "wholesale", u_last, u_prev)}

    def net(parts):
        r = cell(parts, "retail", n_last, n_prev)
        cost = _val(d, parts, "wholesale", n_last)
        r["cost"] = cost
        r["margin"] = ((r["cur"] - cost) / r["cur"] * 100
                       if r["cur"] and cost is not None else None)
        return {"retail": r}

    uw_s = {"overall": uw(allp), "brands": [], "products": []}
    nm_s = {"overall": net(allp), "brands": [], "products": []}
    for b, color in (("Catalyst", CAT), ("Feline Fresh", FF)):
        bp = brand_parts(d, b)
        uw_s["brands"].append({"name": b, "color": color, **uw(bp)})
        nm_s["brands"].append({"name": b, "color": color, **net(bp)})
    for p, v in d["products"].items():
        if _val(d, [p], "units", u_last):
            uw_s["products"].append({"short": v["short"], "brand": v["brand"],
                                     **uw([p])})
        blk = net([p])
        if blk["retail"]["cur"]:
            nm_s["products"].append({"short": v["short"], "brand": v["brand"],
                                     **blk})
    uw_s["products"].sort(key=lambda x: -(x["wholesale"]["cur"] or 0))
    nm_s["products"].sort(key=lambda x: -(x["retail"]["cur"] or 0))

    def discount(parts):
        imp = _val(d, parts, "implied", n_last)
        r = _val(d, parts, "retail", n_last)
        return (imp - r) / imp * 100 if imp and r is not None else None
    disc = {"overall": discount(allp)}
    for b in ("Catalyst", "Feline Fresh"):
        disc[b] = discount(brand_parts(d, b))

    return {"u_last": u_last, "u_prev": u_prev, "n_last": n_last,
            "n_prev": n_prev, "uw": uw_s, "nm": nm_s, "discount": disc}


def _g(v):
    """MoM % as a colored pill."""
    if v is None:
        return '<span style="color:#c7c7c7;">–</span>'
    up = v >= 0
    c, bg = ("#0b7a4b", "#e4f4ec") if up else ("#c0392b", "#fbe8e6")
    return (f'<span style="background:{bg};color:{c};font-weight:700;'
            f'padding:2px 8px;border-radius:20px;font-size:12px;'
            f'white-space:nowrap;">{"▲" if up else "▼"}{abs(v):.0f}%</span>')


def _mval(metric, v):
    return num(v) if metric == "units" else money(v)


UW_METRICS = [("units", "Units"), ("wholesale", "Wholesale $")]


def render_html(d, s, img_src):
    ul, up = s["u_last"], s["u_prev"]
    nl = s["n_last"]

    def big_card(label, value, mom):
        return f"""
      <div style="background:#fff;border-radius:12px;border-left:5px solid {CAT};
           padding:13px 18px;margin-bottom:9px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <table role="presentation" width="100%" style="border-collapse:collapse;"><tr>
          <td style="vertical-align:middle;">
            <div style="font-size:12px;font-weight:700;color:#7a7a7a;text-transform:uppercase;letter-spacing:.04em;">{label}</div>
            <div style="font-size:25px;font-weight:800;color:#16281f;line-height:1.1;margin-top:2px;">{value}</div>
          </td>
          <td style="vertical-align:middle;text-align:right;white-space:nowrap;">{_g(mom)}</td>
        </tr></table>
      </div>"""

    over = ""
    for metric, label in UW_METRICS:
        c = s["uw"]["overall"][metric]
        over += big_card(label, _mval(metric, c["cur"]), c["mom"])

    def vcell(metric, mc):
        return f"""<td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;">
          <div style="font-size:12.5px;font-weight:700;color:#222;">{_mval(metric, mc['cur'])}</div>
          <div style="margin-top:3px;">{_g(mc['mom'])}</div></td>"""

    def uw_rows(rows_data):
        out = ""
        for r in rows_data:
            color = r.get("color") or (CAT if r.get("brand") == "Catalyst" else FF)
            name = r.get("name") or r.get("short")
            out += f"""<tr>
          <td style="padding:8px 4px;border-top:1px solid #eee;">
            <span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:{color};margin-right:6px;"></span>
            <span style="font-size:12.5px;color:#222;">{name}</span></td>
          {vcell('units', r['units'])}{vcell('wholesale', r['wholesale'])}
        </tr>"""
        return out

    def uw_card(title, rows_data):
        return f"""
      <div style="background:#fff;border-radius:12px;padding:14px 12px 6px;margin-top:14px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 6px 8px;">{title}</div>
        <table role="presentation" width="100%" style="border-collapse:collapse;table-layout:fixed;">
          <tr style="color:#9a9a9a;font-size:9.5px;text-transform:uppercase;">
            <td style="padding:0 4px 4px;width:44%;"></td>
            <td style="padding:0 3px 4px;text-align:center;">Units</td>
            <td style="padding:0 3px 4px;text-align:center;">Wholesale&nbsp;$</td></tr>
          {uw_rows(rows_data)}
        </table>
        <div style="font-size:9.5px;color:#b3b3b3;margin:6px 6px 2px;">value = {mlabel(ul)} · pill = MoM vs {mlabel(up)}</div>
      </div>"""

    nc = s["nm"]["overall"]["retail"]
    n_over = big_card("Net Sales $", money(nc["cur"]), nc["mom"])

    dsc = s["discount"]

    def dchip(name, v, color):
        val = "—" if v is None else f"{v:.0f}%"
        return f"""<td style="text-align:center;padding:6px;">
          <div style="font-size:11px;color:#7a7a7a;"><span style="color:{color};">●</span> {name}</div>
          <div style="font-size:20px;font-weight:800;color:#16281f;">{val}</div></td>"""
    disc_card = f"""
      <div style="background:#fff;border-radius:12px;padding:13px 14px;margin-top:12px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 6px 4px;">Discount vs list price</div>
        <div style="font-size:11px;color:#9a9a9a;margin:0 6px 6px;">net sell-through below list retail</div>
        <table role="presentation" width="100%" style="border-collapse:collapse;"><tr>
          {dchip("Overall", dsc["overall"], "#555")}
          {dchip("Catalyst", dsc["Catalyst"], CAT)}
          {dchip("Feline Fresh", dsc["Feline Fresh"], FF)}
        </tr></table>
      </div>"""

    def mrow(name, color, r, bold=False):
        net_, cost, m = r["cur"], r["cost"], r["margin"]
        if m is None:
            badge = '<span style="color:#c7c7c7;">—</span>'
        else:
            mc = "#0b7a4b" if m >= 15 else ("#c0392b" if m < 5 else "#c98a00")
            badge = f'<span style="color:{mc};font-weight:800;">{m:.0f}%</span>'
        fw = "800" if bold else "400"
        return f"""<tr>
          <td style="padding:8px 4px;border-top:1px solid #eee;">
            <span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:{color};margin-right:6px;"></span>
            <span style="font-size:12.5px;font-weight:{fw};color:#222;">{name}</span></td>
          <td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;font-size:12px;color:#444;">{money(net_)}</td>
          <td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;font-size:12px;color:#444;">{money(cost)}</td>
          <td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;font-size:14px;">{badge}</td>
        </tr>"""

    ov = s["nm"]["overall"]["retail"]
    mrows = mrow("All products", "#555", ov, True)
    for b in s["nm"]["brands"]:
        mrows += mrow(b["name"], b["color"], b["retail"], True)
    for p in s["nm"]["products"]:
        mrows += mrow(p["short"], CAT if p["brand"] == "Catalyst" else FF, p["retail"])
    margin_card = f"""
      <div style="background:#fff;border-radius:12px;padding:14px 12px 8px;margin-top:12px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 6px 3px;">Chewy net margin</div>
        <div style="font-size:10.5px;color:#9a9a9a;margin:0 6px 6px;">(net sales − wholesale cost) ÷ net sales</div>
        <table role="presentation" width="100%" style="border-collapse:collapse;table-layout:fixed;">
          <tr style="color:#9a9a9a;font-size:9.5px;text-transform:uppercase;">
            <td style="padding:0 4px 4px;width:40%;"></td>
            <td style="padding:0 3px 4px;text-align:center;">Net&nbsp;$</td>
            <td style="padding:0 3px 4px;text-align:center;">Cost&nbsp;$</td>
            <td style="padding:0 3px 4px;text-align:center;">Margin</td></tr>
          {mrows}
        </table>
      </div>"""

    return f"""<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;background:#eef1f0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#16281f;">
  <div style="max-width:460px;margin:0 auto;padding:14px;">
    <div style="background:#0d5c43;color:#fff;border-radius:12px;padding:16px 18px;">
      <div style="font-size:18px;font-weight:800;">🐾 Chewy Sales — {mlabel(ul)}</div>
      <div style="font-size:12px;opacity:.82;margin-top:3px;">Catalyst &amp; Feline Fresh · units through {mlabel(ul)} · net sales through {mlabel(nl)}</div>
    </div>

    <div style="background:#fff;border-radius:12px;padding:12px 10px 6px;margin-top:14px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
      <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 8px 6px;">Monthly unit sales · last 12 mo</div>
      <img src="{img_src}" width="100%" style="display:block;max-width:100%;border-radius:6px;" alt="Monthly unit sales">
    </div>

    <div style="font-size:11px;font-weight:800;color:#7a7a7a;text-transform:uppercase;letter-spacing:.05em;margin:16px 6px 4px;">Units &amp; wholesale — {mlabel(ul)}</div>
    <div style="font-size:10.5px;color:#9a9a9a;margin:0 6px 8px;">from the monthly Brand Snapshot (units actual; wholesale $ = units × your cost)</div>
    {over}
    {uw_card("By Brand", s["uw"]["brands"])}
    {uw_card("By Product", s["uw"]["products"])}

    <div style="font-size:11px;font-weight:800;color:#7a7a7a;text-transform:uppercase;letter-spacing:.05em;margin:20px 6px 4px;">Net sales &amp; margin — {mlabel(nl)}</div>
    <div style="font-size:10.5px;color:#9a9a9a;margin:0 6px 8px;">latest reported (Chewy's L52W report; the snapshot has no $, so this trails units)</div>
    {n_over}
    {disc_card}
    {margin_card}

    <div style="text-align:center;margin-top:20px;">
      <a href="{DASHBOARD_URL}" style="display:inline-block;background:#12805c;color:#fff;text-decoration:none;padding:12px 28px;border-radius:8px;font-weight:700;font-size:15px;">View full dashboard →</a>
      <div style="font-size:11px;color:#9a9a9a;margin-top:9px;">Password: pellets123 · Chewy sell-through · MoM = vs prior month</div>
    </div>
    <div style="text-align:center;font-size:10px;color:#b6b6b6;margin-top:14px;padding-bottom:6px;">
      Sent from {SENDER} · Catalyst / Feline Fresh internal use only · Private &amp; Confidential
    </div>
  </div>
</body></html>"""


def read_state():
    if os.path.exists(STATE_FILE):
        return open(STATE_FILE, encoding="utf-8").read().strip()
    return ""


def write_state(month):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        f.write(month)


def send(html, subject, recipients, png):
    msg = MIMEMultipart("related")
    msg["Subject"] = subject
    msg["From"] = SENDER
    msg["To"] = ", ".join(recipients)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html, "html"))
    msg.attach(alt)
    img = MIMEImage(png, "png")
    img.add_header("Content-ID", "<unitschart>")
    img.add_header("Content-Disposition", "inline", filename="units.png")
    msg.attach(img)
    pw = _cred("EMAIL_APP_PASSWORD")
    if not pw:
        raise SystemExit("[chewy_email] EMAIL_APP_PASSWORD not set")
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx) as srv:
        srv.login(SENDER, pw)
        srv.sendmail(SENDER, recipients, msg.as_string())


def main():
    global STATE_FILE
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="Render preview HTML, do not send.")
    ap.add_argument("--dev-only", action="store_true",
                    help="Send to the dev recipient only.")
    ap.add_argument("--force", action="store_true",
                    help="Send even if this month was already emailed.")
    ap.add_argument("--state-file", default=None,
                    help="Path to the last-emailed-month state file.")
    args = ap.parse_args()
    if args.state_file:
        STATE_FILE = args.state_file

    d = load_payload()
    s = build_summary(d)
    png = make_units_png(d)
    subject = f"Chewy Sales Recap — {mlabel(s['u_last'])} (Catalyst & Feline Fresh)"

    data_uri = "data:image/png;base64," + base64.b64encode(png).decode()
    with open(PREVIEW, "w", encoding="utf-8") as f:
        f.write(render_html(d, s, data_uri))
    print(f"[chewy_email] latest month {s['u_last']}, preview -> {PREVIEW}")

    if args.dry_run:
        return

    if not args.force and read_state() == s["u_last"]:
        print(f"[chewy_email] {s['u_last']} already emailed — skipping "
              "(use --force to resend).")
        return

    # Wait until BOTH brands have data for the latest month (a single-brand
    # snapshot arriving first shouldn't trigger a half-empty recap).
    present = {b["name"] for b in s["uw"]["brands"] if b["units"]["cur"] > 0}
    if not args.force and not {"Catalyst", "Feline Fresh"} <= present:
        missing = {"Catalyst", "Feline Fresh"} - present
        print(f"[chewy_email] {s['u_last']} incomplete — waiting for "
              f"{', '.join(sorted(missing))} snapshot (use --force to send).")
        return

    recipients = DEV_RECIPIENTS if args.dev_only else EMAIL_TO
    send(render_html(d, s, "cid:unitschart"), subject, recipients, png)
    print(f"[chewy_email] sent '{subject}' to {len(recipients)} recipient(s).")
    if not args.dev_only:
        write_state(s["u_last"])


if __name__ == "__main__":
    main()
