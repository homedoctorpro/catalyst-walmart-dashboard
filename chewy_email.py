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
    tot = 0
    for p in parts:
        if metric == "wholesale":
            src = d["wholesale"].get(p, {})
        elif metric == "implied":
            src = d["implied"].get(p, {})
        else:
            src = d["series"][p][metric]
        tot += src.get(month, 0)
    return tot


def _growth(a, b):
    return (a - b) / b * 100 if b else None


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


def _cell(d, parts, metric, last, prev, yoy_month, months):
    cur = _val(d, parts, metric, last)
    return {
        "cur": cur,
        "mom": _growth(cur, _val(d, parts, metric, prev)),
        "yoy": _growth(cur, _val(d, parts, metric, yoy_month))
        if yoy_month in months else None,
    }


def build_summary(d):
    months = d["months"]
    last, prev = months[-1], months[-2]
    yoy_month = f"{int(last[:4])-1}-{last[5:]}"
    allp = list(d["products"].keys())

    def block(parts):
        return {m: _cell(d, parts, m, last, prev, yoy_month, months)
                for m, _ in METRICS}

    overall = block(allp)
    brands = []
    for b, color in (("Catalyst", CAT), ("Feline Fresh", FF)):
        brands.append({"name": b, "color": color, **block(brand_parts(d, b))})

    products = []
    for p, v in d["products"].items():
        blk = block([p])
        if blk["retail"]["cur"] or blk["units"]["cur"]:
            products.append({"short": v["short"], "brand": v["brand"], **blk})
    products.sort(key=lambda x: -x["retail"]["cur"])

    # implied discount = (list retail - net sales) / list retail, latest month
    def discount(parts):
        imp = _val(d, parts, "implied", last)
        net = _val(d, parts, "retail", last)
        return (imp - net) / imp * 100 if imp else None
    disc = {"overall": discount(allp)}
    for b in ("Catalyst", "Feline Fresh"):
        disc[b] = discount(brand_parts(d, b))

    return {"last": last, "prev": prev, "yoy_month": yoy_month,
            "overall": overall, "brands": brands, "products": products,
            "discount": disc}


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


def render_html(d, s, img_src):
    last, prev = s["last"], s["prev"]

    # ---- Overall: three full-width stacked cards (value + MoM) ----
    over = ""
    for metric, label in METRICS:
        c = s["overall"][metric]
        over += f"""
      <div style="background:#fff;border-radius:12px;border-left:5px solid {CAT};
           padding:13px 18px;margin-bottom:9px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <table role="presentation" width="100%" style="border-collapse:collapse;"><tr>
          <td style="vertical-align:middle;">
            <div style="font-size:12px;font-weight:700;color:#7a7a7a;text-transform:uppercase;letter-spacing:.04em;">{label}</div>
            <div style="font-size:25px;font-weight:800;color:#16281f;line-height:1.1;margin-top:2px;">{_mval(metric, c['cur'])}</div>
          </td>
          <td style="vertical-align:middle;text-align:right;white-space:nowrap;">{_g(c['mom'])}</td>
        </tr></table>
      </div>"""

    # ---- implied discount vs list ----
    dsc = s["discount"]

    def dchip(name, v, color):
        val = "—" if v is None else f"{v:.0f}%"
        return f"""<td style="text-align:center;padding:6px;">
          <div style="font-size:11px;color:#7a7a7a;"><span style="color:{color};">●</span> {name}</div>
          <div style="font-size:20px;font-weight:800;color:#16281f;">{val}</div></td>"""
    disc_card = f"""
      <div style="background:#fff;border-radius:12px;padding:13px 14px;margin-top:14px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 6px 4px;">Discount vs list price</div>
        <div style="font-size:11px;color:#9a9a9a;margin:0 6px 6px;">how far net sell-through sits below list retail ({mlabel(last)})</div>
        <table role="presentation" width="100%" style="border-collapse:collapse;"><tr>
          {dchip("Overall", dsc["overall"], "#555")}
          {dchip("Catalyst", dsc["Catalyst"], CAT)}
          {dchip("Feline Fresh", dsc["Feline Fresh"], FF)}
        </tr></table>
      </div>"""

    # ---- value + MoM table (each cell: value over a MoM pill) ----
    def vcell(metric, mc):
        return f"""<td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;">
          <div style="font-size:12.5px;font-weight:700;color:#222;">{_mval(metric, mc['cur'])}</div>
          <div style="margin-top:3px;">{_g(mc['mom'])}</div></td>"""

    def rows_html(rows_data):
        out = ""
        for r in rows_data:
            color = r.get("color") or (CAT if r.get("brand") == "Catalyst" else FF)
            name = r.get("name") or r.get("short")
            out += f"""<tr>
          <td style="padding:8px 4px;border-top:1px solid #eee;">
            <span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:{color};margin-right:6px;"></span>
            <span style="font-size:12.5px;color:#222;">{name}</span></td>
          {vcell('units', r['units'])}{vcell('wholesale', r['wholesale'])}{vcell('retail', r['retail'])}
        </tr>"""
        return out

    def card(title, rows_data):
        return f"""
      <div style="background:#fff;border-radius:12px;padding:14px 12px 6px;margin-top:14px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 6px 8px;">{title}</div>
        <table role="presentation" width="100%" style="border-collapse:collapse;table-layout:fixed;">
          <tr style="color:#9a9a9a;font-size:9.5px;text-transform:uppercase;letter-spacing:.02em;">
            <td style="padding:0 4px 4px;width:34%;"></td>
            <td style="padding:0 3px 4px;text-align:center;">Units</td>
            <td style="padding:0 3px 4px;text-align:center;">Whsl&nbsp;$</td>
            <td style="padding:0 3px 4px;text-align:center;">Retail&nbsp;$</td></tr>
          {rows_html(rows_data)}
        </table>
        <div style="font-size:9.5px;color:#b3b3b3;margin:6px 6px 2px;">value = {mlabel(last)} · pill = MoM vs {mlabel(prev)}</div>
      </div>"""

    # ---- Chewy net margin table (net sell-through vs wholesale cost) ----
    def marginpct(net, cost):
        return (net - cost) / net * 100 if net else None

    def mrow(name, color, net, cost, bold=False):
        m = marginpct(net, cost)
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
          <td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;font-size:12px;color:#444;">{money(net)}</td>
          <td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;font-size:12px;color:#444;">{money(cost)}</td>
          <td style="padding:8px 3px;border-top:1px solid #eee;text-align:center;font-size:14px;">{badge}</td>
        </tr>"""

    ov = s["overall"]
    mrows = mrow("All products", "#555", ov["retail"]["cur"],
                 ov["wholesale"]["cur"], True)
    for b in s["brands"]:
        mrows += mrow(b["name"], b["color"], b["retail"]["cur"],
                      b["wholesale"]["cur"], True)
    for p in s["products"]:
        mrows += mrow(p["short"], CAT if p["brand"] == "Catalyst" else FF,
                      p["retail"]["cur"], p["wholesale"]["cur"])
    margin_card = f"""
      <div style="background:#fff;border-radius:12px;padding:14px 12px 8px;margin-top:14px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 6px 3px;">Chewy net margin — {mlabel(last)}</div>
        <div style="font-size:10.5px;color:#9a9a9a;margin:0 6px 6px;">(net sales − wholesale cost) ÷ net sales · what Chewy keeps after discounts</div>
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
      <div style="font-size:18px;font-weight:800;">🐾 Chewy Sales — {mlabel(last)}</div>
      <div style="font-size:12px;opacity:.82;margin-top:3px;">Catalyst &amp; Feline Fresh · growth vs {mlabel(prev)}</div>
    </div>

    <div style="background:#fff;border-radius:12px;padding:12px 10px 6px;margin-top:14px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
      <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 8px 6px;">Monthly unit sales · last 12 mo</div>
      <img src="{img_src}" width="100%" style="display:block;max-width:100%;border-radius:6px;" alt="Monthly unit sales">
    </div>

    <div style="font-size:11px;font-weight:800;color:#7a7a7a;text-transform:uppercase;letter-spacing:.05em;margin:16px 6px 8px;">Overall — {mlabel(last)}</div>
    {over}
    {disc_card}
    {card("By Brand", s["brands"])}
    {card("By Product", s["products"])}
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
    subject = f"Chewy Sales Recap — {mlabel(s['last'])} (Catalyst & Feline Fresh)"

    data_uri = "data:image/png;base64," + base64.b64encode(png).decode()
    with open(PREVIEW, "w", encoding="utf-8") as f:
        f.write(render_html(d, s, data_uri))
    print(f"[chewy_email] latest month {s['last']}, preview -> {PREVIEW}")

    if args.dry_run:
        return

    if not args.force and read_state() == s["last"]:
        print(f"[chewy_email] {s['last']} already emailed — skipping "
              "(use --force to resend).")
        return

    # Wait until BOTH brands have data for the latest month (a single-brand
    # snapshot arriving first shouldn't trigger a half-empty recap).
    present = {b["name"] for b in s["brands"] if b["units"]["cur"] > 0}
    if not args.force and not {"Catalyst", "Feline Fresh"} <= present:
        missing = {"Catalyst", "Feline Fresh"} - present
        print(f"[chewy_email] {s['last']} incomplete — waiting for "
              f"{', '.join(sorted(missing))} snapshot (use --force to send).")
        return

    recipients = DEV_RECIPIENTS if args.dev_only else EMAIL_TO
    send(render_html(d, s, "cid:unitschart"), subject, recipients, png)
    print(f"[chewy_email] sent '{subject}' to {len(recipients)} recipient(s).")
    if not args.dev_only:
        write_state(s["last"])


if __name__ == "__main__":
    main()
