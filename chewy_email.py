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
import ssl
import json
import smtplib
import argparse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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

    return {"last": last, "prev": prev, "yoy_month": yoy_month,
            "overall": overall, "brands": brands, "products": products}


def _g(v):
    """Colored MoM/YoY badge."""
    if v is None:
        return '<span style="color:#bbb;">—</span>'
    c = "#12805c" if v >= 0 else "#d32f2f"
    return f'<span style="color:{c};font-weight:700;">{"▲" if v>=0 else "▼"}{abs(v):.0f}%</span>'


def _mval(metric, v):
    return num(v) if metric == "units" else money(v)


def render_html(d, s):
    last = s["last"]

    # ---- overall: one growth card per metric (Units / Wholesale / Retail) ----
    cards = ""
    for metric, label in METRICS:
        c = s["overall"][metric]
        yoy = (f'<div style="font-size:11px;color:#aaa;margin-top:3px;">'
               f'{_g(c["yoy"])} vs {mlabel(s["yoy_month"])}</div>'
               if c["yoy"] is not None else
               '<div style="font-size:11px;color:#ccc;margin-top:3px;">no YoY</div>')
        cards += f"""<td style="padding:6px;width:33%;">
          <div style="background:#fff;border-radius:10px;border-left:4px solid {CAT};padding:13px 15px;box-shadow:0 1px 4px rgba(0,0,0,.08);">
            <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.04em;">{label}</div>
            <div style="font-size:20px;font-weight:800;color:#1a1a2e;margin-top:3px;">{_mval(metric, c['cur'])} {_g(c['mom'])}</div>
            <div style="font-size:10px;color:#bbb;margin-top:1px;">MoM</div>{yoy}
          </div></td>"""

    # ---- shared growth table (3 metric cols, value + MoM% each) ----
    def grow_table(title, rows_data, name_key):
        rows = ""
        for r in rows_data:
            cells = ""
            for metric, _ in METRICS:
                cells += (f'<td style="padding:6px 8px;text-align:right;white-space:nowrap;">'
                          f'{_mval(metric, r[metric]["cur"])} &nbsp;{_g(r[metric]["mom"])}</td>')
            dot = (f'<span style="display:inline-block;width:8px;height:8px;border-radius:50%;'
                   f'background:{r.get("color", CAT if r.get("brand")=="Catalyst" else FF)};margin-right:6px;"></span>'
                   if name_key != "plain" else "")
            rows += (f'<tr><td style="padding:6px 8px;">{dot}{r["name"] if "name" in r else r["short"]}</td>{cells}</tr>')
        return f"""<div style="background:#fff;border-radius:12px;padding:16px 18px;margin-top:14px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow-x:auto;">
      <div style="font-size:12px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px;">{title}</div>
      <table style="width:100%;border-collapse:collapse;font-size:12.5px;">
        <tr style="color:#888;font-size:10px;text-transform:uppercase;">
          <td style="padding:5px 8px;"></td>
          <td style="padding:5px 8px;text-align:right;">Units (MoM)</td>
          <td style="padding:5px 8px;text-align:right;">Wholesale $ (MoM)</td>
          <td style="padding:5px 8px;text-align:right;">Retail $ (MoM)</td></tr>
        {rows}</table></div>"""

    brand_tbl = grow_table(f"By Brand — {mlabel(last)}", s["brands"], "brand")
    prod_tbl = grow_table(f"By Product — {mlabel(last)}", s["products"], "brand")

    return f"""<!DOCTYPE html><html><body style="margin:0;background:#f0f2f5;font-family:-apple-system,Segoe UI,Roboto,sans-serif;color:#1a1a2e;">
  <div style="max-width:700px;margin:0 auto;padding:20px;">
    <div style="background:linear-gradient(90deg,#0b3d2e,#0d5c43);color:#fff;border-radius:12px;padding:18px 22px;">
      <div style="font-size:19px;font-weight:800;">🐾 Chewy Sales — {mlabel(last)}</div>
      <div style="font-size:13px;opacity:.8;margin-top:2px;">Catalyst Pet &amp; Feline Fresh · growth in units, wholesale $ &amp; retail $</div>
    </div>

    <div style="font-size:12px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:.05em;margin:14px 6px 2px;">Overall — {mlabel(last)}</div>
    <table style="width:100%;border-collapse:collapse;"><tr>{cards}</tr></table>

    {brand_tbl}
    {prod_tbl}

    <div style="text-align:center;margin-top:18px;">
      <a href="{DASHBOARD_URL}" style="display:inline-block;background:#12805c;color:#fff;text-decoration:none;padding:11px 26px;border-radius:8px;font-weight:700;font-size:14px;">View full dashboard →</a>
      <div style="font-size:11px;color:#aaa;margin-top:8px;">Password: pellets123 · figures are Chewy sell-through · MoM = vs prior month</div>
    </div>
    <div style="text-align:center;font-size:11px;color:#bbb;margin-top:14px;">
      Sent from {SENDER} · Catalyst Pet / Feline Fresh internal use only · Private &amp; Confidential
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


def send(html, subject, recipients):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = SENDER
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(html, "html"))
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
    html = render_html(d, s)
    subject = f"Chewy Sales Recap — {mlabel(s['last'])} (Catalyst & Feline Fresh)"

    with open(PREVIEW, "w", encoding="utf-8") as f:
        f.write(html)
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
    send(html, subject, recipients)
    print(f"[chewy_email] sent '{subject}' to {len(recipients)} recipient(s).")
    if not args.dev_only:
        write_state(s["last"])


if __name__ == "__main__":
    main()
