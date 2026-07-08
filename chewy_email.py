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


def render_html(d, s):
    last, prev = s["last"], s["prev"]

    # ---- Overall: three full-width stacked cards (stack fine on mobile) ----
    over = ""
    for metric, label in METRICS:
        c = s["overall"][metric]
        over += f"""
      <div style="background:#fff;border-radius:12px;border-left:5px solid {CAT};
           padding:14px 18px;margin-bottom:10px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <table role="presentation" width="100%" style="border-collapse:collapse;"><tr>
          <td style="vertical-align:middle;">
            <div style="font-size:12px;font-weight:700;color:#7a7a7a;text-transform:uppercase;letter-spacing:.04em;">{label}</div>
            <div style="font-size:26px;font-weight:800;color:#16281f;line-height:1.1;margin-top:2px;">{_mval(metric, c['cur'])}</div>
          </td>
          <td style="vertical-align:middle;text-align:right;white-space:nowrap;">{_g(c['mom'])}</td>
        </tr></table>
      </div>"""

    # ---- Narrow growth table: name + 3 short MoM% columns (mobile-safe) ----
    def rows_html(rows_data):
        out = ""
        for r in rows_data:
            color = r.get("color") or (CAT if r.get("brand") == "Catalyst" else FF)
            name = r.get("name") or r.get("short")
            out += f"""<tr>
          <td style="padding:9px 6px;border-top:1px solid #eee;">
            <span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:{color};margin-right:7px;"></span>
            <span style="font-size:13px;color:#222;">{name}</span></td>
          <td style="padding:9px 4px;border-top:1px solid #eee;text-align:center;">{_g(r['units']['mom'])}</td>
          <td style="padding:9px 4px;border-top:1px solid #eee;text-align:center;">{_g(r['wholesale']['mom'])}</td>
          <td style="padding:9px 4px;border-top:1px solid #eee;text-align:center;">{_g(r['retail']['mom'])}</td>
        </tr>"""
        return out

    def card(title, rows_data):
        return f"""
      <div style="background:#fff;border-radius:12px;padding:14px 14px 6px;margin-top:14px;box-shadow:0 1px 3px rgba(0,0,0,.1);">
        <div style="font-size:13px;font-weight:800;color:#0d5c43;margin:2px 6px 8px;">{title}</div>
        <table role="presentation" width="100%" style="border-collapse:collapse;">
          <tr style="color:#9a9a9a;font-size:10px;text-transform:uppercase;letter-spacing:.03em;">
            <td style="padding:0 6px 4px;">MoM growth</td>
            <td style="padding:0 4px 4px;text-align:center;">Units</td>
            <td style="padding:0 4px 4px;text-align:center;">Whsl&nbsp;$</td>
            <td style="padding:0 4px 4px;text-align:center;">Retail&nbsp;$</td></tr>
          {rows_html(rows_data)}
        </table>
      </div>"""

    return f"""<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;background:#eef1f0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#16281f;">
  <div style="max-width:460px;margin:0 auto;padding:14px;">
    <div style="background:#0d5c43;color:#fff;border-radius:12px;padding:16px 18px;">
      <div style="font-size:18px;font-weight:800;">🐾 Chewy Sales — {mlabel(last)}</div>
      <div style="font-size:12px;opacity:.82;margin-top:3px;">Catalyst &amp; Feline Fresh · growth vs {mlabel(prev)}</div>
    </div>

    <div style="font-size:11px;font-weight:800;color:#7a7a7a;text-transform:uppercase;letter-spacing:.05em;margin:16px 6px 8px;">Overall</div>
    {over}
    {card("By Brand", s["brands"])}
    {card("By Product", s["products"])}

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
