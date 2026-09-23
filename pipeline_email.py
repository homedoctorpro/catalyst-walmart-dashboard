"""
pipeline_email.py — Catalyst Pet retailer-pipeline daily change email.

Reads the Retailers sheet through its Apps Script web app, compares it with the
last snapshot, and emails the group only when something actually moved. Also
flags deadlines that have passed.

  python pipeline_email.py --dry-run     # write a preview, send nothing
  python pipeline_email.py               # send if there are changes
  python pipeline_email.py --force       # send even with nothing to report
  python pipeline_email.py --seed        # save today's state without emailing

The sync URL comes from RETAILERS_SYNC_URL (env or credentials.py). The snapshot
lives in Reports/pipeline_state.json and is committed, so the cloud run and a
local run share one history.
"""

import os, sys, json, argparse, smtplib, urllib.request
from datetime import date, datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

HERE = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.path.join(HERE, "Reports", "pipeline_state.json")
DASHBOARD_URL = "https://homedoctorpro.github.io/catalyst-walmart-dashboard/dashboard.html"

try:
    from credentials import EMAIL_USER, EMAIL_APP_PASSWORD
except ImportError:
    EMAIL_USER = os.environ.get("EMAIL_USER", "ligneticsdata@gmail.com")
    EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")

try:
    from credentials import RETAILERS_SYNC_URL
except ImportError:
    RETAILERS_SYNC_URL = os.environ.get("RETAILERS_SYNC_URL", "")

EMAIL_FROM = f"Catalyst Pet Reports <{EMAIL_USER}>"
EMAIL_TO = [
    "pross@lignetics.com",
    "ckohagen@lignetics.com",      # Cindee Kohagen
    "mscanlon@lignetics.com",      # Michael Scanlon
    "JGallman@Lignetics.com",      # Jim Gallman
    "ckaminski@lignetics.com",     # Cammie Kaminski
]
DEV_TO = ["pross@lignetics.com"]

# Fields worth telling people about, in the order they read best.
TRACKED = [
    ("status",           "Status"),
    ("priority",         "Priority"),
    ("repFirm",          "Rep firm"),
    ("nextSteps",        "Next steps"),
    ("deadline",         "Deadline"),
    ("nextReview",       "Next review"),
    ("resetDate",        "Reset date"),
    ("needsDistributor", "Needs distributor"),
    ("distributorName",  "Distributor"),
    ("keyContact",       "Key contact"),
    ("sfOppStage",       "Opportunity stage"),
]
STATUS_LABEL = {"in": "Currently In", "pitched": "Pitched", "target": "Target",
                "non-target": "Non-Target", "declined": "Declined", "": "—"}

BLUE, INK, MUTED = "#0057e7", "#1a1a2e", "#888"


# ── data ─────────────────────────────────────────────────────────────────────

def fetch_sheet(url):
    with urllib.request.urlopen(url + "?action=get", timeout=120) as r:
        payload = json.load(r)
    if not payload.get("ok"):
        raise RuntimeError(payload.get("error") or "sync endpoint returned ok:false")
    return payload.get("data") or {}, payload.get("unlinkedOpps") or []


def retailer_names():
    """id → display name, read straight out of the dashboard template."""
    import re
    src = open(os.path.join(HERE, "dashboard_template.html"), encoding="utf-8").read()
    block = src[src.index("const RT_RETAILERS"):src.index("const RT_STATUS_OPTIONS")]
    out = {}
    for rid, name in re.findall(r'\{\s*id:\s*"([^"]+)",\s*name:\s*"([^"]+)"', block):
        out[rid] = name
    return out


def load_state():
    try:
        with open(STATE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_state(rows, sent_on=None, overdue=None):
    os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
    prev = load_state()
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "savedAt": datetime.now().isoformat(timespec="seconds"),
            "lastSentOn": sent_on or prev.get("lastSentOn"),
            "overdueSeen": sorted(overdue if overdue is not None else prev.get("overdueSeen", [])),
            "rows": {rid: {k: v for k, v in row.items() if k in dict(TRACKED)}
                     for rid, row in rows.items()},
        }, f, indent=1, sort_keys=True)


# ── diffing ──────────────────────────────────────────────────────────────────

def fmt(field, value):
    if value in (None, ""):
        return "—"
    if field == "status":
        return STATUS_LABEL.get(str(value), str(value))
    if field == "needsDistributor":
        return "yes"
    return str(value)


def diff_rows(old_rows, new_rows, names):
    """[{id, name, changes:[(label, before, after)], isNew}] sorted by name."""
    out = []
    for rid, new in new_rows.items():
        old = old_rows.get(rid)
        changes = []
        for field, label in TRACKED:
            before, after = (old or {}).get(field, ""), new.get(field, "")
            if str(before or "") != str(after or ""):
                changes.append((label, fmt(field, before), fmt(field, after)))
        if changes:
            out.append({"id": rid, "name": names.get(rid, rid),
                        "changes": changes, "isNew": old is None})
    for rid in old_rows:
        if rid not in new_rows:
            out.append({"id": rid, "name": names.get(rid, rid), "gone": True,
                        "changes": [("Row", "tracked", "no longer on the list")]})
    return sorted(out, key=lambda r: r["name"].lower())


def overdue_deadlines(rows, names, today):
    out = []
    for rid, row in rows.items():
        dl = str(row.get("deadline") or "")
        if len(dl) == 10 and dl < today and row.get("status") in ("target", "pitched"):
            out.append({"id": rid, "name": names.get(rid, rid), "deadline": dl,
                        "status": STATUS_LABEL.get(row.get("status"), ""),
                        "owner": row.get("repFirm") or "Unassigned",
                        "nextSteps": row.get("nextSteps") or ""})
    return sorted(out, key=lambda r: r["deadline"])


# ── email ────────────────────────────────────────────────────────────────────

def build_html(changes, overdue, new_overdue, unlinked, counts, today):
    def esc(t):
        return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    rows_html = ""
    for c in changes:
        tag = ""
        if c.get("isNew"):
            tag = f' <span style="font-size:11px;color:{BLUE};font-weight:700">NEW</span>'
        elif c.get("gone"):
            tag = ' <span style="font-size:11px;color:#b71c1c;font-weight:700">REMOVED</span>'
        detail = "<br>".join(
            f'<span style="color:{MUTED}">{esc(label)}:</span> {esc(before)} '
            f'<span style="color:{MUTED}">→</span> <strong>{esc(after)}</strong>'
            for label, before, after in c["changes"])
        rows_html += (
            f'<tr><td style="padding:9px 12px;border-bottom:1px solid #eee;vertical-align:top;'
            f'white-space:nowrap"><strong>{esc(c["name"])}</strong>{tag}</td>'
            f'<td style="padding:9px 12px;border-bottom:1px solid #eee;font-size:13px;'
            f'line-height:1.6">{detail}</td></tr>')

    overdue_html = ""
    if overdue:
        lines = ""
        for o in overdue:
            fresh = ' <span style="font-size:11px;color:#b71c1c;font-weight:700">NEW</span>' \
                    if o["id"] in new_overdue else ""
            lines += (
                f'<tr><td style="padding:8px 12px;border-bottom:1px solid #f3dede;white-space:nowrap">'
                f'<strong>{esc(o["name"])}</strong>{fresh}</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #f3dede;color:#b71c1c;'
                f'font-weight:700;white-space:nowrap">{esc(o["deadline"])}</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #f3dede;white-space:nowrap">{esc(o["status"])}</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #f3dede;white-space:nowrap">{esc(o["owner"])}</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #f3dede;font-size:13px">{esc(o["nextSteps"])}</td></tr>')
        overdue_html = f"""
        <h2 style="font-size:14px;text-transform:uppercase;letter-spacing:.05em;color:#b71c1c;
                   margin:26px 0 8px">Deadlines passed ({len(overdue)})</h2>
        <table style="border-collapse:collapse;width:100%;background:#fff6f6;border-radius:8px">
          <tr style="font-size:11px;text-transform:uppercase;color:{MUTED};text-align:left">
            <th style="padding:6px 12px">Retailer</th><th style="padding:6px 12px">Deadline</th>
            <th style="padding:6px 12px">Status</th><th style="padding:6px 12px">Rep</th>
            <th style="padding:6px 12px">Next steps</th></tr>
          {lines}
        </table>"""

    unlinked_html = ""
    if unlinked:
        items = "".join(
            f'<li style="margin-bottom:3px">{esc(o["account"])} — {esc(o["name"])} '
            f'<span style="color:{MUTED}">({esc(o["stage"])})</span></li>'
            for o in unlinked if o.get("stage") and "Closed" not in str(o.get("stage")))
        if items:
            unlinked_html = f"""
        <h2 style="font-size:14px;text-transform:uppercase;letter-spacing:.05em;color:{MUTED};
                   margin:26px 0 8px">Open litter opportunities not on the list</h2>
        <ul style="font-size:13px;line-height:1.6;padding-left:18px;margin:0">{items}</ul>"""

    body = rows_html and f"""
        <h2 style="font-size:14px;text-transform:uppercase;letter-spacing:.05em;color:{MUTED};
                   margin:22px 0 8px">Changes ({len(changes)})</h2>
        <table style="border-collapse:collapse;width:100%">{rows_html}</table>""" or ""

    return f"""<!doctype html><html><body style="margin:0;background:#f4f5f7;
      font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;color:{INK}">
      <div style="max-width:720px;margin:0 auto;padding:22px">
        <div style="background:#fff;border-radius:12px;padding:22px 24px;box-shadow:0 2px 8px rgba(0,0,0,.08)">
          <div style="font-size:12px;color:{MUTED};text-transform:uppercase;letter-spacing:.06em">
            Catalyst Pet · Retailer pipeline</div>
          <h1 style="font-size:20px;margin:4px 0 2px">What moved since the last update</h1>
          <div style="font-size:13px;color:{MUTED}">{today} · {counts['working']} retailers in the working list ·
            {counts['target']} target · {counts['pitched']} pitched</div>
          {body}
          {overdue_html}
          {unlinked_html}
          <p style="margin:26px 0 0">
            <a href="{DASHBOARD_URL}" style="background:{BLUE};color:#fff;text-decoration:none;
               padding:9px 16px;border-radius:8px;font-size:13px;font-weight:700">Open the dashboard</a>
          </p>
          <div style="font-size:11px;color:#aaa;margin-top:16px;line-height:1.5">
            Sent only when something changed. Status, priority, rep firm, next steps, dates,
            distributor and Salesforce opportunity stage are all watched.
          </div>
        </div>
      </div></body></html>"""


def send(html, subject, dev_only):
    to = DEV_TO if dev_only else EMAIL_TO
    if not EMAIL_APP_PASSWORD:
        print("  [Email] no app password configured — nothing sent")
        return False
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = subject, EMAIL_FROM, ", ".join(to)
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(EMAIL_USER, EMAIL_APP_PASSWORD)
            s.sendmail(EMAIL_USER, to, msg.as_string())
    except Exception as e:
        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as s:
                s.starttls()
                s.login(EMAIL_USER, EMAIL_APP_PASSWORD)
                s.sendmail(EMAIL_USER, to, msg.as_string())
        except Exception as e2:
            print(f"  [Email] FAILED (465: {e}) (587: {e2})")
            return False
    print(f"  [Email] sent to {len(to)} recipient(s): {subject}")
    return True


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="write a preview, send nothing")
    ap.add_argument("--force", action="store_true", help="send even with nothing to report")
    ap.add_argument("--seed", action="store_true", help="store today's state, send nothing")
    ap.add_argument("--dev", action="store_true", help="send to the dev address only")
    ap.add_argument("--once-per-day", action="store_true",
                    help="skip if an email already went out today (for the hourly cloud run)")
    args = ap.parse_args()

    if not RETAILERS_SYNC_URL:
        print("  [Pipeline] RETAILERS_SYNC_URL not set — nothing to read")
        return 0

    today = date.today().isoformat()
    state = load_state()
    if args.once_per_day and state.get("lastSentOn") == today and not args.force:
        print(f"  [Pipeline] already sent today ({today})")
        return 0

    rows, unlinked = fetch_sheet(RETAILERS_SYNC_URL)
    names = retailer_names()
    changes = diff_rows(state.get("rows", {}), rows, names)

    overdue = overdue_deadlines(rows, names, today)
    seen = set(state.get("overdueSeen", []))
    new_overdue = {o["id"] for o in overdue} - seen

    counts = {
        "working": sum(1 for r in rows.values()
                       if r.get("status") not in ("in", "non-target", "declined")),
        "target": sum(1 for r in rows.values() if r.get("status") == "target"),
        "pitched": sum(1 for r in rows.values() if r.get("status") == "pitched"),
    }

    first_run = not state.get("rows")
    if first_run and not args.force:
        save_state(rows, overdue=[o["id"] for o in overdue])
        print(f"  [Pipeline] first run — stored {len(rows)} rows, no email")
        return 0

    if not changes and not new_overdue and not args.force:
        save_state(rows, overdue=[o["id"] for o in overdue])
        print("  [Pipeline] nothing changed — no email")
        return 0

    bits = []
    if changes:
        bits.append(f"{len(changes)} retailer{'' if len(changes) == 1 else 's'} updated")
    if new_overdue:
        bits.append(f"{len(new_overdue)} deadline{'' if len(new_overdue) == 1 else 's'} passed")
    subject = "Catalyst retailer pipeline — " + (", ".join(bits) or "no changes")

    html = build_html(changes, overdue, new_overdue, unlinked, counts, today)

    if args.dry_run:
        out = os.path.join(HERE, "_pipeline_email_preview.html")
        with open(out, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  [Pipeline] dry run — {len(changes)} change(s), {len(overdue)} overdue "
              f"({len(new_overdue)} new). Preview: {out}")
        return 0

    if send(html, subject, args.dev):
        save_state(rows, sent_on=today, overdue=[o["id"] for o in overdue])
    return 0


if __name__ == "__main__":
    sys.exit(main())
