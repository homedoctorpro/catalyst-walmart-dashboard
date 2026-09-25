"""
prospeo_retry.py — finish the contact lookups Prospeo's rate limit refused.

Reads Reports/prospeo_candidates.csv, retries the rows that have a person_id
but no email yet, and writes whatever it gets into the sheet's
suggested_contacts column. Designed to be run on a schedule: it does a handful
per run, stops cleanly when the rate limit bites, and exits quietly when there
is nothing left to do.

    python prospeo_retry.py            # retry up to MAX_PER_RUN
    python prospeo_retry.py --all      # keep going until the limit says no
"""

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.request
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(HERE, "Reports", "prospeo_candidates.csv")
MAX_PER_RUN = 8
PAUSE = 20            # the free plan is unhappy with anything faster

try:
    from credentials import PROSPEO_KEY as KEY
except ImportError:
    KEY = os.environ.get("PROSPEO_KEY", "")
try:
    from credentials import RETAILERS_SYNC_URL as SHEET_URL
except ImportError:
    SHEET_URL = os.environ.get("RETAILERS_SYNC_URL", "")


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None


_opener = urllib.request.build_opener(_NoRedirect)


def prospeo(path, payload):
    req = urllib.request.Request(f"https://api.prospeo.io/{path}", method="POST",
                                 data=json.dumps(payload).encode(),
                                 headers={"Content-Type": "application/json", "X-KEY": KEY})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try:
            return json.loads(e.read().decode())
        except json.JSONDecodeError:
            return {"error": True, "error_code": f"HTTP {e.code}"}


def sheet(payload=None, query=""):
    req = urllib.request.Request(SHEET_URL + query, method="POST" if payload else "GET",
                                 data=json.dumps(payload).encode() if payload else None,
                                 headers={"Content-Type": "application/json"} if payload else {})
    try:
        with _opener.open(req, timeout=120) as r:
            body = r.read().decode()
    except urllib.error.HTTPError as e:
        loc = e.headers.get("Location") if e.headers else None
        if not loc:
            raise
        with urllib.request.urlopen(loc, timeout=120) as r2:
            body = r2.read().decode()
    return json.loads(body)


def push_to_sheet(rows_by_id):
    """Rewrite suggested_contacts for every retailer that has leads."""
    live = {r["id"]: r for r in sheet(query="?action=rows")["rows"]}
    leads = defaultdict(list)
    for r in rows_by_id:
        if r.get("email"):
            leads[r["retailer_id"]].append(f'{r["name"]} <{r["email"]}> · {r["title"]}')
    for rid, lines in leads.items():
        row = live.get(rid)
        if not row:
            continue
        if row.get("suggested_contacts", "") == "\n".join(lines):
            continue                                   # already current
        fields = {
            "status": row.get("status", ""), "priority": row.get("priority") or "",
            "repFirm": row.get("rep_firm", ""), "nextSteps": row.get("next_steps", ""),
            "deadline": row.get("deadline", ""), "nextReview": row.get("next_review", ""),
            "resetDate": row.get("reset_date", ""),
            "needsDistributor": "1" if row.get("needs_distributor") else "",
            "distributorName": row.get("distributor_name", ""),
            "keyContact": row.get("key_contact", ""),
            "suggestedContacts": "\n".join(lines),
        }
        out = sheet({"action": "upsert", "item": {"retailerId": rid, "fields": fields}})
        print(f"  sheet: {row['name']} ← {len(lines)} lead(s) "
              f"{'ok' if out.get('ok') else out}", flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true", help="keep going past MAX_PER_RUN")
    args = ap.parse_args()

    if not KEY:
        print("[prospeo] no API key configured — nothing to do")
        return 0
    if not os.path.exists(CSV_PATH):
        print("[prospeo] no candidates file — nothing to do")
        return 0

    rows = list(csv.DictReader(open(CSV_PATH, encoding="utf-8")))
    pending = [r for r in rows if r.get("person_id") and not r.get("email")
               and "no email on record" not in (r.get("email_status") or "")]
    if not pending:
        print("[prospeo] nothing pending")
        return 0

    budget = len(pending) if args.all else min(MAX_PER_RUN, len(pending))
    print(f"[prospeo] {len(pending)} pending, trying {budget}", flush=True)

    done = 0
    for r in pending[:budget]:
        out = prospeo("enrich-person", {"data": {"person_id": r["person_id"],
                                                 "company_website": r.get("domain", "")}})
        code = str(out.get("error_code") or "")
        if code.lower().startswith("rate limit"):
            print(f"  rate limited at {r['name']} — stopping, will pick up next run", flush=True)
            break
        if out.get("error"):
            r["email_status"] = code or "error"
        else:
            em = (out.get("person") or {}).get("email")
            if isinstance(em, dict):
                r["email"] = em.get("email") or ""
                r["email_status"] = em.get("status") or ("found" if r["email"] else "no email on record")
            else:
                r["email"] = em or ""
                r["email_status"] = "found" if r["email"] else "no email on record"
            done += 1
        print(f"  {r['retailer'][:26]:28s} {r['name'][:24]:26s} "
              f"{r['email'] or r['email_status']}", flush=True)
        with open(CSV_PATH, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=rows[0].keys())
            w.writeheader()
            w.writerows(rows)
        time.sleep(PAUSE)

    if done and SHEET_URL:
        push_to_sheet(rows)

    still = [r for r in rows if r.get("person_id") and not r.get("email")
             and "no email on record" not in (r.get("email_status") or "")]
    print(f"[prospeo] {done} revealed this run · {len(still)} still pending")
    return 0


if __name__ == "__main__":
    sys.exit(main())
