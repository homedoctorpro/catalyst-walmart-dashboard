"""
prospeo_contacts.py — find cat-litter buyers for retailers that have no contact.

Two steps, because that is how Prospeo works:
  1. /search-person   finds people at a company (1 credit per page of 25)
  2. /enrich-person   reveals one person's email (credits per person)

Nothing is written to the sheet or Salesforce. Results land in
Reports/prospeo_candidates.csv for review.

    python prospeo_contacts.py --probe            # check the key + filter shape, no credits
    python prospeo_contacts.py --scope target     # Target/Pitched rows with no contact
    python prospeo_contacts.py --scope all        # every row with no contact
    python prospeo_contacts.py --scope target --enrich   # also reveal emails (spends credits)
"""

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(HERE, "Reports", "prospeo_candidates.csv")
API = "https://api.prospeo.io"

try:
    from credentials import PROSPEO_KEY as KEY
except ImportError:
    try:
        from credentials import PROSPEO_API_KEY as KEY
    except ImportError:
        KEY = os.environ.get("PROSPEO_KEY") or os.environ.get("PROSPEO_API_KEY", "")

try:
    from credentials import RETAILERS_SYNC_URL
except ImportError:
    RETAILERS_SYNC_URL = os.environ.get("RETAILERS_SYNC_URL", "")

# Titles worth a buyer conversation, best first.
TITLE_TERMS = [
    "cat litter buyer", "litter buyer", "pet buyer", "pet category manager",
    "category manager pet", "buyer pet", "merchant pet", "category manager",
    "senior buyer", "buyer", "merchant", "merchandise manager",
]
# Retailer id → corporate domain. Only what we need; anything missing is skipped
# and reported rather than guessed at.
DOMAINS = {
    "pet-supplies-plus": "petsuppliesplus.com", "pet-supermarket": "petsupermarket.com",
    "earthwise-pet": "earthwisepet.com", "kahoots": "kahoots.com",
    "tomlinsons": "tomlinsons.com", "kriser": "krisers.com",
    "petsense": "petsense.com", "feeders": "feederspetsupply.com",
    "concord-pet": "concordpetfoods.com", "aldi": "aldi.us",
    "trader-joes": "traderjoes.com", "whole-foods": "wholefoodsmarket.com",
    "heb": "heb.com", "walgreens": "walgreens.com", "rite-aid": "riteaid.com",
    "dollar-tree": "dollartree.com", "big-lots": "biglots.com",
    "five-below": "fivebelow.com", "ace": "acehardware.com",
    "true-value": "truevalue.com", "do-it-best": "doitbest.com",
    "kehe": "kehe.com", "jeffers": "jefferspet.com", "bradley-caldwell": "bradleycaldwell.com",
    "hall-roberts": "hallroberts.com", "awg": "awginc.com", "cs-wholesale": "cswg.com",
    "mdi": "mdi-inc.com", "bozzutos": "bozzutos.com", "southern-states": "southernstates.com",
    "db-supply": "dbsupply.com", "big-y": "bigy.com", "market-basket": "shopmarketbasket.com",
    "price-chopper": "pricechopper.com", "tops": "topsmarkets.com",
    "southeastern": "winndixie.com", "lowes-foods": "lowesfoods.com",
    "ingles": "ingles-markets.com", "weis": "weismarkets.com",
    "woodmans": "woodmans-food.com", "redners": "rednersmarkets.com",
    "fresh-market": "thefreshmarket.com", "sprouts": "sprouts.com",
    "wegmans": "wegmans.com", "giant-eagle": "gianteagle.com", "winco": "wincofoods.com",
    "publix": "publix.com", "meijer": "meijer.com", "hyvee": "hy-vee.com",
    "pet-food-experts": "petfoodexperts.com", "phillips-pet": "phillipspet.com",
    "unfi": "unfi.com", "walmart-canada": "walmart.ca", "menards": "menards.com",
    "rural-king": "ruralking.com", "bomgaars": "bomgaars.com", "atwoods": "atwoods.com",
    "fleet-farm": "fleetfarm.com", "farm-fleet": "farmandfleet.com",
    "murdochs": "murdochs.com", "cal-ranch": "calranch.com", "big-r": "bigronline.com",
    "food-lion": "foodlion.com", "stop-shop": "stopandshop.com",
    "giant-company": "giantfoodstores.com", "hannaford": "hannaford.com",
    "giant-food": "giantfood.com", "albertsons": "albertsons.com", "kroger": "kroger.com",
    "tractor-supply": "tractorsupply.com", "petsmart": "petsmart.com", "petco": "petco.com",
    "target": "target.com", "costco": "costco.com", "sams-club": "samsclub.com",
    "bjs": "bjs.com", "dollar-general": "dollargeneral.com", "family-dollar": "familydollar.com",
    "home-depot": "homedepot.com", "lowes": "lowes.com", "cvs": "cvs.com",
    "hollywood-feed": "hollywoodfeed.com", "walmart": "walmart.com",
}


def api(path, payload):
    req = urllib.request.Request(
        f"{API}/{path}", method="POST", data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json", "X-KEY": KEY})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        try:
            return json.loads(body)
        except json.JSONDecodeError:
            return {"error": True, "error_code": f"HTTP {e.code}", "body": body[:200]}


def sheet_rows():
    """Retailers straight from the pipeline, so 'has a contact' is the live answer."""
    if not RETAILERS_SYNC_URL:
        sys.exit("RETAILERS_SYNC_URL is not set (credentials.py or env)")

    class _NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, *a, **k):
            return None

    opener = urllib.request.build_opener(_NoRedirect)
    try:
        with opener.open(RETAILERS_SYNC_URL + "?action=rows", timeout=90) as r:
            body = r.read().decode()
    except urllib.error.HTTPError as e:
        loc = e.headers.get("Location")
        with urllib.request.urlopen(loc, timeout=90) as r2:
            body = r2.read().decode()
    return json.loads(body).get("rows") or []


def needs_contact(row):
    try:
        return not json.loads(row.get("contacts_json") or "[]")
    except json.JSONDecodeError:
        return True


def search(domain, page=1, tries=4):
    """One page of people at a domain. Retries the free plan's rate limit."""
    for attempt in range(tries):
        out = _search_once(domain, page)
        if str(out.get("error_code", "")).lower().startswith("rate limit"):
            wait = 5 * (attempt + 1)
            print(f"    rate limited, waiting {wait}s")
            time.sleep(wait)
            continue
        return out
    return out


def _search_once(domain, page=1):
    return api("search-person", {
        "page": page,
        "filters": {
            "company": {"websites": {"include": [domain]}},
            "person_job_title": {"include": TITLE_TERMS},
        },
    })


def title_score(title):
    t = (title or "").lower()
    for i, term in enumerate(TITLE_TERMS):
        if all(w in t for w in term.split()):
            return len(TITLE_TERMS) - i
    return 0


def probe():
    print("key loaded:", bool(KEY), "| length:", len(KEY))
    acct = api("account-information", {})
    print("account-information:", json.dumps(acct)[:300])
    print("\nprobing the filter shape on petco.com (no enrichment)...")
    out = search("petco.com")
    if out.get("error"):
        print("  error:", json.dumps(out)[:400])
        return
    results = out.get("results") or []
    print(f"  {len(results)} result(s), free={out.get('free')}, "
          f"pagination={json.dumps(out.get('pagination') or {})}")
    for r in results[:5]:
        p = r.get("person") or {}
        print("   ·", p.get("full_name"), "|", p.get("current_job_title"),
              "| id:", str(p.get("person_id"))[:14])


def run(scope, enrich, limit_per_retailer, max_retailers):
    rows = [r for r in sheet_rows() if needs_contact(r)]
    if scope == "target":
        rows = [r for r in rows if r.get("status") in ("target", "pitched")]
    rows.sort(key=lambda r: (r.get("priority") or 99, -(r.get("us_stores") or 0)))
    rows = rows[:max_retailers]
    print(f"{len(rows)} retailer(s) with no contact in scope\n")

    found, skipped = [], []
    for row in rows:
        domain = DOMAINS.get(row["id"])
        if not domain:
            skipped.append(row["name"])
            continue
        out = search(domain)
        if out.get("error"):
            print(f"  {row['name']:34s} {out.get('error_code')}")
            if out.get("error_code") == "INSUFFICIENT_CREDITS":
                break
            continue
        people = sorted(
            ({"person": (r.get("person") or {}), "company": (r.get("company") or {})}
             for r in out.get("results") or []),
            key=lambda r: -title_score(r["person"].get("current_job_title")))
        picks = [p for p in people if title_score(p["person"].get("current_job_title"))][:limit_per_retailer]
        print(f"  {row['name']:34s} {domain:26s} {len(picks)} candidate(s)")
        for p in picks:
            person = p["person"]
            rec = {
                "retailer": row["name"], "retailer_id": row["id"],
                "status": row.get("status", ""), "rep_firm": row.get("rep_firm", ""),
                "domain": domain,
                "name": person.get("full_name") or " ".join(
                    filter(None, [person.get("first_name"), person.get("last_name")])),
                "title": person.get("current_job_title", ""),
                "linkedin": person.get("linkedin_url", ""),
                "location": person.get("location", ""),
                "email": "", "email_status": "not enriched",
                "person_id": person.get("person_id") or "",
            }
            found.append(rec)
        time.sleep(2.5)

    if enrich:
        print("\nenriching emails...")
        for rec in found:
            if not rec["person_id"]:
                rec["email_status"] = "no person_id"
                continue
            out = api("enrich-person", {"person_id": rec["person_id"], "email": True})
            if out.get("error"):
                rec["email_status"] = out.get("error_code", "error")
                continue
            person = (out.get("response") or out).get("person") or out.get("response") or {}
            email = person.get("email") or {}
            rec["email"] = email.get("email") if isinstance(email, dict) else (email or "")
            rec["email_status"] = (email.get("status") if isinstance(email, dict) else "") or "found"
            print(f"  {rec['retailer']:28s} {rec['name']:26s} {rec['email'] or rec['email_status']}")
            time.sleep(2.5)

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    cols = ["retailer", "retailer_id", "status", "rep_firm", "domain", "name", "title",
            "email", "email_status", "linkedin", "location", "person_id"]
    with open(OUT_PATH, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(found)
    print(f"\nwrote {len(found)} candidate(s) to {OUT_PATH}")
    if skipped:
        print("no domain on file (add to DOMAINS): " + ", ".join(skipped))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--probe", action="store_true", help="check the key and filter shape, no credits")
    ap.add_argument("--scope", choices=["target", "all"], default="target")
    ap.add_argument("--enrich", action="store_true", help="reveal emails (spends credits)")
    ap.add_argument("--per-retailer", type=int, default=3)
    ap.add_argument("--max-retailers", type=int, default=15)
    a = ap.parse_args()
    if not KEY:
        sys.exit("No Prospeo key: add PROSPEO_KEY to credentials.py")
    if a.probe:
        probe()
    else:
        run(a.scope, a.enrich, a.per_retailer, a.max_retailers)
