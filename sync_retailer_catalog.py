"""Keep the Retailers Google Sheet's dashboard-owned columns in step with the dashboard.

Sends retailers_catalog.json (written by extract_data.py on every build) to the
Apps Script's `syncstatic` action. Only name, channel and store count change on
existing rows; missing retailers are appended. Edits and Salesforce links are
never touched, so this runs unattended from the retailer-catalog-sync workflow
whenever the catalog changes — nobody has to click "Push ALL retailers to Sheet".

    RETAILERS_SYNC_URL=... python sync_retailer_catalog.py
"""
import json
import os
import sys
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))


def build_items(catalog):
    return [{
        "retailerId":   rid,
        "retailerName": m["name"],
        "channel":      m["channel"],
        "usStores":     m["us_stores"],
        "defaultUsw":   m.get("default_usw"),
    } for rid, m in sorted(catalog.items())]


def post(url, payload, tries=3):
    body = json.dumps(payload).encode("utf-8")
    for i in range(tries):
        try:
            req = urllib.request.Request(url, data=body, method="POST",
                                         headers={"Content-Type": "text/plain;charset=utf-8"})
            with urllib.request.urlopen(req, timeout=180) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:  # Apps Script throws transient 5xx under load
            if i == tries - 1:
                raise
            print(f"  retry {i + 1}: {e}")
            time.sleep(10 * (i + 1))


def main():
    url = os.environ.get("RETAILERS_SYNC_URL", "").strip()
    if not url:
        print("RETAILERS_SYNC_URL not set — nothing to sync")
        return 1
    with open(os.path.join(HERE, "retailers_catalog.json"), encoding="utf-8") as f:
        catalog = json.load(f)["retailers"]
    out = post(url, {"action": "syncstatic", "items": build_items(catalog)})
    if not out.get("ok"):
        print(f"Sheet sync failed: {out.get('error')}")
        return 1
    changed, added = out.get("changed", []), out.get("added", [])
    print(f"Sheet synced: {len(catalog)} retailers, {len(changed)} updated, {len(added)} added")
    for line in changed:
        print(f"  updated {line}")
    for rid in added:
        print(f"  added   {rid}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
