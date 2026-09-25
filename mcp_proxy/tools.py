"""
The MCP tools themselves.

They live here rather than in Apps Script because a JSON-RPC round trip through
a web app's doPost is slow (20-55s under the write lock) and occasionally comes
back with the wrong payload. Reads now use a plain GET (?action=rows), writes
the same upsert the dashboard uses, and both are reliable.
"""

import json
import os
import time
import urllib.error
import urllib.request

UPSTREAM_TIMEOUT = 90
CACHE_TTL = 20            # seconds; a conversation asks several questions in a row
_CACHE = {}

HERE = os.path.dirname(os.path.abspath(__file__))
# Published by every dashboard build; the file baked into the image is only the
# fallback for when GitHub Pages is unreachable.
CATALOG_URL = os.environ.get(
    "CATALOG_URL",
    "https://homedoctorpro.github.io/catalyst-walmart-dashboard/retailers_catalog.json")
CATALOG_TTL = 900
_catalog_cache = {"at": 0, "data": {}}

try:
    with open(os.path.join(HERE, "catalog.json"), encoding="utf-8") as f:
        BAKED_CATALOG = json.load(f)    # id -> name, channel, us_stores
except (FileNotFoundError, json.JSONDecodeError):
    BAKED_CATALOG = {}


def catalog():
    """The dashboard's retailer list, refreshed from Pages every 15 minutes."""
    if _catalog_cache["data"] and time.time() - _catalog_cache["at"] < CATALOG_TTL:
        return _catalog_cache["data"]
    try:
        with urllib.request.urlopen(CATALOG_URL, timeout=30) as r:
            data = json.loads(r.read().decode("utf-8")).get("retailers") or {}
        if data:
            _catalog_cache.update(at=time.time(), data=data)
            return data
    except Exception as e:
        print(f"[catalog] fetch failed, using the baked copy: {e}", flush=True)
    _catalog_cache.update(at=time.time(), data=BAKED_CATALOG)
    return BAKED_CATALOG


def _today():
    from datetime import date
    return date.today().isoformat()

STATUSES = ["in", "pitched", "target", "non-target", "declined", ""]
REP_FIRMS = ["Unassigned", "Brian Schlager", "Internal", "Jeff Day", "PSE", "StoR"]
# tool argument → the field name the sheet's upsert expects
WRITABLE = {
    "status": "status", "priority": "priority", "rep_firm": "repFirm",
    "next_steps": "nextSteps", "deadline": "deadline", "next_review": "nextReview",
    "reset_date": "resetDate", "needs_distributor": "needsDistributor",
    "distributor_name": "distributorName", "key_contact": "keyContact",
    "suggested_contacts": "suggestedContacts",
}
DATE_FIELDS = {"deadline", "next_review", "reset_date"}


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """Apps Script parks the real body behind a 302; follow it by hand."""
    def redirect_request(self, *args, **kwargs):
        return None


_opener = urllib.request.build_opener(_NoRedirect)


def _fetch(url, payload=None):
    req = urllib.request.Request(
        url,
        method="POST" if payload is not None else "GET",
        data=json.dumps(payload).encode() if payload is not None else None,
        headers={"Content-Type": "application/json"} if payload is not None else {},
    )
    try:
        with _opener.open(req, timeout=UPSTREAM_TIMEOUT) as r:
            body = r.read().decode("utf-8").strip()
    except urllib.error.HTTPError as e:
        location = e.headers.get("Location") if e.headers else None
        if e.code in (301, 302, 303, 307, 308) and location:
            with urllib.request.urlopen(location, timeout=UPSTREAM_TIMEOUT) as r2:
                body = r2.read().decode("utf-8").strip()
        else:
            raise
    return json.loads(body) if body else {}


def _retry(fn, attempts=3):
    last = None
    for i in range(attempts):
        try:
            return fn()
        except Exception as e:
            last = e
            print(f"[upstream] attempt {i + 1}: {e}", flush=True)
    raise last


class Pipeline:
    def __init__(self, base_url):
        self.base = base_url

    # ── data ──────────────────────────────────────────────────────────────
    def rows(self):
        """Prefer the sheet's own row feed; fall back to the overrides feed plus
        the catalog baked into this image, so a not-yet-redeployed script still
        works (names, channels and door counts then come from the dashboard)."""
        cached = _CACHE.get(self.base)
        if cached and time.time() - cached[0] < CACHE_TTL:
            return cached[1], cached[2]

        try:                                    # one probe only — it is missing
            out = _fetch(self.base + "?action=rows")   # until the script is redeployed
        except Exception as e:
            print(f"[upstream] rows probe: {e}", flush=True)
            out = {}
        if out.get("ok") and out.get("rows"):
            rows = self._seed_missing(out["rows"])
            today = out.get("today", _today())
            _CACHE[self.base] = (time.time(), rows, today)
            return rows, today

        out = _retry(lambda: _fetch(self.base + "?action=get"))
        if not out.get("ok"):
            raise RuntimeError(out.get("error") or "sheet returned ok:false")
        data, rows = out.get("data") or {}, []
        for rid, meta in catalog().items():
            v = data.get(rid, {})
            rows.append({
                "id": rid, "name": meta["name"], "channel": meta["channel"],
                "us_stores": meta["us_stores"],
                "status": v.get("status", ""), "priority": v.get("priority"),
                "rep_firm": v.get("repFirm", ""), "next_steps": v.get("nextSteps", ""),
                "deadline": v.get("deadline", ""), "next_review": v.get("nextReview", ""),
                "reset_date": v.get("resetDate", ""),
                "needs_distributor": bool(v.get("needsDistributor")),
                "distributor_name": v.get("distributorName", ""),
                "key_contact": v.get("keyContact", ""),
                "suggested_contacts": v.get("suggestedContacts", ""),
                "sf_account": v.get("sfAccountName", ""),
                "sf_opportunity_stage": v.get("sfOppStage", ""),
                "contacts_json": json.dumps(v.get("contacts") or []),
            })
        today = _today()
        _CACHE[self.base] = (time.time(), rows, today)
        return rows, today

    def _seed_missing(self, rows):
        """Add rows for retailers the dashboard has and the sheet doesn't.

        Keeps the sheet, Salesforce and this connector in step with the
        dashboard without anyone running "Push ALL retailers to Sheet".
        """
        known = {r["id"] for r in rows}
        missing = {rid: m for rid, m in catalog().items() if rid not in known}
        if not missing:
            return rows
        print(f"[seed] adding {len(missing)} retailer(s) the sheet lacks: "
              + ", ".join(sorted(missing)), flush=True)
        for rid, meta in missing.items():
            try:
                _retry(lambda: _fetch(self.base, {"action": "upsert", "item": {
                    "retailerId": rid, "retailerName": meta["name"],
                    "channel": meta["channel"], "usStores": meta["us_stores"],
                    "fields": {},
                }}), attempts=2)
                rows.append({"id": rid, "name": meta["name"], "channel": meta["channel"],
                             "us_stores": meta["us_stores"], "status": "", "priority": None,
                             "rep_firm": "", "next_steps": "", "deadline": "", "next_review": "",
                             "reset_date": "", "needs_distributor": False, "distributor_name": "",
                             "key_contact": "", "sf_account": "", "sf_opportunity_stage": "",
                             "contacts_json": "[]"})
            except Exception as e:
                print(f"[seed] {rid} failed: {e}", flush=True)
        return rows

    def _write(self, retailer_id, fields, us_stores=None):
        item = {"retailerId": retailer_id, "fields": fields}
        if us_stores is not None:
            item["usStores"] = us_stores
        out = _retry(lambda: _fetch(self.base, {"action": "upsert", "item": item}))
        if not out.get("ok"):
            raise RuntimeError(out.get("error") or "write refused")
        _CACHE.pop(self.base, None)             # the next read must see the change

    # ── helpers ───────────────────────────────────────────────────────────
    @staticmethod
    def _match(rows, term):
        t = str(term or "").strip().lower()
        if not t:
            return []
        exact = [r for r in rows if r["id"] == t or r["name"].lower() == t]
        return exact or [r for r in rows if t in r["id"] or t in r["name"].lower()]

    @staticmethod
    def _line(r, today=""):
        bits = [f'{r["name"]} ({r["id"]})', r["channel"]]
        if r.get("us_stores"):
            bits.append(f'{r["us_stores"]:,} doors')
        bits.append("status: " + (r.get("status") or "—"))
        if r.get("priority"):
            bits.append(f'priority {r["priority"]}')
        if r.get("rep_firm"):
            bits.append("rep: " + r["rep_firm"])
        if r.get("deadline"):
            bits.append("deadline " + r["deadline"] + (" (PASSED)" if today and r["deadline"] < today else ""))
        if r.get("next_review"):
            bits.append("review " + r["next_review"])
        if r.get("reset_date"):
            bits.append("reset " + r["reset_date"])
        if r.get("needs_distributor"):
            bits.append("needs distributor" + (": " + r["distributor_name"] if r.get("distributor_name") else ""))
        if r.get("sf_account"):
            bits.append("SF: " + r["sf_account"] + (" · opp " + r["sf_opportunity_stage"] if r.get("sf_opportunity_stage") else ""))
        if r.get("suggested_contacts"):
            first = r["suggested_contacts"].split("\n")[0]
            bits.append("suggested contact (unverified): " + first)
        if r.get("next_steps"):
            bits.append("next steps: " + r["next_steps"])
        return "- " + " · ".join(bits)

    # ── tools ─────────────────────────────────────────────────────────────
    def find_retailers(self, args):
        rows, today = self.rows()
        out = rows
        if args.get("query"):
            out = self._match(out, args["query"])
        if args.get("status"):
            out = [r for r in out if r.get("status") == str(args["status"]).lower()]
        if args.get("rep_firm"):
            out = [r for r in out if (r.get("rep_firm") or "").lower() == str(args["rep_firm"]).lower()]
        if args.get("channel"):
            out = [r for r in out if str(args["channel"]).lower() in (r.get("channel") or "").lower()]
        if args.get("overdue"):
            out = [r for r in out if r.get("deadline") and r["deadline"] < today]
        limit = int(args.get("limit") or 40)
        head = f"{len(out)} retailer(s)" + (f", showing {limit}" if len(out) > limit else "")
        return head + "\n" + "\n".join(self._line(r, today) for r in out[:limit])

    def get_retailer(self, args):
        rows, today = self.rows()
        hits = self._match(rows, args.get("retailer"))
        if not hits:
            return f'No retailer matches "{args.get("retailer")}".'
        if len(hits) > 1:
            return f'That matches {len(hits)}: ' + ", ".join(r["name"] for r in hits) + ". Be more specific."
        r = hits[0]
        out = self._line(r, today)
        try:
            contacts = json.loads(r.get("contacts_json") or "[]")
        except json.JSONDecodeError:
            contacts = []
        if contacts:
            out += "\nSalesforce contacts:"
            for c in contacts:
                star = "  ★ key contact" if r.get("key_contact") and c.get("email") == r["key_contact"] else ""
                out += (f'\n  · {c.get("name","")}' + (f' — {c["title"]}' if c.get("title") else "") +
                        (f' · {c["email"]}' if c.get("email") else "") +
                        (f' · {c["phone"]}' if c.get("phone") else "") + star)
        elif r.get("sf_account"):
            out += "\nNo Salesforce contacts on the linked account."
        return out

    def update_retailer(self, args):
        rows, _ = self.rows()
        hits = self._match(rows, args.get("retailer"))
        if not hits:
            return f'No retailer matches "{args.get("retailer")}". Nothing changed.'
        if len(hits) > 1:
            return (f'That matches {len(hits)}: ' + ", ".join(r["name"] for r in hits) +
                    ". Nothing changed — name one.")
        r = hits[0]

        # start from what the row already holds, so an upsert can't blank a field
        fields = {}
        for arg, key in WRITABLE.items():
            v = r.get(arg)
            if arg == "needs_distributor":
                v = "1" if v else ""
            fields[key] = "" if v is None else v

        changed = []
        for arg, key in WRITABLE.items():
            if arg not in args:
                continue
            v = args[arg]
            if arg == "status":
                v = str(v or "").lower()
                if v not in STATUSES:
                    return "Status must be one of: " + ", ".join(s for s in STATUSES if s) + ". Nothing changed."
            elif arg == "priority":
                v = int(v) if v and 1 <= int(v) <= 10 else ""
            elif arg == "needs_distributor":
                v = "1" if v else ""
            elif arg == "rep_firm":
                if v and v not in REP_FIRMS:
                    return "Rep firm must be one of: " + ", ".join(REP_FIRMS) + ". Nothing changed."
            elif arg in DATE_FIELDS and v:
                v = str(v).strip()
                if len(v) != 10 or v[4] != "-" or v[7] != "-":
                    return f"{arg} must look like yyyy-mm-dd. Nothing changed."
            before = fields[key]
            if str(before) == str(v):
                continue
            fields[key] = v
            changed.append(f'{arg}: {before if before != "" else "—"} → {v if v != "" else "—"}')

        stores = None
        if "us_stores" in args and int(args["us_stores"]) >= 0 and int(args["us_stores"]) != r.get("us_stores"):
            stores = int(args["us_stores"])
            changed.append(f'us_stores: {r.get("us_stores")} → {stores} '
                           "(a dashboard rebuild resets this to the count in the dashboard code)")

        if not changed:
            return f'Nothing to change on {r["name"]}.'
        self._write(r["id"], fields, stores)
        return f'Updated {r["name"]}:\n  ' + "\n  ".join(changed)

    def pipeline_summary(self, args=None):
        rows, today = self.rows()
        by_status, by_rep = {}, {}
        overdue, upcoming = [], []
        for r in rows:
            by_status[r.get("status") or "(none)"] = by_status.get(r.get("status") or "(none)", 0) + 1
            by_rep[r.get("rep_firm") or "Unassigned"] = by_rep.get(r.get("rep_firm") or "Unassigned", 0) + 1
            if r.get("deadline") and r["deadline"] < today:
                overdue.append(r)
            if r.get("next_review") and r["next_review"] >= today:
                upcoming.append(r)
        upcoming.sort(key=lambda r: r["next_review"])
        return (
            f"Retailers: {len(rows)}\n"
            "By status: " + " · ".join(f"{k} {v}" for k, v in by_status.items()) + "\n"
            "By rep firm: " + " · ".join(f"{k} {v}" for k, v in by_rep.items()) + "\n"
            "Deadlines passed: " + (", ".join(f'{r["name"]} ({r["deadline"]})' for r in overdue) or "none") + "\n"
            "Next reviews: " + (", ".join(f'{r["name"]} {r["next_review"]}' for r in upcoming[:5]) or "none set")
        )


TOOLS = [
    {
        "name": "find_retailers",
        "title": "Find retailers",
        "description": "Search the Catalyst retailer pipeline. Filters combine; omit them all for the "
                       "whole list. Each result shows status, priority, rep firm, next steps, dates, "
                       "doors, channel and the Salesforce link.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": 'Part of a retailer name, e.g. "petco"'},
                "status": {"type": "string", "enum": [s for s in STATUSES if s]},
                "rep_firm": {"type": "string", "description": "e.g. Brian Schlager, PSE, Internal, StoR"},
                "channel": {"type": "string", "description": 'e.g. "Grocery", "Pet Specialty"'},
                "overdue": {"type": "boolean", "description": "Only rows whose deadline has passed"},
                "limit": {"type": "number", "description": "Max rows (default 40)"},
            },
        },
    },
    {
        "name": "get_retailer",
        "title": "Get one retailer",
        "description": "Everything on one retailer, including Salesforce contacts and its cat-litter "
                       "opportunity stage.",
        "inputSchema": {
            "type": "object",
            "properties": {"retailer": {"type": "string", "description": "Retailer name or id"}},
            "required": ["retailer"],
        },
    },
    {
        "name": "update_retailer",
        "title": "Update a retailer",
        "description": "Change one retailer's planning fields; only the arguments passed are touched. "
                       "Target or Pitched puts the row into Salesforce and creates its cat-litter "
                       "opportunity. Currently In, Non-Target and Declined park it at the bottom of "
                       "the dashboard. Dates are yyyy-mm-dd.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "retailer": {"type": "string", "description": "Retailer name or id"},
                "status": {"type": "string", "enum": STATUSES},
                "priority": {"type": "number", "description": "1-10, 1 is highest; 0 clears it"},
                "rep_firm": {"type": "string", "enum": REP_FIRMS},
                "next_steps": {"type": "string"},
                "deadline": {"type": "string", "description": "yyyy-mm-dd, empty clears"},
                "next_review": {"type": "string", "description": "yyyy-mm-dd, empty clears"},
                "reset_date": {"type": "string", "description": "yyyy-mm-dd shelf reset, empty clears"},
                "needs_distributor": {"type": "boolean"},
                "distributor_name": {"type": "string"},
                "key_contact": {"type": "string", "description": "Email of the Salesforce contact to star"},
                "suggested_contacts": {"type": "string", "description": "Leads from a lookup tool, one per line; not Salesforce contacts"},
                "us_stores": {"type": "number", "description": "Door count (a dashboard rebuild resets it)"},
            },
            "required": ["retailer"],
        },
    },
    {
        "name": "pipeline_summary",
        "title": "Pipeline summary",
        "description": "Counts by status and rep firm, deadlines that have passed, and the next reviews.",
        "inputSchema": {"type": "object", "properties": {}},
    },
]


def call(base_url, name, args):
    p = Pipeline(base_url)
    fn = {
        "find_retailers": p.find_retailers,
        "get_retailer": p.get_retailer,
        "update_retailer": p.update_retailer,
        "pipeline_summary": p.pipeline_summary,
    }.get(name)
    if not fn:
        raise ValueError(f"unknown tool: {name}")
    return fn(args or {})
