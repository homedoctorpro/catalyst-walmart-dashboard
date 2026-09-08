"""
Endcap Set Map
==============
One dot per Catalyst store, in three groups the field program cares about:

  set        -- endcap stores a merchandiser confirmed set, one colour per set
                wave (WK27 sweep, Aug 9-11 follow-up, WK30 sweep, ...)
  notset     -- endcap stores visited but not set (popup carries the reason)
  other      -- every other store carrying Catalyst in the latest weekly feed

Set / not-set comes from build_endcap_report.load_survey (WK27 survey +
follow-up list + every re-sweep, same definition as the dashboard). Sales in
the popups are 15 lb Original units in the pre-endcap week (202626) and the
latest week on disk.

    python build_endcap_set_map.py      -> endcap_set_map.html

No password gate, same as endcap_rollout_map.html, so it can be shared.
"""
import csv
import json
import os

from build_endcap_report import (ENDCAP_WEEK, HERE, PRIOR_WEEK, REASONS, load_bystore,
                                 load_survey, wave_meta)

OUTPUT = os.path.join(HERE, "endcap_set_map.html")
GEO = os.path.join(HERE, "stores_geo.json")
NOTSET_COLOR, OTHER_COLOR, UNVISITED_COLOR = "#d73027", "#9aa5ad", "#7f8c8d"


def main():
    survey = load_survey()
    waves = wave_meta(survey)
    endcap = {}
    with open(os.path.join(HERE, "endcap_stores.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            endcap[int(r["store_number"])] = r
    q26, _, _, _, meta26 = load_bystore(PRIOR_WEEK)
    qcur, _, _, qtot, meta = load_bystore(ENDCAP_WEEK)
    geo = json.load(open(GEO, encoding="utf-8"))
    reason_label = {k: l for _a, (k, l, _c) in REASONS.items()}

    groups = [{"key": "set_" + w["key"], "label": "Set — " + w["label"], "color": w["color"],
               "kind": "set"} for w in waves]
    groups += [{"key": "notset", "label": "Endcap store, not set", "color": NOTSET_COLOR, "kind": "notset"},
               {"key": "unvisited", "label": "Endcap store, not visited", "color": UNVISITED_COLOR, "kind": "notset"},
               {"key": "other", "label": "Non-endcap Catalyst store", "color": OTHER_COLOR, "kind": "other"}]

    def latlon(store, row=None):
        if row and row.get("latitude") and row.get("longitude"):
            return float(row["latitude"]), float(row["longitude"])
        zip5 = (row or {}).get("zip") or meta.get(store, ("",))[0] or meta26.get(store, ("",))[0]
        g = geo.get(str(zip5)[:5])
        return (g["lat"], g["lon"]) if g else (None, None)

    stores, skipped = [], 0
    for s, r in endcap.items():
        lat, lon = latlon(s, r)
        if lat is None:
            skipped += 1
            continue
        v = survey.get(s)
        if v and v["set"]:
            key = "set_" + v["wave"]
        elif v:
            key = "notset"
        else:
            key = "unvisited"
        stores.append({
            "store": s, "city": (r["city"] or "").title(), "state": r["state"],
            "lat": round(lat, 4), "lon": round(lon, 4), "g": key,
            "date": (v or {}).get("date", ""),
            "where": (v or {}).get("where", ""),
            "reason": reason_label.get((v or {}).get("seg", ""), ""),
            "detail": (v or {}).get("detail", ""),
            "prior": reason_label.get(
                {a: k for a, (k, _l, _c) in REASONS.items()}.get((v or {}).get("prior_reason", "")), ""),
            "q26": round(q26.get(s, 0)), "qcur": round(qcur.get(s, 0)),
        })
    for s in qtot:
        if s in endcap:
            continue
        lat, lon = latlon(s)
        if lat is None:
            skipped += 1
            continue
        zip5, city, state = meta.get(s, ("", "", ""))
        stores.append({"store": s, "city": city, "state": state,
                       "lat": round(lat, 4), "lon": round(lon, 4), "g": "other",
                       "q26": round(q26.get(s, 0)), "qcur": round(qcur.get(s, 0))})

    counts = {g["key"]: sum(1 for x in stores if x["g"] == g["key"]) for g in groups}
    usw = {}
    for g in groups:
        sel = [x for x in stores if x["g"] == g["key"]]
        n = len(sel)
        usw[g["key"]] = {"pre": round(sum(x["q26"] for x in sel) / n, 2) if n else None,
                         "cur": round(sum(x["qcur"] for x in sel) / n, 2) if n else None}
    payload = {"week": ENDCAP_WEEK, "prior_week": PRIOR_WEEK, "n_endcap": len(endcap),
               "groups": groups, "counts": counts, "usw": usw, "stores": stores}
    html = TEMPLATE.replace("/*DATA*/", json.dumps(payload, separators=(",", ":")))
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[ok] {len(stores)} stores mapped, {skipped} without a location; "
          + ", ".join(f"{k}={v}" for k, v in counts.items()))
    print(f"[ok] wrote {OUTPUT}")


TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Catalyst Endcap Set Map</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
  html, body { margin:0; height:100%; font-family:system-ui,Segoe UI,Arial,sans-serif; }
  #map { position:absolute; inset:0; }
  .hdr { position:absolute; top:10px; left:50px; right:10px; z-index:1000;
         background:#fff; border-radius:8px; box-shadow:0 1px 6px rgba(0,0,0,.3);
         padding:8px 14px; display:inline-block; max-width:calc(100% - 80px); }
  .hdr h1 { margin:0; font-size:16px; }
  .hdr .sub { color:#666; font-size:12px; margin-top:2px; }
  .legend { position:absolute; bottom:20px; left:10px; z-index:1000; background:#fff;
            border-radius:8px; box-shadow:0 1px 6px rgba(0,0,0,.3); padding:10px 14px;
            font-size:13px; min-width:330px; }
  .legend table { border-collapse:collapse; width:100%; }
  .legend th { font-size:10px; color:#888; text-transform:uppercase; letter-spacing:.03em;
               font-weight:600; text-align:right; padding:0 0 4px 10px; }
  .legend th:first-child { text-align:left; padding-left:0; }
  .legend tr.row { cursor:pointer; user-select:none; }
  .legend tr.row td { padding:3px 0 3px 10px; text-align:right; white-space:nowrap; }
  .legend tr.row td:first-child { text-align:left; padding-left:0; }
  .legend tr.off { opacity:.35; }
  .legend .dot { display:inline-block; width:11px; height:11px; border-radius:50%;
                 margin-right:7px; vertical-align:-1px; }
  .legend .hint { color:#999; font-size:11px; margin-top:8px; }
  .popup b { font-size:13px; }
  .popup table { border-collapse:collapse; margin-top:4px; font-size:12px; }
  .popup td { padding:1px 8px 1px 0; vertical-align:top; }
</style>
</head>
<body>
<div id="map"></div>
<div class="hdr">
  <h1>Catalyst 15-lb Endcap &mdash; set, not set, and everyone else</h1>
  <div class="sub" id="sub"></div>
</div>
<div class="legend" id="legend"></div>
<script>
const DATA = /*DATA*/;
const fmt = n => n.toLocaleString();
const map = L.map('map').setView([39.5, -96.5], 4);
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
  { attribution:'&copy; OpenStreetMap contributors', maxZoom:18 }).addTo(map);

const nSet = DATA.groups.filter(g => g.kind === 'set').reduce((a, g) => a + DATA.counts[g.key], 0);
document.getElementById('sub').textContent =
  fmt(DATA.n_endcap) + ' endcap program stores \\u00b7 ' + fmt(nSet) + ' confirmed set \\u00b7 ' +
  fmt(DATA.counts.other) + ' other Catalyst stores \\u00b7 sales through wk ' + DATA.week;

const groups = {}, meta = {};
for (const g of DATA.groups) { groups[g.key] = L.layerGroup(); meta[g.key] = g; }
// draw the grey background first so program stores sit on top
for (const g of [...DATA.groups].reverse()) groups[g.key].addTo(map);

for (const s of DATA.stores) {
  const g = meta[s.g], bg = g.kind === 'other';
  const m = L.circleMarker([s.lat, s.lon], { radius: bg ? 3 : 5, color: '#333',
    weight: bg ? 0 : .5, fillColor: g.color, fillOpacity: bg ? .45 : .88 });
  let rows = '<tr><td>15-lb units wk ' + DATA.prior_week + ' (pre-endcap)</td><td>' + s.q26 + '</td></tr>' +
             '<tr><td>15-lb units wk ' + DATA.week + '</td><td><b>' + s.qcur + '</b></td></tr>';
  if (g.kind === 'set') {
    rows += '<tr><td>Set on</td><td>' + (s.date || '\\u2014') + (s.where ? ' \\u00b7 ' + s.where : '') + '</td></tr>';
    if (s.prior) rows += '<tr><td>Earlier visit said</td><td>' + s.prior + '</td></tr>';
  } else if (s.g === 'notset') {
    rows += '<tr><td>Reason (' + s.date + ')</td><td>' + (s.detail || s.reason || '\\u2014') + '</td></tr>';
  }
  m.bindPopup('<div class="popup"><b>Store #' + s.store + '</b> \\u2014 ' + s.city + ', ' + s.state +
    '<br><span style="color:' + g.color + ';font-weight:600">' + g.label + '</span>' +
    '<table>' + rows + '</table></div>');
  m.addTo(groups[s.g]);
}

const legend = document.getElementById('legend');
let html = '<table><tr><th>Group</th><th>Stores</th><th>U/S/W wk' + DATA.prior_week +
           '</th><th>U/S/W wk' + DATA.week + '</th></tr>';
for (const g of DATA.groups) {
  const n = DATA.counts[g.key], u = DATA.usw[g.key];
  const pct = g.kind === 'other' ? '' : ' <span style="color:#999">(' + Math.round(100 * n / DATA.n_endcap) + '%)</span>';
  html += '<tr class="row" data-key="' + g.key + '"><td><span class="dot" style="background:' + g.color +
    '"></span>' + g.label + '</td><td>' + fmt(n) + pct + '</td><td>' +
    (u.pre == null ? '\\u2014' : u.pre.toFixed(2)) + '</td><td><b>' +
    (u.cur == null ? '\\u2014' : u.cur.toFixed(2)) + '</b></td></tr>';
}
html += '</table><div class="hint">Click a row to show/hide \\u00b7 click a dot for detail \\u00b7 ' +
        'U/S/W = 15-lb Original units per store per week</div>';
legend.innerHTML = html;
legend.querySelectorAll('tr.row').forEach(row => {
  row.onclick = () => {
    const key = row.dataset.key, off = row.classList.toggle('off');
    off ? map.removeLayer(groups[key]) : map.addLayer(groups[key]);
  };
});
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
