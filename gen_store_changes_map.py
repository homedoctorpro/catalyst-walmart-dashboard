"""
Generate `store_changes_map.html` — a standalone Leaflet map of Walmart store
distribution CHANGE from the first store-week to the latest, with trait status
factored in.

Reads the embedded `const DATA = {...}` from dashboard.html (so run it after
extract_data.py). Categorizes every store that was in the start feed and/or the
end feed:

  retained_active    in start, in end, still traited      (kept)
  lost_trait         in start, in end, but de-traited      (lost the listing)
  dropped_feed       in start, gone from end feed          (lost entirely)
  gained_active      new since start, traited              (real new distribution)
  gained_nontraited  new in feed since start, not traited  (showing but unauthorized)

Trait status comes from the current traited snapshot (DATA.traited); we have no
historical trait flag, so start stores are assumed active at launch.
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).parent

CATS = {
    "gained_active":     {"label": "Gained — authorized (traited)", "color": "#2e7d32", "r": 7, "op": 0.9,  "on": True},
    "gained_nontraited": {"label": "Gained — in feed, not traited", "color": "#26a69a", "r": 6, "op": 0.75, "on": True},
    "lost_trait":        {"label": "Lost authorization (de-traited)", "color": "#e67300", "r": 7, "op": 0.9, "on": True},
    "dropped_feed":      {"label": "Dropped from feed entirely",     "color": "#d32f2f", "r": 8, "op": 0.95, "on": True},
    "retained_active":   {"label": "Retained (active)",             "color": "#9e9e9e", "r": 3, "op": 0.30, "on": False},
}
ORDER = ["gained_active", "gained_nontraited", "lost_trait", "dropped_feed", "retained_active"]


def load_dashboard_data() -> dict:
    html = (ROOT / "dashboard.html").read_text(encoding="utf-8")
    i = html.index("const DATA = ") + len("const DATA = ")
    return json.JSONDecoder().raw_decode(html, i)[0]


def categorize(DATA: dict):
    ws = DATA["store_weeks"]
    start_w, end_w = ws[0], ws[-1]
    wsd = DATA["weekly_stores"]
    S, E = set(wsd[start_w]), set(wsd[end_w])
    traited = set()
    for lst in DATA.get("traited", {}).get("by_sku", {}).values():
        traited.update(lst)

    out = []
    counts = {k: 0 for k in CATS}
    for sn in S | E:
        ins, ine, ist = sn in S, sn in E, sn in traited
        if ins and ine and ist:        c = "retained_active"
        elif ins and ine and not ist:  c = "lost_trait"
        elif ins and not ine:          c = "dropped_feed"
        elif not ins and ine and ist:  c = "gained_active"
        else:                          c = "gained_nontraited"
        info = DATA["stores"].get(sn, {})
        lat, lon = info.get("lat"), info.get("lon")
        if lat is None or lon is None:
            continue
        counts[c] += 1
        out.append({
            "sn": sn, "lat": lat, "lon": lon,
            "city": (info.get("city") or "").title(), "state": info.get("state") or "",
            "cat": c,
        })
    meta = {
        "start_week": start_w, "end_week": end_w,
        "start_active": len(S), "end_feed": len(E), "end_active": len(E & traited),
        "snapshot_date": DATA.get("traited", {}).get("snapshot_date"),
    }
    return out, counts, meta


def build_html(points, counts, meta) -> str:
    payload = {"points": points, "counts": counts, "meta": meta, "cats": CATS, "order": ORDER}
    net_raw = meta["end_feed"] - meta["start_active"]
    net_adj = meta["end_active"] - meta["start_active"]
    sub = (f"{meta['start_week']} → {meta['end_week']} &nbsp;|&nbsp; "
           f"net feed {net_raw:+,} &nbsp;|&nbsp; net active {net_adj:+,}"
           f"{' &nbsp;|&nbsp; trait snapshot ' + meta['snapshot_date'] if meta.get('snapshot_date') else ''}")
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Catalyst Pet – Walmart Distribution Change Map</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css"/>
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css"/>
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
<style>
  *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
  html,body{{height:100%}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f0f2f5;color:#1a1a2e;display:flex;flex-direction:column}}
  header{{background:linear-gradient(90deg,#1a1a2e 0%,#16213e 100%);color:#fff;padding:13px 24px;box-shadow:0 2px 8px rgba(0,0,0,.3);flex-shrink:0}}
  header h1{{font-size:1.12rem;font-weight:700}}
  header span{{font-size:0.78rem;opacity:.7}}
  #map{{flex:1;width:100%;z-index:0}}
  .legend{{background:#fff;padding:12px 14px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.2);min-width:230px;font-family:inherit}}
  .legend h4{{font-size:0.82rem;font-weight:700;color:#333;margin-bottom:8px}}
  .legend-item{{display:flex;align-items:center;gap:8px;padding:5px 6px;border-radius:4px;cursor:pointer;font-size:0.8rem;user-select:none}}
  .legend-item:hover{{background:#f0f4ff}}
  .legend-item.off{{opacity:.4}}
  .legend-dot{{width:14px;height:14px;border-radius:50%;border:1.5px solid #333;flex-shrink:0}}
  .info{{background:#fff;padding:8px 12px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.15);font-size:0.8rem;line-height:1.5;max-width:240px}}
</style>
</head>
<body>
<header>
  <h1>\U0001F4CD Catalyst Pet – Walmart Distribution Change</h1>
  <span>{sub}</span>
</header>
<div id="map"></div>
<script>
const D = {json.dumps(payload, separators=(",", ":"))};

const map = L.map("map", {{ center: [39.5, -98.35], zoom: 5 }});
L.tileLayer("https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png", {{
  maxZoom: 18, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
}}).addTo(map);

const groups = {{}};
D.order.forEach(cat => {{
  const c = D.cats[cat];
  groups[cat] = L.markerClusterGroup({{ maxClusterRadius: 45, disableClusteringAtZoom: 9, showCoverageOnHover: false }});
}});
const LBL = D.cats;
D.points.forEach(p => {{
  const c = D.cats[p.cat];
  const popup = `<strong>Store #${{p.sn}}</strong><br>${{p.city}}, ${{p.state}}<br><span style="color:${{c.color}};font-weight:700">${{c.label}}</span>`;
  L.circleMarker([p.lat, p.lon], {{ radius: c.r, fillColor: c.color, color: "#222", weight: 1, opacity: 1, fillOpacity: c.op }})
    .bindPopup(popup).addTo(groups[p.cat]);
}});
D.order.forEach(cat => {{ if (D.cats[cat].on) map.addLayer(groups[cat]); }});

const info = L.control({{ position: "topright" }});
info.onAdd = () => {{
  const d = L.DomUtil.create("div", "info");
  const m = D.meta;
  d.innerHTML = `<strong>Start (${{m.start_week}}):</strong> ${{m.start_active.toLocaleString()}} stores<br>`
    + `<strong>End (${{m.end_week}}):</strong> ${{m.end_feed.toLocaleString()}} in feed<br>`
    + `<strong>End active (traited):</strong> ${{m.end_active.toLocaleString()}}<br>`
    + `<span style="color:#2e7d32;font-weight:700">Net active ${{(m.end_active-m.start_active>=0?'+':'')}}${{(m.end_active-m.start_active).toLocaleString()}}</span>`;
  return d;
}};
info.addTo(map);

const legend = L.control({{ position: "bottomleft" }});
legend.onAdd = () => {{
  const d = L.DomUtil.create("div", "legend");
  L.DomEvent.disableClickPropagation(d);
  d.innerHTML = `<h4>Distribution change <span style="font-weight:400;font-size:0.74rem">(click to toggle)</span></h4>`
    + D.order.map(cat => {{
        const c = D.cats[cat];
        return `<div class="legend-item${{c.on?'':' off'}}" data-cat="${{cat}}">
          <span class="legend-dot" style="background:${{c.color}}"></span>
          <span><strong>${{(D.counts[cat]||0).toLocaleString()}}</strong> &nbsp;${{c.label}}</span></div>`;
      }}).join("");
  d.querySelectorAll(".legend-item").forEach(el => el.addEventListener("click", () => {{
    const cat = el.dataset.cat;
    if (map.hasLayer(groups[cat])) {{ map.removeLayer(groups[cat]); el.classList.add("off"); }}
    else {{ map.addLayer(groups[cat]); el.classList.remove("off"); }}
  }}));
  return d;
}};
legend.addTo(map);
</script>
</body>
</html>
"""


def main():
    DATA = load_dashboard_data()
    points, counts, meta = categorize(DATA)
    html = build_html(points, counts, meta)
    out = ROOT / "store_changes_map.html"
    out.write_text(html, encoding="utf-8")
    print(f"Wrote {out.name}: {len(points):,} stores plotted")
    for k in ORDER:
        print(f"  {k:18s}: {counts[k]:>5,}")
    print(f"  net feed {meta['end_feed']-meta['start_active']:+,} | "
          f"net active {meta['end_active']-meta['start_active']:+,}")


if __name__ == "__main__":
    main()
