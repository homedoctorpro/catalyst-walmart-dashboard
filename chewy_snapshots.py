"""
Chewy Brand Snapshot PDF parser
===============================
Parses any Chewy "Brand Snapshot" PDF into structured monthly data so new
months auto-ingest (no hardcoding). Used by chewy_extract.py.

Each snapshot PDF is one brand for one month. We pull:
  - brand ("Catalyst" / "Feline Fresh") and month ("YYYY-MM")
  - avg customer rating
  - top-10 states by units
  - per-SKU customer sales units (US lb SKUs keyed by Chewy part number;
    Canada kg SKUs keyed by size label)

Uses pymupdf (fitz) positional word extraction: each product/state row is read
by its y-position, taking the part number (leftmost digits) and the unit value
(rightmost number). Chewy prints each value twice at the same spot, so we dedupe
by position — no fragile text-stream munging.
"""
import os
import re
import glob

import fitz  # pymupdf

KG_LABEL = {"18.14": "40-lb (18.14 kg)", "9.07": "20-lb (9.07 kg)"}
MONTHS = {"January": "01", "February": "02", "March": "03", "April": "04",
          "May": "05", "June": "06", "July": "07", "August": "08",
          "September": "09", "October": "10", "November": "11",
          "December": "12"}
STATES = {"AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI",
          "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI",
          "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC",
          "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT",
          "VT", "VA", "WA", "WV", "WI", "WY", "DC"}


def _num(tok):
    return int(tok.replace(",", ""))


def _value_on_line(values, y, tol=4.0):
    """Rightmost numeric value whose token sits on ~the same baseline as y."""
    same = [(x, v) for (yy, x, v) in values if abs(yy - y) <= tol]
    return max(same)[1] if same else None


def _value_near(values, y, xmin=200.0, max_dy=25.0):
    """Value in the right column nearest the part-number baseline (product
    names may wrap, pushing the bar label a few lines down)."""
    cand = [(abs(yy - y), -x, v) for (yy, x, v) in values
            if x >= xmin and abs(yy - y) <= max_dy]
    return min(cand)[2] if cand else None


def parse_snapshot_pdf(path):
    """Return a dict of parsed fields, or None if it isn't a snapshot PDF."""
    doc = fitz.open(path)
    full = re.sub(r"\s+", " ",
                  " ".join(page.get_text() for page in doc))
    if "Customer Sales Units" not in full:
        return None

    mo = re.search(r"Month\s+(\d{2})/\d{2}/(\d{4})", full)
    mn = re.search(r"\b(January|February|March|April|May|June|July|August|"
                   r"September|October|November|December)\s+(\d{4})", full)
    if mo:
        month = f"{mo.group(2)}-{mo.group(1)}"
    elif mn:                                       # older format: "| March 2026"
        month = f"{mn.group(2)}-{MONTHS[mn.group(1)]}"
    else:
        # Jul-2026+ exports truncate the "Shipped Month" text ("07/01/20�");
        # fall back to the newest axis date on the Monthly Customer Sales
        # Units chart, which always ends at the snapshot month.
        dates = re.findall(r"\b(\d{4})-(\d{2})-01\b", full)
        if not dates:
            return None
        month = "-".join(max(dates))

    rm = re.search(r"([\d.]+)\s*Avg Custom", full)  # "Rating" may be truncated
    rating = float(rm.group(1)) if rm else None

    # Collect positioned tokens per page, then match each SKU/state to the
    # value that shares its baseline. Robust to wrapped product names and
    # layout shifts; newer exports put the per-SKU chart on page 2.
    sku_units, canada, top_states = {}, {}, []
    seen_state = set()
    for page in doc:
        words = page.get_text("words")
        # numeric unit-count tokens (exclude bare 6-7 digit part numbers)
        values = [(y0, x0, _num(t)) for x0, y0, x1, y1, t, *_ in words
                  if re.fullmatch(r"[\d,]+", t)
                  and not re.fullmatch(r"\d{6,7}", t)]
        for x0, y0, x1, y1, t, *_ in words:
            if re.fullmatch(r"\d{6,7}", t):            # US part number
                v = _value_near(values, y0)
                if v is not None and v < 100000:   # exclude a stray part-number
                    sku_units[t] = v
            elif t in STATES and t not in seen_state:  # top-10 state
                v = _value_on_line(values, y0)
                if v is not None:
                    top_states.append([t, v])
                    seen_state.add(t)
            else:                                      # Canada kg SKU
                km = re.fullmatch(r"([\d.]+)-kg", t)
                if km and km.group(1) in KG_LABEL:
                    v = _value_near(values, y0)
                    if v is not None:
                        canada[KG_LABEL[km.group(1)]] = v

    top_states = sorted(top_states, key=lambda s: -s[1])[:10]
    # Brand from the SKU part numbers (Feline Fresh parts all start 1932),
    # falling back to the title text — robust across PDF template vintages.
    if any(p.startswith("1932") for p in sku_units) or canada:
        brand = "Feline Fresh"
    elif sku_units:
        brand = "Catalyst"
    else:
        brand = "Feline Fresh" if "Feline Fresh" in full else "Catalyst"
    return {"brand": brand, "month": month, "rating": rating,
            "top_states": top_states, "sku_units": sku_units,
            "canada": canada, "path": os.path.basename(path)}


def parse_folder(folder):
    """Parse every Brand Snapshot PDF in a folder -> list of parsed dicts."""
    out = []
    for p in sorted(glob.glob(os.path.join(folder, "*.pdf"))):
        try:
            d = parse_snapshot_pdf(p)
        except Exception as e:                      # noqa: BLE001
            print(f"[WARN] could not parse {os.path.basename(p)}: {e}")
            continue
        if d:
            out.append(d)
    return out


if __name__ == "__main__":
    import sys
    folder = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "Chewy Brand Snapshots")
    for d in parse_folder(folder):
        print(f"\n{d['brand']:12} {d['month']}  rating={d['rating']}  "
              f"({d['path']})")
        print("  SKUs:", d["sku_units"])
        if d["canada"]:
            print("  Canada:", d["canada"])
        print("  states:", d["top_states"])
