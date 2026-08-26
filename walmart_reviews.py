"""
Collect Walmart verified-purchase reviews for Catalyst Pet items.

Walmart server-renders review data as JSON inside the page's __NEXT_DATA__ script
tag, so no browser or DOM parsing is needed -- fetch the HTML, pull one JSON blob.

Each product page rolls up both pack sizes; the per-review `features` list carries
Size ("15 lbs" / "34 lbs"), which is how the SKU split is recovered.

Output (all in Reviews/):
    reviews_raw_<itemid>.json   full API objects, one file per product
    catalyst_reviews.csv        flattened, both products
    catalyst_reviews.xlsx       same + monthly trend sheet

Usage:
    python walmart_reviews.py            # incremental: keeps existing, adds new
    python walmart_reviews.py --full     # refetch every page from scratch
"""

import argparse
import json
import os
import random
import re
import sys
import time
from datetime import datetime

import pandas as pd
import requests

# Both listings roll up their 15 lb and 34 lb variants.
PRODUCTS = {
    "17951853567": "Original",
    "17944250041": "Unscented",
}

OUTDIR = "Reviews"
BASE = "https://www.walmart.com/reviews/product/{item}"
NEXT_DATA_RE = re.compile(r'id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Paced to stay well inside normal human browsing speed. ~30 pages per product.
DELAY_RANGE = (2.5, 4.5)
MAX_RETRIES = 4
MAX_REPAIR_PASSES = 3


def fetch_page(session, item, page):
    """Return the `reviews` object from one review page, or None if unavailable."""
    url = BASE.format(item=item)
    params = {"vp": "true", "sort": "submission-desc", "page": page}
    for attempt in range(MAX_RETRIES):
        try:
            r = session.get(url, params=params, timeout=45)
            if r.status_code == 200:
                m = NEXT_DATA_RE.search(r.text)
                if not m:
                    # Served a page but without embedded data -- usually an
                    # interstitial. Back off hard rather than hammering.
                    print(f"    [warn] no __NEXT_DATA__ on page {page} "
                          f"(attempt {attempt + 1})")
                    time.sleep(15 * (attempt + 1))
                    continue
                data = json.loads(m.group(1))
                return data["props"]["pageProps"]["initialData"]["data"]["reviews"]
            if r.status_code in (403, 429, 503):
                wait = 20 * (attempt + 1)
                print(f"    [warn] HTTP {r.status_code} on page {page}, "
                      f"waiting {wait}s")
                time.sleep(wait)
                continue
            print(f"    [warn] HTTP {r.status_code} on page {page}")
            return None
        except (requests.RequestException, KeyError, json.JSONDecodeError) as e:
            print(f"    [warn] {type(e).__name__} on page {page}: {e}")
            time.sleep(10 * (attempt + 1))
    return None


def is_verified(review):
    return any(b.get("id") == "VerifiedPurchaser" for b in (review.get("badges") or []))


def _walk(session, item, n_pages, found, skip_first=None):
    """Fetch pages 1..n_pages, adding any unseen reviews to `found` (id -> review)."""
    for page in range(1, n_pages + 1):
        if page == 1 and skip_first is not None:
            batch = skip_first.get("customerReviews") or []
        else:
            time.sleep(random.uniform(*DELAY_RANGE))
            rv = fetch_page(session, item, page)
            if rv is None:
                print(f"    [warn] page {page} unavailable, continuing")
                continue
            batch = rv.get("customerReviews") or []
        for r in batch:
            # vp=true should make the verified check a no-op, but verify not trust.
            if is_verified(r):
                found.setdefault(r["reviewId"], r)
        if page % 5 == 0 or page == n_pages:
            print(f"    page {page}/{n_pages} -> {len(found)} unique")


def collect(item, label):
    """Walk every page of verified-purchase reviews for one item.

    Walmart paginates by offset over a list sorted newest-first, so a review
    submitted mid-walk shifts every subsequent row down one -- which both repeats
    a row at the page boundary and drops the row that slid past it. A single pass
    therefore reliably comes up a few short. Union by reviewId across repeat
    passes until the count matches the advertised total.
    """
    session = requests.Session()
    session.headers.update(HEADERS)

    first = fetch_page(session, item, 1)
    if first is None:
        print(f"  [FAIL] {label} ({item}): could not load page 1")
        return []

    total = first["pagination"]["total"]
    n_pages = -(-total // 10)  # ceil
    print(f"  {label} ({item}): {total} verified reviews across {n_pages} pages")
    print(f"    catalog total: {first.get('totalReviewCount')} ratings, "
          f"avg {first.get('roundedAverageOverallRating')}")

    found = {}
    _walk(session, item, n_pages, found, skip_first=first)

    # Repair passes for anything pagination drift skipped.
    for attempt in range(1, MAX_REPAIR_PASSES + 1):
        if len(found) >= total:
            break
        missing = total - len(found)
        print(f"    [repair {attempt}] {len(found)}/{total} "
              f"({missing} missing), re-walking")
        before = len(found)
        _walk(session, item, n_pages, found)
        if len(found) == before:
            print("    [repair] no new reviews found, stopping")
            break

    if len(found) < total:
        print(f"    [WARN] collected {len(found)} of {total} advertised "
              f"-- {total - len(found)} unrecovered")
    else:
        print(f"    [OK] {len(found)}/{total} complete")

    reviews = list(found.values())
    for r in reviews:
        r["_product"] = label
        r["_itemPageId"] = item
        r["_advertisedTotal"] = total
        # Catalog-wide figures include syndicated reviews; verified-only is a
        # subset, so keep both to show the gap on the dashboard.
        r["_catalogTotal"] = first.get("totalReviewCount")
        r["_catalogAvg"] = first.get("roundedAverageOverallRating")
    return reviews


def flatten(reviews):
    rows = []
    for r in reviews:
        size = None
        for f in r.get("features") or []:
            if f.get("name") == "Size":
                size = f.get("value")
        rows.append({
            "product": r["_product"],
            "size": size,
            "date": pd.to_datetime(r.get("reviewSubmissionTime"), errors="coerce"),
            "rating": r.get("rating"),
            "title": r.get("reviewTitle"),
            "text": r.get("reviewText"),
            "nickname": r.get("userNickname"),
            "helpful_yes": r.get("positiveFeedback"),
            "helpful_no": r.get("negativeFeedback"),
            "photos": len(r.get("photos") or []),
            "fulfilled_by": r.get("fulfilledBy"),
            "seller": r.get("sellerName"),
            "syndication_source": r.get("syndicationSource"),
            "review_id": r["reviewId"],
        })
    df = pd.DataFrame(rows).sort_values("date", ascending=False)
    return df.reset_index(drop=True)


# Regex themes for the dashboard's "what people are saying" panel. The signal is
# the GAP between the negative and positive rate, not the raw rate -- odor control
# scores high in both camps (polarizing), while tracking is genuinely lopsided.
THEMES = {
    "Tracking / mess":      r"track|scatter|everywhere|all over|mess|kick",
    "Packaging / bag":      r"bag|packag|torn|rip|seal|spill",
    "Odor control":         r"odor|smell|stink|stench",
    "Dust":                 r"dust",
    "Clumping":             r"clump",
    "Scooping":             r"scoop|sift",
    "Scent / fragrance":    r"scent|fragrance|perfum|pine",
    "Texture / pellet size": r"pellet|texture|sawdust|chunk|grain|too fine|too big",
    "Sticks to box":        r"stick|stuck|cement|cling|scrape",
    "Price / value":        r"price|value|cheap|expensive|worth|money|cost",
    "Cat rejected it":      r"(?:cat|cats|kitty|kitten).{0,25}(?:won'?t|refus|hate|didn'?t like|did not like|dislike|avoid)",
    "Absorbency":           r"absorb|soak",
    "Auto / self-clean box": r"auto|automatic|self.clean|litter.robot|robot",
    "Lightweight":          r"lightweight|light weight|so light",
}


def build_summary(df, reviews):
    """Aggregate into the payload the dashboard embeds."""
    df = df.copy()
    df["month"] = df["date"].dt.to_period("M").astype(str)
    txt = (df["title"].fillna("") + " " + df["text"].fillna("")).str.lower()

    def block(sub):
        if not len(sub):
            return {"n": 0, "avg": None, "pct_low": None}
        return {
            "n": int(len(sub)),
            "avg": round(float(sub["rating"].mean()), 2),
            "pct_low": round(float((sub["rating"] <= 2).mean() * 100), 1),
        }

    monthly = []
    for month, sub in df.groupby("month"):
        row = {"month": month, **block(sub)}
        for prod, psub in sub.groupby("product"):
            row[prod] = block(psub)
        monthly.append(row)

    by_variant = []
    for (prod, size), sub in df.groupby(["product", "size"], dropna=False):
        by_variant.append({"product": prod, "size": size, **block(sub)})

    neg, pos = df[df.rating <= 2], df[df.rating >= 4]
    neg_txt, pos_txt = txt[neg.index], txt[pos.index]
    themes = []
    for name, pat in THEMES.items():
        n = int(neg_txt.str.contains(pat, regex=True, na=False).sum())
        p = int(pos_txt.str.contains(pat, regex=True, na=False).sum())
        themes.append({
            "theme": name,
            "neg_n": n, "neg_pct": round(n / max(len(neg), 1) * 100, 1),
            "pos_n": p, "pos_pct": round(p / max(len(pos), 1) * 100, 1),
            "gap": round(n / max(len(neg), 1) * 100 - p / max(len(pos), 1) * 100, 1),
        })
    themes.sort(key=lambda t: -t["gap"])

    # Catalog totals include syndicated reviews; verified-only is a subset.
    catalog = {}
    for r in reviews:
        catalog.setdefault(r["_product"], {
            "ratings": r.get("_catalogTotal"),
            "avg": r.get("_catalogAvg"),
            "item_id": r.get("_itemPageId"),
        })

    # NaN is not valid JSON and JSON.parse rejects it, so scrub explicitly --
    # DataFrame.where() leaves NaN in place on object columns.
    def clean(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        return v

    rows = df.to_dict("records")
    for row in rows:
        for k in list(row):
            row[k] = clean(row[k])
        d = row.get("date")
        row["date"] = d.strftime("%Y-%m-%d") if d is not None and pd.notna(d) else None
        for k in ("helpful_yes", "helpful_no", "photos", "rating"):
            if row.get(k) is not None:
                row[k] = int(row[k])

    return {
        "as_of": datetime.now().strftime("%Y-%m-%d"),
        "n": int(len(df)),
        "avg": round(float(df["rating"].mean()), 2),
        "date_min": df["date"].min().strftime("%Y-%m-%d"),
        "date_max": df["date"].max().strftime("%Y-%m-%d"),
        "dist": {str(k): int(v) for k, v in df["rating"].value_counts().items()},
        "pct_low": round(float((df["rating"] <= 2).mean() * 100), 1),
        "catalog": catalog,
        "monthly": sorted(monthly, key=lambda r: r["month"]),
        "by_variant": sorted(by_variant, key=lambda r: (r["product"], str(r["size"]))),
        "themes": themes,
        "products": sorted(df["product"].dropna().unique().tolist()),
        "sizes": sorted(df["size"].dropna().unique().tolist()),
        "reviews": rows,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true",
                    help="refetch all pages instead of reusing cached raw JSON")
    args = ap.parse_args()

    os.makedirs(OUTDIR, exist_ok=True)
    all_reviews = []

    for item, label in PRODUCTS.items():
        raw_path = os.path.join(OUTDIR, f"reviews_raw_{item}.json")
        if not args.full and os.path.exists(raw_path):
            with open(raw_path, encoding="utf-8") as fh:
                cached = json.load(fh)
            print(f"  {label} ({item}): {len(cached)} reviews from cache "
                  f"(use --full to refetch)")
            all_reviews.extend(cached)
            continue

        reviews = collect(item, label)
        if reviews:
            with open(raw_path, "w", encoding="utf-8") as fh:
                json.dump(reviews, fh, indent=1)
            all_reviews.extend(reviews)

    if not all_reviews:
        print("[FAIL] no reviews collected")
        return 1

    df = flatten(all_reviews)
    csv_path = os.path.join(OUTDIR, "catalyst_reviews.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    monthly = (df.assign(month=df["date"].dt.to_period("M").astype(str))
                 .groupby(["month", "product"])
                 .agg(reviews=("rating", "size"), avg_rating=("rating", "mean"))
                 .round(2).reset_index())

    xlsx_path = os.path.join(OUTDIR, "catalyst_reviews.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Reviews", index=False)
        monthly.to_excel(xw, sheet_name="Monthly Trend", index=False)

    # Aggregated payload the dashboard embeds (see load_reviews in extract_data.py).
    summary = build_summary(df, all_reviews)
    summary_path = os.path.join(OUTDIR, "reviews_summary.json")
    with open(summary_path, "w", encoding="utf-8") as fh:
        # allow_nan=False so an unscrubbed NaN fails loudly here rather than
        # silently producing JSON the browser cannot parse.
        json.dump(summary, fh, separators=(",", ":"), allow_nan=False)
    print(f"[OK] dashboard payload -> {summary_path} "
          f"({os.path.getsize(summary_path) / 1024:.0f} KB)")

    print(f"\n[OK] {len(df)} verified reviews -> {csv_path}")
    print(f"     date range {df['date'].min():%Y-%m-%d} to {df['date'].max():%Y-%m-%d}")
    print(f"     overall avg rating {df['rating'].mean():.2f}")
    print(df.groupby(["product", "size"])["rating"]
            .agg(["size", "mean"]).round(2).to_string())
    return 0


if __name__ == "__main__":
    sys.exit(main())
