"""
REG-018: "Channel Pattern 2025" column on the Full Portfolio sheet.

Unlike the other update_*.py scripts, this one reads bob_salesdata_2024/
2025.xlsx (the OTHER workstream's source files, top-level data/ folder) --
because franchise-level FY24 vs FY25 Wholesale/DTC growth needs a prior
year to compare against, and the SS26 exports only cover FY25 onward.
Core-scope filtering here therefore uses is_core_customer_group() (exact
match on Customer Group), NOT is_core_market() -- see ss26_lib's
docstring / REG-008.

Sales Channel -> bucket mapping: {Wholesale, Marketplace} -> "Wholesale",
{Retail, E-com} -> "DTC".

Classification per franchise (both WH and DTC compared against
MIN_RELIABLE=50,000 SEK independently, in BOTH years):
  - both reliable in both years -> compare FY25 vs FY24 growth:
      "Substitution (DTC ↑ / Wholesale ↓)" / "Substitution (Wholesale ↑ / DTC ↓)"
      "Co-growth (both ↑)" / "Co-decline (both ↓)"
  - Wholesale reliable, DTC negligible both years -> "Wholesale-only / DTC negligible"
  - DTC reliable, Wholesale negligible both years -> "DTC-only / Wholesale negligible"
  - anything else (thin/mixed data) -> "Insufficient data"
  - franchise absent from the FY24/25 salesdata entirely -> "Not in FY24/25 dataset"

The two negligible-bucket categories were split out from a single
"Insufficient data" bucket after the user found 36% of that original
bucket actually had real, single-channel-concentrated sales -- worth
distinguishing from genuinely thin/no data.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_channel_pattern.py
"""
from __future__ import annotations

from collections import Counter, defaultdict

import ss26_lib as lib

SALESDATA_FILES = [
    (lib.REPO_ROOT / "data" / "bob_salesdata_2024.xlsx", 2024),
    (lib.REPO_ROOT / "data" / "bob_salesdata_2025.xlsx", 2025),
]


def channel_bucket(sales_channel):
    if sales_channel in ("Wholesale", "Marketplace"):
        return "Wholesale"
    if sales_channel in ("Retail", "E-com"):
        return "DTC"
    return None


def classify(d):
    wh24, wh25 = d.get("Wholesale_2024", 0.0), d.get("Wholesale_2025", 0.0)
    dtc24, dtc25 = d.get("DTC_2024", 0.0), d.get("DTC_2025", 0.0)

    wh_reliable = abs(wh24) >= lib.MIN_RELIABLE and abs(wh25) >= lib.MIN_RELIABLE
    dtc_reliable = abs(dtc24) >= lib.MIN_RELIABLE and abs(dtc25) >= lib.MIN_RELIABLE
    wh_small_both = abs(wh24) < lib.MIN_RELIABLE and abs(wh25) < lib.MIN_RELIABLE
    dtc_small_both = abs(dtc24) < lib.MIN_RELIABLE and abs(dtc25) < lib.MIN_RELIABLE

    if wh_reliable and dtc_reliable:
        wh_growth = (wh25 - wh24) / wh24 * 100
        dtc_growth = (dtc25 - dtc24) / dtc24 * 100
        if dtc_growth > 0 and wh_growth < 0:
            return "Substitution (DTC ↑ / Wholesale ↓)"
        elif wh_growth > 0 and dtc_growth < 0:
            return "Substitution (Wholesale ↑ / DTC ↓)"
        elif wh_growth > 0 and dtc_growth > 0:
            return "Co-growth (both ↑)"
        else:
            return "Co-decline (both ↓)"
    elif wh_reliable and dtc_small_both:
        return "Wholesale-only / DTC negligible"
    elif dtc_reliable and wh_small_both:
        return "DTC-only / Wholesale negligible"
    else:
        return "Insufficient data"


def main():
    franchise_year_bucket = defaultdict(lambda: defaultdict(float))
    excluded_sales = total_sales = 0.0
    for path, year in SALESDATA_FILES:
        header, rows = lib.load_raw_export(path)
        idx = {n: i for i, n in enumerate(header)}
        for row in rows:
            if row[idx["Year"]] == "Total":
                continue
            a = row[idx["Article"]]
            if not a or a == "*MISSING*":
                continue
            a = a.strip()
            sales = row[idx["Garp SEK Sales"]] or 0
            total_sales += sales
            if not lib.is_core_customer_group(row[idx["Customer Group"]]):
                excluded_sales += sales
                continue
            bucket = channel_bucket(row[idx["Sales Channel"]])
            if bucket is None:
                continue
            key = lib.franchise_key(a)
            franchise_year_bucket[key][f"{bucket}_{year}"] += sales

    pct = excluded_sales / total_sales * 100 if total_sales else 0
    print(f"total_sales(raw): {total_sales:,.0f}  excluded(special accts): {excluded_sales:,.0f} ({pct:.1f}%)")
    print("distinct franchises with any data:", len(franchise_year_bucket))

    wb, ws = lib.load_full_portfolio()
    s = lib.styles()

    ws["T4"] = "Channel Pattern 2025"
    ws["T4"].font = s["header_font"]
    ws["T4"].fill = s["header_fill"]
    ws.column_dimensions["T"].width = 30

    counts = Counter()
    for r in range(lib.FIRST_DATA_ROW, ws.max_row + 1):
        b = ws.cell(row=r, column=1).value
        if b is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        key = (b.strip(), g)
        d = franchise_year_bucket.get(key)
        label = classify(d) if d is not None else "Not in FY24/25 dataset"
        counts[label] += 1
        ws.cell(row=r, column=20, value=label).font = s["body_font"]

    print(counts)
    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
