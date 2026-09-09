"""
REG-023: Region-relative tiering -- "6. Full Portfolio (Nordic)" and
"7. Full Portfolio (Non-Nordic)" sheets.

CONTEXT: attempted to reverse-engineer the Global tiering's exact
Hero/Near-Hero/Workhorse/Harvest/Problem Child cut rule from Sales_2025,
GM%_2025, Growth%, and Pace% (the only fields available for it) so the
same rule could be reapplied to Nordic-only and Non-Nordic-only
populations. It does NOT reduce to any simple formula on these fields --
tested absolute thresholds, top-N by sales, top-N per Layer, and several
composite scores (sales x growth, sales x margin, etc.); the best of
these only explained ~41% of the actual Hero+Near-Hero set, no better
than noise. Likely a business/merchandising judgment call made with more
context than these columns carry. Business-directed (9-Sep-2026): build
a NEW, transparent, explicitly-documented rule instead of guessing at
the old one, using the SAME 7 tier names for continuity, with
REGION-RELATIVE thresholds (each region ranked against its own
population, not a shared global SEK bar) -- see RULE below for the
exact, adjustable cut points.

DATA SOURCE for growth: Sheet 5 (REG-022) only has FY25 + YTD2026 by
country -- no FY24, so no growth-by-region is possible from it alone.
Pulled FY24 and FY25 by country fresh from bob_salesdata_2024.xlsx /
bob_salesdata_2025.xlsx instead (REG-018's precedent: same two files,
clean native "Sales Market" field, no store-name parsing needed, unlike
the SS26 exports). Core-scope filtered via is_core_customer_group() --
the correct filter for THIS file family (REG-008), not is_core_market()
(that's for the SS26 exports). Carries REG-009's known ~2.3% cross-
workstream total-sales gap vs the SS26 exports -- read growth as
directional, not audited. A literal "Total" row (877.15M SEK, matching
REG-009's cited total exactly) and a single zero-value "Restored" row
are excluded, not real market data.

RULE (v1, adjustable -- these are the parameters to revisit if the
business wants different cut points):
  Within EACH region (Nordic, Non-Nordic) independently:
    1. Exited: region FY25 Sales < MIN_RELIABLE, region FY24 Sales >= MIN_RELIABLE
       (had real regional sales before, negligible now).
    2. New/Test: region FY24 Sales < MIN_RELIABLE, region FY25 Sales >= MIN_RELIABLE
       (negligible/no regional history before, real sales now).
    3. Thin/Immaterial: both FY24 and FY25 region Sales < MIN_RELIABLE.
    4. Clearance: Sheet 2's existing Clearance Flag == "Fully" carried over
       as-is -- REG-010's China JV/Zalando SKU lists aren't region-specific,
       so this tier isn't region-relative like the rest.
    5. Everything else ("Continuing" pool, region-local) ranked by region
       FY25 Sales descending, split into the SAME proportions Global uses
       within its own Continuing pool (34/615 Hero+Near-Hero, 431/615
       Workhorse+Harvest, 150/615 Problem Child -- chosen only to keep the
       tier-size shape recognizable, not because it's "correct"):
         - Top slice (Hero+Near-Hero share) AND region Growth% >= HERO_GROWTH_FLOOR
           (-5%, i.e. roughly flat-or-better) -- else falls through to the
           Workhorse+Harvest slice instead. Split in half by region sales:
           top half = Hero, bottom half = Near-Hero.
         - Next slice (Workhorse+Harvest share): Workhorse if region Growth%
           >= WORKHORSE_GROWTH_FLOOR (-5%), else Harvest.
         - Remaining (Problem Child share): Problem Child regardless of
           growth sign -- Global's own Problem Child spans a huge growth
           range too, so "smallest of the continuing performers" is this
           rule's read of what that tier actually captures, not "declining."

Country breakdown shown alongside: reuses Sheet 5's already-validated
per-country FY25/YTD2026 rows (REG-022), filtered to the region, rather
than re-deriving a second country split from bob_salesdata -- avoids two
different "Sales by country" numbers with different provenance sitting
in the same sheet.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_regional_tiering.py
"""
from __future__ import annotations

from collections import defaultdict

import ss26_lib as lib
import update_country_breakdown as cb  # reuse NORDIC set + Region Group logic

MIN_RELIABLE = lib.MIN_RELIABLE
HERO_GROWTH_FLOOR = -5.0
WORKHORSE_GROWTH_FLOOR = -5.0

BOB_SALES_FILES = {
    2024: lib.REPO_ROOT / "data" / "bob_salesdata_2024.xlsx",
    2025: lib.REPO_ROOT / "data" / "bob_salesdata_2025.xlsx",
}
NON_MARKET_VALUES = {"Total", "Restored"}  # grand-total / adjustment rows, not real markets


def load_bob_sales_by_country(year):
    """(base, gender, country) -> Garp SEK Sales, Core-scope only, for one FY."""
    import openpyxl

    path = BOB_SALES_FILES[year]
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {n: i for i, n in enumerate(header)}
    out = defaultdict(float)
    for row in rows:
        market = row[idx["Sales Market"]]
        if market in NON_MARKET_VALUES or market is None:
            continue
        if market == "XXL":
            continue  # account, not a geography (REG-008) -- dropped, not "a country"
        if not lib.is_core_customer_group(row[idx["Customer Group"]]):
            continue
        a = row[idx["Article"]]
        if not a:
            continue
        base, gender = lib.franchise_key(a.strip())
        country = market if market != "Export Other" else cb.UNMAPPED
        region = "Nordic" if country in cb.NORDIC else ("Unmapped" if country == cb.UNMAPPED else "Non-Nordic")
        out[(base, gender, region)] += row[idx["Garp SEK Sales"]] or 0
    return out


def classify_region(pool):
    """pool: list of dicts with base/gender/fy24/fy25/growth for ONE region's
    Continuing-pool franchises (already excluded: Exited/New-Test/Thin/
    Clearance). Returns {(base,gender): sub-tier} using this region's own
    sales ranking -- every franchise in `pool` ends up in exactly one of
    the three buckets below, no fallthrough/edge cases."""
    n = len(pool)
    n_hero_nh = round(n * 34 / 615)
    n_problem_child = round(n * 150 / 615)
    ranked = sorted(pool, key=lambda p: -p["fy25"])  # sales-ranked, region-local

    # Take the top-by-sales candidates that ALSO clear the growth floor as
    # Hero/Near-Hero; anyone in that top slice who doesn't clear it falls
    # through to the Workhorse/Harvest pool instead (still sales-ranked).
    hero_candidates, remaining = [], []
    for p in ranked:
        if len(hero_candidates) < n_hero_nh and p["growth"] is not None and p["growth"] >= HERO_GROWTH_FLOOR:
            hero_candidates.append(p)
        else:
            remaining.append(p)

    result = {}
    half = len(hero_candidates) / 2
    for i, p in enumerate(hero_candidates):
        result[(p["base"], p["gender"])] = "Hero" if i < half else "Near-Hero"

    n_workhorse_harvest = max(len(remaining) - n_problem_child, 0)
    for i, p in enumerate(remaining):
        key = (p["base"], p["gender"])
        if i < n_workhorse_harvest:
            result[key] = "Workhorse" if (p["growth"] is not None and p["growth"] >= WORKHORSE_GROWTH_FLOOR) else "Harvest"
        else:
            result[key] = "Problem Child"
    return result


CONSOLIDATED = {
    "Hero": "Hero + Near-Hero", "Near-Hero": "Hero + Near-Hero",
    "Workhorse": "Workhorse + Harvest", "Harvest": "Workhorse + Harvest",
    "Problem Child": "Problem Child",
}


def build_region_data(fy24, fy25, franchise_universe, region_name):
    """Returns {(base,gender): {"tier": consolidated tier, "origTier": raw sub-tier,
    "sales24": .., "sales25": .., "growth": ..}} for every franchise in the universe."""
    out = {}
    continuing_pool = []
    for base, gender, clearance_flag in franchise_universe:
        key = (base, gender)
        s24 = fy24.get((base, gender, region_name), 0.0)
        s25 = fy25.get((base, gender, region_name), 0.0)
        growth = round((s25 - s24) / s24 * 100, 1) if s24 >= MIN_RELIABLE else None
        if clearance_flag == "Fully":
            tier, orig = "Clearance — Ex China/Zalando", "Clearance — Ex China/Zalando"
        elif s25 < MIN_RELIABLE and s24 >= MIN_RELIABLE:
            tier, orig = "Exited", "Exited"
        elif s24 < MIN_RELIABLE and s25 >= MIN_RELIABLE:
            tier, orig = "New / Test", "New / Test"
        elif s24 < MIN_RELIABLE and s25 < MIN_RELIABLE:
            tier, orig = "Thin / Immaterial", "Thin / Immaterial"
        else:
            tier, orig = None, None  # resolved after ranking the Continuing pool
            continuing_pool.append({"base": base, "gender": gender, "fy24": s24, "fy25": s25, "growth": growth})
        out[key] = {"sales24": s24, "sales25": s25, "growth": growth, "tier": tier, "origTier": orig}

    sub_tiers = classify_region(continuing_pool)
    for key, sub in sub_tiers.items():
        out[key]["origTier"] = sub
        out[key]["tier"] = CONSOLIDATED[sub]
    return out


def write_region_sheet(wb, sheet_name, region_group_value, region_data, country_rows, franchise_meta, s):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    headers = [
        "Base", "Gender", "Layer", f"Regional Tier ({region_group_value}-relative)",
        f"Sales_2024 ({region_group_value}, bob_salesdata basis)",
        f"Sales_2025 ({region_group_value}, bob_salesdata basis)",
        f"Growth% ({region_group_value}, FY24→FY25)",
        "Country/Market", "Sales_2025 (SEK, country, SS26 basis)", "Units_2025 (country, SS26 basis)",
        "Sales_YTD2026 (SEK, country, SS26 basis, raw)", "Units_YTD2026 (country, SS26 basis, raw)",
    ]
    header_row = lib.FIRST_DATA_ROW - 1
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = s["header_font"]
        cell.fill = s["header_fill"]
        ws.column_dimensions[cell.column_letter].width = max(14, len(h) + 2)

    r = lib.FIRST_DATA_ROW
    n_written = 0
    for crow in country_rows:
        key = (crow["base"], crow["gender"])
        rd = region_data.get(key)
        if rd is None:
            continue
        meta = franchise_meta.get(key, {})
        row_vals = [
            crow["base"], crow["gender"], meta.get("layer"), rd["tier"],
            round(rd["sales24"]) or None, round(rd["sales25"]) or None, rd["growth"],
            crow["country"], crow["sales2025"], crow["units2025"], crow["salesYtd26"], crow["unitsYtd26"],
        ]
        for c, v in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = s["body_font"]
        r += 1
        n_written += 1
    ws.freeze_panes = ws.cell(row=lib.FIRST_DATA_ROW, column=1)
    print(f"{sheet_name}: {n_written} country rows written")


def main():
    fy24 = load_bob_sales_by_country(2024)
    fy25 = load_bob_sales_by_country(2025)

    wb, portfolio_ws = lib.load_full_portfolio()
    s = lib.styles()
    franchise_meta = lib.full_portfolio_lookup(portfolio_ws)

    franchise_universe = []
    for r in range(lib.FIRST_DATA_ROW, portfolio_ws.max_row + 1):
        b = portfolio_ws.cell(row=r, column=1).value
        if b is None:
            continue
        g = (portfolio_ws.cell(row=r, column=2).value or "").strip()
        clearance_flag = portfolio_ws.cell(row=r, column=14).value
        franchise_universe.append((b.strip(), g, clearance_flag))

    nordic_data = build_region_data(fy24, fy25, franchise_universe, "Nordic")
    non_nordic_data = build_region_data(fy24, fy25, franchise_universe, "Non-Nordic")

    for label, data in [("Nordic", nordic_data), ("Non-Nordic", non_nordic_data)]:
        from collections import Counter
        print(f"{label} tier distribution: {Counter(v['tier'] for v in data.values())}")

    if "5. Sales by Country" not in wb.sheetnames:
        raise SystemExit("Sheet 5 (REG-022) not found -- run update_country_breakdown.py first")
    country_ws = wb["5. Sales by Country"]
    all_country_rows = []
    for r in range(lib.FIRST_DATA_ROW, country_ws.max_row + 1):
        b = country_ws.cell(row=r, column=1).value
        if b is None:
            continue
        all_country_rows.append({
            "base": b, "gender": country_ws.cell(row=r, column=2).value,
            "region_group": country_ws.cell(row=r, column=6).value,
            "country": country_ws.cell(row=r, column=5).value,
            "sales2025": country_ws.cell(row=r, column=9).value,
            "units2025": country_ws.cell(row=r, column=10).value,
            "salesYtd26": country_ws.cell(row=r, column=11).value,
            "unitsYtd26": country_ws.cell(row=r, column=12).value,
        })

    nordic_country_rows = [r for r in all_country_rows if r["region_group"] == "Nordic"]
    non_nordic_country_rows = [r for r in all_country_rows if r["region_group"] == "Non-Nordic"]

    write_region_sheet(wb, "6. Full Portfolio (Nordic)", "Nordic", nordic_data, nordic_country_rows, franchise_meta, s)
    write_region_sheet(wb, "7. Full Portfolio (Non-Nordic)", "Non-Nordic", non_nordic_data, non_nordic_country_rows, franchise_meta, s)

    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
