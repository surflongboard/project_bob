"""
REG-023 (rewritten 9-Sep-2026): Region-relative tiering using the REAL,
now-confirmed base rule (REG-024) -- "6. Full Portfolio (Nordic)" and
"7. Full Portfolio (Non-Nordic)" sheets.

HISTORY: the first version of this script (see git history) invented a
brand-new percentile-based rule after an extensive attempt to
reverse-engineer Sheet 2's actual tier-cut logic failed (best composite
score only explained 41% of the real Hero+Near-Hero set). The business
then supplied the real rule directly -- see REG-024 in the register and
`verify_base_tiering.py`, which validates it against Sheet 2 at 99.8%
accuracy. This version reapplies THAT REAL rule (via ss26_lib.classify_tier)
with region-scoped inputs, not an invented one.

REGION-SCALED THRESHOLDS (business-directed 9-Sep-2026): the confirmed
rule's SEK floors (2M Hero, 1M Near-Hero, 100K/50K cross-channel) are
GLOBAL, absolute numbers. Applying them unscaled to each region separately
would just be "same absolute thresholds as Global" -- the option
explicitly NOT chosen (see chat history) -- since Non-Nordic's smaller
regional pool would rarely clear a bar calibrated to global-scale revenue.
Instead, each region's SEK floors are scaled by that region's own share of
total FY25 Core sales (Nordic ~68.1%, Non-Nordic ~31.9%, computed fresh
from bob_salesdata_2025.xlsx each run -- not hardcoded). GM%/growth cuts
are NOT scaled -- they're already relative (%), not absolute SEK, so
scaling them would double-count the adjustment. The 100-unit "real
trading" threshold is also NOT scaled -- it's a physical/operational
reality check (did this genuinely sell), not a revenue-scale artifact.

DATA SOURCE: bob_salesdata_2024.xlsx / bob_salesdata_2025.xlsx (REG-018's
precedent file family -- clean native "Sales Market" field, Core-scope via
is_core_customer_group()). Carries REG-009's known ~2.3% cross-workstream
gap vs. the SS26 exports. A literal "Total" row and a single zero-value
"Restored" row are excluded, not real market data.

Country breakdown alongside the regional tier reuses Sheet 5's
already-validated per-country FY25/YTD2026 rows (REG-022), filtered to the
region -- avoids a second, differently-sourced country split in the same
sheet. Franchises with a computed regional tier but NO country-level rows
in that region (no sales there at all) still get one summary row, with
the country column reading "— (no {region} sales)" -- the first version of
this script silently omitted them (733 of 1,860 were missing from the
Non-Nordic sheet), which is now fixed.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_regional_tiering.py
"""
from __future__ import annotations

from collections import Counter, defaultdict

import ss26_lib as lib
import update_country_breakdown as cb  # reuse NORDIC set + Region Group logic

# Global (unscaled) floors from the confirmed rule -- see REG-024.
HERO_SALES_FLOOR_GLOBAL = 2_000_000
NEAR_HERO_SALES_FLOOR_GLOBAL = 1_000_000
WHOLESALE_FLOOR_GLOBAL = 100_000
DTC_FLOOR_GLOBAL = 50_000

BOB_SALES_FILES = {
    2024: lib.REPO_ROOT / "data" / "bob_salesdata_2024.xlsx",
    2025: lib.REPO_ROOT / "data" / "bob_salesdata_2025.xlsx",
}
NON_MARKET_VALUES = {"Total", "Restored"}
NO_PRESENCE = "— (no regional sales)"


def load_franchise_region_totals(year):
    """(base, gender, region) -> {sales, units, margin, wholesale, dtc}, Core-scope only.
    region is 'Nordic', 'Non-Nordic', or 'Unmapped' (REG-022's country->region map)."""
    import openpyxl

    path = BOB_SALES_FILES[year]
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {n: i for i, n in enumerate(header)}
    out = defaultdict(lambda: {"sales": 0.0, "units": 0.0, "margin": 0.0, "wholesale": 0.0, "dtc": 0.0})
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
        key = (base, gender, region)
        sales = row[idx["Garp SEK Sales"]] or 0
        out[key]["sales"] += sales
        out[key]["units"] += row[idx["Units Sold"]] or 0
        out[key]["margin"] += row[idx["Garp SEK Margin"]] or 0
        if row[idx["Sales Channel"]] in ("Wholesale", "Marketplace"):  # REG-018 precedent
            out[key]["wholesale"] += sales
        else:
            out[key]["dtc"] += sales
    return out


def region_revenue_share(fy25, region_name):
    region_total = sum(v["sales"] for (b, g, r), v in fy25.items() if r == region_name)
    grand_total = sum(v["sales"] for (b, g, r), v in fy25.items() if r in ("Nordic", "Non-Nordic"))
    return region_total / grand_total if grand_total else 0.0


CONSOLIDATED = {
    "Hero": "Hero + Near-Hero", "Near-Hero (Rising Star)": "Hero + Near-Hero",
    "Workhorse": "Workhorse + Harvest", "Harvest (Cash Cow)": "Workhorse + Harvest",
    "Problem Child": "Problem Child", "New / Test": "New / Test",
    "Thin / Immaterial": "Thin / Immaterial", "Exited": "Exited",
}


def build_region_data(fy24, fy25, franchise_universe, region_name, revenue_share):
    """Returns {(base,gender): {...}} with the region-scoped inputs AND the
    resulting tier (REG-013 consolidated naming) for every franchise."""
    hero_floor = HERO_SALES_FLOOR_GLOBAL * revenue_share
    near_hero_floor = NEAR_HERO_SALES_FLOOR_GLOBAL * revenue_share
    wholesale_floor = WHOLESALE_FLOOR_GLOBAL * revenue_share
    dtc_floor = DTC_FLOOR_GLOBAL * revenue_share

    out = {}
    for base, gender, clearance_flag in franchise_universe:
        key = (base, gender)
        a24 = fy24.get((base, gender, region_name), {"sales": 0, "units": 0, "margin": 0, "wholesale": 0, "dtc": 0})
        a25 = fy25.get((base, gender, region_name), {"sales": 0, "units": 0, "margin": 0, "wholesale": 0, "dtc": 0})
        sales25 = a25["sales"]
        gm25 = round(a25["margin"] / sales25 * 100, 1) if sales25 else None
        growth = round((sales25 - a24["sales"]) / a24["sales"] * 100, 1) if a24["sales"] else None

        if clearance_flag == "Fully":
            sub_tier = "Clearance — Ex China/Zalando"
        else:
            sub_tier = lib.classify_tier(
                sales=sales25, units_prior=a24["units"], units_current=a25["units"],
                growth=growth, gm_pct=gm25 or 0.0, wholesale_sek=a25["wholesale"], dtc_sek=a25["dtc"],
                sales_floor_hero=hero_floor, sales_floor_near_hero=near_hero_floor,
                wholesale_floor=wholesale_floor, dtc_floor=dtc_floor,
            )
        tier = CONSOLIDATED.get(sub_tier, sub_tier)
        out[key] = {
            "tier": tier, "subTier": sub_tier,
            "sales24": round(a24["sales"]) or None, "sales25": round(sales25) or None,
            "units25": round(a25["units"]) or None, "gm25": gm25, "growth": growth,
            "wholesale25": round(a25["wholesale"]) or None, "dtc25": round(a25["dtc"]) or None,
        }
    return out


def write_region_sheet(wb, sheet_name, region_name, region_data, country_rows_by_key, franchise_meta, s):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    headers = [
        "Base", "Gender", "Layer", f"Regional Tier ({region_name})",
        f"Sales_2024 ({region_name}, bob_salesdata basis)",
        f"Sales_2025 ({region_name}, bob_salesdata basis)",
        f"Units_2025 ({region_name}, bob_salesdata basis)",
        f"GM%_2025 ({region_name})", f"Growth% ({region_name}, FY24→FY25)",
        f"Wholesale_2025 ({region_name}, SEK)", f"DTC_2025 ({region_name}, SEK)",
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
    n_with_country = n_no_presence = 0
    for key, rd in region_data.items():
        base, gender = key
        meta = franchise_meta.get(key, {})
        franchise_cols = [
            base, gender, meta.get("layer"), rd["tier"],
            rd["sales24"], rd["sales25"], rd["units25"], rd["gm25"], rd["growth"],
            rd["wholesale25"], rd["dtc25"],
        ]
        crows = country_rows_by_key.get(key)
        if crows:
            n_with_country += 1
            for crow in crows:
                row_vals = franchise_cols + [
                    crow["country"], crow["sales2025"], crow["units2025"], crow["salesYtd26"], crow["unitsYtd26"],
                ]
                for c, v in enumerate(row_vals, start=1):
                    ws.cell(row=r, column=c, value=v).font = s["body_font"]
                r += 1
        else:
            n_no_presence += 1
            row_vals = franchise_cols + [NO_PRESENCE, None, None, None, None]
            for c, v in enumerate(row_vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = s["gray_font"] if c == 12 else s["body_font"]
            r += 1
    ws.freeze_panes = ws.cell(row=lib.FIRST_DATA_ROW, column=1)
    print(f"{sheet_name}: {n_with_country} franchises with country rows, "
          f"{n_no_presence} with a tier but no regional sales at all, {r - lib.FIRST_DATA_ROW} total rows")


def main():
    fy24 = load_franchise_region_totals(2024)
    fy25 = load_franchise_region_totals(2025)

    nordic_share = region_revenue_share(fy25, "Nordic")
    non_nordic_share = region_revenue_share(fy25, "Non-Nordic")
    print(f"FY25 Core revenue share -- Nordic: {nordic_share*100:.1f}%  Non-Nordic: {non_nordic_share*100:.1f}%")
    print(f"Scaled Hero floor -- Nordic: {HERO_SALES_FLOOR_GLOBAL*nordic_share:,.0f} SEK  "
          f"Non-Nordic: {HERO_SALES_FLOOR_GLOBAL*non_nordic_share:,.0f} SEK")

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

    nordic_data = build_region_data(fy24, fy25, franchise_universe, "Nordic", nordic_share)
    non_nordic_data = build_region_data(fy24, fy25, franchise_universe, "Non-Nordic", non_nordic_share)

    for label, data in [("Nordic", nordic_data), ("Non-Nordic", non_nordic_data)]:
        print(f"{label} tier distribution: {dict(Counter(v['tier'] for v in data.values()))}")

    if "5. Sales by Country" not in wb.sheetnames:
        raise SystemExit("Sheet 5 (REG-022) not found -- run update_country_breakdown.py first")
    country_ws = wb["5. Sales by Country"]
    nordic_country_rows = defaultdict(list)
    non_nordic_country_rows = defaultdict(list)
    for r in range(lib.FIRST_DATA_ROW, country_ws.max_row + 1):
        b = country_ws.cell(row=r, column=1).value
        if b is None:
            continue
        key = (b, country_ws.cell(row=r, column=2).value)
        region_group = country_ws.cell(row=r, column=6).value
        crow = {
            "country": country_ws.cell(row=r, column=5).value,
            "sales2025": country_ws.cell(row=r, column=9).value,
            "units2025": country_ws.cell(row=r, column=10).value,
            "salesYtd26": country_ws.cell(row=r, column=11).value,
            "unitsYtd26": country_ws.cell(row=r, column=12).value,
        }
        if region_group == "Nordic":
            nordic_country_rows[key].append(crow)
        elif region_group == "Non-Nordic":
            non_nordic_country_rows[key].append(crow)

    write_region_sheet(wb, "6. Full Portfolio (Nordic)", "Nordic", nordic_data, nordic_country_rows, franchise_meta, s)
    write_region_sheet(wb, "7. Full Portfolio (Non-Nordic)", "Non-Nordic", non_nordic_data, non_nordic_country_rows, franchise_meta, s)

    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
