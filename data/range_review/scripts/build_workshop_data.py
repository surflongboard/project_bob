"""
Merges the franchise review list, MOQ exceptions, and inventory review
into one JSON dataset for the Nov 12 layer-by-layer workshop page
(the web-based review format, chosen over Excel because embedded images
didn't work well there -- RR-015).

Join key is the BASE name (version-stripped, `_base` on each franchise
row), not the displayed "Franchise" column (version-kept, RR-003) --
build_moq_exceptions.py and build_inventory_review.py both key their own
"Franchise" field on the base name already, so joining on the displayed
name would silently miss every franchise with a version token (e.g.
"Alert GTX Jacket" vs "Alert GTX II Jacket" -- the bug this script
exists to not repeat).

Carries both the FW27+SS28 combined forecast (forecastQty/forecastGP, as
shown on the franchise list) and per-season splits (forecastQtyFW27/
forecastGPFW27, forecastQtySS28/forecastGPSS28) computed in
build_franchise_list.py -- the workshop page's FW27/SS28/All toggle shows
whichever one matches the selected season rather than always showing the
combined total, which it did at first and was misleading once the toggle
existed.

Inventory review moved to its own "Inventory Planning" workstream on
30-Oct-2026 (data/inventory_planning/, see its ASSUMPTIONS_REGISTER.md
and RR-018/019/020/024's pointer entries here) -- imported cross-
workstream below, the same way `ss26_lib` is imported from the SS26
Portfolio Tiering workstream. Re-run
`data/inventory_planning/scripts/build_inventory_review.py` first if
its output is stale.

The old "Generation timeline" card (IP-005's stock-snapshot-proxy
dates, via `build_generation_timeline.load_generation_groups`/
`aggregate_by_generation`) was REMOVED from this script 3-Nov-2026
(RR-032) -- IP-011's `build_generation_release_timeline.build_generation_table()`
(the Product master's real per-code release seasons) superseded it,
and keeping both was duplicate information on the Board. IP-005's own
script/output still exist in inventory_planning for other uses; only
the cross-workstream wiring into this file was removed.

Run: python3 build_workshop_data.py
Output: ../outputs/workshop_data.json
"""
from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

RR_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RR_DIR / "scripts"))
from build_franchise_list import build_rows  # noqa: E402
from build_forecast_discrepancies import load_ss28_qty_by_code  # noqa: E402
from build_historical_sales import (  # noqa: E402
    build_by_base as build_historical_sales_by_base,
    build_by_base_and_generation as build_historical_sales_by_base_and_generation,
)

REPO_ROOT = RR_DIR.parents[1]
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402

IP_DIR = REPO_ROOT / "data" / "inventory_planning"
sys.path.insert(0, str(IP_DIR / "scripts"))
from build_generation_release_timeline import build_generation_table, season_key, seasons_between  # noqa: E402

GENERATION_AGE_FLAG_SEASONS = 4  # RR-029/IP-011/RR-030 -- no confirmed renewal cadence exists; 4
                                  # seasons (2 years, user-set 3-Nov-2026, was 6) is a judgment
                                  # call for both the card alert and the "Generation age" toggle,
                                  # not a business rule -- adjust freely.

OUTPUTS_DIR = RR_DIR / "outputs"
MOQ_PATH = OUTPUTS_DIR / "SS28_MOQ_Exceptions.xlsx"
INVENTORY_PATH = IP_DIR / "outputs" / "SS27_Inventory_Review.xlsx"


def load_moq_by_base() -> dict[str, list[dict]]:
    """Each MOQ-exception row also carries the SS28 forecast CSV's own Qty
    for the same style Code, and whether it disagrees with the MOQ email's
    FC1 demand (RR-023) -- surfaced as its own "SS28 forecast discrepancy"
    flag/toggle on the board, separate from the MOQ shortfall itself."""
    ss28_csv_qty_by_code = load_ss28_qty_by_code()
    wb = openpyxl.load_workbook(MOQ_PATH, data_only=True)
    ws = wb["SS28 MOQ Exceptions"]
    by_base: dict[str, list[dict]] = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        base = r[2]
        code, demand_fc1 = r[0], r[11]
        csv_qty = ss28_csv_qty_by_code.get(code)
        has_discrepancy = csv_qty is None or round(csv_qty) != demand_fc1
        by_base[base].append({
            "code": code, "style": r[1], "colorway": r[6],
            "moq_style": r[9], "demand_fc1": demand_fc1, "shortfall": r[12],
            "buying_comment": r[13], "merch_comment": r[14],
            "ss28CsvQty": round(csv_qty) if csv_qty is not None else None,
            "hasDiscrepancy": has_discrepancy,
        })
    return dict(by_base)


def load_inventory_by_base() -> dict[str, dict]:
    wb = openpyxl.load_workbook(INVENTORY_PATH, data_only=True)
    ws = wb["Franchise Summary"]
    by_base = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        base = r[0]
        by_base[base] = {
            "available_active": r[3],
            "available_not_active": r[4],
            "skus_total": r[5],
            "skus_not_in_range": r[6],
        }
    return by_base


def load_tail_styles_by_base() -> dict[str, list[dict]]:
    """Per franchise, which NOT-in-active-range style code(s) make up its
    inventory tail -- e.g. "Mimic Alert 2.0 Hood" (current) carries leftover
    stock attributed to it because the pre-"2.0" generation ("Mimic Alert
    Hood", codes 607444/607447) shares the same version-stripped base. The
    generic "N units in codes no longer in the active range" alert gave no
    way to see WHICH old style that was -- this breaks it out by style so
    the card can name it, same as the MOQ alert names each flagged style."""
    wb = openpyxl.load_workbook(INVENTORY_PATH, data_only=True)
    ws = wb["SKU Detail"]
    totals: dict[tuple[str, str, str], int] = defaultdict(int)
    for r in ws.iter_rows(min_row=2, values_only=True):
        in_active_range, franchise, style_code, article, available_stock = r[6], r[7], r[1], r[2], r[9]
        if in_active_range == "Y" or not franchise:
            continue
        key = (franchise, style_code, (article or "").strip())
        totals[key] += available_stock or 0

    by_base: dict[str, list[dict]] = defaultdict(list)
    for (base, style_code, article), stock in totals.items():
        if stock <= 0:
            continue
        by_base[base].append({"code": style_code, "style": article, "availableStock": stock})
    for base in by_base:
        by_base[base].sort(key=lambda s: -s["availableStock"])
    return dict(by_base)


_GEN_RANK = {"": 0, "Generation 1": 0, "2.0": 1, "II": 1, "III": 2, "IV": 3}


def load_generation_age_by_base(historical_sales_by_base_and_gen: dict[str, dict]) -> dict[str, dict]:
    """Per active-range franchise: how old its CURRENT generation is (catalog label, for the
    "New generation" chip) AND how old whatever is ACTUALLY SELLING is (for the "Generation age"
    alert) -- RR-041, these can diverge badly. IP-011's Release Season is when a generation was
    added to the Product master/catalog, not when it actually launched in the market: "Rosson
    Mid II Jacket/Hood" is catalogued since SS27 (2 seasons old) but has sold a grand total of 5
    sample units (all in an excluded market) through Oct 2026 -- every real unit sold in
    2024-2026 is still under the Generation-1 codes, which released FW24 (7 seasons old). Found
    from a user question on the Board, 3-Nov-2026 ("seems they are around +3 years given the
    historical data... why aren't they flagged +4?") -- checked and confirmed 22 active-range
    franchises have this same pattern (catalog-current generation with ZERO real sales anywhere).
    User direction: age whichever generation is ACTUALLY selling, not the catalog's latest.
    Walks from the newest generation down (by `_GEN_RANK`) until it finds one with real recorded
    sales (`historical_sales_by_base_and_gen`, RR-037/RR-039's per-generation sales, matched by
    the same sales-side version-token bucket) -- falls back to the catalog-current generation
    only if NONE of a franchise's generations have any recorded sales at all."""
    gen_table = build_generation_table()
    by_franchise: dict[str, list[dict]] = defaultdict(list)
    for r in gen_table:
        by_franchise[r["Franchise"]].append(r)

    by_base = {}
    for base, rows in by_franchise.items():
        rows_sorted = sorted(rows, key=lambda r: _GEN_RANK.get(r["Generation"], 99))
        catalog_cur = next((r for r in rows_sorted if r["Is Current Generation"] == "Y"), rows_sorted[-1])
        sales_for_base = historical_sales_by_base_and_gen.get(base, {})

        def has_real_sales(row):
            sales_key = "" if row["Generation"] in ("", "Generation 1") else row["Generation"]
            years = sales_for_base.get(sales_key, {})
            return sum(y.get("units", 0) for y in years.values()) > 0

        effective = next((r for r in reversed(rows_sorted) if has_real_sales(r)), None)
        if effective is None:
            effective = catalog_cur

        by_base[base] = {
            "label": catalog_cur["Generation"],
            "releaseSeason": catalog_cur["Release Season"],
            "isNew": catalog_cur["Release Season"] in ("FW27", "SS28", "FW28"),  # RR-042 -- FW27 added, the earliest season this active range covers
            "ageLabel": effective["Generation"],
            "ageReleaseSeason": effective["Release Season"],
            "ageSeasons": seasons_between(season_key(effective["Release Season"]), season_key("SS28"))
                          if effective["Release Season"] else None,
            "agedByPreviousGeneration": effective is not catalog_cur,
        }
    return by_base


def load_historical_sales_by_generation_by_base(by_base_gen: dict[str, dict],
                                                 multi_gen_bases: set[str]) -> dict[str, list[dict]]:
    """RR-039 -- per franchise, a list of {generation, years} entries, oldest first, so a
    multi-generation franchise's card can show its PREVIOUS generation's historical sales
    alongside its current one ("as one franchise, what's the historic sales" -- the blended
    `historicalSales` field alone hides exactly this, since a new "II"/"III" generation barely
    has any sales yet and the blended number is almost entirely the old generation's).

    Gated on `multi_gen_bases` (2+ generations in the Product master's own catalog, IP-011) --
    NOT on whether 2+ of them have their own sales bucket. RR-043 fix: the original version
    gated on sales-bucket count, so a franchise whose "II" has literally never sold (RR-041's
    22 cases, e.g. "Alert GTX II Jacket") fell through to the single blended row with a generic
    "Historical sales" label -- which happens to equal Generation 1's full history (II
    contributes nothing) but didn't SAY so, reading as if the predecessor's sales were missing
    rather than already being the only number shown. Now any base with 2+ catalog generations
    gets labeled rows for whichever of its generations actually have sales data (often just
    one) -- a single-generation franchise still relies on the plain blended `historicalSales`
    field, no label needed since there's nothing to distinguish it from."""
    out: dict[str, list[dict]] = {}
    for base, gens in by_base_gen.items():
        if base not in multi_gen_bases:
            continue
        rows = []
        for gen in sorted(gens, key=lambda g: _GEN_RANK.get(g, 99)):
            rows.append({"generation": gen or "Generation 1", "years": gens[gen]})
        out[base] = rows
    return out


def build() -> list[dict]:
    franchise_rows = build_rows()
    moq_by_base = load_moq_by_base()
    inventory_by_base = load_inventory_by_base()
    tail_styles_by_base = load_tail_styles_by_base()
    historical_sales_by_base = build_historical_sales_by_base()
    historical_sales_by_base_and_gen = build_historical_sales_by_base_and_generation()
    generation_age_by_base = load_generation_age_by_base(historical_sales_by_base_and_gen)
    gen_counts: dict[str, int] = defaultdict(int)
    for r in build_generation_table():
        gen_counts[r["Franchise"]] += 1
    multi_gen_bases = {base for base, n in gen_counts.items() if n > 1}
    historical_sales_by_gen_by_base = load_historical_sales_by_generation_by_base(
        historical_sales_by_base_and_gen, multi_gen_bases)

    out = []
    for r in franchise_rows:
        base = r["_base"]
        inv = inventory_by_base.get(base, {})
        moq = moq_by_base.get(base, [])
        franchise_version = lib.version_token(r["Franchise"])
        tail_styles = []
        for s in tail_styles_by_base.get(base, []):
            tail_styles.append({**s, "sameGeneration": lib.version_token(s["style"]) == franchise_version
                                 and franchise_version is not None})

        age = generation_age_by_base.get(base)
        active_fw27, active_ss28, active_fw28 = r["Active in FW27"] == "Y", r["Active in SS28"] == "Y", r["Active in FW28"] == "Y"
        age_seasons = age["ageSeasons"] if age else None
        # "New in season" -- either the Product master's own Release Season says so (age["isNew"],
        # RR-042: FW27/SS28/FW28, i.e. anywhere in this active range's own coverage), or the
        # franchise isn't in the Product master at all (brand-new FW28 launch, RR-029's "No
        # Product master match" case) and FW28 is its only active season.
        if age is not None:
            is_new = age["isNew"]
            new_season_label = age["releaseSeason"] if is_new else None
        elif active_fw28 and not active_fw27 and not active_ss28:
            is_new, new_season_label = True, "FW28"
        else:
            is_new, new_season_label = False, None
        # Distinguish a genuinely NEW FRANCHISE (first-ever appearance) from a NEW GENERATION of
        # a franchise that already existed -- e.g. "Alert GTX II Jacket" is active FW27+SS28+FW28
        # (continuous), but its "II" generation's own Release Season is SS28. The first pass
        # (3-Nov-2026) checked "was the franchise active in a season before this generation's own
        # release" -- but that misses a predecessor that already phased out before FW27 even
        # started, e.g. "L.I.M Airak 2,5L Proof II Jacket"/"L.I.M GTX III Pant": their own Gen-1
        # row either predates the Product master's SS20 coverage or just isn't active any more,
        # so active_fw27/active_ss28 were both False and they were wrongly called "franchise"
        # (caught by the user 3-Nov-2026). The reliable signal doesn't need Gen 1 to be visible at
        # all: a version token (II/2.0/III/IV) IN THE CURRENT GENERATION'S OWN NAME is definitional
        # proof something came before it, regardless of whether that predecessor still shows up
        # anywhere in available data. `age["label"]` already IS that token (blank only for a true
        # first generation) -- no separate lookup needed.
        new_kind = None
        if is_new:
            new_kind = "generation" if (age is not None and age["label"]) else "franchise"

        out.append({
            "base": base,
            "franchise": r["Franchise"],
            "layer": r["Layer"],
            "businessArea": r["Business Area"],
            "gender": r["Gender"],
            "genderNote": r["Gender note"],
            "activeFW27": r["Active in FW27"] == "Y",
            "activeSS28": r["Active in SS28"] == "Y",
            "activeFW28": r["Active in FW28"] == "Y",  # RR-026 -- tag only, no FW28 forecast/MOQ data to show yet
            "targetRRP": r["Target RRP"],
            "rrpVaries": r["RRP varies by style"] == "Y",
            "consumerActivity": r["Consumer Activity"],
            "marginTarget": r["Target GM0%"] or None,
            "specialEditionNote": r["Special edition / fit note"],
            "numStyles": r["# styles rolled up"],
            "numColors": r["# active colors"],
            "coreAssortmentFW27": r["Core Assortment FW27"] == "Yes",
            "coreAssortmentSS28": r["Core Assortment SS28"],
            "fc1TopdownDemandSS28": r["SS28 FC1 Demand (Top-Down Comparison)"] or None,  # RR-028 -- kept for audit; no longer its own card stat (RR-038, folded into forecastQtySS28 below)
            "abs": r["ABS"] == "Yes",
            "asiaFit": r["Asia Fit"] == "Yes",
            "statusFW27": r["Status FW27"],
            "statusSS28": r["Status SS28"],
            "forecastQty": r["Forecast Qty (total)"] or None,
            "forecastGP": r["Forecast GP% (blended)"] or None,
            "forecastQtyDTC": r["Forecast Qty (DTC)"] or None,
            "forecastGPDTC": r["Forecast GP% (DTC)"] or None,
            "forecastQtyWS": r["Forecast Qty (Wholesale)"] or None,
            "forecastGPWS": r["Forecast GP% (Wholesale)"] or None,
            "forecastQtyFW27": r["_forecast_fw27_qty"],  # FW27 FC3, unchanged (RR-012)
            "forecastGPFW27": r["_forecast_fw27_gp"],
            # RR-038 -- SS28's card figure now comes from the FC1 Top-Down vs Bottom-Up file
            # (RR-028's fc1TopdownDemandSS28), per user direction 3-Nov-2026, NOT the original
            # bottom-up forecast CSV (RR-011/RR-023) any more -- that figure is kept under
            # forecastQtySS28BottomUp for audit, since RR-023 found it doesn't reconcile with
            # the MOQ email's FC1 figure either and this swap doesn't resolve that, just picks
            # which number the card shows.
            "forecastQtySS28": r["SS28 FC1 Demand (Top-Down Comparison)"] or None,
            "forecastGPSS28": r["_forecast_ss28_gp"],
            "forecastQtySS28BottomUp": r["_forecast_ss28_qty"],
            "forecastQtyFW28": None,  # RR-026/RR-034 -- no FW28 forecast file exists yet
            "historicalSales": historical_sales_by_base.get(base, {}),  # RR-037 -- {"2024": {...}, "2025": {...}, "2026": {...}}
            "historicalSalesByGeneration": historical_sales_by_gen_by_base.get(base, []),  # RR-039 -- only set for 2+ generations
            "imageFile": r["_image_file"],
            "moq": moq,
            "hasForecastDiscrepancy": any(m["hasDiscrepancy"] for m in moq),
            "availableStockActive": inv.get("available_active", 0),
            "availableStockTail": inv.get("available_not_active", 0),
            "skusNotInRange": inv.get("skus_not_in_range", 0),
            "tailStyles": tail_styles,
            "generationLabel": age["label"] if age else "",  # IP-011 -- CATALOG-current generation's own label, e.g. "III" (chip/hint only)
            "generationReleaseSeason": age["releaseSeason"] if age else "",
            # RR-041 -- the generation actually used for the age alert/flag may NOT be the
            # catalog-current one above, if that generation has never actually sold (see
            # load_generation_age_by_base's docstring). ageGenerationLabel/ageGenerationReleaseSeason
            # are what the age alert displays; agedByPreviousGeneration flags when a fallback happened.
            "ageGenerationLabel": age["ageLabel"] if age else "",
            "ageGenerationReleaseSeason": age["ageReleaseSeason"] if age else "",
            "agedByPreviousGeneration": age["agedByPreviousGeneration"] if age else False,
            "generationAgeSeasons": age_seasons,
            "generationAgeFlag": age_seasons is not None and age_seasons >= GENERATION_AGE_FLAG_SEASONS,
            "isNewInSeason": is_new,  # RR-029 -- first season for this generation (or brand-new FW28 launch)
            "newKind": new_kind,  # "generation" (franchise continues, only this generation is new) or "franchise" (first-ever appearance)
            "newSeasonLabel": new_season_label,
        })
    return out


if __name__ == "__main__":
    data = build()
    out_path = OUTPUTS_DIR / "workshop_data.json"
    with open(out_path, "w") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    n_moq = sum(1 for d in data if d["moq"])
    n_discrepancy = sum(1 for d in data if d["hasForecastDiscrepancy"])
    n_tail = sum(1 for d in data if d["availableStockTail"])
    n_images = sum(1 for d in data if d["imageFile"])
    n_same_gen = sum(1 for d in data for s in d["tailStyles"] if s["sameGeneration"])
    print(f"{len(data)} franchises written to {out_path}")
    print(f"{n_moq} with MOQ flags ({n_discrepancy} with a forecast discrepancy), "
          f"{n_tail} with inventory tail, {n_images} with images")
    print(f"{n_same_gen} tail-style lines are a retired code of the SAME generation, not an older one")

    n_age_known = sum(1 for d in data if d["generationAgeSeasons"] is not None)
    n_age_flagged = sum(1 for d in data if d["generationAgeFlag"])
    n_new = sum(1 for d in data if d["isNewInSeason"])
    print(f"{n_age_known} franchises have a known current-generation age (IP-011); "
          f"{n_age_flagged} flagged at {GENERATION_AGE_FLAG_SEASONS}+ seasons old")
    print(f"{n_new} franchises are new in FW27/SS28/FW28 (first season for their current generation)")
