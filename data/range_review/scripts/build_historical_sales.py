"""
Per-franchise historical sales (2024 / 2025 / 2026 YTD) for the Range
Review Board, per the user's request to show "2024-2026 units and GP%
per year" on each card.

**2024 uses `data/bob_salesdata_2024.xlsx` (the existing root-workstream
file), NOT the new Sales_Data_2024_H1/H2.csv export** -- cross-checking
the new export's 2024 figures against this file found its Wholesale
channel badly undercounted for 2024 specifically (82.7M SEK vs this
file's 589.3M SEK, a ~7x gap; row density is ~5-6x lower per month than
2025/2026 in the same new export). E-com/Retail matched closely, and
2025 reconciled almost exactly between old and new sources, so the gap
is specific to the new export's 2024 Wholesale pull, not a systemic
issue -- confirmed 3-Nov-2026 by inspecting the user's original,
unsplit master file (`Sales_2024_YTD_oct_2026.xlsx`), which reproduces
the identical low 2024 Wholesale figures row-for-row, so it's not a
splitting artifact either. User direction, 3-Nov-2026: "use the old
file for 2024, as 2025 new and old data seems to match anyway."

**2025/2026 use the new `inputs/sales/Sales_Data_<year>_H1/H2.csv`**
(received 3-Nov-2026) -- monthly, style-color grain (Model/Color ID/SCO
codes, Article text, Sales Channel, Sales Market/Store). 2026 is
YTD-through-October only (the user's own files stop there) -- labelled
as such, not presented as a full year.

Matched to franchise BASE (not per-style) via the same text-based
`strip_special_edition` + `ss26_lib.base_name()` used everywhere else
in this workstream, applied to the OLD file's "Article" column and the
NEW files' "Article" column alike -- same methodology for both sources,
deliberately, rather than code-matching one and text-matching the
other. This means a franchise's 2024/2025 historical sales include its
PREDECESSOR generation's sales too (e.g. "Alert GTX II Jacket"'s 2024
figure is really "Alert GTX Jacket" Gen 1's 2024 sales, since "II"
didn't exist yet) -- the intended read is "this product line's sales
history," not "this exact SKU's sales history."

**Core-account exclusion** (CLAUDE.md's "Core account" rule --
XXL/China/Zalando/Zalando Marketplace/Stadium Outlet) uses the two
field-specific functions already built for this in `ss26_lib.py` rather
than reimplementing: `is_core_customer_group()` for the OLD file's
"Customer Group" field (catches all 5 accounts exactly), and
`is_core_market()` for the NEW files' "Sales Market/Store" field
(catches only XXL/China by substring -- Zalando/Zalando Marketplace/
Stadium Outlet are NOT distinguishable in the new export's fields at
all, confirmed 3-Nov-2026; their sales are invisibly folded into their
home country's total there). This is a real, asymmetric limitation
across years, not hidden: 2024 excludes all 5 named accounts, 2025/2026
only exclude 2 of them.

GP% is blanked (not zero) below `ss26_lib.MIN_RELIABLE` (50,000 SEK of
that franchise's sales in that year) -- the shared small-denominator
convention (CLAUDE.md), applied the same way as everywhere else in
these workstreams. Units Sold is never blanked by this threshold.

Run: python3 build_historical_sales.py
Output: ../outputs/Historical_Sales_By_Franchise.xlsx
"""
from __future__ import annotations

import csv
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RR_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RR_DIR / "scripts"))
from build_franchise_list import strip_special_edition  # noqa: E402

REPO_ROOT = RR_DIR.parents[1]
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402

OLD_2024_PATH = REPO_ROOT / "data" / "bob_salesdata_2024.xlsx"
SALES_DIR = RR_DIR / "inputs" / "sales"
OUTPUTS_DIR = RR_DIR / "outputs"

NEW_CSV_YEARS = {
    2025: ["Sales_Data_2025_H1.csv", "Sales_Data_2025_H2.csv"],
    2026: ["Sales_Data_2026_H1.csv", "Sales_Data_2026_H2.csv"],
}


def base_of(article: str) -> str:
    stripped, _ = strip_special_edition(article)
    return lib.base_name(stripped)


def gen_of(article: str) -> str:
    """Raw version token ('', 'II', 'III', 'IV'...) -- the generation bucket a sale belongs to,
    not yet relabeled to 'Generation 1' (that's only meaningful once we know the franchise has
    2+ generations, done downstream in build_workshop_data.py, same convention as IP-011)."""
    stripped, _ = strip_special_edition(article)
    return lib.version_token(stripped) or ""


def load_2024_from_old_file() -> dict[tuple[str, str], dict]:
    wb = openpyxl.load_workbook(OLD_2024_PATH, read_only=True, data_only=True)
    ws = wb["2024"]
    rows = ws.iter_rows(values_only=True)
    next(rows)  # header
    agg: dict[tuple[str, str], list[float]] = defaultdict(lambda: [0.0, 0.0, 0.0])  # sek, units, margin
    for r in rows:
        article, customer_group = r[5], r[3]
        if not article or article == "Total":
            continue
        if not lib.is_core_customer_group(customer_group):
            continue
        key = (base_of(article), gen_of(article))
        a = agg[key]
        a[0] += r[7] or 0   # Garp SEK Sales
        a[1] += r[9] or 0   # Units Sold
        a[2] += r[15] or 0  # Garp SEK Margin
    return {k: {"sek": v[0], "units": v[1], "margin": v[2]} for k, v in agg.items()}


def load_year_from_new_csvs(filenames: list[str]) -> dict[tuple[str, str], dict]:
    agg: dict[tuple[str, str], list[float]] = defaultdict(lambda: [0.0, 0.0, 0.0])
    for fn in filenames:
        path = SALES_DIR / fn
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            next(reader)  # header
            for row in reader:
                article, market = row[2], row[10]
                if not article:
                    continue
                if not lib.is_core_market(market):
                    continue
                key = (base_of(article), gen_of(article))
                a = agg[key]
                try:
                    a[0] += float(row[11]) if row[11] else 0.0  # Garp SEK Sales
                except ValueError:
                    pass
                try:
                    a[1] += float(row[12]) if row[12] else 0.0  # Units Sold
                except ValueError:
                    pass
                try:
                    a[2] += float(row[13]) if row[13] else 0.0  # Garp SEK Margin
                except ValueError:
                    pass
    return {k: {"sek": v[0], "units": v[1], "margin": v[2]} for k, v in agg.items()}


def _load_all_years_by_base_and_gen() -> dict[tuple[str, str], dict]:
    """(base, generation) -> {2024: {...}, 2025: {...}, 2026: {...}}, each with units/gpPct/sek."""
    by_year = {
        2024: load_2024_from_old_file(),
        2025: load_year_from_new_csvs(NEW_CSV_YEARS[2025]),
        2026: load_year_from_new_csvs(NEW_CSV_YEARS[2026]),
    }
    out: dict[tuple[str, str], dict] = defaultdict(dict)
    for year, data in by_year.items():
        for key, v in data.items():
            gp_pct = (v["margin"] / v["sek"] * 100) if v["sek"] else None
            if gp_pct is not None and v["sek"] < lib.MIN_RELIABLE:
                gp_pct = None  # small-denominator convention -- blank, not a misleading number
            out[key][year] = {
                "units": round(v["units"]) if v["units"] else 0,
                "gpPct": round(gp_pct, 1) if gp_pct is not None else None,
                "sek": round(v["sek"]),
            }
    return dict(out)


def build_by_base() -> dict[str, dict]:
    """base -> {2024: {...}, 2025: {...}, 2026: {...}}, BLENDED across every generation sold
    under that base -- the franchise's whole historical total, same as before 3-Nov-2026's
    per-generation split was added (RR-039)."""
    by_base_gen = _load_all_years_by_base_and_gen()
    out: dict[str, dict] = defaultdict(lambda: defaultdict(lambda: {"units": 0, "sek": 0, "_margin": 0.0}))
    for (base, _gen), years in by_base_gen.items():
        for year, v in years.items():
            o = out[base][year]
            o["units"] += v["units"]
            o["sek"] += v["sek"]
            o["_margin"] += (v["sek"] * (v["gpPct"] / 100)) if v["gpPct"] is not None else 0.0
    result: dict[str, dict] = {}
    for base, years in out.items():
        result[base] = {}
        for year, v in years.items():
            gp_pct = (v["_margin"] / v["sek"] * 100) if v["sek"] else None
            if gp_pct is not None and v["sek"] < lib.MIN_RELIABLE:
                gp_pct = None
            result[base][year] = {"units": v["units"], "sek": v["sek"],
                                   "gpPct": round(gp_pct, 1) if gp_pct is not None else None}
    return result


def build_by_base_and_generation() -> dict[str, dict]:
    """base -> {generation_token ('' for no-token, 'II', 'III', ...): {2024: {...}, ...}} --
    the SAME source data as build_by_base(), not blended, so a franchise's sales can be shown
    broken out by generation (RR-039 -- "show the previous generation's sales too, so you know
    as one franchise what's the historic sales"). Generation-token relabeling to "Generation 1"
    happens downstream (build_workshop_data.py), same convention as IP-011/RR-025."""
    by_base_gen = _load_all_years_by_base_and_gen()
    out: dict[str, dict] = defaultdict(dict)
    for (base, gen), years in by_base_gen.items():
        out[base][gen] = years
    return dict(out)


def write_workbook(by_base: dict[str, dict], by_base_gen: dict[str, dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def write_sheet(title: str, headers: list[str], data_rows: list[list]) -> None:
        ws = wb.create_sheet(title)
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
        for row in data_rows:
            ws.append(row)
        for i, h in enumerate(headers, start=1):
            ws.column_dimensions[get_column_letter(i)].width = max(len(h) + 2, 14)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    headers = ["Franchise (base)", "2024 Units", "2024 GP%", "2024 SEK",
               "2025 Units", "2025 GP%", "2025 SEK",
               "2026 YTD Units", "2026 YTD GP%", "2026 YTD SEK"]
    rows = []
    for base in sorted(by_base):
        row = [base]
        for year in (2024, 2025, 2026):
            v = by_base[base].get(year, {})
            row += [v.get("units", ""), v.get("gpPct", ""), v.get("sek", "")]
        rows.append(row)
    write_sheet("Historical Sales By Franchise", headers, rows)

    # RR-039 -- same data, NOT blended across generations, so a multi-generation franchise's
    # history can be read generation by generation (e.g. Gen 1's declining sales alongside a
    # new "II"/"III" generation just starting to ramp up), not just as one combined number.
    gen_headers = ["Franchise (base)", "Generation", "2024 Units", "2024 GP%",
                   "2025 Units", "2025 GP%", "2026 YTD Units", "2026 YTD GP%"]
    gen_rows = []
    rank = {"": 0, "2.0": 1, "II": 1, "III": 2, "IV": 3}
    for base in sorted(by_base_gen):
        gens = by_base_gen[base]
        multi_gen = len(gens) > 1
        for gen in sorted(gens, key=lambda g: rank.get(g, 99)):
            label = (gen or "Generation 1") if multi_gen else gen
            row = [base, label]
            for year in (2024, 2025, 2026):
                v = gens[gen].get(year, {})
                row += [v.get("units", ""), v.get("gpPct", "")]
            gen_rows.append(row)
    write_sheet("Historical Sales By Generation", gen_headers, gen_rows)

    notes = wb.create_sheet("Notes")
    notes.append(["What this is", "Per-franchise (base-name, version-stripped) historical "
                                    "Units Sold / GP% / SEK Sales for 2024, 2025, and 2026 "
                                    "YTD-through-October. Feeds the Range Review Board's card "
                                    "stats, not a standalone deliverable."])
    notes.append(["2024 source", "data/bob_salesdata_2024.xlsx (existing root-workstream file), "
                                   "NOT the new Sales_Data_2024_H1/H2.csv export -- that export's "
                                   "2024 Wholesale channel is badly undercounted (82.7M vs this "
                                   "file's 589.3M SEK), confirmed against the user's own unsplit "
                                   "master file too, so it's a source-export issue, not a "
                                   "splitting artifact. User direction 3-Nov-2026."])
    notes.append(["2025/2026 source", "inputs/sales/Sales_Data_2025_H1/H2.csv and "
                                        "Sales_Data_2026_H1/H2.csv (received 3-Nov-2026). 2025 "
                                        "reconciles almost exactly against the old bob_salesdata_"
                                        "2025.xlsx. 2026 is YTD through October only."])
    notes.append(["Matching", "Text-based: strip_special_edition() + ss26_lib.base_name() on "
                                "each row's Article field, same methodology for every source -- "
                                "NOT a style-code match. A franchise's historical figure includes "
                                "its predecessor generation's sales (e.g. a franchise whose "
                                "current 'II' generation launched in SS28 shows real 2024/2025 "
                                "sales under its Gen-1 name) -- this is the product LINE's sales "
                                "history, not this exact SKU's."])
    notes.append(["Core-account exclusion -- asymmetric by source", "2024 (old file's "
                                                                       "'Customer Group' field) "
                                                                       "excludes all 5 named "
                                                                       "accounts (XXL SE/NO, "
                                                                       "ZALANDO, ZALANDO "
                                                                       "MARKETPLACE, China, "
                                                                       "STADIUM Outlet) exactly, "
                                                                       "via ss26_lib."
                                                                       "is_core_customer_group(). "
                                                                       "2025/2026 (new files' "
                                                                       "'Sales Market/Store' "
                                                                       "field) can only exclude "
                                                                       "XXL and China -- Zalando/"
                                                                       "Zalando Marketplace/"
                                                                       "Stadium Outlet are not "
                                                                       "distinguishable in that "
                                                                       "field at all and are "
                                                                       "folded into their home "
                                                                       "country's total, via "
                                                                       "ss26_lib.is_core_market(). "
                                                                       "This is a real asymmetry, "
                                                                       "not an oversight."])
    notes.append(["GP% blanking", f"Blank (not 0%) when that franchise-year's SEK sales are "
                                    f"under {lib.MIN_RELIABLE:,} SEK -- the shared small-"
                                    f"denominator convention. Units Sold is never blanked."])
    notes.append(["By Generation sheet (RR-039)", "Same underlying data as the main sheet, "
                                                     "NOT blended across generations -- one row "
                                                     "per (franchise, generation) so a franchise "
                                                     "currently on its 'II'/'III' generation shows "
                                                     "its earlier generation's sales as its own "
                                                     "row too. 'Generation 1' only labels the "
                                                     "no-token bucket when the franchise has 2+ "
                                                     "generations in this data; a single-generation "
                                                     "franchise's one row is left unlabeled, same "
                                                     "convention as IP-011/RR-025."])
    notes.append(["Status", "Provisional -- methodology is consistent and documented, but the "
                             "2025/2026 account-exclusion gap and the predecessor-generation "
                             "rollup are both judgment calls worth flagging if used for anything "
                             "high-stakes."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 26
    notes.column_dimensions["B"].width = 100

    wb.save(path)


if __name__ == "__main__":
    by_base = build_by_base()
    by_base_gen = build_by_base_and_generation()
    out_path = OUTPUTS_DIR / "Historical_Sales_By_Franchise.xlsx"
    write_workbook(by_base, by_base_gen, out_path)
    n_2024 = sum(1 for b in by_base if 2024 in by_base[b] and by_base[b][2024]["units"])
    n_2025 = sum(1 for b in by_base if 2025 in by_base[b] and by_base[b][2025]["units"])
    n_2026 = sum(1 for b in by_base if 2026 in by_base[b] and by_base[b][2026]["units"])
    n_multi_gen = sum(1 for b in by_base_gen if len(by_base_gen[b]) > 1)
    print(f"{len(by_base)} franchise bases written to {out_path}")
    print(f"{n_2024} with 2024 units, {n_2025} with 2025 units, {n_2026} with 2026 YTD units")
    print(f"{n_multi_gen} franchise bases have 2+ generations with their own sales history")
