"""
REG-022: "5. Sales by Country" sheet -- franchise x country FY25/YTD2026
sales & units, so a question like "Hero products in Japan and Germany,
excluding Scandinavian countries" can be answered directly instead of
only at the whole-portfolio grain.

SOURCE PROBLEM (REG-007, now resolved for this purpose): the raw exports
don't have a clean "country" field. data_3 (FY25 Sep-Dec) has a fairly
clean "Sales Market" field; data_1/data_2 (SS26YTD, FY25 Jan-Aug) mix
DTC store/channel identity into the SAME field ("Sales Market/Store"),
e.g. "Outlet Barkarby", "E-com Sweden", "Brand Store Sthlm" -- these are
store names and e-com-by-country labels, not country names themselves.
MARKET_TO_COUNTRY below maps every one of the 50 distinct raw values
seen across all three files to a country (or an explicit non-country
bucket), built from the raw store names' real-world locations (e.g.
"Outlet Haparanda" -> Sweden, "Brand Store Chamonix" -> France,
"Outlet Helsinki"/"Brand Store Helsinki" -> Finland).

REG-022 originally flagged "Scandinavian vs Nordic" as an open
definitional choice rather than silently picking one. RESOLVED
9-Sep-2026, business-directed: the standard exclusion group is
**Nordic (Sweden/Norway/Denmark/Finland)**, not the stricter
Scandinavian-only cut. The `Region Group (Nordic vs Non-Nordic)` column
is the one to filter/group on for normal use. The stricter
`Scandinavian (SE/NO/DK)` and the underlying `Nordic (+FI)` Yes/No
columns are kept alongside it (not removed) for anyone who specifically
wants the narrower Sweden/Norway/Denmark-only cut.

ONE THING STILL LEFT DELIBERATELY OPEN (do not silently resolve if
extending this): "Pop-Up Sales Haglöfs" (994 rows) and "Export Other"
(145 rows) carry no city/country in the label itself and are NOT
guessed at -- mapped to the explicit bucket "Unmapped / Other," not
silently assigned to Sweden or dropped. Region Group gives this bucket
its own "Unmapped" value rather than folding it into "Non-Nordic" --
lumping it in would quietly inflate a Non-Nordic total with a
geography nobody actually confirmed (501 of 13,438 franchise-country
rows, 3.7% -- a small share of rows, but real activity, not noise).
XXL (523 rows) is an account, not a geography (REG-008) -- dropped
entirely, same as everywhere else in this workbook.

Sales_2025/Units_2025 come from data_2+data_3 (FY25 halves), matching
Units_2025's own basis (REG-020). Sales_YTD2026/Units_YTD2026 come from
data_1 (SS26YTD) -- this is a raw-recomputed YTD2026 figure, NOT the
published Sheet-2 Units_YTD2026/Sales_YTD2026 column, so don't expect
these to sum exactly to that column's total for a franchise (REG-012's
ordertype-gap caveat applies here too).

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_country_breakdown.py
"""
from __future__ import annotations

from collections import defaultdict

import ss26_lib as lib

COUNTRY_SHEET = "5. Sales by Country"

FY25_EXPORTS = [
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_2_FY25_Jan_to_Aug.xlsx",
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_3_FY25_Sep_to_Dec.xlsx",
]
YTD26_EXPORTS = [
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_1_SS26YTD.xlsx",
]

UNMAPPED = "Unmapped / Other"

# Every distinct raw Sales Market/Store value seen across all three SS26
# exports (confirmed by direct inspection, Sep 2026) -> country, or an
# explicit non-country bucket. XXL is dropped separately (see main()),
# not listed here as "a country."
MARKET_TO_COUNTRY = {
    # already-clean country names (data_3, and Wholesale rows in data_1/2)
    "Austria": "Austria", "Belgium": "Belgium", "China": "China",
    "Czech Republic": "Czech Republic", "Denmark": "Denmark",
    "Finland": "Finland", "France": "France", "Germany": "Germany",
    "Hong Kong": "Hong Kong", "Italy": "Italy", "Japan": "Japan",
    "Korea": "Korea", "Netherlands": "Netherlands", "Norway": "Norway",
    "Poland": "Poland", "Spain": "Spain", "Sweden": "Sweden",
    "Switzerland": "Switzerland", "Taiwan": "Taiwan",
    "United Kingdom": "United Kingdom",
    # E-com <country> -> that country
    "E-com Austria": "Austria", "E-com Belgium": "Belgium",
    "E-com Denmark": "Denmark", "E-com Finland": "Finland",
    "E-com France": "France", "E-com Germany": "Germany",
    "E-com Italy": "Italy", "E-com Netherlands": "Netherlands",
    "E-com Norway": "Norway", "E-com Spain": "Spain",
    "E-com Sweden": "Sweden", "E-com UK": "United Kingdom",
    # DTC store names -> the country that city is actually in
    "Brand Store Chamonix": "France",       # Chamonix-Mont-Blanc, French Alps
    "Brand Store Göteborg": "Sweden",       # Gothenburg
    "Brand Store Helsinki": "Finland",
    "Brand Store Sthlm": "Sweden",          # Stockholm
    "Brand Store Åre": "Sweden",            # Åre, Jämtland
    "Outlet Avesta": "Sweden",              # Avesta, Dalarna
    "Outlet Barkarby": "Sweden",            # Järfälla, Stockholm County
    "Outlet Haparanda": "Sweden",           # Norrbotten, at the Finland border
    "Outlet Hede": "Sweden",                # Härjedalen
    "Outlet Helsinki": "Finland",
    "Outlet Norway Oslo": "Norway",
    "Pop-up Insjön": "Sweden",              # Insjön, Dalarna
    "Pop-up Båstad": "Sweden",              # Båstad, Skåne
    # regional buckets that are NOT a single country -- kept as their own
    # label rather than force-mapped to one member country
    "Benelux": "Benelux (region)",
    "North America": "North America (region)",
    # genuinely unresolvable from the label alone -- see module docstring
    "Export Other": UNMAPPED,
    "Pop-Up Sales Haglöfs": UNMAPPED,
}

SCANDINAVIAN = {"Sweden", "Norway", "Denmark"}
NORDIC = SCANDINAVIAN | {"Finland"}


def aggregate(exports):
    """(base, gender, country) -> {"sales": float, "units": float}"""
    out = defaultdict(lambda: {"sales": 0.0, "units": 0.0})
    unmapped_seen = set()
    for path in exports:
        header, rows = lib.load_raw_export(path)
        idx = {n: i for i, n in enumerate(header)}
        market_field = lib.SS26_MARKET_FIELD[path.name]
        for row in rows:
            a = row[idx["Article"]]
            if not a:
                continue
            a = a.strip()
            raw_market = row[idx[market_field]]
            if raw_market == "XXL":
                continue  # account, not a geography (REG-008) -- dropped, not a country
            country = MARKET_TO_COUNTRY.get(raw_market)
            if country is None:
                unmapped_seen.add(raw_market)
                country = UNMAPPED
            base, gender = lib.franchise_key(a)
            key = (base, gender, country)
            out[key]["sales"] += row[idx["Garp SEK Sales"]] or 0
            out[key]["units"] += row[idx["Units Sold"]] or 0
    if unmapped_seen:
        print(f"WARNING: raw market value(s) not in MARKET_TO_COUNTRY, treated as {UNMAPPED}: {sorted(unmapped_seen)}")
    return out


def main():
    fy25 = aggregate(FY25_EXPORTS)
    ytd26 = aggregate(YTD26_EXPORTS)

    all_keys = set(fy25) | set(ytd26)
    print(f"{len(all_keys)} distinct (franchise, country) combinations with any activity")

    wb, portfolio_ws = lib.load_full_portfolio()
    franchise_meta = lib.full_portfolio_lookup(portfolio_ws)  # (base, gender) -> {tier, layer, ...}
    s = lib.styles()

    if COUNTRY_SHEET in wb.sheetnames:
        del wb[COUNTRY_SHEET]
    ws = wb.create_sheet(COUNTRY_SHEET)

    headers = [
        "Base", "Gender", "Layer", "Tier", "Country/Market",
        "Region Group (Nordic vs Non-Nordic)",
        "Scandinavian (SE/NO/DK)", "Nordic (+FI)",
        "Sales_2025 (SEK)", "Units_2025",
        "Sales_YTD2026 (SEK, raw)", "Units_YTD2026 (raw)",
    ]
    header_row = lib.FIRST_DATA_ROW - 1
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = s["header_font"]
        cell.fill = s["header_fill"]
        ws.column_dimensions[cell.column_letter].width = max(14, len(h) + 2)

    n_matched_franchise = n_unmatched_franchise = 0
    r = lib.FIRST_DATA_ROW
    for base, gender, country in sorted(all_keys):
        meta = franchise_meta.get((base, gender))
        if meta is None:
            n_unmatched_franchise += 1
            continue
        n_matched_franchise += 1
        f25 = fy25.get((base, gender, country), {"sales": 0, "units": 0})
        y26 = ytd26.get((base, gender, country), {"sales": 0, "units": 0})
        if country == UNMAPPED:
            region_group = "Unmapped"
        elif country in NORDIC:
            region_group = "Nordic"
        else:
            region_group = "Non-Nordic"
        row_vals = [
            base, gender, meta["layer"], meta["tier"], country,
            region_group,
            "Yes" if country in SCANDINAVIAN else "No",
            "Yes" if country in NORDIC else "No",
            round(f25["sales"]) or None, round(f25["units"]) or None,
            round(y26["sales"]) or None, round(y26["units"]) or None,
        ]
        for c, v in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = s["body_font"]
            if c in (9, 11):
                cell.number_format = s["num_fmt"]
        r += 1

    ws.freeze_panes = ws.cell(row=lib.FIRST_DATA_ROW, column=1)
    print(f"rows written: {r - lib.FIRST_DATA_ROW}  franchise-matched: {n_matched_franchise}  "
          f"franchise-key not in Sheet 2 (skipped): {n_unmatched_franchise}")

    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
