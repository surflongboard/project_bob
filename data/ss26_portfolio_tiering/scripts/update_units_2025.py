"""
REG-020: "Units_2025" column on the Full Portfolio sheet.

Reads the two FY25 SS26 exports (Jan-Aug, Sep-Dec) and sums the raw
"Units Sold" column per franchise (Core-scope only, via is_core_market) --
the same two files and the same franchise-key/Core-scope method already
used for Sales_2025 (predates these scripts) and for REG-017/REG-018.

Written as a raw unit count, NOT blanked below MIN_RELIABLE -- consistent
with how Sales_2025 itself (SEK) is a raw total, not gated by the 50,000
SEK/year reliability threshold. MIN_RELIABLE only gates *rate* metrics
(GM%, growth%, share%) built on a small denominator, not sums.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_units_2025.py
"""
from __future__ import annotations

from collections import defaultdict

import ss26_lib as lib

FY25_EXPORTS = [
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_2_FY25_Jan_to_Aug.xlsx",
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_3_FY25_Sep_to_Dec.xlsx",
]


def main():
    franchise_units = defaultdict(float)
    for path in FY25_EXPORTS:
        header, rows = lib.load_raw_export(path)
        idx = {n: i for i, n in enumerate(header)}
        market_field = lib.SS26_MARKET_FIELD[path.name]
        for row in rows:
            a = row[idx["Article"]]
            if not a:
                continue
            a = a.strip()
            if not lib.is_core_market(row[idx[market_field]]):
                continue
            key = lib.franchise_key(a)
            franchise_units[key] += row[idx["Units Sold"]] or 0

    wb, ws = lib.load_full_portfolio()
    s = lib.styles()

    ws["U4"] = "Units_2025"
    ws["U4"].font = s["header_font"]
    ws["U4"].fill = s["header_fill"]
    ws.column_dimensions["U"].width = 14

    n_matched = n_no_data = 0
    for r in range(lib.FIRST_DATA_ROW, ws.max_row + 1):
        b = ws.cell(row=r, column=1).value
        if b is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        key = (b.strip(), g)
        units = franchise_units.get(key)
        if units:
            n_matched += 1
            val = round(units)
        else:
            n_no_data += 1
            val = None
        cell = ws.cell(row=r, column=21, value=val)
        cell.font = s["body_font"]
        if val is not None:
            cell.number_format = s["num_fmt"]

    print(f"matched (units > 0): {n_matched}  no raw units at all: {n_no_data}")
    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
