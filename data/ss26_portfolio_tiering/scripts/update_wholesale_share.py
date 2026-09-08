"""
REG-017: "Wholesale Share % (FY25, raw)" column on the Full Portfolio sheet.

Reads the two FY25 SS26 exports (Jan-Aug, Sep-Dec), sums Garp SEK Sales by
Sales Channel per franchise (Core-scope only, via is_core_market), and
computes Wholesale / (Wholesale + DTC channels). Blanked below
MIN_RELIABLE (50,000 SEK) total, per REG-004's precedent -- a
small-denominator % here is misleading, not just imprecise.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_wholesale_share.py
"""
from __future__ import annotations

from collections import defaultdict

import ss26_lib as lib

FY25_EXPORTS = [
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_2_FY25_Jan_to_Aug.xlsx",
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_3_FY25_Sep_to_Dec.xlsx",
]


def main():
    franchise_channel = defaultdict(lambda: defaultdict(float))
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
            franchise_channel[key][row[idx["Sales Channel"]]] += row[idx["Garp SEK Sales"]] or 0

    wb, ws = lib.load_full_portfolio()
    s = lib.styles()

    ws["S4"] = "Wholesale Share % (FY25, raw)"
    ws["S4"].font = s["header_font"]
    ws["S4"].fill = s["header_fill"]
    ws.column_dimensions["S"].width = 22

    n_valid = n_blank_small = n_no_data = 0
    for r in range(lib.FIRST_DATA_ROW, ws.max_row + 1):
        b = ws.cell(row=r, column=1).value
        if b is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        key = (b.strip(), g)
        chd = franchise_channel.get(key)
        val = None
        if chd:
            total = sum(chd.values())
            if abs(total) >= lib.MIN_RELIABLE:
                val = round(chd.get("Wholesale", 0) / total * 100, 1)
                n_valid += 1
            else:
                n_blank_small += 1
        else:
            n_no_data += 1
        cell = ws.cell(row=r, column=19, value=val)
        cell.font = s["body_font"]
        if val is not None:
            cell.number_format = s["pct_fmt"]

    print(f"valid: {n_valid}  blanked (below 50k): {n_blank_small}  no raw data at all: {n_no_data}")
    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
