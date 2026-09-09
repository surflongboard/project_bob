"""
Read-only "Browse" sheet: the same 1,860 rows as Sheet 2, reordered for
human scanning (Sales/Units/GM% grouped per year) via live formulas
pointing back at Sheet 2 -- so this view can never go stale, and Sheet
2's own column order/positions (which every update_*.py script hardcodes)
never has to change.

Do NOT hand-edit this sheet -- it's formulas, not data. To change the
column grouping, edit COLUMN_ORDER below and re-run.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/build_browse_view.py
"""
from __future__ import annotations

import ss26_lib as lib

BROWSE_SHEET = "2b. Full Portfolio (Browse)"

# (source column letter on Sheet 2, header text to show). Same 22 columns,
# just grouped: identity -> FY25 (Sales/Units/GM%) -> YTD2026 (Sales/Units/
# GM%) -> Pace/Growth -> Style Code -> everything else in its existing order.
COLUMN_ORDER = [
    ("A", "Base"),
    ("B", "Gender"),
    ("C", "Layer"),
    ("D", "Tier"),
    ("E", "Bucket"),
    ("F", "Sales_2025 (SEK)"),
    ("U", "Units_2025"),
    ("G", "GM%_2025"),
    ("H", "Sales_YTD2026 (SEK)"),
    ("I", "Units_YTD2026"),
    ("J", "GM%_YTD2026"),
    ("K", "Pace %"),
    ("L", "Growth% (FY25 vs FY24)"),
    ("V", "Style Code(s)"),
    ("M", "SpecialAcctExposure%"),
    ("N", "Clearance Flag"),
    ("O", "Clearance Detail"),
    ("P", "Original Tier (pre-consolidation)"),
    ("Q", "Core Assortment FW27"),
    ("R", "FW27 Collection"),
    ("S", "Wholesale Share % (FY25, raw)"),
    ("T", "Channel Pattern 2025"),
]


def main():
    wb, _ = lib.load_full_portfolio()
    src = lib.FULL_PORTFOLIO_SHEET
    src_ws = wb[src]
    s = lib.styles()

    if BROWSE_SHEET in wb.sheetnames:
        del wb[BROWSE_SHEET]
    ws = wb.create_sheet(BROWSE_SHEET, index=wb.sheetnames.index(src) + 1)

    ws["A1"] = (
        "Read-only view of '" + src + "' -- same rows, columns grouped for "
        "browsing (Sales next to Units next to GM% per year). Formulas, not "
        "data -- edit the source sheet, not this one. Rebuild with "
        "scripts/build_browse_view.py."
    )
    ws["A1"].font = s["gray_font"]

    header_row = lib.FIRST_DATA_ROW - 1  # row 4, matching every other sheet
    for col_idx, (src_letter, header) in enumerate(COLUMN_ORDER, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = s["header_font"]
        cell.fill = s["header_fill"]
        ws.column_dimensions[cell.column_letter].width = max(14, len(header) + 2)

    for r in range(lib.FIRST_DATA_ROW, src_ws.max_row + 1):
        base = src_ws.cell(row=r, column=1).value
        if base is None:
            continue
        for col_idx, (src_letter, _) in enumerate(COLUMN_ORDER, start=1):
            ws.cell(row=r, column=col_idx, value=f"='{src}'!{src_letter}{r}")

    ws.freeze_panes = ws.cell(row=lib.FIRST_DATA_ROW, column=1)

    print(f"built '{BROWSE_SHEET}': {src_ws.max_row - header_row} data rows, {len(COLUMN_ORDER)} columns")
    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
