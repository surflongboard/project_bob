"""
REG-021: "Style Code(s)" column on the Full Portfolio sheet.

Reads all three SS26 exports (SS26YTD, FY25 Jan-Aug, FY25 Sep-Dec) and
collects the distinct raw `Model` values seen for each franchise
(Base+Gender), so people tracking a franchise elsewhere (e.g. against a
PLM/ERP system) have the underlying style/model code(s) to search on --
the published workbook otherwise carries no code at all, only
franchise names.

IMPORTANT caveat, confirmed while building this: Model is NOT 1:1 with
franchise, or even with the exact Article text. The same Article name can
carry several Model codes across seasons (a re-launch keeps the product
name but gets a new internal Model number) -- 1,353 of 1,731 franchises
with any raw sales (78%) have exactly one Model code, but 378 (22%) have
2-6. This column lists every distinct code seen, comma-separated, sorted
-- not a single canonical "the" style code. Franchises with zero raw
sales rows (129 of 1,860, mostly Thin/Immaterial or Exited with no
history in these three files) get no value here, same as other
raw-export-derived columns (REG-012's caveat).

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_style_codes.py
"""
from __future__ import annotations

from collections import defaultdict

import ss26_lib as lib


def main():
    franchise_models = defaultdict(set)
    for path in lib.SS26_EXPORTS:
        header, rows = lib.load_raw_export(path)
        idx = {n: i for i, n in enumerate(header)}
        for row in rows:
            a = row[idx["Article"]]
            if not a:
                continue
            a = a.strip()
            model = row[idx["Model"]]
            if model is None:
                continue
            key = lib.franchise_key(a)
            franchise_models[key].add(str(model).strip())

    wb, ws = lib.load_full_portfolio()
    s = lib.styles()

    ws["V4"] = "Style Code(s)"
    ws["V4"].font = s["header_font"]
    ws["V4"].fill = s["header_fill"]
    ws.column_dimensions["V"].width = 22

    n_single = n_multi = n_none = 0
    for r in range(lib.FIRST_DATA_ROW, ws.max_row + 1):
        b = ws.cell(row=r, column=1).value
        if b is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        key = (b.strip(), g)
        models = franchise_models.get(key)
        if models:
            val = ", ".join(sorted(models))
            if len(models) == 1:
                n_single += 1
            else:
                n_multi += 1
        else:
            val = None
            n_none += 1
        cell = ws.cell(row=r, column=22, value=val)
        cell.font = s["body_font"]

    print(f"single code: {n_single}  multiple codes: {n_multi}  no raw data at all: {n_none}")
    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
