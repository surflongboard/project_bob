"""
Rebuilds qa_data/country_breakdown.json (Tier Bench's `queryByCountry`
tool) by reading the master workbook's "5. Sales by Country" sheet
directly, row for row -- NOT by re-deriving the franchise x country
matching from the three raw SS26 exports (that logic lives in
scripts/update_country_breakdown.py, REG-022, and stays the single
source of truth for Sheet 5 itself). This script's job is only the last
step: turn the published sheet into Tier Bench's JSON shape, so re-run
scripts/update_country_breakdown.py FIRST if Sheet 5 itself needs
refreshing (a new data drop, a matching-logic fix), then this script.

Usage (from this directory):
    python3 extract_country_breakdown.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import ss26_lib as lib

COUNTRY_SHEET = "5. Sales by Country"


def main():
    wb = openpyxl.load_workbook(lib.WORKBOOK_PATH, read_only=True, data_only=True)
    ws = wb[COUNTRY_SHEET]

    out_rows = []
    for row in ws.iter_rows(min_row=lib.FIRST_DATA_ROW, values_only=True):
        base = row[0]
        if base is None:
            continue
        sales25 = row[8]
        units25 = row[9]
        sales_ytd26 = row[10]
        units_ytd26 = row[11]
        out_rows.append({
            "base": base,
            "gender": row[1],
            "layer": row[2],
            "tier": row[3],
            "country": row[4],
            "regionGroup": row[5],
            "scandinavian": row[6] == "Yes",
            "nordic": row[7] == "Yes",
            "sales25": round(sales25) if sales25 is not None else None,
            "units25": round(units25) if units25 is not None else None,
            "salesYtd26": round(sales_ytd26) if sales_ytd26 is not None else None,
            "unitsYtd26": round(units_ytd26) if units_ytd26 is not None else None,
        })
    wb.close()
    print("country_breakdown rows:", len(out_rows))

    out_dir = Path(__file__).resolve().parent / "qa_data"
    out_dir.mkdir(exist_ok=True)
    with open(out_dir / "country_breakdown.json", "w") as f:
        json.dump({"rows": out_rows}, f)
    print("written", out_dir / "country_breakdown.json", len(json.dumps({"rows": out_rows})), "bytes")


if __name__ == "__main__":
    main()
