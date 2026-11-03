"""
Cross-checks the "cleaned up" FW28 tiering draft
(inputs/F28_Acces_Tiering_1st_Draft_20261006.xlsx, received 3-Nov-2026)
against the FW28 file already integrated into the franchise review list
(inputs/FW28_2_Nov_2026.xlsx, RR-026) -- every discrepancy found, for the
team to confirm before either file replaces or merges into the Board.

NOT integrated anywhere yet -- this is a comparison report only. See
RR-027 for the full write-up; don't resolve any of these by picking an
assumption in code.

Discrepancy types:
  - "Internal contradiction (new file)": the SAME code appears twice in
    the new file with conflicting Active/Status values.
  - "Active status conflict": same code, Active differs between files.
  - "Missing from new file": a code Active in the old file doesn't
    appear anywhere in the new file (by code OR by matching style name).
  - "Possibly renumbered": a code missing from the new file, but its
    exact style name appears there under a different code.
  - "Layer reclassified" / "Collection Type reclassified": same code,
    same style, a categorization field changed between the two files.

Also lists, as a separate category, the codes that exist ONLY in the
new file -- not really a "discrepancy" (the old file was Apparel-only;
the new one spans Apparel/Hardware/Footwear) but worth the team seeing
the scale of it before signing off.

Run: python3 build_fw28_discrepancies.py
Output: ../outputs/FW28_Discrepancies.xlsx
"""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RR_DIR = Path(__file__).resolve().parent.parent
INPUTS_DIR = RR_DIR / "inputs"
OUTPUTS_DIR = RR_DIR / "outputs"
OLD_PATH = INPUTS_DIR / "FW28_2_Nov_2026.xlsx"
NEW_PATH = INPUTS_DIR / "F28_Acces_Tiering_1st_Draft_20261006.xlsx"


def load_old() -> dict[str, dict]:
    """FW28_2_Nov_2026.xlsx -- flat one-row-per-colorway, Apparel-only (RR-026)."""
    wb = openpyxl.load_workbook(OLD_PATH, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    by_code: dict[str, dict] = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        code = r[2]
        by_code.setdefault(code, {
            "style": r[3], "active": r[1], "layer": r[14], "gender": r[15],
            "business_area": r[16], "collection_type": r[23],
        })
    return by_code


def load_new() -> tuple[dict[str, dict], dict[str, list[dict]], list[dict]]:
    """F28_Acces_Tiering_1st_Draft_20261006.xlsx -- style grain already (no
    colorway sub-rows), but 4 codes appear twice with contradictory data
    (RR-027) -- those go in `duplicates`, keyed by code, not collapsed.
    Rows with no code assigned yet ("xxx", proposed new styles) go in
    `uncoded`, kept separate since they can't be compared by code at all."""
    wb = openpyxl.load_workbook(NEW_PATH, read_only=True, data_only=True)
    ws = wb["FW28"]
    grouped: dict[str, list[dict]] = defaultdict(list)
    uncoded: list[dict] = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        code = r[2]
        entry = {
            "style": r[3], "active": r[1], "business_area": r[4], "status": r[7],
            "layer": r[12], "collection_type": r[13], "gender": r[14],
        }
        if code is None or code == "xxx":
            uncoded.append(entry)
            continue
        grouped[code].append(entry)

    by_code: dict[str, dict] = {}
    duplicates: dict[str, list[dict]] = {}
    for code, entries in grouped.items():
        if len(entries) > 1:
            duplicates[code] = entries
            active_true = [e for e in entries if e["active"]]
            by_code[code] = active_true[0] if active_true else entries[0]
        else:
            by_code[code] = entries[0]
    return by_code, duplicates, uncoded


def build_rows() -> tuple[list[dict], list[dict]]:
    old = load_old()
    new, duplicates, uncoded = load_new()

    discrepancies: list[dict] = []

    for code, entries in duplicates.items():
        statuses = " vs ".join(f"Active={e['active']}/Status={e['status']}" for e in entries)
        discrepancies.append({
            "Code": code, "Style": entries[0]["style"],
            "Discrepancy Type": "Internal contradiction (new file)",
            "Field": "Active / Status",
            "Old File Value (FW28_2_Nov_2026.xlsx)": "n/a",
            "New File Value (F28_Acces_Tiering..._20261006.xlsx)": statuses,
            "Notes": "Same code appears twice in the new file with opposite Active/Status "
                     "values. Needs a decision from whoever owns this file -- not resolved here.",
        })

    old_codes, new_codes = set(old), set(new)
    common = old_codes & new_codes
    only_old = old_codes - new_codes

    new_styles_index: dict[str, str] = {}
    for c, e in new.items():
        new_styles_index.setdefault(e["style"], c)

    for code in only_old:
        o = old[code]
        if not o["active"]:
            continue  # only flag codes that were actually Active in the old file
        renamed_to = new_styles_index.get(o["style"])
        if renamed_to:
            discrepancies.append({
                "Code": code, "Style": o["style"],
                "Discrepancy Type": "Possibly renumbered",
                "Field": "Code",
                "Old File Value (FW28_2_Nov_2026.xlsx)": code,
                "New File Value (F28_Acces_Tiering..._20261006.xlsx)": renamed_to,
                "Notes": "Same exact style name found under a different code in the new file. "
                         "Confirm this is the same style before treating it as a straight rename.",
            })
        else:
            discrepancies.append({
                "Code": code, "Style": o["style"],
                "Discrepancy Type": "Missing from new file",
                "Field": "Code",
                "Old File Value (FW28_2_Nov_2026.xlsx)": f"Active, {o['layer']}",
                "New File Value (F28_Acces_Tiering..._20261006.xlsx)": "not found (by code or style name)",
                "Notes": "Active in the Nov file; no trace of this code or style name anywhere "
                         "in the new draft. Confirm whether it was dropped or just not in this draft yet.",
            })

    for code in sorted(common):
        o, n = old[code], new[code]
        if bool(o["active"]) != bool(n["active"]):
            discrepancies.append({
                "Code": code, "Style": o["style"],
                "Discrepancy Type": "Active status conflict",
                "Field": "Active",
                "Old File Value (FW28_2_Nov_2026.xlsx)": o["active"],
                "New File Value (F28_Acces_Tiering..._20261006.xlsx)": f"{n['active']} (Status: {n['status']})",
                "Notes": "",
            })
        if o["layer"] != n["layer"]:
            discrepancies.append({
                "Code": code, "Style": o["style"],
                "Discrepancy Type": "Layer reclassified",
                "Field": "Category/Layer",
                "Old File Value (FW28_2_Nov_2026.xlsx)": o["layer"],
                "New File Value (F28_Acces_Tiering..._20261006.xlsx)": n["layer"],
                "Notes": "",
            })
        if o["collection_type"] != n["collection_type"]:
            discrepancies.append({
                "Code": code, "Style": o["style"],
                "Discrepancy Type": "Collection Type reclassified",
                "Field": "Collection Type",
                "Old File Value (FW28_2_Nov_2026.xlsx)": o["collection_type"],
                "New File Value (F28_Acces_Tiering..._20261006.xlsx)": n["collection_type"],
                "Notes": "ABS <-> Inline changes main-range scope (RR-002/RR-006)." if
                         {o["collection_type"], n["collection_type"]} & {"ABS", "ABS Exclusive"} else "",
            })
        if o["gender"] != n["gender"]:
            discrepancies.append({
                "Code": code, "Style": o["style"],
                "Discrepancy Type": "Gender reclassified",
                "Field": "Gender",
                "Old File Value (FW28_2_Nov_2026.xlsx)": o["gender"],
                "New File Value (F28_Acces_Tiering..._20261006.xlsx)": n["gender"],
                "Notes": "",
            })

    discrepancies.sort(key=lambda r: (r["Discrepancy Type"], r["Code"]))

    new_only_rows = []
    for code in sorted(new_codes - old_codes):
        n = new[code]
        new_only_rows.append({
            "Code": code, "Style": n["style"], "Business Area": n["business_area"],
            "Layer": n["layer"], "Active": n["active"], "Status": n["status"],
        })
    for e in uncoded:
        new_only_rows.append({
            "Code": "(no code yet)", "Style": e["style"], "Business Area": e["business_area"],
            "Layer": e["layer"], "Active": e["active"], "Status": e["status"],
        })

    return discrepancies, new_only_rows


def write_workbook(discrepancies: list[dict], new_only_rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def write_sheet(title: str, rows: list[dict]) -> None:
        ws = wb.create_sheet(title)
        headers = list(rows[0].keys())
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
        for r in rows:
            ws.append(list(r.values()))
        for i, h in enumerate(headers, start=1):
            width = max(len(h), max((len(str(r[h])) for r in rows), default=0)) + 2
            ws.column_dimensions[get_column_letter(i)].width = min(width, 55)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    write_sheet("Discrepancies", discrepancies)
    write_sheet("Only in new draft (320)", new_only_rows)

    notes = wb.create_sheet("Notes")
    notes.append(["What this is", "Every discrepancy found comparing "
                                   "F28_Acces_Tiering_1st_Draft_20261006.xlsx (received 3-Nov-2026, "
                                   "the 'cleaned up' FW28 list) against FW28_2_Nov_2026.xlsx "
                                   "(already integrated into the franchise review list as RR-026). "
                                   "Nothing here has been integrated into the Range Review Board -- "
                                   "for the team to confirm first."])
    notes.append(["Date flag", "The new file's name ends 20261006 (6-Oct-2026) -- about a month "
                                "BEFORE FW28_2_Nov_2026.xlsx. 'Cleaned up' does not necessarily "
                                "mean 'more recent' here. Confirm which draft is actually current "
                                "before treating either as authoritative."])
    notes.append(["Internal contradiction (new file)", "4 codes (the L.I.M Airak 2,5L Proof II "
                                                          "Jacket/Pant Men/Women family) appear "
                                                          "TWICE in the new file's own FW28 sheet, "
                                                          "once CANCELED/inactive and once CO/active. "
                                                          "Not a cross-file issue -- two contradictory "
                                                          "rows in the same sheet."])
    notes.append(["Scope difference", "The new draft covers Apparel + Hardware + Footwear (611 "
                                       "coded styles + 8 proposed-but-uncoded 'NEW' styles); the "
                                       "old file was Apparel-only (292 codes). Most of the 320 "
                                       "'only in new draft' codes are explained by that wider scope, "
                                       "not a real mismatch -- listed on its own sheet for visibility, "
                                       "not mixed into the Discrepancies sheet."])
    notes.append(["Status", "Open -- every row needs a human confirmation. Don't integrate either "
                             "file into the Range Review Board, or treat one as replacing the "
                             "other, until the team has signed off on this list."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 22
    notes.column_dimensions["B"].width = 100

    wb.save(path)


if __name__ == "__main__":
    discrepancies, new_only_rows = build_rows()
    out_path = OUTPUTS_DIR / "FW28_Discrepancies.xlsx"
    write_workbook(discrepancies, new_only_rows, out_path)
    by_type = defaultdict(int)
    for r in discrepancies:
        by_type[r["Discrepancy Type"]] += 1
    print(f"{len(discrepancies)} discrepancies written to {out_path}:")
    for t, n in sorted(by_type.items()):
        print(f"  {t}: {n}")
    print(f"{len(new_only_rows)} codes only in the new draft (separate sheet, not counted above)")
