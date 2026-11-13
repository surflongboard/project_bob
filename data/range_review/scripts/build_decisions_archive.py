"""
Freezes a point-in-time, protected copy of the Range Review Board's live
`decisions` database (RR-022/RR-097/RR-098/RR-099) once a review session
is done and the input should no longer change -- per the user, 12-Nov-
2026 ("I'd like to lock in the input so that it's not touched. Very
valuable input that I wouldn't be able to get again.") after the Nov 12
workshop concluded (RR-101).

Unlike build_decisions_export.py's "Range_Review_Decisions.xlsx" (a
*running* snapshot, meant to be refreshed and overwritten after every
layer as the workshop progresses), this script's output is a *dated,
sheet-protected archive* that is never overwritten -- a new workshop
gets a new dated file. Two outputs, same pull:
  - outputs/archive/decisions_raw_<date>.json -- the exact ArtifactData
    pull, byte-for-byte (list of {"id", "data"}), for a future script to
    re-derive anything from.
  - outputs/archive/Range_Review_Decisions_FINAL_<date>.xlsx -- the same
    columns as the running export, but with the sheet protected (Excel
    Review > Protect Sheet, no password -- this is a "don't edit this by
    accident" guard, not a security boundary) so it opens read-only in
    practice.

Run: python3 build_decisions_archive.py <path-to-decisions-pull.json> <date, e.g. 2026-11-12>
Output: ../outputs/archive/Range_Review_Decisions_FINAL_<date>.xlsx
        ../outputs/archive/decisions_raw_<date>.json
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

from openpyxl.worksheet.protection import SheetProtection

from build_decisions_export import build_rows, write_workbook

RR_DIR = Path(__file__).resolve().parents[1]
ARCHIVE_DIR = RR_DIR / "outputs" / "archive"


def main(pull_path: Path, date_str: str) -> None:
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)

    raw_out = ARCHIVE_DIR / f"decisions_raw_{date_str}.json"
    shutil.copy(pull_path, raw_out)

    rows = build_rows(pull_path)
    xlsx_out = ARCHIVE_DIR / f"Range_Review_Decisions_FINAL_{date_str}.xlsx"
    write_workbook(rows, xlsx_out)

    # Protect both sheets so the archive doesn't get casually edited --
    # all cells are locked by default; protection just enforces that.
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_out)
    for ws in wb.worksheets:
        ws.protection = SheetProtection(sheet=True, password=None, selectLockedCells=True)
    notes = wb["Notes"]
    notes.append([])
    notes.append(["Frozen archive", f"This file is a protected, point-in-time copy made {date_str}, "
                                      "once the workshop was done and the input needed to stop "
                                      "changing. It will never be overwritten by a later export -- "
                                      "a later workshop gets its own dated file. The sheet is "
                                      "protected (Excel Review > Unprotect Sheet, no password) so "
                                      "it isn't edited by accident; this is a guard against "
                                      "accidents, not a security boundary."])
    wb.save(xlsx_out)

    n_decided = sum(1 for r in rows if r["Tier"] != "Undecided")
    n_deep = sum(1 for r in rows if r["Deep review"] == "Yes")
    print(f"{len(rows)} franchises frozen to {xlsx_out}")
    print(f"Raw pull copied to {raw_out}")
    print(f"{n_decided} have a Tier decision, {n_deep} flagged Deep review")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 build_decisions_archive.py <path-to-decisions-pull.json> <date, e.g. 2026-11-12>")
        sys.exit(1)
    main(Path(sys.argv[1]), sys.argv[2])
