"""
Builds a working sheet for untangling the Nov 12 workshop's "Deep review"
flagged franchises (RR-099's checkbox) -- per the user, 13-Nov-2026: the
raw Action notes on these cards mix several topics at once (naming
confusion, consolidation calls, open questions addressed to a named
person, timing/TBD, pricing/color review, ...) and need sorting one by
one, without touching the frozen original (RR-101).

Each flagged franchise's original comment is carried over VERBATIM into
a locked column -- this sheet is for untangling it, not rewriting it.
Everything to the right (Topic, Owner, Action needed, Resolved?, Notes)
is blank and unlocked for the user to fill in. The sheet is NOT
protected (unlike the frozen archive) since it's meant to be worked in;
only the original-data columns are cell-locked, so a careless edit still
can't overwrite the source comment, but the sheet itself isn't locked.

Run: python3 build_deep_review_worklist.py <path-to-decisions-pull.json>
Output: ../outputs/archive/Deep_Review_Worklist_<date>.xlsx
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.protection import SheetProtection

RR_DIR = Path(__file__).resolve().parents[1]
OUTPUTS_DIR = RR_DIR / "outputs"
ARCHIVE_DIR = OUTPUTS_DIR / "archive"
WORKSHOP_DATA_PATH = OUTPUTS_DIR / "workshop_data.json"

LOCKED_HEADERS = ["Layer", "Franchise", "Tier", "Flags", "Original comment (frozen)"]
WORK_HEADERS = ["Topic(s) untangled", "Owner to follow up", "Action needed", "Resolved?", "Notes"]


def fmt_updated(ms: int | None) -> str:
    if not ms:
        return ""
    return datetime.fromtimestamp(ms / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def build_rows(decisions_pull_path: Path) -> list[dict]:
    franchises = json.loads(WORKSHOP_DATA_PATH.read_text())
    by_base = {f["base"]: f for f in franchises}

    rows = []
    for doc in json.loads(decisions_pull_path.read_text()):
        d = doc["data"]
        if not d.get("deepReview"):
            continue
        f = by_base.get(d.get("base", doc["id"]), {})
        flags = []
        if d.get("nordicAbs"):
            flags.append("Nordic ABS")
        if d.get("qtmCandidate"):
            flags.append("QTM candidate")
        if d.get("consolidate"):
            flags.append("Consolidate")
        rows.append({
            "Layer": f.get("layer", ""),
            "Franchise": d.get("franchise") or doc["id"],
            "Tier": d.get("tier") or "Undecided",
            "Flags": ", ".join(flags),
            "Original comment (frozen)": d.get("action") or "",
            "Topic(s) untangled": "",
            "Owner to follow up": "",
            "Action needed": "",
            "Resolved?": "",
            "Notes": "",
        })
    rows.sort(key=lambda r: (r["Layer"], r["Franchise"]))
    return rows


def write_workbook(rows: list[dict], path: Path, date_str: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Deep Review")
    headers = LOCKED_HEADERS + WORK_HEADERS
    ws.append(headers)
    for i, cell in enumerate(ws[1], start=1):
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="AE4B3E" if i <= len(LOCKED_HEADERS) else "2F7A6B")

    for r in rows:
        ws.append([r[h] for h in headers])

    n_locked = len(LOCKED_HEADERS)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i, cell in enumerate(row, start=1):
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.protection = openpyxl.styles.Protection(locked=(i <= n_locked))

    widths = {"Layer": 14, "Franchise": 26, "Tier": 11, "Flags": 20,
              "Original comment (frozen)": 55, "Topic(s) untangled": 24,
              "Owner to follow up": 18, "Action needed": 32, "Resolved?": 11, "Notes": 28}
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(h, 18)
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = ws.dimensions
    # Sheet is NOT protected (meant to be worked in) -- only the individual
    # original-data cells are locked, which has no effect unless/until the
    # user turns on sheet protection themselves (Review > Protect Sheet).

    notes = wb.create_sheet("Notes")
    notes.append(["What this is", f"A worklist for untangling the {len(rows)} franchises flagged "
                                    "\"Deep review\" at the Nov 12 workshop -- their Action notes "
                                    "mix several topics at once. The first five columns are a "
                                    "frozen copy of the original comment from that day; the last "
                                    "five are blank for you to fill in as you sort each one out."])
    notes.append(["Protecting the original", "The first five columns' cells are individually "
                                                "locked. That only takes effect if you turn on "
                                                "Review > Protect Sheet yourself -- it's off by "
                                                "default so you can work in the other columns "
                                                "freely. Turn it on once you're confident you want "
                                                "the original comments to stay fixed while you edit."])
    notes.append(["Source", f"Pulled from the Board's live `decisions` database, {date_str}, "
                              "after the Board itself was locked to admin-write (RR-101) -- "
                              "these 79 comments won't change underneath you."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 22
    notes.column_dimensions["B"].width = 100

    wb.save(path)


def main(pull_path: Path, date_str: str) -> None:
    rows = build_rows(pull_path)
    out_path = ARCHIVE_DIR / f"Deep_Review_Worklist_{date_str}.xlsx"
    write_workbook(rows, out_path, date_str)
    print(f"{len(rows)} franchises written to {out_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 build_deep_review_worklist.py <path-to-decisions-pull.json> <date, e.g. 2026-11-12>")
        sys.exit(1)
    main(Path(sys.argv[1]), sys.argv[2])
