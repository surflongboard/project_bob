"""
Exports the Range Review Board's live workshop decisions (Tier / Last
season to sell / Consolidate / Action notes, captured in the artifact's
own `decisions` database, RR-022) into an Excel sign-off sheet.

**Not a fully automated script** -- there is no standing credential for
a local script to query the artifact's live database directly. The
actual export is a two-step, Claude-assisted process each time it's
needed:
  1. Claude pulls the current `decisions` collection via the
     ArtifactData tool (`action: "list"`, `collection: "decisions"`)
     and saves it to a local JSON file (list of `{"id", "data"}`).
  2. This script merges that pull against the franchise list
     (`outputs/workshop_data.json`, for Layer/Business Area/display
     name -- the DB only stores what the Board itself writes) and
     writes the Excel.

Run: python3 build_decisions_export.py <path-to-decisions-pull.json>
Output: ../outputs/Range_Review_Decisions.xlsx
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RR_DIR = Path(__file__).resolve().parents[1]
OUTPUTS_DIR = RR_DIR / "outputs"
WORKSHOP_DATA_PATH = OUTPUTS_DIR / "workshop_data.json"


def fmt_updated(ms: int | None) -> str:
    if not ms:
        return ""
    return datetime.fromtimestamp(ms / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def build_rows(decisions_pull_path: Path) -> list[dict]:
    franchises = json.loads(WORKSHOP_DATA_PATH.read_text())
    by_base = {f["base"]: f for f in franchises}

    decisions_by_base = {}
    for doc in json.loads(decisions_pull_path.read_text()):
        d = doc["data"]
        decisions_by_base[d.get("base", doc["id"])] = d

    rows = []
    for f in franchises:
        dec = decisions_by_base.get(f["base"], {})
        rows.append({
            "Layer": f["layer"],
            "Business Area": f["businessArea"],
            "Franchise": f["franchise"],
            "Tier": dec.get("tier") or "Undecided",
            "Last season to sell": dec.get("lastSeasonToSell") or "Undecided",
            "Consolidate with other models": "Yes" if dec.get("consolidate") else "",
            "Action / notes": dec.get("action") or "",
            "Decision updated": fmt_updated(dec.get("updatedAt")),
        })
    rows.sort(key=lambda r: (r["Layer"], r["Business Area"], r["Franchise"]))
    return rows


def write_workbook(rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Decisions")
    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for r in rows:
        ws.append(list(r.values()))
    for i, h in enumerate(headers, start=1):
        width = max(len(h), max((len(str(r[h])) for r in rows), default=0)) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 45)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    notes = wb.create_sheet("Notes")
    notes.append(["What this is", "A point-in-time export of the Range Review Board's live "
                                    "workshop decisions (Tier / Last season to sell / "
                                    "Consolidate / Action notes) -- the Board's own `decisions` "
                                    "database (RR-022) is the live source of truth; this file is "
                                    "a snapshot as of when it was pulled, not a live link."])
    notes.append(["How to refresh", "Ask Claude to re-export -- it pulls the current `decisions` "
                                      "collection via the ArtifactData tool and re-runs this "
                                      "script. There is no automated/scheduled export; it's a "
                                      "deliberate pull each time."])
    notes.append(["\"Undecided\"", "A franchise with no Tier or Last-season-to-sell entry yet "
                                     "shows \"Undecided\", not blank -- distinguishes a franchise "
                                     "nobody has reviewed yet from one Merch decided to leave "
                                     "blank on purpose (the Board itself still shows blank for "
                                     "both cases; this export makes the distinction explicit)."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 20
    notes.column_dimensions["B"].width = 100

    wb.save(path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 build_decisions_export.py <path-to-decisions-pull.json>")
        sys.exit(1)
    rows = build_rows(Path(sys.argv[1]))
    out_path = OUTPUTS_DIR / "Range_Review_Decisions.xlsx"
    write_workbook(rows, out_path)
    n_decided = sum(1 for r in rows if r["Tier"] != "Undecided")
    print(f"{len(rows)} franchises written to {out_path}")
    print(f"{n_decided} have a Tier decision so far")
