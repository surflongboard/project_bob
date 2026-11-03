"""
"Last season to sell" -- a SUGGESTED signal, not an answer. RR-021 found
no field anywhere in this workstream's existing sources that says when
an active franchise is planned to exit the range; the Board's own
dropdown (RR-022) stays a human workshop decision for that reason.

The Product master file (received 3-Nov-2026, data/inventory_planning/
inputs/product_master/) adds a ProductLastSeason field per code -- but
it's a TRAILING "last season this code was seen," not a forward-looking
exit plan, and the file's own coverage ends at SS28 (RR-026's FW28 data
postdates it). The user's own framing: a code tagged "last season SS28"
might genuinely be ending, OR might simply continue into FW28 without
the Product master export having caught up yet -- can't tell without
cross-checking FW28 activity directly.

This script does exactly that cross-check, per active-range franchise
(reusing data/inventory_planning/scripts/build_generation_release_timeline.py's
`build_generation_table()` for the per-franchise Last Known Season
roll-up/Release Season, and inventory_planning's `build_clearance_view.
load_style_seasonality()` for the range files' own Style Seasonality
field -- neither re-derived here) and produces ONE suggested signal per
franchise, in 7 buckets. Two refinements added 3-Nov-2026 after the
user pushed back on the first pass's framing:
  - "Active through FW28" (was "Continues into FW28") -- being active
    in FW28 doesn't mean the SS29/FW29 renewal-or-cut question is
    settled; the wording and note no longer imply it is.
  - A franchise active in SS28 but absent from FW28 is NOT automatically
    "last season: SS28" -- it's first checked against Style Seasonality
    (a Spring/Summer-only item has no Fall/Winter counterpart by design,
    so FW28 absence is expected, not an ending signal) and against its
    own Release Season (a generation that just launched in SS28 hasn't
    had a chance to appear in any FW season yet, so it's too early to
    read anything into it either). Only once both are ruled out does
    "Suggested last season: SS28" apply -- and even then, whether the
    franchise also had FW27 presence changes how strong that signal is
    (noted per row, not just asserted).
  - "Suggested last season: FW27" -- active in FW27 only; if the
    franchise's Style Seasonality is Fall/Winter, the note says so
    (skipping SS28 is then expected on its own; the real signal is
    also being absent from FW28).
  - "Check -- Product master shows an earlier exit than current activity"
    -- the mismatch case: franchise is active now but Last Known Season
    trails behind that, most likely the export simply hasn't caught up
    to the current season yet. Flagged for a sanity check, not trusted.
  - "No Product master match" -- the franchise's base name isn't found
    in the Product master file at all (e.g. brand-new FW28-only launch).

This is a DERIVED/SUGGESTED column only -- it does NOT write to the
Board's live "Last season to sell" dropdown (that stays a human
decision in the `decisions` database, RR-021/RR-022). Treat this as a
starting point for Merch to confirm or override during the workshop.

Run: python3 build_last_season_signal.py
Output: ../outputs/Last_Season_Signal.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RR_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = RR_DIR.parents[1]
sys.path.insert(0, str(RR_DIR / "scripts"))
from build_franchise_list import build_rows  # noqa: E402

IP_DIR = REPO_ROOT / "data" / "inventory_planning"
sys.path.insert(0, str(IP_DIR / "scripts"))
from build_generation_release_timeline import build_generation_table, season_key  # noqa: E402
from build_clearance_view import load_style_seasonality  # noqa: E402

OUTPUTS_DIR = RR_DIR / "outputs"


def build_signal_rows() -> list[dict]:
    active_rows = build_rows()
    gen_table = build_generation_table()
    current_gen_by_base = {r["Franchise"]: r for r in gen_table if r["Is Current Generation"] == "Y"}
    _, seasonality_by_base, _ = load_style_seasonality()

    out = []
    for r in active_rows:
        base = r["_base"]
        flags = []
        if r["Active in FW27"] == "Y":
            flags.append("FW27")
        if r["Active in SS28"] == "Y":
            flags.append("SS28")
        if r["Active in FW28"] == "Y":
            flags.append("FW28")

        gen_row = current_gen_by_base.get(base)
        last_known = gen_row["Last Known Season"] if gen_row else ""
        last_known_key = season_key(last_known) if last_known else None
        release_season = gen_row["Release Season"] if gen_row else ""
        seasonality = seasonality_by_base.get(base, "")

        if gen_row is None:
            signal = "No Product master match"
            note = "Franchise base name not found in the Product master file -- likely a brand-new " \
                   "launch (e.g. FW28-only) not yet covered by that export."
        elif "FW28" in flags:
            signal = "Active through FW28"
            note = ("Being active in FW28 doesn't settle whether it continues into SS29/FW29 -- that "
                     "renewal-or-cut decision is separate and still open (see the Generation age flag).")
            if last_known_key is not None and last_known_key < season_key("FW28"):
                note += (f" (Product master's Last Known Season, {last_known}, trails behind FW28 "
                          f"activity -- expected, since that file's own coverage ends at SS28.)")
        elif "SS28" in flags:
            if seasonality == "Spring/Summer":
                signal = "Active SS28 -- seasonal item, FW28 absence expected"
                note = ("Style Seasonality is Spring/Summer -- this franchise isn't expected to have "
                         "a Fall/Winter counterpart at all, so missing from FW28 isn't an ending signal "
                         "by itself; renewal/cut is still a question for SS29.")
            elif release_season == "SS28":
                signal = "New in SS28 -- too early to signal an ending"
                note = ("Product master's own Release Season for this generation is SS28 -- it just "
                         "launched, so absence from FW28 (or FW27) says nothing about ending; too "
                         "early to read as a last-season signal.")
            elif not last_known or last_known == "SS28":
                signal = "Suggested last season: SS28"
                if "FW27" in flags:
                    note = ("Had FW27 AND SS28 presence (not a Spring/Summer-only item) but isn't in "
                             "the FW28 plan -- the stronger of the two SS28 cases, since it held a full "
                             "FW+SS pattern before dropping the FW side.")
                else:
                    note = ("Not in FW27 either, and this isn't its first season (not a brand-new "
                             "SS28 launch) or a Spring/Summer-only item -- per the user's own check, "
                             "this is the likelier genuine 'last season' case, not just seasonality.")
            elif last_known_key is not None and last_known_key < season_key("SS28"):
                signal = "Check -- Product master shows an earlier exit than current activity"
                note = (f"Franchise is active in SS28 now, but Product master's Last Known Season "
                        f"is {last_known} (earlier) -- most likely the export hasn't caught up to "
                        f"the current season yet; sanity-check before trusting either way.")
            else:
                signal = "Check -- unexpected season value"
                note = f"Last Known Season ({last_known}) is later than SS28 but FW28 isn't flagged active."
        elif "FW27" in flags:
            signal = "Suggested last season: FW27"
            if seasonality == "Fall/Winter":
                note = ("Style Seasonality is Fall/Winter, so skipping SS28 is expected on its own -- "
                         "the real signal is that it's ALSO absent from FW28, the next season it would "
                         "actually appear in.")
            else:
                note = "Already absent from SS28/FW28 in Range Review's own active-range data."
        else:
            signal = "No active-range season flag"
            note = "Franchise has no FW27/SS28/FW28 active flag at all -- check the franchise list."

        out.append({
            "Franchise": r["Franchise"],
            "Layer": r["Layer"],
            "Active seasons (Range Review)": "; ".join(flags),
            "Style Seasonality": seasonality,
            "Last Known Season (Product master)": last_known,
            "Suggested signal": signal,
            "Note": note,
        })
    out.sort(key=lambda r: (r["Layer"], r["Franchise"]))
    return out


def write_workbook(rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Last Season Signal")
    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for r in rows:
        ws.append(list(r.values()))
    for i, h in enumerate(headers, start=1):
        width = max(len(h), max((len(str(r[h])) for r in rows), default=0)) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    notes = wb.create_sheet("Notes")
    notes.append(["What this is", "A SUGGESTED 'last season to sell' signal per active-range franchise, "
                                   "derived from the Product master file's ProductLastSeason field "
                                   "cross-checked against Range Review's own FW27/SS28/FW28 active "
                                   "flags. It does NOT auto-populate the Range Review Board's live "
                                   "'Last season to sell' dropdown (RR-021/RR-022) -- that stays a human "
                                   "workshop decision, captured in the Board's own `decisions` database. "
                                   "Use this as a starting point for Merch to confirm or override."])
    notes.append(["Why not just trust ProductLastSeason directly", "It's a TRAILING field -- the last "
                                                                      "season a code was observed -- not "
                                                                      "a forward-looking exit plan. A code "
                                                                      "tagged 'last season SS28' might be "
                                                                      "genuinely ending OR might simply "
                                                                      "continue into FW28 without the "
                                                                      "Product master export (whose own "
                                                                      "coverage ends at SS28) having "
                                                                      "caught up yet. This script resolves "
                                                                      "that ambiguity by checking FW28 "
                                                                      "activity directly rather than "
                                                                      "reading the field at face value."])
    notes.append(["The 6 signal buckets", "'Active through FW28' -- active in FW28 now; does NOT mean "
                                            "the SS29/FW29 renewal-or-cut question is settled, just that "
                                            "it hasn't dropped out yet. 'Suggested last season: SS28' -- "
                                            "active in SS28, not FW28, NOT a Spring/Summer-only style, and "
                                            "not a brand-new SS28 launch -- the genuine candidate. 'Active "
                                            "SS28 -- seasonal item, FW28 absence expected' -- Style "
                                            "Seasonality is Spring/Summer, so no Fall/Winter counterpart "
                                            "is expected anyway; not an ending signal. 'New in SS28 -- too "
                                            "early to signal an ending' -- this generation's own Release "
                                            "Season is SS28, so it hasn't had a chance to appear in any FW "
                                            "season yet. 'Suggested last season: FW27' -- active in FW27 "
                                            "only. 'Check -- Product master shows an earlier exit than "
                                            "current activity' -- a mismatch, most likely export lag. "
                                            "'No Product master match' -- base name not found in that "
                                            "file at all (e.g. brand-new FW28 launch)."])
    notes.append(["Refined 3-Nov-2026", "The first pass lumped every SS28-active/non-FW28 franchise "
                                          "into 'last season: SS28' -- the user correctly pushed back: "
                                          "that ignores Spring/Summer-only items (which never have an FW "
                                          "counterpart by design) and brand-new SS28 launches (too early "
                                          "to read anything into). Both are now split out as their own, "
                                          "non-alarming buckets using the range files' own Style "
                                          "Seasonality field and the generation's own Release Season."])
    notes.append(["Status", "Provisional -- a reasonable derivation, not independently confirmed per "
                             "franchise. RR-021 stays Open until Merch actually fills in the live "
                             "dropdown; this signal is meant to speed that up, not replace it."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 24
    notes.column_dimensions["B"].width = 100

    wb.save(path)


if __name__ == "__main__":
    rows = build_signal_rows()
    out_path = OUTPUTS_DIR / "Last_Season_Signal.xlsx"
    write_workbook(rows, out_path)

    from collections import Counter
    counts = Counter(r["Suggested signal"] for r in rows)
    print(f"{len(rows)} franchises written to {out_path}")
    for signal, n in counts.most_common():
        print(f"  {n:3d}  {signal}")
