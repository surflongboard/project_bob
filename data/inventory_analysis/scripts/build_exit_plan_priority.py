"""
Sell-Down Priority: joins the FW27 range assortment plan's planned exit
season to current warehouse available stock, so stock sitting on a
franchise that's about to leave the collection surfaces BEFORE it retires,
not after.

This is a different signal from analyze_inventory.py's Season-Recency view
(REG-INV-008/009), and deliberately complements rather than replaces it:
  - analyze_inventory.py asks "how OLD is this stock?" (backward-looking,
    from the stock file's own Season Article code on each unit).
  - This script asks "how much RUNWAY does this stock's franchise have
    left in the collection?" (forward-looking, from the assortment plan's
    own planned exit season for that style).
A franchise can be low on the first (nearly-new stock) and high-risk on
the second (already planned to exit next season) — that's exactly the
case this script is built to catch early enough to act on.

Source: data/inventory_analysis/inputs/assortment/Assortment Attribution
Review_11092026.xlsx, FW27 tab only (per the business's own framing: FW27
is "the workable season" — see ASSUMPTIONS_REGISTER.md REG-INV-010). Two
columns carry the signal, both business-identified by column letter:
  - Column AI, "LSO / Exit Season" — the planned Last Season Offered.
  - Column AN, "NEW MAPPING: Activity" — a merchandising Activity grouping
    (Day Hiking, Trekking, Active Essentials, ...), carried through for
    context/filtering, not used in the ranking itself.

Matched to franchise (Base+Gender) via ss26_lib.franchise_key() on the
FW27 tab's own `Style` text — the tab's `Franchise/Family` column is
almost entirely blank (REG-INV-011), so this is the only usable join key,
same fallback already used for the stock file itself.

CAVEATS (see ASSUMPTIONS_REGISTER.md for the full reasoning):
  - Only 392 of ~1,860 published franchises appear on the FW27 tab at
    all — this plan is scoped to FW27's own active/carry-over/new range,
    not the full historical franchise universe. Stock on a franchise with
    no FW27-tab row is reported separately (Sheet 4), not silently
    dropped and not assumed safe.
  - Sheet 4 is NOT safe to blanket-label "retired before FW27" (REG-INV-012):
    it's cross-checked against the SS27 tab (the season immediately before
    FW27), and a real share of it was still Active/Carry-Over there. That
    subset needs a merch check, not an assumption — see the sheet's own
    flag column and REG-INV-012 for two confirmed examples where the
    likely explanation is a name change (e.g. "Rosson Softshell Hood" ->
    "Rosson Mid II Hood"), not a real exit, which Base+Gender matching
    can't catch (a known limitation, documented in config.py itself).
  - "REVIEW" (Exit Season undecided) is NOT force-ranked into the sell-
    down priority list — REG-INV-010 documents this as a genuine Open
    item per this repo's register discipline. Those franchises get their
    own sheet (Sheet 3) so the stock at stake is visible without
    pretending the timing is known.
  - "FW27 or SS28" (an explicitly ambiguous exit season in the source
    data) is kept as its own tier rather than collapsed into either
    season — see REG-INV-010.

Usage:
    python3 data/inventory_analysis/scripts/build_exit_plan_priority.py \
        [--assortment PATH] [--stock PATH]
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402

WORKSTREAM_DIR = REPO_ROOT / "data" / "inventory_analysis"
DEFAULT_ASSORTMENT = WORKSTREAM_DIR / "inputs" / "assortment" / "Assortment Attribution Review_11092026.xlsx"
DEFAULT_STOCK = REPO_ROOT / "data" / "ss26_portfolio_tiering" / "inputs" / "stock" / "Available_stock_260825.xlsx"
ASSORTMENT_SHEET = "FW27"
PRIOR_SEASON_SHEET = "SS27"  # cross-check for Sheet 4 -- see REG-INV-012

# --------------------------------------------------------------------------
# REG-INV-010: exit-urgency tiers, read directly off the FW27 tab's own
# "LSO / Exit Season" values (column AI) — only these five values exist in
# the current pull, so this is an explicit lookup, not season arithmetic.
# Ranked list order = TIER_RANK order; "Pending decision" is NEVER placed
# in the ranked list (see module docstring).
# --------------------------------------------------------------------------
TIER_RANK = {
    "FW27": (1, "Exiting FW27 (last season offered)"),
    "FW27 or SS28": (2, "Exiting FW27 or SS28 (ambiguous — treat as near-term)"),
    "SS28": (3, "One season left (exits after SS28)"),
    "FW28+": (4, "Long runway (continues through FW28+)"),
}
PENDING_VALUE = "REVIEW"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--assortment", default=str(DEFAULT_ASSORTMENT))
    ap.add_argument("--stock", default=str(DEFAULT_STOCK))
    args = ap.parse_args()

    # --- FW27 exit plan ---
    wb_a = openpyxl.load_workbook(args.assortment, data_only=True)
    ws_a = wb_a[ASSORTMENT_SHEET]
    rows = list(ws_a.iter_rows(values_only=True))
    hdr = rows[1]
    idx = {h: i for i, h in enumerate(hdr) if h}
    for col in ("Code", "Style", "LSO / Exit Season ", "NEW MAPPING: Activity", "Active"):
        if col not in idx:
            sys.exit(f"Expected column '{col}' not found on {ASSORTMENT_SHEET} header: {hdr}")
    plan = {}
    for r in rows[2:]:
        if r[idx["Code"]] in (None, ""):
            continue
        key = lib.franchise_key(str(r[idx["Style"]]))
        plan[key] = {
            "style": r[idx["Style"]],
            "active": r[idx["Active"]],
            "exit_season": r[idx["LSO / Exit Season "]],
            "activity": r[idx["NEW MAPPING: Activity"]],
        }
    print(f"FW27 exit-plan rows: {len(plan)}")

    # --- SS27 tab, for the Sheet-4 "actually retired?" cross-check (REG-INV-012) ---
    ws_prior = wb_a[PRIOR_SEASON_SHEET]
    prows = list(ws_prior.iter_rows(values_only=True))
    phdr = prows[1]
    pidx = {h: i for i, h in enumerate(phdr) if h}
    prior_season = {}
    for r in prows[2:]:
        if r[pidx["Code"]] in (None, ""):
            continue
        key = lib.franchise_key(str(r[pidx["Style"]]))
        prior_season[key] = {"active": r[pidx["Active"]], "status": r[pidx.get("STATUS", -1)]}
    print(f"{PRIOR_SEASON_SHEET} rows (for cross-check): {len(prior_season)}")

    # --- current stock, same matching as analyze_inventory.py / update_stock.py ---
    wb_s = openpyxl.load_workbook(args.stock, data_only=True, read_only=True)
    candidates = [n for n in wb_s.sheetnames if n.lower().startswith("available stock")]
    ws_s = wb_s[candidates[0]] if candidates else wb_s[wb_s.sheetnames[0]]
    srows = list(ws_s.iter_rows(values_only=True))
    shdr = srows[0]
    sidx = {h: i for i, h in enumerate(shdr)}
    stock_units = defaultdict(float)
    for r in srows[2:]:
        article = r[sidx["Article"]]
        units = r[sidx["Available stock"]] or 0
        if not article:
            continue
        stock_units[lib.franchise_key(str(article).strip())] += units
    print(f"stock franchises with units: {len(stock_units)}")

    # --- join ---
    ranked = []       # franchises with stock, matched to a ranked exit tier
    pending = []       # franchises with stock, exit season = REVIEW
    unplanned = []     # franchises with stock, no FW27-tab row at all
    for key, units in stock_units.items():
        if units <= 0:
            continue
        info = plan.get(key)
        if info is None:
            unplanned.append((key, units))
            continue
        exit_season = info["exit_season"]
        if exit_season == PENDING_VALUE:
            pending.append((key, units, info))
        elif exit_season in TIER_RANK:
            tier_num, tier_label = TIER_RANK[exit_season]
            ranked.append((key, units, info, tier_num, tier_label))
        else:
            # An exit-season value outside the five known ones -- don't guess
            # its urgency, route it alongside "pending" so it's visible.
            pending.append((key, units, {**info, "exit_season": f"UNRECOGNIZED: {exit_season!r}"}))

    ranked.sort(key=lambda x: (x[3], -x[1]))
    pending.sort(key=lambda x: -x[1])
    unplanned.sort(key=lambda x: -x[1])

    print(f"matched to a ranked exit tier: {len(ranked)}")
    print(f"pending merch decision (REVIEW): {len(pending)}")
    print(f"no FW27-tab row at all: {len(unplanned)}")
    still_active_ss27 = [(k, u) for k, u in unplanned if prior_season.get(k, {}).get("active") is True]
    print(f"  of which still Active in {PRIOR_SEASON_SHEET}: {len(still_active_ss27)} "
          f"({sum(u for k, u in still_active_ss27):,.0f} units) -- needs a merch check, not an assumption")
    total_stock = sum(stock_units.values())
    print(f"total stock units: {total_stock:,.0f}  "
          f"ranked: {sum(x[1] for x in ranked):,.0f}  "
          f"pending: {sum(x[1] for x in pending):,.0f}  "
          f"unplanned: {sum(x[1] for x in unplanned):,.0f}")

    # ---------------------------------------------------------------------
    # Build the workbook
    # ---------------------------------------------------------------------
    s = lib.styles()
    HDR_FONT, HDR_FILL, BODY_FONT, GRAY_FONT = s["header_font"], s["header_fill"], s["body_font"], s["gray_font"]
    TITLE_FONT = Font(name="Arial", bold=True, color="FF1F3864")
    WRAP = Alignment(vertical="top", wrap_text=True)
    NUM_FMT = s["num_fmt"]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet("1. Overview")
    ws1["A1"] = "Sell-Down Priority — FW27 Exit Plan x Current Stock"
    ws1["A1"].font = TITLE_FONT
    ws1.merge_cells("A1:F1")
    ws1["A3"] = (
        "Joins the FW27 range assortment plan's own planned exit season (\"LSO / Exit Season\", "
        "column AI) to current warehouse stock, so a franchise already planned to leave the "
        "collection surfaces here BEFORE it retires -- not discovered afterward as dead stock. "
        "Matched to franchise (Base+Gender) via ss26_lib.franchise_key() on the assortment plan's "
        "own Style text (its Franchise/Family column is not usable -- see REG-INV-011)."
    )
    ws1["A3"].font = BODY_FONT
    ws1["A3"].alignment = WRAP
    ws1.merge_cells("A3:F3")
    ws1.row_dimensions[3].height = 75
    ws1["A6"] = (
        f"{total_stock:,.0f} total stock units across all franchises with any available stock. "
        f"{sum(x[1] for x in ranked):,.0f} matched a franchise with a known FW27 exit-season tier "
        f"(Sheet 2). {sum(x[1] for x in pending):,.0f} matched a franchise still at \"REVIEW\" -- "
        "exit timing undecided, NOT force-ranked (Sheet 3). "
        f"{sum(x[1] for x in unplanned):,.0f} has no FW27-tab row at all (Sheet 4) -- of which "
        f"{sum(u for k, u in still_active_ss27):,.0f} units ({len(still_active_ss27)} franchises) were "
        f"still Active in {PRIOR_SEASON_SHEET} and are NOT safe to assume retired (REG-INV-012)."
    )
    ws1["A6"].font = GRAY_FONT
    ws1["A6"].alignment = WRAP
    ws1.merge_cells("A6:F6")
    ws1.row_dimensions[6].height = 75
    for col, w in zip("ABCDEF", (20,) * 6):
        ws1.column_dimensions[col].width = w

    def write_table(ws, headers, col_widths, title, note=None):
        ws["A1"] = title
        ws["A1"].font = TITLE_FONT
        ws.merge_cells(f"A1:{openpyxl.utils.get_column_letter(len(headers))}1")
        start = 3
        if note:
            ws["A3"] = note
            ws["A3"].font = GRAY_FONT
            ws["A3"].alignment = WRAP
            ws.merge_cells(f"A3:{openpyxl.utils.get_column_letter(len(headers))}3")
            ws.row_dimensions[3].height = 45
            start = 5
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=start, column=c, value=h)
            cell.font, cell.fill = HDR_FONT, HDR_FILL
        for c, w in enumerate(col_widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
        return start + 1

    # --- Sheet 2: ranked sell-down priority ---
    ws2 = wb.create_sheet("2. Sell-Down Priority")
    r = write_table(
        ws2,
        ["Priority Tier", "Base", "Gender", "Activity", "Exit Season", "Units in Stock", "Active?"],
        (34, 30, 10, 20, 16, 16, 10),
        "Sell-Down Priority (most urgent exit, most stock, first)",
        "Sorted by exit urgency (Sheet 1's tier order), then by units in stock, descending. "
        "\"FW28+\" rows are long-runway, low-urgency -- included for completeness, not action.",
    )
    for key, units, info, tier_num, tier_label in ranked:
        base, g = key
        ws2.cell(row=r, column=1, value=tier_label).font = BODY_FONT
        ws2.cell(row=r, column=2, value=base).font = BODY_FONT
        ws2.cell(row=r, column=3, value=g).font = BODY_FONT
        ws2.cell(row=r, column=4, value=info["activity"] or "").font = BODY_FONT
        ws2.cell(row=r, column=5, value=info["exit_season"]).font = BODY_FONT
        c = ws2.cell(row=r, column=6, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
        ws2.cell(row=r, column=7, value=bool(info["active"])).font = BODY_FONT
        r += 1

    # --- Sheet 3: pending merch decision ---
    ws3 = wb.create_sheet("3. Pending Review")
    r = write_table(
        ws3,
        ["Base", "Gender", "Activity", "Exit Season (raw)", "Units in Stock", "Active?"],
        (34, 10, 20, 22, 16, 10),
        "Pending Merch Decision — Exit Timing Not Yet Set",
        "These franchises carry real stock but the FW27 plan hasn't fixed an exit season yet "
        "(\"REVIEW\"). Not ranked above -- ranking these would mean guessing a business decision "
        "that hasn't been made (REG-INV-010). Sorted by units in stock, descending.",
    )
    for key, units, info in pending:
        base, g = key
        ws3.cell(row=r, column=1, value=base).font = BODY_FONT
        ws3.cell(row=r, column=2, value=g).font = BODY_FONT
        ws3.cell(row=r, column=3, value=info["activity"] or "").font = BODY_FONT
        ws3.cell(row=r, column=4, value=info["exit_season"]).font = BODY_FONT
        c = ws3.cell(row=r, column=5, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
        ws3.cell(row=r, column=6, value=bool(info["active"])).font = BODY_FONT
        r += 1

    # --- Sheet 4: unplanned stock (no FW27-tab row), cross-checked against SS27 ---
    still_active_prior = sum(1 for k, u in unplanned if prior_season.get(k, {}).get("active") is True)
    still_active_prior_units = sum(u for k, u in unplanned if prior_season.get(k, {}).get("active") is True)
    ws4 = wb.create_sheet("4. Not on FW27 Plan")
    r = write_table(
        ws4,
        ["Base", "Gender", "Units in Stock", f"In {PRIOR_SEASON_SHEET} Tab?", f"{PRIOR_SEASON_SHEET} Active?", f"{PRIOR_SEASON_SHEET} Status"],
        (34, 10, 16, 14, 14, 12),
        "Stock With No Row on the FW27 Tab",
        "NOT safe to blanket-label \"retired before FW27\" (REG-INV-012): cross-checked against "
        f"{PRIOR_SEASON_SHEET} (the season immediately before FW27) below. "
        f"{still_active_prior} franchises ({still_active_prior_units:,.0f} units) were still Active "
        f"there and need a real merch check -- likely explanations include a genuine drop, but also "
        "a name change too large for Base+Gender matching to catch (e.g. \"Rosson Softshell Hood\" -> "
        f"\"Rosson Mid II Hood\", confirmed both exist, neither name matches the other). Rows with "
        f"\"{PRIOR_SEASON_SHEET} Active? = False\" or blank (not in {PRIOR_SEASON_SHEET} either) are the "
        "safer default for \"likely already retired.\" Sorted by units in stock, descending.",
    )
    for key, units in unplanned:
        base, g = key
        prior = prior_season.get(key)
        ws4.cell(row=r, column=1, value=base).font = BODY_FONT
        ws4.cell(row=r, column=2, value=g).font = BODY_FONT
        c = ws4.cell(row=r, column=3, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
        ws4.cell(row=r, column=4, value="Yes" if prior else "No").font = BODY_FONT
        ws4.cell(row=r, column=5, value=(prior["active"] if prior else None)).font = BODY_FONT
        ws4.cell(row=r, column=6, value=(prior["status"] if prior else None)).font = BODY_FONT
        r += 1

    out_name = f"Project_Bob_Sell_Down_Priority_{date.today().strftime('%d%m%Y')}.xlsx"
    out_path = WORKSTREAM_DIR / out_name
    wb.save(out_path)
    print("saved:", out_path)


if __name__ == "__main__":
    main()
