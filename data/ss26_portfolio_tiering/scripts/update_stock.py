"""
REG-019: rebuilds Sheet 4 "Stock by Tier" from a warehouse available-stock
snapshot.

Reads inputs/stock/Available_stock_260825.xlsx (or whichever file is
passed via --source), matches each row's Article field to a franchise via
the same base_name/gender parsing used everywhere else, joins to that
franchise's current Tier (from Sheet 2), and produces:
  - a summary table (units + SKU count + % of matched stock + a
    recommended action per tier)
  - detail tables for Thin/Immaterial, Exited, and Problem Child (the
    last one included for CONTRAST, not as a clearance candidate --
    Problem Child needs a pricing/cost fix, not a fire sale)

CAVEAT: this is UNITS only. The source file has no cost/price column, so
no SEK stock value is computed -- don't multiply by an average price
without checking with Finance; unit economics vary widely by article.

ASSUMPTION -- column layout: Article is column index 6 (0-indexed, i.e.
column G) and available-stock units is column index 10 (column K) in the
source file, with data starting row 3. This is NOT derived from a header
lookup (the source file's header row doesn't cleanly name these columns)
-- if a future stock file has a different layout, these indices will need
updating; the script will silently mismatch columns rather than error, so
sanity-check the printed match-rate against the file before trusting the
output.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_stock.py \
        [--source PATH] [--sheet NAME]
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import ss26_lib as lib

ARTICLE_COL_IDX = 6   # 0-indexed -- see docstring ASSUMPTION
STOCK_COL_IDX = 10    # 0-indexed
DATA_MIN_ROW = 3

TIER_ORDER = lib.TIER_ORDER
ACTIONS = {
    "Hero + Near-Hero": "Protect availability — top performer; prioritize replenishment, avoid stockouts.",
    "Workhorse + Harvest": "Maintain — steady sellers; standard replenishment cadence, no special action.",
    "Problem Child": "Fix or watch, not clear — real revenue exists; review pricing/cost before reordering. Large stock on a real seller is a margin-fix candidate, not a clearance one.",
    "New / Test": "Monitor sell-through — too early to judge; don't over-commit to further stock yet.",
    "Thin / Immaterial": "CLEAR VIA SALE — below materiality threshold; liquidate rather than carry, especially SKUs with ~zero FY25 sales sitting on real stock.",
    "Exited": "CLEAR VIA SALE — discontinued; liquidate remaining stock, expect it to trend to zero.",
    "Clearance — Ex China/Zalando": "Already earmarked — this stock should be flowing to the close-out channel; monitor sell-through, don't reorder.",
}
DEAD_STOCK_SALES_THRESHOLD = 1000  # SEK -- below this, FY25 sales treated as ~zero for the "dead stock" signal
TOP_N_DETAIL = 10


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=str(lib.INPUTS_DIR / "stock" / "Available_stock_260825.xlsx"))
    ap.add_argument("--sheet", default=None, help="defaults to the first sheet whose name starts with 'Available stock', else the workbook's first sheet")
    args = ap.parse_args()

    wb_stock = openpyxl.load_workbook(args.source, data_only=True)
    if args.sheet:
        ws_stock = wb_stock[args.sheet]
    else:
        candidates = [n for n in wb_stock.sheetnames if n.lower().startswith("available stock")]
        ws_stock = wb_stock[candidates[0]] if candidates else wb_stock[wb_stock.sheetnames[0]]
    rows = list(ws_stock.iter_rows(min_row=DATA_MIN_ROW, values_only=True))

    wb, ws2 = lib.load_full_portfolio()
    pub_lookup = {}
    for r in range(lib.FIRST_DATA_ROW, ws2.max_row + 1):
        b = ws2.cell(row=r, column=1).value
        if b is None:
            continue
        g = (ws2.cell(row=r, column=2).value or "").strip()
        pub_lookup[(b.strip(), g)] = {
            "tier": ws2.cell(row=r, column=4).value,
            "sales25": ws2.cell(row=r, column=6).value or 0,
        }

    franchise_units = defaultdict(float)
    total_rows = matched_rows = 0
    for row in rows:
        article = row[ARTICLE_COL_IDX]
        stock = row[STOCK_COL_IDX] or 0
        if not article:
            continue
        total_rows += 1
        key = lib.franchise_key(article.strip())
        if key not in pub_lookup:
            continue
        franchise_units[key] += stock
        matched_rows += 1

    total_units = sum(row[STOCK_COL_IDX] or 0 for row in rows if row[ARTICLE_COL_IDX])
    matched_units = sum(franchise_units.values())
    pct_matched = matched_units / total_units * 100 if total_units else 0
    print(f"stock rows: {total_rows}  matched an existing franchise: {matched_rows}")
    print(f"total units: {total_units:,.0f}  matched units: {matched_units:,.0f} ({pct_matched:.1f}%)")

    tier_units = Counter()
    for key, units in franchise_units.items():
        tier_units[pub_lookup[key]["tier"]] += units

    if lib.STOCK_SHEET in wb.sheetnames:
        del wb[lib.STOCK_SHEET]
    ws = wb.create_sheet(lib.STOCK_SHEET)

    TITLE_FONT = Font(name="Arial", bold=True, color="FF1F3864")
    s = lib.styles()
    HDR_FONT, HDR_FILL, BODY_FONT, GRAY_FONT = s["header_font"], s["header_fill"], s["body_font"], s["gray_font"]
    ACTION_FILL = PatternFill(start_color="FFF3E4D2", end_color="FFF3E4D2", fill_type="solid")
    DEAD_FILL = ACTION_FILL
    WRAP = Alignment(vertical="top", wrap_text=True)
    NUM_FMT, PCT_FMT = s["num_fmt"], s["pct_fmt"]

    ws["A1"] = "Available Stock by Tier — What's Actionable"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:F1")

    src_name = args.source.split("/")[-1]
    ws["A3"] = (
        f"Source: {src_name} (warehouse available-stock snapshot, per its own filename). Matched to franchise "
        "(Base+Gender) via the Article field's own gender/version-token parsing (same method as REG-010/014/015), "
        f"then joined to each franchise's current consolidated Tier. {total_units:,.0f} total units in the source "
        f"file; {matched_units:,.0f} ({pct_matched:.1f}%) matched an existing published franchise. The remainder "
        "are styles not yet in the tiering file — likely brand-new FW27 launches with no FY24/25 history to tier "
        "against yet."
    )
    ws["A3"].font = BODY_FONT
    ws["A3"].alignment = WRAP
    ws.merge_cells("A3:F3")
    ws.row_dimensions[3].height = 62

    ws["A5"] = (
        "CAVEAT: this is UNITS only — the source file has no cost/price column, so no SEK stock value is computed "
        "here. Don't multiply by an average price without checking with Finance; unit economics vary widely by article."
    )
    ws["A5"].font = GRAY_FONT
    ws["A5"].alignment = WRAP
    ws.merge_cells("A5:F5")

    hdr_row = 7
    headers = ["Tier", "# SKUs (distinct)", "Total Units in Stock", "% of Matched Stock", "Recommended Action"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=hdr_row, column=c, value=h)
        cell.font, cell.fill = HDR_FONT, HDR_FILL

    r = hdr_row + 1
    for tier in TIER_ORDER:
        n_skus = sum(1 for key in franchise_units if pub_lookup[key]["tier"] == tier)
        units = tier_units.get(tier, 0)
        pct = units / matched_units * 100 if matched_units else 0
        ws.cell(row=r, column=1, value=tier).font = BODY_FONT
        ws.cell(row=r, column=2, value=n_skus).font = BODY_FONT
        c = ws.cell(row=r, column=3, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
        c = ws.cell(row=r, column=4, value=round(pct, 1)); c.number_format, c.font = PCT_FMT, BODY_FONT
        action_cell = ws.cell(row=r, column=5, value=ACTIONS[tier])
        action_cell.font, action_cell.alignment = BODY_FONT, WRAP
        if tier in ("Thin / Immaterial", "Exited"):
            for col in range(1, 6):
                ws.cell(row=r, column=col).fill = ACTION_FILL
        r += 1
    r += 1

    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 26, 16
    ws.column_dimensions["C"].width, ws.column_dimensions["D"].width = 18, 16
    ws.column_dimensions["E"].width = 60
    for rr in range(hdr_row + 1, r):
        ws.row_dimensions[rr].height = 45

    # --- detail tables ---
    def top_items(tier_name):
        items = [(k, franchise_units[k], pub_lookup[k]["sales25"]) for k in franchise_units if pub_lookup[k]["tier"] == tier_name]
        items.sort(key=lambda x: -x[1])
        return items[:TOP_N_DETAIL]

    def table(start_row, items, tier_label, extra_note=None):
        rr = start_row
        ws.cell(row=rr, column=1, value=f"Top stock-holding franchises in {tier_label}").font = TITLE_FONT
        rr += 1
        hdr = ["Base", "Gender", "Units in Stock", "FY25 Sales (SEK)", "Signal"]
        for c, h in enumerate(hdr, start=1):
            cell = ws.cell(row=rr, column=c, value=h)
            cell.font, cell.fill = HDR_FONT, HDR_FILL
        rr += 1
        for (b, g), units, sales in items:
            dead = sales < DEAD_STOCK_SALES_THRESHOLD
            ws.cell(row=rr, column=1, value=b).font = BODY_FONT
            ws.cell(row=rr, column=2, value=g).font = BODY_FONT
            c = ws.cell(row=rr, column=3, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
            c = ws.cell(row=rr, column=4, value=round(sales)); c.number_format, c.font = NUM_FMT, BODY_FONT
            signal = "Dead stock — near-zero FY25 sales" if dead else "Slow-moving, some sell-through"
            ws.cell(row=rr, column=5, value=signal).font = BODY_FONT
            if dead:
                for col in range(1, 6):
                    ws.cell(row=rr, column=col).fill = DEAD_FILL
            rr += 1
        if extra_note:
            rr += 1
            ws.cell(row=rr, column=1, value=extra_note).font = GRAY_FONT
            ws.merge_cells(f"A{rr}:E{rr}")
            ws.cell(row=rr, column=1).alignment = WRAP
            rr += 1
        return rr + 1

    r = table(r, top_items("Thin / Immaterial"), "Thin / Immaterial", (
        "Franchises with stock in this tier are the clearest \"clear via sale\" set: stock that isn't moving "
        "through normal channels at all -- watch for large unit counts against near-zero FY25 sales."
    ))
    r = table(r, top_items("Exited"), "Exited", (
        "Exited franchises still holding stock -- consistent with these being wound down; any large residual is "
        "worth a clearance push."
    ))
    r = table(r, top_items("Problem Child"), "Problem Child (contrast — NOT a clearance case)", (
        "Included for contrast, not action: real, actively-selling franchises with a margin problem, not dead "
        "stock. Clearing this the way Thin/Immaterial or Exited stock should be cleared would be the wrong move; "
        "per the tier definition, Problem Child needs a pricing/cost fix, not a fire sale."
    ))

    wb.save(lib.WORKBOOK_PATH)
    print("saved, last row:", r)


if __name__ == "__main__":
    main()
