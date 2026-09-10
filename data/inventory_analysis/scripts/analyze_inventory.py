"""
Inventory Analysis workstream — builds the Season-Recency / aging view of
warehouse available stock that the SS26 Portfolio Tiering workstream's own
stock sheet (REG-019 there) doesn't cover.

This is a NEW, separate workstream (data/inventory_analysis/) from SS26
Portfolio Tiering, but it deliberately reuses that workstream's source file,
matching functions, and several of its confirmed definitions rather than
re-deriving them:
  - Source file: the exact same warehouse available-stock snapshot already
    checked in at data/ss26_portfolio_tiering/inputs/stock/
    Available_stock_260825.xlsx (byte-identical, confirmed via md5 — see
    REG-INV-001). Not duplicated here.
  - Franchise matching: ss26_lib.franchise_key() (Base+Gender, REG-010/014/
    015/019's method) — same import, not re-copied.
  - Tier join: ss26_lib.load_full_portfolio() against the published SS26
    tiering workbook, exactly as REG-019's own update_stock.py does it.
  - REG-004 (50,000 SEK/year reliability threshold), REG-008 (Core-account
    scope), REG-012 (Ordertype gap caveat), REG-019 (stock methodology:
    units only, no SEK value, no cost/price column in the source) — cited
    inline below and in ASSUMPTIONS_REGISTER.md rather than re-derived.

What's NEW here (not in REG-019's Stock-by-Tier sheet): a Season-Recency
lens using the source file's own `Season Article` field (e.g. "F26",
"S27"), which REG-019 never used. See REG-INV-004 for the bucket
definition and REG-INV-005 for why it's Provisional.

Usage:
    python3 data/inventory_analysis/scripts/analyze_inventory.py \
        [--source PATH]
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402

WORKSTREAM_DIR = REPO_ROOT / "data" / "inventory_analysis"
DEFAULT_SOURCE = REPO_ROOT / "data" / "ss26_portfolio_tiering" / "inputs" / "stock" / "Available_stock_260825.xlsx"
DATA_MIN_ROW = 3  # row 1 = header, row 2 = "Total" summary row (see REG-INV-002)

# --------------------------------------------------------------------------
# REG-INV-004: Season-Recency buckets. Provisional — see
# ASSUMPTIONS_REGISTER.md for the reasoning and the open questions.
# "Incoming/Future" = the two newest season codes present in this snapshot
# (pre-season stock already in the warehouse ahead of that season's main
# selling window). "Current season" = the season code matching this
# workstream's naming sibling (S26 = the SS26 season). Everything else
# (F25 and older) = "Aged".
# --------------------------------------------------------------------------
FUTURE_SEASONS = {"S27", "F26"}
CURRENT_SEASONS = {"S26"}


def season_bucket(season) -> str:
    if season in FUTURE_SEASONS:
        return "Incoming / Future (F26, S27)"
    if season in CURRENT_SEASONS:
        return "Current season (S26)"
    return "Aged (F25 and older)"


TIER_ORDER = lib.TIER_ORDER + ["UNMATCHED (not in tiering file)"]
DEAD_STOCK_SALES_THRESHOLD = 1000  # SEK — same signal threshold as REG-019's update_stock.py
TOP_N_DETAIL = 15


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=str(DEFAULT_SOURCE))
    args = ap.parse_args()

    wb_stock = openpyxl.load_workbook(args.source, data_only=True, read_only=True)
    candidates = [n for n in wb_stock.sheetnames if n.lower().startswith("available stock")]
    ws_stock = wb_stock[candidates[0]] if candidates else wb_stock[wb_stock.sheetnames[0]]
    rows = list(ws_stock.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header)}
    for col in ("Article", "Available stock", "Season Article", "Layer"):
        if col not in idx:
            sys.exit(f"Expected column '{col}' not found in source header: {header}")
    data = rows[DATA_MIN_ROW - 1:]

    # REG-019's tier join: same SS26 tiering workbook, same lookup shape.
    wb_tiering, ws2 = lib.load_full_portfolio()
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

    total_units = 0.0
    matched_units = 0.0
    padded_article_rows = 0
    season_tier_units: dict[str, Counter] = defaultdict(Counter)
    layer_units = Counter()
    unmatched_franchise_units = Counter()
    aged_flagged = defaultdict(float)  # (key, tier) -> units, Aged bucket only, flagged tiers only

    for row in data:
        article = row[idx["Article"]]
        stock = row[idx["Available stock"]] or 0
        season = row[idx["Season Article"]]
        layer = row[idx["Layer"]]
        if not article:
            continue
        total_units += stock
        if str(article) != str(article).strip():
            padded_article_rows += 1
        layer_units[layer] += stock

        key = lib.franchise_key(str(article).strip())
        bucket = season_bucket(season)
        if key in pub_lookup:
            tier = pub_lookup[key]["tier"]
            matched_units += stock
        else:
            tier = "UNMATCHED (not in tiering file)"
            unmatched_franchise_units[key] += stock
        season_tier_units[bucket][tier] += stock

        if bucket == "Aged (F25 and older)" and tier in ("Thin / Immaterial", "Exited", "Problem Child"):
            aged_flagged[(key, tier)] += stock

    pct_matched = matched_units / total_units * 100 if total_units else 0
    print(f"total units: {total_units:,.0f}  matched an existing franchise: {matched_units:,.0f} ({pct_matched:.1f}%)")
    print(f"rows with padded Article text: {padded_article_rows}")
    for bucket in ["Incoming / Future (F26, S27)", "Current season (S26)", "Aged (F25 and older)"]:
        print(f"-- {bucket}: {sum(season_tier_units[bucket].values()):,.0f} units")

    # ---------------------------------------------------------------------
    # Build the workbook
    # ---------------------------------------------------------------------
    s = lib.styles()
    HDR_FONT, HDR_FILL, BODY_FONT, GRAY_FONT = s["header_font"], s["header_fill"], s["body_font"], s["gray_font"]
    TITLE_FONT = Font(name="Arial", bold=True, color="FF1F3864")
    WRAP = Alignment(vertical="top", wrap_text=True)
    NUM_FMT, PCT_FMT = s["num_fmt"], s["pct_fmt"]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- Sheet 1: Overview ---
    ws1 = wb.create_sheet("1. Overview")
    ws1["A1"] = "Inventory Analysis — Season-Recency View of Available Stock"
    ws1["A1"].font = TITLE_FONT
    ws1.merge_cells("A1:F1")
    src_name = Path(args.source).name
    overview_text = (
        f"Source: {src_name} (warehouse available-stock snapshot; units by Article No./SKU), "
        "the same file already used for the SS26 Portfolio Tiering workstream's Stock-by-Tier "
        "sheet (REG-019 there) — reused directly, not duplicated (REG-INV-001). Matched to "
        "franchise (Base+Gender) via ss26_lib.franchise_key(), the same parsing used across the "
        "SS26 workstream (REG-010/014/015/019), then joined to each franchise's current "
        "consolidated Tier from the published tiering workbook. "
        f"{total_units:,.0f} total units in the source file; {matched_units:,.0f} ({pct_matched:.1f}%) "
        "matched an existing published franchise (see Sheet 5 for the remainder)."
    )
    ws1["A3"] = overview_text
    ws1["A3"].font = BODY_FONT
    ws1["A3"].alignment = WRAP
    ws1.merge_cells("A3:F3")
    ws1.row_dimensions[3].height = 90

    caveats = (
        "CAVEATS (inherited, see ASSUMPTIONS_REGISTER.md for full citations):\n"
        "- REG-019: units only — the source file has no cost/price column, so no SEK stock value "
        "is computed here either.\n"
        "- REG-004: the 50,000 SEK/year reliability threshold applies to the FY25 Sales figures "
        "pulled in from the tiering workbook for the \"dead stock\" signal below — those figures "
        "are already blanked/computed per that rule upstream.\n"
        "- REG-008: this stock file has no account/channel field of its own, so the Core-account "
        "exclusion doesn't apply to these rows directly — but every FY25 Sales figure joined in "
        "from the tiering workbook is already Core-scoped.\n"
        "- REG-012: any such joined FY25 Sales figure inherits the unresolved \"3-Close out order\" "
        "reconciliation gap — read as directional, not an audited total.\n"
        "- REG-INV-004/005 (new to this workstream): the Season-Recency buckets below are a "
        "Provisional definition, not yet confirmed with the business — see the register."
    )
    ws1["A6"] = caveats
    ws1["A6"].font = GRAY_FONT
    ws1["A6"].alignment = WRAP
    ws1.merge_cells("A6:F6")
    ws1.row_dimensions[6].height = 140
    for col, w in zip("ABCDEF", (20, 20, 20, 20, 20, 20)):
        ws1.column_dimensions[col].width = w

    # --- Sheet 2: Stock by Season Recency ---
    ws2s = wb.create_sheet("2. Stock by Season Recency")
    ws2s["A1"] = "Available Stock by Season Recency x Tier"
    ws2s["A1"].font = TITLE_FONT
    ws2s.merge_cells("A1:H1")
    ws2s["A3"] = (
        "Season Recency read from the source file's own \"Season Article\" field (e.g. \"F26\" = "
        "Fall/Winter 2026, \"S27\" = Spring/Summer 2027) — not previously used anywhere in this "
        "repo. Bucket definition is Provisional (REG-INV-004) — not yet confirmed with the "
        "business."
    )
    ws2s["A3"].font = GRAY_FONT
    ws2s["A3"].alignment = WRAP
    ws2s.merge_cells("A3:H3")
    ws2s.row_dimensions[3].height = 45

    bucket_order = ["Incoming / Future (F26, S27)", "Current season (S26)", "Aged (F25 and older)"]
    hdr_row = 5
    headers = ["Season Recency Bucket"] + TIER_ORDER + ["Bucket Total"]
    for c, h in enumerate(headers, start=1):
        cell = ws2s.cell(row=hdr_row, column=c, value=h)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = WRAP
    r = hdr_row + 1
    for bucket in bucket_order:
        ws2s.cell(row=r, column=1, value=bucket).font = BODY_FONT
        row_total = 0
        for c, tier in enumerate(TIER_ORDER, start=2):
            v = round(season_tier_units[bucket].get(tier, 0))
            row_total += v
            cell = ws2s.cell(row=r, column=c, value=v)
            cell.number_format, cell.font = NUM_FMT, BODY_FONT
        cell = ws2s.cell(row=r, column=len(TIER_ORDER) + 2, value=row_total)
        cell.number_format, cell.font = NUM_FMT, Font(name="Arial", bold=True)
        r += 1
    # column totals
    ws2s.cell(row=r, column=1, value="Tier Total").font = Font(name="Arial", bold=True)
    grand_total = 0
    for c, tier in enumerate(TIER_ORDER, start=2):
        col_total = sum(season_tier_units[b].get(tier, 0) for b in bucket_order)
        grand_total += col_total
        cell = ws2s.cell(row=r, column=c, value=round(col_total))
        cell.number_format, cell.font = NUM_FMT, Font(name="Arial", bold=True)
    cell = ws2s.cell(row=r, column=len(TIER_ORDER) + 2, value=round(grand_total))
    cell.number_format, cell.font = NUM_FMT, Font(name="Arial", bold=True)

    ws2s.column_dimensions["A"].width = 28
    for i in range(2, len(TIER_ORDER) + 3):
        ws2s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 16

    # --- Sheet 3: Stock by Layer ---
    ws3 = wb.create_sheet("3. Stock by Layer")
    ws3["A1"] = "Available Stock by Layer"
    ws3["A1"].font = TITLE_FONT
    ws3.merge_cells("A1:C1")
    ws3["A3"] = (
        "Layer field is used as-is — already in the same canonical casing as config.ALL_LAYERS "
        "(no config.LAYER_NAME_FIXES entries triggered on this file — see REG-INV-006)."
    )
    ws3["A3"].font = GRAY_FONT
    ws3["A3"].alignment = WRAP
    ws3.merge_cells("A3:C3")
    hdr_row = 5
    for c, h in enumerate(["Layer", "Units in Stock", "% of Total"], start=1):
        cell = ws3.cell(row=hdr_row, column=c, value=h)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
    r = hdr_row + 1
    for layer, units in layer_units.most_common():
        ws3.cell(row=r, column=1, value=layer or "(blank)").font = BODY_FONT
        c = ws3.cell(row=r, column=2, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
        pct = units / total_units * 100 if total_units else 0
        c = ws3.cell(row=r, column=3, value=round(pct, 1)); c.number_format, c.font = PCT_FMT, BODY_FONT
        r += 1
    ws3.column_dimensions["A"].width = 24
    ws3.column_dimensions["B"].width = 18
    ws3.column_dimensions["C"].width = 14

    # --- Sheet 4: Aged Stock Detail (the new, actionable cut) ---
    ws4 = wb.create_sheet("4. Aged Stock Detail")
    ws4["A1"] = "Aged Stock (F25-and-older season code) in Clearance-Relevant Tiers"
    ws4["A1"].font = TITLE_FONT
    ws4.merge_cells("A1:E1")
    ws4["A3"] = (
        "Franchises whose stock carries an F25-or-older Season Article code AND sit in a tier "
        "REG-019 already flags for review: Thin/Immaterial and Exited (\"clear via sale\") are the "
        "clearest double-confirmed candidates (old season code + tier says clear). Problem Child "
        "is included per REG-019's own contrast framing — a margin problem, not automatically dead "
        "stock — but an old season code on a Problem Child line is worth a merch check for whether "
        "it's a superseded colorway rather than current-line stock (Open, not resolved here)."
    )
    ws4["A3"].font = GRAY_FONT
    ws4["A3"].alignment = WRAP
    ws4.merge_cells("A3:E3")
    ws4.row_dimensions[3].height = 75
    hdr_row = 6
    for c, h in enumerate(["Base", "Gender", "Tier", "Aged Units in Stock", "FY25 Sales (SEK)"], start=1):
        cell = ws4.cell(row=hdr_row, column=c, value=h)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
    r = hdr_row + 1
    for (key, tier), units in sorted(aged_flagged.items(), key=lambda x: -x[1])[:TOP_N_DETAIL]:
        base, g = key
        sales25 = pub_lookup[key]["sales25"]
        ws4.cell(row=r, column=1, value=base).font = BODY_FONT
        ws4.cell(row=r, column=2, value=g).font = BODY_FONT
        ws4.cell(row=r, column=3, value=tier).font = BODY_FONT
        c = ws4.cell(row=r, column=4, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
        c = ws4.cell(row=r, column=5, value=round(sales25)); c.number_format, c.font = NUM_FMT, BODY_FONT
        r += 1
    for col, w in zip("ABCDE", (32, 10, 18, 18, 18)):
        ws4.column_dimensions[col].width = w

    # --- Sheet 5: Unmatched / Unscoped stock ---
    ws5 = wb.create_sheet("5. Unmatched Stock")
    ws5["A1"] = "Stock Not Matched to a Published Franchise"
    ws5["A1"].font = TITLE_FONT
    ws5.merge_cells("A1:C1")
    unmatched_total = sum(unmatched_franchise_units.values())
    ws5["A3"] = (
        f"{unmatched_total:,.0f} units ({unmatched_total/total_units*100:.1f}% of total) across "
        f"{len(unmatched_franchise_units)} distinct franchises don't match anything in the 1,860-row "
        "tiering file — consistent with REG-019's own explanation (likely brand-new FW27 launches "
        "with no FY24/25 sales history to tier against yet). Not independently confirmed as ALL "
        "being new launches vs. a residual name-matching miss (Open)."
    )
    ws5["A3"].font = GRAY_FONT
    ws5["A3"].alignment = WRAP
    ws5.merge_cells("A3:C3")
    ws5.row_dimensions[3].height = 60
    hdr_row = 6
    for c, h in enumerate(["Base", "Gender", "Units in Stock"], start=1):
        cell = ws5.cell(row=hdr_row, column=c, value=h)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
    r = hdr_row + 1
    for (base, g), units in unmatched_franchise_units.most_common(TOP_N_DETAIL):
        ws5.cell(row=r, column=1, value=base).font = BODY_FONT
        ws5.cell(row=r, column=2, value=g).font = BODY_FONT
        c = ws5.cell(row=r, column=3, value=round(units)); c.number_format, c.font = NUM_FMT, BODY_FONT
        r += 1
    for col, w in zip("ABC", (32, 10, 16)):
        ws5.column_dimensions[col].width = w

    out_name = f"Project_Bob_Inventory_Analysis_{date.today().strftime('%d%m%Y')}.xlsx"
    out_path = WORKSTREAM_DIR / out_name
    wb.save(out_path)
    print("saved:", out_path)


if __name__ == "__main__":
    main()
