"""
REG-010: Clearance Flag / Clearance Detail, and the Clearance tier reassignment.

Reads the two Aug-2026 close-out lists (China JV + Zalando) from
inputs/clearance/, matches them to franchises via SCO -> (Base, Gender),
and classifies each franchise as:
  - "Fully"   -- every SCO ever seen for this franchise across the three
                 SS26 raw exports is on a clearance list. Tier AND
                 Original Tier are reassigned to CLEARANCE_TIER (the
                 whole franchise is being wound down/shipped out).
  - "Partial" -- some but not all of the franchise's SCOs are on a
                 clearance list (e.g. an old generation SKU, while a
                 newer generation of the same franchise continues).
                 Tier is left untouched -- see REG-010's methodology
                 note in Sheet 1 for why (a franchise can be both "Hero"
                 and have clearance exposure at the same time).
  - "None"    -- no overlap.

Writes columns N (Clearance Flag) and O (Clearance Detail) on the
Full Portfolio sheet, and updates column D (Tier) / column P (Original
Tier) in place for "Fully" franchises. Does NOT re-sort the sheet --
sorting by tier was a one-time formatting step at initial construction;
re-running this script preserves whatever row order the sheet is
currently in, since other update_*.py scripts key off (Base, Gender)
rather than row position.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_clearance.py \
        [--jv PATH] [--zalando PATH]

Defaults to inputs/clearance/Close_out_delivery_JV_August_2026.xlsx and
inputs/clearance/Zalando_close_out_order_August_2026.xlsx -- point at a
new pair of files for a future close-out round.
"""
from __future__ import annotations

import argparse
from collections import defaultdict

import openpyxl

import ss26_lib as lib


def load_clearance(path, sheet=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    return rows[1:]  # drop header


def sco_sources(rows, source):
    """First 9 characters of Article No. = SCO (Model+Color ID); Article No. itself is 12 chars (SCO+size)."""
    s = defaultdict(set)
    for row in rows:
        art_no = row[0]
        if art_no is None or art_no == "Total":
            continue
        sco = str(art_no)[:9]
        s[sco].add(source)
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--jv", default=str(lib.INPUTS_DIR / "clearance" / "Close_out_delivery_JV_August_2026.xlsx"))
    ap.add_argument("--zalando", default=str(lib.INPUTS_DIR / "clearance" / "Zalando_close_out_order_August_2026.xlsx"))
    ap.add_argument("--zalando-sheet", default="Zalando close out order August")
    args = ap.parse_args()

    # 1. every SCO ever seen for each franchise, across all three raw exports
    franchise_scos = defaultdict(set)
    sco_info = {}
    for path in lib.SS26_EXPORTS:
        header, rows = lib.load_raw_export(path)
        idx = {n: i for i, n in enumerate(header)}
        for row in rows:
            sco = row[idx["SCO"]]
            article = row[idx["Article"]]
            if sco is None or not article:
                continue
            a = article.strip()
            key = lib.franchise_key(a)
            franchise_scos[key].add(sco)
            if sco not in sco_info:
                sco_info[sco] = {
                    "article": a,
                    "model": row[idx["Model"]],
                    "start_season": row[idx["Start Season"]],
                }

    # 2. clearance SCOs from the two close-out lists
    jv_rows = load_clearance(args.jv)
    za_rows = load_clearance(args.zalando, args.zalando_sheet)
    clearance_scos = defaultdict(set)
    for rows, source in [(jv_rows, "JV/China"), (za_rows, "Zalando")]:
        for sco, srcs in sco_sources(rows, source).items():
            clearance_scos[sco] |= srcs

    unmatched_no_raw = [s for s in clearance_scos if s not in sco_info]

    # 3. group clearance SCOs by franchise
    franchise_clearance_scos = defaultdict(set)
    franchise_sources = defaultdict(set)
    for sco, srcs in clearance_scos.items():
        if sco not in sco_info:
            continue
        key = lib.franchise_key(sco_info[sco]["article"])
        franchise_clearance_scos[key].add(sco)
        franchise_sources[key] |= srcs

    # 4. classify Fully vs Partial
    fully, partial = {}, {}
    for key, c_scos in franchise_clearance_scos.items():
        all_scos = franchise_scos.get(key, set())
        if all_scos and c_scos >= all_scos:
            fully[key] = c_scos
        else:
            partial[key] = c_scos

    def detail_string(key, c_scos):
        models = sorted(set(str(sco_info[s]["model"]) for s in c_scos))
        srcs = " + ".join(sorted(franchise_sources[key]))
        seasons = sorted(set(
            str(sco_info[s]["start_season"]) for s in c_scos
            if sco_info[s]["start_season"] and str(sco_info[s]["start_season"]) != "0"
        ))
        season_str = f", season {'/'.join(seasons)}" if seasons else ""
        return f"{len(c_scos)} SKU(s) (Model {', '.join(models)}) via {srcs}{season_str}"

    fully_detail = {k: detail_string(k, v) for k, v in fully.items()}
    partial_detail = {k: detail_string(k, v) for k, v in partial.items()}

    print(f"Total distinct clearance SCOs: {len(clearance_scos)}")
    print(f"Unmatched (no raw export history at all): {len(unmatched_no_raw)}")
    print(f"Franchises touched: {len(franchise_clearance_scos)}  Fully: {len(fully)}  Partial: {len(partial)}")

    # 5. write to workbook
    wb, ws = lib.load_full_portfolio()
    s = lib.styles()

    ws["N4"], ws["O4"] = "Clearance Flag", "Clearance Detail"
    for coord in ("N4", "O4"):
        ws[coord].font = s["header_font"]
        ws[coord].fill = s["header_fill"]

    n_fully = n_partial = 0
    for r in range(lib.FIRST_DATA_ROW, ws.max_row + 1):
        base = ws.cell(row=r, column=1).value
        if base is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        key = (base.strip(), g)

        if key in fully:
            ws.cell(row=r, column=4, value=lib.CLEARANCE_TIER).font = s["body_font"]   # Tier
            ws.cell(row=r, column=16, value=lib.CLEARANCE_TIER).font = s["body_font"]  # Original Tier
            ws.cell(row=r, column=5, value="Clearance").font = s["body_font"]          # Bucket
            ws.cell(row=r, column=14, value="Fully").font = s["body_font"]
            ws.cell(row=r, column=15, value=fully_detail[key]).font = s["body_font"]
            n_fully += 1
        elif key in partial:
            ws.cell(row=r, column=14, value="Partial").font = s["body_font"]
            ws.cell(row=r, column=15, value=partial_detail[key]).font = s["body_font"]
            n_partial += 1
        else:
            ws.cell(row=r, column=14, value="None").font = s["body_font"]
            ws.cell(row=r, column=15, value=None).font = s["body_font"]

    ws.column_dimensions["N"].width = 14
    ws.column_dimensions["O"].width = 55

    wb.save(lib.WORKBOOK_PATH)
    print(f"Saved. Rows marked Fully: {n_fully}  Partial: {n_partial}")


if __name__ == "__main__":
    main()
