"""
REG-015: "Core Assortment FW27" column (Yes/No) on the Full Portfolio sheet.

Reads inputs/core_assortment/FW27_CORE_Assortment_Styles.xlsx (Sheet1,
column B = STYLE_NAME) and matches directly on the style-name text (via
the same base_name/gender parsing as everywhere else) rather than via
sales history -- an earlier attempt matched through raw-sales-history
lookup and produced false negatives for not-yet-shipped successor SKUs
with zero sales history (e.g. "Mimic Alert 2.0 Hood", "Rosson Mid II
Jacket Women"). Matching on the style name directly fixed that.

NEWS_PACKAGE_STYLES below is a small hardcoded fallback: 27 FW27 "News
package" items whose model-code -> style-name mapping was manually
transcribed from a slide deck (Core_assortment_FW27__send_out.pdf) before
the proper Excel export arrived. Kept here in case a future core-
assortment file is missing recently-launched items again; delete entries
once the source file reliably includes them.

Only styles that match an EXISTING published franchise (i.e. one with
FY24/25 or YTD26 sales history already in Sheet 2) get marked -- a
core-assortment style with no matching row is a genuinely new FW27
launch not yet in the tiering file at all, and is reported but not
written anywhere (there's no row to write it to).

This script prints the "Hero + Near-Hero gaps" list (Hero-tier franchises
NOT in Core Assortment FW27) -- if the numbers here differ materially
from Sheet 1's callout table, update that table by hand; this script only
touches Sheet 2.

Usage:
    python3 data/ss26_portfolio_tiering/scripts/update_core_assortment.py \
        [--source PATH]
"""
from __future__ import annotations

import argparse

import openpyxl

import ss26_lib as lib

# Manually transcribed from Core_assortment_FW27__send_out.pdf's "News
# package" slide (model code -> style name) -- see docstring above.
NEWS_PACKAGE_STYLES = {
    "608670": "Tarre Proof Insulated Jacket Men", "608674": "Tarre Proof Insulated Jacket Women",
    "608671": "Saulo Proof Jacket Men", "608676": "Saulo Proof Jacket Women",
    "608673": "Tarre Proof Insulated Pant Men", "608675": "Tarre Proof Insulated Pant Women",
    "608672": "Saulo Proof Pant Men", "608677": "Saulo Proof Pant Women",
    "608692": "Skalvik Softshell Parka Men", "608693": "Skalvik Softshell Parka Women",
    "608678": "Mimic Alert 2.0 Hood Men", "608680": "Mimic Alert 2.0 Hood Women",
    "608679": "Mimic Alert 2.0 Jacket Men", "608681": "Mimic Alert 2.0 Jacket Women",
    "608694": "Lugna Softshell Overshirt Men", "608695": "Lugna Softshell Overshirt Women",
    "608686": "Swook Mid Halfzip W", "608687": "Swook Mid Sweater M",
    "608683": "Brisen Tech Mid Hood Men", "608684": "Brisen Tech Mid Hood Women",
    "608696": "Sunna LS Crewneck Men", "608697": "Sunna LS Crewneck Women",
    "608691": "Tova Wool Beanie", "608690": "Aresk Beanie", "608720": "Rosson Beanie",
    "608702": "Haglöfs Vertigo QL PROOF Low Men", "608704": "Haglöfs Vertigo QL PROOF Low Women",
    "608703": "L.I.M Horizon GTX Mid Men", "608705": "L.I.M Horizon GTX Mid Women",
    "608700": "Sling pouch 2L", "608701": "Courier Messenger Bag",
}


def gender_with_abbrev_override(style_name: str) -> str:
    """Two News-package items use an abbreviated "W"/"M" suffix instead of "Women"/"Men"."""
    a = style_name.strip()
    if a == "Swook Mid Halfzip W":
        return "Women"
    if a == "Swook Mid Sweater M":
        return "Men"
    return lib.gender(a)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=str(lib.INPUTS_DIR / "core_assortment" / "FW27_CORE_Assortment_Styles.xlsx"))
    args = ap.parse_args()

    wb_src = openpyxl.load_workbook(args.source, data_only=True)
    ws_src = wb_src["Sheet1"]
    style_items = [(row[0], str(row[1]).strip()) for row in ws_src.iter_rows(min_row=2, values_only=True) if row[1]]
    all_items = list(style_items) + [(v, k) for k, v in NEWS_PACKAGE_STYLES.items()]
    print("total core-assortment style entries (file + News package fallback):", len(all_items))

    core_keys_direct = set()
    for style_name, _model in all_items:
        core_keys_direct.add((lib.base_name(style_name), gender_with_abbrev_override(style_name)))
    print("distinct franchise keys (direct style-name parse):", len(core_keys_direct))

    wb, ws = lib.load_full_portfolio()
    pub_keys = set()
    for r in range(lib.FIRST_DATA_ROW, ws.max_row + 1):
        b = ws.cell(row=r, column=1).value
        if b is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        pub_keys.add((b.strip(), g))

    core_keys = core_keys_direct & pub_keys
    unmatched = core_keys_direct - pub_keys
    print("matched to an existing published franchise:", len(core_keys))
    print("unmatched (genuinely new FW27 launch, no FY24/25/YTD26 history yet):", len(unmatched))
    for u in sorted(unmatched):
        print(" ", u)

    s = lib.styles()
    ws["Q4"] = "Core Assortment FW27"
    ws["Q4"].font = s["header_font"]
    ws["Q4"].fill = s["header_fill"]
    ws.column_dimensions["Q"].width = 18

    hero_gaps = []
    n_yes = n_no = 0
    for r in range(lib.FIRST_DATA_ROW, ws.max_row + 1):
        b = ws.cell(row=r, column=1).value
        if b is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        tier = ws.cell(row=r, column=4).value
        key = (b.strip(), g)
        in_core = key in core_keys
        val = "Yes" if in_core else "No"
        ws.cell(row=r, column=17, value=val).font = s["body_font"]
        if in_core:
            n_yes += 1
        else:
            n_no += 1
            if tier == "Hero + Near-Hero":
                hero_gaps.append((b, g, ws.cell(row=r, column=6).value or 0, ws.cell(row=r, column=7).value))

    print(f"Yes: {n_yes}  No: {n_no}")
    print(f"Hero + Near-Hero gaps: {len(hero_gaps)}")
    for h in sorted(hero_gaps, key=lambda x: -x[2]):
        gm = f"{h[3]:.1f}" if h[3] is not None else "n/a"
        print(f"  {h[0]:<28} {h[1]:<7} Sales_2025={h[2]:,.0f}  GM%_2025={gm}")

    wb.save(lib.WORKBOOK_PATH)
    print("saved")


if __name__ == "__main__":
    main()
