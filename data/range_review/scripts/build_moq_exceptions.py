"""
Builds a structured view of the SS28 FC1 MOQ exceptions flagged by Buying
(Kristina Thoren/Fanny) and Merch (Isabelle Witschard), cross-referenced
against the franchise each style belongs to in the FW27/SS28 range files.

Source: inputs/correspondence/Re__SS28_FC1_stilar_som_inte_m_ter_MOQ.eml
(email thread, 28-Oct-2026) -- see RR-016. The 30-row table was parsed
from the email's own HTML table (html.parser, not hand-transcribed) and
cross-checked against its plain-text rendering; both matched exactly.
Hardcoded here rather than re-parsed from the .eml at run time, since an
email is correspondence, not a stable machine-readable export -- the
source is cited so this can be re-verified against the original.

These are FC1 (not FC3) demand numbers against each style's MOQ -- a
different forecast round than build_forecast_summary.py's FC3-based
Qty/GP%, not directly comparable to it.

Run: python3 build_moq_exceptions.py
Output: ../outputs/SS28_MOQ_Exceptions.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RR_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RR_DIR / "scripts"))
from build_franchise_list import FW27_PATH, SS28_PATH, load_style_rows, strip_special_edition  # noqa: E402

REPO_ROOT = RR_DIR.parents[1]
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402

OUTPUTS_DIR = RR_DIR / "outputs"

# Code, Style, Colorway Code, Colorway, SS28 Range, FW27 Range, MOQ Style,
# MOQ Color, SS28 Demand FC1 Qty, Buying Comment, Merch Comment
MOQ_TABLE = [
    ("608172", "Alert GTX Jacket Men", "3N5", "Tarn Blue", "n/a", "FW27", 1000, 300, 820, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608192", "Alert GTX Jacket Women", "6GB", "Walnut Brown", "n/a", "FW27", 1000, 300, 740, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608178", "Alert Mid Shorts Women", "6E1", "Haze Blue", "SS28", "n/a", 1000, 1, 960, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608816", "Kaise Wool Tank Top", "6H5", "Cedar Brown", "SS28", "n/a", 1000, 300, 900, "", "Increase to MOQ"),
    ("609018", "Klyva GTX II Jacket Men", "3N5", "Tarn Blue", "SS28", "n/a", 1000, 300, 790, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("609019", "Klyva GTX II Jacket Women", "5YH", "Misty Mauve", "SS28", "n/a", 1000, 300, 665, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("607779", "Korp Mid Halfzip Men", "2C5", "True Black", "SS28", "FW27", 1000, 300, 700, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("607782", "Korp Mid Halfzip Women", "2C5", "True Black", "SS28", "FW27", 1000, 300, 600, "MOQ is 1000 pcs", "340pcs requested from Markets -Drop?"),
    ("608179", "L.I.M Pace Shorts Women", "6HJ", "Birch Glow", "SS28", "n/a", 1000, 300, 940, "", "Increase to MOQ"),
    ("608694", "Lugna Softshell Overshirt Men", "5YJ", "Stone Brick", "SS28", "FW27", 1000, 300, 600, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608695", "Lugna Softshell Overshirt Women", "3X3", "Stone Grey", "SS28", "FW27", 1000, 300, 600, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("607707", "Magma Rugged Pant Women", "4Q9", "Seaweed Green", "SS28", "FW27", 1000, 300, 750, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608680", "Mimic Alert 2.0 Hood Women", "2C5", "True Black", "SS28", "FW27", 1000, 500, 820, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608679", "Mimic Alert 2.0 Jacket Men", "2C5", "True Black", "SS28", "FW27", 1000, 500, 800, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608681", "Mimic Alert 2.0 Jacket Women", "6E1", "Haze Blue", "SS28", "FW27", 1000, 500, 500, "LOW FC should style be kept MOQ is 1000 pcs", "340pcs requested from Markets -Drop?"),
    ("608345", "Nallo Touring Hybrid Mid Men", "2C5", "True Black", "SS28", "FW27", 1000, 300, 600, "MOQ is 1000 pcs", "467 pcs requested from Markets -Only keep FW season? Or buy with surcharge SS28?"),
    ("608346", "Nallo Touring Hybrid Mid Women", "2C5", "True Black", "SS28", "FW27", 1000, 300, 600, "MOQ is 1000 pcs", "390 pcs requested from Markets -Only keep FW season? Or buy with surcharge SS28?"),
    ("608173", "Ozka Proof Jacket Men", "6GB", "Walnut Brown", "SS28", "FW27", 1000, 300, 860, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608193", "Ozka Proof Jacket Women", "3N5", "Tarn Blue", "SS28", "FW27", 1000, 300, 700, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608208", "Ozka Proof Pant Men", "2C5", "True Black", "SS28", "FW27", 1000, 300, 400, "LOW FC should style be kept MOQ is 1000 pcs", "382 pcs requested from Markets -Only keep FW season? Or buy with surcharge SS28?"),
    ("608209", "Ozka Proof Pant Women", "2C5", "True Black", "SS28", "FW27", 1000, 300, 400, "LOW FC should style be kept MOQ is 1000 pcs", "368 pcs requested from Markets -Only keep FW season? Or buy with surcharge SS28?"),
    ("608198", "Salo Shirt II Women", "5S9", "Chalk Beige", "SS28", "n/a", 1000, 300, 690, "LOW FC should style be kept MOQ is 1000 pcs", "Increase to MOQ"),
    ("608696", "Selme LS Crewneck Men", "3N5", "Tarn Blue", "SS28", "FW27", 1000, 300, 900, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608697", "Selme LS Crewneck Women", "3N5", "Tarn Blue", "SS28", "FW27", 1000, 300, 500, "LOW FC should style be kept MOQ is 1000 pcs", "Increase to MOQ or buy with surcharge?"),
    ("608156", "Solis Zip Off Pant Men", "2C5", "True Black", "SS28", "FW27", 1000, 300, 900, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608686", "Swook Mid Halfzip W", "6EF", "Haze Blue/Lt Haze Blue", "SS28", "FW27", 1000, 300, 500, "LOW FC should style be kept MOQ is 1000 pcs", "340pcs requested from Markets -Drop?"),
    ("608687", "Swook Mid Sweater M", "6EF", "Haze Blue/Lt Haze Blue", "SS28", "FW27", 1000, 300, 500, "LOW FC should style be kept MOQ is 1000 pcs", "340pcs requested from Markets -Drop?"),
    ("608367", "Särna Mimic II Hood Women", "2C5", "True Black", "SS28", "FW27", 1000, 300, 900, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608477", "Vara Alpha Sweater Men", "2C5", "True Black", "SS28", "FW27", 1000, 300, 900, "MOQ is 1000 pcs", "Increase to MOQ"),
    ("608478", "Vara Alpha Sweater Women", "2C5", "True Black", "SS28", "FW27", 1000, 300, 800, "MOQ is 1000 pcs", "Increase to MOQ"),
]


def load_code_lookup() -> dict[str, dict]:
    """Code -> franchise Base/Layer/Business Area, from ALL style rows (not
    just the main-range filter -- every code in MOQ_TABLE is a regular
    Inline SS28 style, so no ABS/Active filtering is needed here)."""
    rows = load_style_rows(FW27_PATH, "FW27") + load_style_rows(SS28_PATH, "SS28")
    lookup = {}
    for r in rows:
        if r["code"] in lookup:
            continue
        stripped, _ = strip_special_edition(r["style"])
        lookup[str(r["code"])] = {
            "base": lib.base_name(stripped),
            "layer": r["layer"],
            "business_area": r["business_area"],
        }
    return lookup


def build_rows() -> list[dict]:
    code_lookup = load_code_lookup()
    out_rows = []
    for (code, style, cw_code, cw_name, ss28_range, fw27_range, moq_style,
         moq_color, demand_qty, buying_comment, merch_comment) in MOQ_TABLE:
        info = code_lookup.get(code, {})
        out_rows.append({
            "Code": code,
            "Style": style,
            "Franchise": info.get("base", ""),
            "Layer": info.get("layer", ""),
            "Business Area": info.get("business_area", ""),
            "Colorway Code": cw_code,
            "Colorway": cw_name,
            "SS28 Range": ss28_range,
            "FW27 Range": fw27_range,
            "MOQ (Style)": moq_style,
            "MOQ (Color)": moq_color,
            "SS28 Demand FC1 (Qty)": demand_qty,
            "Shortfall vs MOQ": moq_style - demand_qty,
            "Buying Comment": buying_comment,
            "Merch Comment (Isabelle)": merch_comment,
        })
    out_rows.sort(key=lambda r: (r["Layer"] or "", r["Franchise"]))
    return out_rows


def write_workbook(rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SS28 MOQ Exceptions"

    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for r in rows:
        ws.append(list(r.values()))

    for i, h in enumerate(headers, start=1):
        width = max(len(h), max((len(str(r[h])) for r in rows), default=0)) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 50)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    notes = wb.create_sheet("Notes")
    notes.append(["Source", "inputs/correspondence/Re__SS28_FC1_stilar_som_inte_m_ter_MOQ.eml "
                             "(email thread, 28-Oct-2026) -- Kristina Thoren/Fanny (Buying) ran "
                             "SS28 FC1 demand against each style's MOQ; Isabelle Witschard (Merch) "
                             "added per-style suggestions. See RR-016."])
    notes.append(["Status", "These are FACTS (which styles miss MOQ, by how much) plus Isabelle's "
                             "SUGGESTED actions -- not yet decided. Linden Mallory's reply says "
                             "specific feedback comes after the Nov 12 workshop, so 'Drop?' / "
                             "'Only keep FW season? Or buy with surcharge?' rows are open questions "
                             "for that discussion, not settled calls."])
    notes.append(["Forecast round", "FC1, not FC3 -- a different, earlier forecast round than "
                                     "build_forecast_summary.py's Qty/GP% (which uses FC3). Not "
                                     "directly comparable without accounting for that."])
    notes.append(["Franchise / Layer", "Joined from the FW27/SS28 range files by style Code, same "
                                        "matching as build_franchise_list.py."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 18
    notes.column_dimensions["B"].width = 100

    wb.save(path)


if __name__ == "__main__":
    rows = build_rows()
    out_path = OUTPUTS_DIR / "SS28_MOQ_Exceptions.xlsx"
    write_workbook(rows, out_path)
    print(f"{len(rows)} MOQ-exception styles written to {out_path}")
