"""
Builds the "separate ABS review" RR-006 flagged as follow-up work: every
style tagged Collection Type = ABS or ABS Exclusive across the FW27/SS28/
FW28 range files, with its ABS Customer (the account the style is built
for), grouped by franchise (RR-001: base name, gender not a separate
split) and by account.

Reuses load_style_rows()/load_style_rows_flat() from build_franchise_list.py
(same FW27/SS28 nested vs. FW28 flat layout handling, RR-026) rather than
re-parsing the range files -- these are explicitly OUT of the main Nov 12
range review (RR-006/RR-007), so this is a standalone side-review, not a
column added to the main franchise list.

ABS Customer data hygiene note: two case variants of the same account
appear in the source data -- "ASIA" (176 of 191 Asia-account rows) and
"Asia" (15 of 191) -- not distinguished by this script (folded together
as "Asia" for grouping) but worth a question to Tobias/Isabelle on whether
they mean the same thing.

Run: python3 build_abs_review.py
Output: ../outputs/ABS_Account_Review.xlsx
"""
from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RR_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RR_DIR / "scripts"))
from build_franchise_list import (  # noqa: E402
    FW27_PATH,
    FW28_PATH,
    SS28_PATH,
    load_style_rows,
    load_style_rows_flat,
    strip_special_edition,
)

REPO_ROOT = RR_DIR.parents[1]
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402

OUTPUTS_DIR = RR_DIR / "outputs"

ABS_COLLECTION_TYPES = ("ABS", "ABS Exclusive")


def normalize_account(raw: str | None) -> str:
    if not raw:
        return "(blank)"
    s = str(raw).strip()
    if s.upper() == "ASIA":
        return "Asia"
    return s


def load_abs_rows() -> list[dict]:
    rows = (
        load_style_rows(FW27_PATH, "FW27")
        + load_style_rows(SS28_PATH, "SS28")
        + load_style_rows_flat(FW28_PATH, "FW28")
    )
    abs_rows = [r for r in rows if r["collection_type"] in ABS_COLLECTION_TYPES]
    for r in abs_rows:
        stripped, special = strip_special_edition(r["style"])
        r["base"] = lib.base_name(stripped)
        r["special_edition"] = special
        r["account"] = normalize_account(r["abs_customer"])
    return abs_rows


def group_by_franchise(abs_rows: list[dict]) -> list[dict]:
    """One row per (base, account) -- the same franchise can have ABS
    variants for more than one account (rare but real, e.g. a style with
    both an XXL cut and an ASIA Fit cut)."""
    groups: dict[tuple[str, str], dict] = {}
    for r in abs_rows:
        key = (r["base"], r["account"])
        g = groups.setdefault(
            key,
            {
                "base": r["base"],
                "account": r["account"],
                "layer": r["layer"],
                "collection_types": set(),
                "seasons_present": set(),
                "seasons_active": set(),
                "codes": set(),
                "styles": set(),
            },
        )
        g["collection_types"].add(r["collection_type"])
        g["seasons_present"].add(r["season_label"])
        if r["active"]:
            g["seasons_active"].add(r["season_label"])
        g["codes"].add(str(r["code"]))
        g["styles"].add(r["style"])
    season_order = {"FW27": 0, "SS28": 1, "FW28": 2}

    def season_str(s: set[str]) -> str:
        return "/".join(sorted(s, key=lambda x: season_order.get(x, 99)))

    out = []
    for (base, account), g in groups.items():
        out.append(
            {
                "franchise": base,
                "account": account,
                "layer": g["layer"],
                "collection_type": "/".join(sorted(g["collection_types"])),
                "seasons_present": season_str(g["seasons_present"]),
                "seasons_active": season_str(g["seasons_active"]) or "(none active)",
                "codes": ", ".join(sorted(g["codes"])),
                "styles": "; ".join(sorted(g["styles"])),
            }
        )
    out.sort(key=lambda r: (r["account"], r["layer"] or "", r["franchise"]))
    return out


HEADER_FILL = PatternFill("solid", fgColor="1B211C")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
TITLE_FONT = Font(bold=True, size=13)


def autosize(ws, rows: list[dict], cols: list[str]) -> None:
    for i, col in enumerate(cols, start=1):
        width = max(len(col), *(len(str(r.get(col, ""))) for r in rows)) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 60)


def write_franchise_sheet(ws, rows: list[dict]) -> None:
    ws.title = "ABS by franchise"
    cols = ["franchise", "account", "layer", "collection_type", "seasons_present", "seasons_active", "codes", "styles"]
    headers = ["Franchise", "ABS Account", "Layer", "Collection Type", "Seasons present", "Seasons active", "Style code(s)", "Style name(s)"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws.freeze_panes = "A2"
    for r in rows:
        ws.append([r[c] for c in cols])
    autosize(ws, rows, cols)


def write_account_summary_sheet(ws, rows: list[dict]) -> None:
    ws.title = "Summary by account"
    by_account: dict[str, dict] = defaultdict(lambda: {"franchises": 0, "layers": defaultdict(int), "active": 0})
    for r in rows:
        acc = by_account[r["account"]]
        acc["franchises"] += 1
        acc["layers"][r["layer"] or "(unknown)"] += 1
        if r["seasons_active"] != "(none active)":
            acc["active"] += 1

    ws.append(["ABS Account", "Franchises", "Active in ≥ 1 season", "By layer"])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws.freeze_panes = "A2"
    for acc, d in sorted(by_account.items(), key=lambda x: -x[1]["franchises"]):
        layer_str = ", ".join(f"{k} ({v})" for k, v in sorted(d["layers"].items(), key=lambda x: -x[1]))
        ws.append([acc, d["franchises"], d["active"], layer_str])
    for i, w in enumerate([22, 14, 20, 70], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_notes_sheet(ws, raw_row_count: int, franchise_row_count: int) -> None:
    ws.title = "Notes"
    ws.column_dimensions["A"].width = 100
    lines = [
        ("ABS / ABS Account Review", TITLE_FONT),
        ("", None),
        ("Scope: every style tagged Collection Type = ABS or ABS Exclusive in the", None),
        ("FW27/SS28/FW28 range files, grouped by franchise and by account (the", None),
        ("ABS Customer field). See RR-006/RR-007 in the Assumptions Register --", None),
        ("these styles are deliberately OUT of the main Nov 12 Range Review (and", None),
        ("out of the Range Review Board / Price Maps), reviewed here as its own,", None),
        ("separate topic per the user's original framing.", None),
        ("", None),
        (f"{raw_row_count} ABS/ABS Exclusive style rows across FW27+SS28+FW28,", None),
        (f"consolidating to {franchise_row_count} franchise-account rows (one franchise can", None),
        ("have ABS variants for more than one account).", None),
        ("", None),
        ("Data hygiene note: the source data has two case variants of the same", None),
        ("account -- \"ASIA\" and \"Asia\" -- folded together here as \"Asia\"; worth", None),
        ("confirming with Tobias/Isabelle whether that's correct.", None),
        ("", None),
        ("\"Seasons active\" uses each row's own Active flag per season -- a style", None),
        ("present in a season file but marked inactive there shows in \"Seasons", None),
        ("present\" but not \"Seasons active\". \"(none active)\" means the ABS variant", None),
        ("exists in the data but isn't currently active in any of FW27/SS28/FW28.", None),
    ]
    for text, font in lines:
        ws.append([text])
        if font:
            ws.cell(row=ws.max_row, column=1).font = font
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True)


def main() -> None:
    abs_rows = load_abs_rows()
    franchise_rows = group_by_franchise(abs_rows)

    wb = openpyxl.Workbook()
    write_franchise_sheet(wb.active, franchise_rows)
    write_account_summary_sheet(wb.create_sheet(), franchise_rows)
    write_notes_sheet(wb.create_sheet(), len(abs_rows), len(franchise_rows))

    out_path = OUTPUTS_DIR / "ABS_Account_Review.xlsx"
    wb.save(out_path)
    print(f"{len(abs_rows)} ABS/ABS Exclusive style rows -> {len(franchise_rows)} franchise-account rows")
    print(f"Wrote {out_path.relative_to(REPO_ROOT)}")


if __name__ == "__main__":
    main()
