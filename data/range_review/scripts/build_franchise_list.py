"""
Builds the FW27/SS28 franchise-level review list for the Nov 12 Range
Review workshop prep: one row per franchise (RR-001 definition -- base
name, gender NOT a separate split), with layer, business area, season
coverage, gender coverage, and Target RRP.

Margin target is NOT present anywhere in the FW27/SS28 source files (see
RR-008) -- left blank here rather than guessed; Tobias/Finance need to say
where that number lives before it can be added.

Scope: Active styles only (RR-002), excluding ABS/ABS Exclusive (RR-006).
Representative name = latest active generation, gender stripped but
version kept (RR-003). "Blueberry "-prefixed and " AF"-suffixed styles are
folded into their base franchise rather than listed as separate franchises
(RR-005/RR-007) -- flagged via the "Special edition / fit note" column
instead.

Adds a plain "Core Assortment FW27" Yes/No column, sourced from the flag
already in the SS26 tiering workbook (Sheet 2), rolled up from Base+Gender
to Base grain (Yes if either gender was Yes). Not framed as a historic-
sales cross-check here -- that comparison is a later step, once updated
sales data is in. "Core Assortment SS28" has no source yet (the flag was
only ever defined for FW27) -- left as "To confirm" for every row.

Run: python3 build_franchise_list.py
Output: ../outputs/FW27_SS28_Franchise_Review_List.xlsx
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))
from config import GENDER_TOKENS  # noqa: E402

RR_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402

INPUTS_DIR = RR_DIR / "inputs"
OUTPUTS_DIR = RR_DIR / "outputs"
FW27_PATH = INPUTS_DIR / "FW27_28_Oct_2026.xlsx"
SS28_PATH = INPUTS_DIR / "SS28_28_Oct_2026.xlsx"

# Gender-only strip (keeps version token) -- reuses GENDER_TOKENS from
# config.py rather than redefining it, per CLAUDE.md's shared-token-set rule.
_GENDER_ONLY_PATTERN = r"\b(" + "|".join(re.escape(t) for t in GENDER_TOKENS) + r")\b"


def strip_gender_only(style: str) -> str:
    s = re.sub(_GENDER_ONLY_PATTERN, "", str(style).strip())
    s = re.sub(r"\s+", " ", s).strip()
    for suf in ("Women", "Men"):
        if s.endswith(suf) and len(s) > len(suf):
            s = s[: -len(suf)].strip()
            break
    return s


def strip_special_edition(style: str) -> tuple[str, str | None]:
    """Strip 'Blueberry '/' AF' before franchise matching (RR-005/RR-007)."""
    s = str(style).strip()
    if s.startswith("Blueberry "):
        return s[len("Blueberry "):].strip(), "Special edition: Blueberry"
    if s.endswith(" AF"):
        return s[: -len(" AF")].strip(), "Asian Fit (AF)"
    return s, None


TIERING_SHEET = "2. Full Portfolio (1860)"


def load_core_assortment_fw27() -> dict[str, str]:
    """Base -> 'Yes'/'No', from the SS26 tiering workbook's Core Assortment FW27 flag.

    Base+Gender grain in the source, rolled up to Base (Yes if either
    gender was Yes). A base not found here isn't marked at all -- treated
    the same as 'No' rather than a separate bucket.
    """
    wb = openpyxl.load_workbook(lib.WORKBOOK_PATH, read_only=True, data_only=True)
    ws = wb[TIERING_SHEET]
    by_base: dict[str, set[str]] = defaultdict(set)
    for r in ws.iter_rows(min_row=lib.FIRST_DATA_ROW, values_only=True):
        base = r[0]
        if base is None:
            continue
        by_base[base].add(r[16])  # "Core Assortment FW27" column
    return {base: ("Yes" if "Yes" in flags else "No") for base, flags in by_base.items()}


def load_style_rows(path: Path, season_label: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[0] is None:
            continue  # colorway-only sub-row
        rows.append({
            "season_label": season_label, "active": r[1], "code": r[2], "style": r[3],
            "target_rrp": r[10], "layer": r[14], "gender": r[15],
            "business_area": r[16], "collection_type": r[23], "abs_customer": r[24],
        })
    return rows


def build_rows() -> list[dict]:
    core_fw27 = load_core_assortment_fw27()
    rows = load_style_rows(FW27_PATH, "FW27") + load_style_rows(SS28_PATH, "SS28")
    main_range = [r for r in rows
                  if r["active"] and r["collection_type"] not in ("ABS", "ABS Exclusive")]

    franchises: dict[str, list[dict]] = defaultdict(list)
    for r in main_range:
        stripped_style, special_flag = strip_special_edition(r["style"])
        base = lib.base_name(stripped_style)
        r["special_flag"] = special_flag
        r["display_name"] = strip_gender_only(stripped_style)
        r["version_rank"] = lib.RANK.get(lib.version_token(stripped_style), 0)
        franchises[base].append(r)

    out_rows = []
    for base, items in franchises.items():
        best = max(items, key=lambda r: (r["version_rank"], r["season_label"] == "SS28"))
        layers = sorted({r["layer"] for r in items if r["layer"]})
        areas = sorted({r["business_area"] for r in items if r["business_area"]})
        genders = sorted({r["gender"] for r in items if r["gender"]})
        seasons = sorted({r["season_label"] for r in items})
        rrps = sorted({r["target_rrp"] for r in items if r["target_rrp"] is not None})
        special_flags = sorted({r["special_flag"] for r in items if r["special_flag"]})

        gender_note = ""
        if len(genders) == 1 and genders[0] != "Unisex":
            gender_note = f"{genders[0]} only"

        rrp_display = rrps[0] if len(rrps) <= 1 else f"{min(rrps)}-{max(rrps)}"

        out_rows.append({
            "Franchise": best["display_name"],
            "Layer": "; ".join(layers),
            "Business Area": "; ".join(areas),
            "Gender": "; ".join(genders),
            "Gender note": gender_note,
            "Active in FW27": "Y" if "FW27" in seasons else "",
            "Active in SS28": "Y" if "SS28" in seasons else "",
            "Target RRP": rrp_display,
            "RRP varies by style": "Y" if len(rrps) > 1 else "",
            "Margin target": "",  # RR-008 -- not in source data, see register
            "Special edition / fit note": "; ".join(special_flags),
            "# styles rolled up": len(items),
            "Core Assortment FW27": core_fw27.get(base, "No"),
            "Core Assortment SS28": "To confirm",
        })

    out_rows.sort(key=lambda r: (r["Layer"], r["Business Area"], r["Franchise"]))
    return out_rows


def write_workbook(rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Franchise Review List"

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
    notes.append(["Scope & definitions -- see data/range_review/ASSUMPTIONS_REGISTER.md"])
    notes.append([])
    notes.append(["Franchise", "Base product name, gender NOT split out (RR-001). "
                                "A franchise can roll up several style codes "
                                "(different genders and/or generations)."])
    notes.append(["Style", "One row in the FW27/SS28 source file = one style "
                            "Code, with its own Gender, Layer, Target RRP, "
                            "Active flag. Multiple colorways nest under one style."])
    notes.append(["Scope of this list", "Active=True styles only (RR-002), "
                                         "excluding ABS/ABS Exclusive (RR-006, separate review). "
                                         "Blueberry special editions and AF/Asian Fit variants "
                                         "are folded into their base franchise, flagged in the "
                                         "note column, not listed separately (RR-005/RR-007)."])
    notes.append(["Margin target", "NOT available in the FW27/SS28 source files -- "
                                    "only Target RRP exists. Left blank. See RR-008."])
    notes.append(["Core Assortment FW27", "Yes/No, from the flag already in the SS26 tiering "
                                           "workbook (57/1,860 Base+Gender rows = 'Yes'), rolled "
                                           "up to Base grain (Yes if either gender was Yes)."])
    notes.append(["Core Assortment SS28", "No source yet -- the flag was only ever defined for "
                                           "FW27. Left as 'To confirm' for every row until "
                                           "Tobias/Isabelle define it for SS28."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 20
    notes.column_dimensions["B"].width = 90

    wb.save(path)


if __name__ == "__main__":
    rows = build_rows()
    out_path = OUTPUTS_DIR / "FW27_SS28_Franchise_Review_List.xlsx"
    write_workbook(rows, out_path)
    print(f"{len(rows)} franchises written to {out_path}")
