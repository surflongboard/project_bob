"""
Builds the FW27/SS28 franchise-level review list for the Nov 12 Range
Review workshop prep: one row per franchise (RR-001 definition -- base
name, gender NOT a separate split), with layer, business area, season
coverage, gender coverage, and Target RRP.

Margin target: the FW27/SS28 range files have no such field (RR-008), but
the FW27 and SS28 "Pic Centric" workbooks do (Target GM0%) -- pulled in
here via build_product_images.py's combined outputs/gm_by_code.json
(both seasons, SS28 winning on the rare conflict -- see that script's
docstring). Looked up by the franchise's own representative ("best")
code, so it works regardless of which season that code is from. Still
blank for the small number of franchises with no Pic Centric match at
all in either season.

No embedded product images in this list (tried, didn't look good -- see
RR-015). The extracted photos stay in outputs/product_images/ as a
reusable repository, keyed by style Code, for pulling into a better-suited
format later.

"Consumer Activity" is the raw multi-select field from the range files
(e.g. "Trekking / Hiking, Everyday Outdoor"), shown as-is per franchise --
NOT collapsed into a single category. It's a reference signal only; the
"strategic intent" category framing it might relate to (RR-017) is
explicitly not adopted as this list's structure.

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

Also adds "# active colors" (sum of colorway-level Active flags across
every style rolled into the franchise -- see RR-010 for a caveat on rows
with zero colorway sub-rows), and "ABS"/"Asia Fit" Yes/No flags so a
reviewer knows an ABS/ABS Exclusive or AF/Asian-Fit variant exists
somewhere for this franchise even though RR-006/RR-007 keep those rows
out of the main-range count.

Merges in the forecast data from build_forecast_summary.py (reusing its
loaders and `totals()` rather than redefining them): Qty/GP% blended and
split DTC vs Wholesale, summed across every style (both genders, both
seasons) rolled into the franchise, plus Status shown separately per
season (Status FW27 / Status SS28) since a franchise can be e.g. New in
FW27 but Carry Over in SS28. Joined on (season, style Code) -- 591/616
main-range style rows have a forecast match; 25 don't (RR-014) and show
blank forecast columns rather than 0, which would misrepresent "no data"
as "zero forecast."

Run: python3 build_franchise_list.py
Output: ../outputs/FW27_SS28_Franchise_Review_List.xlsx
"""
from __future__ import annotations

import json
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

sys.path.insert(0, str(RR_DIR / "scripts"))
from build_forecast_summary import (  # noqa: E402
    FW27_PATH as FORECAST_FW27_PATH, SS28_PATH as FORECAST_SS28_PATH,
    STATUS_LABELS, load_fw27 as load_forecast_fw27, load_ss28 as load_forecast_ss28, totals,
)

INPUTS_DIR = RR_DIR / "inputs"
OUTPUTS_DIR = RR_DIR / "outputs"
FW27_PATH = INPUTS_DIR / "FW27_28_Oct_2026.xlsx"
SS28_PATH = INPUTS_DIR / "SS28_28_Oct_2026.xlsx"
GM_BY_CODE_PATH = OUTPUTS_DIR / "gm_by_code.json"


def load_gm_by_code() -> dict[str, dict]:
    with open(GM_BY_CODE_PATH) as f:
        return json.load(f)

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
    """One dict per style row, with its nested colorway sub-rows counted.

    A style row has Season populated; its colorways are the following rows
    with Season=None, each with its own Active flag (col 4) -- a colorway
    carried in the sheet for reference but not actually active this season
    is NOT counted, since "how many colors" means colors on sale, not
    every colorway ever listed.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    current = None
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[0] is not None:  # a style row
            current = {
                "season_label": season_label, "active": r[1], "code": r[2], "style": r[3],
                "target_rrp": r[10], "layer": r[14], "gender": r[15],
                "business_area": r[16], "collection_type": r[23], "abs_customer": r[24],
                "consumer_activity": r[37], "active_colors": 0,
            }
            rows.append(current)
        elif current is not None:  # a colorway sub-row of the preceding style
            if r[4]:  # colorway-level Active flag
                current["active_colors"] += 1
    return rows


def compute_abs_af_flags(all_rows: list[dict]) -> tuple[set[str], set[str]]:
    """Base -> whether an ABS/ABS Exclusive variant or an AF/Asian-Fit variant
    exists anywhere for it, scanning ALL style rows (not just the main-range
    filter) -- RR-006/RR-007 keep these out of the main review's row count,
    but a reviewer looking at a franchise still needs to know one exists."""
    abs_bases, af_bases = set(), set()
    for r in all_rows:
        stripped, special_flag = strip_special_edition(r["style"])
        base = lib.base_name(stripped)
        if r["collection_type"] in ("ABS", "ABS Exclusive"):
            abs_bases.add(base)
        if special_flag == "Asian Fit (AF)":
            af_bases.add(base)
    return abs_bases, af_bases


def load_forecast_by_key() -> dict[tuple[str, str], list[dict]]:
    """(season, Code) -> its forecast rows (one per Market/Channel/Region)."""
    forecast_rows = load_forecast_fw27(FORECAST_FW27_PATH) + load_forecast_ss28(FORECAST_SS28_PATH)
    by_key: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for r in forecast_rows:
        by_key[(r["season"], r["code"])].append(r)
    return by_key


def build_rows() -> list[dict]:
    core_fw27 = load_core_assortment_fw27()
    forecast_by_key = load_forecast_by_key()
    gm_by_code = load_gm_by_code()
    rows = load_style_rows(FW27_PATH, "FW27") + load_style_rows(SS28_PATH, "SS28")
    abs_bases, af_bases = compute_abs_af_flags(rows)
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
        activities = sorted({
            tok.strip() for r in items if r["consumer_activity"]
            for tok in r["consumer_activity"].split(",")
        })

        gender_note = ""
        if len(genders) == 1 and genders[0] != "Unisex":
            gender_note = f"{genders[0]} only"

        rrp_display = rrps[0] if len(rrps) <= 1 else f"{min(rrps)}-{max(rrps)}"
        total_colors = sum(r["active_colors"] for r in items)

        gm_entry = gm_by_code.get(str(best["code"]))
        if gm_entry is None:  # fall back to any other item's code with a Pic Centric match
            for r in items:
                gm_entry = gm_by_code.get(str(r["code"]))
                if gm_entry:
                    break
        margin_target = gm_entry["target_gm0_pct"] if gm_entry else ""
        image_file = gm_entry["image_file"] if gm_entry else None

        forecast_rows = [fr for r in items for fr in forecast_by_key.get((r["season_label"], r["code"]), [])]
        qty, gross, cogs, gp_pct = totals(forecast_rows)
        dtc_qty, dtc_gross, dtc_cogs, dtc_gp = totals([r for r in forecast_rows if r["channel"] == "DTC"])
        ws_qty, ws_gross, ws_cogs, ws_gp = totals([r for r in forecast_rows if r["channel"] == "Wholesale"])

        def status_for(season_label: str) -> str:
            season_rows = [fr for r in items if r["season_label"] == season_label
                           for fr in forecast_by_key.get((season_label, r["code"]), [])]
            statuses = {fr["status"] for fr in season_rows}
            return "; ".join(sorted(STATUS_LABELS.get(s, s) for s in statuses))

        out_rows.append({
            "_base": base,  # internal join key (version-stripped) for downstream scripts
            "_image_file": image_file,  # internal: filename in outputs/product_images/, for downstream scripts
            "Franchise": best["display_name"],
            "Layer": "; ".join(layers),
            "Business Area": "; ".join(areas),
            "Gender": "; ".join(genders),
            "Gender note": gender_note,
            "Active in FW27": "Y" if "FW27" in seasons else "",
            "Active in SS28": "Y" if "SS28" in seasons else "",
            "Target RRP": rrp_display,
            "RRP varies by style": "Y" if len(rrps) > 1 else "",
            "Consumer Activity": "; ".join(activities),
            "Margin target": margin_target,  # RR-008 -- SS28 only (Pic Centric), FW27 pending
            "Special edition / fit note": "; ".join(special_flags),
            "# styles rolled up": len(items),
            "# active colors": total_colors,
            "Core Assortment FW27": core_fw27.get(base, "No"),
            "Core Assortment SS28": "To confirm",
            "ABS": "Yes" if base in abs_bases else "No",
            "Asia Fit": "Yes" if base in af_bases else "No",
            "Status FW27": status_for("FW27"),
            "Status SS28": status_for("SS28"),
            "Forecast Qty (total)": round(qty) if forecast_rows else "",
            "Forecast GP% (blended)": round(gp_pct, 1) if gp_pct is not None else "",
            "Forecast Qty (DTC)": round(dtc_qty) if dtc_qty else "",
            "Forecast GP% (DTC)": round(dtc_gp, 1) if dtc_gp is not None else "",
            "Forecast Qty (Wholesale)": round(ws_qty) if ws_qty else "",
            "Forecast GP% (Wholesale)": round(ws_gp, 1) if ws_gp is not None else "",
        })

    out_rows.sort(key=lambda r: (r["Layer"], r["Business Area"], r["Franchise"]))
    return out_rows


def write_workbook(rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Franchise Review List"

    headers = [h for h in rows[0].keys() if not h.startswith("_")]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for r in rows:
        ws.append([r[h] for h in headers])

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
    notes.append(["Margin target", "Target GM0% from the FW27+SS28 'Pic Centric' workbooks -- "
                                    "the FW27/SS28 range files themselves have no margin field "
                                    "(RR-008). Blank only for franchises with no Pic Centric "
                                    "match in either season."])
    notes.append(["Consumer Activity", "Raw multi-select field from the range files (e.g. "
                                        "'Trekking / Hiking, Everyday Outdoor') -- union of every "
                                        "distinct activity tag across the franchise's styles, "
                                        "shown as-is, not collapsed into one category. 10 distinct "
                                        "tags exist overall; more than half of all styles carry "
                                        "2+ at once, so this is a reference signal, not a clean "
                                        "single grouping. See RR-017 for why the 'strategic "
                                        "intent' category framing (Trail/Hike+Essentials etc.) "
                                        "this might relate to is NOT used as this list's "
                                        "structure."])
    notes.append(["Core Assortment FW27", "Yes/No, from the flag already in the SS26 tiering "
                                           "workbook (57/1,860 Base+Gender rows = 'Yes'), rolled "
                                           "up to Base grain (Yes if either gender was Yes)."])
    notes.append(["Core Assortment SS28", "No source yet -- the flag was only ever defined for "
                                           "FW27. Left as 'To confirm' for every row until "
                                           "Tobias/Isabelle define it for SS28."])
    notes.append(["# active colors", "Sum of active colorways (the colorway-level Active flag, "
                                      "not just a row listed in the sheet) across every style "
                                      "rolled into this franchise -- both genders, both seasons. "
                                      "A colorway still in the file but not Active this season "
                                      "(e.g. a discontinued or carried-forward-for-reference "
                                      "color) is not counted. CAVEAT: 88/616 main-range style "
                                      "rows (incl. 18 Shell styles) have NO colorway sub-rows at "
                                      "all in this export -- showing 0 here, which may be a real "
                                      "single-color item or a data gap in the export. Worth "
                                      "confirming with Tobias before treating 0 as fact."])
    notes.append(["ABS", "Yes if an ABS or ABS Exclusive variant of this franchise exists "
                          "anywhere in FW27/SS28 (even though RR-006 keeps those styles out "
                          "of the main range count above) -- so a reviewer knows one exists "
                          "before deciding the franchise's main-range fate."])
    notes.append(["Asia Fit", "Yes if an 'AF'-suffixed (Asian Fit) variant of this franchise "
                               "exists anywhere in FW27/SS28 (RR-007) -- always also ABS=Yes, "
                               "since every AF style is tagged ABS Exclusive in the source data."])
    notes.append(["Status FW27 / Status SS28", "Carry Over / New / Update, from the forecast "
                                                 "files' Status field -- shown per season since a "
                                                 "franchise can differ between them (e.g. New in "
                                                 "FW27, Carry Over in SS28). Blank if no forecast "
                                                 "row matched that season (see RR-014)."])
    notes.append(["Forecast Qty / GP% (total, DTC, Wholesale)", "Summed across every style and "
                                                                  "season rolled into the "
                                                                  "franchise, from the two "
                                                                  "forecast files (RR-011/RR-012). "
                                                                  "Blank (not 0) when none of the "
                                                                  "franchise's styles have a "
                                                                  "forecast match -- 25/616 "
                                                                  "main-range style rows don't "
                                                                  "(RR-014). GP% blended vs. "
                                                                  "DTC/Wholesale split -- blended "
                                                                  "hides a real spread, so both "
                                                                  "are shown."])
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
