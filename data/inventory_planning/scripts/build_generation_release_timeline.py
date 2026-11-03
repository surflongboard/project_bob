"""
Generation-change-season list for the lifecycle/renewal discussion --
"as of SS28 or FW28, how old is the current generation, and is it time
to renew it (SS29? FW29?)."

Source: inputs/product_master/Product_2026_11_02_14_49_01.xlsx (received
3-Nov-2026), one row per product Code with its own ProductReleaseSeason
and ProductLastSeason fields -- a real release-season record per code,
not a proxy. This SUPERSEDES IP-005's stock-file-based proxy for dating
generations: IP-005 could only see a generation if it still had leftover
stock in one Oct-2026 snapshot (a floor on history, invisible once sold
through); this file carries every code's release season regardless of
current stock, back to SS20. IP-005's output/script are left in place
(the Range Review Board's "Generation timeline" card still reads it,
RR-025) -- swapping the Board's own feed to this source is a separate,
not-yet-made decision; see IP-011.

Generation grouping is base-name only (RR-001: Range Review drops the
gender split SS26 Portfolio Tiering keeps), same as every other
range_review/inventory_planning script -- `strip_special_edition` +
`ss26_lib.base_name` + `ss26_lib.version_token`, imported, never
re-derived.

Two real caveats, carried into the Notes sheet:
  - 884 of 3,268 codes (27%) have no ProductReleaseSeason at all --
    excluded from a generation's Release Season when other codes in the
    same generation have one; if ALL of a generation's codes are blank,
    its Release Season is blank too ("blank not zero", RR-014).
  - This file's own season coverage ends at SS28 -- a code newly
    launched in FW28 (per Range Review's RR-026 FW28 integration) won't
    appear here with a FW28 release tag yet. Cross-referenced against
    `build_franchise_list.build_rows()` so a franchise's CURRENT
    generation is correctly identified even when its latest codes are
    FW28-only and absent from this file.

Run: python3 build_generation_release_timeline.py
Output: ../outputs/Generation_Release_Timeline.xlsx
"""
from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

IP_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = IP_DIR.parents[1]
RR_DIR = REPO_ROOT / "data" / "range_review"
sys.path.insert(0, str(REPO_ROOT / "data" / "ss26_portfolio_tiering" / "scripts"))
import ss26_lib as lib  # noqa: E402
sys.path.insert(0, str(RR_DIR / "scripts"))
from build_franchise_list import strip_special_edition, build_rows  # noqa: E402

PRODUCT_PATH = IP_DIR / "inputs" / "product_master" / "Product_2026_11_02_14_49_01.xlsx"
OUTPUTS_DIR = IP_DIR / "outputs"

NOW = "SS28"  # the lifecycle discussion's reference point -- the last season this file itself covers


def season_key(s: str | None) -> tuple[int, int] | None:
    if not s:
        return None
    s = s.strip()
    if len(s) < 3:
        return None
    prefix, year = s[:2], s[2:]
    if prefix not in ("SS", "FW") or not year.isdigit():
        return None
    return (int(year), 0 if prefix == "SS" else 1)


def season_display(key: tuple[int, int] | None) -> str:
    if key is None:
        return ""
    year, half = key
    return ("SS" if half == 0 else "FW") + f"{year:02d}"


def seasons_between(k1: tuple[int, int] | None, k2: tuple[int, int] | None) -> int | None:
    """Count of half-year steps from k1 to k2 (e.g. SS28 -> FW28 = 1)."""
    if k1 is None or k2 is None:
        return None
    return (k2[0] * 2 + k2[1]) - (k1[0] * 2 + k1[1])


def load_product_rows() -> list[dict]:
    wb = openpyxl.load_workbook(PRODUCT_PATH, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        code, style, gender, area = r[1], r[2], r[3], r[4]
        if not style:
            continue
        out.append({
            "code": str(code), "style": style, "gender": gender, "area": area,
            "release": season_key(r[9]), "last": season_key(r[10]),
        })
    return out


def group_by_generation(rows: list[dict]) -> list[dict]:
    groups: dict[tuple[str, str | None], dict] = {}
    for r in rows:
        stripped, _ = strip_special_edition(r["style"])
        base = lib.base_name(stripped)
        version = lib.version_token(stripped)
        key = (base, version)
        g = groups.setdefault(key, {
            "base": base, "version": version, "_rank": lib.RANK.get(version, 0),
            "codes": [], "styles": set(), "areas": set(),
            "release_keys": [], "last_keys": [],
        })
        g["codes"].append(r["code"])
        g["styles"].add(r["style"])
        if r["area"]:
            g["areas"].add(r["area"])
        if r["release"] is not None:
            g["release_keys"].append(r["release"])
        if r["last"] is not None:
            g["last_keys"].append(r["last"])

    out = []
    for g in groups.values():
        out.append({
            "base": g["base"], "version": g["version"] or "", "_rank": g["_rank"],
            "codes": sorted(set(g["codes"])),
            "styles": sorted(g["styles"]),
            "areas": sorted(g["areas"]),
            "release_key": min(g["release_keys"]) if g["release_keys"] else None,
            "last_key": max(g["last_keys"]) if g["last_keys"] else None,
            "n_codes_no_release": sum(1 for r in rows
                                       if str(r["code"]) in g["codes"] and r["release"] is None),
        })
    return out


def build_generation_table() -> list[dict]:
    rows = load_product_rows()
    gens = group_by_generation(rows)

    by_base: dict[str, list[dict]] = defaultdict(list)
    for g in gens:
        by_base[g["base"]].append(g)
    multi_gen_bases = {b for b, gs in by_base.items() if len(gs) > 1}

    # cross-reference the active range for "is this franchise's current generation still sold,
    # and in which season(s)" -- needed because a brand-new FW28 code may not carry a release
    # season in this file at all (it postdates the file's own SS28 ceiling), so without this
    # cross-check the "current generation" could look stale when it's actually just-launched.
    active_rows = build_rows()
    active_by_base = {r["_base"]: r for r in active_rows}

    out = []
    for base, gs in by_base.items():
        gs_sorted = sorted(gs, key=lambda g: (g["_rank"], g["release_key"] or (9999, 9)))
        is_multi = base in multi_gen_bases
        active = active_by_base.get(base)
        for i, g in enumerate(gs_sorted):
            label = g["version"]
            if not label:
                label = "Generation 1" if is_multi else ""
            next_gen = gs_sorted[i + 1] if i + 1 < len(gs_sorted) else None
            is_current = (i == len(gs_sorted) - 1)
            seasons_ran = (seasons_between(g["release_key"], next_gen["release_key"])
                           if next_gen and g["release_key"] and next_gen["release_key"] else None)
            age_as_of_now = seasons_between(g["release_key"], season_key(NOW)) if (
                is_current and g["release_key"]) else None
            out.append({
                "Franchise": base,
                "Generation": label,
                "Codes": ", ".join(g["codes"]),
                "Style Names": "; ".join(g["styles"]),
                "Business Area": "; ".join(g["areas"]),
                "Release Season": season_display(g["release_key"]),
                "Last Known Season": season_display(g["last_key"]),
                "Succeeded By (next gen. Release Season)": season_display(next_gen["release_key"]) if next_gen else "",
                "Seasons This Generation Ran": seasons_ran if seasons_ran is not None else "",
                "Is Current Generation": "Y" if is_current else "N",
                "Active in range now": (
                    "; ".join(s for s, flag in (
                        ("FW27", active["Active in FW27"] == "Y"),
                        ("SS28", active["Active in SS28"] == "Y"),
                        ("FW28", active["Active in FW28"] == "Y"),
                    ) if flag) if (active and is_current) else ""
                ),
                "Seasons Since Release (as of SS28)": age_as_of_now if age_as_of_now is not None else "",
                "# codes with no Release Season tag": g["n_codes_no_release"],
            })
    out.sort(key=lambda r: (r["Franchise"], r["Release Season"] or "zzz"))
    return out


def write_workbook(rows: list[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def write_sheet(title: str, data: list[dict]) -> None:
        ws = wb.create_sheet(title)
        if not data:
            return
        headers = list(data[0].keys())
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
        for r in data:
            ws.append(list(r.values()))
        for i, h in enumerate(headers, start=1):
            width = max(len(h), max((len(str(r[h])) for r in data), default=0)) + 2
            ws.column_dimensions[get_column_letter(i)].width = min(width, 50)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    by_franchise: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        by_franchise[r["Franchise"]].append(r)
    multi_gen_rows = [r for fr, rs in by_franchise.items() if len(rs) > 1 for r in rs]

    change_rows = [r for r in multi_gen_rows if r["Release Season"] or r["Succeeded By (next gen. Release Season)"]]

    current_active_rows = [r for r in rows if r["Is Current Generation"] == "Y" and r["Active in range now"]]
    current_active_rows.sort(key=lambda r: -(r["Seasons Since Release (as of SS28)"]
                                              if isinstance(r["Seasons Since Release (as of SS28)"], int) else -1))

    write_sheet("Generation Change Seasons", change_rows)
    write_sheet("Current Generation Age (active)", current_active_rows)
    write_sheet("All Generations (full history)", rows)

    notes = wb.create_sheet("Notes")
    notes.append(["What this is", "A generation-by-generation release/retirement timeline built from "
                                   "the Product master file's own ProductReleaseSeason/ProductLastSeason "
                                   "fields -- a real per-code record, not the stock-snapshot proxy IP-005 "
                                   "used. 'Generation Change Seasons' is the main deliverable: for every "
                                   "multi-generation franchise, when each generation released and when the "
                                   "NEXT one succeeded it. 'Current Generation Age (active)' restricts to "
                                   "franchises still active in FW27/SS28/FW28 (per Range Review's own "
                                   "franchise list), sorted oldest-current-generation first, to flag "
                                   "renewal candidates for the team to judge -- no age threshold is set "
                                   "here, since no confirmed renewal cadence exists (same Open/Provisional "
                                   "discipline as IP-005)."])
    notes.append(["Supersedes IP-005 for dating accuracy", "IP-005's proxy only saw a generation if it "
                                                             "still had leftover stock in the 29-Oct-2026 "
                                                             "stock snapshot -- a generation fully sold "
                                                             "through was invisible. This file carries "
                                                             "every code's own release season back to "
                                                             "SS20 regardless of current stock. The Range "
                                                             "Review Board's 'Generation timeline' card "
                                                             "still reads IP-005's output (RR-025) -- "
                                                             "swapping it to this source is a separate, "
                                                             "not-yet-made decision."])
    notes.append(["Known limitation 1", "884 of 3,268 codes (27%) have no ProductReleaseSeason tag at "
                                         "all. A generation's Release Season is blank only if EVERY code "
                                         "in it is blank; otherwise the other codes' dates are used."])
    notes.append(["Known limitation 2", "This file's own season coverage ends at SS28 -- a code launched "
                                         "only in FW28 (Range Review's RR-026 integration) has no release "
                                         "tag here yet. 'Current Generation Age' is cross-referenced "
                                         "against Range Review's own FW27/SS28/FW28 active flags so a "
                                         "just-launched FW28 generation isn't miscounted as having no "
                                         "current generation at all -- but its age in seasons can't be "
                                         "computed from this file until a later Product master export "
                                         "tags its FW28 release."])
    notes.append(["Status", "Provisional -- release/last-season dates are a real per-code record (better "
                             "than IP-005's proxy) but not independently confirmed against a launch "
                             "calendar. Flag any specific date to the team before treating it as settled."])
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    notes.column_dimensions["A"].width = 24
    notes.column_dimensions["B"].width = 100

    wb.save(path)


if __name__ == "__main__":
    rows = build_generation_table()
    out_path = OUTPUTS_DIR / "Generation_Release_Timeline.xlsx"
    write_workbook(rows, out_path)

    by_franchise: dict[str, int] = defaultdict(int)
    for r in rows:
        by_franchise[r["Franchise"]] += 1
    n_multi = sum(1 for v in by_franchise.values() if v > 1)
    n_active_current = sum(1 for r in rows if r["Is Current Generation"] == "Y" and r["Active in range now"])
    print(f"{len(rows)} generation rows across {len(by_franchise)} franchises written to {out_path}")
    print(f"{n_multi} franchises show 2+ generations across this file's full history (SS20-SS28)")
    print(f"{n_active_current} currently-active franchises have an age-since-release figure")
