"""
Builds the per-franchise data feeding the 5 equipment Price Maps (RR-092) --
Footwear, Packs, Bags, Sleeping Bags, Accessories -- grouping the 11
remaining (non-apparel) layers that weren't covered by the Shell/
Insulation/Midlayer/Legwear/Tops/Softshell/Baselayer maps (build_price_map_
data.py).

Same Activity/Silhouette source and code-then-name matching as
build_price_map_data.py (see that script's docstring and RR-091) --
`data/Assortment Attribution Review(F27).xlsx` covers Hardware/Footwear
rows too, not just Apparel. But Silhouette is too coarse or uniform for
some of these groups (e.g. "Backpack" for every Daypack/Hiking Pack/
Mountain Pack/Trekking Pack alike, "Accessory" for most Accessories), so
this script ALSO pulls the range files' own **Product Type** field
(Code -> capacity band / accessory type), which differentiates where
Silhouette doesn't -- not previously part of the franchise list pipeline,
so pulled directly here via a small loader rather than extending
build_franchise_list.py's load_style_rows()/load_style_rows_flat() (which
don't expose it) for a one-off need.

Grouping (RR-092, user-directed 10-Nov-2026):
- Footwear: Hiking Shoes + Trail Shoes + Trekking Shoes + Casual Shoes
- Packs: Daypacks + Hiking Packs + Mountain Packs + Trekking Packs
- Bags: standalone
- Sleeping Bags: standalone
- Accessories: standalone

Run: python3 build_equipment_price_map_data.py
Output: ../outputs/price_map_data/<group>.json
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent.parent.parent
WORKSHOP_DATA = HERE.parent / "outputs" / "workshop_data.json"
ATTRIBUTION_FILE = REPO_ROOT / "data" / "Assortment Attribution Review(F27).xlsx"
IMAGE_DIR = HERE.parent / "outputs" / "product_images"
OUT_DIR = HERE.parent / "outputs" / "price_map_data"

RANGE_FILES_NESTED = ["FW27_28_Oct_2026.xlsx", "SS28_28_Oct_2026.xlsx"]
RANGE_FILE_FLAT = "FW28_2_Nov_2026.xlsx"

GROUPS = {
    "footwear": ["Hiking Shoes", "Trail Shoes", "Trekking Shoes", "Casual Shoes"],
    "packs": ["Daypacks", "Hiking Packs", "Mountain Packs", "Trekking Packs"],
    "bags": ["Bags"],
    "sleeping_bags": ["Sleeping Bags"],
    "accessories": ["Accessories"],
}

GENDER_AF_SUFFIXES = [" AF", " Men", " Women", " Unisex"]


def norm_name(style: str | None) -> str | None:
    if not style:
        return None
    s = style.strip()
    changed = True
    while changed:
        changed = False
        for tok in GENDER_AF_SUFFIXES:
            if s.endswith(tok):
                s = s[: -len(tok)].strip()
                changed = True
    return s or None


def load_attribution():
    wb = openpyxl.load_workbook(ATTRIBUTION_FILE, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(min_row=3, values_only=True))

    def get(r, i):
        return r[i] if i < len(r) else None

    code_lookup: dict[str, dict] = {}
    name_lookup: dict[str, list[tuple]] = {}
    for r in rows:
        code = get(r, 1)
        style = get(r, 2)
        activity = get(r, 31)
        silhouette = get(r, 33)
        if code is not None:
            code_str = str(int(code)) if isinstance(code, float) else str(code)
            code_lookup[code_str] = {"activity": activity, "silhouette": silhouette}
        n = norm_name(style)
        if n and (activity or silhouette):
            name_lookup.setdefault(n, []).append((activity, silhouette))
    return code_lookup, name_lookup


def load_product_type_by_code() -> dict[str, str]:
    """Code -> Product Type (range files' own column -- capacity band for
    Packs/Bags, Beanie/Cap/.../"Other Accessories" for Accessories, Low/
    Mid/High for Footwear). First value wins across files/seasons -- not
    observed to vary per code (same convention as build_price_map_data.py's
    Activity/Silhouette, which also found this stable)."""
    out: dict[str, str] = {}
    for fname in RANGE_FILES_NESTED:
        wb = openpyxl.load_workbook(HERE.parent / "inputs" / fname, read_only=True)
        ws = wb["Sheet"]
        for r in ws.iter_rows(min_row=3, values_only=True):
            if r[0] is not None and r[2] is not None:
                code = str(int(r[2])) if isinstance(r[2], float) else str(r[2])
                if r[13]:
                    out.setdefault(code, r[13])
    wb = openpyxl.load_workbook(HERE.parent / "inputs" / RANGE_FILE_FLAT, read_only=True)
    ws = wb["Sheet"]
    seen: set[str] = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[2] is not None:
            code = str(int(r[2])) if isinstance(r[2], float) else str(r[2])
            if code not in seen:
                seen.add(code)
                if r[13]:
                    out.setdefault(code, r[13])
    return out


def resolve_activity_silhouette(franchise, code_lookup, name_lookup):
    codes = [str(c) for c in franchise.get("styleCodes") or []]
    activity = silhouette = None
    for c in codes:
        hit = code_lookup.get(c)
        if hit:
            activity = activity or hit["activity"]
            silhouette = silhouette or hit["silhouette"]
    if activity is None and silhouette is None:
        for key in (franchise.get("base"), franchise.get("franchise")):
            n = norm_name(key)
            if n and n in name_lookup:
                vals = name_lookup[n]
                acts = {v[0] for v in vals if v[0]}
                sils = {v[1] for v in vals if v[1]}
                activity = next(iter(acts), None)
                silhouette = next(iter(sils), None)
                break
    return activity, silhouette


def resolve_product_type(franchise, pt_lookup) -> str | None:
    for c in [str(x) for x in (franchise.get("styleCodes") or [])]:
        if c in pt_lookup:
            return pt_lookup[c]
    return None


_RANGE_RE = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\s*$")


def parse_rrp(raw) -> tuple[float | None, str | None, bool]:
    """Target RRP is usually a number, but a few franchises (rrpVaries=true)
    carry a string range like "2000-2200" instead -- plot the midpoint,
    keep the original range for display, and flag it so the tooltip can
    say "varies" rather than imply a single exact price."""
    if raw is None:
        return None, None, False
    if isinstance(raw, (int, float)):
        return raw, None, False
    m = _RANGE_RE.match(str(raw))
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        return (lo + hi) / 2, f"{lo}–{hi}", True
    return None, None, False


def build_group(franchises, code_lookup, name_lookup, pt_lookup):
    out = []
    for d in franchises:
        activity, silhouette = resolve_activity_silhouette(d, code_lookup, name_lookup)
        product_type = resolve_product_type(d, pt_lookup)
        hs = (d.get("historicalSales") or {}).get("2026", {})
        wh = hs.get("wholesale", {})
        img_file = d.get("imageFile")
        img = None
        if img_file and (IMAGE_DIR / img_file).exists():
            img = "images/" + img_file
        plot_rrp, rrp_display, rrp_varies = parse_rrp(d.get("targetRRP"))

        out.append(
            {
                "name": d["franchise"],
                "origLayer": d["layer"],
                "rrp": plot_rrp,
                "rrpDisplay": rrp_display,
                "rrpVaries": rrp_varies,
                "abs": bool(d.get("abs")),
                "asiaFit": bool(d.get("asiaFit")),
                "img": img,
                "codes": [str(c) for c in (d.get("styleCodes") or [])],
                "sales2026": hs.get("units"),
                "gp2026": hs.get("gpPct"),
                "whSales2026": wh.get("units"),
                "whGp2026": wh.get("gpPct"),
                "activeFW27": bool(d.get("activeFW27")),
                "activeSS28": bool(d.get("activeSS28")),
                "activeFW28": bool(d.get("activeFW28")),
                "activity": activity,
                "silhouette": silhouette,
                "productType": product_type,
            }
        )
    return out


# RR-092 follow-up: a focused sub-view of Accessories for its head-worn items only,
# per the user's "create Accessories (Headwear)" request -- defined by Product Type
# (the range files' own field) rather than Linden's coarser Silhouette tag ("Headwear"),
# since Product Type also catches 5 franchises Linden's attribution export never
# covered at all (no Activity/Silhouette match) but which still have a real Product Type.
HEADWEAR_PRODUCT_TYPES = {"Beanie", "Cap", "Hat", "Headband"}


def main() -> None:
    workshop_data = json.loads(WORKSHOP_DATA.read_text())
    code_lookup, name_lookup = load_attribution()
    pt_lookup = load_product_type_by_code()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    accessories_out = None
    for group, layers in GROUPS.items():
        franchises = [d for d in workshop_data if d.get("layer") in layers]
        out = build_group(franchises, code_lookup, name_lookup, pt_lookup)
        no_rrp = sum(1 for o in out if not o["rrp"])
        no_activity = sum(1 for o in out if not o["activity"])
        outfile = OUT_DIR / f"{group}.json"
        outfile.write_text(json.dumps(out, ensure_ascii=False, indent=None))
        print(
            f"{group}: {len(out)} franchises, {no_rrp} no RRP, {no_activity} no activity "
            f"-> {outfile.relative_to(REPO_ROOT)}"
        )
        if group == "accessories":
            accessories_out = out

    headwear = [o for o in accessories_out if o["productType"] in HEADWEAR_PRODUCT_TYPES]
    headwear_file = OUT_DIR / "accessories_headwear.json"
    headwear_file.write_text(json.dumps(headwear, ensure_ascii=False, indent=None))
    print(f"accessories_headwear: {len(headwear)} franchises -> {headwear_file.relative_to(REPO_ROOT)}")


if __name__ == "__main__":
    main()
