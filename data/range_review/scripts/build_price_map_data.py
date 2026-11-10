"""
Builds the per-franchise data feeding the Price x Activity Map artifacts
(one standalone HTML chart per layer -- Shell/Insulation/Midlayer already
published, see RR-091 for the Legwear/Tops/Softshell/Baselayer follow-on).

Each franchise needs a single-valued **Activity** (one of the 9 Linden
consumer-activity categories) and a **Silhouette** (Jacket/Vest/Pant/Tee/...)
-- neither field exists in `outputs/workshop_data.json` or the FW27/SS28/FW28
range files themselves. The only source for both is
`data/Assortment Attribution Review(F27).xlsx`, a one-time FW27-dated export
from Linden Mallory (Merch) -- see RR-017's email thread. It is NOT kept in
sync with the range files, so a franchise added/regenerated after that
export won't have a code it covers (RR-091): matching falls back from style
Code to a gender/AF-stripped name match against the same file before giving
up and leaving Activity/Silhouette blank ("Unclassified"/"Undefined" on the
chart, not a guess).

Run: python3 build_price_map_data.py
Output: ../outputs/price_map_data/<layer>.json (one file per LAYERS entry)
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent.parent.parent
WORKSHOP_DATA = HERE.parent / "outputs" / "workshop_data.json"
ATTRIBUTION_FILE = REPO_ROOT / "data" / "Assortment Attribution Review(F27).xlsx"
IMAGE_DIR = HERE.parent / "outputs" / "product_images"
OUT_DIR = HERE.parent / "outputs" / "price_map_data"

# Layers with a published (or about-to-be-published) Price Map. Shell/
# Insulation/Midlayer's originally-published data was hand-built in a prior
# session (not reproducible from this script -- a gap this script closes
# going forward) -- re-running it for those three is expected to recover
# nearly the same but not byte-identical matches, since RR-091 added the
# name-based fallback match that the original build didn't have.
LAYERS = ["Shell", "Insulation", "Midlayer", "Legwear", "Tops", "Softshell", "Baselayer"]

GENDER_AF_SUFFIXES = [" AF", " Men", " Women", " Unisex"]


def norm_name(style: str | None) -> str | None:
    """Strip trailing gender/Asia-Fit tokens so e.g. 'Korp Lite Pant Men AF'
    and the franchise name 'Korp Lite Pant' can match."""
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


def build_layer(layer, franchises, code_lookup, name_lookup):
    out = []
    for d in franchises:
        activity, silhouette = resolve_activity_silhouette(d, code_lookup, name_lookup)
        hs = (d.get("historicalSales") or {}).get("2026", {})
        wh = hs.get("wholesale", {})
        img_file = d.get("imageFile")
        img = None
        if img_file and (IMAGE_DIR / img_file).exists():
            img = "images/" + img_file

        out.append(
            {
                "name": d["franchise"],
                "rrp": d.get("targetRRP"),
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
                # Activity has no actual per-season variation in this source
                # (one FW27-dated export, no SS28/FW28 equivalent) -- the
                # same resolved value is repeated across all four fields,
                # same convention the original Shell/Insulation/Midlayer
                # data used.
                "activityFW27": activity,
                "activitySS28": activity,
                "activityFW28": activity,
                "activityAll": activity,
                "silhouette": silhouette,
            }
        )
    return out


def main():
    workshop_data = json.loads(WORKSHOP_DATA.read_text())
    code_lookup, name_lookup = load_attribution()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    for layer in LAYERS:
        franchises = [d for d in workshop_data if d.get("layer") == layer]
        out = build_layer(layer, franchises, code_lookup, name_lookup)
        unclassified = sum(1 for o in out if o["silhouette"] is None)
        outfile = OUT_DIR / f"{layer.lower()}.json"
        outfile.write_text(json.dumps(out, ensure_ascii=False, indent=None))
        print(
            f"{layer}: {len(out)} franchises, {unclassified} unclassified "
            f"({unclassified / len(out):.0%}) -> {outfile.relative_to(REPO_ROOT)}"
        )


if __name__ == "__main__":
    main()
