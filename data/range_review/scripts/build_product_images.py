"""
Extracts a product-image repository and the GM% fields from the FW27 and
SS28 "Pic Centric" workbooks -- the first sources in this workstream that
actually have a margin field (RR-008 flagged Target RRP-only, no margin,
in the FW27/SS28 range files; these have Target GM0%/Sample GM0%/GM0%).

Images: each sheet re-anchors the SAME two images (drawing column 3 =
"Active Colorways Images", a real colorway photo; column 4 = "Image", a
flat technical sketch) on every colorway sub-row of a style, not just
once. Column 3's photo is used as the representative image (more
recognizable at a glance than a line sketch) -- saved one per style Code
to outputs/product_images/<Code>.jpeg.

GM%: Target GM0% is the only true *target* field (matches "margin target"
as asked); Sample GM0%/GM0% are also pulled since they're right there,
but are NOT used as a margin-target substitute anywhere else in this
workstream -- a target and a realized/sample number are different things
(same principle as RR-008).

Both files share the exact same column layout (verified) and 215 of
FW27's 379 / SS28's 336 Active codes overlap (a style running both
seasons under the same Code) -- for those, Target GM0% matches exactly in
213/215 cases, RRP differs in the other 2 (a planned FW27->SS28 price
change, not a data error). Processed FW27 first, SS28 second, so SS28's
value wins on the rare conflict -- consistent with this workstream's
general "latest/most-current generation wins" convention (RR-003)
elsewhere. One combined lookup, not two separate per-season files, since
build_franchise_list.py just needs "does this franchise have a Pic
Centric match anywhere" rather than a season-specific answer.

Run: python3 build_product_images.py
Outputs: ../outputs/product_images/<Code>.jpeg (one per style, either season)
         ../outputs/gm_by_code.json (Code -> GM%/image/source-season
         fields, used by build_franchise_list.py to fill in Margin target)
"""
from __future__ import annotations

import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl

RR_DIR = Path(__file__).resolve().parents[1]
IMAGES_INPUT_DIR = RR_DIR / "inputs" / "images"
FW27_PIC_CENTRIC_PATH = IMAGES_INPUT_DIR / "FW27_Pic_Centric_28_oct.xlsx"
SS28_PIC_CENTRIC_PATH = IMAGES_INPUT_DIR / "SS28_Pic_Centric_28_oct.xlsx"
OUTPUTS_DIR = RR_DIR / "outputs"
IMAGES_DIR = OUTPUTS_DIR / "product_images"

NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
PHOTO_COL = 3  # 0-indexed drawing column for "Active Colorways Images"


def parse_image_anchors(zf: zipfile.ZipFile) -> dict[int, str]:
    """Excel row (1-indexed) -> media filename, for the representative photo column only."""
    rels_root = ET.fromstring(zf.read("xl/drawings/_rels/drawing1.xml.rels"))
    rid_to_media = {
        rel.get("Id"): rel.get("Target").rsplit("/", 1)[-1]
        for rel in rels_root
    }
    root = ET.fromstring(zf.read("xl/drawings/drawing1.xml"))
    row_to_media = {}
    for anchor in root.findall("xdr:twoCellAnchor", NS):
        frm = anchor.find("xdr:from", NS)
        col = int(frm.find("xdr:col", NS).text)
        if col != PHOTO_COL:
            continue
        row = int(frm.find("xdr:row", NS).text) + 1  # xdr rows are 0-indexed
        blip = anchor.find(".//a:blip", NS)
        rid = blip.get(R_EMBED) if blip is not None else None
        if rid in rid_to_media:
            row_to_media[row] = rid_to_media[rid]
    return row_to_media


def _pct(v):
    return round(v * 100, 1) if isinstance(v, (int, float)) else None


def process_file(path: Path, season_label: str, gm_by_code: dict[str, dict]) -> tuple[int, int]:
    zf = zipfile.ZipFile(path)
    row_to_media = parse_image_anchors(zf)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Sheet"]

    seen_codes = set()
    saved_images = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        code = row[5]
        if code is None or code in seen_codes:
            continue
        seen_codes.add(code)

        entry = {
            "style": row[0],
            "source_season": season_label,
            "target_rrp": row[14],
            "target_gm0_pct": _pct(row[19]),
            "sample_gm0_pct": _pct(row[20]),
            "gm0_pct": _pct(row[21]),
            "image_file": None,
        }
        media = row_to_media.get(row_idx)
        if media:
            out_path = IMAGES_DIR / f"{code}.jpeg"
            if not out_path.exists():
                out_path.write_bytes(zf.read(f"xl/media/{media}"))
                saved_images += 1
            entry["image_file"] = out_path.name

        gm_by_code[str(code)] = entry  # SS28 processed after FW27 wins on overlap

    return len(seen_codes), saved_images


def build() -> None:
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    gm_by_code: dict[str, dict] = {}

    fw27_codes, fw27_images = process_file(FW27_PIC_CENTRIC_PATH, "FW27", gm_by_code)
    ss28_codes, ss28_images = process_file(SS28_PIC_CENTRIC_PATH, "SS28", gm_by_code)

    with open(OUTPUTS_DIR / "gm_by_code.json", "w") as f:
        json.dump(gm_by_code, f, indent=2)

    print(f"FW27: {fw27_codes} unique codes, {fw27_images} new images")
    print(f"SS28: {ss28_codes} unique codes, {ss28_images} new images")
    print(f"combined: {len(gm_by_code)} unique codes in {OUTPUTS_DIR / 'gm_by_code.json'}")
    missing_image = sum(1 for v in gm_by_code.values() if v["image_file"] is None)
    print(f"{missing_image} codes with no representative image found")


if __name__ == "__main__":
    build()
