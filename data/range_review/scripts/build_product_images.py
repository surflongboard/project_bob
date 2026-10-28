"""
Extracts a product-image repository and the GM% fields from the SS28
"Pic Centric" workbook -- the first source in this workstream that
actually has a margin field (RR-008 flagged Target RRP-only, no margin,
in the FW27/SS28 range files; this file has Target GM0%/Sample GM0%/GM0%).

Images: the sheet re-anchors the SAME two images (drawing column 3 =
"Active Colorways Images", a real colorway photo; column 4 = "Image", a
flat technical sketch) on every colorway sub-row of a style, not just
once -- 4,192 anchors but only 636 distinct embedded images (confirmed by
parsing xl/drawings/drawing1.xml's actual blip r:embed, not just anchor
count). Column 3's photo is used as the representative image (more
recognizable at a glance than a line sketch) -- saved one per style Code
to outputs/product_images/<Code>.jpeg.

GM%: Target GM0% (col T) is the only true *target* field (matches
"margin target" as asked); Sample GM0%/GM0% are also pulled since they're
right there, but are NOT used as a margin-target substitute anywhere else
in this workstream -- a target and a realized/sample number are different
things (same principle as RR-008).

This file is SS28 only (per its filename) -- covers 336 Active SS28
styles. No FW27 equivalent yet.

Run: python3 build_product_images.py
Outputs: ../outputs/product_images/<Code>.jpeg (one per style)
         ../outputs/ss28_gm_by_code.json (Code -> GM%/image fields, used
         by build_franchise_list.py to fill in Margin target for SS28)
"""
from __future__ import annotations

import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl

RR_DIR = Path(__file__).resolve().parents[1]
PIC_CENTRIC_PATH = RR_DIR / "inputs" / "images" / "SS28_Pic_Centric_28_oct.xlsx"
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


def build() -> None:
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    zf = zipfile.ZipFile(PIC_CENTRIC_PATH)
    row_to_media = parse_image_anchors(zf)

    wb = openpyxl.load_workbook(PIC_CENTRIC_PATH, read_only=True, data_only=True)
    ws = wb["Sheet"]

    gm_by_code: dict[str, dict] = {}
    saved_images = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        code = row[5]
        if code is None or code in gm_by_code:
            continue
        def pct(v):
            return round(v * 100, 1) if isinstance(v, (int, float)) else None

        gm_by_code[str(code)] = {
            "style": row[0],
            "target_rrp": row[14],
            "target_gm0_pct": pct(row[19]),
            "sample_gm0_pct": pct(row[20]),
            "gm0_pct": pct(row[21]),
            "image_file": None,
        }
        media = row_to_media.get(row_idx)
        if media:
            out_path = IMAGES_DIR / f"{code}.jpeg"
            if not out_path.exists():
                out_path.write_bytes(zf.read(f"xl/media/{media}"))
                saved_images += 1
            gm_by_code[str(code)]["image_file"] = out_path.name

    with open(OUTPUTS_DIR / "ss28_gm_by_code.json", "w") as f:
        json.dump(gm_by_code, f, indent=2)

    print(f"{len(gm_by_code)} unique SS28 codes, {saved_images} images saved to {IMAGES_DIR}")
    missing_image = sum(1 for v in gm_by_code.values() if v["image_file"] is None)
    print(f"{missing_image} codes with no representative image found")


if __name__ == "__main__":
    build()
