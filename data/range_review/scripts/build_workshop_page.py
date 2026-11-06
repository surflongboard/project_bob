"""
Splices workshop_data.json into workshop_page_template.html to produce the
page republished to the live "Range Review Board" artifact
(https://claude.ai/artifact/QaZKbo6yKxFUdHAuWNBx95) for the Nov 12
workshop. Web format chosen over Excel because embedded product images in
the franchise review list "didn't look good" (RR-015) and couldn't survive
Excel's sort/filter anchoring; a web page holds up under both and shows
the image inline.

This script does NOT upload images -- that one-time step (Artifact tool
asset upload, outputs/product_images/<Code>.jpeg -> an asset url) already
happened and its result is recorded in outputs/image_asset_map.json
(local filename -> "/_blob/<id>", 258 entries as of this writing). This
script applies that map automatically (patch_image_urls(), below) so
every build already has real, servable image src values -- it no longer
depends on a manual post-build patch. RR-088: the previous "patch it on
the live artifact after publishing" approach (this docstring's own
earlier wording) was never actually durable -- the patch lived only in
that one publish, not in anything this script produced, so the NEXT
ordinary rebuild-and-republish (no images touched, nothing about that
session suggested checking) silently regenerated the page with bare
local filenames again and broke every image on the live Board. Twice
(RR-086, RR-087), unnoticed until the user reported it. If a NEW image
is added later (new franchise/code), upload it via the Artifact tool's
asset upload and add its filename -> url entry to image_asset_map.json
by hand -- same one-time step as before, just persisted in the repo
afterward instead of living only in the live artifact.

Decisions (Tier/Action per franchise) live in the artifact's own `db`
capability (collection "decisions", doc id = slugified franchise base),
captured live by whoever has the page open during the workshop -- not in
this repo. Pull them out after the workshop with ArtifactData (action
"list", collection "decisions") if they need to feed back into the
franchise review list.

Run: python3 build_workshop_page.py
Output: ../outputs/workshop_page.html
"""
from __future__ import annotations

import json
from pathlib import Path

RR_DIR = Path(__file__).resolve().parents[1]
OUTPUTS_DIR = RR_DIR / "outputs"
TEMPLATE_PATH = Path(__file__).resolve().parent / "workshop_page_template.html"
DATA_PATH = OUTPUTS_DIR / "workshop_data.json"
IMAGE_ASSET_MAP_PATH = OUTPUTS_DIR / "image_asset_map.json"


def patch_image_urls(data: list[dict]) -> int:
    """Replace each franchise's imageFile local filename with its uploaded asset
    url, per RR-088 -- without this, every rebuild regenerates bare filenames
    that render as broken images on the live Board. Returns the count patched;
    leaves any filename with no map entry untouched (not null) so it's still
    visible in the output as something that needs a fresh asset upload, same
    as it would have shown before this function existed."""
    if not IMAGE_ASSET_MAP_PATH.exists():
        print(f"WARNING: {IMAGE_ASSET_MAP_PATH} not found -- image src values "
              "will be bare local filenames and every image will be broken on "
              "the live Board.")
        return 0
    with open(IMAGE_ASSET_MAP_PATH) as f:
        image_map = json.load(f)
    patched = 0
    unmapped = []
    for r in data:
        fname = r.get("imageFile")
        if not fname:
            continue
        url = image_map.get(fname)
        if url:
            r["imageFile"] = url
            patched += 1
        else:
            unmapped.append(fname)
    if unmapped:
        print(f"WARNING: {len(unmapped)} imageFile filename(s) have no entry in "
              f"{IMAGE_ASSET_MAP_PATH.name} -- these will render as broken "
              f"images until uploaded and added to the map: {unmapped}")
    return patched


def build() -> None:
    with open(DATA_PATH) as f:
        data = json.load(f)
    patched = patch_image_urls(data)
    template = TEMPLATE_PATH.read_text()
    out = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    out_path = OUTPUTS_DIR / "workshop_page.html"
    out_path.write_text(out)
    print(f"{len(out)} bytes written to {out_path}")
    print(f"{patched} imageFile values patched to asset URLs from {IMAGE_ASSET_MAP_PATH.name}.")


if __name__ == "__main__":
    build()
