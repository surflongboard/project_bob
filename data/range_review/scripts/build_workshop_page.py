"""
Splices workshop_data.json into workshop_page_template.html to produce the
page republished to the live "Range Review Board" artifact
(https://claude.ai/artifact/QaZKbo6yKxFUdHAuWNBx95) for the Nov 12
workshop. Web format chosen over Excel because embedded product images in
the franchise review list "didn't look good" (RR-015) and couldn't survive
Excel's sort/filter anchoring; a web page holds up under both and shows
the image inline.

This script does NOT upload images or publish the page -- those steps go
through Claude's Artifact tool (asset upload, then publish), not this
repo's Python pipeline, since they talk to the claude.ai artifact service
rather than producing a local file. Run build_workshop_data.py first; the
output of this script (outputs/workshop_page.html) still has each
franchise's imageFile as a local filename (e.g. "608172.jpeg") rather than
a real asset URL -- swap those in after uploading each file in
outputs/product_images/ to the artifact and getting back its url, same
as the live page already does.

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


def build() -> None:
    with open(DATA_PATH) as f:
        data = json.load(f)
    template = TEMPLATE_PATH.read_text()
    out = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    out_path = OUTPUTS_DIR / "workshop_page.html"
    out_path.write_text(out)
    print(f"{len(out)} bytes written to {out_path}")
    print("NOTE: imageFile values are still local filenames -- upload "
          "outputs/product_images/ to the artifact and patch in the real "
          "asset URLs before publishing.")


if __name__ == "__main__":
    build()
