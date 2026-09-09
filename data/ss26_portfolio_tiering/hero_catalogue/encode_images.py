"""
Regenerates catalogue_images_b64.json from catalogue_final/*.jpg -- the
34 uniform-size, product-only JPEGs produced by the auto_crop.py ->
final_crop.py pipeline. gen_catalogue.py reads catalogue_images_b64.json
directly (embedding it as base64 data URIs, since the published artifact
can't reference external image files -- see ASSUMPTIONS_REGISTER.md).

Run from this directory after re-cropping a new/updated photo set:
    python3 encode_images.py
"""
import base64
import json
import os

SRC_DIR = "catalogue_final/"
OUT_PATH = "catalogue_images_b64.json"

images = {}
for fname in sorted(os.listdir(SRC_DIR)):
    if not fname.endswith(".jpg"):
        continue
    key = fname[:-4]
    with open(SRC_DIR + fname, "rb") as f:
        images[key] = base64.b64encode(f.read()).decode("ascii")

with open(OUT_PATH, "w") as f:
    json.dump(images, f)

print(f"encoded {len(images)} images -> {OUT_PATH} ({os.path.getsize(OUT_PATH):,} bytes)")
