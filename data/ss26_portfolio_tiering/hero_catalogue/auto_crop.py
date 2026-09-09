"""
Detect the studio-backdrop band (product photo area) in each screenshot by
scanning row-wise edge brightness, and crop tightly to it. Haglofs.com
product pages consistently use a light grey/beige studio backdrop for the
photo, distinctly darker than the white/cream nav bar and caption/price
areas above and below it.
"""
from PIL import Image
import numpy as np
import os, json

SRC_DIR = "catalogue_crops/"
OUT_DIR = "catalogue_final/"
os.makedirs(OUT_DIR, exist_ok=True)

# manual overrides for images with no distinct backdrop band (pure white bg
# retailer pages): (top, bottom) in the ALREADY-CROPPED (top=290,bottom=260
# removed) source image's own pixel coordinates.
MANUAL_BOUNDS = {
    "lim_mid_multi_hood_m": (260, 1780),   # Galaxus: below breadcrumb, above nothing (goes to bottom)
    "chaos_gtx_jacket_w": (0, 1350),        # Intersport: crop already tight from custom top/bottom crop
}

WHITE_THRESHOLD = 236
MIN_RUN = 300  # px, filters out small non-backdrop dips (badges, banners)

def find_backdrop_band(arr):
    h, w, _ = arr.shape
    left = arr[:, 10:30].mean(axis=(1, 2))
    right = arr[:, w - 30:w - 10].mean(axis=(1, 2))
    prof = (left + right) / 2
    notwhite = prof < WHITE_THRESHOLD
    runs = []
    start = None
    for y in range(h):
        if notwhite[y] and start is None:
            start = y
        elif not notwhite[y] and start is not None:
            runs.append((start, y))
            start = None
    if start is not None:
        runs.append((start, h))
    runs = [r for r in runs if r[1] - r[0] >= MIN_RUN]
    if not runs:
        return None
    # longest run
    return max(runs, key=lambda r: r[1] - r[0])

results = {}
for fname in sorted(os.listdir(SRC_DIR)):
    if not fname.endswith(".png"):
        continue
    key = fname[:-4]
    im = Image.open(SRC_DIR + fname).convert("RGB")
    arr = np.array(im).astype(int)
    h, w, _ = arr.shape

    if key in MANUAL_BOUNDS:
        top, bottom = MANUAL_BOUNDS[key]
    else:
        band = find_backdrop_band(arr)
        if band is None:
            print(f"!! no band found for {key}, using full image")
            top, bottom = 0, h
        else:
            top, bottom = band

    pad = 14
    top = max(0, top - pad)
    bottom = min(h, bottom + pad)
    results[key] = (top, bottom, h)
    print(f"{key:32s} h={h:5d}  crop=({top},{bottom})  kept={bottom-top}px ({(bottom-top)/h*100:.0f}%)")

with open("crop_bounds.json", "w") as f:
    json.dump(results, f)
