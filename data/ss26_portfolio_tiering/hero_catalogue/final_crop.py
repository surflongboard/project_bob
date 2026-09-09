"""
Final tight crop: auto-detected studio-backdrop band, with a top floor to
clear nav headers whose background color reads too close to the backdrop
for pure brightness thresholding, and a proportional bottom trim to clear
sticky search-bar/Add-to-bag overlays that don't touch the sampled edge
columns (so the plain background-color scan can't see them).
"""
from PIL import Image
import numpy as np
import os, json

SRC_DIR = "catalogue_crops/"
OUT_DIR = "catalogue_final/"
os.makedirs(OUT_DIR, exist_ok=True)

MANUAL_BOUNDS = {
    # (top, bottom) in the source (already top=290/bottom=260-cropped) image's own coordinates
    "lim_mid_multi_hood_m": (850, 2006),   # Galaxus: plain white bg, no backdrop band to detect
    "chaos_gtx_jacket_w": (20, 1100),       # Intersport: plain white bg, no backdrop band to detect
}
TOP_FLOOR_OVERRIDE = {
    "rosson_mid_jacket_m": 420,  # unusually tall in-page header on this one
}

WHITE_THRESHOLD = 236
MIN_RUN = 300
TOP_FLOOR = 360
BOTTOM_KEEP_RATIO = 0.78  # trim the trailing ~22% of the detected band (search bar / Add to bag overlay)

TARGET_W, TARGET_H = 900, 1100  # common canvas size, product scaled to fit (contain), centered


def find_backdrop_band(arr):
    h, w, _ = arr.shape
    left = arr[:, 10:30].mean(axis=(1, 2))
    right = arr[:, w - 30:w - 10].mean(axis=(1, 2))
    prof = (left + right) / 2
    notwhite = prof < WHITE_THRESHOLD
    runs, start = [], None
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
    return max(runs, key=lambda r: r[1] - r[0])


def backdrop_fill_color(arr, top, bottom):
    """Sample the crop's own backdrop color (edge columns, mid-band rows) to use as canvas padding."""
    h, w, _ = arr.shape
    mid = (top + bottom) // 2
    band = arr[max(top, mid - 40):min(bottom, mid + 40), :]
    left = band[:, 10:30].reshape(-1, 3)
    right = band[:, w - 30:w - 10].reshape(-1, 3)
    sample = np.concatenate([left, right], axis=0)
    return tuple(int(c) for c in sample.mean(axis=0))


manifest = {}
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
        raw_top, raw_bottom = band if band else (0, h)
        floor = TOP_FLOOR_OVERRIDE.get(key, TOP_FLOOR)
        top = max(raw_top, floor)
        top = min(top, raw_bottom - 200)  # never cross into/past the bottom
        band_len = raw_bottom - top
        bottom = top + band_len * BOTTOM_KEEP_RATIO

    top, bottom = int(top), int(bottom)
    fill = backdrop_fill_color(arr, top, bottom)

    cropped = im.crop((0, top, w, bottom))
    cw, ch = cropped.size
    scale = min(TARGET_W / cw, TARGET_H / ch)
    new_w, new_h = int(cw * scale), int(ch * scale)
    resized = cropped.resize((new_w, new_h), Image.LANCZOS)

    canvas = Image.new("RGB", (TARGET_W, TARGET_H), fill)
    canvas.paste(resized, ((TARGET_W - new_w) // 2, (TARGET_H - new_h) // 2))
    canvas.save(OUT_DIR + key + ".jpg", "JPEG", quality=85, optimize=True)

    manifest[key] = {"top": top, "bottom": bottom, "fill": fill}
    print(f"{key:32s} crop=({top},{bottom}) fill={fill}")

with open("final_crop_manifest.json", "w") as f:
    json.dump(manifest, f)
print("done:", len(manifest), "images")
