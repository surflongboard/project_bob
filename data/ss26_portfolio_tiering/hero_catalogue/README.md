# Hero Tier Catalogue — generator

A visual-review test build: one photo per Hero/Near-Hero franchise (34
total), grouped by layer, with FY25/YTD2026 Sales/Units/GM% per card and a
Keep/Merge/Update decision control — an alternative to reviewing the same
34 franchises as a text list. Published as an artifact (see
`ASSUMPTIONS_REGISTER.md`'s live-links section for the current URL).

**This was a chat-only build for several iterations before being
committed here** — flagged as a known gap in `LAUNCH_READINESS_CHECKLIST.md`
Section E and now closed, images included (per business direction,
9 Sep 2026: internal test use only — see the caveat below, unchanged).

## What's here

| File | What it does |
|---|---|
| `gen_catalogue.py` | Builds `hero_catalogue.html` from the inputs below. Run from this directory: `python3 gen_catalogue.py`. |
| `catalogue_template.html` | Static shell (CSS, header, footer, decision-tracking JS) — `gen_catalogue.py` fills in the per-layer card grid and the Known Issues table. |
| `franchise_financials.json` | Sales_2025 / Units_2025 / GM%_2025 / Sales_YTD2026 / Units_YTD2026 / GM%_YTD2026 for the 34 franchises, extracted from `Project_Bob_Portfolio_Tiering_08092026.xlsx` (Sheet 2, keyed `"{Base}|{Gender}"`). Re-extract this if the workbook is refreshed — there's no script for that extraction yet (a good next `update_*`-style script if this catalogue becomes a recurring build rather than a one-off test). |
| `auto_crop.py`, `final_crop.py` | The two-pass image-cropping pipeline used to turn raw full-page screenshots into uniform product-only photos (detects the studio-backdrop band via row-wise edge-brightness sampling, then a proportional bottom trim to clear sticky page overlays). Kept for reference/reuse if this set is ever re-shot or extended. |
| `catalogue_final/` | The 34 final, cropped, uniform-size (900×1100) product-only JPEGs — one per franchise, filenamed by the `img` key used in `gen_catalogue.py`'s `PRODUCTS` list. |
| `catalogue_images_b64.json` | Same 34 images, base64-encoded (`{image_key: base64_jpeg_string}`) — what `gen_catalogue.py` actually loads and embeds as data URIs (the published artifact can't reference external image files). Regenerate from `catalogue_final/` with `python3 encode_images.py` if the photo set changes. |

## Image provenance — internal test use only

The 34 photos in `catalogue_final/`/`catalogue_images_b64.json` are
screenshots of live haglofs.com and third-party retailer product pages,
cropped to product-only — **not licensed imagery**. Committed per
business direction (9 Sep 2026) on the basis that this stays **internal
test use**, not something republished or shared externally as-is. If
this catalogue ever moves beyond internal review (e.g. shared outside
the company, or the mechanic gets scaled to the full portfolio), source
proper licensed/owned photography first rather than carrying these
forward.

## Known caveats (see ASSUMPTIONS_REGISTER.md REG-020 for the Units figures)

- 7 of the 34 photographed franchises are flagged in the page's own
  "Known issues" table: 6 show a "II"/successor generation rather than
  the exact SKU the sales figure is attributed to, 1 has an uncertain
  size/variant match (URL slug vs. page title disagree).
- FY25 and YTD2026 Sales/GM% are the published, audited Sheet-2 basis.
  FY25 and YTD2026 Units are both real (REG-020 for FY25), but
  raw-recomputed from the source exports rather than an audited
  published total — same caveat as Tier Bench's `units25` field.
- Keep/Merge/Update decisions are stored in each viewer's own browser
  (`localStorage`), not shared or synced anywhere — a UI test, not a
  real review record yet.
