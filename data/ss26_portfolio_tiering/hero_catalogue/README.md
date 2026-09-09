# Hero Tier Catalogue — generator

A visual-review test build: one photo per Hero/Near-Hero franchise (34
total), grouped by layer, with FY25/YTD2026 Sales/Units/GM% per card and a
Keep/Merge/Update decision control — an alternative to reviewing the same
34 franchises as a text list. Published as an artifact (see
`ASSUMPTIONS_REGISTER.md`'s live-links section for the current URL).

**This was a chat-only build for several iterations before being
committed here** — flagged as a known gap in `LAUNCH_READINESS_CHECKLIST.md`
Section E and now closed for the code/data half of it (see "What's
intentionally NOT committed" below for the other half).

## What's here

| File | What it does |
|---|---|
| `gen_catalogue.py` | Builds `hero_catalogue.html` from the three inputs below. Run from this directory: `python3 gen_catalogue.py`. |
| `catalogue_template.html` | Static shell (CSS, header, footer, decision-tracking JS) — `gen_catalogue.py` fills in the per-layer card grid and the Known Issues table. |
| `franchise_financials.json` | Sales_2025 / Units_2025 / GM%_2025 / Sales_YTD2026 / Units_YTD2026 / GM%_YTD2026 for the 34 franchises, extracted from `Project_Bob_Portfolio_Tiering_08092026.xlsx` (Sheet 2, keyed `"{Base}|{Gender}"`). Re-extract this if the workbook is refreshed — there's no script for that extraction yet (a good next `update_*`-style script if this catalogue becomes a recurring build rather than a one-off test). |
| `auto_crop.py`, `final_crop.py` | The two-pass image-cropping pipeline used to turn raw full-page screenshots into uniform product-only photos (detects the studio-backdrop band via row-wise edge-brightness sampling, then a proportional bottom trim to clear sticky page overlays). Kept for reference/reuse — see "What's intentionally NOT committed" for why they have nothing to run against here.

## What's intentionally NOT committed

The 34 product photos themselves — at every stage (raw screenshots,
cropped intermediates, the final base64-encoded JSON `gen_catalogue.py`
embeds into the page) — are **not** in this folder or the repo.

They're screenshots of live haglofs.com and third-party retailer product
pages, not licensed imagery (the published catalogue's own footer says as
much: "not something to publish externally as-is"). Putting them in git
history is a materially different, harder-to-reverse step than putting
them in a private, unlisted artifact — it means distributing someone
else's product photography via the repo. That's a call worth making
explicitly rather than defaulting into via a routine commit; ask before
adding them here.

To regenerate the page, `gen_catalogue.py` expects `catalogue_images_b64.json`
(`{image_key: base64_jpeg_string}`, one entry per `img` key in `PRODUCTS`)
in this directory — re-source and re-run the crop pipeline to produce it.

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
