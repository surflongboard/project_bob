# Project Bob — Key-Article Margin Analysis (handoff package)

> **Note:** this repo also holds a second, separate workstream —
> `data/ss26_portfolio_tiering/` (8-tier SS26 DTC & Wholesale portfolio
> analysis, its own assumptions register). Same business, different
> question and different source pull; see that folder's
> `ASSUMPTIONS_REGISTER.md` before assuming a number or account rule from
> one workstream applies to the other — a few (e.g. how the special
> accounts XXL/Zalando/China/Stadium are treated) are intentionally
> different between the two.

This is a working starting point, not a finished tool. It packages up the
pipeline and assumptions from the first 8 layers analyzed in chat (Shell,
Insulation, Mid layer, Legwear, Soft shell, Tops, Daypacks, Accessories) so
the remaining 8 layers — and future data drops (FY2026 YTD) — don't require
re-deriving the same logic from scratch.

## Files

- **`config.py`** — every analytical assumption made so far, in one place,
  with reasoning in the comments. Read this first. This is the file to edit
  when a judgment call needs revisiting (e.g. the NEW/RAMP unit threshold,
  the XXL/Zalando/China account notes, the version-family matching logic).
- **`load_data.py`** — loads and cleans the two source workbooks into one
  tidy DataFrame. Run `python load_data.py` after dropping new source files
  into `data/` — it prints a sanity check (row counts, total sales by year,
  layer list) so you can eyeball that nothing broke before trusting anything
  downstream.
- **`analysis.py`** — the reusable functions: `article_summary`,
  `account_view`, `channel_view`, `ordertype_view`, `reorder_channel_split`,
  `find_version_families`. Each takes a layer-filtered DataFrame and returns
  a plain pandas DataFrame. No plotting or file-writing in here on purpose —
  compose these into whatever output format you need (notebook, script,
  xlsx export).
- **`data/`** — put `bob_salesdata_2024.xlsx` / `bob_salesdata_2025.xlsx`
  (or their FY2026 successors) here.
- **`output/`** — `load_data.py` writes a cached `full.pkl` here so you don't
  re-parse the Excel files every run.

## Quick start

```python
from load_data import load_all
from analysis import article_summary, account_view, channel_view, ordertype_view, reorder_channel_split

full = load_all("data/bob_salesdata_2024.xlsx", "data/bob_salesdata_2025.xlsx")

layer = full[full["Layer"] == "Sleepingbags"]   # pick any of config.ALL_LAYERS
print(article_summary(layer).head(15))
print(account_view(layer))
print(channel_view(layer))
print(ordertype_view(layer))
print(reorder_channel_split(layer))
```

## What's already been found (see `config.CROSS_LAYER_FINDINGS_LOG` for the
maintained version of this)

1. **Wholesale Reorder cost inflation** — COGS/unit rising 12–44% in
   Wholesale Reorder while E-com Reorder stays flat/improving. Confirmed in
   7 of 8 layers checked. Daypacks (Hardware) was the one exception — working
   theory is this is an apparel-specific problem (small cut-and-sew reorder
   batches, rush freight), not universal. **Next step: test this against the
   footwear/hardware layers still to come** (Hiking shoes, Trekking shoes,
   Trail running shoes, Bags, Sleepingbags, Hiking packs, Trekking packs) —
   if the pattern holds (mild in hardware/footwear, sharp in apparel), that's
   a much more targeted finding for Sourcing than "fix Reorder everywhere."
2. **FW27 target-margin gaps** — top articles in every layer checked are
   running 10–30pt below their own FW27 target margin, regardless of whether
   the layer's overall trend is up or down.
3. **`Active = FALSE` anomaly** — recurring in the Assortment Attribution
   Review on FW27 successor styles that are clearly still trading (Bield
   Down II, Zircon Slim II, Gabbro II). Worth a direct question to Product.
4. **Version-transition lifecycle** — successions split into CLEAN CUTOVER /
   PARTIAL OVERLAP / HEAVY OVERLAP / GAP YEAR (a new pattern found in
   Gabbro Pant — old style wound down before the successor has any sales).
   Almost no evidence anywhere of deliberate markdown pricing on outgoing
   versions.
5. **XXL is layer-inconsistent** — exiting in Shell/Insulation/Legwear/
   Accessories, growing in Mid layer/Soft shell. Not yet reconciled — worth
   a direct question to Sales on whether this is intentional category
   consolidation.

## Known gaps / things to watch

- The Assortment Attribution Review (FW27 target margins, carry-over status)
  is a PDF, not structured data. Target-margin figures used in this analysis
  were pulled by grep/manual lookup against `pdftotext -layout` output, not a
  clean programmatic join — if this becomes a recurring need, it's worth
  asking whoever owns that file for a proper export (CSV/Sheet) instead.
- `find_version_families` is a string-matching heuristic (strips "Men"/
  "Women"/"II"/"III"/"2.0" and groups by what's left). It will miss any
  transition that involves a real name change rather than a version-token
  bump. A REF_CODE/franchise-based join against the assortment file would be
  more robust if that field turns out to be populated reliably.
- FY2026 YTD data was not available as of this handoff — everything above is
  FY2024 vs FY2025 only.
