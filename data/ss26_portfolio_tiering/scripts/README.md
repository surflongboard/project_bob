# SS26 Portfolio Tiering — pipeline scripts

Reusable, committed versions of the matching/parsing logic used to build
`Project_Bob_Portfolio_Tiering_08092026.xlsx`. Every column this session
added to the published workbook after the base 1,860-row tiering existed
now has a corresponding `update_*.py` script here, so a future data drop
is "rerun this script," not "re-derive this from scratch."

**The base 8→7-tier classification itself** (columns A–M — Tier, Sales_2025,
GM%, Pace%, Growth% etc.) predates this session's scratch-script work and
was never captured as a reusable script here — but as of 9-Sep-2026 the
actual rule behind it **is** documented and validated: see REG-024 in the
register, `ss26_lib.classify_tier()`, and `verify_base_tiering.py` (a
read-only audit — 99.8% match against the published Tier column, but it
does NOT rewrite Sheet 2; the published tiers stay authoritative). Nothing
here rebuilds Sheet 2's Tier column from scratch and overwrites it — that
remains true, and is a deliberate choice (see `verify_base_tiering.py`'s
docstring for why), not an oversight.

## What each script does

| Script | Register ID | Writes | Source input |
|---|---|---|---|
| `update_clearance.py` | REG-010 | Sheet 2 cols D, E, N, O, P | `inputs/clearance/*.xlsx` |
| `update_generation_detail.py` | REG-014 | Sheet 3 (full rebuild) | SS26 exports + `inputs/clearance/*.xlsx` |
| `update_core_assortment.py` | REG-015 | Sheet 2 col Q | `inputs/core_assortment/*.xlsx` |
| `update_fw27_collection.py` | REG-016 | Sheet 2 col R | `data/Assortment Attribution Review(F27).xlsx` |
| `update_wholesale_share.py` | REG-017 | Sheet 2 col S | SS26 exports (FY25 only) |
| `update_channel_pattern.py` | REG-018 | Sheet 2 col T | `data/bob_salesdata_2024/2025.xlsx` |
| `update_stock.py` | REG-019 | Sheet 4 (full rebuild) | `inputs/stock/*.xlsx` |
| `update_units_2025.py` | REG-020 | Sheet 2 col U | SS26 exports (FY25 only) |
| `update_style_codes.py` | REG-021 | Sheet 2 col V | SS26 exports (all three) |
| `update_country_breakdown.py` | REG-022 | Sheet 5 (full rebuild) | SS26 exports (all three) |
| `update_regional_tiering.py` | REG-023 | Sheets 6 & 7 (full rebuild) | `bob_salesdata_2024/2025.xlsx` + Sheet 5 |
| `verify_base_tiering.py` | REG-024 | Nothing — read-only audit, prints a match report | `bob_salesdata_2024/2025.xlsx` |
| `build_distribution_workbook.py` | — (packaging, no new methodology) | A brand-new dated file, `Project_Bob_Portfolio_Tiering_<DDMMYYYY>.xlsx` — never touches the master workbook | The master workbook itself |
| `build_browse_view.py` | — (no data, just a layout view) | Sheet 2b (full rebuild) | Sheet 2 (formulas only, no source file) |

`ss26_lib.py` is the shared library every script above imports from —
franchise-key parsing (`base_name`/`gender`/`version_token`), the two
Core-scope filters (REG-008), `MIN_RELIABLE`, the workbook path, and
common cell styling. Read its module docstring before touching any
script's matching logic — the two Core-scope filters are NOT
interchangeable across file types.

## Running a script

From the repo root:

```bash
cd data/ss26_portfolio_tiering/scripts
python3 update_clearance.py               # uses default input paths below
python3 update_clearance.py --jv path/to/new_jv_file.xlsx --zalando path/to/new_zalando_file.xlsx
```

Each script edits `Project_Bob_Portfolio_Tiering_08092026.xlsx` **in
place** and prints a summary (counts, match rates, any exceptions worth a
human look) — read the printed output before trusting the save. None of
them re-sort the sheet or touch columns they don't own, so scripts can be
re-run in any order, any number of times, except:

- Run `update_clearance.py` before `update_generation_detail.py` or
  `update_stock.py` if the clearance lists changed — both read Sheet 2's
  Tier column, and `update_clearance.py` is what updates it.
- `update_core_assortment.py`, `update_fw27_collection.py`,
  `update_wholesale_share.py`, `update_channel_pattern.py`,
  `update_units_2025.py`, `update_style_codes.py`,
  `update_country_breakdown.py` are independent of each other and of
  clearance.
- Run `update_country_breakdown.py` (Sheet 5) BEFORE
  `update_regional_tiering.py` (Sheets 6 & 7) — the latter reads Sheet
  5's country rows directly and exits with an error if it's missing.
  Also re-run `update_clearance.py` first if the clearance lists
  changed, since `update_regional_tiering.py` reads Sheet 2's Clearance
  Flag column too.
- Run `build_browse_view.py` last (or any time after) if you've
  re-run any Sheet 2 script — it just re-points formulas at Sheet 2's
  current column positions and has no data of its own.
- `verify_base_tiering.py` is read-only (no dependency, no write) — run
  it any time to re-confirm the REG-024 rule still matches Sheet 2's
  published tiers, e.g. after a new data drop, before trusting the rule
  for something new.

Validated (Sep 2026) by running every script against a scratch copy of
the published workbook and diffing the result cell-by-cell against the
real file: all match exactly except (a) floating-point noise at the
10th+ significant digit from summation order, and (b) two intentional
simplifications in `update_stock.py` — its per-tier detail tables cap at
the top 10 franchises by stock units (the original one-time build listed
all matching franchises, e.g. 124 for Thin/Immaterial, which doesn't
scale as a rerunnable sheet), and its callout text is written generically
rather than hardcoding specific example article names — re-read the
printed summary for anything an original hand-written callout used to
surface.

## Input files — intake convention

Manual data drops land under `inputs/<category>/`, keeping their
original filename (so it stays traceable to whatever the sender called
it) rather than being renamed on arrival:

```
inputs/
  clearance/        close-out / clearance lists (China JV, Zalando, ...)
  core_assortment/  FW27 Core Assortment style lists
  stock/            warehouse available-stock snapshots
```

For a **new** data drop in an existing category (e.g. the next quarterly
stock snapshot), add the new file alongside the old one rather than
overwriting it — keep the old file for traceability — and pass its path
explicitly via the script's `--source`/`--jv`/`--zalando` flag rather
than renaming it to match the current default. Once a new file is the
one you want scripts to use by default, update the script's default path
constant and note the change in `ASSUMPTIONS_REGISTER.md`.

`data/Assortment Attribution Review(F27).xlsx` and
`data/bob_salesdata_2024/2025.xlsx` are shared with the other workstream
(Key-Article Margin Analysis) and stay at the top-level `data/` folder
rather than moving under `inputs/` — don't relocate them.

## Distribution cuts

`build_distribution_workbook.py` is different from every script above —
it doesn't add a column or sheet to the master workbook, it **produces a
separate, dated file** (`Project_Bob_Portfolio_Tiering_<DDMMYYYY>.xlsx`)
for sharing outside the working file: Global Full Portfolio + Global
Generation Detail + the two regional tiering sheets, with the internal
working-only sheets (Browse view, Stock by Tier, Sales by Country)
dropped, and Sheet 1 rebuilt to include a sheet index (labeling each
sheet Global vs. Regional) plus the exact Global and regional tier rule
definitions (REG-023/024) that were never written into Sheet 1 before.
Re-run it any time a fresh dated cut is needed — it always reads the
current master workbook and writes a new file named for today, never
overwriting the master or a previous cut.

## Adding a new column/script

Follow the pattern above: read source file(s) with `ss26_lib` helpers,
match via `franchise_key()`, write one clearly-named column with the
shared header styling from `ss26_lib.styles()`, print a match-rate/
sanity-check summary, save. Add a row to the table above and a new
REG-### entry in `ASSUMPTIONS_REGISTER.md` describing the source, the
matching approach, and any caveats.
