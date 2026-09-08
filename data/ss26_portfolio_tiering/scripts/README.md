# SS26 Portfolio Tiering — pipeline scripts

Reusable, committed versions of the matching/parsing logic used to build
`Project_Bob_Portfolio_Tiering_08092026.xlsx`. Every column this session
added to the published workbook after the base 1,860-row tiering existed
now has a corresponding `update_*.py` script here, so a future data drop
is "rerun this script," not "re-derive this from scratch."

**Not covered by these scripts:** the base 8→7-tier classification itself
(columns A–M — Tier, Sales_2025, GM%, Pace%, Growth% etc.) predates this
session's scratch-script work and wasn't captured as a reusable script
here. Everything below assumes that base workbook already exists and only
adds/refreshes columns N onward, plus Sheets 3 and 4. Rebuilding the base
tiering from raw sales data from scratch is a real gap — flagged, not
silently out of scope; see `LAUNCH_READINESS_CHECKLIST.md`.

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
  `update_wholesale_share.py`, `update_channel_pattern.py` are
  independent of each other and of clearance.

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

## Adding a new column/script

Follow the pattern above: read source file(s) with `ss26_lib` helpers,
match via `franchise_key()`, write one clearly-named column with the
shared header styling from `ss26_lib.styles()`, print a match-rate/
sanity-check summary, save. Add a row to the table above and a new
REG-### entry in `ASSUMPTIONS_REGISTER.md` describing the source, the
matching approach, and any caveats.
