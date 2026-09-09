# Tier Bench — generator

The chat-query tool over the current tiering dataset — seven tools
(franchise, generation-comparison, channel-mix, tier-summary, stock,
country-breakdown, regional-tiering) that a viewer can ask questions
against via `sample()`, plus a filterable table view with CSV export per
tool. Published as an artifact (see `ASSUMPTIONS_REGISTER.md`'s live-links
section for the current URL).

**This was a chat-only build across many iterations before being
committed here** — flagged as a known gap in
`LAUNCH_READINESS_CHECKLIST.md` Section E and now partly closed (the page
generator itself; see "What's still not reproducible" below for what
isn't).

## What's here

| File | What it does |
|---|---|
| `gen_page.py` | Builds `tier_bench.html` from `qa_data/*.json`. Run from this directory: `python3 gen_page.py`. Self-contained — the full page markup, CSS, and client-side JS (tool definitions, `SYSTEM_PROMPT`, the `sample()` call, the filter/CSV-export table view) live in one big string in this file, unlike the Hero Tier Catalogue's separate `catalogue_template.html`. |
| `qa_data/*.json` | The seven datasets the page's tools query, as a point-in-time extract from `Project_Bob_Portfolio_Tiering_08092026.xlsx` (`franchises_chunk_0..4.json` — Sheet 2, chunked to keep each JSON doc a manageable size; `generation_comparison.json` — Sheet 3, recomputed per-pair; `channel_mix.json` — per-article Wholesale/Retail/E-com split; `tier_summary.json` — Sheet 1's 7-tier roll-up; `stock_by_franchise.json` — Sheet 4; `country_breakdown.json` — Sheet 5; `regional_tiering.json` — Sheets 6/7, REG-023/024). |
| `build_art_channel.py` | Rebuilds `art_channel.pkl` — the per-article FY25 Wholesale/Retail/E-com Sales+GM lookup `build_qa_data.py` needs — from the raw SS26 exports' own `Sales Channel` field (Core-scope filtered, REG-008). Not committed itself, same as its own output (`art_channel.pkl`) and `tier_bench.html` — a build intermediate, regenerate on demand. |
| `build_qa_data.py` | Rebuilds `franchises_chunk_*.json`, `tier_summary.json`, `channel_mix.json`, and `generation_comparison.json` from the master workbook — needs `art_channel.pkl` (run `build_art_channel.py` first). |
| `extract_stock.py` | Rebuilds `stock_by_franchise.json` from the raw warehouse stock snapshot, reusing `scripts/update_stock.py`'s exact matching logic (REG-019). Verified 9-Sep-2026: exact match against the previously-published file, structurally and in row order. |
| `extract_country_breakdown.py` | Rebuilds `country_breakdown.json` by reading the master workbook's Sheet 5 directly (REG-022) — doesn't re-derive the franchise↔country matching itself; re-run `scripts/update_country_breakdown.py` first if Sheet 5 needs refreshing. Verified 9-Sep-2026: exact match, all 13,438 rows. |
| `extract_regional_tiering.py` | Rebuilds `regional_tiering.json` by reading the master workbook's Sheets 6/7 directly (REG-023/024), deduped to one row per franchise per region and joined to Sheet 2's Tier for `globalTier`. Re-run `scripts/update_regional_tiering.py` first if the regional sheets need refreshing. Verified 9-Sep-2026: exact match, all 3,720 rows. |

## Regenerating the full dataset

As of 9-Sep-2026, all 7 `qa_data/*.json` files have a committed,
verified regeneration path from the master workbook + raw source files
— matching the SS26 pipeline scripts' own standard (`scripts/update_*.py`,
each with a clean, rerunnable path from source file to workbook column).
In order:

```bash
cd data/ss26_portfolio_tiering/tier_bench
python3 build_art_channel.py         # -> art_channel.pkl (not committed, a build intermediate)
python3 build_qa_data.py             # -> franchises_chunk_*.json, tier_summary.json,
                                      #    channel_mix.json, generation_comparison.json
python3 extract_stock.py             # -> stock_by_franchise.json
python3 extract_country_breakdown.py # -> country_breakdown.json
python3 extract_regional_tiering.py  # -> regional_tiering.json
python3 gen_page.py                  # -> tier_bench.html
```

`build_art_channel.py`, `extract_stock.py`, `extract_country_breakdown.py`,
and `extract_regional_tiering.py` read straight from raw source files or
the master workbook and don't depend on each other. If the underlying
workbook sheets themselves need refreshing first (a new data drop, not
just a Tier Bench rebuild), re-run the relevant `scripts/update_*.py`
before the `extract_*.py` that reads its sheet — see `scripts/README.md`'s
ordering (`update_stock.py`/`update_country_breakdown.py`/
`update_regional_tiering.py` before Tier Bench touches those sheets at
all).

**Every one of the 7 files was verified structurally equal (not just
matching row counts) against the previously-published version before
this was trusted** — `build_art_channel.py`'s rebuild reproduced
`channel_mix.json`/`generation_comparison.json`/`tier_summary.json`
exactly on the first run; `extract_country_breakdown.py` and
`extract_regional_tiering.py` each needed one fix (a genuinely-blank
sheet cell was defaulting to `0` instead of staying `None`); and
re-verifying `franchises_chunk_*.json` end-to-end (previously untested,
since `build_qa_data.py` used to fail before reaching that output) caught
a real, separate bug — `build_qa_data.py`'s franchise loader predated
REG-015 through REG-021 being added to Sheet 2 and was silently missing
7 fields (`coreAssortmentFw27`, `fw27Collection`, `wholesaleSharePct`,
`channelPattern2025`, `units25`, `unitsYtd26`, `styleCodes`) that the
*previously-committed* `qa_data/franchises_chunk_*.json` files already
had. Fixed by reading all 22 of Sheet 2's current columns rather than
the 15 the script originally knew about.

## Running the generator

```bash
cd data/ss26_portfolio_tiering/tier_bench
python3 gen_page.py
```

Writes `tier_bench.html` next to this README — not committed (a build
output, like `hero_catalogue.html`, not source). Validate before
publishing: `node --check` on the extracted `<script>` block, and
independently exercise the seven tools' `execute()` functions against
realistic query parameters (a Node harness reading the same block) before
trusting a change to a tool's logic or description — `sample()` can't be
driven from an agent session directly (see
`LAUNCH_READINESS_CHECKLIST.md` Section C for why), so this is the
closest available substitute for a live QA pass.

**One easy way to break every query at once, found the hard way
(9-Sep-2026):** the `sample()` API enforces a **1KB (1024-byte, UTF-8)
limit on each tool's `description` field**. Blowing it doesn't just
disable that one tool — it fails the *entire* tool-calling request, so
every question breaks, including ones that never touch the oversized
tool, with no warning beforehand. Check `Buffer.byteLength(desc, 'utf8')`
for every tool description before publishing a change to one.

## Known caveats

- Franchise-level (`queryFranchises`) figures are the audited, published
  Sheet-2 basis. Generation-, channel-, country-, and regional-tiering-
  level figures are raw-recomputed and directional (REG-012) — Tier
  Bench's own system prompt says so when a question turns on an exact
  SEK figure.
- `queryRegionalTiering`'s `regionalTier` and `globalTier` are meant to
  disagree sometimes — region-relative thresholds mean a franchise can
  legitimately rank Hero in one region and Workhorse+Harvest in the
  other (REG-023). Not a bug if you see it.
- See `ASSUMPTIONS_REGISTER.md` for every dataset's own methodology and
  caveats — this README only covers how the page itself is built.
