# Project Bob — conventions for working in this repo

Read this first, before `README.md` or either workstream's own docs — it
tells you which of those to read next depending on what you're doing.

## Two separate workstreams, same business

- **Key-Article Margin Analysis** (repo root: `config.py`, `load_data.py`,
  `analysis.py`, `data/bob_salesdata_*.xlsx`, `data/bob_financialdata_*.xlsx`,
  `data/bob_ecomdiscountdata_*.xlsx`). FY2024 vs FY2025 margin/channel
  analysis by product layer. Start at the top-level `README.md`.
- **SS26 Portfolio Tiering** (`data/ss26_portfolio_tiering/`). An 1,860-
  franchise DTC & Wholesale tier classification, built from a separate
  `SS26_DTC_Wholesale_w34_data_*.xlsx` export. Start at
  `data/ss26_portfolio_tiering/ASSUMPTIONS_REGISTER.md`.

They analyze the same underlying business but from different source pulls,
and **a few rules are deliberately different between them** — most notably
how the "Core account" exclusions (XXL, China, Zalando, Zalando
Marketplace, Stadium Outlet) are applied, because the two source exports
encode account identity in different fields. Don't assume a rule from one
workstream applies to the other; check that workstream's own docs.

## If you're picking up a new data drop

1. Figure out which workstream it belongs to (source filename/sheet
   layout is usually enough — ask if not obvious).
2. For the SS26 Portfolio Tiering workstream: drop the raw file under
   `data/ss26_portfolio_tiering/inputs/<category>/`, keeping its original
   filename (see `scripts/README.md` for the intake convention), then run
   the matching `update_*.py` script rather than writing new one-off
   Python. If no script exists yet for this kind of update, write one
   following the existing `update_*.py` pattern and commit it — don't
   leave the logic only in a scratch file or a chat transcript.
3. For the Key-Article Margin Analysis workstream: drop the file in
   `data/`, extend `load_data.py`/`analysis.py`/`config.py` following
   their existing structure (load/clean separated from analyze;
   assumptions documented in docstrings/comments, not just in chat).
4. **Every material assumption, cleaning rule, or judgment call gets a
   register entry** in the relevant workstream's `ASSUMPTIONS_REGISTER.md`
   — not just an explanation in the chat session. If it's not written
   down there, it doesn't exist for the next person (or the next Claude
   Code session) picking this up. Cite entries by ID (e.g. "per REG-010")
   in code comments and script docstrings that depend on them.
5. The SS26 workstream's register also has a live, filterable HTML
   version (linked at the top of its `ASSUMPTIONS_REGISTER.md`) and a
   "Tier Bench" chat-query tool over the dataset — keep both in sync with
   the file whenever you touch a register entry or the underlying data.
   The checked-in `.md` file is the source of truth if the two disagree.

## Shared matching methodology (SS26 workstream)

Franchise identity is **Base + Gender**, derived from the raw `Article`
text by stripping gender tokens (`Men`/`Women`/`Unisex`/`Junior`) and
version tokens (`2.0`/`II`/`III`/`IV`), including glued-text export
artifacts (no space before the token). The canonical token sets live in
`config.py` (`GENDER_TOKENS`/`VERSION_TOKENS`) and are reused by BOTH
workstreams rather than redefined — don't fork a second copy of these
lists. The SS26 workstream's parsing functions live in
`data/ss26_portfolio_tiering/scripts/ss26_lib.py`
(`base_name`/`gender`/`version_token`/`franchise_key`) — import from
there rather than re-copying the regex.

A **50,000 SEK/year reliability threshold** (`MIN_SALES_FOR_RELIABLE_MARGIN_PCT`
in `config.py`, `MIN_RELIABLE` in `ss26_lib.py`) blanks/excludes GM%,
growth, and share calculations built on a small denominator — reused
everywhere in both workstreams. Don't compute a new margin/growth/share
metric without applying it.

## Register discipline

- **Confirmed** — a fact with one right answer, verified against a real
  source (a document, a cross-check, business confirmation). Safe to
  build on.
- **Provisional** — a reasonable methodology choice, not yet independently
  confirmed. Usable, but flag it if it feeds something high-stakes.
- **Open** — a genuine unresolved question. Don't silently resolve it by
  picking an assumption in code; either ask, or build around it and say
  so.

Status changes are currently made in real time by whoever is driving the
session — see `data/ss26_portfolio_tiering/LAUNCH_READINESS_CHECKLIST.md`
Section B for the current (deliberately informal, revisit-later) approval
model.

## Before presenting anything from this repo to the wider team

Check `data/ss26_portfolio_tiering/LAUNCH_READINESS_CHECKLIST.md` first —
it tracks what's still open (unresolved register items, Tier Bench QA,
sharing scope) before either workstream is ready for an audience beyond
the person driving it today.
