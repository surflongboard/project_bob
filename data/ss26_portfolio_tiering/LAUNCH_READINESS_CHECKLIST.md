# SS26 Portfolio Tiering — Launch Readiness Checklist

Tracked list of what to close out before presenting this work (the tiering
workbook, the assumptions register, and Tier Bench) to the wider team. Pulled
together 8 Sep 2026 after a review of the whole workstream — see that
conversation for the full reasoning behind each item.

**How to use this:** check items off as they're resolved, update the "Status"
column, and note who owns what. Re-run the QA pass (Section C) after any
change to Tier Bench's underlying data before showing it to anyone new.

---

## A. Open assumptions — resolve or explicitly own before the meeting

A skeptical team will find these if you don't name them first. For each:
either resolve it, or walk in with a stated owner and next step.

| ID | Item | Status | Owner | Next step |
|---|---|---|---|---|
| REG-002 | Main Segment taxonomy inconsistent across exports (OUT/OUTDOOR, SNW/SNOW, etc.) | Open | Merchandising (TBD) | Confirm canonical list; likely moot for this deliverable (tiering doesn't use Main Segment) but worth closing so it doesn't linger |
| REG-005 | "SS26 & SS25" flag column meaning undocumented | Open | TBD | Ask whoever built the original export what this column is for |
| REG-007 | Sales Market / store-type granularity differs between export files | Open | Data & Analytics | Define one parsing rule + market-to-region roll-up if this granularity is ever needed downstream |
| REG-009 | FY25 total sales gap (~2.3%) between SS26 exports and `bob_salesdata_2025.xlsx` | Open | Data & Analytics | Explain to Finance/whoever owns both pulls — likely pull-date drift, not confirmed |
| REG-011 | Case-sensitive duplicate franchise names (e.g. "Roc Sight SoftshellJacket" vs "ROC Sight Softshell Jacket") | Open | Data & Analytics | Check with data owner whether this is an isolated case or systemic |
| REG-012 | `Ordertype = "3-Close out order"` treatment in Core-scope Sales_2025 unclear (13.5% of all FY25 sales) | Open | Data & Analytics | Needed before any future SEK-level reconstruction from raw data is attempted; currently blocks a more precise Wholesale Share % and Generation Detail sheet |

---

## B. Process / governance

- [ ] **Get PR #2 reviewed and merged** (or make a deliberate call to keep iterating on the branch until after team review). Right now every change this session is unreviewed — "the team verifying assumptions" should mean verifying *approved* ones, not draft ones.
- [ ] **Decide sharing scope** for the repo and both artifacts (Assumptions Register, Tier Bench) before sending links to the wider team — both are private by default.
- [ ] **Decide on an approval workflow** for future assumption changes — right now status (Confirmed/Provisional/Open) is set in real time during a chat session. Does a Confirmed status need a second sign-off once more people depend on it?

---

## C. Tier Bench pre-launch checklist

- [ ] **Manually test the live chat end-to-end** — the `sample()` capability powering it has never been verified to actually fire in this session; test before anyone else sees it.
- [ ] **Confirm the org has the `sample` capability enabled for artifacts** — check with your Claude admin. If it's not enabled, the first team member to try it hits a wall.
- [ ] **Add export/download capability** — you asked that the team still be able to export data; Tier Bench currently answers in chat only, no file download yet.
- [ ] **Run a QA pass** — 5–10 questions with known answers pulled directly from the workbook, confirm Tier Bench gets them right, before the team's first look. Re-run after any data resync.
- [ ] **Re-confirm sharing settings** on the Tier Bench artifact itself once ready to distribute.

---

## D. Data pipeline / reusability

- [x] **Commit reusable matching/parsing scripts to the repo**, not just their outputs. *Done 8 Sep 2026 — `scripts/` now holds a shared library (`ss26_lib.py`) plus one `update_*.py` per REG-010/014–019 column/sheet, following the `load_data.py`/`analysis.py` pattern. Each script was validated by running it against a scratch copy of the published workbook and diffing every touched cell against the real file — all match exactly (float-rounding noise aside). Two intentional simplifications in `update_stock.py` (top-10-only detail tables, generic callout text) are documented in `scripts/README.md`. **Known gap:** the base 8→7-tier classification itself (cols A–M) predates this session's scratch work and still has no reusable script — see `scripts/README.md`'s "Not covered" note.*
- [x] **Establish a file-naming/intake convention** for manual uploads. *Done 8 Sep 2026 — raw source files (clearance lists, core assortment list, stock snapshot) copied from the chat upload into `inputs/<category>/`, keeping original filenames; convention documented in `scripts/README.md`.*

---

## E. Documentation / continuity

- [ ] **Add a "Recent Changes" section** to the top of the Assumptions Register — last 5 updates in plain language, so a first-time reader sees the pace of iteration without reading all 19 entries.
- [ ] **Consider a `CLAUDE.md`** summarizing conventions (matching methodology, file locations, register discipline) so a future Claude Code session — or a teammate — ramps up faster than re-reading the whole register.
- [ ] **Confirm nothing load-bearing only exists in this chat transcript** — do a pass to make sure every material decision is written into the register, not just explained once in conversation.

---

*Compiled 8 Sep 2026. Cross-reference: `ASSUMPTIONS_REGISTER.md` for the full data assumptions this checklist points at.*
