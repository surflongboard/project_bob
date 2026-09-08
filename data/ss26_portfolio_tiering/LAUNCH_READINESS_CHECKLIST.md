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

- [x] **Get PR #2 reviewed and merged.** *Done 8 Sep 2026 — merged directly (no CI configured on this repo, no reviewers assigned yet pre-team-launch) into `main` at `5d1e8c1`. All of Section D's new pipeline scripts and inputs are now on `main`, not just a branch. **Carries forward:** this was a deliberate call to skip formal review for now, not a standing policy — the next PR onto `main` should go through actual team review once there's a team to review it, per this section's own framing ("the team verifying assumptions" should mean verifying *approved* ones).*
- [x] **Decide sharing scope** for the repo and both artifacts. *Decided 8 Sep 2026 — staying private (just the owner) for now; owner is reviewing the full workstream tomorrow before deciding who else gets access. Revisit before any link goes to the wider team.*
- [x] **Decide on an approval workflow** for future assumption changes. *Decided 8 Sep 2026 — keeping the current model (single owner sets Confirmed/Provisional/Open status in real time) for now. Revisit once more people are actually depending on the register — a second sign-off requirement is the natural next step at that point, not before.*

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

- [x] **Add a "Recent Changes" section** to the top of the Assumptions Register. *Done 8 Sep 2026 — added to both the checked-in file and the live artifact (last 5 updates, plain language).*
- [x] **Add a `CLAUDE.md`.** *Done 8 Sep 2026 — repo root, covers the two-workstream split, where a new data drop goes, the shared matching methodology, register discipline (Confirmed/Provisional/Open), and a pointer back to this checklist.*
- [x] **Confirm nothing load-bearing only exists in this chat transcript.** *Audited 8 Sep 2026. Found and fixed one gap: the "Wholesale expansion opportunity" glossary mapping existed only in Tier Bench's system prompt, not in the register — added to both. Code-level judgment calls (News package fallback list, stock-file column-index assumption, why scripts don't re-sort the sheet, the top-10 cap on stock detail tables) were already captured as docstrings/comments in `scripts/`, which is the right home for implementation detail rather than the register.*
  **Found and NOT fixed — flagged instead, since it's a real build, not a doc fix:** Tier Bench's own data-refresh pipeline (the script that turns the workbook into the JSON embedded in the Tier Bench page) exists only as a one-off script in this session's scratchpad — never committed. Right now, refreshing Tier Bench's dataset after a future data drop would mean re-deriving that export logic from scratch, the exact problem Section D solved for the Excel workbook itself. Not addressed here — raised as a candidate next item if Tier Bench moves from pilot to something refreshed on a cadence.

---

*Compiled 8 Sep 2026. Cross-reference: `ASSUMPTIONS_REGISTER.md` for the full data assumptions this checklist points at.*
