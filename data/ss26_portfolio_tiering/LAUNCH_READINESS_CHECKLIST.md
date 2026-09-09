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

- [x] **Run a QA pass.** *Done 8 Sep 2026 — but not by asking the live chat: `sample()` can't be driven from this session (it needs a real authenticated browser tab, which this agent doesn't have). Instead, extracted Tier Bench's actual query-tool code (the JS `execute()` functions each tool call runs) into a standalone Node harness and ran 12 known-answer checks straight against it — the same code path `sample()`'s tool-calling invokes, just without the live model round-trip. Every check matched the register/workbook exactly: 20 Hero+Near-Hero franchises missing from Core Assortment FW27 (REG-015), 48 franchises "Fully" clearance (REG-010), 109/221 generation-comparison rows (REG-014), 59/83 DTC-vs-Wholesale-expansion counts (REG-018), the 649.76M SEK tier-summary TOTAL (REG-013), dead-stock top-5 matching Sheet 4 exactly (REG-019).**
  **Found and fixed a real bug in the process:** `queryFranchises`, `queryGenerationComparison`, `queryChannelMix`, and `queryStock` were passing their whole tool input — including the `limit`, `sortBy`, and `sortDir` control fields — straight into the row-filter function, which treated every key as a literal field match. Since no data row has a field called `limit`, ANY tool call that specified a limit, a sort field, or a sort direction returned zero rows, silently. That's not an edge case — it's most real questions (a page description said "Default 30" for a reason; "top N" / "sorted by" questions are exactly this tool's core use case). Confirmed the bug, fixed it (control keys now excluded from the filter pass before it runs), re-ran all 12 checks clean. Also fixed a smaller issue while in there: the system prompt said "four tools" when there are five. Republished with the fix.*
- [x] **Add export/download capability.** *Done 8 Sep 2026 — added the `downloads` capability. Every answer that used a query tool now gets an "Export as CSV" button per tool result underneath it, so the team can pull the underlying rows into Excel rather than only reading the chat answer.*
- [ ] **Manually test the live chat end-to-end.** Not done — needs a real click from an authenticated browser tab, which isn't something this session can do. With the bug above now fixed and the query layer independently verified correct, the main known risk going into that test is gone — but do one real round-trip (ask it something, confirm an answer comes back and an Export button appears) before showing anyone else.
- [ ] **Confirm the org has the `sample` capability enabled for artifacts** — check with your Claude admin; can't be verified from here.
- [x] **Re-confirm sharing settings.** *Already private (owner-only) per the Section B decision — confirmed on this republish ("sharing owner"). Revisit before sending the link to the wider team.*

---

## D. Data pipeline / reusability

- [x] **Commit reusable matching/parsing scripts to the repo**, not just their outputs. *Done 8 Sep 2026 — `scripts/` now holds a shared library (`ss26_lib.py`) plus one `update_*.py` per REG-010/014–019 column/sheet, following the `load_data.py`/`analysis.py` pattern. Each script was validated by running it against a scratch copy of the published workbook and diffing every touched cell against the real file — all match exactly (float-rounding noise aside). Two intentional simplifications in `update_stock.py` (top-10-only detail tables, generic callout text) are documented in `scripts/README.md`. **Known gap:** the base 8→7-tier classification itself (cols A–M) predates this session's scratch work and still has no reusable script — see `scripts/README.md`'s "Not covered" note.*
- [x] **Establish a file-naming/intake convention** for manual uploads. *Done 8 Sep 2026 — raw source files (clearance lists, core assortment list, stock snapshot) copied from the chat upload into `inputs/<category>/`, keeping original filenames; convention documented in `scripts/README.md`.*

---

## E. Documentation / continuity

- [x] **Add a "Recent Changes" section** to the top of the Assumptions Register. *Done 8 Sep 2026 — added to both the checked-in file and the live artifact (last 5 updates, plain language).*
- [x] **Add a `CLAUDE.md`.** *Done 8 Sep 2026 — repo root, covers the two-workstream split, where a new data drop goes, the shared matching methodology, register discipline (Confirmed/Provisional/Open), and a pointer back to this checklist.*
- [x] **Confirm nothing load-bearing only exists in this chat transcript.** *Audited 8 Sep 2026. Found and fixed one gap: the "Wholesale expansion opportunity" glossary mapping existed only in Tier Bench's system prompt, not in the register — added to both. Code-level judgment calls (News package fallback list, stock-file column-index assumption, why scripts don't re-sort the sheet, the top-10 cap on stock detail tables) were already captured as docstrings/comments in `scripts/`, which is the right home for implementation detail rather than the register.*
  **Partly closed 9 Sep 2026, partly still open:** Tier Bench's page generator (`gen_page.py`, the script that turns `qa_data/*.json` into the Tier Bench page) is now committed at `data/ss26_portfolio_tiering/tier_bench/`, along with the seven `qa_data/*.json` files it reads and a partial rebuild script (`build_qa_data.py`, 4 of 7 files). **Still genuinely open:** `build_qa_data.py` itself depends on an uncommitted intermediate (`art_channel.pkl`) whose own builder was never saved, and 3 of the 7 `qa_data/*.json` files (stock, country-breakdown, regional-tiering) have no committed regeneration path at all — their extraction logic only ever existed as one-off chat scripts, now gone. So refreshing Tier Bench's dataset after a future data drop is not yet a one-command rerun the way the Excel pipeline (Section D) is — see `tier_bench/README.md` for the exact gap. Raised as a candidate next item if Tier Bench moves from pilot to something refreshed on a cadence.

---

*Compiled 8 Sep 2026. Cross-reference: `ASSUMPTIONS_REGISTER.md` for the full data assumptions this checklist points at.*
