# Range Review tiering guidance (Hero / Harvest / Emerging / Exit)

Proposed with the user 3-Nov-2026, then fully revised same day (RR-049
supersedes RR-048's first pass entirely — Core Assortment, generation
age, and declining-trend are no longer part of this logic at all).
Maps to CLAUDE.md's "layer-by-layer keep/consolidate/improve/kill
proposal" framing for this workstream. Implemented as
`compute_suggested_tier()` in `scripts/build_workshop_data.py` — this
document and that function are the same guidance, not two independent
copies; if one changes, so should the other.

**This is a suggestion engine, not a decision engine.** The Range
Review Board shows a "Suggested: <Tier> — <reason>" hint on any
undecided card these rules have an opinion on, but never writes to the
Tier dropdown itself — Tier stays a human call, captured in the
Board's own `decisions` database, exactly like every other decision
field (Last season to sell, Consolidate). Treat every suggestion here
as a starting point for the Nov 12 workshop discussion, not a verdict.

## Rules (first match wins)

1. **New FRANCHISE in FW27, SS28, or FW28** (`isNewInSeason` and
   `newKind == "franchise"`) → **Emerging**, unconditionally. **Does
   NOT fire for a new GENERATION of an already-established franchise**
   (`newKind == "generation"`, e.g. "Alert GTX II Jacket") — RR-054,
   fixed 4-Nov-2026 after the user flagged it: a generation transition
   isn't a new, unproven product the way a genuinely new franchise is,
   and the blended-by-base historical sales figure (RR-037) already
   rolls up the predecessor generation's sales on purpose, so a
   generation case falls through to the same sales-based rules below
   as any other franchise instead of being auto-suggested Emerging.
2. **New in 2026** — no recorded sales in 2024 or 2025 at all (both
   genuinely absent, not zero — "blank not zero") but has 2026 sales
   → **Emerging**. A sales-based newness signal, independent of rule
   1 — catches a franchise whose Product master release season isn't
   cleanly FW27/SS28/FW28 but whose real sales history clearly starts
   in 2026, a new-generation franchise included (RR-054: 2 of the 24
   new-generation cases still land here, e.g. "Alert GTX II Jacket" —
   not a bug, their blended sales history genuinely has zero matched
   2024 sales and only core-excluded China sales in 2025, so from this
   workstream's own 2024–2026 data their real market history starts in
   2026 regardless of generation labelling; flagged to the user as a
   data-window limitation, not resolved further without their input).
3. **2026 units > 10,000 AND 2026 GP% > 50%** → **Hero**. Per the
   user, a Hero suggestion from this rule still needs the team's own
   manual confirmation — the card's reason text says so explicitly
   ("— confirm with team"); this tier is treated as less settled than
   the other three.
4. **2026 units > 5,000 AND 2026 GP% > 45%** → **Harvest**. Checked
   *after* Hero, not before — Harvest's bar is strictly looser (5k/45%
   vs. 10k/50%), so a franchise clearing both gets the stronger Hero
   call, not double-counted into Harvest.
5. **2026 units < 500, and not already Emerging** (rule 1 or 2) →
   **Exit**.
6. **Everything else** → **no suggestion** (left undecided). Expected
   to be the majority under these tighter, sales-only thresholds — see
   the distribution below.

## Thresholds — all user-set, 3-Nov-2026

| Tier | Units | GP% |
|---|---|---|
| Hero | > 10,000 | > 50% |
| Harvest | > 5,000 | > 45% |
| Exit | < 500 | — |

No judgment-call thresholds remain in this version (RR-048's one
unconfirmed Harvest floor is gone along with the rest of that
approach).

## Current distribution (4-Nov-2026, before the workshop)

| Suggested tier | Count |
|---|---|
| No suggestion | 139 |
| Emerging | 89 (56 new-franchise-in-season + 31 new-in-2026, 2 of which are new-generation cases per rule 2) |
| Harvest | 24 |
| Exit | 22 |
| Hero | 3 |

These are **suggestions on 277 active-range franchises**, most of
which have no real Tier decision yet (RR-047's export: 1 as of
3-Nov-2026) — this distribution will shift as more 2026 sales data
comes in through the rest of the year, and as the remaining
unmatched-sales gap closes (82 franchises still have no historical
sales matched at all, RR-037/RR-040).

## Revision history

- **3-Nov-2026, first pass (RR-048, superseded)**: Core Assortment +
  GP% thresholds (50%/30%) + generation-age + declining-trend. Full
  details preserved in RR-048's own register entry, not repeated here.
- **3-Nov-2026, RR-049**: user said "let me revise" and replaced the
  entire rule set with pure 2026 sales/GP% thresholds — simpler, no
  Core Assortment or generation-age dependency.
- **3-Nov-2026, RR-051**: Hero's GP% floor lowered 55% → 50%;
  Exit's units ceiling raised 200 → 500. Same rule structure as RR-049,
  thresholds only.
- **4-Nov-2026, RR-054 (current)**: rule 1 narrowed to new FRANCHISES
  only — a new GENERATION of an already-established franchise no
  longer auto-suggests Emerging, falling through to the same
  sales-based rules as any other franchise instead. 22 of 24
  new-generation cases moved off Emerging as a result (mostly to "no
  suggestion"; 2 to Harvest, 1 to Exit); 2 remain Emerging via the
  independent rule-2 sales signal, not rule 1.
