# Range Review tiering guidance (Hero / Harvest / Emerging / Exit)

Proposed and refined with the user 3-Nov-2026 (RR-048). Maps to
CLAUDE.md's "layer-by-layer keep/consolidate/improve/kill proposal"
framing for this workstream. Implemented as `compute_suggested_tier()`
in `scripts/build_workshop_data.py` — this document and that function
are the same guidance, not two independent copies; if one changes, so
should the other.

**This is a suggestion engine, not a decision engine.** The Range
Review Board shows a "Suggested: <Tier> — <reason>" hint on any
undecided card these rules have an opinion on, but never writes to the
Tier dropdown itself — Tier stays a human call, captured in the
Board's own `decisions` database, exactly like every other decision
field (Last season to sell, Consolidate). Treat every suggestion here
as a starting point for the Nov 12 workshop discussion, not a verdict.

## Definitions

| Tier | Meaning |
|---|---|
| **Hero** | Proven, strong performer — invest further, a safe bet to buy more. |
| **Harvest** | Not Core Assortment, but healthy — keep selling, no fresh investment needed. |
| **Emerging** | New, unproven, or Core Assortment but not yet clearing the margin bar — give it a season or two, or watch closely. |
| **Exit** | Old, declining, and under-margin — the clearest kill/consolidate candidate. |

## Rules (first match wins)

1. **New in FW27, SS28, or FW28** (`isNewInSeason`) → **Emerging**,
   unconditionally. Too early to judge anything else about it — this
   overrides every other rule, including an otherwise-Exit-looking
   aging/declining read.
2. **Aging generation** (`generationAgeFlag` — the generation that is
   *actually selling*, per RR-041, not necessarily the catalog-current
   one) **AND declining** (2025 units < 2024 units, both full
   calendar years — 2026 is YTD-through-October and deliberately left
   out of this specific comparison to avoid an apples-to-oranges
   partial-year read) **AND 2026 GP% < 30%** → **Exit**. Applies
   regardless of Core Assortment status.
3. **Core Assortment** (FW27 or SS28, per RR-008/RR-028):
   - 2026 GP% **> 50%** → **Hero**
   - 2026 GP% **≤ 50%** (or no 2026 GP% to judge by) → **Emerging**
4. **Not Core Assortment**:
   - 2026 GP% **≥ 30%** → **Harvest**
   - 2026 GP% **< 30%**, or no 2026 GP% at all → **no suggestion**
     (left for a human to judge — not forced into a tier without the
     data to back it)

## Thresholds, and which are confirmed vs. a judgment call

- **50% GP% (Hero/Emerging split within Core)** — user-set, 3-Nov-2026. Confirmed.
- **30% GP% (Exit floor)** — user-set, 3-Nov-2026. Confirmed.
- **30% GP% (Harvest floor, non-Core)** — **not explicitly set by the
  user**; reuses the Exit floor as "the natural not-clearly-struggling
  line," my own judgment call (`TIER_HARVEST_MIN_GP_PCT` in the code).
  Flag if a different number is wanted.
- **4+ seasons (generation age flag)** — set earlier in this
  workstream (RR-029/RR-030, lowered from 6 to 4 per user direction),
  reused here rather than inventing a separate aging threshold for
  tiering specifically.
- **2024→2025 full-year comparison for "declining"** — a methodology
  choice to avoid comparing a partial 2026 YTD figure against full
  years; not separately confirmed with the user.

## Known gap: the 30% Exit floor hasn't fired yet

As of 3-Nov-2026, **23 franchises** are aging *and* declining
2024→2025, but **none** currently has 2026 GP% under 30% — the lowest
is 37.7% ("Vassi GTX II Jacket"). So rule 2 (Exit) currently suggests
nothing. This is a real finding, not a bug: margins on declining-volume
old generations in this portfolio are mostly holding up, at least
through October 2026. Worth flagging at the workshop — either the 30%
floor is set correctly and genuinely nothing qualifies yet, or the
team may want a higher floor (e.g. the Harvest floor, 30%, vs. a
tighter Exit-specific number) to actually catch anything.

## Current distribution (3-Nov-2026, before the workshop)

| Suggested tier | Count |
|---|---|
| Harvest | 128 |
| Emerging | 93 (80 new-in-season + 13 Core Assortment, GP% ≤ 50%) |
| Hero | 17 |
| Exit | 0 (see gap above) |
| No suggestion | 39 |

These are **suggestions on 277 active-range franchises**, most of
which have no real Tier decision yet (RR-047's export: 1 as of
3-Nov-2026) — this distribution will shift as real sales data and Core
Assortment coverage fill in gaps (82 franchises still have no matched
historical sales at all, RR-037/RR-040).
