# Range Review tiering guidance (Hero / Harvest / Emerging / Exit)

Proposed with the user 3-Nov-2026, then fully revised same day (RR-049
supersedes RR-048's first pass entirely — Core Assortment, generation
age, and declining-trend are no longer part of this logic at all).
Maps to CLAUDE.md's "layer-by-layer keep/consolidate/improve/kill
proposal" framing for this workstream. Implemented as
`compute_suggested_tier()` in `scripts/build_workshop_data.py` — this
document and that function are the same guidance, not two independent
copies; if one changes, so should the other.

> **Pending taxonomy change (16-Nov-2026, RR-105) — not yet
> implemented.** The user has proposed replacing this tier set with
> **Core / Harvest / Emerging / Exit** — Hero is dropped as a tier
> value. In its place, **Icon** becomes a separate, cross-tier
> designation: a style can be an Icon regardless of which of the four
> tiers it's in (a Core Icon, a Harvest Icon, even an Exit Icon),
> rather than Icon/Hero being one of the mutually-exclusive options
> itself. This formalizes language that was already showing up
> organically in Action notes during the Nov 12 workshop and the
> Deep Review pass (e.g. "a potential Icon," "Spitz GTX Pro Jacket
> (Icon)") rather than introducing a new idea from scratch. Per the
> user, explicitly **"no need to update yet"** — this note captures
> the target definition only. Nothing below (the rules, the
> thresholds, the Board's Tier dropdown, `compute_suggested_tier()`)
> has been changed. See **Open questions** below before implementing.

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

## Resolved (16-Nov-2026, RR-106)

- **"Core" connects to the existing Core Assortment field** — per the
  user: "to connect to the existing[,] however there were a few cases
  that were wrongly labeled as Core. Follow the note, but
  Merchandising to confirm." So the new Core tier is meant to track
  `coreAssortmentFW27`/`coreAssortmentSS28`, not stay independent of
  it — but where the Deep Review pass already flagged a franchise as
  mislabeled (Brand Five Panel Cap, Hede Cargo II Pant, Magma Rugged
  Pant, and the pre-workshop Niklas-input gaps: Haglöfs L.I.M Horizon
  GTX Low/Low, Utility pouch 3L, Skuta 15, L.I.M Mountain Trail 25),
  the Action note's correction wins over the stale flag value. **Status:
  Provisional** — the connection and the override rule are both
  decided, but Merchandising still needs to confirm the corrected
  cases before this is fully Confirmed.
- **Icon stays a pure human call** — no suggestion-engine rule, unlike
  Hero's current 2026-units/GP% thresholds. Confirmed.
- **A "Potential Icon" tag was added to the Board** (not a rename of
  Tier itself, which still says Hero today) — see RR-106's register
  entry and the picks inventory below.

## Open questions (for the Core/Harvest/Emerging/Exit rename itself)

- **Scope of the rename.** Everywhere "Hero" appears today — this
  document's rules 3/4 and threshold table, `compute_suggested_tier()`,
  the Board's Tier `<select>` options, the `tier-definition-box` copy,
  `LIFECYCLE_GUIDANCE.md`'s Tier table, and every existing decision
  doc already carrying `tier: "Hero"` — needs a coordinated update
  when this is implemented, not a partial rename. Still **not
  implemented** — "no need to update yet" per the user.

## Icon picks and candidates (16-Nov-2026, RR-106)

Supplied by the user as a reference inventory, organized by status and
layer. "On Board?" says whether the franchise exists as a card today —
only those get the new **Potential Icon** checkbox; a style with no
card yet has nothing to check.

| Status | Apparel | Footwear | Hardware |
|---|---|---|---|
| **Established Icon** | Vassi, Spitz | — | Tight family |
| **Potential Icon — current range** | Front Proof (II Jacket), L.I.M GTX III **Jacket only** | — | — |
| **Potential Icon — future season** | FW28 Hede 2.0, L.I.M BirchGrid Fleece, L.I.M Light Jkt, SS29 Pant Platform | Vertigo family, SS28 L.I.M UL GTX Shoe | — |

**Established Icon** (Vassi, Spitz, Tight family) — reference only,
no tag applied. These are family-level names spanning several SKUs
each (e.g. "Vassi" matches 8 Board franchises across Shell/Legwear/
Midlayer/Accessories/Packs); the user asked specifically for a
**Potential** Icon tag, not an Established one, so the exact per-SKU
scope within each family was never resolved and isn't needed yet.

**Potential Icon — tagged on the Board** (`potentialIcon: true`,
confirmed matches only):
- Front Proof II Jacket
- L.I.M GTX III Jacket (**not** L.I.M GTX III Pant — confirmed with the
  user; the Pant was never discussed in Icon terms, the Jacket was)
- Haglöfs Vertigo QL PROOF Low (Vertigo family)
- Vertigo Proof Low (Vertigo family)
- L.I.M Ultralight GTX Low (= "SS28 L.I.M UL GTX Shoe")

**Potential Icon — not yet on the Board** (future/unreleased, nothing
to tag): FW28 Hede 2.0, L.I.M BirchGrid Fleece, L.I.M Light Jkt
(confirmed with the user: not the existing L.I.M Light Proof Jacket —
a genuinely different, unreleased product), SS29 Pant Platform. Revisit
once these exist as real franchises in `workshop_data.json`.

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
- **16-Nov-2026, RR-105 (proposed, not implemented)**: Core/Harvest/
  Emerging/Exit replaces Hero/Harvest/Emerging/Exit as the tier set;
  Icon introduced as a cross-tier designation, not a tier value. See
  the note at the top of this document and **Open questions** above.
- **16-Nov-2026, RR-106**: both RR-105 open questions resolved — Core
  connects to the existing Core Assortment field (Provisional, pending
  Merchandising confirmation on the known mislabeled cases); Icon
  confirmed as a pure human call, no suggestion-engine rule. A
  **Potential Icon** checkbox added to the Board and applied to 5
  confirmed franchises; the full Icon picks/candidates inventory
  recorded above. The Core/Harvest/Emerging/Exit rename itself is
  still not implemented.
