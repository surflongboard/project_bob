# Range Review lifecycle-date guidance (Last season to sell / No new POs as of)

Companion to `TIERING_GUIDANCE.md`, which covers how **Tier** (Hero /
Harvest / Emerging / Exit) gets suggested. This document covers the two
date fields that sit alongside Tier on every Board card — what each one
means, who it's for, and what value it should carry for each Tier.
Proposed with the user 16-Nov-2026 (RR-103/RR-104), after a consistency
review of Tier vs. the dates across all 269 decisions surfaced that
"Last season to sell" was doing two jobs for two different audiences at
once.

## Two fields, two audiences

- **Last season to sell** — **sales-facing**. What the sales tool shows;
  how long a style is still fair game to sell/market.
- **No new POs as of** — **purchasing-facing**. The cutoff after which
  Purchasing stops placing new orders.

These two dates can legitimately differ. An Exit style with real stock
still on hand is the clearest case: Purchasing stops ordering
immediately, but Sales can — and should — keep selling it until that
stock actually runs out, which may be a later season than the
purchasing cutoff. Treating one field as if it served both questions is
what produced apparent "inconsistencies" in the Nov 12 data that weren't
really inconsistencies — see RR-103's register entry.

**No separate "sell from stock" checkbox** — considered and rejected
16-Nov-2026. The recurring free-text pattern ("No POs, sell from the
stocks") that shows up across many Action notes is exactly what the
**No new POs as of** field already captures in structured form: once
that field carries a season, it *is* the sell-from-stock state. Adding
a second, redundant flag for the same thing was the wrong move; the
field itself is the mechanism.

## Rules by Tier

| Tier | No new POs as of | Last season to sell |
|---|---|---|
| **Hero / Harvest** | Open-ended by default — keep producing as long as there's a forecast. Only gets a cutoff when the Action note states one (a "sell from stock" case). | Mirrors No new POs as of — no separate cutoff unless the note gives one. |
| **Emerging** | Not applicable yet — still in its demand-proving phase. | **Undecided** by default. Stays open until the style's performance is confirmed; only gets a value if the note explicitly states an exit point. |
| **Exit** | **Immediate — SS27.** Not "soon," not "next season": the cutoff is now, as of the season already in progress when this review happened. Applies across the board, with Hardware the clearest/most consistent case so far. | Can run later than the purchasing cutoff, reflecting how long remaining stock is still sold — the Action note's rationale should support whatever season is picked here. |

## Why SS27, not FW27

"Immediate" for an Exit style means **SS27** — the season already
underway at the time of this review, and the same boundary RR-020/
RR-021 already established as this workstream's end-of-life cutoff
("anything not carried into FW27 or FW28 is end-of-life as of end of
SS27"). The dropdown (`LAST_SEASON_OPTIONS` in
`scripts/workshop_page_template.html`) didn't originally include SS27
at all — it started at FW27, since FW27/SS28 is this review's own
active-range scope. People picking "FW27" for an immediate exit were
reaching for the earliest option available, not actually choosing
FW27 over SS27; the option they needed didn't exist. Added 16-Nov-2026
(RR-104) — both the Last season to sell and No new POs as of dropdowns
now include SS27, positioned before FW27 in the sequence.

## Open edges

- **Tier taxonomy change pending.** `TIERING_GUIDANCE.md` (RR-105,
  16-Nov-2026) proposes replacing Hero with Core as a tier, and
  introducing Icon as a cross-tier designation instead. Not yet
  implemented — the rules-by-Tier table above still says "Hero /
  Harvest" because that's what the Board's Tier field still offers
  today. When that change lands, this table's row needs the same
  rename; it should behave identically either way (open-ended
  production unless the note says otherwise), since the rename
  doesn't change the lifecycle-date logic itself.
- **Apparel Exit timing** isn't confirmed the way Hardware's is. The
  "immediate = SS27" rule was stated with Hardware specifically in
  mind; whether the same immediacy applies to Apparel Exit cases is
  still open (6 non-Hardware Exit franchises had a non-FW27/non-SS27
  last season as of the Nov 12 data: Essential Tights, L.I.M Airak GTX
  Jacket, L.I.M Delta Crop Tee, Outsiders Graphic Tee, Outsiders Tee,
  Vassi Mid Hood — not yet revisited against this guidance).
- **Backfill**: `noNewPOsAsOf` exists now but started blank for all 269
  existing decisions. Filling it in for the Hardware Exit cases (per
  this guidance, `SS27`) is a separate, explicit pass — not done
  automatically just because the field and the rule now exist.

## Revision history

- **16-Nov-2026 (RR-103)**: fields split — No new POs as of added,
  separate from Last season to sell.
- **16-Nov-2026 (RR-104)**: this document written; SS27 added to the
  season dropdown; confirmed no separate "sell from stock" checkbox is
  needed (the No new POs as of field already serves that role).
