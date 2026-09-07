# Assumptions Register — SS26 Portfolio Tiering

This documents the data definitions, cleaning rules, and methodology
decisions behind the 8-tier SS26 DTC & Wholesale portfolio analysis
(`Project_Bob_Portfolio_Tiering_20082026.xlsx`), built from the three raw
exports in this folder (`SS26_DTC_Wholesale_w34_data_*.xlsx`).

**Live version (filterable, shareable):**
https://claude.ai/code/artifact/1b4cce0a-d4aa-4f87-9c8f-10926b1d28c4

This file is a durable snapshot of that page, checked in so it survives
independently of any one chat session. If the two ever disagree, the live
artifact is more current — update this file to match, not the other way
around.

**Reference style:** cite an entry by ID (e.g. "per REG-003") in any
formula, prompt, or report that depends on it.

---

## How this relates to `config.py` (Key-Article Margin Analysis)

This repo already contains a related, earlier workstream — the
Key-Article Margin Analysis package (`config.py` / `load_data.py` /
`analysis.py` / `data/bob_*.xlsx`), covering FY2024 vs FY2025 margin
trends across 16 product layers. Both workstreams touch the same
underlying business and some of the same special accounts, but they are
**separate exercises with separate source pulls** — don't assume a number
from one applies to the other without checking.

Three specific overlaps worth knowing about:

1. **Layer casing** — `config.LAYER_NAME_FIXES` already normalizes casing
   drift in the `Layer` field (e.g. "mid layer" → "Mid layer"). The SS26
   portfolio tiering file's `Layer` column is independently clean (verified,
   see REG-002 below) — consistent with the same fix having been applied,
   though the portfolio tiering pipeline itself doesn't reuse this repo's
   code, so this hasn't been confirmed as the *same* applied fix vs.
   independently rebuilt.
2. **Margin % outlier threshold** — `config.MIN_SALES_FOR_RELIABLE_MARGIN_PCT`
   (50,000 SEK per article per year; below it, GM% is blanked rather than
   shown) is a precedent worth reusing for REG-004 below, which is the same
   underlying problem (wild GM% on tiny denominators) in a different
   dataset. Not yet confirmed as adopted for portfolio tiering — see REG-004.
3. **Special-account treatment is genuinely different by design, not a
   conflict** — `config.ACCOUNT_CONTEXT` (XXL, Zalando/Zalando Lounge,
   China, STADIUM Outlet) **includes** these accounts in margin analysis
   but reads their trend with special context (e.g. Zalando's margin drop
   is expected outlet behavior, not demand erosion). The portfolio tiering
   file (REG-008) **excludes** the same accounts entirely from "Core."
   Both are legitimate choices for their respective questions — margin
   *interpretation* vs. portfolio *sizing* — but flag this explicitly to
   anyone using both files so it doesn't read as an inconsistency.

---

## Register

| ID | Domain | Field / item | Status | Rule & rationale |
|---|---|---|---|---|
| REG-001 | Data Hygiene | Article (and other free-text fields) | **Confirmed** | Strip leading/trailing whitespace on all text fields (Article, Color, Sales Market/Store, Ordertype, Model) before grouping/filtering/joins. Export tool pads some Article names with trailing spaces; confirmed across all three SS26 w.34 files, 25–45% of rows affected, 124–130 distinct articles silently split into two pivot lines per file. |
| REG-002 | Portfolio Analysis | Main Segment taxonomy | Open | Raw exports mix OUT/OUTDOOR, SNW/SNOW, MOUNTAIN, ACCESSORIES/ABS, SLEEPING BAGS inconsistently. **Note:** the published tiering file doesn't use `Main Segment` at all (only `Layer`, which is already clean) — so this may be moot for this specific deliverable. Still worth resolving if `Main Segment` is used elsewhere. |
| REG-003 | Portfolio Analysis | Canonical revenue metric | **Confirmed** | `Garp SEK Sales`, not `Sales`. Verified by reconciling the raw exports against the published tiering file at franchise (Base+Gender) level: 85% exact match on the SS26 YTD2026 cut (median difference 0.00 SEK); `Sales` matched exactly on only 19% of franchises. |
| REG-004 | Portfolio Analysis | Garp SEK Margin % outliers | Provisional | Raw exports show GM% outliers from roughly −696,000% to +2,491,000%, concentrated in claims/returns lines. `config.MIN_SALES_FOR_RELIABLE_MARGIN_PCT` (50,000 SEK/article/year, blank GM% below it rather than showing it) is an existing precedent for the same problem — proposed as the rule here too, not yet confirmed as adopted. Verified: outliers do **not** affect the tier assignment itself (none fall in Hero/Near-Hero/Harvest/Workhorse/Problem Child on the FY25 basis that drives tiering); a handful affect the supplementary YTD2026 GM% column but move sales-weighted tier averages by <0.1pt. |
| REG-005 | Portfolio Analysis | "SS26 & SS25" flag column | Open | Meaning undocumented; mostly/entirely blank across all three source files. Can't infer intent from data alone. |
| REG-006 | Portfolio Analysis | Season comparison scope | Provisional | Current SS26 season-to-date (Jan–Aug 2026) vs. full FY25 actuals (Jan–Dec 2025), Pace % = YTD2026 / full-FY25 (explicitly *not* annualized — a pre-read convention, per the tiering file's own intro sheet). Matches how the tiering file frames it; not yet independently confirmed as the intended workshop framing beyond what's stated there. |
| REG-007 | Portfolio Analysis | Sales Market / store-type granularity | Open | Some exports combine market + store type into one string (e.g. "Outlet Barkarby"); others (`data_3`) split them. Needs one parsing rule plus a market-to-region roll-up if this granularity is needed downstream. |
| REG-008 | Portfolio Analysis | "Core accounts" scope | **Confirmed**, with one open sub-item | Core excludes 5 named accounts: XXL, China, Zalando, Zalando Marketplace, Stadium Outlet. `XXL` and `China` are independently verifiable in the raw exports as distinct Sales Market values. `Zalando`, `Zalando Marketplace`, and `Stadium Outlet` don't appear as distinct values at this export's grain (wholesale rows aren't broken out to account name here) — taken on the documented definition. **Open sub-item:** `config.ACCOUNT_CONTEXT` lists "Zalando" (= Zalando Lounge) but no separate "Zalando Marketplace" — confirmed (2026-09-07) to be treated as a genuinely distinct account, not a naming variant, pending its own confirmation the way Zalando Lounge/China have already been confirmed with the business. |
| REG-009 | Cross-workstream | FY25 total sales discrepancy vs. `bob_salesdata_2025.xlsx` | Open | This repo's existing `data/bob_salesdata_2025.xlsx` (Key-Article Margin Analysis, "2026-07 data pull") totals 877.15M SEK Garp SEK Sales for FY25, all customer groups. The SS26 w.34 raw exports (data_2 + data_3, "core" scope not yet applied) total 856.78M SEK for the same period — a ~2.3% gap. Likely explained by pull-date drift (July vs. week-34/September pull) and/or account scope, but not confirmed. Don't treat the two FY25 totals as interchangeable until this is resolved. |

---

*Seeded 4 Sep 2026 from the SS26 DTC & Wholesale portfolio review. Reconciled against `config.py` 7 Sep 2026. Update the live artifact first, then re-sync this file.*
