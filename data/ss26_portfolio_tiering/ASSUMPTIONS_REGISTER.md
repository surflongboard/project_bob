# Assumptions Register — SS26 Portfolio Tiering

This documents the data definitions, cleaning rules, and methodology
decisions behind the SS26 DTC & Wholesale portfolio analysis, built from
the three raw exports in this folder (`SS26_DTC_Wholesale_w34_data_*.xlsx`).

- `Project_Bob_Portfolio_Tiering_20082026.xlsx` — original 8-tier version.
- `Project_Bob_Portfolio_Tiering_08092026.xlsx` — current version, adds a
  9th tier (Clearance — Ex China/Zalando) per REG-010. Use this one going
  forward; the original is kept for reference only.

**Live version (filterable, shareable):**
https://claude.ai/code/artifact/1b4cce0a-d4aa-4f87-9c8f-10926b1d28c4

**Ask the data directly (no Excel/pivoting needed):**
https://claude.ai/code/artifact/aac8af79-3113-4f05-a0a5-557231bdd958 — "Tier Bench," a chat tool over the current tiering dataset (franchise, generation-comparison, and channel-mix tables). Answers are Claude-generated from live queries against the data below, not pre-written — verify anything going into a decision deck against the source workbook.

**Before presenting this to the wider team:**
`LAUNCH_READINESS_CHECKLIST.md` (same folder) — the open items, process gaps, and Tier Bench QA steps to close out first.

**Reusable pipeline scripts:** `scripts/` (same folder) — every REG entry
below whose column/sheet was added via one-off Python now has a committed,
documented, re-runnable script (`update_*.py`) instead. Raw manual-upload
source files live in `inputs/` alongside them. See `scripts/README.md` for
what each script does and the intake convention for new data drops.

This file is a durable snapshot of that page, checked in so it survives
independently of any one chat session. If the two ever disagree, the live
artifact is more current — update this file to match, not the other way
around.

**Reference style:** cite an entry by ID (e.g. "per REG-003") in any
formula, prompt, or report that depends on it.

---

## Recent changes

Newest first, plain language — so a first-time reader sees the pace of
iteration without reading all 24 register entries below.

1. **10 Sep 2026 — Flowed Core Assortment FW27 (REG-015) into the Hero
   Tier Catalogue.** A third card badge, "Core FW27: Yes/No," read
   straight from Sheet 2 — no new methodology. Same split REG-015 has
   documented since 8-Sep: 14 of the 34 Global Hero+Near-Hero franchises
   are on the confirmed 88-style list, 20 are not (a merch check, not a
   discontinuation signal — matched by style name, not sales history, so
   an unshipped successor doesn't wrongly show absent).
2. **9 Sep 2026 — Flowed regional tiering (REG-023/024) into the Hero
   Tier Catalogue.** Each of the 34 cards now shows a Nordic and a ROW
   (Rest of World / Non-Nordic) badge alongside the existing Global tier
   chip — highlighted when that franchise is also a Hero within that
   region alone, muted (with the actual regional tier on hover) when it
   isn't. Read straight from Sheets 6/7, no new methodology. Real split
   across the 34: 4 Hero in both regions, 17 Nordic-only, 7 ROW-only, 6
   neither (Global Hero from combining two mid-sized regions, not from
   dominating one) — summarized on the page itself, not just per-card.
3. **9 Sep 2026 — Flowed regional tiering (REG-023/024) into Tier
   Bench.** New seventh query tool, `queryRegionalTiering` — one row per
   franchise per region (Nordic/Non-Nordic, 3,720 rows), with
   `globalTier` alongside `regionalTier` for direct comparison. Verified
   against the workbook before publishing (32 Nordic / 19 Non-Nordic
   Hero+Near-Hero, exact match) and confirmed it correctly surfaces real
   regional-vs-global disagreements rather than treating them as errors
   — e.g. "ROC Flash Down Hood" (Men) is Workhorse+Harvest globally but
   Hero+Near-Hero in Nordic (185.9% regional growth). (Also fixed a
   same-day Tier Bench outage this caused — the new tool's description
   field exceeded a previously-unknown 1KB per-tool limit enforced by the
   underlying chat API, which silently broke every query, not just ones
   using that tool — trimmed the description and confirmed all seven
   tools work via an independent test harness before republishing.)
4. **9 Sep 2026 — The base tiering rule is confirmed (REG-024), and the
   regional sheets from earlier today were rebuilt to use it.** After
   yesterday's failed reverse-engineering attempt (below), the business
   supplied the actual rule directly: absolute SEK/GM%/growth cuts
   (2M SEK + growth >0% + GM% ≥46.6% + cross-channel for Hero, etc. —
   full detail in REG-024). Tested it against Sheet 2's own published
   tiers: **1,808 of 1,812 non-Clearance franchises match exactly
   (99.8%)** — the 4 mismatches are article-parsing edge cases, not rule
   failures. Sheet 2's Tier column is NOT rewritten (stays authoritative
   even though it's now reproducible — see REG-024 for why). Rebuilt
   Sheets 6/7 (REG-023) to use this real rule instead of yesterday's
   invented placeholder: since the rule's SEK floors are global/absolute,
   scaled each region's floors by its own share of FY25 Core revenue
   (Nordic ≈68%, Non-Nordic ≈32%) rather than reusing the unscaled
   global numbers — confirmed with the business first, since reusing
   them unscaled would have been "same absolute thresholds as Global," an
   option already declined. Also fixed a real gap found while rebuilding:
   733 of 1,860 franchises had been silently missing from the Non-Nordic
   sheet (any franchise with zero Non-Nordic sales) — every franchise now
   gets at least one row per region.
5. **9 Sep 2026 — Added region-relative tiering for Nordic and
   Non-Nordic (REG-023 — since superseded, see above), after finding the
   original tier rule can't be reverse-engineered.** Team wanted the
   same Hero/Near-Hero/etc. structure computed separately per region.
   Tried hard to recover Sheet 2's actual cut rule first — tested
   absolute thresholds, top-N by sales, top-N per product layer, and
   several composite scores; none reproduce the real Hero+Near-Hero set
   (best got 41%, no better than noise) — looked at the time like a
   business judgment call with no recoverable formula (it turned out to
   be a real, precise formula — see above, it just wasn't derivable from
   Sheet 2's columns alone). Built a placeholder rule to keep moving;
   superseded same day once the real rule arrived.
6. **9 Sep 2026 — Added a Sales-by-Country sheet + a browse-order view
   (REG-022), then resolved the open definition question same day.**
   Team feedback: they want to see tiers by region, specifically able to
   exclude the home/legacy market from a query like "Hero products in
   Japan and Germany." New Sheet 5 breaks FY25/YTD2026 sales & units out
   by country per franchise (13,438 rows), hand-mapped from the raw
   exports' messy store/market field since it has no clean country
   field. Initially left "Scandinavian vs Nordic" open as a genuine
   definition choice; **business directed the same day: use Nordic
   (Sweden/Norway/Denmark/Finland) as the standard exclusion group** — a
   new `Region Group (Nordic vs Non-Nordic)` column is the one to filter
   on, with the stricter Scandinavian-only (no Finland) flag kept
   alongside for anyone who wants that narrower cut. Also fixed a real
   error caught before it published: a small "Unmapped / Other" bucket
   (3.7% of rows, no confirmed country) was initially going to silently
   count as "Non-Nordic" — corrected to its own "Unmapped" value so it
   doesn't inflate that total. Also added a read-only "Browse" sheet
   grouping Sales/Units/GM% together per year for easier scanning,
   without reordering Sheet 2 itself (which all 9 pipeline scripts
   hardcode column positions against).
7. **9 Sep 2026 — Added Style Code(s) (REG-021), same day someone asked
   Tier Bench for one.** Tier Bench correctly said it didn't have a SKU/
   style code field rather than guessing — none of the published Sheet-2
   data ever carried one. Added a `Style Code(s)` column sourced from the
   raw exports' `Model` field: most franchises resolve to one code, but
   ~23% carry 2–6 (the same product name can get a new Model number
   across seasons) — read as "codes seen for this franchise," not a
   single canonical code. Now in the workbook, Tier Bench, and its CSV
   exports.
8. **9 Sep 2026 — Fixed a second Tier Bench data gap: YTD2026 units were
   missing.** `Units_YTD2026` has been a published Sheet-2 column since
   before this session's work (it's part of the original base tiering
   build, alongside Sales_2025/Sales_YTD2026/GM%/Pace%/Growth%) — but
   Tier Bench's franchise data was built without ever pulling that
   column in, so `queryFranchises` could answer "FY25 units" (once
   REG-020 landed) but not "FY26 units," even though the workbook always
   had both. Found because someone asked Tier Bench for it directly and
   got a (correct, not fabricated) "I don't have that field" answer
   instead of a wrong number. Added `unitsYtd26` alongside the existing
   `units25`; republished.
9. **9 Sep 2026 — Added real FY2025 Units (REG-020).** A `Units_2025`
   column on Sheet 2, sourced from the same raw exports' "Units Sold"
   field — not an estimate. Flowed through to Tier Bench and the Hero
   Tier Catalogue's financial detail (previously showed "—" for FY25
   units since no real figure existed yet).
10. **8 Sep 2026 — Fixed a real bug in Tier Bench, found via QA.** Its four
   filterable query tools silently returned ZERO results whenever a
   question needed a limit, a sort field, or a sort direction (i.e. most
   "top N" / "sorted by" questions — a large share of real use). Found by
   testing the tool code directly rather than through the live chat;
   fixed and re-verified against 12 known answers pulled from the
   workbook/register. Also added CSV export per query result.
11. **8 Sep 2026 — Pipeline made reusable.** Every column/sheet below that
   was built from a manual data upload now has a committed, re-runnable
   script instead of one-off chat Python (`scripts/`); the raw source
   files themselves are checked in too (`inputs/`). PR #2 (everything in
   this workstream so far) merged to `main`.
12. **8 Sep 2026 — Added a stock breakdown by tier (REG-019).** A warehouse
   available-stock snapshot, matched to franchise and tier, so it's clear
   which stock is a clearance candidate (Thin/Immaterial, Exited) vs. a
   real seller with a margin problem (Problem Child) that shouldn't be
   fire-saled.
13. **8 Sep 2026 — Added channel-mix columns (REG-017, REG-018).** Each
   franchise's Wholesale vs. DTC sales split, plus a year-over-year
   pattern label (e.g. "DTC growing while Wholesale shrinks"). Revised
   same day after the first cut mislabeled a third of one bucket as
   "insufficient data" when they actually had real, single-channel sales.
14. **8 Sep 2026 — Added FW27 assortment-status columns (REG-015,
   REG-016).** Whether each franchise is in the Core Assortment for FW27,
   and whether it's marked Active in the separate Assortment Attribution
   Review — flagging where Hero-tier products are missing from either
   list, for a merch check.
   with "on the clearance list" — they're not the same thing.

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
| REG-010 | Portfolio Analysis | Clearance treatment (China JV / Zalando close-out stock) | **Confirmed** (2026-09-08) | Two Aug-2026 stock clean-up lists (`Close_out_delivery_JV_August_2026.xlsx` = China JV distributor, 321 SKUs; `Zalando_close_out_order_August_2026.xlsx`, 127 SKUs; 421 distinct SKUs combined) identify old-generation SKUs being exited from the assortment. Matched to franchises via SCO (Model+Color ID) → Article → Base/Gender, reusing `config.GENDER_TOKENS`/`VERSION_TOKENS` from the Key-Article Margin Analysis workstream (`find_version_families` precedent), plus two fallbacks for glued-text export artifacts (no space before "Men"/"Women" or before "II"). 418/421 SKUs matched a raw sales history; all 418 matched an existing published franchise; 3 SKUs never appear in any raw export (no impact, noted for completeness). **Rule (business-directed, not derived):** per franchise, compare its full historical SKU set against the clearance-list SKUs. "Fully" (48 franchises) = every known SKU is on the list → **physically moved** into a new tier, `Clearance — Ex China/Zalando`, in `Project_Bob_Portfolio_Tiering_08092026.xlsx` (tier totals adjusted; grand total unchanged — a re-label, not a deletion/addition). "Partial" (212 franchises, incl. 3 Hero, 7 Near-Hero, 58 Problem Child) = a current-generation SKU survives → **franchise stays in its original tier untouched**, only annotated via new `Clearance Flag`/`Clearance Detail` columns on the Full Portfolio sheet. No SEK amount was split out of Partial franchises' totals — see REG-012 for why a $-level carve-out was attempted and abandoned. |
| REG-011 | Data Hygiene | Case-sensitive duplicate franchise names | Open | Found while matching REG-010: the raw `Article` field produces at least one case/spacing-sensitive duplicate franchise pair in the published tiering file — `"Roc Sight SoftshellJacket"` (Harvest/Problem Child tiers) vs. `"ROC Sight Softshell Jacket"` (New/Test, Thin/Immaterial tiers) — apparently the same product line split into two "franchises" by a casing/spacing inconsistency upstream. Not corrected in either tiering workbook; worth a check with the data owner before assuming the full franchise list is deduplicated. Likely not an isolated case — only found because this one happened to intersect the clearance list. |
| REG-012 | Portfolio Analysis | `Ordertype = "3-Close out order"` in Core-scope Sales_2025 | Open | Discovered attempting a SEK-level clearance carve-out (abandoned in favor of REG-010's SKU-membership rule): reconstructing a franchise's FY25 Core sales from raw rows (Sales Market ≠ XXL/China, per REG-008) does not reconcile against the published `Sales_2025` figure whenever the franchise has meaningful `Ordertype = "3-Close out order"` activity — a **standing** wholesale ordertype, not specific to the Aug-2026 lists, accounting for **13.5% of all FY25 Garp SEK Sales** (116M of 860M) across the full raw dataset. Example: "Roc Mimic Hood" (Women), published Sales_2025 = 300,637 SEK; all-Core-rows reconstruction = 467,996 SEK; excluding `3-Close out order` rows = 126,469 SEK — the published figure sits between the two, so the published Core-scope definition applies some rule to this ordertype that isn't reproducible from the raw exports alone. Likely entangled with REG-008's open sub-item (Zalando Marketplace/Stadium Outlet not distinguishable at this grain — a "close out order" row may *be* one of those accounts routed through a generic country market). Needed before any future SEK-level reconstruction from raw data is attempted against the published totals. |
| REG-013 | Portfolio Analysis | Tier consolidation (Hero+Near-Hero, Workhorse+Harvest) | **Confirmed** (business-directed, 2026-09-08) | Hero and Near-Hero (Rising Star) merged into one tier, "Hero + Near-Hero"; Workhorse and Harvest (Cash Cow) merged into one tier, "Workhorse + Harvest" — because SS26 YTD2026 sales growth and margin trends were moving in ways that made the original size-based cut lines between each pair a less consistent decision boundary. Applied on top of the already-current (post-REG-010) row-level data: each combined tier's Sales_2025/GM%_2025/Sales_YTD2026/GM%_YTD2026/Pace% is the sales-weighted roll-up of its two source tiers, computed from row-level figures, not the rounded summary-table inputs — grand TOTAL unaffected (649.76M SEK). The pre-consolidation tier (9-tier scheme including the REG-010 Clearance move) is preserved per-franchise in a new `Original Tier (pre-consolidation)` column (Sheet 2, column P) so the original split can always be reconstructed. |
| REG-024 | Portfolio Analysis | Base 8-tier commercial portfolio rule (Sheet 2's Hero/Near-Hero/Workhorse/Harvest/Problem Child/New-Test/Exited/Thin-Immaterial classification) | **Confirmed** (business-supplied 9-Sep-2026, validated 99.8%) | Previously undocumented anywhere in this repo — an extensive attempt to reverse-engineer it from Sales_2025/GM%_2025/Growth%/Pace% (see REG-023's history) failed; the business then supplied the actual rule directly: **New/Test** = no real trading (≥100 units, `MIN_UNITS_FOR_REAL_TRADING_YEAR` in `config.py`) in FY24, real trading in FY25; **Exited** = real trading in FY24, not FY25; **Thin/Immaterial** = below the 100-unit real-trading threshold in either year; **Hero** = FY25 sales ≥2M SEK, YoY growth >0%, cross-channel (Wholesale >100K SEK AND DTC >50K SEK), GM% ≥46.6% (the core-company average — the "good" bar, not a round number); **Near-Hero (Rising Star)** = FY25 sales ≥1M SEK, YoY growth >15%, same cross-channel test, GM% ≥46.6%; **Workhorse** = not Hero/Near-Hero, GM% ≥40% (a separate, lower floor), YoY growth > −10%; **Harvest (Cash Cow)** = same as Workhorse but YoY growth ≤ −10%; **Problem Child** = not Hero/Near-Hero, GM% <40% regardless of growth. Implemented as `ss26_lib.classify_tier()`, validated by `scripts/verify_base_tiering.py` (read-only audit, does not rewrite Sheet 2). | Recomputed from `bob_salesdata_2024.xlsx`/`bob_salesdata_2025.xlsx` (Core-scope via `is_core_customer_group()`, "Marketplace" grouped into Wholesale per REG-018's precedent) and compared against Sheet 2's own "Original Tier (pre-consolidation)" column: **1,808 of 1,812 non-Clearance franchises match exactly (99.8%)**. The 4 mismatches are article-name parsing edge cases in this reconstruction (a literal `"*MISSING*"` placeholder base name, a glued gender suffix like "RT1Women"), not rule failures. Sheet 2's published Tier column is **not** rewritten by this script — it stays authoritative even though it's now reproducible, since (a) the 0.2% gap could reflect either this reconstruction's parsing noise or the original build's access to more precise source data, and (b) overwriting audited, published tiers with a 99.8%-but-not-100%-matching recomputation would be a regression, not an improvement. This rule is what REG-023 (below) now actually implements for the regional sheets, replacing that entry's original invented placeholder. |
| REG-023 | Portfolio Analysis | Region-relative tiering — Sheets 6 & 7, "Full Portfolio (Nordic)" / "(Non-Nordic)" | **Confirmed** (uses REG-024's real, validated rule; SEK-floor scaling method **business-directed** 9-Sep-2026), figures **Provisional** | **Rewritten after REG-024 was confirmed** — the original version of this entry (see git history for the full account) invented a new percentile-based rule after failing to reverse-engineer Sheet 2's logic; once the business supplied the real rule (REG-024), this was rebuilt to use it properly instead. The real rule's SEK floors (2M Hero, 1M Near-Hero, 100K/50K cross-channel) are **global, absolute** numbers — applying them unscaled to each region separately would just be "same absolute thresholds as Global," an option already declined, since Non-Nordic's smaller pool would rarely clear a bar calibrated to global-scale revenue. **Resolved: each region's SEK floors are scaled by that region's own share of total FY25 Core sales** (computed fresh each run from `bob_salesdata_2025.xlsx`, not hardcoded — currently Nordic ≈68.1%, Non-Nordic ≈31.9%, e.g. Hero floor ≈1.36M SEK for Nordic, ≈638K for Non-Nordic). GM%/growth cuts and the 100-unit real-trading threshold are **not** scaled — GM%/growth are already relative (%), and 100 units is a physical reality-check, not a revenue-scale artifact. Clearance is carried over unchanged from Sheet 2 (SKU-list membership isn't region-specific). | Same data source and Core-scope filter as REG-024 (`bob_salesdata_2024/2025.xlsx`), aggregated per region instead of globally — carries the same REG-009 ~2.3% cross-workstream gap caveat. **Concrete result of using the real rule instead of the old placeholder:** "Astral GTX Jacket" Men now stays Workhorse in *both* Nordic and Non-Nordic (GM% 42.9%/44.5%, both below the 46.6% Hero bar) — consistent with its Global tier, for the actual reason (margin, not size) — where the old placeholder rule had it flip to Hero in both regions. Result: **32 Nordic Hero+Near-Hero vs. 19 Non-Nordic** (region-relative — reflects rank within a smaller home population at a proportionally-scaled bar, not superior absolute performance). Country breakdown per region reuses Sheet 5's already-validated rows (REG-022), filtered to the region. **Fixed a real completeness gap from the first version:** every franchise now gets at least one row per region sheet, even with zero regional sales (marked "— (no regional sales)") — the first version silently omitted 733 of 1,860 franchises from the Non-Nordic sheet (any franchise with no Non-Nordic country activity at all). **Flowed into Tier Bench (9-Sep-2026) as a seventh query tool, `queryRegionalTiering`** — one row per franchise per region (3,720 rows), with `globalTier` alongside `regionalTier` for direct comparison. Independently tested against the tool's actual execute() code before publishing: Nordic/Non-Nordic Hero+Near-Hero counts match the workbook exactly (32/19), and querying for rows where the two tiers disagree correctly surfaces real, expected cases — e.g. "ROC Flash Down Hood" (Men) is Workhorse+Harvest globally but Hero+Near-Hero in Nordic (185.9% regional growth). Tier Bench's system prompt is explicit that a regional/global mismatch is the intended effect of region-relative thresholds, not a data error, so it doesn't get reported as one. **Also flowed into the Hero Tier Catalogue (9-Sep-2026)** — a Nordic and a ROW (Rest of World / Non-Nordic) badge on each of the 34 cards, highlighted when that franchise is a Hero within that region alone: 4 of 34 in both regions, 17 Nordic-only, 7 ROW-only, 6 in neither region alone. |
| REG-022 | Portfolio Analysis | Sales by Country (Sheet 5) | **Confirmed** (methodology; "Nordic vs Non-Nordic" grouping **business-directed 9-Sep-2026**), figures **Provisional** per REG-012 | New Sheet 5, "5. Sales by Country," reusable script `scripts/update_country_breakdown.py`: FY25/YTD2026 Sales/Units per franchise **per country** (13,438 franchise-country rows), so a question like "Hero products in Japan and Germany, excluding Nordic countries" can be answered directly rather than only at the whole-portfolio grain. **Source problem (REG-007):** the raw exports have no clean country field — `data_3` has a mostly-clean `Sales Market`, but `data_1`/`data_2`'s `Sales Market/Store` field mixes DTC store identity into the same field (e.g. "Outlet Barkarby," "Brand Store Sthlm," "E-com Sweden"). Built a hand-verified mapping of all 50 distinct raw values across the three files to a country (e.g. "Outlet Haparanda" → Sweden, "Brand Store Chamonix" → France, "Outlet Helsinki"/"Brand Store Helsinki" → Finland — resolved from each store name's real-world location). **The "Scandinavian vs Nordic" question originally flagged Open here was resolved same day:** the standard exclusion group is `Region Group (Nordic vs Non-Nordic)` = Sweden/Norway/Denmark/Finland vs. everywhere else; the stricter `Scandinavian (SE/NO/DK)` flag (no Finland) stays available as a separate column for anyone who specifically wants that narrower cut. | **One thing still deliberately left open, not silently resolved:** 994 "Pop-Up Sales Haglöfs" + 145 "Export Other" raw transaction rows (501 of the 13,438 franchise-country rows in this sheet, 3.7%) carry no city/country in the raw label at all — mapped to an explicit `Unmapped / Other` country, and **`Region Group` gives this its own "Unmapped" value rather than folding it into "Non-Nordic"** (folding it in would quietly inflate a Non-Nordic total with a geography nobody actually confirmed). `XXL` (523 raw rows) is dropped entirely — an account, not a geography (REG-008), consistent with every other sheet. Coverage: 13,438 of 13,717 franchise-country combinations with any raw activity matched an existing Sheet-2 franchise (279 did not — prototypes/non-apparel/discontinued items outside the curated 1,860-franchise scope, the same class of gap documented in REG-019). **Figures are raw-recomputed, not the published Sheet-2 totals** — summing a franchise's country rows will run higher than its audited Sales_2025 (this sheet is NOT Core-scope filtered the way Sheet 2 is — China appears as its own country row here, on purpose, so a real geography question isn't silently missing a market). Sanity-checked: "Astral GTX Jacket" (Women) — Sweden+Finland+Denmark together are ~77% of its total reconstructed sales, a concrete illustration of exactly the home-market skew this sheet was built to let people exclude. Flowed into Tier Bench as a sixth query tool, `queryByCountry`, with `regionGroup` as the primary filter. |
| REG-021 | Portfolio Analysis | Style Code(s) (franchise → raw Model code lookup) | **Confirmed** (methodology), **not a canonical 1:1 code** | New `Style Code(s)` column on Sheet 2 (column V), reusable script `scripts/update_style_codes.py`. Collects every distinct raw `Model` value seen for a franchise (Base+Gender) across all three SS26 exports (SS26YTD + both FY25 halves), comma-separated, sorted — added so people tracking a franchise in another system (PLM/ERP) have a code to search on; the published workbook otherwise carries no code at all, only franchise names. **Model is NOT 1:1 with franchise, or even with the exact Article text** — the same article name can carry several Model codes across seasons (a re-launch keeps the product name but gets a new internal Model number): of 1,636 franchises with any matching raw data, 1,261 (77%) have exactly one code, 375 (23%) have 2–6. Read this column as "codes seen for this franchise," not "the" style code. Coverage: 224 of 1,860 franchises have no raw data at all (no code) — consistent with the same class of small join-gap seen in REG-017/020 (raw-export franchise-key text not exactly matching the published workbook's). Cross-checked: "ROC Flash Down Hood" (Women) resolves to a single code, 607466 — matches the Model number already cited in that same franchise's existing `Clearance Detail` note (REG-010), an independent confirmation the join is correct. |
| REG-020 | Portfolio Analysis | Units_2025 (real FY25 unit sales) | **Confirmed** (methodology), figures **Provisional** per REG-012 | New `Units_2025` column on Sheet 2 (column U), reusable script `scripts/update_units_2025.py`. Sums the raw `Units Sold` field (present in all three SS26 exports, not previously pulled) from the two FY25 exports (Jan–Aug, Sep–Dec), Core-scope only (`is_core_market`, REG-008), matched by franchise (Base+Gender) — the exact same two files and method already used for REG-017/018. **Not blanked below MIN_RELIABLE** — unlike GM%/growth/share, this is a raw count (like Sales_2025 itself), so a small-denominator rate problem doesn't apply. Coverage: 1,527 of 1,860 franchises have a nonzero value; 333 have no matching raw units at all — consistent with REG-017's 330 "no raw sales data" count (small 3-franchise difference likely reflects a franchise with reconstructed sales but net-zero/negative units from returns, or vice versa). Sanity-checked: implied per-unit price (Sales_2025 / Units_2025) for the top 10 Hero+Near-Hero franchises ranges 290–1,900 SEK, consistent with Haglöfs jacket/pant/baselayer retail pricing. **Inherits REG-012's caveat** — reconstructed from raw rows rather than an audited published total, so treat as directional/consistent-methodology rather than an audited unit count. Added to Tier Bench's franchise data and the Hero Tier Catalogue's financial detail (previously showed "—" for FY25 units). |
| REG-019 | Portfolio Analysis | Available Stock by Tier (Sheet 4) | **Confirmed** (methodology) | New Sheet 4, `Available_stock_260825.xlsx` (warehouse available-stock snapshot, units by Article No./SKU) matched to franchise (Base+Gender) via the file's own `Article` field, same gender/version-token parsing as REG-010/014/015. 527,963 total units in the source file; **513,640 (97.3%) matched** an existing published franchise; the remaining 2.7% (14,323 units, e.g. "Midnattssol Crewneck") are styles not yet in the tiering file — likely brand-new FW27 launches with no FY24/25 history to tier against. **Units only — no SEK stock value**, the source file carries no cost/price column; not multiplied by an average price. Tier-level breakdown with a business-directed Recommended Action per tier: Hero+Near-Hero (protect availability), Workhorse+Harvest (maintain), Problem Child (fix or watch, NOT clear — large stock on a real seller like "Front Proof Jacket" — 4,337 units, 9.2M SEK FY25 sales — is a margin problem, not dead stock), New/Test (monitor), **Thin/Immaterial and Exited (clear via sale)** — the two tiers holding stock with materially zero/near-zero FY25 sales (e.g. "Spacelite -1," 3,405 units, 0 SEK FY25 sales), Clearance tier (already earmarked, monitor). Detail tables on Sheet 4 list the top stock-holding franchises per tier with a per-line "dead stock" signal (FY25 sales < 1,000 SEK). |
| REG-018 | Portfolio Analysis | Channel Pattern 2025 (DTC x Wholesale substitution classification) | **Confirmed** (business-directed taxonomy, revised same day) | New `Channel Pattern 2025` column on Sheet 2, classifying each franchise's 2024→2025 Wholesale vs. DTC (Retail+E-com) pattern. **Source: `data/bob_salesdata_2024.xlsx` + `bob_salesdata_2025.xlsx`** (Key-Article Margin Analysis workstream, full two-year annual pull) — not the SS26 w.34 exports used for REG-012/014/017 — since "that channel analysis" referred to `analysis.py`'s existing `channel_view()` function, which uses this data source. Filtered to the same Core-account exclusions as the rest of Sheet 2 (REG-008); cross-checked the excluded China total (49.6M SEK combined FY24+FY25) against `config.py`'s own documented figure — matches almost exactly. **Assumption:** the `Marketplace` Sales Channel (Sport-Scheck, About You, Zalando Marketplace, etc.) is grouped into Wholesale, not DTC — not independently confirmed. **Taxonomy, revised same day:** the first version applied the 50,000 SEK/year threshold per channel per year and lumped every failure into one "Insufficient data" bucket (1,555 franchises) — median sales for that bucket was 17,360 SEK (34 units), mostly genuinely small, but 564 of them (36%) had Sales_2025 above the threshold, revealing they were single-channel, not small. Split into: Substitution (DTC↑/Wholesale↓) — 135, the most common pattern; Substitution (Wholesale↑/DTC↓) — 26; Co-growth (both↑) — 66; Co-decline (both↓) — 70; **Wholesale-only / DTC negligible** (Wholesale reliable both years, DTC <50K in *both* years — a structural fact, not a data gap) — 59, incl. 37 Workhorse+Harvest; **DTC-only / Wholesale negligible** (the reverse) — 83, incl. 29 Workhorse+Harvest; Insufficient data (now reserved for genuinely-small-both-channels or a channel with mixed year-to-year reliability) — 1,413, incl. 4 Hero+Near-Hero; Not in FY24/25 dataset — 8. Classification is sign-based with no materiality band beyond the reliability gate. **Cross-workstream caveat (REG-009):** this dataset's totals have a known ~2.3% gap vs. the SS26 exports for the same period; franchise identity uses the same Base+Gender text matching as REG-010/014/015/017. |
| REG-017 | Portfolio Analysis | Wholesale Share % per franchise (DTC x Wholesale channel mix) | **Confirmed** (methodology), figures **Provisional** per REG-012 | New `Wholesale Share % (FY25, raw)` column on Sheet 2: Wholesale channel's share of a franchise's FY25 Core-scope sales (vs. Retail+E-com combined, i.e. DTC), recomputed from the raw SS26 exports at franchise grain — same method/caveat as REG-014's Generation Detail sheet. Company-wide FY25 Core baseline for reference: Wholesale 40.7% GM%, Retail 48.9%, E-com 61.1% (Wholesale lowest-margin). Coverage: 866 of 1,860 franchises have a value; 664 blanked below the 50,000 SEK/year threshold (REG-004 precedent); 330 have no matching raw sales at all (mostly Thin/Immaterial or Exited). **Of the 1,530 franchises with any reconstructed sales, 366 (24%) differ from the published Sales_2025 by more than 15%** — the same unresolved "3-Close out order" ordertype gap as REG-012, which skews toward overstating Wholesale share wherever it's present (close-out activity is wholesale-channel). Read as directional (wholesale-heavy vs. DTC-heavy), not an audited percentage. One outlier: "Asp 3-in-1 GTX Parka" (Women, Exited, 203K SEK FY25) shows a small negative value from returns exceeding gross wholesale sales on a low base — left as computed. |
| REG-016 | Portfolio Analysis | FW27 Collection status (Active/Not active/Not in review) | **Confirmed** (business-directed), field is a **known-imperfect source, kept as-is** | New `FW27 Collection` column on Sheet 2: `Active` / `Not active` / `Not in review`, sourced from the `Active` field in `data/Assortment Attribution Review(F27).xlsx` (Key-Article Margin Analysis workstream), matched by franchise via the same text parsing as REG-010/014/015. `Active` = at least one matching style row is Active=True; `Not active` = every matching row is Active=False; `Not in review` = the franchise isn't in this 392-row file at all (1,568 of 1,860 — this file only covers 292 franchises, so absence ≠ a negative signal). **This repo's own README documents a known false-negative problem** on this field (successor styles clearly still trading shown Active=False) — confirmed still live: `Zircon Slim II Pant` is Active=False in the source file, but the matching franchise `Zircon Slim Pant` (Men) has 546,751 SEK in YTD2026 sales. **Per business direction, the source value is kept as-is, not overridden** — instead, a 24-franchise exception list (`Not active` + YTD2026 sales ≥ 50,000 SEK, REG-004's reliability threshold) is on Sheet 1 for review. Several exceptions (Velum Jacket, Zodiac Jacket, Spacelite, Steep Proof 3L Jacket) show near-zero FY25 sales with substantial YTD2026 sales — likely new SS26 launches postdating the attribution snapshot, a different pattern than the Zircon-style successor-naming anomaly; not distinguished from each other since the source file carries no refresh date to tell them apart. |
| REG-015 | Portfolio Analysis | Core Assortment FW27 inclusion check | **Confirmed** (methodology, **updated 08-Sep-2026 with the authoritative Excel list**) | New `Core Assortment FW27` (Yes/No) column on Sheet 2, matched by franchise (Base+Gender) against **`FW27_CORE_Assortment_Styles.xlsx`** (57 carry-over styles, each with an exact Style Number/Model code) **plus** the 31-style "News package" from the original send-out deck (its own Model-coded table) — 88 items combined. **Supersedes the first version**, which was transcribed from a PDF slide deck's image captions (labeled "version 1" on its own title slide, no Model codes for the carry-over section, a different and less complete ~64-style set). Matched by parsing each item's own style/article name (same gender/version-token method as REG-010/014) — deliberately **not** by looking up whether that Model code has ever appeared in the SS26 raw exports, since several 2.0/II successor styles (e.g. "Mimic Alert 2.0 Hood", "Rosson Mid II Jacket Women") haven't shipped yet and have zero sales history; a sales-history lookup would have wrongly marked their whole franchise absent. Still true from the first version: **this list is a deliberately curated subset, not the full continuing catalog** — absence isn't automatic discontinuation, but is worth a merch-team check for a Hero/Near-Hero performer. Match quality: 57 of 88 items matched an existing published franchise; the 31 unmatched are genuinely new (the News package minus "Mimic Alert," which matched; plus "Korp Softshell Hood/Pant," independently confirmed to have no FY24/25 sales history either way). **Finding: 20 of 34 Hero + Near-Hero franchises (59%) are not in the list** (down from 21) — "Long Down Parka" (Women, 6.3M SEK) dropped off the gap list because the authoritative file confirms its successor "Long Down II Parka Women" is in Core Assortment; listed on Sheet 1 for merch review. **Flowed into the Hero Tier Catalogue (10-Sep-2026)** as a third card badge ("Core FW27: Yes/No") — same 14-of-34 Yes / 20-of-34 No split, read straight from Sheet 2, no new methodology. |
| REG-014 | Portfolio Analysis | Generation Detail sheet (article-naming version split) | **Confirmed** (methodology, revised 2026-09-08), figures **Provisional** per REG-012 | New Sheet 3 in `Project_Bob_Portfolio_Tiering_08092026.xlsx`: one row per distinct Article, for the 109 franchises (of 1,860) where the raw exports show a genuine naming-based version succession (an "Original" article and a later II/III/IV/2.0-suffixed article under the same franchise). **Generation is determined purely by the article name's version token — NOT by the Aug-2026 China JV/Zalando close-out lists.** (First build conflated the two: it bucketed SKUs by close-out-list membership, which produced the same article name appearing in both an "Old-gen" and "Continuing" row whenever only some of that article's colorways were on the list — confusing and wrong, corrected same day.) Close-out-list membership is now shown only as an informational "Clearance Exposure" column, decoupled from the generation label. Franchises with only one continuously-named article aren't included (nothing to compare). Example: "Astral GTX Jacket" (Women) — Original 48.9% GM% vs. Version II 48.8% GM% (comparable); "L.I.M Fuse Pant" (Men) — Original 48.8% GM% vs. Version II 41.7% GM% (the *new* generation is lower-margin, the opposite of what you'd assume). **Caveat:** Sales/GM% recomputed directly from raw SKU-level exports (same Core-scope rule as REG-008), not the published Sheet-2 figures — an article's rows won't necessarily sum to its franchise's audited Sales_2025 (REG-012's unresolved ordertype gap). Read as relative/directional, not an audited SEK split. GM% blanked below 50,000 SEK/year per REG-004's precedent. |

---

## Glossary — Business Term Mappings

Plain-language terms the team uses in conversation or when querying Tier Bench, mapped to the exact column/value they mean — so the term reads the same way on this sheet, in Tier Bench, and in this register.

| Term | Maps to |
|---|---|
| "DTC expansion opportunity" | `Channel Pattern 2025` = `Wholesale-only / DTC negligible` (REG-018) — proven Wholesale demand, DTC below 50,000 SEK in both 2024 and 2025. 59 franchises (37 Workhorse+Harvest, 14 Problem Child); breakdown and top candidates on Sheet 1. |
| "Wholesale expansion opportunity" | The mirror case: `Channel Pattern 2025` = `DTC-only / Wholesale negligible` (REG-018) — proven DTC demand, Wholesale below 50,000 SEK in both years. 83 franchises. |

---

*Seeded 4 Sep 2026 from the SS26 DTC & Wholesale portfolio review. Reconciled against `config.py` 7 Sep 2026. Clearance treatment (REG-010/011/012), tier consolidation (REG-013), Generation Detail sheet (REG-014), Core Assortment FW27 check (REG-015), FW27 Collection status (REG-016), Wholesale Share % (REG-017), and Channel Pattern 2025 (REG-018) added 8 Sep 2026; Core Assortment FW27 (REG-015) updated same day with the authoritative Excel list. Available Stock by Tier (REG-019, Sheet 4) added 8 Sep 2026. Units_2025 (REG-020), Style Code(s) (REG-021), Sales by Country (REG-022, Sheet 5), and region-relative tiering (REG-023, Sheets 6 & 7) added 9 Sep 2026; REG-023 rebuilt same day once the base tiering rule was confirmed (REG-024). Update the live artifact first, then re-sync this file.*
