# Assumptions Register — Inventory Analysis

A new, third workstream (see top-level `CLAUDE.md`), built from a single
warehouse available-stock snapshot: `Available_stock_260825.xlsx`.

**This workstream deliberately reuses definitions from the SS26 Portfolio
Tiering workstream** rather than re-deriving them — cited by ID below and
in `scripts/analyze_inventory.py`. Read `data/ss26_portfolio_tiering/
ASSUMPTIONS_REGISTER.md` for the full reasoning behind each cited entry;
this file only records what's genuinely new here.

**Source file:** not duplicated. This workstream reads the copy already
checked in at `data/ss26_portfolio_tiering/inputs/stock/
Available_stock_260825.xlsx` directly (see REG-INV-001) — the file
uploaded to start this workstream is byte-identical (confirmed via md5).

**Reusable pipeline script:** `scripts/analyze_inventory.py` — reads the
source file, matches to franchise via `ss26_lib.franchise_key()`, joins
tier from the SS26 tiering workbook, and writes a dated output workbook
`Project_Bob_Inventory_Analysis_<DDMMYYYY>.xlsx` (never overwrites a
previous dated cut — same convention as `build_distribution_workbook.py`).
Re-run it any time the source snapshot or the tiering workbook it joins
against is updated.

**Reference style:** cite an entry by ID (e.g. "per REG-INV-004") in any
code, prompt, or report that depends on it.

---

## Register

| ID | Domain | Field / item | Status | Rule & rationale |
|---|---|---|---|---|
| REG-INV-001 | Data Hygiene | Source file identity | **Confirmed** | The file used to start this workstream is byte-identical (md5 `20a1c473408f8cad42f29143ed88ff1e`) to `data/ss26_portfolio_tiering/inputs/stock/Available_stock_260825.xlsx`, already checked in for REG-019 (SS26 register). Not duplicated here — `scripts/analyze_inventory.py` reads that copy directly via a relative path. If a future data drop supersedes it, land the new file under that same `inputs/stock/` folder per the SS26 workstream's own intake convention (`scripts/README.md` there), and update this workstream's `--source` default alongside REG-019's. |
| REG-INV-002 | Data Hygiene | Article field trailing whitespace | **Confirmed** | Same class of issue as REG-001 (SS26 register): 737 of 7,568 stock rows (9.7%) have a padded `Article` value (e.g. `"Furudal Mimic ParkaWomen            "`). `ss26_lib.franchise_key()` strips this before matching (it calls `.strip()` internally via `base_name()`), so no separate fix was needed here, but it's the same underlying export-tool behavior REG-001 documents — cited rather than re-diagnosed. |
| REG-INV-003 | Portfolio Analysis | Franchise matching (Base+Gender) | **Confirmed** (methodology, reused as-is) | Reuses `ss26_lib.franchise_key()` un-modified (REG-010/014/015/019's method — gender/version-token stripping from `config.py`'s `GENDER_TOKENS`/`VERSION_TOKENS`). 513,640 of 527,963 units (97.3%) matched an existing published franchise in the SS26 tiering workbook — identical match rate to REG-019, as expected since it's the same source file and matching code. See REG-INV-007 for what the unmatched 2.7% looks like. |
| REG-INV-004 | Portfolio Analysis | Core-account scope (REG-008) | **Confirmed — not applicable to this file directly** | This stock file has no account/channel field at all (it's a warehouse snapshot, not a sales export), so REG-008's Core-account exclusion (XXL, China, Zalando, Zalando Marketplace, Stadium Outlet) cannot be and is not applied to its rows. Every FY25 Sales figure this workstream joins in (Sheet 4's detail table) comes from the SS26 tiering workbook's own `Sales_2025` column, which is already Core-scoped per REG-008 upstream — no double-handling needed, just inherited. |
| REG-INV-005 | Portfolio Analysis | Reliability threshold (REG-004) | **Confirmed — applies only to joined figures, not to stock units** | `MIN_RELIABLE` (50,000 SEK/year, REG-004's precedent, reused via `ss26_lib.MIN_RELIABLE` elsewhere in the SS26 workstream) blanks GM%/growth/share figures built on a small sales denominator. `Available stock` itself is a unit count, not a rate — the threshold doesn't apply to it directly. It matters here only because Sheet 4's FY25 Sales column is pulled from the tiering workbook, where that rule (and REG-012's caveat) already governs it — read those figures as directional, not audited. |
| REG-INV-006 | Portfolio Analysis | Stock methodology & caveats (REG-019) | **Confirmed (inherited)** | Same file, same core caveat as REG-019: **units only**. The source file carries no cost/price column, so no SEK stock value is computed here either — don't multiply by an average price without checking with Finance. `Layer` field is used as-is: its 19 distinct values already match `config.ALL_LAYERS`' canonical post-fix casing exactly (e.g. "Mid layer", "Hiking shoes") — no `config.LAYER_NAME_FIXES` casing variant appears in this file, consistent with the SS26 register's note that the tiering file's own `Layer` column is independently clean. One data-quality wrinkle carried over from the source, not from REG-019: 5,628 units (1.1%) carry `Layer = "*MISSING*"` — the same placeholder text flagged as a parsing edge case in REG-024 (SS26 register), here appearing directly in the source file's own field rather than from a matching failure. Not resolved — flagged for a data-owner check. |
| REG-INV-007 | Portfolio Analysis | Unmatched stock (not in the 1,860-franchise tiering file) | **Open** | 14,323 units (2.7%) across 27 distinct franchises don't match any row in the published SS26 tiering workbook — identical total to REG-019's own figure (same source file, same match code, so this is expected, not a new finding). Breakdown by Product Area: 10,764 Apparel, 3,364 Hardware, 195 Footwear. Two classes visible in the top 10 by units: (a) clearly out-of-scope non-apparel items (`LO Packing cubes`, `LO Wallets` — accessories the 1,860-franchise tiering universe was never built to cover), and (b) apparel-looking names not yet in that universe (`Klyva GTX Jacket`, `Kyligt Softshell Pant`, `Pollux 1/2 zip Jacket`, `Zircon Shorts`, `Kaja Proof Jacket`) — consistent with REG-019's stated explanation (brand-new FW27 launches with no FY24/25 sales history to tier against yet), but **not independently confirmed** as all being genuinely new vs. a residual `franchise_key()` naming miss on an existing franchise. Worth a merch/data-owner check before assuming this list is exhaustively "new launches." |
| REG-INV-008 | Portfolio Analysis | Season-Recency buckets (`Season Article` field) | **Provisional** | New to this workstream — the source file's `Season Article` field (e.g. `F26`, `S27`) was present but unused in REG-019's Stock-by-Tier build. Distinct values span `S22`…`S27` (11 codes), read as `<S\|F><YY>` = Spring/Summer or Fall/Winter of that year — consistent with this repo's own `SS26` naming (`S26` = the same "SS26" season this repo's other workstream is named for). **Bucket definition (not yet confirmed with the business):** `Incoming / Future` = the two newest codes present, `F26` and `S27` (436,625 units, 82.7% of total stock) — pre-season stock already in the warehouse ahead of that season's main selling window; `Current season` = `S26` alone (34,056 units, 6.5%); `Aged` = everything else, `F25` and older (57,282 units, 10.9%). **Open sub-item:** the snapshot's own filename, `Available_stock_260825`, doesn't unambiguously encode its pull date — could be read as DDMMYY (26-Aug-2025) or YYMMDD (2026-08-25); the latter is far more consistent with the season mix actually observed (F26/S27 dominating implies a pull shortly before the F26 season starts shipping, i.e. mid/late-2026, not 2025) and with REG-019 already using this same file on 8-Sep-2026, but neither reading is confirmed with whoever exported it. If the pull date is materially different from "late Aug 2026," the Current/Future/Aged cutoffs above should be revisited. |
| REG-INV-009 | Portfolio Analysis | Aged stock in clearance-relevant tiers (Sheet 4 of the output workbook) | **Provisional** (depends on REG-INV-008) | Cross-tabbing Season-Recency (REG-INV-008) against Tier: within the 57,282-unit Aged bucket, 2,858 units sit in Thin/Immaterial and 447 in Exited — REG-019 already calls both tiers "clear via sale" on tier grounds alone, so an old season code on top is a second, independent signal pointing the same way (doubly-confirmed candidates, not a new claim). Also present: 9,546 Aged units in Problem Child and 31,483 in Workhorse+Harvest — per REG-019's own framing these tiers are NOT clearance candidates (Problem Child needs a pricing/cost fix; Workhorse+Harvest is a real active seller), so an old season code there is presented as a merch-check flag only ("is this a superseded colorway sitting on top of an otherwise-healthy franchise, or just normal replenishment tail?"), not a recommendation to clear it — that determination isn't made here. |

---

## Glossary

| Term | Maps to |
|---|---|
| "Aged stock" | `Season Article` = `F25` or older per REG-INV-008 — **not** the same thing as REG-019's "clear via sale" tiers (Thin/Immaterial, Exited); the two overlap partially (REG-INV-009) but are independent signals. |
| "Incoming / Future stock" | `Season Article` = `F26` or `S27` per REG-INV-008 — pre-season stock, not a risk signal by itself. |

---

*Seeded 10 Sep 2026, starting the Inventory Analysis workstream from the
`Available_stock_260825.xlsx` warehouse snapshot (same file as SS26
Portfolio Tiering's REG-019). REG-INV-001 through REG-INV-007 restate/
apply SS26-register definitions (REG-001, REG-004, REG-008, REG-010/014/
015, REG-012, REG-019, REG-024) to this file; REG-INV-008/009 (Season-
Recency) are new to this workstream and Provisional pending business
confirmation.*
