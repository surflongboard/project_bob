# Inventory Analysis

A third workstream (see top-level `CLAUDE.md`), analyzing warehouse
available stock — specifically the Season-Recency ("how old is this
stock?") view that the SS26 Portfolio Tiering workstream's own stock sheet
(REG-019 there) doesn't cover.

**Start here:** `ASSUMPTIONS_REGISTER.md` — every definition used below is
either cited from the SS26 register or documented fresh as REG-INV-###.

## What's in this folder

- `ASSUMPTIONS_REGISTER.md` — the register (Confirmed/Provisional/Open
  discipline, same as the SS26 workstream).
- `scripts/analyze_inventory.py` — reusable, re-runnable build script.
  Reads the warehouse stock snapshot already checked in at
  `data/ss26_portfolio_tiering/inputs/stock/Available_stock_260825.xlsx`
  (not duplicated here — see REG-INV-001), matches to franchise via
  `ss26_lib.franchise_key()`, joins each franchise's current Tier from the
  SS26 tiering workbook, and writes a dated output workbook.
- `Project_Bob_Inventory_Analysis_<DDMMYYYY>.xlsx` — the output: Overview,
  Stock by Season Recency x Tier, Stock by Layer, Aged Stock Detail, and
  Unmatched Stock. Re-run the script for a fresh dated cut rather than
  editing a workbook by hand.
- `inputs/assortment/` — the FW27 range assortment plan
  (`Assortment Attribution Review_11092026.xlsx`, from Google Drive),
  kept under its original filename per the SS26 workstream's own intake
  convention.
- `scripts/build_exit_plan_priority.py` — joins the FW27 tab's own planned
  exit season (`LSO / Exit Season`, column AI) to current stock, so a
  franchise already scheduled to leave the collection surfaces before it's
  discovered as dead stock. A forward-looking complement to
  `analyze_inventory.py`'s backward-looking Season-Recency view — see
  REG-INV-010.
- `Project_Bob_Sell_Down_Priority_<DDMMYYYY>.xlsx` — the output: Overview,
  ranked Sell-Down Priority, Pending Review (exit timing not yet decided —
  deliberately not ranked), stock not on the FW27 plan at all (cross-checked
  against SS27 — REG-INV-012), a heuristic rename-candidate list for
  that cross-check's still-Active-in-SS27 subset (REG-INV-013), and an
  inventory value estimate — Retail/Wholesale/Est. Cost, with a summary
  by Priority Tier and by Activity (REG-INV-014). Landed cost is
  reconstructed (`WP x (1-GM0%)`) since the plan's own cost field is
  broken — see REG-INV-014 for the validation and caveats. Sheet 4 gets
  a separate fallback value (Historical ASP from actual FY25 sales in
  the SS26 tiering workbook) covering 79% of its otherwise-unpriced
  units — see REG-INV-015; it's a different basis from REG-INV-014's
  Target-RRP value, not directly comparable to it.

## Running the script

From the repo root:

```bash
python3 data/inventory_analysis/scripts/analyze_inventory.py
```

Prints a summary (total/matched units, per-bucket totals) before saving —
read it before trusting the output. Pass `--source PATH` to point at a
future stock snapshot once one exists; land the raw file under the SS26
workstream's own `inputs/stock/` folder first (its intake convention), then
add a corresponding register entry here.

## Relationship to the SS26 Portfolio Tiering workstream

Same source file, same franchise-matching code, same tiering workbook —
this is a genuinely separate deliverable (a different question: stock
aging, not portfolio sizing), not a fork of REG-019's logic. See the
register's header for exactly which SS26-register entries are reused
as-is vs. what's new here.
