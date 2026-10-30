# Assumptions Register — Inventory Planning

A new, separate workstream carved out of **FW27/SS28 Range Review**
(`data/range_review/`) on 30-Oct-2026 — same underlying stock/purchase
data, different question: not "which franchise gets what Tier," but
"what do we do with the inventory and generational tail itself" — SS27
stock position, end-of-life timing, and per-franchise generation
history, as a dedicated planning topic in its own right rather than a
side-concern bolted onto the Range Review Board.

**Why this was split out:** all five entries below (IP-001 through
IP-005) started life as RR-018/019/020/024/025 inside Range Review's
own register. That register still has short pointer rows at those IDs
— this file is now the authoritative source for their full content;
don't let the two drift apart by editing one without the other.

**What moved, concretely, on 30-Oct-2026:**
- `inputs/stock/Stock_and_purchase_261029.xlsx` (the Inventory
  Manager's stock/purchase file) — was `data/range_review/inputs/stock/`.
- `scripts/build_inventory_review.py` and
  `scripts/build_generation_timeline.py` — were in
  `data/range_review/scripts/`. Both still import `build_franchise_list`
  (Range Review) and `ss26_lib` (SS26 Portfolio Tiering) cross-workstream,
  the same pattern this repo already uses elsewhere — not a reason to
  re-derive franchise identity or token-stripping logic here.
- `outputs/SS27_Inventory_Review.xlsx` and
  `outputs/Franchise_Generation_Timeline.xlsx` — were in
  `data/range_review/outputs/`.

**What did NOT move, and why:** the Range Review Board
(`data/range_review/scripts/workshop_page_template.html` +
`build_workshop_data.py`) still shows inventory-derived signals —
the "Non-active item overstocks" and "Generation timeline" flags/
alerts on each franchise card. `build_workshop_data.py` now imports
this workstream's `build_generation_timeline.py` cross-workstream and
reads `SS27_Inventory_Review.xlsx` from this new location directly.
**If you change this workstream's methodology, re-run
`build_inventory_review.py`/`build_generation_timeline.py` here, then
re-run Range Review's `build_workshop_data.py` and republish the
board** — the two are linked by a live file dependency, not a one-time
copy.

**Source data:** `inputs/stock/Stock_and_purchase_261029.xlsx` —
SKU-level (Article No. = 6-digit style Code + 3-char Color Code +
3-digit Size Code) stock/purchase/orderbook data from the Inventory
Manager, received 29-Oct-2026 into the Range Review workstream, now
owned here. 4 sheets: Physical stock, Purchase SS27 (undelivered),
Available stock (the headline, already-netted figure), Orderbook. See
IP-001 for the Inventory Manager's own caveats on what "Available
stock" does and doesn't net out.

Franchise identity (which active-range franchise a SKU's Code belongs
to, when it's in the active range) comes from Range Review's own
`build_franchise_list.py` — don't re-derive it here. Franchise
attribution for codes NOT in the active range (old generations, fully
discontinued lines) is derived directly from the stock file's own
Article text via `base_name()`/`version_token()` (`ss26_lib`, SS26
Portfolio Tiering) — see IP-002.

**Status legend** (same discipline as every other register in this
repo):
- **Confirmed** — a fact with one right answer, verified against a source
  or business confirmation. Safe to build on.
- **Provisional** — a reasonable methodology choice, not yet independently
  confirmed. Usable, but flag it if it feeds something high-stakes.
- **Open** — a genuine unresolved question. Don't silently resolve it by
  picking an assumption in code; either ask, or build around it and say so.

**Pending follow-up work** (handed off from Range Review's own "Next
steps," 30-Oct-2026 — see that register's entry 8 for the original
framing):
1. Review the **16 active multi-generation franchises** (IP-005 /
   Range Review Board "Generation timeline" flag) to decide an
   explicit update-cycle policy per franchise.
2. Confirm the **17 non-active multi-generation franchises** (Astral
   GTX Pant, Bow Windstopper Glove, Chaos GTX Pant, Gabbro Pant,
   Haglöfs Krusa GTX Low, Husk Jacket, Husk Pant, L.I.M Alpha Hood,
   L.I.M Down Jacket, L.I.M Fuse Pant, L.I.M Fuse Shorts, L.I.M Mid
   Fast Hood, L.I.M Mimic Hood, Puffy Mimic Hood, Riksgränsen Down 850
   Hood, Scree Pant, Zodiac Mid Jacket) are indeed being discontinued
   at end of SS27 per IP-003's EOL cutoff — several (L.I.M Fuse
   Pant/Shorts, Haglöfs Krusa GTX Low, Scree Pant) show stock as late
   as S27/F27, worth a sanity-check that this is intentional range
   narrowing and not a data gap.
3. **Reconcile the SS27 "Available stock" headline figure** against
   SS27 DTC demand and FW27 allocations taken from current stock —
   per IP-001, the 445,528-unit headline likely overstates what's
   truly free for SS27, not yet by a known amount. The user plans to
   request an SS27 forecast separately to refine this.

---

| ID | Domain | Field | Status | Rule | Rationale |
|---|---|---|---|---|---|
| IP-001 | Portfolio Analysis | SS27 inventory data ("List 2") — Inventory Manager's scope caveats on "Available stock" | **Confirmed** (business-directed, 29-Oct-2026 — verbatim from the Inventory Manager) | `inputs/stock/Stock_and_purchase_261029.xlsx` has 4 sheets: Physical stock (615,297 units total), Purchase SS27 — undelivered (527,277 units), Available stock (445,528 units, the headline figure), Orderbook (378,351 units). Per the Inventory Manager's own comment, verbatim: **"in available stock the SS27 purchase Is included"**; **"All orders in the orderbook is deducted from the physical stock and purchase in that view"**; **"The orderbook is only on 'hard' orders so no DTC booking for SS27 is included"**; **"Everything we book from stock for FW27 is not in the orderbook either."** In plain terms: `Available stock = Physical stock + Purchase SS27 − Orderbook (wholesale hard orders only)`, and it is **NOT** further netted against SS27 DTC demand or FW27 allocations taken from current stock — both still need to come out of it, so **the 445,528 headline figure likely overstates what's truly free for SS27**, not yet by a known amount. Shown as-is per user direction — not adjusted — with this caveat carried into the output. Article No. (SKU) decodes as 6-digit style Code + 3-char Color Code + 3-digit Size Code — matches the Range Review franchise list's Code field directly, no re-derivation needed. A `Season Article` field on the Available-stock sheet tags each SKU's original season back to S22/F22 — useful for spotting genuinely old/dead stock, not just this-cycle leftovers, and is the basis for IP-005's generation timeline. | Received 29-Oct-2026 (into what was then the Range Review workstream; this workstream now owns it). User plans to separately request an SS27 forecast to refine the true-availability picture later — not pursued here. Originally RR-018. |
| IP-002 | Portfolio Analysis | SS27 Inventory Review build — SKU-grain calculation, franchise-grain summary | **Provisional** — user-directed 29-Oct-2026 (SKU grain for calculation, franchise rollup for summary) | `scripts/build_inventory_review.py` builds the List 2/List 3 view at **SKU grain** (Article No., per user direction — "otherwise not accurate"), joining all 4 sheets by Article No. (union of their populations, not just Available stock's, so a SKU with e.g. only an open order still appears). Rolled up into a **Franchise Summary** sheet for readability. **Franchise attribution for non-Active codes**: initially every non-Active SKU fell into an unusable "(no franchise match)" bucket (144,756 units, 100% of the EOL candidates) because the lookup only covered styles Active in the current range. Fixed by parsing a franchise name directly from the stock file's own Article text (same `base_name()` stripping used everywhere else) when the Code isn't in the active range — since version tokens are stripped either way, an old-generation code's parsed base matches its replacement's base, so old leftover stock of a franchise you're KEEPING correctly rolls up under that same franchise name (e.g. "Rosson Mid Hood": Active=Y, 8,214 units of old-code stock still attributed to it) rather than vanishing into an unmapped bucket. **This rollup is correct but was opaque on the Range Review Board** — the user flagged not understanding why an active franchise ("Mimic Alert 2.0 Hood") showed a 6,903-unit tail alert when its own current codes carry none. Fixed 29-Oct-2026: Range Review's `build_workshop_data.py`'s `load_tail_styles_by_base()` breaks the tail total down by the actual non-active style code(s) feeding it (here: "Mimic Alert Hood" — no "2.0" — codes 607444 Men/3,651 units + 607447 Women/3,252 units, the pre-"2.0" generation), and the card's inventory-tail alert now names them, the same way the MOQ alert names its flagged style. | Built 29-Oct-2026. Result: **3,883 SKUs (512 distinct style Codes) not in the active range, 144,756 of 445,528 Available Stock units (33%)** — reconciles exactly against the file's own 445,528 total. Of those 512 codes, 435 carry positive Available Stock; the other 77 have zero Available Stock but still show up via Physical/Purchase/Orderbook presence. **Note:** the user's own estimate going in was "~425 articles" — closer to the 435-with-stock figure than the 512-total-codes figure, but not an exact match either way; worth a quick sanity check against whatever the user was recalling, rather than assuming either number is "right." Originally RR-019. |
| IP-003 | Portfolio Analysis | EOL cutoff season = SS27 | **Confirmed** (user-directed, 29-Oct-2026 — "cut off is SS27, my mistake") | The 512 non-active-range style codes / 144,756 Available Stock units in `SS27_Inventory_Review.xlsx` are end-of-life as of **end of SS27** — consistent with the rest of this topic (the original Nov 12 prep email's "List 2/List 3," "by end of SS27" in the sent email to Tobias). The earlier "end of SS26" reference was confirmed a typo, not a real scope change. | Flagged 29-Oct-2026 as likely a repeat of an earlier SS26/SS27 slip elsewhere in this business's analysis; confirmed same day. Originally RR-020. |
| IP-004 | Portfolio Analysis | Inventory tail isn't always an OLDER generation — some is a retired code of the SAME generation | **Confirmed** (methodology fix, 29-Oct-2026) | User asked to check all 49 flagged franchises (on the Range Review Board at the time) for the same situation as "Mimic Alert 2.0 Hood." Found: 43 are genuine prior-generation stock (IP-002's intended case — e.g. "Mimic Alert Hood" rolling up under current "Mimic Alert 2.0 Hood"). **But 11 tail-style lines across 6 franchises (Mid Standard II Shorts, Spitz GTX Pro II Jacket, Vassi GTX II Jacket, Vassi GTX II Pant, Vassi GTX Pro II Bib, Vassi GTX Pro II Jacket) carry the SAME version token as the active franchise itself** — e.g. "Vassi GTX II Jacket" (active, current codes 608574/608578) has tail stock under codes 607791/607793, both still named "Vassi GTX II Jacket." Confirmed in the range files these old codes simply aren't listed in the current range (last season's SKU for the same "II" style, superseded by a new code) — not a predecessor model at all. Calling this "old generation" would mislead the team. Range Review's `build_workshop_data.py` tags each `tailStyles` entry with `sameGeneration` (comparing `ss26_lib.version_token()` of the tail style's name against the active franchise's own displayed name); the card's alert wording splits accordingly — "old generation in stock" for genuine predecessors, "a retired code of this same style" for same-generation leftovers, and both named with their respective totals when a franchise has a mix (4 franchises: Mid Standard II Shorts, Vassi GTX II Jacket, Vassi GTX II Pant, Vassi GTX Pro II Jacket). | Found 29-Oct-2026 checking all 49 flagged franchises per the user's request, not just the one already investigated. 2 franchises (Spitz GTX Pro II Jacket, Vassi GTX Pro II Bib) are pure same-generation with no genuine older stock at all. Originally RR-024. |
| IP-005 | Portfolio Analysis | Generation-by-generation timeline per franchise, for the lifecycle discussion | **Provisional** — methodology is reasonable but each date is an inferred proxy, not a confirmed source | User asked for a list of products with their generations and when each was introduced, for the Nov 12 lifecycle discussion. **No "launch date"/"introduced" field exists in any source this business's data has** — checked exhaustively, the same way Range Review's RR-021 checked for "last season to sell." Best available proxy: each SKU's own "Season Article" tag in the SS27 stock file (S22 through F27 observed; 31% of rows blank) — the earliest season seen across a generation's style Code(s) is used as an approximate "existed by at least this season" marker, built by `scripts/build_generation_timeline.py` → `outputs/Franchise_Generation_Timeline.xlsx`. **Two real limitations, not hidden**: (1) only SKUs still present in this one stock/purchase snapshot show up at all — a generation that sold through completely with zero remaining stock anywhere in the file is invisible, so this is a floor on history, not a complete one; (2) a season tag can reflect a later re-buy of an existing style rather than its true first season. Generations are grouped by `ss26_lib.version_token()` (no token / "II"/"2.0" / "III" / "IV"), with Men/Women/Asia-Fit codes of the same generation collapsed into one row (the raw per-style-code grouping inflated "generation count" by counting gender variants as separate generations — fixed before this was presented). The no-token generation is labeled **"Generation 1"** only within a franchise that actually has 2+ generations — a single-generation franchise's sole style stays unlabeled, since there's no sequence to number it against (fixed 30-Oct-2026 after the first pass mislabeled all 488 franchises in the stock file, not just the 33 multi-generation ones). **Surfaced on the Range Review Board itself**: `build_workshop_data.py`'s `load_generation_timeline_by_base()` reuses this script's own `load_generation_groups`/`aggregate_by_generation` functions (not a re-implementation, imported cross-workstream after this script moved here) to attach a `generationTimeline` list to the 16 of these 33 franchises that are also currently Active (the other 17 are fully discontinued, with no active-range successor, so they're not on the board at all — see this register's "Pending follow-up work" above for the full list). A card with one shows a "Generation timeline" alert naming each generation, and a matching "Generation timeline" flag-only toggle filters to just those cards. | Found 29-Oct-2026. **33 of 488 franchises in the stock file show 2+ distinct generations**; one (Puffy Mimic Hood) shows 3 ("Generation 1" F24, "II" F25, "III" F26). The season-to-season gaps in the output (e.g. Front Proof Jacket: "Generation 1" earliest F24 → "II" earliest S26, ~2 seasons) are directionally plausible as a refresh cadence but not independently verified — flag any specific date to the team before treating it as settled. Originally RR-025; moved to this workstream and the "Generation 1" labeling bug fixed same day, 30-Oct-2026. |

---

*Carved out of `data/range_review/ASSUMPTIONS_REGISTER.md` on
30-Oct-2026. Cross-reference that file for the Range Review Board's own
UI (Tier/Action/Last season to sell/Consolidate decisions, the three
flag-only toggles, the Summary panel) and
`data/ss26_portfolio_tiering/ASSUMPTIONS_REGISTER.md` for the shared
franchise-matching methodology (`base_name`/`gender`/`version_token`)
both workstreams import rather than redefine.*
