# Assumptions Register — FW27/SS28 Range Review

A new, separate workstream from both **Key-Article Margin Analysis** (repo
root) and **SS26 Portfolio Tiering** (`data/ss26_portfolio_tiering/`) — same
business, different question: narrowing the active range (FW27 + SS28) down
from existing + incoming SS27 inventory, tiering what survives (Hero /
Harvest / Emerging / Exit), and preparing a layer-by-layer keep/consolidate/
improve/kill proposal for the Nov 12 Range Review workshop.

**Reuses, rather than re-derives:** franchise-key parsing (`base_name`/
`gender`/`version_token`) from `config.py` and
`data/ss26_portfolio_tiering/scripts/ss26_lib.py` — confirmed working
directly against this workstream's own data (see RR-002/RR-003 below). Don't
assume every SS26-workstream rule carries over unchanged, though — this
register exists because several genuinely don't (see RR-001, RR-003).

**Source data:** `inputs/FW27_28_Oct_2026.xlsx` and
`inputs/SS28_28_Oct_2026.xlsx` — the FW27 and SS28 range files supplied by
Tobias (Product Controller), received 28-Oct-2026. Style-level rows (Code,
Style name, Gender, Category/Layer, Active, Carry Over, Collection Type, ABS
Customer, ...) with nested colorway sub-rows underneath each style.

**Status legend** (same discipline as the SS26 register):
- **Confirmed** — a fact with one right answer, verified against a source or
  business confirmation. Safe to build on.
- **Provisional** — a reasonable methodology choice, not yet independently
  confirmed. Usable, but flag it if it feeds something high-stakes.
- **Open** — a genuine unresolved question. Don't silently resolve it by
  picking an assumption in code; either ask, or build around it and say so.

---

| ID | Domain | Field | Status | Rule | Rationale |
|---|---|---|---|---|---|
| RR-001 | Portfolio Analysis | Franchise identity (gender separation) | **Confirmed** (business-directed, 28-Oct-2026) | Franchise identity for this workstream is **Base name only** — gender is NOT a separate split, to reduce the number of decisions needed for the Nov 12 workshop. Where a Base name exists in only one gender, that's noted as an attribute on the franchise (e.g. "Women only"), not treated as a defining part of its identity. | Differs deliberately from SS26 Portfolio Tiering, where franchise identity is Base + Gender (REG-013/REG-024 there). Checked before deciding: of 1,241 existing Base names, 571 (46%) have both Men and Women versions under separate style codes — collapsing those loses the gender split for roughly half the catalog, a real simplification, not a free one. 401 of the remaining single-gender bases are Unisex (no gender split existed anyway); 244 are genuinely single-gender apparel (134 Men-only, 110 Women-only). |
| RR-002 | Portfolio Analysis | Range-membership filter (Active flag) | **Confirmed** (business-directed, 28-Oct-2026) | Use each style's own `Active` flag in the FW27/SS28 range files as the range-membership signal — not mere presence in the file. A franchise absent from the Active set in both FW27 and SS28 is the end-of-life candidate population (List 3). | Evidence from the FW27/SS28 files themselves: "Alert GTX Jacket Men" (old code 608172) appears in both files but is marked Active=False in SS28, while its successor "Alert GTX II Jacket Men" (new code 609012) is Active=True — the same pattern holds for "Hede Cargo Pant" and "Hede Shorts" (Men/Women). Confirms Active, not row presence, is the right signal. |
| RR-003 | Portfolio Analysis | Representative style name (generation display) | **Provisional** — user-directed 28-Oct-2026, explicitly flagged for reconsideration | Each franchise is represented by its **latest active generation's actual style name** (gender-stripped), not a generic version-stripped base — e.g. once "Alert GTX II Jacket" is the Active generation, the franchise displays as "Alert GTX II Jacket," not a generic "Alert GTX Jacket." Where multiple generations are Active at once across genders (e.g. Men still Original, Women already on II), the higher-ranked (most recent) generation's name wins as the shared label. | Deliberately different from SS26 Portfolio Tiering's Base-name convention (REG-014/`config.py`), which strips version tokens entirely so "Astral GTX Jacket" represents every generation, with version tracked separately on Sheet 3. For this workstream the user wants the generation visible in the name itself at a glance, since old-generation style codes are already filtered out by RR-002 (Active=False). **Explicitly flagged by the user as a call to revisit, not a settled convention** — confirm before this feeds anything beyond the Nov 12 workshop. |
| RR-004 | Data Hygiene | "Hede Shorts" inconsistent version-token placement | Open | Two different SS28 style codes for the same franchise use "II" in different positions: "Hede II Shorts Men" (609006, Active=True) vs. "Hede Shorts II Men" (609010, Active=False). Both parse to the same franchise under our matching, so this doesn't currently cause a wrong result — flagged for Tobias/Isabelle as a naming-convention inconsistency worth raising, not a blocking issue. | Found 28-Oct-2026 while validating franchise matching against the FW27/SS28 range files. |
| RR-005 | Portfolio Analysis | "Blueberry"-prefixed styles = special editions | **Provisional** — user-directed 28-Oct-2026, to confirm with Tobias/Isabelle next week | The 8 SS28 styles prefixed "Blueberry " (e.g. "Blueberry Tight Medium") are read as a **special edition of the existing franchise**, not a new or distinct item — noted on the franchise overview as a special-edition flag (e.g. "Tight Medium — special edition: Blueberry"), rather than folded silently into the base franchise or counted as new. **Out of scope for the main Nov 12 range review** — these don't get their own keep/consolidate/improve/kill call. | Stripping "Blueberry " matches an existing franchise for 6 of 8 (the other 2, both "L.I.M Mimic Barrier Hood," don't match either way — still worth a direct check). Found 28-Oct-2026 cross-referencing FW27/SS28 Style names against the existing tiering workbook. |
| RR-006 | Portfolio Analysis | Account-exclusive (ABS/ABS Exclusive) styles — scope for the layer review | Open — to verify with Tobias/Isabelle next week | ~50–58 FW27 and ~21–28 SS28 styles are tagged Collection Type = ABS or ABS Exclusive, with an ABS Customer field naming the account (Intersport, XXL, ASIA, DTC, "KA YO Collab," "Team Sportia, Eventyr Sport"). Undecided whether to include these in the general layer-by-layer keep/consolidate/improve/kill review, or scope them out separately since low/narrow sales there is by design, not a performance signal. See RR-007 for the ASIA/"AF" subset specifically. | Same FW27/SS28 upload (28-Oct-2026); this field also answers the earlier "built specifically for a Key account" question without needing a separate ask of Tobias. |
| RR-007 | Portfolio Analysis | "AF"-suffixed styles = Asian Fit assortment | **Provisional** — user-directed 28-Oct-2026, to confirm with Tobias/Isabelle next week | The 24 FW27 styles suffixed " AF" (all `ABS Customer = ASIA`) are read as **"Asian Fit"** — a parallel, fit-adjusted assortment line running alongside the Global assortment, not a regional sales-channel exclusive. **Global assortment vs. Asian Fit assortment is a separate topic**, out of scope for the main Nov 12 range review — these don't get their own keep/consolidate/improve/kill call there either. | Stripping " AF" matches an existing franchise by name for about half of the 24 (the rest don't match either way). Found 28-Oct-2026 cross-referencing FW27/SS28 Style names against the existing tiering workbook. |

---

*Seeded 28 Oct 2026, following the FW27/SS28 range data arriving from
Tobias. Cross-reference `data/ss26_portfolio_tiering/ASSUMPTIONS_REGISTER.md`
for the shared franchise-matching methodology this workstream reuses rather
than redefines.*
