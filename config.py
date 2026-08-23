"""
Project Bob — Key-Article Margin Analysis
All analytical ASSUMPTIONS live here, in one place, so they're visible and easy
to revisit. Nothing in analysis.py should hardcode a judgment call — it should
read from this file. If a finding changes because someone edits a constant
below, that's the point.
"""

# ---------------------------------------------------------------------------
# 1. Layer name normalization
# ---------------------------------------------------------------------------
# The raw sales files have inconsistent casing for a handful of layers
# (e.g. "Mid layer" vs "mid layer"). This map fixes that before anything
# else runs. Re-check this against new data drops — if a new casing variant
# appears, sales for that layer will silently split into two buckets.
LAYER_NAME_FIXES = {
    "mid layer": "Mid layer",
    "Hiking Shoes": "Hiking shoes",
    "Trekking Shoes": "Trekking shoes",
    "Mountain Packs": "Mountain packs",
    "Trail Running Shoes": "Trail running shoes",
}

# The 16 layers the brief asks us to cover (canonical names, post-fix).
ALL_LAYERS = [
    "Shell", "Insulation", "Mid layer", "Legwear", "Soft shell", "Tops",
    "Daypacks", "Accessories", "Hiking shoes", "Hiking packs", "Bags",
    "Sleepingbags", "Base layer", "Trekking shoes", "Trekking packs",
    "Mountain packs", "Trail running shoes", "Casual shoes", "Mountain shoes",
]
# Note: some of the above are small/adjacent categories not in the brief's
# original 16-name list (e.g. Casual shoes, Mountain shoes) — worth confirming
# with whoever owns the layer taxonomy whether these should be folded into a
# neighboring layer or reported separately.

# ---------------------------------------------------------------------------
# 2. Article lifecycle classification (NEW/RAMP, CONTINUING, DISCONTINUED)
# ---------------------------------------------------------------------------
# An article is classified by comparing FY24 vs FY25 units against this
# threshold. Below the threshold, a year's sales are treated as noise
# (returns, sample sales, etc.) rather than a real trading base — this
# matters a lot for margin-% comparisons, which get wild on tiny denominators.
MIN_UNITS_FOR_REAL_TRADING_YEAR = 100

def classify_lifecycle(units_fy24: float, units_fy25: float) -> str:
    """Returns NEW/RAMP, DISCONTINUED, CONTINUING, or MINOR."""
    real_24 = units_fy24 >= MIN_UNITS_FOR_REAL_TRADING_YEAR
    real_25 = units_fy25 >= MIN_UNITS_FOR_REAL_TRADING_YEAR
    if real_25 and not real_24:
        return "NEW/RAMP"
    if real_24 and not real_25:
        return "DISCONTINUED"
    if real_24 and real_25:
        return "CONTINUING"
    return "MINOR"

# A margin-% figure computed on a tiny sales base is noise, not signal --
# small numbers of returns/samples can produce GM% figures like -650% or
# +290% that mean nothing. Below this SEK threshold, treat GM% as unreliable
# (analysis.py blanks it out rather than showing a wild number). This does
# NOT affect Sales/Units totals, only the margin-% column.
MIN_SALES_FOR_RELIABLE_MARGIN_PCT = 50_000  # SEK, per article per year


# ---------------------------------------------------------------------------
# 3. Version-family detection (v1 -> v2/v3 successions, e.g. "Astral GTX
#    Jacket Men" -> "Astral GTX II Jacket Men")
# ---------------------------------------------------------------------------
# Strategy: strip gender words and roman-numeral / "2.0" version tokens from
# an article name to get a "base" family name; articles sharing a base+gender
# are treated as the same style family across versions.
#
# ASSUMPTION WORTH REVISITING: this is a pure string-matching heuristic. It
# will miss version transitions that involve a genuine name change (e.g. if
# "Zircon Shorts" were renamed "Summit Shorts" for FW27, this won't catch it —
# that would need a REF_CODE / franchise mapping from the assortment file
# instead, which is more reliable but requires the PDF to be parsed cleanly).
VERSION_TOKENS = {"2.0", "II", "III", "IV"}
GENDER_TOKENS = {"Men", "Women", "Unisex", "Junior"}

MIN_FAMILY_SALES_FY25 = 200_000  # SEK — below this, don't bother surfacing a "family"

# Transition classification thresholds, based on the outgoing version's share
# of total family UNITS in the most recent year:
CLEAN_CUTOVER_MAX_SHARE = 0.15      # old version < 15% of family units -> clean cutover
HEAVY_OVERLAP_MIN_SHARE = 0.35      # old version > 35% of family units -> heavy overlap
                                     # in between -> "partial overlap"

def classify_transition(old_version_share_of_family_units: float) -> str:
    if old_version_share_of_family_units is None:
        return "n/a"
    if old_version_share_of_family_units < CLEAN_CUTOVER_MAX_SHARE:
        return "CLEAN CUTOVER"
    if old_version_share_of_family_units > HEAVY_OVERLAP_MIN_SHARE:
        return "HEAVY OVERLAP"
    return "PARTIAL OVERLAP"


# ---------------------------------------------------------------------------
# 4. Account context overrides — CONFIRMED WITH THE BUSINESS (2026-07)
# ---------------------------------------------------------------------------
# These three accounts are structurally different from a normal wholesale
# partner. Their raw margin/sales trend should NOT be read the same way as
# other accounts — see notes. Confirmed directly by the business; if this
# changes (e.g. XXL's exit completes, or China's terms change), update here
# and every downstream analysis picks it up automatically.
ACCOUNT_CONTEXT = {
    "XXL SE": "XXL is being exited from Key Accounts (planned wind-down, both markets). "
              "NOTE: this is layer-inconsistent in the data — XXL is growing in Mid Layer "
              "and Soft shell while shrinking in Shell/Insulation/Legwear/Accessories. "
              "Confirm whether this is a deliberate category-level rationalization.",
    "XXL NO": "See XXL SE note — same account family, same caveat.",
    "ZALANDO": "This is Zalando Lounge — an OUTLET/CLEARANCE channel, not a normal "
               "wholesale/retail partner. Sales-up-margin-down here is expected outlet "
               "behavior, not discount-driven demand destruction. Read Zalando's *volume* "
               "growth as an inventory/clearance signal, not its margin % as an erosion signal. "
               "NOTE (2026-07): the raw Customer Group value is 'ZALANDO' (all caps) — a prior "
               "version of this key was 'Zalando', which silently fails any case-sensitive match "
               "against the raw data. Confirmed against actuals: FY25 Sales 52.6M SEK, GM 30.1%.",
    "ZALANDO MARKETPLACE": "A separate, smaller Zalando entry (FY25: SEK 0.46M, GM 69.3%) — "
               "different commercial mechanism (marketplace/commission model, not wholesale-in), "
               "materially higher margin than 'ZALANDO' (Lounge/outlet). Keep these two separate "
               "in any account-level view; don't blend them.",
    "China": "Special cost-plus commercial arrangement — structurally different economics "
             "from other accounts. Its low blended margin (~11% company-wide FY25) is by "
             "design, not a quality-of-account issue, and isn't comparable like-for-like to "
             "XXL/Zalando/Intersport. Watch the volume trend (SEK 0.5M FY24 -> SEK 49.1M "
             "FY25 company-wide), not the margin %.",
    "STADIUM Outlet": "A second outlet-type channel identified in Legwear FY25 (new, "
                       "SEK ~2.0M). Not yet confirmed with the business the way Zalando/"
                       "China have been — treat provisionally the same way (volume signal, "
                       "not a margin-erosion signal) until confirmed.",
}

# ---------------------------------------------------------------------------
# 5. Ordertype x Channel structure
# ---------------------------------------------------------------------------
# Ordertype and Channel are NOT independent in this data. Confirmed pattern
# (holds across all layers checked so far):
#   - Retail channel  -> ~100% "Retail" ordertype
#   - E-com/Marketplace -> ~100% "Reorder Group" ordertype
#   - Wholesale channel -> the only channel with a real Preorder/Reorder/
#     Closeout mix
# This is WHY the Reorder-Group-split-by-channel view matters: blending
# channels within Reorder can mask a real Wholesale-specific problem behind
# E-com's stability. Always check the split, don't just read blended Reorder.
CHANNELS_WITH_REAL_ORDERTYPE_MIX = ["Wholesale"]
REORDER_SPLIT_CHANNELS = ["Wholesale", "E-com"]  # the two worth comparing directly

# ---------------------------------------------------------------------------
# 6. Cross-layer findings log (running, update as more layers are analyzed)
# ---------------------------------------------------------------------------
# This isn't code the pipeline reads — it's a plain-English running tally so
# whoever picks this up next (human or Claude) knows what's been confirmed
# vs. still provisional. Update after each layer.
CROSS_LAYER_FINDINGS_LOG = """
1. Wholesale Reorder cost inflation (COGS/unit rising 12-44% while E-com
   Reorder stays flat/improving) -- confirmed in 7/8 layers checked
   (Shell, Insulation, Mid layer, Legwear, Soft shell, Tops, Accessories).
   Daypacks (Hardware) is the one exception found so far (+5.3% only) --
   working theory: this is an APPAREL-specific problem (small cut-and-sew
   reorder batches, rush freight) rather than truly company-wide. Needs
   testing against the remaining hardware/footwear layers to confirm.

2. Top articles in every layer checked are running meaningfully BELOW
   their FW27 target margin (per the Assortment Attribution Review),
   typically by 10-30 points, regardless of whether the layer's overall
   margin trend is up or down. This looks systematic, not article-specific.

3. Recurring assortment-file anomaly: several FW27 "II"/successor styles
   are flagged Active = FALSE despite the base style still selling real
   volume (Bield Down II Hood - Insulation, Zircon Slim II Pant - Legwear,
   Gabbro II Pant - Soft shell). Worth a direct question to Product/the
   planning tool owner about what "Active" is supposed to capture.

4. Version-transition lifecycle: FY24->FY25 successions split cleanly into
   CLEAN CUTOVER / PARTIAL OVERLAP / HEAVY OVERLAP / (new) GAP YEAR (old
   style already wound down, successor not yet selling -- e.g. Gabbro Pant
   in Soft shell). Across nearly all HEAVY OVERLAP cases, the outgoing
   version's ASP barely moves or even rises -- little evidence of active
   markdown/clearance pricing being used to manage the handover.

5. XXL account behavior is genuinely split by layer: exiting in Shell,
   Insulation, Legwear, Accessories; growing in Mid layer and Soft shell.
   Confirm with Sales whether this is intentional category consolidation.
"""
