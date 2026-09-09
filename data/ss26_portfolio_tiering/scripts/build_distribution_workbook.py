"""
Builds a dated DISTRIBUTION cut of the master tiering workbook: Global
Full Portfolio + Global Generation Detail + the two regional tiering
sheets (Nordic, Non-Nordic) -- drops the Browse view, Stock by Tier, and
Sales by Country sheets, which stay internal-working-file-only.

This does NOT modify the master workbook (Project_Bob_Portfolio_Tiering_
08092026.xlsx) -- it reads from it and writes a brand-new file named with
today's date, following the existing DDMMYYYY convention.

Sheet 1 ("Intro & Definitions") is rebuilt, not copied verbatim: all of
its existing content is kept (none of it was specific to the two dropped
sheets), plus three new sections -- a sheet index labeling each sheet
Global vs Regional, the exact Global tier rule (REG-024, previously
undocumented anywhere in this workbook), and the regional-tiering
methodology (REG-023's region-scaled SEK floors).

Usage:
    python3 data/ss26_portfolio_tiering/scripts/build_distribution_workbook.py
"""
from __future__ import annotations

import datetime
import math

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

import ss26_lib as lib
import update_regional_tiering as rt

TODAY = datetime.date.today().strftime("%d%m%Y")
OUT_PATH = lib.TIERING_DIR / f"Project_Bob_Portfolio_Tiering_{TODAY}.xlsx"

SHEETS_TO_DROP = ["2b. Full Portfolio (Browse)", "4. Stock by Tier", "5. Sales by Country"]
RENAMES = {
    "2. Full Portfolio (1860)": "2. Full Portfolio (Global)",
    "3. Generation Detail": "3. Generation Details (Global)",
    "6. Full Portfolio (Nordic)": "4. Full Portfolio (Nordic)",
    "7. Full Portfolio (Non-Nordic)": "5. Full Portfolio (Non-Nordic)",
}

TITLE_COLOR = "FF1F3864"
GRAY = "FF555555"
HEADER_FILL = lib.styles()["header_fill"]
HEADER_FONT = lib.styles()["header_font"]

COL_WIDTHS = {"A": 26.0, "B": 12.0, "C": 17.0, "D": 10.0, "E": 18.0, "F": 12.0, "G": 9.0}
CHARS_PER_LINE = 180  # calibrated against the existing sheet's own row heights


def para_height(text: str) -> float:
    lines = math.ceil(len(text) / CHARS_PER_LINE)
    return max(15.75, lines * 14)


class SheetWriter:
    """Small helper so the intro-sheet content below reads as a list of
    blocks (title / section / subsection / paragraph / table) instead of
    raw cell coordinates."""

    def __init__(self, ws):
        self.ws = ws
        self.r = 1
        for col, width in COL_WIDTHS.items():
            ws.column_dimensions[col].width = width

    def _merge_row(self):
        self.ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=7)

    def blank(self, n=1):
        self.r += n

    def title(self, text):
        cell = self.ws.cell(row=self.r, column=1, value=text)
        cell.font = Font(name="Arial", bold=True, size=16, color=TITLE_COLOR)
        self.r += 1

    def section(self, text):
        cell = self.ws.cell(row=self.r, column=1, value=text)
        cell.font = Font(name="Arial", bold=True, color=TITLE_COLOR)
        self.r += 1

    def subsection(self, text):
        self.section(text)  # same style, used at a nested level

    def para(self, text, gray=False):
        self._merge_row()
        cell = self.ws.cell(row=self.r, column=1, value=text)
        cell.font = Font(name="Arial", color=GRAY if gray else None)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        self.ws.row_dimensions[self.r].height = para_height(text)
        self.r += 1

    def table_header(self, headers):
        for c, h in enumerate(headers, start=1):
            cell = self.ws.cell(row=self.r, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        self.r += 1

    def table_row(self, values):
        for c, v in enumerate(values, start=1):
            self.ws.cell(row=self.r, column=c, value=v)
        self.r += 1


def build_intro_sheet(ws):
    w = SheetWriter(ws)
    w.title("Project Bob — Portfolio Tiering (Global + Regional)")
    w.blank()
    w.para(
        "Every continuing Core franchise (Base + Gender) is assigned to a tier based on its FY2024/FY2025 sales "
        "size, growth, and margin. This distribution cut, generated "
        f"{datetime.date.today().strftime('%d %b %Y')}, contains the Global base tiering, the Global Generation "
        "Detail sheet, and the two regional (Nordic / Non-Nordic) re-applications of the same tier rule. It drops "
        "the Browse-order view, Available Stock by Tier, and Sales by Country sheets from the full internal "
        "working file — those live in the master workbook and the Assumptions Register, not here."
    )
    w.blank()

    w.section("What's in this file")
    w.para(
        "Each sheet below is labeled Global (worldwide, Core accounts) or Regional (Nordic- or Non-Nordic-only "
        "sales). The two regional sheets use the SAME tier rule as the Global sheet, region-scoped — see "
        "\"Regional Tiering Methodology\" further down for exactly how.",
        gray=True,
    )
    w.table_header(["Sheet", "Scope", "What it shows"])
    w.table_row(["2. Full Portfolio (Global)", "Global (worldwide)",
                 "The audited 1,860-franchise base tiering — Sales_2025/YTD2026, GM%, Growth%, Pace%, plus "
                 "assortment/channel columns. Source of truth; the two regional sheets are a different, "
                 "region-scoped cut of the same rule."])
    w.table_row(["3. Generation Details (Global)", "Global",
                 "Article-naming version successions (Original vs II/III/IV) for the 109 franchises with a "
                 "genuine generation change — margin/sales comparison per generation, worldwide."])
    w.table_row(["4. Full Portfolio (Nordic)", "Regional — Nordic (Sweden, Norway, Denmark, Finland)",
                 "The same tier rule as Sheet 2, reapplied using ONLY Nordic sales, with Nordic-scaled SEK floors."])
    w.table_row(["5. Full Portfolio (Non-Nordic)", "Regional — Non-Nordic (everywhere else)",
                 "The same tier rule as Sheet 2, reapplied using ONLY non-Nordic sales, with Non-Nordic-scaled "
                 "SEK floors."])
    w.blank()

    w.section("Summary by Tier")
    w.para(
        "Sales_2025 = full FY2025 (Core accounts). Sales_YTD2026 = Jan-Aug 2026 (Core accounts), same franchise "
        "match methodology (Base+Gender). Pace % = Sales_YTD2026 / Sales_2025 (pre-read convention, not "
        "annualized). \"Not Yet Tiered\" (5.2M SEK YTD2026 sales) = brand-new SS26 launches with no FY24/25 "
        "history, excluded from this tier table by definition. Clearance — Ex China/Zalando (REG-010) = "
        "franchises fully moved out of their original tier because every known SKU they have is on the Aug-2026 "
        "China JV / Zalando close-out lists. Hero+Near-Hero and Workhorse+Harvest (REG-013) are consolidated "
        "tiers — see Sheet 2 column P for each franchise's tier before either change. TOTAL row is unaffected by "
        "both changes — these are re-labelings of existing rows, not new/removed sales.",
        gray=True,
    )
    w.table_header(["Tier", "# Franchises", "Sales_2025 (SEK)", "GM%_2025", "Sales_YTD2026 (SEK)", "GM%_YTD2026"])
    for row in [
        ["Hero + Near-Hero", 34, 97931546, 53.2, 50733948, 54.3],
        ["Workhorse + Harvest", 431, 316679716, 48.7, 139975555, 47.7],
        ["Problem Child", 150, 94513733, 33.5, 32955867, 32.9],
        ["New / Test", 189, 122822803, 46.5, 92223739, 51.6],
        ["Thin / Immaterial", 837, 6822109, 28.5, 58964006, 47.6],
        ["Exited", 171, 7179342, 40.4, 4975003, 36.4],
        ["Clearance — Ex China/Zalando", 48, 3810858, 37.8, 2668099, 36.1],
        ["TOTAL", 1860, 649760108, 46.3, 382496216, 48],
    ]:
        w.table_row(row)
    w.blank()

    w.section("Global Tier Definitions — The Exact Rule (added 09-Sep-2026, REG-024)")
    w.para(
        "Previously undocumented anywhere in this workbook — confirmed directly by the business on 9-Sep-2026 "
        "and validated against this sheet's own published tiers at 99.8% accuracy (1,808 of 1,812 non-Clearance "
        "franchises match exactly; the 4 mismatches are article-name parsing edge cases in the validation "
        "reconstruction, not rule failures — see scripts/verify_base_tiering.py). This is the actual rule "
        "behind every Tier value on Sheet 2."
    )
    w.table_header(["Order", "Tier", "Rule"])
    for row in [
        [1, "New / Test", "No real trading in 2024, real trading in 2025 (i.e., a genuine launch)"],
        [2, "Exited", "Real trading in 2024, not in 2025"],
        [3, "Thin / Immaterial", "Below the 100-unit real-trading threshold in one or both years"],
        [4, "Hero", "2025 sales ≥2M SEK, YoY growth >0%, sells in both Wholesale and DTC, GM% ≥46.6%"],
        [5, "Near-Hero (Rising Star)", "2025 sales ≥1M SEK, YoY growth >15%, cross-channel, GM% ≥46.6%"],
        [6, "Workhorse", "Not Hero/Near-Hero, GM% ≥40%, YoY growth > −10%"],
        [7, "Harvest (Cash Cow)", "Not Hero/Near-Hero, GM% ≥40%, YoY growth ≤ −10%"],
        [8, "Problem Child", "Not Hero/Near-Hero, GM% <40% (regardless of growth)"],
    ]:
        w.table_row(row)
    w.para(
        "Notes on the terms — \"Real trading\" = ≥100 units sold in that year (MIN_UNITS_FOR_REAL_TRADING_YEAR "
        "in config.py) — the same threshold used throughout the Key-Article Margin Analysis workstream. 46.6% "
        "is the core-company average GM% — the bar for \"good,\" not an arbitrary round number. 40% is a "
        "separate, lower floor — Workhorse/Harvest/Problem Child are split by whether they clear it, not by the "
        "46.6% Hero bar. −10% YoY is the growth line that separates Workhorse (holding up) from Harvest "
        "(declining but still margin-healthy) — both tiers require GM% ≥40%; only the growth direction differs. "
        "Cross-channel for Hero/Near-Hero specifically means Wholesale >100,000 SEK AND DTC >50,000 SEK (this "
        "exact threshold) — not just \"any presence in both.\"",
        gray=True,
    )
    w.blank()

    w.section("Regional Tiering Methodology — Nordic & Non-Nordic (added 09-Sep-2026, REG-023)")

    fy25 = rt.load_franchise_region_totals(2025)
    nordic_share = rt.region_revenue_share(fy25, "Nordic")
    non_nordic_share = rt.region_revenue_share(fy25, "Non-Nordic")

    w.para(
        "Sheets 4 and 5 apply the EXACT SAME rule above, but the rule's SEK floors (2M Hero, 1M Near-Hero, "
        "100K/50K cross-channel) are global, absolute numbers — reusing them unscaled for each region separately "
        "would just be \"same absolute thresholds as Global,\" which was considered and declined, since a "
        "smaller region's pool would rarely clear a bar calibrated to global-scale revenue. Instead, each "
        "region's SEK floors are scaled by that region's own share of total FY25 Core sales (recomputed fresh "
        "each time this file is built, not hardcoded — currently as shown below). GM%/growth cuts and the "
        "100-unit real-trading threshold are NOT scaled — GM%/growth are already relative (%), and 100 units is "
        "a physical reality-check, not a revenue-scale artifact. Clearance is carried over unchanged from Sheet "
        "2 (China JV/Zalando SKU-list membership isn't region-specific)."
    )
    w.table_header(["Region", "Share of FY25 Core Sales", "Hero Sales Floor (SEK)", "Near-Hero Sales Floor (SEK)",
                     "Wholesale Floor (SEK)", "DTC Floor (SEK)"])
    for label, share in [("Nordic", nordic_share), ("Non-Nordic", non_nordic_share)]:
        w.table_row([
            label, round(share * 100, 1),
            round(rt.HERO_SALES_FLOOR_GLOBAL * share), round(rt.NEAR_HERO_SALES_FLOOR_GLOBAL * share),
            round(rt.WHOLESALE_FLOOR_GLOBAL * share), round(rt.DTC_FLOOR_GLOBAL * share),
        ])
    w.para(
        "Concrete illustration of why this matters: \"Astral GTX Jacket\" Men stays Workhorse in BOTH Nordic and "
        "Non-Nordic (GM% 42.9%/44.5%, both below the 46.6% Hero bar) — consistent with its Global tier, for the "
        "actual reason (margin, not size). Result: 32 Nordic Hero+Near-Hero franchises vs. 19 Non-Nordic — this "
        "reflects rank within a smaller home (Nordic) population at a proportionally-scaled bar, not superior "
        "absolute performance. Figures on Sheets 4/5 are recomputed from bob_salesdata_2024/2025.xlsx (Core-"
        "scope via is_core_customer_group()) — full detail, caveats, and the completeness fix (every franchise "
        "now gets a row even with zero regional sales) are in REG-023/REG-024 in the Assumptions Register.",
        gray=True,
    )
    w.blank()

    w.section("Tier Definitions")
    for name, text in [
        ("Hero + Near-Hero",
         "Continuing franchises with the best margin in the portfolio (53.2% GM%, blending the former Hero and "
         "Near-Hero tiers). The proof that ~53%+ margin is achievable, at both large scale (former Hero) and "
         "smaller/fast-growing scale (former Near-Hero). Consolidated 08-Sep-2026 (REG-013): YTD2026 growth and "
         "margin trends between the two were converging, making the original size-based cut line less "
         "meaningful — see Sheet 2 column P to split back into the original two if needed."),
        ("Workhorse + Harvest",
         "The largest tier by both sales (316.7M SEK FY25) and count (431 franchises) — continuing franchises "
         "with mid-range-to-solid margin (48.7%), blending the former Workhorse (mid-margin, modest growth) and "
         "Harvest/Cash Cow (solid margin, flat-to-declining growth, median -18% YoY combined) tiers. "
         "Consolidated 08-Sep-2026 (REG-013) because YTD2026 pace (44.2%) and margin trends no longer separated "
         "the two tiers consistently. The volume backbone of the portfolio."),
        ("Problem Child",
         "Continuing franchises with real revenue but structurally low margin (33.5%, well below the ~53% "
         "target and the lowest of any continuing tier). Needs a pricing/cost fix, not necessarily an exit — "
         "see the Wholesale-specific margin gap flagged elsewhere in the pre-read. 58 of these 150 franchises "
         "also carry a Partial Clearance flag (an old-generation SKU exiting via China JV/Zalando) — read the "
         "Clearance Detail column before prioritizing a fix; the margin issue may already be shrinking on its "
         "own."),
        ("New / Test",
         "Franchises launched too recently to have a full FY2024 history, so they aren't yet tiered on the "
         "same growth/margin basis as the continuing tiers. Being watched to see which graduate into Hero, "
         "Workhorse, or Problem Child. Running well ahead of pace YTD2026 (75.1%, GM% up to 51.6%)."),
        ("Thin / Immaterial",
         "Below the materiality threshold on FY24/25 sales — individually too small to manage at the franchise "
         "level. By far the largest tier by count (837, 45% of all franchises) but only 1.0% of FY25 Core "
         "sales. Its YTD2026 Pace% (864.3%) is a base-effect artifact of a tiny FY25 denominator, not real "
         "growth — read the GM% (28.5%→47.6%) instead."),
        ("Exited",
         "Franchises the business has discontinued — residual/wind-down sales only, expected to trend to "
         "zero."),
        ("Clearance — Ex China/Zalando",
         "Franchises whose entire known SKU range (every Model/Color combination ever sold) appears on the "
         "Aug-2026 China JV close-out delivery or Zalando close-out order lists — i.e. nothing is left behind "
         "to carry the franchise forward. Moved here from their original tier on 08-Sep-2026 so the other "
         "tiers only show franchises with a live go-forward assortment. Sales_2025/GM%_2025 are the "
         "franchise's actual FY25 history, unchanged — this is a re-label, not a deletion. 48 franchises, "
         "mostly ex-Thin/Immaterial (27) and ex-New/Test (8); only 1 came out of Problem Child and none out "
         "of Hero or Near-Hero, because no Hero/Near-Hero franchise had its entire SKU range on either "
         "close-out list — the current-generation article always survived. See REG-010 for the matching "
         "methodology and caveats."),
    ]:
        w.subsection(name)
        w.para(text)
    w.blank()

    w.section("Clearance Methodology Note (added 08-Sep-2026, REG-010)")
    w.para(
        "Source: Close_out_delivery_JV_August_2026.xlsx (China JV distributor, 321 SKUs) + "
        "Zalando_close_out_order_August_2026.xlsx (127 SKUs) — 421 distinct SKUs (SCO = Model+Color ID) "
        "combined, provided by the business as the Aug-2026 stock clean-up lists."
    )
    w.para(
        "Matching: each SKU's SCO was matched to the raw SS26 exports to find its Article name, then rolled up "
        "to franchise (Base+Gender) using the same gender/version-token stripping already established for this "
        "business (config.GENDER_TOKENS / VERSION_TOKENS from the Key-Article Margin Analysis workstream), plus "
        "two fallbacks for glued-text export artifacts (no space before \"Men\"/\"Women\", or before \"II\"). "
        "418 of 421 SKUs matched a raw sales history; all 418 matched an existing published franchise. 3 SKUs "
        "(6045574DL, 6056693JR, 6056694T9) never appear in any of the three raw exports — no impact on any "
        "franchise's figures, noted for completeness."
    )
    w.para(
        "Treatment rule: for each of the 260 franchises touched, compared the full set of SKUs the franchise "
        "has ever had against the SKUs on the close-out lists. \"Fully\" (48 franchises) = every known SKU is "
        "on the list — moved wholesale into the new Clearance tier, tier totals adjusted accordingly. "
        "\"Partial\" (212 franchises) = a current-generation SKU survives — franchise stays in its original "
        "tier untouched (Tier/Sales_2025/GM% unchanged), only flagged via the Clearance Flag/Detail columns on "
        "Sheet 2 so a reader knows part of its FY25 history was old-generation stock now exiting. No dollar "
        "amount was split out of any Partial franchise's totals — an earlier attempt to reconstruct a SEK "
        "carve-out from the raw exports did not reconcile against the published Sales_2025 figures (traced to "
        "how the raw data's \"3-Close out order\" ordertype, ~13.5% of all FY25 sales, is treated in the "
        "published Core-scope definition — not yet confirmed, see REG-012) and was abandoned in favor of this "
        "simpler, dollar-free rule."
    )
    w.para(
        "Known data-quality item (not corrected here, flagged as REG-011): the raw Article text contains at "
        "least one case-sensitive duplicate franchise pair — e.g. \"Roc Sight SoftshellJacket\" and \"ROC Sight "
        "Softshell Jacket\" exist as two separate published Base rows for what looks like one product line. "
        "Worth a check with the data owner; not assumed away."
    )
    w.blank()

    w.section("Tier Consolidation Note (added 08-Sep-2026, REG-013)")
    w.para(
        "Hero and Near-Hero (Rising Star) were merged into \"Hero + Near-Hero\"; Workhorse and Harvest (Cash "
        "Cow) were merged into \"Workhorse + Harvest\" — both on your direction, because SS26 YTD2026 sales "
        "growth and margin trends were moving in ways that made the original size-based cut lines between each "
        "pair less consistent as a decision boundary. This is a re-labeling on top of the already-current "
        "(post-Clearance) row-level data: each combined tier's Sales_2025/GM%_2025/Sales_YTD2026/GM%_YTD2026/"
        "Pace% is the sales-weighted roll-up of its two source tiers, computed from row-level figures (not from "
        "the rounded summary-table inputs), so grand TOTAL is unaffected."
    )
    w.para(
        "Column P on Sheet 2 (\"Original Tier (pre-consolidation)\") keeps each franchise's tier immediately "
        "before this consolidation — i.e. the 9-tier scheme including the Clearance move from REG-010 — so the "
        "original Hero/Near-Hero/Workhorse/Harvest split can always be reconstructed by filtering on that "
        "column."
    )
    w.blank()

    w.section("Core Assortment FW27 Check (added 08-Sep-2026, updated 08-Sep-2026 with the authoritative Excel list, REG-015)")
    w.para(
        "New column on Sheet 2, \"Core Assortment FW27\": Yes/No, matched against the authoritative "
        "FW27_CORE_Assortment_Styles.xlsx (57 carry-over styles, each with an exact Style Number/Model code) "
        "plus the 31-style \"News package\" from the earlier send-out deck (Model codes from that deck's own "
        "table) — 88 items combined. This supersedes the first version of this check, which was transcribed "
        "from a PDF slide deck's image captions (no Model codes for the carry-over section) and covered a "
        "different, less complete 64-style carry-over set — the deck itself was labeled \"version 1.\" Matched "
        "by parsing each item's own style/article name (same gender/version-token method as REG-010/014), NOT "
        "by looking up whether that exact Model code has ever appeared in the SS26 raw exports — several "
        "2.0/II successor styles (e.g. \"Mimic Alert 2.0 Hood\", \"Rosson Mid II Jacket Women\") haven't "
        "shipped yet and have zero sales history, so a sales-history lookup would have wrongly marked their "
        "whole franchise absent from Core Assortment; matching on the name itself avoids that."
    )
    w.para(
        "Match quality: 57 of 88 items matched an existing published franchise; 31 unmatched are genuinely new "
        "(the News package minus \"Mimic Alert\", which matched; plus \"Korp Softshell Hood/Pant\", "
        "independently confirmed to have no FY24/25 sales history either way)."
    )
    w.subsection("Hero + Near-Hero franchises NOT in the Core Assortment FW27 list — for merch review")
    w.table_header(["Base", "Gender", "Sales_2025 (SEK)", "GM%_2025"])
    for row in [
        ["Korp Proof Jacket", "Men", 3886774, 55.2], ["Astral GTX Pant", "Women", 3536140, 49],
        ["Salix Proof Mimic Parka", "Men", 3359528, 53], ["Rosson Down Hood", "Women", 2920336, 52.4],
        ["Chaos GTX Jacket", "Women", 2602402, 46.7], ["Korp Proof Jacket", "Women", 2319994, 57],
        ["Morän Softshell Standard Pant", "Men", 2118928, 47.1], ["Vassi GTX Pro Jacket", "Women", 2015935, 46.8],
        ["Jarve Multi 28", "Unisex", 1977999, 49.2], ["L.I.M Fuse Pant", "Men", 1704736, 47],
        ["L.I.M Mid Multi Hood", "Men", 1557975, 50.9], ["L.I.M Fuse Shorts", "Women", 1451642, 50],
        ["Fjatla 60", "Unisex", 1410361, 51.5], ["Lava 30", "Unisex", 1359786, 49.5],
        ["Mossa Pile Jacket", "Women", 1356082, 58.6], ["Tarius -5", "Unisex", 1348641, 48.1],
        ["Rugged Slim Pant", "Women", 1201422, 50], ["L.I.M Mid Multi Hood", "Women", 1179954, 52.8],
        ["Korp Proof Pant", "Men", 1104990, 53.3], ["L.I.M Proof Pant", "Women", 1014937, 54.6],
    ]:
        w.table_row(row)
    w.blank()
    w.para(
        "20 of 34 Hero + Near-Hero franchises (59%) are not in the Core Assortment FW27 list — down from 21 "
        "under the earlier PDF-based version; \"Long Down Parka\" (Women, 6.3M SEK) dropped off this list "
        "because the authoritative file confirms its successor, \"Long Down II Parka Women,\" is in Core "
        "Assortment. Whether the remaining 20 are intentional cuts or oversights is still a merch-team call, "
        "not a data question."
    )
    w.blank()

    w.section("FW27 Collection Check (added 08-Sep-2026, REG-016)")
    w.para(
        "New column on Sheet 2, \"FW27 Collection\": Active / Not active / Not in review — sourced from the "
        "\"Active\" field in data/Assortment Attribution Review(F27).xlsx, matched by franchise (Base+Gender), "
        "same text parsing as REG-010/014/015. \"Active\" = at least one matching style is Active=True. \"Not "
        "active\" = every matching style is Active=False. \"Not in review\" = the franchise doesn't appear in "
        "this file at all (1,568 of 1,860) — this file covers 292 of the 1,860 franchises, so absence is not "
        "itself a negative signal, just no data."
    )
    w.para(
        "Known caveat (per this repo's README, Key-Article Margin Analysis workstream): this \"Active\" field "
        "has a documented false-negative problem — successor styles that are clearly still trading have shown "
        "Active=False before. Confirmed still current: \"Zircon Slim II Pant\" is Active=False in the source "
        "file, but the matching franchise \"Zircon Slim Pant\" (Men) has 546,751 SEK in YTD2026 sales. Per your "
        "direction, the source file's value is kept as-is (not overridden) — the table below is the exception "
        "list instead: every \"Not active\" franchise with material (≥50,000 SEK) YTD2026 sales, for review "
        "against whether the underlying attribution data needs a refresh."
    )
    w.subsection("\"Not active\" franchises with material YTD2026 sales — for review")
    w.table_header(["Base", "Gender", "Tier", "Sales_YTD2026 (SEK)", "Sales_2025 (SEK)"])
    for row in [
        ["Velum Jacket", "Men", "New / Test", 1514939, 220891], ["Velum Jacket", "Women", "New / Test", 921906, 200652],
        ["Zodiac Jacket", "Men", "Thin / Immaterial", 835543, 4440], ["Bield Down Hood", "Men", "Exited", 778509, 74586],
        ["Zodiac Jacket", "Women", "Thin / Immaterial", 592478, 10766], ["Jarve Single 20", "Unisex", "Workhorse + Harvest", 569800, 306169],
        ["Zircon Slim Pant", "Men", "Thin / Immaterial", 546751, 32238], ["Kaja Proof Jacket", "Men", "Thin / Immaterial", 535334, 10243],
        ["Steep Proof 3L Jacket", "Men", "Thin / Immaterial", 367272, 0], ["Spacelite -1", "Unisex", "Thin / Immaterial", 284609, 0],
        ["Spacelite +7", "Unisex", "Thin / Immaterial", 254514, 0], ["Steep Proof 3L Jacket", "Women", "Thin / Immaterial", 251926, 0],
        ["Apex Hood", "Men", "Thin / Immaterial", 232640, 0], ["Zed Jacket", "Men", "Thin / Immaterial", 180044, 3680],
        ["Scree Pant", "Men", "Thin / Immaterial", 177795, 53906], ["Scree Pant", "Women", "Thin / Immaterial", 139775, 41554],
        ["Zed Jacket", "Women", "Thin / Immaterial", 120133, 12682], ["Zircon Slim Pant", "Women", "New / Test", 119433, 108591],
        ["Apex Hood", "Women", "Thin / Immaterial", 85493, 5211], ["Gabbro Pant", "Men", "Thin / Immaterial", 82445, 0],
        ["Gabbro Pant", "Women", "Thin / Immaterial", 79991, 4170], ["Nordic Expedition Down -12", "Unisex", "Thin / Immaterial", 75131, 92688],
        ["Nordic Expedition Down -2", "Unisex", "Exited", 53857, 107477], ["Scand Down +5", "Unisex", "Thin / Immaterial", 51503, 83251],
    ]:
        w.table_row(row)
    w.blank()
    w.para(
        "24 franchises meet this bar. Several (e.g. Velum Jacket, Zodiac Jacket, Spacelite, Steep Proof 3L "
        "Jacket) show FY25 sales near zero with substantial YTD2026 sales — likely brand-new SS26 launches "
        "that postdate this attribution snapshot, not the same successor-naming anomaly as Zircon Slim II. "
        "Distinguishing \"snapshot is stale\" from \"genuine naming anomaly\" per line needs the attribution "
        "file's own refresh date, which isn't available here — flagging both patterns together as one review "
        "list rather than guessing which is which."
    )
    w.blank()

    w.section("Wholesale Share % Column (added 08-Sep-2026, REG-017)")
    w.para(
        "New column on Sheet 2, \"Wholesale Share % (FY25, raw)\": Wholesale channel's share of the "
        "franchise's FY25 Core-scope sales (Wholesale vs. Retail + E-com / \"DTC\" combined), recomputed from "
        "the raw SS26 exports at franchise grain — same method and same REG-012 caveat as the Generation "
        "Detail sheet. Company-wide FY25 Core baseline for reference: Wholesale 40.7% GM%, Retail 48.9%, "
        "E-com 61.1% (Wholesale is the lowest-margin channel) — this column lets a reader connect a "
        "franchise's margin to its channel mix without a separate pivot."
    )
    w.para(
        "Coverage and reliability: 866 of 1,860 franchises have a value; 664 are blanked (total reconstructed "
        "sales below the 50,000 SEK/year threshold, REG-004's precedent); 330 have no matching raw Core-scope "
        "sales at all (mostly Thin/Immaterial or Exited). Of the 1,530 franchises with any reconstructed "
        "sales, 366 (24%) differ from the published Sales_2025 by more than 15% — the same unresolved "
        "\"3-Close out order\" ordertype gap as REG-012, which skews toward overstating Wholesale share "
        "(close-out activity is wholesale-channel) wherever it's present. Read this column as directional, "
        "not audited — useful for spotting a wholesale-heavy vs. DTC-heavy franchise, not as a precise "
        "percentage. One residual/wind-down franchise (\"Asp 3-in-1 GTX Parka\" Women, Exited, 203K SEK FY25) "
        "shows a small negative value from returns exceeding gross wholesale sales on a low base — left as "
        "computed rather than suppressed."
    )
    w.blank()

    w.section("Channel Pattern 2025 (added 08-Sep-2026, revised same day, REG-018)")
    w.para(
        "New column on Sheet 2, \"Channel Pattern 2025\": classifies each franchise's 2024->2025 Wholesale vs. "
        "DTC (Retail + E-com) pattern. Source: data/bob_salesdata_2024.xlsx + bob_salesdata_2025.xlsx "
        "(Key-Article Margin Analysis workstream's own two-year annual pull — this is what \"that channel "
        "analysis\" referred to, analysis.py's channel_view() function), filtered to the same Core-account "
        "exclusions as the rest of Sheet 2 (REG-008: all XXL market variants, Zalando, Zalando Marketplace, "
        "China, Stadium Outlet — cross-checked: excluded China total ties to config.py's own documented 49.6M "
        "SEK FY24+FY25 figure). \"Marketplace\" Sales Channel (Sport-Scheck, About You, Zalando Marketplace, "
        "etc.) is grouped into Wholesale, not DTC — an assumption, not independently confirmed with the "
        "business."
    )
    w.para(
        "Revised same day: the first version applied the 50,000 SEK/year reliability threshold per channel "
        "per year and lumped every franchise that failed it into one \"Insufficient data\" bucket — but that "
        "conflated two different things: a franchise that's genuinely tiny overall, and a franchise with "
        "real, sometimes substantial sales that's just structurally concentrated in one channel (so the "
        "OTHER channel never has enough volume to compute a meaningful growth rate, in either year). Median "
        "sales for the original 1,555-franchise bucket was 17,360 SEK (34 units) — mostly genuinely small — "
        "but 564 of them (36%) had Sales_2025 above the threshold, revealing they were single-channel, not "
        "small. Split out below."
    )
    w.subsection("Categories")
    for name, text in [
        ("Substitution (DTC ↑ / Wholesale ↓)", "135 franchises — DTC grew, Wholesale declined, 2024->2025. Both channels clear the reliability bar in both years. The single most common classified pattern."),
        ("Substitution (Wholesale ↑ / DTC ↓)", "26 franchises — the reverse: Wholesale grew, DTC declined."),
        ("Co-growth (both ↑)", "66 franchises — both channels grew."),
        ("Co-decline (both ↓)", "70 franchises — both channels declined (or flat/zero, grouped with decline)."),
        ("Wholesale-only / DTC negligible", "59 franchises — Wholesale clears the reliability bar in both years; DTC is below 50,000 SEK in BOTH years (a stable structural fact, not a data gap) — no meaningful DTC growth rate to report. 37 of these are Workhorse + Harvest, 14 Problem Child."),
        ("DTC-only / Wholesale negligible", "83 franchises — the reverse: DTC reliable both years, Wholesale below 50,000 SEK in both years. 29 Workhorse + Harvest, 21 Exited, 15 Problem Child."),
        ("Insufficient data", "1,413 franchises — doesn't fit any of the above: either both channels are small in both years (the genuinely tiny/immaterial majority — 823 of these are Thin/Immaterial), or one channel's reliability is mixed across years (clears the bar in one year but not the other), which is a real data-reliability gap, not a stable pattern. Includes 4 Hero + Near-Hero franchises."),
        ("Not in FY24/25 dataset", "8 franchises — no matching rows found in bob_salesdata_2024/2025.xlsx at all."),
    ]:
        w.subsection(name)
        w.para(text)
    w.para(
        "Classification is sign-based (any positive vs. any negative growth) for the four growth categories, "
        "no materiality band beyond the 50,000 SEK reliability gate — a franchise moving from +1% to -1% would "
        "flip categories. Cross-workstream caveat (REG-009): this dataset's totals have a known ~2.3% gap vs. "
        "the SS26 w.34 exports for the same period; franchise identity uses the same Base+Gender text matching "
        "as REG-010/014/015/017, not a guaranteed reconciliation to this sheet's own Sales_2025 column."
    )
    w.blank()

    w.section("Business Term Mapping: \"DTC Expansion Opportunity\" (added 08-Sep-2026)")
    w.para(
        "\"DTC expansion opportunity\" = Channel Pattern 2025 = \"Wholesale-only / DTC negligible\" (REG-018): "
        "a franchise with proven, reliable Wholesale demand but DTC (Retail+E-com) sales below 50,000 SEK in "
        "BOTH 2024 and 2025 — i.e. established market pull with essentially no direct-to-consumer presence to "
        "capture it. Logged here as the standing definition so the term reads the same way whether someone is "
        "scanning this sheet, asking Tier Bench, or reviewing this register."
    )
    w.table_header(["Tier", "# Franchises", "Wholesale-side Sales_2025 (SEK)"])
    for row in [
        ["Workhorse + Harvest", 37, 17541625], ["Problem Child", 14, 10276743], ["Exited", 4, 484514],
        ["Thin / Immaterial", 3, 360470], ["Clearance — Ex China/Zalando", 1, 254398],
    ]:
        w.table_row(row)
    w.blank()
    w.para(
        "37 of 59 are Workhorse + Harvest — proven, currently-performing franchises with no DTC foothold at "
        "all, the clearest whitespace. 14 are Problem Child, where opening a DTC channel could also help the "
        "margin story (DTC typically runs higher-margin than Wholesale — see REG-017's company baseline: "
        "Wholesale 40.7% GM% vs. Retail 48.9% vs. E-com 61.1%). Top individual opportunities by Wholesale-side "
        "Sales_2025: Haglöfs Krusa GTX Low (Men, Problem Child, 3.40M SEK), Husk Jacket (Men, Workhorse + "
        "Harvest, 2.75M SEK), Haglöfs Path GTX Low (Men, Problem Child, 1.59M SEK). Full list: filter Sheet 2, "
        "Channel Pattern 2025 = \"Wholesale-only / DTC negligible\"."
    )


def main():
    src = openpyxl.load_workbook(lib.WORKBOOK_PATH)

    for name in SHEETS_TO_DROP:
        if name in src.sheetnames:
            del src[name]
            print(f"dropped: {name}")

    intro = src["1. Intro & Definitions"]
    for merged in list(intro.merged_cells.ranges):
        intro.unmerge_cells(str(merged))
    intro.delete_rows(1, intro.max_row)
    build_intro_sheet(intro)
    print("rebuilt: 1. Intro & Definitions")

    for old_name, new_name in RENAMES.items():
        if old_name in src.sheetnames:
            src[old_name].title = new_name
            print(f"renamed: {old_name!r} -> {new_name!r}")

    src.save(OUT_PATH)
    print(f"saved: {OUT_PATH}")
    print("final sheet order:", src.sheetnames)


if __name__ == "__main__":
    main()
