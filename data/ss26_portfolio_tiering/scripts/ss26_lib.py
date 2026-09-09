"""
Shared parsing/matching helpers for the SS26 Portfolio Tiering workstream.

Every `update_*.py` script in this folder imports from here rather than
redefining its own copy. This is the consolidated, final version of logic
that was originally worked out iteratively (and inconsistently -- earlier
scratch attempts diverged) in one-off scripts during the Sep 2026 build;
see ASSUMPTIONS_REGISTER.md for the reasoning behind each rule.

ASSUMPTIONS (see ASSUMPTIONS_REGISTER.md for the full reasoning):
  - Franchise identity is Base+Gender: derived from the raw `Article` text
    by stripping gender tokens and version tokens (REG-010/014/015 and
    onward all key off this same (base, gender) tuple). Reuses the token
    sets from the OTHER workstream's config.py (Key-Article Margin
    Analysis) rather than redefining them, since both workstreams parse
    the same underlying Article-naming convention.
  - Two Core-scope exclusion rules exist because two different raw exports
    encode "which account this sale belongs to" differently (REG-008):
      * SS26 exports (SS26_DTC_Wholesale_w34_data_*.xlsx) -- substring
        match on the Sales Market/Store or Sales Market field. Only XXL
        and China are independently verifiable in this field; Zalando /
        Zalando Marketplace / Stadium Outlet do not appear as distinct
        market values in these exports.
      * bob_salesdata_*.xlsx exports -- exact match on the Customer Group
        field, which does carry all 5 named accounts (XXL SE/NO/FI,
        ZALANDO, ZALANDO MARKETPLACE, China, STADIUM Outlet -- carefully
        distinguished from the legitimate STADIUM account).
    Use is_core_market() for the former, is_core_customer_group() for the
    latter. Do not use one on the other file's fields.
  - MIN_RELIABLE (50,000 SEK/year) blanks/excludes GM%, growth, and share
    calculations with an unreliable small denominator (REG-004's
    precedent, reused everywhere downstream).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))
from config import GENDER_TOKENS, VERSION_TOKENS  # noqa: E402

TIERING_DIR = REPO_ROOT / "data" / "ss26_portfolio_tiering"
INPUTS_DIR = TIERING_DIR / "inputs"
WORKBOOK_PATH = TIERING_DIR / "Project_Bob_Portfolio_Tiering_08092026.xlsx"
FULL_PORTFOLIO_SHEET = "2. Full Portfolio (1860)"
GENERATION_SHEET = "3. Generation Detail"
STOCK_SHEET = "4. Stock by Tier"
FIRST_DATA_ROW = 5  # row 4 is the header row on every sheet built this way

# SS26 export files (SS26YTD / FY25 Jan-Aug / FY25 Sep-Dec), used by most
# of the update scripts as the raw source of Article/SCO-level detail.
SS26_EXPORTS = [
    TIERING_DIR / "SS26_DTC_Wholesale_w34_data_1_SS26YTD.xlsx",
    TIERING_DIR / "SS26_DTC_Wholesale_w34_data_2_FY25_Jan_to_Aug.xlsx",
    TIERING_DIR / "SS26_DTC_Wholesale_w34_data_3_FY25_Sep_to_Dec.xlsx",
]
# The FY25 Sep-Dec file's market column is named differently from the
# other two -- a known export inconsistency (REG-007).
SS26_MARKET_FIELD = {
    "SS26_DTC_Wholesale_w34_data_1_SS26YTD.xlsx": "Sales Market/Store",
    "SS26_DTC_Wholesale_w34_data_2_FY25_Jan_to_Aug.xlsx": "Sales Market/Store",
    "SS26_DTC_Wholesale_w34_data_3_FY25_Sep_to_Dec.xlsx": "Sales Market",
}

MIN_RELIABLE = 50_000  # SEK/year -- REG-004 precedent

# --------------------------------------------------------------------------
# REG-024: the base 8-tier commercial portfolio rule. Confirmed 9-Sep-2026
# by the business (previously undocumented anywhere in this repo -- an
# extensive reverse-engineering attempt from Sales/GM%/Growth/Pace alone had
# failed, best composite score only explained 41% of the real Hero+Near-Hero
# set). Validated against Sheet 2's own "Original Tier (pre-consolidation)"
# column: 1,808 of 1,812 non-Clearance franchises match exactly (99.8%) when
# recomputed from bob_salesdata_2024/2025.xlsx -- the 4 mismatches are
# article-name parsing edge cases ("*MISSING*", a glued gender suffix), not
# rule failures. See scripts/verify_base_tiering.py for the validation run.
# --------------------------------------------------------------------------
MIN_UNITS_FOR_REAL_TRADING_YEAR = 100  # matches config.py's constant of the same name
GM_HERO_FLOOR = 46.6   # core-company average GM% -- the "good" bar, not a round number
GM_WORKHORSE_FLOOR = 40.0  # separate, lower floor -- splits Workhorse/Harvest from Problem Child
HARVEST_GROWTH_LINE = -10.0  # Workhorse (growth > this) vs Harvest (growth <= this)
NEAR_HERO_GROWTH_FLOOR = 15.0


def classify_tier(*, sales, units_prior, units_current, growth, gm_pct,
                   wholesale_sek, dtc_sek, sales_floor_hero, sales_floor_near_hero,
                   wholesale_floor, dtc_floor):
    """The confirmed base tier rule (REG-024), parameterized so the same
    logic can be reused for the Global scale (pass the real 2M/1M/100K/50K
    SEK floors) or a region-scaled scale (pass floors scaled to that
    region's revenue share -- REG-023's region-relative approach applies
    scaling to these SEK floors only; gm_pct/growth are already relative and
    are never scaled). units_prior/units_current and sales/growth/gm_pct/
    wholesale_sek/dtc_sek must all be computed on the SAME year-pair and
    SAME scope (Global sales, or one region's sales) as each other.

    Returns one of: "New / Test", "Exited", "Thin / Immaterial", "Hero",
    "Near-Hero (Rising Star)", "Workhorse", "Harvest (Cash Cow)",
    "Problem Child". Does NOT handle Clearance -- that's a separate,
    SKU-list-based flag carried over from Sheet 2 regardless of scale
    (REG-010's China JV/Zalando lists aren't a sales-rule outcome).
    """
    real_prior = units_prior >= MIN_UNITS_FOR_REAL_TRADING_YEAR
    real_current = units_current >= MIN_UNITS_FOR_REAL_TRADING_YEAR
    if real_current and not real_prior:
        return "New / Test"
    if real_prior and not real_current:
        return "Exited"
    if not real_prior or not real_current:
        return "Thin / Immaterial"

    cross_channel = wholesale_sek > wholesale_floor and dtc_sek > dtc_floor
    if (sales >= sales_floor_hero and growth is not None and growth > 0
            and cross_channel and gm_pct >= GM_HERO_FLOOR):
        return "Hero"
    if (sales >= sales_floor_near_hero and growth is not None and growth > NEAR_HERO_GROWTH_FLOOR
            and cross_channel and gm_pct >= GM_HERO_FLOOR):
        return "Near-Hero (Rising Star)"
    if gm_pct < GM_WORKHORSE_FLOOR:
        return "Problem Child"
    if growth is not None and growth > HARVEST_GROWTH_LINE:
        return "Workhorse"
    return "Harvest (Cash Cow)"

# Tier order, consolidated scheme (REG-013): Hero+Near-Hero merged,
# Workhorse+Harvest merged, Clearance tier (REG-010) added.
TIER_ORDER = [
    "Hero + Near-Hero", "Workhorse + Harvest", "Problem Child", "New / Test",
    "Thin / Immaterial", "Exited", "Clearance — Ex China/Zalando",
]
CLEARANCE_TIER = "Clearance — Ex China/Zalando"

_TOKENS_PATTERN = r"\b(" + "|".join(re.escape(t) for t in (GENDER_TOKENS | VERSION_TOKENS)) + r")\b"
_GENDER_PATTERN = r"\b(" + "|".join(re.escape(t) for t in GENDER_TOKENS) + r")\b"
_VERSION_PATTERN = r"\b(" + "|".join(re.escape(t) for t in VERSION_TOKENS) + r")\b"
# Version tokens that sometimes appear glued to the preceding word with no
# space -- an export artifact (e.g. "Fuse PantII"). Stripped separately
# via a lookbehind/lookahead rather than \b, which requires a boundary.
_GLUED_VERSION_TOKENS = ("III", "II", "IV", "2.0")

RANK = {None: 0, "2.0": 1, "II": 1, "III": 2, "IV": 3}


def base_name(article: str) -> str:
    """Strip gender + version tokens (incl. glued ones) to get the franchise base name."""
    a = str(article).strip()
    a = re.sub(_GENDER_PATTERN, "", a)
    a = re.sub(_VERSION_PATTERN, "", a)
    for t in _GLUED_VERSION_TOKENS:
        a = re.sub(r"(?<=[a-zA-Z])" + re.escape(t) + r"(?=[ A-Z]|$)", "", a)
    a = re.sub(r"\s+", " ", a).strip()
    # fallback: glued gender suffix (export artifact, no space before "Men"/"Women")
    for suf in ("Women", "Men"):
        if a.endswith(suf) and len(a) > len(suf):
            a = a[: -len(suf)].strip()
            break
    return a


def gender(article: str) -> str:
    a = str(article).strip()
    if "Women" in a:
        return "Women"
    if "Men" in a:
        return "Men"
    return "Unisex"


def version_token(article: str) -> str | None:
    """The raw version token found in an article name, or None for the original/unversioned article."""
    a = str(article).strip()
    m = re.search(_VERSION_PATTERN, a)
    if m:
        return m.group(1)
    for t in _GLUED_VERSION_TOKENS:
        if re.search(r"(?<=[a-zA-Z])" + re.escape(t) + r"(?=[ A-Z]|$)", a):
            return t
    return None


def franchise_key(article: str) -> tuple[str, str]:
    """(base, gender) tuple -- the franchise-grain join key used throughout the workbook."""
    return base_name(article), gender(article)


def is_core_market(market) -> bool:
    """Core-scope filter for SS26 export files (Sales Market/Store / Sales Market field).

    Substring match on 'xxl' / 'china' -- REG-008. Only these two of the
    five excluded accounts are independently verifiable in this field;
    Zalando/Zalando Marketplace/Stadium Outlet do not appear distinctly
    here. Use is_core_customer_group() for bob_salesdata files instead.
    """
    if market is None:
        return True
    m = str(market).lower()
    return not any(x in m for x in ("xxl", "china"))


_CUSTOMER_GROUP_EXCLUDED_EXACT = {"ZALANDO", "ZALANDO MARKETPLACE", "China", "STADIUM Outlet"}


def is_core_customer_group(customer_group) -> bool:
    """Core-scope filter for bob_salesdata_*.xlsx files (Customer Group field).

    Exact match against the 5 named accounts (REG-008): XXL SE/NO/FI,
    ZALANDO, ZALANDO MARKETPLACE, China, STADIUM Outlet -- carefully
    distinguished from the legitimate STADIUM account (no "Outlet"
    suffix, which stays Core).
    """
    if customer_group is None:
        return True
    cg = str(customer_group).strip()
    if cg.upper().startswith("XXL "):
        return False
    if cg in _CUSTOMER_GROUP_EXCLUDED_EXACT or cg.upper() in {"ZALANDO", "ZALANDO MARKETPLACE"}:
        return False
    return True


# --------------------------------------------------------------------------
# Workbook I/O helpers
# --------------------------------------------------------------------------

def load_raw_export(path):
    """Load one SS26 export workbook's first sheet. Returns (header_tuple, list_of_row_tuples)."""
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    return header, list(rows)


def load_full_portfolio(workbook_path=WORKBOOK_PATH):
    """Open the tiering workbook and return (wb, ws) for the Full Portfolio sheet, data_only."""
    import openpyxl

    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    ws = wb[FULL_PORTFOLIO_SHEET]
    return wb, ws


def full_portfolio_lookup(ws) -> dict:
    """(base, gender) -> {"row": excel_row_number, "tier": ..., "sales25": ..., ...} for every franchise row."""
    lookup = {}
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        base = ws.cell(row=r, column=1).value
        if base is None:
            continue
        g = (ws.cell(row=r, column=2).value or "").strip()
        lookup[(base.strip(), g)] = {
            "row": r,
            "layer": ws.cell(row=r, column=3).value,
            "tier": ws.cell(row=r, column=4).value,
            "bucket": ws.cell(row=r, column=5).value,
            "sales25": ws.cell(row=r, column=6).value or 0,
            "gm25": ws.cell(row=r, column=7).value,
            "sales_ytd26": ws.cell(row=r, column=8).value or 0,
            "orig_tier": ws.cell(row=r, column=16).value,
        }
    return lookup


# Shared cell styling, so every update script's new columns look identical.
def styles():
    from openpyxl.styles import Font, PatternFill

    return {
        "header_fill": PatternFill(start_color="FF1F3864", end_color="FF1F3864", fill_type="solid"),
        "header_font": Font(name="Arial", bold=True, color="FFFFFFFF"),
        "body_font": Font(name="Arial", bold=False),
        "gray_font": Font(name="Arial", bold=False, color="FF555555"),
        "pct_fmt": '0.0"%"',
        "num_fmt": "#,##0",
    }
