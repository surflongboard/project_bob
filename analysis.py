"""
Reusable analysis functions for the Key-Article Margin Analysis.

Every function takes the full tidy DataFrame (see load_data.load_all) plus a
layer name, and returns a plain pandas DataFrame -- no plotting, no writing to
disk here, so these compose easily in a notebook, a script, or an xlsx-writer.

Import and call what you need, e.g.:

    from load_data import load_all
    from analysis import article_summary, account_view, channel_view, \
        ordertype_view, reorder_channel_split, find_version_families

    full = load_all("data/bob_salesdata_2024.xlsx", "data/bob_salesdata_2025.xlsx")
    shell = full[full["Layer"] == "Shell"]
    print(article_summary(shell).head(15))
"""
from __future__ import annotations
import re
import pandas as pd
from config import (
    classify_lifecycle, VERSION_TOKENS, GENDER_TOKENS,
    MIN_FAMILY_SALES_FY25, classify_transition, MIN_SALES_FOR_RELIABLE_MARGIN_PCT,
)


def _yearly(df: pd.DataFrame, group_col: str, year: int) -> pd.DataFrame:
    d = df[df["Year"] == year]
    g = d.groupby(group_col).agg(
        Sales=("Garp SEK Sales", "sum"),
        Margin=("Garp SEK Margin", "sum"),
        Units=("Units Sold", "sum"),
        COGS=("Garp SEK COGS", "sum"),
    ).reset_index()
    return g


def article_summary(layer_df: pd.DataFrame, fy_current: int = 2025, fy_prior: int = 2024) -> pd.DataFrame:
    """Article-level FY-over-FY comparison with lifecycle classification.
    Sorted descending by current-year sales."""
    cur = _yearly(layer_df, "Article", fy_current).add_suffix(f"_{fy_current}").rename(
        columns={f"Article_{fy_current}": "Article"})
    pri = _yearly(layer_df, "Article", fy_prior).add_suffix(f"_{fy_prior}").rename(
        columns={f"Article_{fy_prior}": "Article"})
    m = cur.merge(pri, on="Article", how="left").fillna(0)
    m["Status"] = m.apply(
        lambda r: classify_lifecycle(r[f"Units_{fy_prior}"], r[f"Units_{fy_current}"]), axis=1)
    m[f"GM_{fy_current}"] = (m[f"Margin_{fy_current}"] / m[f"Sales_{fy_current}"]).where(
        m[f"Sales_{fy_current}"] >= MIN_SALES_FOR_RELIABLE_MARGIN_PCT)
    m[f"GM_{fy_prior}"] = (m[f"Margin_{fy_prior}"] / m[f"Sales_{fy_prior}"]).where(
        m[f"Sales_{fy_prior}"] >= MIN_SALES_FOR_RELIABLE_MARGIN_PCT)
    return m.sort_values(f"Sales_{fy_current}", ascending=False)


def account_view(layer_df: pd.DataFrame, fy_current: int = 2025, fy_prior: int = 2024, top_n: int = 15) -> pd.DataFrame:
    cur = _yearly(layer_df, "Customer Group", fy_current).add_suffix(f"_{fy_current}").rename(
        columns={f"Customer Group_{fy_current}": "Account"})
    pri = _yearly(layer_df, "Customer Group", fy_prior).add_suffix(f"_{fy_prior}").rename(
        columns={f"Customer Group_{fy_prior}": "Account"})
    m = cur.merge(pri, on="Account", how="outer").fillna(0)
    m[f"GM_{fy_current}"] = m[f"Margin_{fy_current}"] / m[f"Sales_{fy_current}"].replace(0, pd.NA)
    m[f"GM_{fy_prior}"] = m[f"Margin_{fy_prior}"] / m[f"Sales_{fy_prior}"].replace(0, pd.NA)
    m["MarginPtChg"] = m[f"GM_{fy_current}"] - m[f"GM_{fy_prior}"]
    return m.sort_values(f"Sales_{fy_current}", ascending=False).head(top_n)


def channel_view(layer_df: pd.DataFrame, fy_current: int = 2025, fy_prior: int = 2024) -> pd.DataFrame:
    cur = _yearly(layer_df, "Sales Channel", fy_current).add_suffix(f"_{fy_current}").rename(
        columns={f"Sales Channel_{fy_current}": "Channel"})
    pri = _yearly(layer_df, "Sales Channel", fy_prior).add_suffix(f"_{fy_prior}").rename(
        columns={f"Sales Channel_{fy_prior}": "Channel"})
    m = cur.merge(pri, on="Channel", how="outer").fillna(0)
    m[f"GM_{fy_current}"] = m[f"Margin_{fy_current}"] / m[f"Sales_{fy_current}"].replace(0, pd.NA)
    m[f"GM_{fy_prior}"] = m[f"Margin_{fy_prior}"] / m[f"Sales_{fy_prior}"].replace(0, pd.NA)
    m["MarginPtChg"] = m[f"GM_{fy_current}"] - m[f"GM_{fy_prior}"]
    return m.sort_values(f"Sales_{fy_current}", ascending=False)


def ordertype_view(layer_df: pd.DataFrame, fy_current: int = 2025, fy_prior: int = 2024) -> pd.DataFrame:
    cur = _yearly(layer_df, "Ordertype Group", fy_current).add_suffix(f"_{fy_current}").rename(
        columns={f"Ordertype Group_{fy_current}": "Ordertype"})
    pri = _yearly(layer_df, "Ordertype Group", fy_prior).add_suffix(f"_{fy_prior}").rename(
        columns={f"Ordertype Group_{fy_prior}": "Ordertype"})
    m = cur.merge(pri, on="Ordertype", how="outer").fillna(0)
    m[f"GM_{fy_current}"] = m[f"Margin_{fy_current}"] / m[f"Sales_{fy_current}"].replace(0, pd.NA)
    m[f"GM_{fy_prior}"] = m[f"Margin_{fy_prior}"] / m[f"Sales_{fy_prior}"].replace(0, pd.NA)
    m["MarginPtChg"] = m[f"GM_{fy_current}"] - m[f"GM_{fy_prior}"]
    m["ShareCurrent"] = m[f"Sales_{fy_current}"] / m[f"Sales_{fy_current}"].sum()
    return m.sort_values(f"Sales_{fy_current}", ascending=False)


def reorder_channel_split(layer_df: pd.DataFrame, fy_current: int = 2025, fy_prior: int = 2024,
                           channels=("Wholesale", "E-com")) -> pd.DataFrame:
    """The Reorder-Group-only view, split by channel. See config.py section 5
    for why this matters -- blended Reorder can mask a Wholesale-specific
    problem behind E-com's stability."""
    ro = layer_df[(layer_df["Ordertype Group"] == "Reorder Group") & (layer_df["Sales Channel"].isin(channels))]
    cur = _yearly(ro, "Sales Channel", fy_current).add_suffix(f"_{fy_current}").rename(
        columns={f"Sales Channel_{fy_current}": "Channel"})
    pri = _yearly(ro, "Sales Channel", fy_prior).add_suffix(f"_{fy_prior}").rename(
        columns={f"Sales Channel_{fy_prior}": "Channel"})
    m = cur.merge(pri, on="Channel", how="outer").fillna(0)
    m[f"GM_{fy_current}"] = m[f"Margin_{fy_current}"] / m[f"Sales_{fy_current}"].replace(0, pd.NA)
    m[f"GM_{fy_prior}"] = m[f"Margin_{fy_prior}"] / m[f"Sales_{fy_prior}"].replace(0, pd.NA)
    m["MarginPtChg"] = m[f"GM_{fy_current}"] - m[f"GM_{fy_prior}"]
    m[f"COGSu_{fy_current}"] = m[f"COGS_{fy_current}"] / m[f"Units_{fy_current}"]
    m[f"COGSu_{fy_prior}"] = m[f"COGS_{fy_prior}"] / m[f"Units_{fy_prior}"]
    m["COGSuChgPct"] = m[f"COGSu_{fy_current}"] / m[f"COGSu_{fy_prior}"] - 1
    return m


# ---------------------------------------------------------------------------
# Version-family / lifecycle-transition detection
# ---------------------------------------------------------------------------

def _base_name(article: str) -> str:
    a = str(article).strip()
    tokens = GENDER_TOKENS | VERSION_TOKENS
    # Some source rows are missing the space before a trailing gender token
    # (e.g. "ROC Flash Down HoodMen" instead of "... Hood Men") -- without this,
    # such rows fail to strip the token and form their own spurious one-off
    # "family", splitting what should be one continuing article/franchise in two.
    for t in tokens:
        a = re.sub(rf"(?<=[a-z])({re.escape(t)})\b", r" \1", a)
    tokens_pattern = r"\b(" + "|".join(re.escape(t) for t in tokens) + r")\b"
    a = re.sub(tokens_pattern, "", a)
    a = re.sub(r"\s+", " ", a).strip()
    return a


def _gender(article: str) -> str:
    if "Women" in article:
        return "Women"
    if "Men" in article:
        return "Men"
    return "Unisex"


def find_version_families(layer_df: pd.DataFrame, fy_current: int = 2025, fy_prior: int = 2024,
                           min_family_sales: float = MIN_FAMILY_SALES_FY25) -> pd.DataFrame:
    """Detect article families that appear to be mid-version-transition
    (e.g. 'Astral GTX Jacket Men' / 'Astral GTX II Jacket Men'), and classify
    each as CLEAN CUTOVER / PARTIAL OVERLAP / HEAVY OVERLAP based on how much
    of the family's current-year units the OLDER-looking article still holds.

    Caveat (see config.py #3): pure string-matching -- will not catch
    transitions with a real name change, only version-token successions.
    """
    d = layer_df.copy()
    d["Base"] = d["Article"].apply(_base_name)
    d["Gender"] = d["Article"].apply(_gender)

    g = d.groupby(["Base", "Gender", "Article", "Year"]).agg(
        Sales=("Garp SEK Sales", "sum"), Units=("Units Sold", "sum")).reset_index()
    piv = g.pivot_table(index=["Base", "Gender", "Article"], columns="Year",
                         values=["Sales", "Units"], fill_value=0)
    piv.columns = [f"{a}_{b}" for a, b in piv.columns]
    piv = piv.reset_index()

    counts = piv.groupby(["Base", "Gender"])["Article"].nunique()
    families = counts[counts > 1].index

    rows = []
    for base, gend in families:
        sub = piv[(piv.Base == base) & (piv.Gender == gend)]
        tot_current = sub.get(f"Sales_{fy_current}", pd.Series(dtype=float)).sum()
        if tot_current < min_family_sales:
            continue
        # crude heuristic: the article with the lower current-year units among
        # the family, if there are exactly two, is treated as "old"; this is a
        # simplification and should be eyeballed for 3+ version families
        sub = sub.sort_values(f"Units_{fy_current}", ascending=False)
        for _, row in sub.iterrows():
            rows.append({
                "Base": base, "Gender": gend, "Article": row["Article"],
                f"Sales_{fy_prior}": row.get(f"Sales_{fy_prior}", 0),
                f"Units_{fy_prior}": row.get(f"Units_{fy_prior}", 0),
                f"Sales_{fy_current}": row.get(f"Sales_{fy_current}", 0),
                f"Units_{fy_current}": row.get(f"Units_{fy_current}", 0),
                "FamilySalesCurrent": tot_current,
            })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out["ShareOfFamilyUnitsCurrent"] = out[f"Units_{fy_current}"] / out.groupby(["Base", "Gender"])[f"Units_{fy_current}"].transform("sum").replace(0, pd.NA)
    return out.sort_values(["FamilySalesCurrent", "Base", "Gender"], ascending=[False, True, True])
