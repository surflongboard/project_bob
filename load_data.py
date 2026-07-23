"""
Load and clean the Bob sales data (bob_salesdata_2024.xlsx / bob_salesdata_2025.xlsx).

Usage:
    from load_data import load_all
    full = load_all("data/bob_salesdata_2024.xlsx", "data/bob_salesdata_2025.xlsx")

Returns a single tidy DataFrame, one row per
(Year, Sales Channel, Sales Market, Customer Group, Layer, Article, Ordertype Group)
with the raw metric columns untouched (Garp SEK Sales, Garp SEK Margin,
Garp SEK COGS, Units Sold, etc).

ASSUMPTIONS (see config.py for the reasoning, this file just applies them):
  - Each source workbook has a single sheet named after its year ("2024", "2025").
  - A row with Year == "Total" is a workbook-level subtotal row and is dropped
    (verified: dropping it does not change any other row; the detail rows sum
    exactly to it in both source files as of the 2026-07 data pull).
  - Article/Layer/Sales Channel/Sales Market are stripped of whitespace and
    have known casing inconsistencies fixed (see config.LAYER_NAME_FIXES).
  - Missing Customer Group (blank cells) are relabeled '*BLANK*' rather than
    dropped -- these are real transactions (own retail / some e-com rows) and
    should stay in every total, just excluded from named-account breakdowns.
  - Missing Ordertype Group similarly becomes '*BLANK*' (a handful of rows,
    immaterial in every layer checked so far, but not silently dropped).
"""
from __future__ import annotations
import pandas as pd
from config import LAYER_NAME_FIXES


def _load_one_year(path: str, year: int) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=str(year))
    df = df[df["Year"] != "Total"].copy()
    df["Year"] = year

    # NOTE: astype(str).str.strip() on a column with real NaNs is a pandas
    # gotcha -- in this pandas version it silently turns NaN back into a
    # float NaN after the .str.strip() call rather than the string "nan".
    # fillna() BEFORE casting avoids it. (~900 rows/year have NaN Article,
    # mostly under Layer == '*MISSING*' -- likely blank/placeholder rows in
    # the source extract; kept rather than dropped, but tagged so they don't
    # silently break downstream string ops.)
    df["Article"] = df["Article"].fillna("*MISSING*").astype(str).str.strip()
    df["Layer"] = df["Layer"].fillna("*MISSING*").astype(str).str.strip().replace(LAYER_NAME_FIXES)
    df["Sales Channel"] = df["Sales Channel"].fillna("*MISSING*").astype(str).str.strip()
    df["Sales Market"] = df["Sales Market"].fillna("*MISSING*").astype(str).str.strip()
    df["Customer Group"] = df["Customer Group"].fillna("*BLANK*").astype(str).str.strip()
    df["Ordertype Group"] = df["Ordertype Group"].fillna("*BLANK*").astype(str).str.strip()
    return df


def load_all(path_2024: str, path_2025: str) -> pd.DataFrame:
    df24 = _load_one_year(path_2024, 2024)
    df25 = _load_one_year(path_2025, 2025)
    full = pd.concat([df24, df25], ignore_index=True)
    return full


def sanity_check(full: pd.DataFrame) -> None:
    """Quick print-out to confirm the load matches expectations. Run this
    after every new data drop before trusting any downstream analysis."""
    print("Rows by year:")
    print(full.groupby("Year").size())
    print("\nTotal sales by year (SEK):")
    print(full.groupby("Year")["Garp SEK Sales"].sum())
    print("\nLayers present:")
    print(sorted(full["Layer"].dropna().unique().tolist()))


if __name__ == "__main__":
    full = load_all("data/bob_salesdata_2024.xlsx", "data/bob_salesdata_2025.xlsx")
    sanity_check(full)
    full.to_pickle("output/full.pkl")
    print("\nSaved output/full.pkl")
