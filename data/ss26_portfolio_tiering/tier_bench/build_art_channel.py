"""
Rebuilds art_channel.pkl -- the per-article FY25 Wholesale/Retail/E-com
Sales+Margin lookup that build_qa_data.py needs to compute
channel_mix.json and generation_comparison.json's channel-split fields.

This intermediate previously existed only as a one-off chat script whose
own source was never saved (see tier_bench/README.md's "What's still not
reproducible" -- now closed). Rebuilt from scratch 9-Sep-2026, following
the same pattern already used throughout this workstream (REG-014/017/018):
read the two FY25-half SS26 exports (data_2 Jan-Aug, data_3 Sep-Dec --
NOT data_1, which is SS26YTD/FY26, out of scope for a FY25 figure),
Core-scope filter each file's own market field (`ss26_lib.is_core_market`,
REG-008), and aggregate Garp SEK Sales / Garp SEK Margin per
(base, gender, article) x Sales Channel.

`Sales Channel` on these exports is a clean 3-way field (Wholesale /
Retail / E-com, verified -- no "Marketplace" or other value to fold in,
unlike REG-018's bob_salesdata Sales Channel field, a different source).

Output: art_channel.pkl, {(base, gender, article): {channel: [sales, margin]}},
matching build_qa_data.py's expected shape exactly (a 2-element sales/margin
list per channel, only for channels with any Core-scope activity).

Usage (from this directory):
    python3 build_art_channel.py
"""
from __future__ import annotations

import pickle
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import ss26_lib as lib

FY25_EXPORTS = [
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_2_FY25_Jan_to_Aug.xlsx",
    lib.TIERING_DIR / "SS26_DTC_Wholesale_w34_data_3_FY25_Sep_to_Dec.xlsx",
]


def main():
    # (base, gender, article) -> {channel: [sales, margin]}
    art_channel: dict[tuple[str, str, str], dict[str, list[float]]] = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))

    total_rows = core_rows = 0
    channels_seen = set()
    for path in FY25_EXPORTS:
        header, rows = lib.load_raw_export(path)
        idx = {n: i for i, n in enumerate(header)}
        market_field = lib.SS26_MARKET_FIELD[path.name]
        for row in rows:
            total_rows += 1
            article = row[idx["Article"]]
            if not article:
                continue
            if not lib.is_core_market(row[idx[market_field]]):
                continue
            core_rows += 1
            a = article.strip()
            base, gend = lib.franchise_key(a)
            channel = row[idx["Sales Channel"]]
            channels_seen.add(channel)
            sales = row[idx["Garp SEK Sales"]] or 0
            margin = row[idx["Garp SEK Margin"]] or 0
            rec = art_channel[(base, gend, a)][channel]
            rec[0] += sales
            rec[1] += margin

    print(f"raw rows: {total_rows}  Core-scope rows: {core_rows}")
    print("channels seen:", sorted(channels_seen))
    print("distinct (base, gender, article) keys:", len(art_channel))

    # plain dict, not defaultdicts, before pickling -- avoid pickling a
    # lambda-based defaultdict (unpicklable) and match the plain-dict
    # shape build_qa_data.py's channel_breakdown() expects.
    out = {key: dict(chans) for key, chans in art_channel.items()}

    out_path = Path(__file__).resolve().parent / "art_channel.pkl"
    with open(out_path, "wb") as f:
        pickle.dump(out, f)
    print("written", out_path)


if __name__ == "__main__":
    main()
