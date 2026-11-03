"""
Pulls the numbers for the workshop fact sheets (Part 1: clearing the
end-of-life overstock; Part 2: product lifecycle) into one JSON, from this
workstream's own outputs -- no new source data.

Part 1 reads build_clearance_view.build() (IP-006 to IP-008). "Excess" is
new here: per style, Available Stock beyond what sells in the 1-Dec-2026 to
1-Sep-2027 window at the 2025 weekly rate of sale (IP-007/IP-008), valued
at the style's average unit landed cost. Styles under MIN_RELIABLE 2025
sales have no rate, so no excess -- counted separately as "no reliable
rate", not as zero. Same optimism caveat as IP-008: 2025 was mostly
full-price selling while the style was active.

Part 2 reads outputs/Franchise_Generation_Timeline.xlsx (IP-005), whose
dates are proxies from the season tag, not launch dates.

Run: python3 build_fact_sheet_data.py   (after build_clearance_view.py)
Output: ../outputs/fact_sheet_data.json
"""
from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_clearance_view as cv  # noqa: E402
from build_inventory_review import build as build_inventory, load_active_range_lookup  # noqa: E402

IP_DIR = Path(__file__).resolve().parents[1]
TIMELINE_PATH = IP_DIR / "outputs" / "Franchise_Generation_Timeline.xlsx"
OUT_PATH = IP_DIR / "outputs" / "fact_sheet_data.json"

WINDOW_WEEKS = (cv.WINDOW_END - list(cv.WINDOWS.values())[0]).days / 7
SEASON_INDEX = {s: i for i, s in enumerate(cv.SEASON_ORDER)}


def part1() -> dict:
    colours = cv.build()
    styles: dict[str, dict] = defaultdict(lambda: {
        "units": 0, "cost": 0.0, "wp": 0.0, "seasons": defaultdict(int)})
    for c in colours:
        s = styles[c["Style Code"]]
        s.update({"article": c["Article"], "layer": c["Layer"],
                  "seasonality": c["Style Seasonality"],
                  "units25": c["Style 2025 Units Sold"],
                  "cover": c["Style Weeks of Cover"]})
        s["units"] += c["Available Stock"]
        s["cost"] += c["Landed Cost SEK"]
        s["wp"] += c["WP SEK"]
        s["seasons"][c["Season Article"]] += c["Available Stock"]

    for s in styles.values():
        if s["cover"] is None:
            s["excess"] = None
        else:
            s["excess"] = max(0.0, s["units"] - s["units25"] / 52 * WINDOW_WEEKS)
        unit_cost = s["cost"] / s["units"] if s["units"] else 0
        s["excess_cost"] = (s["excess"] or 0) * unit_cost
        s["main_season"] = max(s["seasons"], key=s["seasons"].get)

    def total(key):
        return sum(c[key] for c in colours)

    clears_col = next(k for k in colours[0] if k.startswith("Clears by"))
    outlook = defaultdict(lambda: defaultdict(lambda: [0, 0.0]))
    age_sea = defaultdict(lambda: defaultdict(lambda: [0, 0.0]))
    layers = defaultdict(lambda: {"units": 0, "cost": 0.0, "wp": 0.0,
                                  "excess": 0.0, "excess_cost": 0.0})
    for c in colours:
        status = {"Y": "Clears", "N": "Won't clear"}.get(c[clears_col], "No reliable rate")
        for d, k in ((outlook, status), (age_sea, c["Style Seasonality"])):
            d[c["Age"]][k][0] += c["Available Stock"]
            d[c["Age"]][k][1] += c["WP SEK"]
        lay = layers[c["Layer"]]
        lay["units"] += c["Available Stock"]
        lay["cost"] += c["Landed Cost SEK"]
        lay["wp"] += c["WP SEK"]
    for s in styles.values():
        layers[s["layer"]]["excess"] += s["excess"] or 0
        layers[s["layer"]]["excess_cost"] += s["excess_cost"]
    # Style excess spread over its colours by units, to show it by age/seasonality.
    excess_by = {"age": defaultdict(lambda: [0.0, 0.0]), "seasonality": defaultdict(lambda: [0.0, 0.0])}
    for c in colours:
        s = styles[c["Style Code"]]
        if not s["excess"]:
            continue
        share = c["Available Stock"] / s["units"]
        for dim, key in (("age", c["Age"]), ("seasonality", c["Style Seasonality"])):
            excess_by[dim][key][0] += s["excess"] * share
            excess_by[dim][key][1] += s["excess_cost"] * share

    sku_rows = [r for r in build_inventory()[0] if r["In Active Range"] == "N"]
    broken = [c for c in colours if "/" in c["Sizes In Stock / Listed"]
              and (lambda a, b: a * 2 < b)(*map(int, c["Sizes In Stock / Listed"].split("/")))]

    top = sorted(styles.items(), key=lambda kv: -kv[1]["excess_cost"])[:15]
    return {
        "window_weeks": round(WINDOW_WEEKS, 1),
        "kpis": {
            "units": total("Available Stock"),
            "landed_cost": total("Landed Cost SEK"),
            "wp": total("WP SEK"),
            "headroom": 1 - total("Landed Cost SEK") / total("WP SEK"),
            "style_codes": len(styles),
            "colours": len(colours),
            "excess_units": round(sum(s["excess"] or 0 for s in styles.values())),
            "excess_cost": round(sum(s["excess_cost"] for s in styles.values())),
            "no_rate_units": sum(s["units"] for s in styles.values() if s["excess"] is None),
            "undelivered_purchase_units": sum(r["Purchase SS27 (Undelivered)"] for r in sku_rows),
            "broken_size_units": sum(c["Available Stock"] for c in broken),
        },
        "age_order": cv.AGE_BUCKETS,
        "age_by_seasonality": {a: {k: v[0] for k, v in d.items()} for a, d in age_sea.items()},
        "outlook_by_age": {a: {k: {"units": v[0], "wp": round(v[1])} for k, v in d.items()}
                           for a, d in outlook.items()},
        "layers": sorted(({"layer": k, **{kk: round(vv) for kk, vv in v.items()},
                           "headroom": 1 - v["cost"] / v["wp"] if v["wp"] else None}
                          for k, v in layers.items()), key=lambda r: -r["cost"]),
        "excess_by_age": {k: {"units": round(v[0]), "cost": round(v[1])}
                          for k, v in excess_by["age"].items()},
        "excess_by_seasonality": {k: {"units": round(v[0]), "cost": round(v[1])}
                                  for k, v in excess_by["seasonality"].items()},
        "undelivered_in_available_units": sum(
            r["Purchase SS27 (Undelivered)"] for r in sku_rows if r["Season Article"]),
        "top_styles_by_excess_cost": [{
            "code": code, "article": s["article"], "layer": s["layer"],
            "main_season": s["main_season"], "seasonality": s["seasonality"],
            "units": s["units"], "excess": round(s["excess"] or 0),
            "excess_cost": round(s["excess_cost"]), "weeks_cover": s["cover"],
            "headroom": 1 - s["cost"] / s["wp"] if s["wp"] else None,
        } for code, s in top],
    }


def part2() -> dict:
    ws = openpyxl.load_workbook(TIMELINE_PATH, read_only=True, data_only=True)["Multi-Generation Only"]
    franchises: dict[str, list] = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        franchises[r[0]].append({"generation": r[1], "codes": r[2], "active": r[3] == "Y",
                                 "earliest": r[4], "latest": r[5], "available": r[6] or 0})
    # Active = any generation active in the stock file's own codes, OR the
    # franchise base is in the active range (matches the Range Review Board's 16).
    active_bases = {v["base"] for v in load_active_range_lookup().values()}
    active, gaps = [], []
    for name, gens in franchises.items():
        seasons = sorted(SEASON_INDEX[g["earliest"]] for g in gens if g["earliest"])
        gaps += [b - a for a, b in zip(seasons, seasons[1:])]
        if any(g["active"] for g in gens) or name in active_bases:
            active.append({"franchise": name, "generations": gens,
                           "old_generation_stock": sum(g["available"] for g in gens if not g["active"])})
    active.sort(key=lambda f: -f["old_generation_stock"])

    _, summary = build_inventory()
    left_behind = [s for s in summary if s["In Active Range"] == "Y"
                   and s["Available Stock -- NOT Active-Range SKUs (List 3 candidate)"] > 0]
    return {
        "multi_generation_franchises": len(franchises),
        "active_multi_generation": len(active),
        "discontinued_multi_generation": len(franchises) - len(active),
        "gap_seasons": sorted(gaps),
        "active": active,
        "eol_units_in_active_franchises": sum(
            s["Available Stock -- NOT Active-Range SKUs (List 3 candidate)"] for s in left_behind),
        "active_franchises_with_eol_stock": len(left_behind),
        "top_left_behind": [{"franchise": s["Franchise"], "layer": s["Layer"],
                             "units": s["Available Stock -- NOT Active-Range SKUs (List 3 candidate)"]}
                            for s in sorted(left_behind, key=lambda s: -s[
                                "Available Stock -- NOT Active-Range SKUs (List 3 candidate)"])[:10]],
    }


if __name__ == "__main__":
    data = {"part1": part1(), "part2": part2()}
    OUT_PATH.write_text(json.dumps(data, indent=1, ensure_ascii=False, default=str))
    print(json.dumps({k: data["part1"]["kpis"] for k in ["kpis"]}, indent=1))
    print({k: v for k, v in data["part2"].items() if k not in ("active", "top_left_behind")})
