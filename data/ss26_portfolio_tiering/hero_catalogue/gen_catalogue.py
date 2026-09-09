import json

with open("catalogue_images_b64.json") as f:
    IMAGES = json.load(f)

# 34 Hero + Near-Hero franchises, sales-descending within tier (matches the published workbook order)
PRODUCTS = [
    # key, base, gender, layer, tier, orig_tier, sales25, url, image_key, status, status_note
    dict(key="astral_gtx_jacket_w", base="Astral GTX Jacket", gender="Women", layer="Shell", tier="Hero", sales=10811041,
         url="https://www.haglofs.com/en/women/tops-women/tops-jackets-women/astral-gtx-jacket-women-6046694HQ", img="astral_gtx_jacket_w", status="ok"),
    dict(key="sarna_mimic_hood_m", base="Särna Mimic Hood", gender="Men", layer="Insulation", tier="Hero", sales=8042848,
         url="https://www.haglofs.com/en/men/tops-men/tops-jackets-men/sarna-mimic-hood-men-6051312C5", img="sarna_mimic_hood_m", status="ok"),
    dict(key="sarna_mimic_hood_w", base="Särna Mimic Hood", gender="Women", layer="Insulation", tier="Hero", sales=6679131,
         url="https://www.haglofs.com/en/women/activities-women/sarna-mimic-hood-women-6051322C5", img="sarna_mimic_hood_w", status="ok"),
    dict(key="long_down_parka_w", base="Long Down Parka", gender="Women", layer="Insulation", tier="Hero", sales=6341899,
         url="https://www.haglofs.com/en/women/tops-women/tops-jackets-women/long-down-parka-women-6054122C5", img="long_down_parka_w", status="ok"),
    dict(key="tight_large", base="Tight Large", gender="Unisex", layer="Daypacks", tier="Hero", sales=4959238,
         url="https://www.haglofs.com/en/equipment/backpacks-bags/tight-large-3381502C5", img="tight_large", status="ok"),
    dict(key="buteo_mid_jacket_w", base="Buteo Mid Jacket", gender="Women", layer="Mid layer", tier="Hero", sales=4680074,
         url="https://www.haglofs.com/en/women/activities-women/buteo-mid-jacket-women-6050742C5", img="buteo_mid_jacket_w", status="ok"),
    dict(key="roc_flash_down_hood_w", base="ROC Flash Down Hood", gender="Women", layer="Insulation", tier="Hero", sales=4494173,
         url="https://www.haglofs.com/en/women/activities-women/roc-flash-down-hood-women-6074662C5", img="roc_flash_down_hood_w", status="ok"),
    dict(key="korp_proof_jacket_m", base="Korp Proof Jacket", gender="Men", layer="Shell", tier="Hero", sales=3886774,
         url="https://www.haglofs.com/en/men/tops-men/tops-jackets-men/korp-proof-jacket-men-6061322C5", img="korp_proof_jacket_m", status="ok"),
    dict(key="astral_gtx_pant_w", base="Astral GTX Pant", gender="Women", layer="Shell", tier="Hero", sales=3536140,
         url="https://www.haglofs.com/de/damen/hosen-damen/hosen-lange-hosen-damen/astral-gtx-pant-women-6047502C5", img="astral_gtx_pant_w", status="ok"),
    dict(key="salix_proof_mimic_parka_m", base="Salix Proof Mimic Parka", gender="Men", layer="Insulation", tier="Hero", sales=3359528,
         url="https://www.haglofs.com/en/men/tops-men/tops-jackets-men/salix-proof-mimic-parka-men-6054192C5", img="salix_proof_mimic_parka_m",
         status="successor", note="Photo shows the “Salix Proof Mimic II Parka” — the current successor, not the original SKU with the sales history shown here."),
    dict(key="rosson_down_hood_w", base="Rosson Down Hood", gender="Women", layer="Insulation", tier="Hero", sales=2920336,
         url="https://www.haglofs.com/en/women/tops-women/tops-jackets-women/rosson-down-hood-women-6074682C5", img="rosson_down_hood_w", status="ok"),
    dict(key="latnja_gtx_pant_m", base="Latnja GTX Insulated Pant", gender="Men", layer="Shell", tier="Hero", sales=2607570,
         url="https://www.haglofs.com/en/men/latnja-gtx-insulated-pant-men-6074702C5", img="latnja_gtx_pant_m", status="ok"),
    dict(key="chaos_gtx_jacket_w", base="Chaos GTX Jacket", gender="Women", layer="Shell", tier="Hero", sales=2602402,
         url="https://www.haglofs.com/en/women/tops-women/chaos-gtx-jacket-women-6056584T6", img="chaos_gtx_jacket_w",
         status="successor", note="Photo (via Intersport) is labeled “Chaos II GTX Jacket Women” — the successor, not the original SKU with the sales history shown here. An earlier search on this one turned up a Youth (kids') version instead — wrong product, correctly not used."),
    dict(key="rosson_mid_jacket_m", base="Rosson Mid Jacket", gender="Men", layer="Mid layer", tier="Hero", sales=2410730,
         url="https://www.haglofs.com/en/men/tops-men/tops-fleece-midlayers-men/rosson-mid-jacket-men-6074732C5", img="rosson_mid_jacket_m", status="ok"),
    dict(key="korp_proof_jacket_w", base="Korp Proof Jacket", gender="Women", layer="Shell", tier="Hero", sales=2319994,
         url="https://www.haglofs.com/en/women/activities-women/korp-proof-jacket-women-6062192C5", img="korp_proof_jacket_w", status="ok"),
    dict(key="moran_softshell_pant_m", base="Morän Softshell Standard Pant", gender="Men", layer="Soft shell", tier="Hero", sales=2118928,
         url="https://www.haglofs.com/en/men/bottoms-men/bottoms-trousers-men/moran-softshell-standard-pant-men-6054022C5", img="moran_softshell_pant_m",
         status="ok"),
    dict(key="vassi_gtx_pro_jacket_w", base="Vassi GTX Pro Jacket", gender="Women", layer="Shell", tier="Hero", sales=2015935,
         url="https://www.haglofs.com/en/women/activities-women/vassi-gtx-pro-jacket-women-6046894WY", img="vassi_gtx_pro_jacket_w", status="ok"),
    # Near-Hero
    dict(key="jarve_multi_28", base="Jarve Multi 28", gender="Unisex", layer="Daypacks", tier="Near-Hero", sales=1977999,
         url="https://www.haglofs.com/it/men/jarve-multi-28-6069552C5", img="jarve_multi_28",
         status="variant", note="Page title says “Jarve Multi 28” but the URL slug reads “jarve-multi-31” — worth confirming this is the 28L, not the 31L variant."),
    dict(key="roc_flash_down_jacket_m", base="ROC Flash Down Jacket", gender="Men", layer="Insulation", tier="Near-Hero", sales=1922685,
         url="https://www.haglofs.com/en/activities/mountaineering/roc-flash-down-jacket-men-6074625Q1", img="roc_flash_down_jacket_m", status="ok"),
    dict(key="latnja_gtx_pant_w", base="Latnja GTX Insulated Pant", gender="Women", layer="Shell", tier="Near-Hero", sales=1838416,
         url="https://www.haglofs.com/en/activities/ski-snowboarding/ski-snowboarding-trousers-pulloversshorts/latnja-gtx-insulated-pant-women-6074722C5", img="latnja_gtx_pant_w", status="ok"),
    dict(key="lim_fuse_pant_m", base="L.I.M Fuse Pant", gender="Men", layer="Legwear", tier="Near-Hero", sales=1704736,
         url="https://www.haglofs.com/en/men/bottoms-men/lim-fuse-pant-men_54315-6069422C5", img="lim_fuse_pant_m", status="ok"),
    dict(key="lim_mid_multi_hood_m", base="L.I.M Mid Multi Hood", gender="Men", layer="Mid layer", tier="Near-Hero", sales=1557975,
         url="https://www.haglofs.com/en/outlet/lim-mid-multi-hood-men_60217-6070722C5", img="lim_mid_multi_hood_m",
         status="successor", note="Photo (via a third-party retailer) shows the “L.I.M Mid Multi II Hood” successor, not the original SKU."),
    dict(key="lim_fuse_shorts_w", base="L.I.M Fuse Shorts", gender="Women", layer="Legwear", tier="Near-Hero", sales=1451642,
         url="https://www.haglofs.com/en/outlet/outlet-outlet-women/lim-fuse-shorts-women_43776-6053073X3", img="lim_fuse_shorts_w",
         status="successor", note="Photo is labeled “L.I.M Fuse II Shorts Women” on the source page — the successor, not the original SKU with the sales history shown here."),
    dict(key="fjatla_60", base="Fjatla 60", gender="Unisex", layer="Bags", tier="Near-Hero", sales=1410361,
         url="https://www.haglofs.com/de/damen/aktivitaten-damen/lava-recycled-poly-60-6069892C5", img="fjatla_60", status="ok"),
    dict(key="lava_30", base="Lava 30", gender="Unisex", layer="Bags", tier="Near-Hero", sales=1359786,
         url="https://www.haglofs.com/en/women/activities-women/lava-30-3393642C5", img="lava_30", status="ok"),
    dict(key="mossa_pile_jacket_w", base="Mossa Pile Jacket", gender="Women", layer="Mid layer", tier="Near-Hero", sales=1356082,
         url="https://www.haglofs.com/en/activities/hiking/pile-jacket-women-6065384VY", img="mossa_pile_jacket_w", status="ok"),
    dict(key="tarius_5", base="Tarius -5", gender="Unisex", layer="Sleepingbags", tier="Near-Hero", sales=1348641,
         url="https://www.haglofs.com/en/women/equipments-accessories-women/tarius-5-4161504GW", img="tarius_5", status="ok"),
    dict(key="roc_flash_down_vest_m", base="ROC Flash Down Vest", gender="Men", layer="Insulation", tier="Near-Hero", sales=1287921,
         url="https://www.haglofs.com/en/activities/mountaineering/roc-flash-down-vest-men-6074612C5", img="roc_flash_down_vest_m", status="ok"),
    dict(key="mimic_alert_jacket_m", base="Mimic Alert Jacket", gender="Men", layer="Insulation", tier="Near-Hero", sales=1272732,
         url="https://www.haglofs.com/en/men/tops-men/tops-jackets-men/mimic-alert-jacket-men-6074462C5", img="mimic_alert_jacket_m", status="ok"),
    dict(key="rugged_slim_pant_w", base="Rugged Slim Pant", gender="Women", layer="Legwear", tier="Near-Hero", sales=1201422,
         url="https://www.haglofs.com/en/women/bottoms-women/bottoms-trousers-women/rugged-slim-pant-women-6051652C5", img="rugged_slim_pant_w",
         status="ok"),
    dict(key="lim_mid_multi_hood_w", base="L.I.M Mid Multi Hood", gender="Women", layer="Mid layer", tier="Near-Hero", sales=1179954,
         url="https://www.haglofs.com/de/aktivitaten/bergsport/lim-mid-multi-ii-hood-women-6076342AT", img="lim_mid_multi_hood_w",
         status="successor", note="Photo shows the “L.I.M Mid Multi II Hood” — the current successor, not the original SKU with the sales history shown here."),
    dict(key="rosson_mid_jacket_w", base="Rosson Mid Jacket", gender="Women", layer="Mid layer", tier="Near-Hero", sales=1154524,
         url="https://www.haglofs.com/en/women/activities-women/rosson-mid-jacket-women-6076062C5", img="rosson_mid_jacket_w", status="ok"),
    dict(key="korp_proof_pant_m", base="Korp Proof Pant", gender="Men", layer="Shell", tier="Near-Hero", sales=1104990,
         url="https://www.haglofs.com/en/men/bottoms-men/bottoms-trousers-men/korp-proof-pant-men_148139-6080952C5", img="korp_proof_pant_m",
         status="successor", note="Photo is explicitly labeled “Korp Proof II Pant Men” on the source page — the successor, not the original SKU with the sales history shown here."),
    dict(key="lim_proof_pant_w", base="L.I.M Proof Pant", gender="Women", layer="Shell", tier="Near-Hero", sales=1014937,
         url="https://www.haglofs.com/be/en-be/l.i.m-proof-pant-women/p/604508-2C5.html", img="lim_proof_pant_w",
         status="ok"),
]

assert len(PRODUCTS) == 34
n_hero = sum(1 for p in PRODUCTS if p["tier"] == "Hero")
n_photographed = sum(1 for p in PRODUCTS if p["img"])
print("Hero:", n_hero, "Near-Hero:", 34 - n_hero, "photographed:", n_photographed)

# merge in FY25/YTD26 sales + GM% from the published workbook (Sheet 2, cols F/G/H/I/J),
# extracted separately into franchise_financials.json -- see chat for the extraction script.
with open("franchise_financials.json") as f:
    FIN = json.load(f)
for p in PRODUCTS:
    fin = FIN[f'{p["base"]}|{p["gender"]}']
    p["gm25"] = fin["gm25"]
    p["units25"] = fin["units25"]
    p["salesYtd26"] = fin["salesYtd26"]
    p["unitsYtd26"] = fin["unitsYtd26"]
    p["gmYtd26"] = fin["gmYtd26"]
    p["gp25"] = p["sales"] * p["gm25"] / 100
    p["gpYtd26"] = p["salesYtd26"] * p["gmYtd26"] / 100

LAYER_ORDER = ["Insulation", "Shell", "Mid layer", "Daypacks", "Legwear", "Bags", "Sleepingbags", "Soft shell"]

STATUS_LABEL = {
    "ok": None,
    "successor": "SUCCESSOR SHOWN",
    "variant": "VARIANT UNCERTAIN",
    "notify": "OUT OF STOCK",
    "missing": None,
}

def money(n):
    return f"{n:,.0f}"

def money_short(n):
    a = abs(n)
    if a >= 1_000_000:
        return f"{n/1_000_000:.2f}M"
    if a >= 1_000:
        return f"{n/1_000:.0f}K"
    return f"{n:.0f}"

def pct(n):
    return f"{n:.1f}%"

PRODUCTS_JSON = json.dumps([{k: v for k, v in p.items() if k != "img"} | {"hasImg": bool(p["img"])} for p in PRODUCTS])

with open("catalogue_template.html") as f:
    template = f.read()

# build card markup, grouped and sorted by layer (see LAYER_ORDER)
def card_html(p):
    img_src = f'data:image/jpeg;base64,{IMAGES[p["img"]]}'
    tier_cls = "tier-hero" if p["tier"] == "Hero" else "tier-near"
    return f'''
    <article class="card" id="card-{p['key']}" data-key="{p['key']}">
      <div class="card-media">
        <img src="{img_src}" alt="{p['base']} {p['gender']}" loading="lazy">
      </div>
      <div class="card-body">
        <div class="card-top">
          <span class="tier-chip {tier_cls}">{p['tier']}</span>
        </div>
        <h3 class="card-title">{p['base']}</h3>
        <p class="card-meta">{p['gender']} &middot; {p['layer']}</p>
        <div class="fin-table">
          <div class="fin-row fin-head"><span></span><span>Sales (Units)</span><span>GM%</span></div>
          <div class="fin-row"><span>FY25</span><span>{money_short(p['sales'])} ({p['units25']:,.0f})</span><span>{pct(p['gm25'])}</span></div>
          <div class="fin-row"><span>YTD26</span><span>{money_short(p['salesYtd26'])} ({p['unitsYtd26']:,.0f})</span><span>{pct(p['gmYtd26'])}</span></div>
        </div>
        <div class="decision" role="group" aria-label="Decision for {p['base']}">
          <button type="button" class="decision-btn keep-btn" data-action="keep">Keep</button>
          <button type="button" class="decision-btn merge-btn" data-action="merge">Merge</button>
          <button type="button" class="decision-btn update-btn" data-action="update">Update</button>
        </div>
        <a class="decision-link" href="{p['url']}" target="_blank" rel="noopener">Source ↗</a>
      </div>
    </article>'''

tier_rank = {"Hero": 0, "Near-Hero": 1}

def layer_section(layer_name):
    items = [p for p in PRODUCTS if p["img"] and p["layer"] == layer_name]
    items.sort(key=lambda p: (tier_rank[p["tier"]], -p["sales"]))
    cards = "\n".join(card_html(p) for p in items)
    n_hero_l = sum(1 for p in items if p["tier"] == "Hero")
    n_near_l = len(items) - n_hero_l
    return f'''
  <div class="section-head">
    <h2>{layer_name}</h2>
    <span class="section-count">{len(items)} franchise{"s" if len(items) != 1 else ""} &middot; {n_hero_l} Hero, {n_near_l} Near-Hero</span>
  </div>
  <div class="grid">
{cards}
  </div>'''

present_layers = [l for l in LAYER_ORDER if any(p["layer"] == l and p["img"] for p in PRODUCTS)]
layer_sections = "\n".join(layer_section(l) for l in present_layers)

# ---- known-issues summary: every flagged (non-"ok") photographed card, one row each ----
FLAG_CLASS = {"successor": "flag-successor", "variant": "flag-variant", "notify": "flag-notify"}

def issue_row(p):
    badge = f'<span class="issue-badge {FLAG_CLASS[p["status"]]}">{STATUS_LABEL[p["status"]]}</span>'
    return f'''<tr>
      <td>{badge}</td>
      <td><b>{p['base']}</b><br><span class="card-meta">{p['gender']} &middot; {p['tier']} &middot; {p['layer']}</span></td>
      <td class="issue-note">{p['note']}</td>
      <td><a class="jump-link" href="#card-{p['key']}">Jump to card &uarr;</a></td>
    </tr>'''

flagged = [p for p in PRODUCTS if p["img"] and p["status"] != "ok"]
flagged.sort(key=lambda p: (tier_rank[p["tier"]], -p["sales"]))
n_successor = sum(1 for p in flagged if p["status"] == "successor")
n_notify = sum(1 for p in flagged if p["status"] == "notify")
n_variant = sum(1 for p in flagged if p["status"] == "variant")

if flagged:
    issue_rows = "\n".join(issue_row(p) for p in flagged)
    breakdown_parts = []
    if n_successor: breakdown_parts.append(f'{n_successor} successor-generation photo{"s" if n_successor != 1 else ""}')
    if n_notify: breakdown_parts.append(f'{n_notify} out-of-stock reference{"s" if n_notify != 1 else ""}')
    if n_variant: breakdown_parts.append(f'{n_variant} uncertain variant{"s" if n_variant != 1 else ""}')
    breakdown = ", ".join(breakdown_parts)
    issues_section = f'''
  <div class="issues-panel">
    <div class="issues-head">
      <h2>Known issues</h2>
      <span class="issues-sub"><span class="num">{len(flagged)}</span> of {n_photographed} photographed cards flagged &mdash; {breakdown}. Cards above are shown clean; jump links here take you back to the flagged one.</span>
    </div>
    <div class="issues-table-wrap">
      <table class="issues">
        <thead><tr><th>Flag</th><th>Franchise</th><th>Note</th><th></th></tr></thead>
        <tbody>
{issue_rows}
        </tbody>
      </table>
    </div>
  </div>'''
else:
    issues_section = ""

def missing_row(p):
    name_cell = p['base']
    if p.get("missing_note"):
        name_cell += f'<div class="missing-note">{p["missing_note"]}</div>'
    return f'''<tr>
      <td>{name_cell}</td><td>{p['gender']}</td><td>{p['layer']}</td>
      <td class="num">{money(p['sales'])}</td>
      <td><a href="{p['url']}" target="_blank" rel="noopener">product page ↗</a></td>
    </tr>'''

missing_hero = "\n".join(missing_row(p) for p in PRODUCTS if not p["img"] and p["tier"] == "Hero")
missing_near = "\n".join(missing_row(p) for p in PRODUCTS if not p["img"] and p["tier"] == "Near-Hero")
n_missing = 34 - n_photographed

if n_missing == 0:
    missing_section = '''
  <div class="complete-banner">
    <span class="cb-mark">✓</span>
    <p><b>All 34 Hero and Near-Hero franchises photographed.</b> 6 cards show a "II"/successor generation rather than the exact SKU a sales figure is attributed to — see the flagged cards and the note below before treating this as a finished, presentable set.</p>
  </div>'''
else:
    missing_section = f'''
  <div class="section-head">
    <h2>Still needed</h2>
    <span class="section-count"><span class="num">{n_missing}</span> franchises, no photo yet</span>
  </div>
  <div class="missing-table-wrap">
    <table class="missing">
      <thead><tr><th colspan="5">Hero</th></tr></thead>
      <tbody>
{missing_hero}
      </tbody>
      <thead><tr><th colspan="5">Near-Hero</th></tr></thead>
      <tbody>
{missing_near}
      </tbody>
    </table>
  </div>'''

html = template
html = html.replace("__ISSUES_SECTION__", issues_section)
html = html.replace("__LAYER_SECTIONS__", layer_sections)
html = html.replace("__MISSING_SECTION__", missing_section)
html = html.replace("__PRODUCTS_JSON__", PRODUCTS_JSON)
html = html.replace("__N_PHOTOGRAPHED__", str(n_photographed))
html = html.replace("__N_TOTAL__", "34")

with open("hero_catalogue.html", "w") as f:
    f.write(html)
print("written, bytes:", len(html))
