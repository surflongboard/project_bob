import json
from pathlib import Path

qa_dir = str(Path(__file__).resolve().parent / "qa_data") + "/"

franchises = []
for i in range(5):
    franchises += json.load(open(f"{qa_dir}franchises_chunk_{i}.json"))["rows"]
gen_comp = json.load(open(f"{qa_dir}generation_comparison.json"))["rows"]
channel_mix = json.load(open(f"{qa_dir}channel_mix.json"))["rows"]
tier_summary = json.load(open(f"{qa_dir}tier_summary.json"))["rows"]
stock = json.load(open(f"{qa_dir}stock_by_franchise.json"))["rows"]
country = json.load(open(f"{qa_dir}country_breakdown.json"))["rows"]
regional_tiering = json.load(open(f"{qa_dir}regional_tiering.json"))["rows"]

print(len(franchises), len(gen_comp), len(channel_mix), len(tier_summary), len(country), len(regional_tiering))

FRANCHISES_JSON = json.dumps(franchises, separators=(",", ":"))
GEN_COMP_JSON = json.dumps(gen_comp, separators=(",", ":"))
CHANNEL_MIX_JSON = json.dumps(channel_mix, separators=(",", ":"))
TIER_SUMMARY_JSON = json.dumps(tier_summary, separators=(",", ":"))
STOCK_JSON = json.dumps(stock, separators=(",", ":"))
COUNTRY_JSON = json.dumps(country, separators=(",", ":"))
REGIONAL_TIERING_JSON = json.dumps(regional_tiering, separators=(",", ":"))

print("bytes:", len(FRANCHISES_JSON), len(GEN_COMP_JSON), len(CHANNEL_MIX_JSON), len(TIER_SUMMARY_JSON), len(COUNTRY_JSON), len(REGIONAL_TIERING_JSON))

html = r"""<title>Tier Bench</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Big+Shoulders+Display:wght@600;700;800&family=IBM+Plex+Sans:ital,wght@0,400;0,500;0,600;1,400&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
<style>
:root{
  --bg:#F1F3EF;
  --surface:#FFFFFF;
  --surface-2:#E7EAE2;
  --surface-3:#DCE0D6;
  --ink:#182420;
  --ink-muted:#5B6960;
  --ink-faint:#8A968D;
  --line:#D6DACE;
  --line-strong:#BFC5B6;
  --accent:#2B6E5E;
  --accent-strong:#194A3E;
  --accent-soft:#DCEAE4;
  --warn:#B5651D;
  --warn-soft:#F5E4D1;
  --warn-strong:#8A4A12;
  --ok:#4C7A3F;
  --ok-soft:#E3ECDD;
  --user-bubble:#1F3A32;
  --user-ink:#F2F5F1;
  --shadow: 0 1px 2px rgba(24,36,32,0.05), 0 8px 24px -10px rgba(24,36,32,0.16);
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#0F1512;
    --surface:#171F1B;
    --surface-2:#1E2721;
    --surface-3:#28322B;
    --ink:#E8ECE6;
    --ink-muted:#9AA69E;
    --ink-faint:#69756D;
    --line:#2B352E;
    --line-strong:#3A453C;
    --accent:#5FC7AE;
    --accent-strong:#8EDCC8;
    --accent-soft:#1C3630;
    --warn:#E29A54;
    --warn-soft:#3A2A18;
    --warn-strong:#F2B879;
    --ok:#9CC17E;
    --ok-soft:#243422;
    --user-bubble:#274038;
    --user-ink:#EAF1EC;
    --shadow: 0 1px 2px rgba(0,0,0,0.35), 0 10px 28px -10px rgba(0,0,0,0.55);
  }
}
:root[data-theme="dark"]{
  --bg:#0F1512;
  --surface:#171F1B;
  --surface-2:#1E2721;
  --surface-3:#28322B;
  --ink:#E8ECE6;
  --ink-muted:#9AA69E;
  --ink-faint:#69756D;
  --line:#2B352E;
  --line-strong:#3A453C;
  --accent:#5FC7AE;
  --accent-strong:#8EDCC8;
  --accent-soft:#1C3630;
  --warn:#E29A54;
  --warn-soft:#3A2A18;
  --warn-strong:#F2B879;
  --ok:#9CC17E;
  --ok-soft:#243422;
  --user-bubble:#274038;
  --user-ink:#EAF1EC;
  --shadow: 0 1px 2px rgba(0,0,0,0.35), 0 10px 28px -10px rgba(0,0,0,0.55);
}
*{box-sizing:border-box;}
body{
  margin:0; background:var(--bg); color:var(--ink);
  font-family:'IBM Plex Sans', -apple-system, sans-serif;
  font-size:15px; line-height:1.55; -webkit-font-smoothing:antialiased;
}
.display{font-family:'Big Shoulders Display', 'Arial Narrow', sans-serif; text-transform:uppercase;}
.mono, code, .num{font-family:'IBM Plex Mono', ui-monospace, Menlo, monospace; font-variant-numeric:tabular-nums;}
.wrap{max-width:980px; margin:0 auto; padding:36px 22px 64px;}

/* ---- header ---- */
header{display:flex; flex-direction:column; gap:6px; margin-bottom:22px;}
.eyebrow{
  font-family:'IBM Plex Mono', monospace; font-size:11px; letter-spacing:0.14em; text-transform:uppercase;
  color:var(--accent); display:flex; align-items:center; gap:8px;
}
.eyebrow::before{content:"";width:7px;height:7px;background:var(--accent);display:inline-block;}
h1.display{
  font-weight:800; font-size:clamp(34px,6vw,52px); line-height:0.92; letter-spacing:0.01em;
  margin:2px 0 0; color:var(--ink); text-wrap:balance;
}
.subtitle{color:var(--ink-muted); font-size:15px; max-width:60ch; margin:8px 0 0;}
.subtitle a{color:var(--accent-strong); font-weight:500;}

/* ---- tier strip ---- */
.tier-strip{
  display:grid; grid-template-columns:repeat(4,1fr); gap:1px;
  background:var(--line); border:1px solid var(--line); margin:24px 0 26px;
  border-radius:2px; overflow:hidden;
}
.tier-tile{background:var(--surface); padding:13px 14px; display:flex; flex-direction:column; gap:4px; min-width:0;}
.tier-tile .t-name{
  font-family:'IBM Plex Mono',monospace; font-size:10.5px; letter-spacing:0.03em; color:var(--ink-faint);
  text-transform:uppercase; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
}
.tier-tile .t-sales{font-family:'Big Shoulders Display'; font-weight:700; font-size:20px; color:var(--ink);}
.tier-tile .t-gm{font-size:12px; color:var(--ink-muted);}
.tier-tile.warn .t-name{color:var(--warn-strong);}
@media (max-width:680px){ .tier-strip{grid-template-columns:repeat(2,1fr);} }

/* ---- chat panel ---- */
.panel{
  background:var(--surface); border:1px solid var(--line); border-radius:3px; box-shadow:var(--shadow);
  display:flex; flex-direction:column; overflow:hidden;
}
.panel-head{
  padding:14px 18px; border-bottom:1px solid var(--line); display:flex; align-items:center; justify-content:space-between; gap:12px;
  background:var(--surface-2);
}
.panel-head .label{font-family:'Big Shoulders Display'; font-weight:700; font-size:16px; letter-spacing:0.02em;}
.panel-head .count{font-size:12px; color:var(--ink-faint);}

.chips{display:flex; flex-wrap:wrap; gap:8px; padding:14px 18px; border-bottom:1px solid var(--line);}
.chip{
  background:var(--surface-2); border:1px solid var(--line-strong); color:var(--ink-muted);
  font-size:12.5px; padding:7px 12px; border-radius:20px; cursor:pointer; text-align:left;
  transition:background 0.12s, color 0.12s, border-color 0.12s;
}
.chip:hover{background:var(--accent-soft); border-color:var(--accent); color:var(--accent-strong);}

#transcript{padding:18px; display:flex; flex-direction:column; gap:16px; min-height:180px; max-height:56vh; overflow-y:auto;}
.msg{display:flex; flex-direction:column; gap:6px; max-width:88%;}
.msg.user{align-self:flex-end; align-items:flex-end;}
.msg.assistant{align-self:flex-start; align-items:flex-start;}
.bubble{padding:11px 14px; border-radius:3px; font-size:14.5px; line-height:1.6;}
.msg.user .bubble{background:var(--user-bubble); color:var(--user-ink); border-radius:10px 10px 2px 10px;}
.msg.assistant .bubble{background:var(--surface-2); color:var(--ink); border:1px solid var(--line); border-radius:2px 10px 10px 10px; width:100%;}
.msg.assistant .bubble.thinking{color:var(--ink-faint); font-style:italic;}
.bubble p{margin:0 0 10px;}
.bubble p:last-child{margin-bottom:0;}
.bubble ul, .bubble ol{margin:0 0 10px; padding-left:20px;}
.bubble strong{color:var(--accent-strong); font-weight:600;}
.bubble h3{font-family:'Big Shoulders Display'; font-size:15px; text-transform:uppercase; letter-spacing:0.02em; margin:14px 0 6px; color:var(--ink);}
.bubble h3:first-child{margin-top:0;}
.tbl-wrap{overflow-x:auto; margin:4px 0 12px; border:1px solid var(--line); border-radius:2px;}
.bubble table{border-collapse:collapse; width:100%; font-size:13px;}
.bubble th{
  background:var(--surface-3); text-align:left; padding:7px 10px; font-family:'IBM Plex Mono',monospace;
  font-size:10.5px; text-transform:uppercase; letter-spacing:0.03em; color:var(--ink-muted); white-space:nowrap;
}
.bubble td{padding:7px 10px; border-top:1px solid var(--line); font-variant-numeric:tabular-nums; white-space:nowrap;}
.bubble tr:nth-child(even) td{background:var(--surface);}
.err-note{color:var(--warn-strong); font-size:12.5px;}

.export-bar{display:flex; flex-wrap:wrap; gap:8px; margin-top:2px;}
.export-btn{
  background:var(--surface); border:1px solid var(--line-strong); color:var(--ink-muted);
  font-family:'IBM Plex Mono', monospace; font-size:11.5px; padding:5px 10px; border-radius:20px;
  cursor:pointer; transition:background 0.12s, color 0.12s, border-color 0.12s;
}
.export-btn:hover:not(:disabled){background:var(--accent-soft); border-color:var(--accent); color:var(--accent-strong);}
.export-btn:disabled{cursor:default; opacity:0.75;}
.export-btn:focus-visible{outline:2px solid var(--accent); outline-offset:2px;}

.composer{display:flex; gap:10px; padding:14px 18px; border-top:1px solid var(--line); background:var(--surface-2);}
.composer textarea{
  flex:1; resize:none; border:1px solid var(--line-strong); border-radius:3px; padding:10px 12px;
  font-family:inherit; font-size:14px; background:var(--surface); color:var(--ink); line-height:1.5;
}
.composer textarea:focus{outline:2px solid var(--accent); outline-offset:1px;}
.composer button{
  font-family:'Big Shoulders Display'; font-weight:700; letter-spacing:0.03em; text-transform:uppercase;
  background:var(--accent); color:#fff; border:none; border-radius:3px; padding:0 20px; font-size:14px;
  cursor:pointer; transition:background 0.12s;
}
.composer button:hover:not(:disabled){background:var(--accent-strong);}
.composer button:disabled{background:var(--surface-3); color:var(--ink-faint); cursor:not-allowed;}
.composer button:focus-visible, .chip:focus-visible, textarea:focus-visible{outline:2px solid var(--accent); outline-offset:2px;}

/* ---- care label footer ---- */
.care-label{
  margin-top:26px; border:1px dashed var(--line-strong); border-radius:3px; padding:16px 18px;
  background:var(--surface); font-size:12px; color:var(--ink-muted); line-height:1.6;
}
.care-label .cl-head{
  font-family:'IBM Plex Mono',monospace; font-size:10.5px; letter-spacing:0.1em; text-transform:uppercase;
  color:var(--ink-faint); margin-bottom:8px; display:flex; align-items:center; gap:8px;
}
.care-label .cl-head::before{content:"⚠"; font-size:11px;}
.care-label b{color:var(--ink);}
.care-label a{color:var(--accent-strong);}
.care-label .cl-row{display:flex; gap:8px; margin-bottom:4px;}
.care-label .cl-row .sym{font-family:'IBM Plex Mono',monospace; color:var(--ink-faint); flex:none;}

@media (prefers-reduced-motion:no-preference){
  .chip, .composer button{transition-duration:0.12s;}
}
</style>

<div class="wrap">
  <header>
    <div class="eyebrow">Project Bob · SS26 Portfolio Tiering</div>
    <h1 class="display">Tier Bench</h1>
    <p class="subtitle">Ask this the way you'd ask an analyst — margin gaps, generation comparisons, channel mix, clearance exposure. It queries the current tiering dataset directly instead of you pivoting the workbook. Methodology lives in the <a href="__REGISTER_URL__" target="_blank" rel="noopener">Assumptions Register</a>.</p>
  </header>

  <div class="tier-strip" id="tierStrip"></div>

  <div class="panel">
    <div class="panel-head">
      <span class="label">Ask the data</span>
      <span class="count" id="dataCount"></span>
    </div>
    <div class="chips" id="chips"></div>
    <div id="transcript"></div>
    <form class="composer" id="composer">
      <textarea id="question" rows="2" placeholder="e.g. Which Problem Child franchises have the biggest margin gap between generations?"></textarea>
      <button type="submit" id="sendBtn">Ask</button>
    </form>
  </div>

  <div class="care-label">
    <div class="cl-head">Data care instructions</div>
    <div class="cl-row"><span class="sym">§</span><span><b>Franchise figures</b> (Sheet 2 basis, 1,860 rows) are the published, audited Sales_2025 / GM% — Confirmed methodology.</span></div>
    <div class="cl-row"><span class="sym">§</span><span><b>Generation &amp; channel-mix figures</b> are recomputed directly from raw SKU-level exports — directional, not audited (REG-012: an unresolved ordertype gap means these won't always sum to the Sheet-2 totals). Treat rankings and comparisons as reliable; treat exact SEK as approximate.</span></div>
    <div class="cl-row"><span class="sym">§</span><span><b>Units_2025</b> (REG-020) is real, not estimated — summed from the source exports' own Units Sold field — but shares the same raw-recomputed caveat as the two rows above (not blanked below the 50K threshold, since it's a count, not a rate).</span></div>
    <div class="cl-row"><span class="sym">§</span><span><b>Style Code(s)</b> (REG-021) is the raw exports' Model field, comma-separated when a franchise has more than one (~23% do — a re-launch can get a new Model number under the same product name). Not a single canonical style code; read it as "codes seen for this franchise."</span></div>
    <div class="cl-row"><span class="sym">§</span><span><b>Country breakdown</b> (REG-022) is hand-mapped from raw store/market names (e.g. "Outlet Barkarby" → Sweden) since the raw data has no clean country field — same raw-recomputed caveat as generation/channel figures. <b>Region Group = Nordic vs Non-Nordic</b> (Sweden/Norway/Denmark/Finland vs everywhere else) is the standard exclusion split, confirmed with the business 9-Sep-2026; a stricter Scandinavian-only (no Finland) flag is also available if specifically asked for. ~4% of rows ("Pop-Up Sales Haglöfs" + "Export Other") carry no city/country in the raw label at all and are marked "Unmapped," excluded from either side of the Nordic split rather than counted as Non-Nordic.</span></div>
    <div class="cl-row"><span class="sym">§</span><span><b>Regional tiers</b> (REG-023/024) recompute the SAME confirmed rule as globalTier, but ranked within Nordic-only or Non-Nordic-only sales, with each region's SEK floors scaled by its own ~68%/~32% share of FY25 Core revenue. <b>regionalTier and globalTier are meant to differ sometimes</b> — a franchise ranking Hero within a smaller regional population while ranking Workhorse globally is the intended effect, not a data error. Region-scoped sales/GM%/growth come from bob_salesdata_2024/2025.xlsx, not the same source as globalTier's figures.</span></div>
    <div class="cl-row"><span class="sym">§</span><span>GM% is blanked below 50,000 SEK/year of sales (REG-004) rather than shown as a wild percentage.</span></div>
    <div class="cl-row"><span class="sym">§</span><span>Answers are generated by Claude reading this dataset live — verify anything going into a decision deck against the source workbook.</span></div>
    <div class="cl-row"><span class="sym">§</span><span>Business shorthand (e.g. "DTC expansion opportunity") is defined in the <a href="__REGISTER_URL__" target="_blank" rel="noopener">register's glossary section</a> — same terms, same meaning here and there.</span></div>
  </div>
</div>

<script>
const FRANCHISES = __FRANCHISES_JSON__;
const GEN_COMPARISON = __GEN_COMP_JSON__;
const CHANNEL_MIX = __CHANNEL_MIX_JSON__;
const TIER_SUMMARY = __TIER_SUMMARY_JSON__;
const STOCK = __STOCK_JSON__;
const COUNTRY = __COUNTRY_JSON__;
const REGIONAL_TIERING = __REGIONAL_TIERING_JSON__;

document.getElementById('dataCount').textContent =
  FRANCHISES.length.toLocaleString() + ' franchises · ' + GEN_COMPARISON.length + ' generation pairs · ' + CHANNEL_MIX.length + ' articles · ' + STOCK.length + ' with stock on hand · ' + COUNTRY.length.toLocaleString() + ' franchise-country rows · ' + (REGIONAL_TIERING.length/2) + ' franchises with regional tiers';

// ---------- tier strip ----------
const tierStrip = document.getElementById('tierStrip');
TIER_SUMMARY.forEach(t => {
  const div = document.createElement('div');
  div.className = 'tier-tile' + (t.tier === 'Problem Child' || t.tier === 'Clearance — Ex China/Zalando' ? ' warn' : '');
  const sales = (t.sales25/1e6).toFixed(1) + 'M';
  div.innerHTML = `<span class="t-name">${t.tier}</span><span class="t-sales num">${sales} SEK</span><span class="t-gm num">${t.n} fr · ${t.gm25}% GM</span>`;
  tierStrip.appendChild(div);
});

// ---------- suggested questions ----------
const SUGGESTIONS = [
  "Which franchises should we clear via sale — dead stock in Thin/Immaterial or Exited?",
  "Which franchises have a DTC expansion opportunity, broken down by tier?",
  "Which franchises have a Wholesale expansion opportunity?",
  "List franchises where the new generation's margin is lower than the old one's, and say whether it's channel mix or a real problem.",
  "Which Problem Child franchises have the biggest margin gap between generations?",
  "Show franchises where the new generation actually has BETTER margin than the old one.",
  "Show the most Wholesale-heavy Hero franchises — where's margin most exposed to channel mix?",
  "Which franchises are Partially flagged for clearance but still have low margin?",
  "Which franchises moved into the Clearance tier, and what tier were they in before?",
  "Which Hero or Near-Hero franchises are missing from the Core Assortment FW27 list?",
  "Which franchises are marked 'Not active' in FW27 Collection status but still have real YTD2026 sales?",
  "Which Problem Child franchises have a real margin gap AND no Core Assortment FW27 inclusion — the clearest fix-or-cut candidates?",
  "Top 10 Workhorse + Harvest franchises by FY25 sales.",
  "Show Hero + Near-Hero franchises in Japan and Germany, excluding Nordic countries.",
  "Which franchises are Hero in Nordic but not Hero globally?",
];
const chipsEl = document.getElementById('chips');
SUGGESTIONS.forEach(q => {
  const b = document.createElement('button');
  b.className = 'chip'; b.type = 'button'; b.textContent = q;
  b.addEventListener('click', () => { document.getElementById('question').value = q; askQuestion(q); });
  chipsEl.appendChild(b);
});

// ---------- filtering helpers ----------
const CONTROL_KEYS = new Set(['sortBy', 'sortDir', 'limit']);
function applyFilters(rows, filters, textFields) {
  let out = rows;
  for (const [k, v] of Object.entries(filters || {})) {
    if (v === undefined || v === null || v === '') continue;
    if (CONTROL_KEYS.has(k)) continue; // not a data field -- sort/paging control, handled by the caller
    if (k.startsWith('min')) {
      const field = k.slice(3, 4).toLowerCase() + k.slice(4);
      out = out.filter(r => typeof r[field] === 'number' && r[field] >= v);
    } else if (k.startsWith('max')) {
      const field = k.slice(3, 4).toLowerCase() + k.slice(4);
      out = out.filter(r => typeof r[field] === 'number' && r[field] <= v);
    } else if (k.endsWith('Contains')) {
      const field = k.slice(0, -8);
      const needle = String(v).toLowerCase();
      out = out.filter(r => r[field] && String(r[field]).toLowerCase().includes(needle));
    } else {
      out = out.filter(r => r[k] === v);
    }
  }
  return out;
}
function sortRows(rows, sortBy, sortDir) {
  if (!sortBy) return rows;
  const dir = sortDir === 'asc' ? 1 : -1;
  return [...rows].sort((a, b) => {
    const av = a[sortBy], bv = b[sortBy];
    if (av == null && bv == null) return 0;
    if (av == null) return 1;
    if (bv == null) return -1;
    if (av < bv) return -1 * dir;
    if (av > bv) return 1 * dir;
    return 0;
  });
}
function cap(rows, limit, max) {
  const n = Math.min(limit || 30, max || 100);
  return rows.slice(0, n);
}

// ---------- tools ----------
// CONSTRAINT (discovered 9-Sep-2026, the hard way): the sample() API enforces
// a 1KB (1024-byte, UTF-8) limit on EACH tool's `description` field. Blowing
// it doesn't just drop/ignore that one tool -- it fails the ENTIRE tool-calling
// request, so every query breaks, including ones that never touch the
// oversized tool. queryRegionalTiering hit this at 1,250 bytes and took down
// all seven tools; trimmed to ~670 bytes (moved the "regionalTier and
// globalTier are different rankings by design" caveat into SYSTEM_PROMPT's
// Rule 4 instead, since the model only needs that guidance once). Before
// adding or expanding any tool description here, check its byte length
// (Buffer.byteLength(desc, 'utf8')) stays well under 1024 -- there's no
// warning from the API before the whole page's chat breaks, just a generic
// failure surfaced through the description-is-at-most-1KB error message.
const TOOLS = [
  {
    name: "queryFranchises",
    description: "Query the 1,860-franchise portfolio table (published Sheet-2 figures: Tier, Sales_2025, Units_2025, GM%_2025, Sales_YTD2026, Units_YTD2026, GM%_YTD2026, Pace%, Growth%, Clearance Flag, Core Assortment FW27, Style Code(s)). Filter, sort, and limit; returns matched rows plus the total match count.",
    inputSchema: {
      type: "object",
      properties: {
        tier: { type: "string", description: "Exact tier: 'Hero + Near-Hero','Workhorse + Harvest','Problem Child','New / Test','Thin / Immaterial','Exited','Clearance — Ex China/Zalando'" },
        layer: { type: "string" },
        gender: { type: "string", description: "'Men','Women','Unisex'" },
        units25: { type: "number", description: "Real FY25 unit sales (REG-020), raw-recomputed from the source exports' own Units Sold field the same way as wholesaleSharePct/channelPattern2025 — NOT blanked below the 50K reliability threshold (it's a count, like sales25, not a rate). null for 333 franchises with no matching raw units at all. Use minUnits25/maxUnits25 to filter." },
        unitsYtd26: { type: "number", description: "Published Sheet-2 YTD2026 unit sales (audited, same basis as sales25/salesYtd26 — NOT the raw-recomputed basis units25 uses). null for 456 franchises with no YTD2026 activity. Use minUnitsYtd26/maxUnitsYtd26 to filter." },
        styleCodesContains: { type: "string", description: "Substring match on Style Code(s) — use to look up a franchise by a known Model/style code. Sourced from the raw exports' Model field (REG-021), comma-separated when a franchise carries more than one (a re-launch can get a new Model number under the same product name — NOT a single canonical code). null for 224 franchises with no matching raw data at all." },
        clearanceFlag: { type: "string", description: "'Fully','Partial','None'" },
        coreAssortmentFw27: { type: "string", description: "'Yes'/'No' — whether the franchise is in the curated FW27 Core Assortment deck (a deliberately narrow best-sellers subset, not the full continuing catalog — REG-015). 'No' on a Hero/Near-Hero franchise is worth flagging for merch review." },
        fw27Collection: { type: "string", description: "'Active'/'Not active'/'Not in review' — from a separate Assortment Attribution Review file's Active flag (REG-016). That field has a KNOWN false-negative bug (successor styles still trading shown Not active) and is kept as-is per business direction — don't treat 'Not active' as a hard fact, especially alongside real YTD2026 sales. 'Not in review' just means this file doesn't cover that franchise (only 292 of 1,860 are in it) — not a negative signal." },
        wholesaleSharePct: { type: "number", description: "Wholesale channel's % share of FY25 Core sales (rest is DTC = Retail+E-com). Raw-recomputed, directional (REG-012/017) — blank/null for 994 franchises (664 below the 50K reliability threshold, 330 with no matching raw sales). Company baseline: Wholesale 40.7% GM%, Retail 48.9%, E-com 61.1%." },
        channelPattern2025: { type: "string", description: "2024→2025 Wholesale-vs-DTC pattern (REG-018), from a DIFFERENT source (bob_salesdata_2024/2025.xlsx, full annual pulls) than wholesaleSharePct. Values: 'Substitution (DTC ↑ / Wholesale ↓)' (135, most common), 'Substitution (Wholesale ↑ / DTC ↓)' (26), 'Co-growth (both ↑)' (66), 'Co-decline (both ↓)' (70), 'Wholesale-only / DTC negligible' (59 — DTC structurally <50K both years, not a data gap), 'DTC-only / Wholesale negligible' (83), 'Insufficient data' (1,413 — genuinely small both channels, or a channel with inconsistent year-to-year reliability), 'Not in FY24/25 dataset' (8)." },
        baseContains: { type: "string", description: "Case-insensitive substring match on franchise name" },
        minSales25: { type: "number" }, maxSales25: { type: "number" },
        minUnits25: { type: "number" }, maxUnits25: { type: "number" },
        minUnitsYtd26: { type: "number" }, maxUnitsYtd26: { type: "number" },
        minGm25: { type: "number" }, maxGm25: { type: "number" },
        sortBy: { type: "string", description: "Any field name, e.g. sales25, gm25, gmYtd26, pace, growth" },
        sortDir: { type: "string", enum: ["asc","desc"] },
        limit: { type: "number", description: "Default 30, max 150" },
      },
    },
    execute(input) {
      let rows = applyFilters(FRANCHISES, input);
      rows = sortRows(rows, input.sortBy, input.sortDir);
      const total = rows.length;
      return { total, rows: cap(rows, input.limit, 150) };
    },
  },
  {
    name: "queryGenerationComparison",
    description: "Query the 109 franchises with a genuine naming-based version succession (an 'Original' article vs. a later II/III/IV/2.0 article under the same franchise). Each row compares the two generations' Sales/GM% (FY25 and YTD2026), the FY25 margin gap, and a channel-mix verdict on WHY: 'mix-driven' (new-gen skews more Wholesale, a lower-margin channel, and margins comparably or better within Wholesale itself), 'real-gap' (still worse even within the same channel), 'inconclusive', or 'new-gen-same-or-better'.",
    inputSchema: {
      type: "object",
      properties: {
        tier: { type: "string" },
        baseContains: { type: "string" },
        verdict: { type: "string", description: "'mix-driven (channel-mix explains most/all of the gap)', 'real-gap (worse even within same channel)', 'inconclusive (insufficient wholesale-channel data)', 'new-gen-same-or-better'" },
        trend: { type: "string", enum: ["holds","reversed"], description: "Whether the FY25 gap direction still holds in YTD2026" },
        minGapFy25: { type: "number", description: "GM point gap = old GM% - new GM%; positive means new generation margins lower" },
        maxGapFy25: { type: "number" },
        sortBy: { type: "string" }, sortDir: { type: "string", enum: ["asc","desc"] },
        limit: { type: "number", description: "Default 40, max 120 (there are only 109 total rows)" },
      },
    },
    execute(input) {
      let rows = applyFilters(GEN_COMPARISON, input);
      rows = sortRows(rows, input.sortBy, input.sortDir);
      const total = rows.length;
      return { total, rows: cap(rows, input.limit, 120) };
    },
  },
  {
    name: "queryChannelMix",
    description: "Query per-article Sales Channel breakdown (Wholesale/Retail/E-com — share of sales and GM% per channel) for the 221 articles that make up the 109 generation-succession franchises. Company-wide FY25 Core GM% baseline: Wholesale 40.7%, Retail 48.9%, E-com 61.1%.",
    inputSchema: {
      type: "object",
      properties: {
        base: { type: "string" }, gender: { type: "string" },
        baseContains: { type: "string" },
        generation: { type: "string", description: "e.g. 'Original (no version token)', 'Version II'" },
        limit: { type: "number", description: "Default 30, max 60" },
      },
    },
    execute(input) {
      let rows = applyFilters(CHANNEL_MIX, input);
      const total = rows.length;
      return { total, rows: cap(rows, input.limit, 60) };
    },
  },
  {
    name: "queryStock",
    description: "Query available warehouse stock (units, REG-019) by franchise — source Available_stock_260825.xlsx, matched to 656 franchises with any stock on hand. UNITS ONLY, no SEK value (source has no cost/price). Each row also carries deadStock (FY25 sales <1,000 SEK — i.e. stock with essentially no sell-through) for quick filtering.",
    inputSchema: {
      type: "object",
      properties: {
        tier: { type: "string" },
        baseContains: { type: "string" },
        deadStock: { type: "boolean", description: "true = only stock with near-zero FY25 sales (the clearest clear-via-sale candidates)" },
        minUnitsInStock: { type: "number" }, maxUnitsInStock: { type: "number" },
        sortBy: { type: "string", description: "e.g. unitsInStock, sales25" }, sortDir: { type: "string", enum: ["asc","desc"] },
        limit: { type: "number", description: "Default 30, max 150 (656 total rows)" },
      },
    },
    execute(input) {
      let rows = applyFilters(STOCK, input);
      rows = sortRows(rows, input.sortBy || "unitsInStock", input.sortDir || "desc");
      const total = rows.length;
      return { total, rows: cap(rows, input.limit, 150) };
    },
  },
  {
    name: "getTierSummary",
    description: "Returns the 7-tier (+TOTAL) portfolio summary: franchise count, Sales_2025, GM%_2025, Sales_YTD2026, GM%_YTD2026, Pace% per tier.",
    inputSchema: { type: "object", properties: {} },
    execute() { return { rows: TIER_SUMMARY }; },
  },
  {
    name: "queryByCountry",
    description: "Query FY25/YTD2026 sales & units per franchise PER COUNTRY (REG-022) — 13,438 franchise-country rows, so a question like 'Hero products in Japan and Germany, excluding Nordic countries' can be answered directly. Sourced from the raw exports' own market/store field, hand-mapped to country (e.g. 'Outlet Barkarby' -> Sweden) since the raw data has no clean country field. sales25/units25 come from the two FY25 exports; salesYtd26/unitsYtd26 come from the SS26YTD export — BOTH are raw-recomputed, NOT the published Sheet-2 totals (REG-012's ordertype-gap caveat applies), so a franchise's rows here won't sum exactly to its audited Sales_2025. A few markets are genuinely regional, not one country: 'Benelux (region)', 'North America (region)', and 'Unmapped / Other' (994+145 rows with no city/country in the raw label at all — don't treat as any specific country, and don't lump into Non-Nordic when totaling a region split).",
    inputSchema: {
      type: "object",
      properties: {
        tier: { type: "string" },
        layer: { type: "string" },
        country: { type: "string", description: "Exact country/market name, e.g. 'Japan', 'Germany', 'China', 'Benelux (region)', 'North America (region)', 'Unmapped / Other'" },
        countryContains: { type: "string" },
        regionGroup: { type: "string", enum: ["Nordic", "Non-Nordic", "Unmapped"], description: "STANDARD grouping to filter/exclude by (business-directed, 9-Sep-2026): Nordic = Sweden/Norway/Denmark/Finland. Use this for any plain 'exclude the Nordic/home market' question. 'Unmapped' is the ~4% of rows with no confirmed country (REG-022) — exclude it from either side of a Nordic vs Non-Nordic split, don't fold it into Non-Nordic." },
        scandinavian: { type: "boolean", description: "NARROWER cut, only if specifically asked for 'Scandinavian' (not 'Nordic'): Sweden/Norway/Denmark only, excluding Finland. Most questions should use regionGroup instead." },
        nordic: { type: "boolean", description: "Same boundary as regionGroup='Nordic' (Sweden/Norway/Denmark/Finland), as a plain Yes/No flag — prefer regionGroup for a three-way split, use this if you just need a simple include/exclude filter." },
        baseContains: { type: "string" },
        minSales25: { type: "number" }, maxSales25: { type: "number" },
        sortBy: { type: "string", description: "e.g. sales25, units25" }, sortDir: { type: "string", enum: ["asc","desc"] },
        limit: { type: "number", description: "Default 40, max 200 (13,438 total rows)" },
      },
    },
    execute(input) {
      let rows = applyFilters(COUNTRY, input);
      rows = sortRows(rows, input.sortBy || "sales25", input.sortDir || "desc");
      const total = rows.length;
      return { total, rows: cap(rows, input.limit, 200) };
    },
  },
  {
    name: "queryRegionalTiering",
    description: "Each franchise's tier recomputed separately within Nordic-only and Non-Nordic-only sales (REG-023/024) — one row per franchise per region (3,720 rows total). Same base rule as globalTier (sales/growth/GM%/cross-channel cuts, REG-024), but each region's SEK floors are scaled by that region's share of FY25 Core sales (Nordic ~68.1%, Non-Nordic ~31.9%) instead of reusing Global's floors unscaled; GM%/growth cuts aren't scaled. regionalTier and globalTier are different rankings by design — mismatches are expected, not errors. sales24/sales25/growth/gm25/wholesale25/dtc25 here are region-scoped (bob_salesdata_2024/2025.xlsx), not queryFranchises' global figures.",
    inputSchema: {
      type: "object",
      properties: {
        base: { type: "string" }, baseContains: { type: "string" },
        gender: { type: "string" }, layer: { type: "string" },
        region: { type: "string", enum: ["Nordic", "Non-Nordic"], description: "Required in spirit — almost every question means one region or the other; omit only to compare across both." },
        regionalTier: { type: "string", description: "The region-scoped tier: 'Hero + Near-Hero','Workhorse + Harvest','Problem Child','New / Test','Thin / Immaterial','Exited','Clearance — Ex China/Zalando'" },
        globalTier: { type: "string", description: "The Sheet-2 published (Global) tier, for comparison — same value set as regionalTier." },
        minSales25: { type: "number" }, maxSales25: { type: "number" },
        minGrowth: { type: "number" }, maxGrowth: { type: "number" },
        minGm25: { type: "number" }, maxGm25: { type: "number" },
        sortBy: { type: "string", description: "e.g. sales25, growth, gm25" }, sortDir: { type: "string", enum: ["asc","desc"] },
        limit: { type: "number", description: "Default 40, max 150 (3,720 total rows across both regions)" },
      },
    },
    execute(input) {
      let rows = applyFilters(REGIONAL_TIERING, input);
      rows = sortRows(rows, input.sortBy || "sales25", input.sortDir || "desc");
      const total = rows.length;
      return { total, rows: cap(rows, input.limit, 150) };
    },
  },
];

const SYSTEM_PROMPT = `You are a portfolio analyst answering questions about Project Bob's SS26 DTC & Wholesale tiering data (an outdoor-apparel brand's franchise-level sales/margin tiering). You have seven tools to query the actual dataset — always call them rather than guessing or using outside knowledge; never invent a number.

Dataset shape:
- queryFranchises: the 1,860-franchise published tiering table (audited FY25 basis).
- queryGenerationComparison: 109 franchises with an Original-vs-newer-version article succession, with a pre-computed channel-mix verdict.
- queryChannelMix: per-article Wholesale/Retail/E-com sales split and GM%, for the same 109 franchises' articles.
- getTierSummary: the 7-tier roll-up.
- queryStock: warehouse available-stock units by franchise (REG-019), with a deadStock flag (near-zero FY25 sales). "What should we clear via sale" = tier is Thin/Immaterial or Exited AND deadStock=true, sorted by unitsInStock. Don't call Problem Child stock a clearance case even if units are high — check sales25 first (a real seller like Front Proof Jacket has both).
- queryByCountry: FY25/YTD2026 sales & units per franchise per country (REG-022), hand-mapped from the raw exports' store/market field since it has no clean country field natively. Use this for any "in [country]" / "excluding [region]" question.
- queryRegionalTiering: each franchise's tier recomputed SEPARATELY within Nordic-only and Non-Nordic-only sales (REG-023/024), one row per franchise per region, with globalTier alongside for direct comparison. Use this for "is X's tier different in Nordic vs globally" / "show Non-Nordic Hero franchises" / "where does regional disagree with global" questions — do NOT use queryFranchises' tier field and call it regional, and do NOT assume regionalTier and globalTier should match (see below).

Known business term mappings — use these exact filters when the question uses this language, don't re-derive:
- "DTC expansion opportunity" / "DTC whitespace" = queryFranchises with channelPattern2025 = "Wholesale-only / DTC negligible" (proven Wholesale demand, DTC negligible both years — see fw27Collection-style caveats don't apply here, this one's a clean structural read). Group by tier when asked "in which tier group."
- "Wholesale expansion opportunity" (the mirror case) = channelPattern2025 = "DTC-only / Wholesale negligible".
- "Nordic countries" / "home market" / "legacy market" (the usual way this comes up) = queryByCountry's regionGroup = "Nordic" (Sweden/Norway/Denmark/Finland) — this is the standard exclusion group, confirmed with the business 9-Sep-2026. Only use the narrower scandinavian flag (Sweden/Norway/Denmark, no Finland) if someone specifically says "Scandinavian" rather than "Nordic." Either way, exclude regionGroup = "Unmapped" from a region split rather than counting it as Non-Nordic — it's ~4% of rows with no confirmed country, not a real geography answer.
- "Nordic tier" / "regional tier" / "how does X do in Nordic vs globally" = queryRegionalTiering, NOT queryByCountry (that's raw sales by country, not a tier) and NOT queryFranchises (that's the single Global tier only).

Rules:
1. Call tools to get real numbers before answering. Use queryGenerationComparison's verdict/trend fields rather than re-deriving them when the question is about generation margin gaps or channel mix — they're already computed correctly.
2. Cite specific figures from tool results. Round sensibly (1 decimal on %, thousands on SEK).
3. Franchise-level (queryFranchises) figures are audited/Confirmed. Generation-, channel-, country-, and regional-tiering-level figures (queryGenerationComparison, queryChannelMix, queryByCountry, queryRegionalTiering) are raw-recomputed and directional (REG-012) — say so briefly if the question turns on an exact SEK figure, but don't over-caveat every answer.
4. queryRegionalTiering's regionalTier and globalTier are DIFFERENT rankings by design, not the same tier at a smaller scale — each region is ranked against its own population with a proportionally-scaled bar (REG-023), so a franchise can legitimately be Hero in one and Workhorse in the other. Never describe a regional/global tier mismatch as an error or inconsistency — it's the intended effect of region-relative thresholds. Say so plainly if a question's premise assumes they should always match.
5. Format the answer in markdown: use a table when listing more than ~3 items, short paragraphs otherwise. Lead with the direct answer, then supporting detail.
6. Default table columns — include these even if the question didn't explicitly ask for them, whenever the tool you called actually has the field (never invent one it doesn't have):
   - queryFranchises rows: always add GM% FY25 (gm25), GM% YTD2026 (gmYtd26), and Sales YTD2026 (salesYtd26) columns alongside whatever the question asked for — this data exists on every row, so default to showing it rather than waiting to be asked.
   - queryByCountry rows: always add Sales YTD2026 and Units YTD2026 alongside FY25 Sales/Units — that data exists per country row too. But GM% is NOT tracked by country in this dataset (only by franchise, globally) — never show a margin column on a country-level table or imply a per-country margin figure. If the question is really about margin in a country, say plainly that margin isn't broken out by country, and offer the franchise's overall (global, all-country) GM% as separate context via queryFranchises, clearly labeled as global rather than country-specific.
   - queryRegionalTiering rows: always add GM% (gm25, already region-scoped for FY25) by default. This dataset has no YTD2026 figures at all (only FY24/FY25) — don't claim a regional YTD2026 sales or margin number; if asked, say the regional cut only goes through FY25 and point to queryFranchises for the global YTD2026 read.
   - queryGenerationComparison and queryChannelMix rows: they already carry both years' GM% (gmOld25/gmNew25 or gm25) and, for queryGenerationComparison, YTD2026 (gmOld26/gmNew26) — keep showing whichever of those the row has; neither carries a YTD2026 figure for queryChannelMix, so don't add one there.
7. If a question can't be answered from these seven tools (e.g. asks about data not covered), say so plainly rather than fabricating.`;

// ---------- tiny markdown renderer ----------
function escapeHtml(s){
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function inline(s){
  s = escapeHtml(s);
  s = s.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  s = s.replace(/`([^`]+)`/g, '<code>$1</code>');
  return s;
}
function renderMarkdown(md){
  const lines = md.split('\n');
  let html = '', i = 0;
  while (i < lines.length) {
    let line = lines[i];
    if (/^\s*$/.test(line)) { i++; continue; }
    if (/^#{1,3}\s/.test(line)) {
      html += `<h3>${inline(line.replace(/^#{1,3}\s/, ''))}</h3>`; i++; continue;
    }
    if (/^\|.*\|\s*$/.test(line) && lines[i+1] && /^\|?\s*:?-{2,}/.test(lines[i+1])) {
      const rows = [];
      while (i < lines.length && /^\|.*\|\s*$/.test(lines[i])) { rows.push(lines[i]); i++; }
      const cells = rows.map(r => r.trim().replace(/^\||\|$/g,'').split('|').map(c => c.trim()));
      const header = cells[0], body = cells.slice(2);
      html += '<div class="tbl-wrap"><table><thead><tr>' +
        header.map(h => `<th>${inline(h)}</th>`).join('') + '</tr></thead><tbody>' +
        body.map(r => '<tr>' + r.map(c => `<td>${inline(c)}</td>`).join('') + '</tr>').join('') +
        '</tbody></table></div>';
      continue;
    }
    if (/^[-*]\s/.test(line)) {
      const items = [];
      while (i < lines.length && /^[-*]\s/.test(lines[i])) { items.push(lines[i].replace(/^[-*]\s/, '')); i++; }
      html += '<ul>' + items.map(it => `<li>${inline(it)}</li>`).join('') + '</ul>';
      continue;
    }
    if (/^\d+\.\s/.test(line)) {
      const items = [];
      while (i < lines.length && /^\d+\.\s/.test(lines[i])) { items.push(lines[i].replace(/^\d+\.\s/, '')); i++; }
      html += '<ol>' + items.map(it => `<li>${inline(it)}</li>`).join('') + '</ol>';
      continue;
    }
    const para = [];
    while (i < lines.length && !/^\s*$/.test(lines[i]) && !/^[-*]\s/.test(lines[i]) && !/^\d+\.\s/.test(lines[i]) && !/^\|.*\|\s*$/.test(lines[i]) && !/^#{1,3}\s/.test(lines[i])) {
      para.push(lines[i]); i++;
    }
    html += `<p>${inline(para.join(' '))}</p>`;
  }
  return html;
}

// ---------- chat ----------
let sampleFn = null;
let downloadsFn = null;
let history = [];
const transcript = document.getElementById('transcript');
const composer = document.getElementById('composer');
const questionEl = document.getElementById('question');
const sendBtn = document.getElementById('sendBtn');

function addMsg(role, html, isPlaceholder) {
  const wrap = document.createElement('div');
  wrap.className = 'msg ' + role;
  const bubble = document.createElement('div');
  bubble.className = 'bubble' + (isPlaceholder ? ' thinking' : '');
  bubble.innerHTML = html;
  wrap.appendChild(bubble);
  transcript.appendChild(wrap);
  transcript.scrollTop = transcript.scrollHeight;
  return bubble;
}

// ---------- CSV export ----------
function toCSV(rows) {
  if (!rows || !rows.length) return '';
  const cols = Object.keys(rows[0]);
  const esc = (v) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [cols.join(',')];
  for (const r of rows) lines.push(cols.map(c => esc(r[c])).join(','));
  return lines.join('\n');
}

async function exportCSV(toolName, rows, btn) {
  if (!downloadsFn) downloadsFn = await claude.use('downloads').catch(() => null);
  if (!downloadsFn) { btn.textContent = 'Export unavailable in this view'; btn.disabled = true; return; }
  const original = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Saving…';
  const filename = `tier_bench_${toolName}_${new Date().toISOString().slice(0, 10)}.csv`;
  try {
    const res = await downloadsFn.save({ filename, data: toCSV(rows) });
    btn.textContent = res.status === 'saved' ? 'Saved ✓' : 'Sent ✓';
  } catch (e) {
    btn.textContent = (e && e.code === 'declined') ? original : 'Export failed — try again';
  } finally {
    setTimeout(() => { btn.textContent = original; btn.disabled = false; }, 2200);
  }
}

async function askQuestion(q) {
  if (!q || !q.trim()) return;
  questionEl.value = '';
  sendBtn.disabled = true;
  addMsg('user', escapeHtml(q));
  const thinking = addMsg('assistant', 'Thinking…', true);

  if (!sampleFn) {
    sampleFn = await claude.use('sample').catch(() => null);
  }
  if (!sampleFn) {
    thinking.className = 'bubble';
    thinking.innerHTML = '<p class="err-note">Claude isn\'t available in this view, so I can\'t answer live — open this page on claude.ai to ask questions, or read the workbook/register directly.</p>';
    sendBtn.disabled = false;
    return;
  }

  history.push({ role: 'user', content: q });
  const messages = [
    { role: 'user', content: SYSTEM_PROMPT },
    { role: 'assistant', content: 'Understood — ready to query the dataset.' },
    ...history,
  ];

  // wrap each tool so we can offer its result as a CSV export after answering,
  // without changing what the model sees
  const turnResults = [];
  const wrappedTools = TOOLS.map(t => ({
    ...t,
    execute(input) {
      const r = t.execute(input);
      turnResults.push({ tool: t.name, result: r });
      return r;
    },
  }));

  try {
    const result = await sampleFn(messages, {
      tools: wrappedTools,
      modelTier: 'default',
      onText: ({ text }) => {
        thinking.className = 'bubble';
        thinking.innerHTML = renderMarkdown(text);
        transcript.scrollTop = transcript.scrollHeight;
      },
    });
    thinking.className = 'bubble';
    thinking.innerHTML = renderMarkdown(result.text);
    history.push({ role: 'assistant', content: result.text });

    // one export button per distinct tool that returned rows this turn
    // (last call wins if a tool was called more than once)
    const exportable = new Map();
    for (const tr of turnResults) {
      if (tr.result && Array.isArray(tr.result.rows) && tr.result.rows.length) {
        exportable.set(tr.tool, tr.result.rows);
      }
    }
    if (exportable.size) {
      const bar = document.createElement('div');
      bar.className = 'export-bar';
      for (const [toolName, rows] of exportable) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'export-btn';
        btn.textContent = `Export ${toolName} (${rows.length} row${rows.length === 1 ? '' : 's'}) as CSV`;
        btn.addEventListener('click', () => exportCSV(toolName, rows, btn));
        bar.appendChild(btn);
      }
      thinking.parentElement.appendChild(bar);
    }
  } catch (e) {
    console.error('Tier Bench askQuestion failed:', e);
    thinking.className = 'bubble';
    let msg = 'Something went wrong answering that.';
    if (e && e.code === 'not_granted') msg = 'This page needs permission to use Claude — allow it when prompted, then try again.';
    else if (e && e.code === 'rate_limited') msg = 'Rate limited — wait a moment and try again.';
    else if (e && e.text) msg = renderMarkdown(e.text) + '<p class="err-note">(Answer cut off — ' + escapeHtml(e.code || 'error') + ')</p>';
    else {
      // Surface whatever diagnostic detail is actually available instead of a
      // fully opaque message -- e.message/e.name are set for a plain JS
      // exception (e.g. a bug in a tool's execute()); e.code covers other
      // API-level failures. Logged to console too, for anyone with dev
      // tools open, since this page can't otherwise report back what broke.
      const detail = (e && (e.message || e.code || e.name)) || String(e);
      msg = `Something went wrong answering that (${detail}). Try again`;
      if (turnResults.length) msg += ` — the last tool call that ran was ${turnResults[turnResults.length - 1].tool}.`;
      else msg += '.';
    }
    thinking.innerHTML = typeof msg === 'string' && msg.startsWith('<') ? msg : `<p class="err-note">${escapeHtml(msg)}</p>`;
    history.pop();
  } finally {
    sendBtn.disabled = false;
  }
}

composer.addEventListener('submit', (e) => {
  e.preventDefault();
  askQuestion(questionEl.value);
});
questionEl.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); composer.requestSubmit(); }
});
</script>
"""

html = html.replace("__FRANCHISES_JSON__", FRANCHISES_JSON)
html = html.replace("__GEN_COMP_JSON__", GEN_COMP_JSON)
html = html.replace("__CHANNEL_MIX_JSON__", CHANNEL_MIX_JSON)
html = html.replace("__TIER_SUMMARY_JSON__", TIER_SUMMARY_JSON)
html = html.replace("__STOCK_JSON__", STOCK_JSON)
html = html.replace("__COUNTRY_JSON__", COUNTRY_JSON)
html = html.replace("__REGIONAL_TIERING_JSON__", REGIONAL_TIERING_JSON)
html = html.replace("__REGISTER_URL__", "https://claude.ai/code/artifact/1b4cce0a-d4aa-4f87-9c8f-10926b1d28c4")

out_path = str(Path(__file__).resolve().parent / "tier_bench.html")
with open(out_path, "w") as f:
    f.write(html)
print("written", out_path, len(html), "bytes")
