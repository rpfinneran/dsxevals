import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime
import re
import zipfile
import xml.etree.ElementTree as ET
import io
import json

st.set_page_config(page_title="U12 Roster Builder", layout="wide", page_icon="⚽")

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #0f1923; }
  [data-testid="stHeader"] { background: transparent; }
  [data-testid="stSidebar"] { background: #14222e; border-right: 2px solid #1e3448; }
  [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3,[data-testid="stSidebar"] label,
  [data-testid="stSidebar"] p { color: #e2e8f0 !important; }
  h1, h2, h3 { color: #f0f4f8 !important; }
</style>
""", unsafe_allow_html=True)


# ── Pure-Python xlsx reader ───────────────────────────────────────────────────
def _xlsx_to_rows(file_bytes):
    NS     = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        names = zf.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst.findall(f"{{{NS}}}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{NS}}}t")))
        wb_xml  = ET.fromstring(zf.read("xl/workbook.xml"))
        rel_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        sheet_ids = [s.attrib.get(f"{{{REL_NS}}}id")
                     for s in wb_xml.findall(f".//{{{NS}}}sheet")]
        rel_map      = {r.get("Id"): r.get("Target") for r in rel_xml}
        sheet_target = rel_map.get(sheet_ids[0], "").lstrip("/")
        if sheet_target in names:
            sheet_path = sheet_target
        elif f"xl/{sheet_target}" in names:
            sheet_path = f"xl/{sheet_target}"
        else:
            sheet_path = next(
                (n for n in names if "worksheets/sheet" in n and n.endswith(".xml")), None)
            if sheet_path is None:
                raise FileNotFoundError("Could not locate worksheet in xlsx archive.")
        ws = ET.fromstring(zf.read(sheet_path))

    EXCEL_EPOCH = datetime(1899, 12, 30)

    def parse_cell(c):
        t    = c.get("t", "n")
        v_el = c.find(f"{{{NS}}}v")
        v    = v_el.text if v_el is not None else None
        if t == "inlineStr":
            el = c.find(f"{{{NS}}}is/{{{NS}}}t")
            return el.text if el is not None else None
        if v is None:
            return None
        if t == "s":
            return shared[int(v)]
        if t == "b":
            return bool(int(v))
        try:
            fv = float(v)
            if 35000 < fv < 55000:
                return EXCEL_EPOCH + pd.Timedelta(days=fv)
            return fv
        except (ValueError, TypeError):
            return v

    rows = []
    for row_el in ws.findall(f".//{{{NS}}}row"):
        cells = {}
        for c in row_el.findall(f"{{{NS}}}c"):
            ref = c.get("r", "")
            col_letters = re.sub(r"\d", "", ref)
            col_idx = 0
            for ch in col_letters:
                col_idx = col_idx * 26 + (ord(ch.upper()) - ord("A") + 1)
            col_idx -= 1
            cells[col_idx] = parse_cell(c)
        if cells:
            rows.append([cells.get(i) for i in range(max(cells.keys()) + 1)])
    return rows


@st.cache_data
def load_data(file_bytes):
    rows = _xlsx_to_rows(file_bytes)
    if not rows:
        st.error("Could not read any data from the file.")
        st.stop()
    max_cols = max(len(r) for r in rows)
    padded   = [r + [None] * (max_cols - len(r)) for r in rows]
    df = pd.DataFrame(padded)
    if df.shape[1] < 9:
        st.error("Unexpected file format — please upload the standard roster export.")
        st.stop()
    df = df.iloc[:, :10].copy()
    df.columns = ["player_num", "gender", "first_name", "last_name", "dob",
                  "soccer_age", "eval_group_raw", "club_raw", "position", "comments"]
    df = df[df["gender"].apply(lambda x: str(x).strip().lower() == "f" if pd.notna(x) else False)].copy()
    df = df[df["first_name"].notna()].copy()
    df["full_name"] = (df["first_name"].apply(lambda x: str(x).strip()) + " " +
                       df["last_name"].apply(lambda x: str(x).strip() if pd.notna(x) else "")).str.strip()
    df["dob"] = pd.to_datetime(df["dob"], errors="coerce")
    df["dob_str"] = df["dob"].apply(
        lambda d: f"{d.month}/{d.day}/{d.year}" if pd.notna(d) else None)
    today = datetime.today().date()
    df["age"] = df["dob"].apply(
        lambda d: today.year - d.year - ((today.month, today.day) < (d.month, d.day))
        if pd.notna(d) else None)
    def simplify_eval(raw):
        if pd.isna(raw) or raw is None:
            return None
        m = re.search(r"U(\d+)", str(raw), re.IGNORECASE)
        return f"U{m.group(1)}" if m else None
    df["eval_group"] = df["eval_group_raw"].apply(simplify_eval)
    df["is_dsx"] = df["club_raw"].apply(
        lambda x: bool(re.search(r"dsx", str(x), re.IGNORECASE)) if pd.notna(x) else False)
    df = df.drop_duplicates(subset=["full_name", "dob"]).reset_index(drop=True)
    df["id"] = df.index.astype(str)
    return df[["id", "full_name", "dob_str", "age", "eval_group", "is_dsx"]].to_dict("records")


# ── Session state ─────────────────────────────────────────────────────────────
if "roster" not in st.session_state:
    st.session_state.roster = []


# ── Sidebar: upload only ──────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚽ Roster Builder")
    st.markdown("---")
    uploaded_file = st.file_uploader("📂 Upload roster (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.markdown("# ⚽ U12 Roster Builder")
    st.info("👈 Upload your roster export (.xlsx) in the sidebar to get started.")
    st.stop()

if st.session_state.get("last_file") != uploaded_file.name:
    st.session_state.roster = []
    st.session_state["last_file"] = uploaded_file.name

file_bytes = uploaded_file.read()
players    = load_data(file_bytes)

roster_ids    = set(st.session_state.roster)
pool_players  = [p for p in players if p["id"] not in roster_ids]
pitch_players = [p for p in players if p["id"] in roster_ids]

# ── Age colour map (passed into JS) ──────────────────────────────────────────
AGE_COLORS = {
    5:  ("#4c1d95","#ede9fe"), 6:  ("#6d28d9","#ede9fe"), 7:  ("#7c3aed","#ede9fe"),
    8:  ("#2563eb","#dbeafe"), 9:  ("#0891b2","#cffafe"), 10: ("#059669","#d1fae5"),
    11: ("#d97706","#fef3c7"), 12: ("#dc2626","#fee2e2"), 13: ("#9333ea","#f3e8ff"),
    14: ("#0f766e","#ccfbf1"), 15: ("#be185d","#fce7f3"),
}
def age_colors(age):
    if age is None: return AGE_COLORS[12]
    return AGE_COLORS.get(max(5, min(int(age), 15)), AGE_COLORS[12])


# ── Build full HTML component ─────────────────────────────────────────────────
def build_html(pool, roster, all_players):
    # Collect unique ages + eval groups across ALL players (not just pool)
    all_ages  = sorted(set(p["age"]        for p in all_players if p["age"] is not None))
    all_evals = sorted(set(p["eval_group"] for p in all_players if p["eval_group"]),
                       key=lambda x: int(re.search(r'\d+', x).group()))

    def card_html(p, zone):
        bg, fg   = age_colors(p["age"])
        age_lbl  = f"Age {p['age']}" if p["age"] is not None else "Age ?"
        eval_tag = (f'<span style="background:#1e3a5f;color:#93c5fd;border:1px solid #3b82f6;'
                    f'border-radius:9999px;padding:2px 7px;font-size:10px;font-weight:700;white-space:nowrap;">'
                    f'{p["eval_group"]}</span>') if p["eval_group"] else ""
        dsx_tag  = ('<span style="background:#7f1d1d;color:#fca5a5;border:1px solid #ef4444;'
                    'border-radius:9999px;padding:2px 7px;font-size:10px;font-weight:700;white-space:nowrap;">'
                    'DSX</span>') if p["is_dsx"] else ""
        card_bg  = "linear-gradient(135deg,#14532d,#166534)" if zone == "roster" else "linear-gradient(135deg,#1e3448,#243b55)"
        border   = "#4ade80" if zone == "roster" else "#2d4f6b"
        name_esc = p["full_name"].replace("'","&#39;").replace('"',"&quot;")
        eval_val = p["eval_group"] or ""
        age_val  = str(p["age"]) if p["age"] is not None else ""
        is_dsx   = "true" if p["is_dsx"] else "false"
        return (f'<div class="player-card" draggable="true" '
                f'data-id="{p["id"]}" data-zone="{zone}" '
                f'data-age="{age_val}" data-eval="{eval_val}" data-dsx="{is_dsx}" '
                f'data-name="{name_esc}" '
                f'style="background:{card_bg};border:1.5px solid {border};border-radius:10px;'
                f'padding:10px 14px;cursor:grab;user-select:none;margin-bottom:8px;'
                f'transition:transform .15s,box-shadow .15s;">'
                f'<div style="color:#f0f4f8;font-weight:700;font-size:14px;margin-bottom:4px;'
                f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="{name_esc}">{p["full_name"]}</div>'
                f'<div style="color:#94a3b8;font-size:11px;margin-bottom:6px;">DOB: {p["dob_str"] or "—"}</div>'
                f'<div style="display:flex;flex-wrap:wrap;gap:4px;">'
                f'<span style="background:{bg};color:{fg};border-radius:9999px;padding:2px 7px;'
                f'font-size:10px;font-weight:700;white-space:nowrap;">{age_lbl}</span>'
                f'{eval_tag}{dsx_tag}</div></div>')

    pool_html   = "".join(card_html(p, "pool")   for p in pool)
    roster_html = "".join(card_html(p, "roster") for p in roster)

    # Build age chip colours for JS
    age_color_map = {str(a): age_colors(a) for a in all_ages}
    age_color_js  = json.dumps({str(a): {"bg": c[0], "fg": c[1]} for a, c in age_color_map.items()})

    age_chips  = "".join(
        f'<button class="chip chip-age" data-value="{a}" '
        f'style="--chip-bg:{age_colors(a)[0]};--chip-fg:{age_colors(a)[1]};">{a}</button>'
        for a in all_ages
    )
    eval_chips = "".join(
        f'<button class="chip chip-eval" data-value="{e}">{e}</button>'
        for e in all_evals
    )

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: #0f1923;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    display: flex; flex-direction: column; height: 100vh; overflow: hidden;
  }}

  /* ── Filter bar ── */
  #filter-bar {{
    flex-shrink: 0;
    padding: 10px 12px 8px;
    background: #0d1821;
    border-bottom: 1px solid #1e3448;
    display: flex; flex-direction: column; gap: 6px;
  }}
  .filter-row {{
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  }}
  .filter-label {{
    color: #64748b; font-size: 10px; font-weight: 700;
    text-transform: uppercase; letter-spacing: 1px;
    white-space: nowrap; min-width: 28px;
  }}

  /* chips */
  .chip {{
    border: 1.5px solid transparent;
    border-radius: 9999px;
    padding: 3px 10px;
    font-size: 11px; font-weight: 700;
    cursor: pointer;
    transition: opacity .15s, transform .1s;
    white-space: nowrap;
  }}
  .chip-age {{
    background: #1e3448; color: #64748b; border-color: #2d4f6b;
  }}
  .chip-age.active {{
    background: var(--chip-bg); color: var(--chip-fg); border-color: transparent;
  }}
  .chip-eval {{
    background: #1e3448; color: #64748b; border-color: #2d4f6b;
  }}
  .chip-eval.active {{
    background: #1e3a5f; color: #93c5fd; border-color: #3b82f6;
  }}
  .chip-dsx {{
    background: #1e3448; color: #64748b; border-color: #2d4f6b;
  }}
  .chip-dsx.active {{
    background: #7f1d1d; color: #fca5a5; border-color: #ef4444;
  }}
  .chip:hover {{ opacity: .85; transform: translateY(-1px); }}

  /* search */
  #search-input {{
    background: #1e3448; border: 1.5px solid #2d4f6b; border-radius: 8px;
    color: #e2e8f0; font-size: 12px; padding: 4px 10px;
    outline: none; width: 180px;
  }}
  #search-input:focus {{ border-color: #4ade80; }}
  #search-input::placeholder {{ color: #475569; }}

  /* clear button */
  #clear-btn {{
    margin-left: auto;
    background: #1e3448; border: 1.5px solid #2d4f6b; border-radius: 8px;
    color: #94a3b8; font-size: 11px; font-weight: 600;
    padding: 4px 12px; cursor: pointer; white-space: nowrap;
    transition: background .15s;
  }}
  #clear-btn:hover {{ background: #7f1d1d; color: #fca5a5; border-color: #ef4444; }}

  /* ── Board ── */
  #board {{
    flex: 1; display: flex; gap: 10px; padding: 10px; overflow: hidden;
  }}
  .column {{ display: flex; flex-direction: column; flex: 1; min-width: 0; }}
  .col-header {{
    color: #f0f4f8; font-size: 14px; font-weight: 700;
    padding: 6px 10px 6px; display: flex; align-items: center; gap: 8px;
    border-bottom: 1px solid #1e3448; flex-shrink: 0;
  }}
  .badge {{
    background: #4ade80; color: #052e16; border-radius: 9999px;
    padding: 1px 9px; font-size: 12px; font-weight: 800;
  }}
  .drop-zone {{
    flex: 1; overflow-y: auto; padding: 8px;
    border-radius: 10px;
  }}
  .pool-zone   {{ background: #14222e; border: 2px solid #1e3448; }}
  .roster-zone {{
    background: linear-gradient(180deg,#1a5c2a,#1e6b31,#1a5c2a);
    border: 2px dashed #4ade80;
  }}
  .drop-zone.drag-over {{ outline: 2px solid #4ade80; outline-offset: -2px; }}

  .player-card:hover {{ transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,.5); }}
  .player-card.dragging {{ opacity: .35; }}
  .player-card.filtered-out {{ display: none; }}

  .empty-hint {{
    color: #4ade80; opacity: .4; text-align: center;
    padding: 40px 16px; font-size: 13px; pointer-events: none;
  }}

  ::-webkit-scrollbar {{ width: 5px; }}
  ::-webkit-scrollbar-track {{ background: transparent; }}
  ::-webkit-scrollbar-thumb {{ background: #2d4f6b; border-radius: 9999px; }}
</style>
</head>
<body>

<!-- ── Filter bar ── -->
<div id="filter-bar">
  <div class="filter-row">
    <span class="filter-label">Age</span>
    {age_chips}
    <span class="filter-label" style="margin-left:8px;">Eval</span>
    {eval_chips}
    <button class="chip chip-dsx" id="dsx-chip" data-value="dsx">DSX</button>
  </div>
  <div class="filter-row">
    <span class="filter-label">Search</span>
    <input id="search-input" type="text" placeholder="Player name…" autocomplete="off">
    <button id="clear-btn">🗑️ Clear roster</button>
  </div>
</div>

<!-- ── Board ── -->
<div id="board">
  <div class="column">
    <div class="col-header">
      👥 Player Pool <span class="badge" id="pool-count">{len(pool)}</span>
    </div>
    <div class="drop-zone pool-zone" id="pool-zone">
      {pool_html or '<div class="empty-hint" id="pool-empty">All players on roster</div>'}
    </div>
  </div>
  <div class="column">
    <div class="col-header">
      🟢 U12 Roster <span class="badge" id="roster-count">{len(roster)}</span>
    </div>
    <div class="drop-zone roster-zone" id="roster-zone">
      {roster_html or '<div class="empty-hint" id="roster-empty">Drag players here</div>'}
    </div>
  </div>
</div>

<script>
// ── Filter state ──────────────────────────────────────────────────────────────
const activeAges  = new Set({json.dumps([str(a) for a in all_ages])});
const activeEvals = new Set({json.dumps(all_evals)});
let   dsxOnly     = false;
let   searchTerm  = '';

function applyFilters() {{
  let poolVisible = 0;
  document.querySelectorAll('#pool-zone .player-card').forEach(card => {{
    const age   = card.dataset.age;
    const eval_ = card.dataset.eval;
    const dsx   = card.dataset.dsx === 'true';
    const name  = card.dataset.name.toLowerCase();

    const ageOk    = activeAges.size  === 0 || activeAges.has(age);
    const evalOk   = activeEvals.size === 0 || activeEvals.has(eval_);
    const dsxOk    = !dsxOnly || dsx;
    const searchOk = !searchTerm || name.includes(searchTerm);

    const show = ageOk && evalOk && dsxOk && searchOk;
    card.classList.toggle('filtered-out', !show);
    if (show) poolVisible++;
  }});

  // Update pool count (visible only)
  document.getElementById('pool-count').textContent = poolVisible;

  // Pool empty hint
  let hint = document.getElementById('pool-empty');
  const hasVisible = poolVisible > 0;
  if (!hasVisible && !hint) {{
    hint = document.createElement('div');
    hint.className = 'empty-hint'; hint.id = 'pool-empty';
    hint.textContent = 'No players match filters';
    document.getElementById('pool-zone').appendChild(hint);
  }} else if (hasVisible && hint) {{
    hint.remove();
  }}
}}

// ── Age chips ─────────────────────────────────────────────────────────────────
document.querySelectorAll('.chip-age').forEach(chip => {{
  chip.addEventListener('click', () => {{
    const v = chip.dataset.value;
    if (activeAges.has(v)) {{
      activeAges.delete(v);
      chip.classList.remove('active');
    }} else {{
      activeAges.add(v);
      chip.classList.add('active');
    }}
    applyFilters();
  }});
}});

// ── Eval chips ────────────────────────────────────────────────────────────────
document.querySelectorAll('.chip-eval').forEach(chip => {{
  chip.addEventListener('click', () => {{
    const v = chip.dataset.value;
    if (activeEvals.has(v)) {{
      activeEvals.delete(v);
      chip.classList.remove('active');
    }} else {{
      activeEvals.add(v);
      chip.classList.add('active');
    }}
    applyFilters();
  }});
}});

// ── DSX chip ──────────────────────────────────────────────────────────────────
document.getElementById('dsx-chip').addEventListener('click', () => {{
  dsxOnly = !dsxOnly;
  document.getElementById('dsx-chip').classList.toggle('active', dsxOnly);
  applyFilters();
}});

// ── Search ────────────────────────────────────────────────────────────────────
document.getElementById('search-input').addEventListener('input', e => {{
  searchTerm = e.target.value.toLowerCase().trim();
  applyFilters();
}});

// ── Clear roster ──────────────────────────────────────────────────────────────
document.getElementById('clear-btn').addEventListener('click', () => {{
  const rosterZone = document.getElementById('roster-zone');
  const poolZone   = document.getElementById('pool-zone');

  // Move all roster cards back to pool
  rosterZone.querySelectorAll('.player-card').forEach(card => {{
    card.style.background  = 'linear-gradient(135deg,#1e3448,#243b55)';
    card.style.borderColor = '#2d4f6b';
    card.dataset.zone = 'pool';
    poolZone.appendChild(card);
  }});

  // Remove empty hints
  const rEmpty = document.getElementById('roster-empty');
  if (rEmpty) rEmpty.remove();

  // Add roster empty hint
  const hint = document.createElement('div');
  hint.className = 'empty-hint'; hint.id = 'roster-empty';
  hint.textContent = 'Drag players here';
  rosterZone.appendChild(hint);

  updateCounts();
  applyFilters();
  window.parent.postMessage({{type:'streamlit:setComponentValue', value:[]}}, '*');
}});

// ── Drag & drop ───────────────────────────────────────────────────────────────
let dragEl = null;

function attachDrag(card) {{
  card.addEventListener('dragstart', e => {{
    dragEl = card;
    setTimeout(() => card.classList.add('dragging'), 0);
    e.dataTransfer.effectAllowed = 'move';
  }});
  card.addEventListener('dragend', () => {{
    card.classList.remove('dragging');
    dragEl = null;
  }});
}}

document.querySelectorAll('.player-card').forEach(attachDrag);

document.querySelectorAll('.drop-zone').forEach(zone => {{
  zone.addEventListener('dragover', e => {{
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    zone.classList.add('drag-over');
  }});
  zone.addEventListener('dragleave', e => {{
    if (!zone.contains(e.relatedTarget)) zone.classList.remove('drag-over');
  }});
  zone.addEventListener('drop', e => {{
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (!dragEl) return;

    const fromZone = dragEl.dataset.zone;
    const toZone   = zone.id === 'roster-zone' ? 'roster' : 'pool';
    if (fromZone === toZone) return;

    // Remove empty hints in destination
    zone.querySelectorAll('.empty-hint').forEach(h => h.remove());

    zone.appendChild(dragEl);
    dragEl.dataset.zone = toZone;

    if (toZone === 'roster') {{
      dragEl.style.background  = 'linear-gradient(135deg,#14532d,#166534)';
      dragEl.style.borderColor = '#4ade80';
    }} else {{
      dragEl.style.background  = 'linear-gradient(135deg,#1e3448,#243b55)';
      dragEl.style.borderColor = '#2d4f6b';
    }}

    // Restore empty hint in source if needed
    const srcZone = document.getElementById(fromZone === 'roster' ? 'roster-zone' : 'pool-zone');
    if (!srcZone.querySelector('.player-card')) {{
      const hint = document.createElement('div');
      hint.className = 'empty-hint';
      hint.id = fromZone === 'roster' ? 'roster-empty' : 'pool-empty';
      hint.textContent = fromZone === 'roster' ? 'Drag players here' : 'All players on roster';
      srcZone.appendChild(hint);
    }}

    updateCounts();
    applyFilters();

    const newRoster = [...document.querySelectorAll('#roster-zone .player-card')].map(c => c.dataset.id);
    window.parent.postMessage({{type:'streamlit:setComponentValue', value: newRoster}}, '*');
  }});
}});

function updateCounts() {{
  document.getElementById('roster-count').textContent =
    document.querySelectorAll('#roster-zone .player-card').length;
  // pool count updated by applyFilters
}}

// Initial filter pass (all active by default, nothing hidden)
applyFilters();
</script>
</body>
</html>"""

# ── Render ────────────────────────────────────────────────────────────────────
st.markdown("# ⚽ U12 Roster Builder")
st.caption("Drag players from the pool on the left onto the U12 Roster on the right.")

result = components.html(
    build_html(pool_players, pitch_players, players),
    height=780,
    scrolling=False
)

if result is not None and isinstance(result, list):
    if set(result) != set(st.session_state.roster):
        st.session_state.roster = result
        st.rerun()

st.markdown("---")
with st.expander("🎨 Tag Legend"):
    st.markdown("""
| Tag | Meaning |
|-----|---------|
| **Age ##** (coloured pill) | Current age calculated from DOB |
| **U8 / U9 / U10 …** (blue outline) | Player's self-selected evaluation age group |
| **DSX** (red outline) | Currently plays at DSX |
""")
