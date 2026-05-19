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
  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3,
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] p { color: #e2e8f0 !important; }
  h1, h2, h3 { color: #f0f4f8 !important; }
  .stButton button {
    background: #1e3448 !important; color: #94a3b8 !important;
    border: 1px solid #2d4f6b !important; border-radius: 6px !important;
    font-size: 11px !important;
  }
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


# ── Sidebar: upload only ─────────────────────────────────────────────────────
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

# ── Inline filter bar ─────────────────────────────────────────────────────────
all_eval      = sorted(set(p["eval_group"] for p in players if p["eval_group"]))
all_ages      = sorted(set(p["age"] for p in players if p["age"] is not None))

f1, f2, f3, f4, f5 = st.columns([2.2, 2.2, 1, 2, 1])
with f1:
    selected_eval  = st.multiselect("Eval group", all_eval, default=all_eval, label_visibility="collapsed",
                                    placeholder="All eval groups")
with f2:
    selected_ages  = st.multiselect("Age", all_ages, default=all_ages, label_visibility="collapsed",
                                    placeholder="All ages")
with f3:
    dsx_only       = st.checkbox("DSX only")
with f4:
    search         = st.text_input("Search", placeholder="🔍 Search name", label_visibility="collapsed")
with f5:
    if st.button("🗑️ Clear roster"):
        st.session_state.roster = []
        st.rerun()

# Use defaults if nothing selected (treat empty multiselect as "show all")
if not selected_eval:
    selected_eval = all_eval
if not selected_ages:
    selected_ages = all_ages


# ── Filter ────────────────────────────────────────────────────────────────────
def filter_players(pool):
    out = []
    for p in pool:
        if p["eval_group"] not in selected_eval:
            continue
        if p["age"] not in selected_ages:
            continue
        if dsx_only and not p["is_dsx"]:
            continue
        if search and search.lower() not in p["full_name"].lower():
            continue
        out.append(p)
    return out

roster_ids    = set(st.session_state.roster)
pool_players  = [p for p in filter_players(players) if p["id"] not in roster_ids]
pitch_players = [p for p in players if p["id"] in roster_ids]


# ── Age tag colour map ────────────────────────────────────────────────────────
AGE_COLORS = {
    5:  ("#4c1d95","#ede9fe"), 6:  ("#6d28d9","#ede9fe"), 7:  ("#7c3aed","#ede9fe"),
    8:  ("#2563eb","#dbeafe"), 9:  ("#0891b2","#cffafe"), 10: ("#059669","#d1fae5"),
    11: ("#d97706","#fef3c7"), 12: ("#dc2626","#fee2e2"), 13: ("#9333ea","#f3e8ff"),
    14: ("#0f766e","#ccfbf1"), 15: ("#be185d","#fce7f3"),
}

def age_colors(age):
    if age is None:
        return AGE_COLORS[12]
    return AGE_COLORS.get(max(5, min(int(age), 15)), AGE_COLORS[12])


# ── Build drag-and-drop HTML component ───────────────────────────────────────
def build_dnd_html(pool, roster):
    def card_html(p, zone):
        bg, fg = age_colors(p["age"])
        age_lbl = f"Age {p['age']}" if p["age"] is not None else "Age ?"
        eval_tag = (f'<span style="background:#1e3a5f;color:#93c5fd;border:1px solid #3b82f6;'
                    f'border-radius:9999px;padding:2px 7px;font-size:10px;font-weight:700;white-space:nowrap;">'
                    f'{p["eval_group"]}</span>') if p["eval_group"] else ""
        dsx_tag  = ('<span style="background:#7f1d1d;color:#fca5a5;border:1px solid #ef4444;'
                    'border-radius:9999px;padding:2px 7px;font-size:10px;font-weight:700;white-space:nowrap;">'
                    'DSX</span>') if p["is_dsx"] else ""
        dob_str  = p["dob_str"] or "—"
        card_bg  = "linear-gradient(135deg,#14532d,#166534)" if zone == "roster" else "linear-gradient(135deg,#1e3448,#243b55)"
        border   = "#4ade80" if zone == "roster" else "#2d4f6b"
        name_escaped = p["full_name"].replace("'", "&#39;").replace('"', "&quot;")
        return f"""
<div class="player-card" draggable="true"
     data-id="{p['id']}" data-zone="{zone}"
     style="background:{card_bg};border:1.5px solid {border};border-radius:10px;
            padding:10px 14px;cursor:grab;user-select:none;margin-bottom:8px;
            transition:transform .15s,box-shadow .15s;">
  <div style="color:#f0f4f8;font-weight:700;font-size:14px;margin-bottom:4px;
              white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"
       title="{name_escaped}">{p['full_name']}</div>
  <div style="color:#94a3b8;font-size:11px;margin-bottom:6px;">DOB: {dob_str}</div>
  <div style="display:flex;flex-wrap:wrap;gap:4px;">
    <span style="background:{bg};color:{fg};border-radius:9999px;padding:2px 7px;
                 font-size:10px;font-weight:700;white-space:nowrap;">{age_lbl}</span>
    {eval_tag}{dsx_tag}
  </div>
</div>"""

    pool_cards   = "".join(card_html(p, "pool")   for p in pool)
    roster_cards = "".join(card_html(p, "roster") for p in roster)

    pool_ids   = json.dumps([p["id"] for p in pool])
    roster_ids_js = json.dumps([p["id"] for p in roster])

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: #0f1923; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }}
  .layout {{ display: flex; gap: 12px; height: 100vh; padding: 0; }}

  .column {{ display: flex; flex-direction: column; flex: 1; min-width: 0; }}

  .col-header {{
    color: #f0f4f8; font-size: 15px; font-weight: 700;
    padding: 10px 14px 8px; display: flex; align-items: center; gap: 8px;
    border-bottom: 1px solid #1e3448; flex-shrink: 0;
  }}
  .badge {{
    background: #4ade80; color: #052e16; border-radius: 9999px;
    padding: 1px 9px; font-size: 12px; font-weight: 800;
  }}

  .drop-zone {{
    flex: 1; overflow-y: auto; padding: 10px;
    border-radius: 10px; transition: background .2s;
  }}
  .pool-zone   {{ background: #14222e; border: 2px solid #1e3448; }}
  .roster-zone {{
    background: linear-gradient(180deg,#1a5c2a,#1e6b31,#1a5c2a);
    border: 2px dashed #4ade80;
  }}
  .drop-zone.drag-over {{ outline: 2px solid #4ade80; outline-offset: -2px; }}

  .player-card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(0,0,0,.5);
  }}
  .player-card.dragging {{ opacity: .4; }}

  .empty-hint {{
    color: #4ade80; opacity: .45; text-align: center;
    padding: 40px 16px; font-size: 13px;
  }}

  ::-webkit-scrollbar {{ width: 5px; }}
  ::-webkit-scrollbar-track {{ background: transparent; }}
  ::-webkit-scrollbar-thumb {{ background: #2d4f6b; border-radius: 9999px; }}
</style>
</head>
<body>
<div class="layout">

  <!-- LEFT: Pool -->
  <div class="column">
    <div class="col-header">
      👥 Player Pool
      <span class="badge" id="pool-count">{len(pool)}</span>
    </div>
    <div class="drop-zone pool-zone" id="pool-zone">
      {"".join(card_html(p,"pool") for p in pool) or '<div class="empty-hint">All players added to roster</div>'}
    </div>
  </div>

  <!-- RIGHT: Roster -->
  <div class="column">
    <div class="col-header">
      🟢 U12 Roster
      <span class="badge" id="roster-count">{len(roster)}</span>
    </div>
    <div class="drop-zone roster-zone" id="roster-zone">
      {"".join(card_html(p,"roster") for p in roster) or '<div class="empty-hint">Drag players here</div>'}
    </div>
  </div>

</div>

<script>
const poolIds   = {pool_ids};
const rosterIds = {roster_ids_js};

let dragEl = null;

document.querySelectorAll('.player-card').forEach(card => {{
  card.addEventListener('dragstart', e => {{
    dragEl = card;
    setTimeout(() => card.classList.add('dragging'), 0);
    e.dataTransfer.effectAllowed = 'move';
  }});
  card.addEventListener('dragend', () => {{
    card.classList.remove('dragging');
    dragEl = null;
  }});
}});

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
    const toZoneId = zone.id;
    const toZone   = toZoneId === 'roster-zone' ? 'roster' : 'pool';

    if (fromZone === toZone) return;   // no-op, same zone

    // Move card visually
    const emptyHint = zone.querySelector('.empty-hint');
    if (emptyHint) emptyHint.remove();
    zone.appendChild(dragEl);
    dragEl.dataset.zone = toZone;

    // Update border/bg to match destination zone
    if (toZone === 'roster') {{
      dragEl.style.background = 'linear-gradient(135deg,#14532d,#166534)';
      dragEl.style.borderColor = '#4ade80';
    }} else {{
      dragEl.style.background = 'linear-gradient(135deg,#1e3448,#243b55)';
      dragEl.style.borderColor = '#2d4f6b';
    }}

    // If source zone is now empty, show hint
    const srcZone = document.getElementById(fromZone === 'roster' ? 'roster-zone' : 'pool-zone');
    if (!srcZone.querySelector('.player-card')) {{
      const hint = document.createElement('div');
      hint.className = 'empty-hint';
      hint.textContent = fromZone === 'roster' ? 'Drag players here' : 'All players added to roster';
      srcZone.appendChild(hint);
    }}

    // Update counts
    document.getElementById('pool-count').textContent   = document.querySelectorAll('#pool-zone .player-card').length;
    document.getElementById('roster-count').textContent = document.querySelectorAll('#roster-zone .player-card').length;

    // Send updated roster list to Streamlit
    const newRoster = [...document.querySelectorAll('#roster-zone .player-card')].map(c => c.dataset.id);
    window.parent.postMessage({{type: 'streamlit:setComponentValue', value: newRoster}}, '*');
  }});
}});
</script>
</body>
</html>
"""
    return html


# ── Render ────────────────────────────────────────────────────────────────────
st.markdown("# ⚽ U12 Roster Builder")
st.caption("Drag players from the pool on the left onto the U12 Roster on the right.")

dnd_html = build_dnd_html(pool_players, pitch_players)

result = components.html(dnd_html, height=750, scrolling=False)

# components.html returns the postMessage value when the component fires
if result is not None and isinstance(result, list):
    new_roster = result
    if set(new_roster) != set(st.session_state.roster):
        st.session_state.roster = new_roster
        st.rerun()

# ── Legend ────────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("🎨 Tag Legend"):
    st.markdown("""
    | Tag | Meaning |
    |-----|---------|
    | **Age ##** (coloured pill) | Current age calculated from DOB |
    | **U8 / U9 / U10 …** (blue outline) | Player's self-selected evaluation age group |
    | **DSX** (red outline) | Currently plays at DSX |
    """)
