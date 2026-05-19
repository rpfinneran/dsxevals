import streamlit as st
import pandas as pd
from datetime import datetime
import re
import zipfile
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="U12 Roster Builder", layout="wide", page_icon="⚽")

# ── Inline CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #0f1923; }
  [data-testid="stHeader"] { background: transparent; }
  [data-testid="stSidebar"] {
    background: #14222e;
    border-right: 2px solid #1e3448;
  }
  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3,
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] p { color: #e2e8f0 !important; }
  h1, h2, h3 { color: #f0f4f8 !important; }
  .pitch-zone {
    background: linear-gradient(180deg,#1a5c2a 0%,#1e6b31 50%,#1a5c2a 100%);
    border: 3px dashed #4ade80;
    border-radius: 12px;
    min-height: 220px;
    padding: 16px;
  }
  .pitch-zone-label {
    color: #86efac;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 10px;
    text-align: center;
  }
  .player-card {
    background: linear-gradient(135deg,#1e3448,#243b55);
    border: 1.5px solid #2d4f6b;
    border-radius: 10px;
    padding: 10px 14px;
    width: 175px;
    cursor: grab;
    transition: transform .15s, box-shadow .15s;
    user-select: none;
  }
  .player-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 24px rgba(0,0,0,.5);
    border-color: #4ade80;
  }
  .player-card.on-pitch {
    background: linear-gradient(135deg,#14532d,#166534);
    border-color: #4ade80;
  }
  .player-name {
    color: #f0f4f8;
    font-weight: 700;
    font-size: 14px;
    margin-bottom: 5px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .player-dob { color: #94a3b8; font-size: 11px; margin-bottom: 7px; }
  .tags { display: flex; flex-wrap: wrap; gap: 4px; }
  .tag {
    border-radius: 9999px;
    padding: 2px 8px;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .5px;
    white-space: nowrap;
  }
  .tag-age-5  { background:#4c1d95; color:#ede9fe; }
  .tag-age-6  { background:#6d28d9; color:#ede9fe; }
  .tag-age-7  { background:#7c3aed; color:#ede9fe; }
  .tag-age-8  { background:#2563eb; color:#dbeafe; }
  .tag-age-9  { background:#0891b2; color:#cffafe; }
  .tag-age-10 { background:#059669; color:#d1fae5; }
  .tag-age-11 { background:#d97706; color:#fef3c7; }
  .tag-age-12 { background:#dc2626; color:#fee2e2; }
  .tag-age-13 { background:#9333ea; color:#f3e8ff; }
  .tag-age-14 { background:#0f766e; color:#ccfbf1; }
  .tag-age-15 { background:#be185d; color:#fce7f3; }
  .tag-eval { background:#1e3a5f; color:#93c5fd; border: 1px solid #3b82f6; }
  .tag-dsx  { background:#7f1d1d; color:#fca5a5; border: 1px solid #ef4444; }
  .count-badge {
    display: inline-block;
    background: #4ade80;
    color: #052e16;
    border-radius: 9999px;
    padding: 2px 10px;
    font-size: 13px;
    font-weight: 800;
    margin-left: 8px;
  }
  .stButton button {
    background: #7f1d1d !important;
    color: #fca5a5 !important;
    border: 1px solid #ef4444 !important;
    border-radius: 6px !important;
    font-size: 11px !important;
    padding: 2px 8px !important;
    margin-top: 4px;
  }
  .stButton button:hover { background: #991b1b !important; }
</style>
""", unsafe_allow_html=True)


# ── Pure-Python xlsx reader (no openpyxl needed) ─────────────────────────────
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

        # Resolve to a path that actually exists in the zip.
        # Targets can be:
        #   "worksheets/sheet1.xml"     -> prepend "xl/"
        #   "xl/worksheets/sheet1.xml"  -> use as-is
        #   "/xl/worksheets/sheet1.xml" -> strip leading slash (openpyxl quirk)
        if sheet_target in names:
            sheet_path = sheet_target
        elif f"xl/{sheet_target}" in names:
            sheet_path = f"xl/{sheet_target}"
        else:
            # Last resort: scan zip for the first worksheet entry
            sheet_path = next(
                (n for n in names if "worksheets/sheet" in n and n.endswith(".xml")), None
            )
            if sheet_path is None:
                raise FileNotFoundError("Could not locate worksheet in xlsx archive.")

        ws = ET.fromstring(zf.read(sheet_path))

    EXCEL_EPOCH = datetime(1899, 12, 30)

    def parse_cell(c):
        t    = c.get("t", "n")
        v_el = c.find(f"{{{NS}}}v")
        v    = v_el.text if v_el is not None else None
        if v is None:
            return None
        if t == "s":
            return shared[int(v)]
        if t == "b":
            return bool(int(v))
        if t == "inlineStr":
            el = c.find(f"{{{NS}}}is/{{{NS}}}t")
            return el.text if el is not None else None
        try:
            fv = float(v)
            # Heuristic: serial dates in xlsx land between ~1900 and ~2100
            # Col E (DOB) serials for kids born 2010-2022 fall in 40179–44927
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

    # Keep only female players with a first name
    df = df[df["gender"].apply(lambda x: str(x).strip().lower() == "f" if pd.notna(x) else False)].copy()
    df = df[df["first_name"].notna()].copy()

    # Full name
    df["full_name"] = (df["first_name"].apply(lambda x: str(x).strip()) + " " +
                       df["last_name"].apply(lambda x: str(x).strip() if pd.notna(x) else "")).str.strip()

    # DOB display string — column E is already parsed to Timestamp by _xlsx_to_rows
    df["dob"] = pd.to_datetime(df["dob"], errors="coerce")
    df["dob_str"] = df["dob"].apply(
        lambda d: f"{d.month}/{d.day}/{d.year}" if pd.notna(d) else None
    )

    # Soccer age — use the pre-calculated column F from the spreadsheet
    # It already accounts for the Aug 1 season cutoff; no need to recalculate.
    df["age"] = df["soccer_age"].apply(
        lambda x: int(float(x)) if pd.notna(x) and str(x).replace(".", "").isdigit() else None
    )

    # Eval group: simplify "U10 (August 1 2016 - July 31 2017)" → "U10"
    def simplify_eval(raw):
        if pd.isna(raw) or raw is None:
            return None
        m = re.search(r"U(\d+)", str(raw), re.IGNORECASE)
        return f"U{m.group(1)}" if m else None

    df["eval_group"] = df["eval_group_raw"].apply(simplify_eval)

    # DSX flag
    df["is_dsx"] = df["club_raw"].apply(
        lambda x: bool(re.search(r"dsx", str(x), re.IGNORECASE)) if pd.notna(x) else False
    )

    # Deduplicate
    df = df.drop_duplicates(subset=["full_name", "dob"]).reset_index(drop=True)
    df["id"] = df.index.astype(str)

    return df[["id", "full_name", "dob_str", "age", "eval_group", "is_dsx"]].to_dict("records")


# ── Session state ─────────────────────────────────────────────────────────────
if "roster" not in st.session_state:
    st.session_state.roster = []


# ── File upload (sidebar) ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚽ Roster Builder")
    st.markdown("---")
    uploaded_file = st.file_uploader("📂 Upload roster (.xlsx)", type=["xlsx"])
    st.markdown("---")

if uploaded_file is None:
    st.markdown("# ⚽ U12 Roster Builder")
    st.info("👈 Upload your roster export (.xlsx) in the sidebar to get started.")
    st.stop()

# Clear roster when a new file is loaded
if st.session_state.get("last_file") != uploaded_file.name:
    st.session_state.roster = []
    st.session_state["last_file"] = uploaded_file.name

file_bytes = uploaded_file.read()
players    = load_data(file_bytes)


# ── Helper: render one card ───────────────────────────────────────────────────
def age_tag_class(age):
    if age is None:
        return "tag-age-12"
    return f"tag-age-{min(max(int(age), 5), 15)}"

def render_card(p, on_pitch=False):
    card_class = "player-card on-pitch" if on_pitch else "player-card"
    age_cls    = age_tag_class(p["age"])
    age_label  = f"Age {p['age']}" if p["age"] is not None else "Age ?"
    eval_tag   = f'<span class="tag tag-eval">{p["eval_group"]}</span>' if p["eval_group"] else ""
    dsx_tag    = '<span class="tag tag-dsx">DSX</span>' if p["is_dsx"] else ""
    dob_str    = p["dob_str"] if p["dob_str"] else "—"
    return f"""
<div class="{card_class}">
  <div class="player-name" title="{p['full_name']}">{p['full_name']}</div>
  <div class="player-dob">DOB: {dob_str}</div>
  <div class="tags">
    <span class="tag {age_cls}">{age_label}</span>
    {eval_tag}
    {dsx_tag}
  </div>
</div>"""


# ── Sidebar filters ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Filters")

    all_eval = sorted(set(p["eval_group"] for p in players if p["eval_group"]))
    selected_eval = st.multiselect("Eval group", all_eval, default=all_eval)

    all_ages = sorted(set(p["age"] for p in players if p["age"] is not None))
    selected_ages = st.multiselect("Current age", all_ages, default=all_ages)

    dsx_only = st.checkbox("DSX players only")
    st.markdown("---")
    search = st.text_input("🔍 Search name")
    st.markdown("---")
    if st.button("🗑️ Clear roster"):
        st.session_state.roster = []
        st.rerun()


# ── Filter pool ───────────────────────────────────────────────────────────────
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


# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("# ⚽ U12 Roster Builder")
st.markdown("Use the **Add →** buttons to move players onto the U12 team, or **✕ Remove** to send them back to the pool.")
st.markdown("---")


# ── Pitch (U12 Roster) ────────────────────────────────────────────────────────
n = len(pitch_players)
st.markdown(
    f'<h3>🟢 U12 Team Roster <span class="count-badge">{n}</span></h3>',
    unsafe_allow_html=True
)

if not pitch_players:
    st.markdown(
        '<div class="pitch-zone"><div class="pitch-zone-label">U12 Roster</div>'
        '<div style="color:#4ade80;text-align:center;margin-top:30px;opacity:.5;">'
        'No players added yet — use Add → below</div></div>',
        unsafe_allow_html=True
    )
else:
    cols_per_row = 5
    rows_p = [pitch_players[i:i+cols_per_row] for i in range(0, len(pitch_players), cols_per_row)]
    st.markdown('<div class="pitch-zone"><div class="pitch-zone-label">U12 Roster</div>', unsafe_allow_html=True)
    for row in rows_p:
        cols = st.columns(cols_per_row)
        for col, p in zip(cols, row):
            with col:
                st.markdown(render_card(p, on_pitch=True), unsafe_allow_html=True)
                if st.button("✕ Remove", key=f"rem_{p['id']}"):
                    st.session_state.roster.remove(p["id"])
                    st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# ── Player pool ───────────────────────────────────────────────────────────────
st.markdown(
    f'<h3>👥 Player Pool <span class="count-badge">{len(pool_players)}</span></h3>',
    unsafe_allow_html=True
)

if not pool_players:
    st.info("No players match the current filters, or all have been added to the roster.")
else:
    cols_per_row = 5
    rows_p = [pool_players[i:i+cols_per_row] for i in range(0, len(pool_players), cols_per_row)]
    for row in rows_p:
        cols = st.columns(cols_per_row)
        for col, p in zip(cols, row):
            with col:
                st.markdown(render_card(p), unsafe_allow_html=True)
                if st.button("Add →", key=f"add_{p['id']}"):
                    st.session_state.roster.append(p["id"])
                    st.rerun()


# ── Legend ────────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("🎨 Tag Legend"):
    st.markdown("""
    | Tag | Meaning |
    |-----|---------|
    | **Age ##** (coloured pill) | Soccer age for the Fall 2026 / Spring 2027 season |
    | **U8 / U9 / U10 …** (blue outline) | Player's self-selected evaluation age group |
    | **DSX** (red outline) | Currently plays at DSX |
    """)
