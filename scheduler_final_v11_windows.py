import streamlit as st
import pandas as pd
import random
import os
from collections import defaultdict, OrderedDict
from PIL import Image
import base64
from io import BytesIO
from pathlib import Path
import re


# ---------------------------------------------------------------------
# Week-grid rendering (merged in from grid_render.py so this stays a
# single-file deploy — no separate import to go missing on GitHub)
# ---------------------------------------------------------------------

DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']

ISO_ROWS = [f'Zone {c}' for c in 'ABCDEFGH']
QS_ROWS = [f'Zone {i}' for i in range(1, 14) if i != 5]
FLOATER_LETTER_ROWS = [f'Floater {c}' for c in 'ABCDEFGH']
FLOATER_EXTRA_ROW = 'Floater (extra)'
FLOATER_GENERAL_ROW = 'General Floater'

CATEGORY_ROWS = {
    'ISO / TECAN MAINT': ISO_ROWS,
    'QS AUTOMATED EXT': QS_ROWS,
    'HORIZON': ['EXT/NORM/DIL', 'POC Swap'],
    'POC': ['DNEasy/Mix-1'],
    'PGD': ['PGD'],
    'TIH': ['TIH'],
    'TIU': ['TIU'],
    'FLOATERS': FLOATER_LETTER_ROWS + [FLOATER_EXTRA_ROW, FLOATER_GENERAL_ROW],
}
CATEGORY_ORDER = ['ISO / TECAN MAINT', 'QS AUTOMATED EXT', 'HORIZON', 'POC',
                   'PGD', 'TIH', 'TIU', 'FLOATERS', 'TRAINING']

CATEGORY_COLORS = {
    'ISO / TECAN MAINT': '#34d399',
    'QS AUTOMATED EXT':  '#f87171',
    'HORIZON':           '#fb923c',
    'POC':                '#a78bfa',
    'PGD':                '#c084fc',
    'TIH':                '#60a5fa',
    'TIU':                '#38bdf8',
    'FLOATERS':           '#fbbf24',
    'TRAINING':           '#4ade80',
}


def classify_role(token):
    """Map a raw assignment token to (category, subrow_label, optional_tag)."""
    if token.startswith('Tecan Maintenance') or token.startswith('ISO Zone'):
        z = token.split('Zone ')[-1].strip()
        return ('ISO / TECAN MAINT', f'Zone {z}', None)
    if token.startswith('QS Zone'):
        z = token.replace('QS ', '').strip()
        return ('QS AUTOMATED EXT', z, None)
    if token == 'HZN EXT/NORM/DIL':
        return ('HORIZON', 'EXT/NORM/DIL', None)
    if token.startswith('HZN POC Swap'):
        tag = 'H\u2192P' if 'First Half HZN' in token else 'P\u2192H'
        return ('HORIZON', 'POC Swap', tag)
    if token.startswith('DNEasy'):
        return ('POC', 'DNEasy/Mix-1', None)
    if token == 'PGD':
        return ('PGD', 'PGD', None)
    if token.startswith('TIH'):
        tag = 'CLA' if token.endswith('CLA') else None
        return ('TIH', 'TIH', tag)
    if token.startswith('TIU'):
        return ('TIU', 'TIU', None)
    if token.startswith('Floater') or token == 'General Floater':
        m = re.search(r'Floater ([A-H])$', token)
        if m:
            return ('FLOATERS', f'Floater {m.group(1)}', None)
        if re.search(r'Floater \d+$', token):
            return ('FLOATERS', FLOATER_EXTRA_ROW, None)
        return ('FLOATERS', FLOATER_GENERAL_ROW, None)
    if token.endswith('Training'):
        return ('TRAINING', token, None)
    return None


def build_grid(role_long):
    """role_long: list of (day, name, raw_role_token). Returns nested dict
    category -> subrow -> day -> [display strings]."""
    grid = {}
    for day, name, token in role_long:
        cls = classify_role(token)
        if cls is None:
            continue
        cat, subrow, tag = cls
        disp = name if not tag else f"{name} <span class='tag'>{tag}</span>"
        grid.setdefault(cat, OrderedDict()).setdefault(subrow, {}).setdefault(day, []).append(disp)
    return grid


CSS = """
.sched-outer { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
  background: #0e1117; color: #e6e6e6; border-radius: 10px; overflow: hidden;
  border: 1px solid #2a3140; display: inline-block; max-width: 100%; }
.sched-scroll { max-height: 74vh; overflow: auto; }
.sched-table { border-collapse: separate; border-spacing: 0; font-size: 13px;
  table-layout: fixed; width: 1650px; max-width: 100%; }
.sched-table col.role-col { width: 180px; }
.sched-table col.day-col { width: 210px; }
.sched-table thead th { position: sticky; top: 0; z-index: 3; background: #161b22;
  color: #8b949e; font-weight: 600; font-size: 11px; letter-spacing: .04em;
  text-transform: uppercase; padding: 10px 12px; text-align: center;
  border-bottom: 1px solid #2a3140; border-left: 1px solid #1c212b; }
.sched-table thead th:first-child { border-left: none; }
.sched-table thead th.role-head { text-align: left; left: 0; z-index: 4; background: #161b22; }
.sched-table thead .headcount { display: block; margin-top: 2px; font-size: 16px;
  font-weight: 700; color: #e6e6e6; letter-spacing: 0; text-transform: none; }
.sched-table tbody td, .sched-table tbody th { padding: 8px 12px;
  border-bottom: 1px solid #1c212b; border-left: 1px solid #1c212b;
  vertical-align: top; line-height: 1.65; }
.sched-table tbody td:first-child, .sched-table tbody th:first-child { border-left: none; }
.sched-table tbody tr:not(.cat-row):nth-child(even) td,
.sched-table tbody tr:not(.cat-row):nth-child(even) th { background: #10141b; }
.sched-table tbody tr:not(.cat-row):hover td,
.sched-table tbody tr:not(.cat-row):hover th { background: #1b2230; }
.cat-row th, .cat-row td { background: var(--accent-bg) !important; padding: 6px 12px;
  font-size: 11px; font-weight: 700; letter-spacing: .07em; color: var(--accent);
  text-transform: uppercase; border-top: 1px solid #2a3140; border-bottom: 1px solid #2a3140;
  border-left: none; }
.role-cell { position: sticky; left: 0; background: #12161f; z-index: 2;
  border-left: 4px solid var(--accent) !important; font-weight: 500; color: #c9d1d9;
  white-space: nowrap; }
.sched-table tbody tr:not(.cat-row):nth-child(even) .role-cell { background: #14181f; }
.sched-table tbody tr:not(.cat-row):hover .role-cell { background: #1c2330; }
.cell-name { color: #e6e6e6; }
.cell-empty { color: #454c59; text-align: center; }
.tag { display: inline-block; margin-left: 3px; padding: 1px 5px; border-radius: 3px;
  font-size: 10px; font-weight: 700; background: #263041; color: #9fb0c9; white-space: nowrap; }
.name-sep { color: #454c59; }
.legend { display: flex; flex-wrap: wrap; gap: 14px; padding: 10px 14px; background: #161b22;
  border-top: 1px solid #2a3140; font-size: 11px; color: #8b949e; }
.legend .dot { display: inline-block; width: 8px; height: 8px; border-radius: 2px;
  margin-right: 5px; vertical-align: middle; }
"""


def render_week_grid_html(grid, day_headcount=None):
    """Returns a <div class='sched-outer'>...</div> fragment (CSS included via <style>)."""
    day_headcount = day_headcount or {}
    parts = [f"<style>{CSS}</style>",
             "<div class='sched-outer'><div class='sched-scroll'><table class='sched-table'>"]
    parts.append("<colgroup><col class='role-col'>" + "<col class='day-col'>" * 7 + "</colgroup>")
    parts.append("<thead><tr><th class='role-head'>Role / zone</th>")
    for d in DAYS:
        hc = day_headcount.get(d, '')
        hc_html = f"<span class='headcount'>{hc}</span>" if hc != '' else ''
        parts.append(f"<th>{d}{hc_html}</th>")
    parts.append("</tr></thead><tbody>")

    for cat in CATEGORY_ORDER:
        subrows_data = grid.get(cat, {})
        if cat == 'TRAINING':
            if not subrows_data:
                continue
            subrow_labels = list(subrows_data.keys())
        else:
            subrow_labels = CATEGORY_ROWS[cat]
            if not subrows_data and cat not in grid:
                # still show the skeleton so coverage gaps are visible
                pass

        accent = CATEGORY_COLORS[cat]
        parts.append(
            f"<tr class='cat-row' style='--accent:{accent}; --accent-bg:{accent}26'>"
            f"<td colspan='8'>{cat}</td></tr>"
        )
        for subrow in subrow_labels:
            parts.append(f"<tr><th class='role-cell' style='--accent:{accent}'>{subrow}</th>")
            for d in DAYS:
                names = subrows_data.get(subrow, {}).get(d, [])
                if names:
                    joined = "<span class='name-sep'>, </span>".join(
                        f"<span class='cell-name'>{n}</span>" for n in names
                    )
                    parts.append(f"<td>{joined}</td>")
                else:
                    parts.append("<td class='cell-empty'>\u2014</td>")
            parts.append("</tr>")

    parts.append("</tbody></table></div>")
    parts.append(
        "<div class='legend'>" +
        "".join(
            f"<span><span class='dot' style='background:{c}'></span>{cat}</span>"
            for cat, c in CATEGORY_COLORS.items() if cat != 'TRAINING'
        ) +
        "</div></div>"
    )
    return "".join(parts)

# ---------------------------------------------------------------------

st.set_page_config(layout='wide', page_title="3rd Shift Schedule Generator")

BASE_DIR = Path(__file__).resolve().parent

def get_base64_logo(img_path):
    if not os.path.exists(img_path):
        return ""
    img = Image.open(img_path)
    buffered = BytesIO()
    img.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()

logo_base64 = get_base64_logo(BASE_DIR / "natera.png")
st.markdown(f"""
    <style>
        body {{
            background-color: #0e1117;
            color: white;
            font-family: Helvetica, sans-serif;
        }}
        .logo-container {{
            position: absolute;
            top: 10px;
            right: 20px;
            z-index: 1000;
        }}
        .stDownloadButton > button, .stButton > button {{
            background-color: white !important;
            color: black !important;
        }}
        .stDataFrame {{ background-color: white !important; color: black !important; }}
    </style>
    <div class="logo-container">
        <img src="data:image/png;base64,{logo_base64}" width="150">
    </div>
""", unsafe_allow_html=True)

st.title("3rd Shift Schedule Generator")

STD_COLS_MAP = {
    'name': 'Name', 'employee name': 'Name',
    'shift': 'Shift',
    'iso': 'ISO', 'tiu': 'TIU', 'qs': 'QS', 'float': 'FLOAT',
    'cls': 'CLS', 'pgd': 'PGD', 'hzn': 'HZN', 'tih': 'TIH',
    'poc': 'POC', 'cla': 'CLA', 'cls trainee': 'CLS_TRAINEE',
}
REQUIRED_COLS = ['Name','Shift','ISO','TIU','QS','FLOAT','CLS','PGD','HZN','TIH','POC','CLA','CLS_TRAINEE']

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = {}
    for c in df.columns:
        key = c.strip()
        lower = key.lower()
        new_cols[c] = STD_COLS_MAP.get(lower, key)
    df = df.rename(columns=new_cols)
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = ''
    return df

def canon_basic(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = (s.replace('â€"','-').replace('â€"','-').replace('–','-').replace('—','-')
           .replace('_','-').replace('/','-'))
    s = re.sub(r'\bto\b|\bthru\b|\bthrough\b', '-', s, flags=re.IGNORECASE)
    s = re.sub(r'[^A-Za-z\- ]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

DATA_FILE = BASE_DIR / "today_active_workers_corrected.csv"

def load_data():
    if not os.path.exists(DATA_FILE):
        st.error(f"{DATA_FILE} not found.")
        return pd.DataFrame()
    try:
        df = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(DATA_FILE, encoding="cp1252")
    df.insert(0, '__roster_index', range(len(df)))
    df.columns = [c.strip() for c in df.columns]
    df = normalize_columns(df)
    df['Name'] = df['Name'].astype(str).str.strip()
    role_cols = [c for c in REQUIRED_COLS if c not in ['Name','Shift']]
    for c in role_cols:
        df[c] = (
            df[c].astype(str).str.strip().str.lower()
            .replace({'nan':'', 'none':'', 'no':'', 'false':'', '0':''})
        )
    df['Shift'] = df['Shift'].astype(str).map(canon_basic)
    df = df[df['Name'].str.len() > 0]
    df = df[df['Name'].str.lower() != 'nan']
    df = df.sort_values('__roster_index').drop_duplicates(subset=['Name'], keep='first').reset_index(drop=True)
    return df

df = load_data()
if df.empty:
    st.stop()

days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
DAY_ORDER = days[:]

if 'pto_by_day' not in st.session_state:
    st.session_state.pto_by_day = {}

st.subheader("Select Employee PTO/Sick")
name_options = df['Name'].tolist()
if 'selected_name' not in st.session_state or st.session_state.selected_name not in name_options:
    st.session_state.selected_name = name_options[0]
selected_name = st.selectbox(
    "Select employee PTO/sick",
    name_options,
    index=name_options.index(st.session_state.selected_name)
)
st.session_state.selected_name = selected_name

st.markdown(f"**PTO Days for {selected_name}**")
prev_days = st.session_state.pto_by_day.get(selected_name, [])
pto_days_state = {}
for d in days:
    pto_days_state[d] = st.toggle(d, value=d in prev_days, key=f"{selected_name}_{d}")
st.session_state.pto_by_day[selected_name] = [d for d, v in pto_days_state.items() if v]

st.markdown("### Current PTO Selections")
for emp, days_off in st.session_state.pto_by_day.items():
    if days_off:
        if st.button(f"{emp} (edit PTO)", key=f"goto_pto_{emp}"):
            st.session_state.selected_name = emp
        st.markdown(", ".join(days_off))

if 'ot_by_day' not in st.session_state:
    st.session_state.ot_by_day = {}

st.subheader("Select Employee Overtime")
if 'ot_selected_name' not in st.session_state or st.session_state.ot_selected_name not in name_options:
    st.session_state.ot_selected_name = name_options[0]
ot_selected_name = st.selectbox(
    "Select employee OVERTIME",
    name_options,
    index=name_options.index(st.session_state.ot_selected_name),
    key="ot_selectbox"
)
st.session_state.ot_selected_name = ot_selected_name

st.markdown(f"**Overtime Days for {ot_selected_name}**")
prev_ot_days = st.session_state.ot_by_day.get(ot_selected_name, [])
ot_days_state = {}
for d in days:
    ot_days_state[d] = st.toggle(d, value=d in prev_ot_days, key=f"OT_{ot_selected_name}_{d}")
st.session_state.ot_by_day[ot_selected_name] = [d for d, v in ot_days_state.items() if v]

st.markdown("### Current Overtime Selections")
for emp, days_ot in st.session_state.ot_by_day.items():
    if days_ot:
        if st.button(f"{emp} (edit OT)", key=f"goto_ot_{emp}"):
            st.session_state.ot_selected_name = emp
        st.markdown(", ".join(days_ot))

# --- Training Pairs UI ---
TRAINING_WORKFLOWS = ['QS', 'ISO', 'Floating', 'HZN', 'PGD', 'TIH', 'POC', 'DNEasy', 'TIU']
ONE_ON_ONE_WORKFLOWS = ['QS', 'ISO']

if 'training_pairs' not in st.session_state:
    st.session_state.training_pairs = []

st.subheader("Training Assignments")

with st.expander("Add Training Pair"):
    all_names = df['Name'].tolist()
    tr_trainee = st.selectbox("Trainee", all_names, key="tr_trainee")
    tr_trainer  = st.selectbox("Trainer",  all_names, key="tr_trainer")
    tr_workflow = st.selectbox("Workflow being trained", TRAINING_WORKFLOWS, key="tr_workflow")

    st.markdown("**Training Days**")
    tr_days_state = {}
    for d in ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']:
        tr_days_state[d] = st.toggle(d, value=False, key=f"TR_{tr_trainee}_{tr_trainer}_{d}")
    tr_days = [d for d, v in tr_days_state.items() if v]

    if st.button("Add Training Pair"):
        if tr_trainee == tr_trainer:
            st.error("Trainee and trainer must be different people.")
        elif not tr_days:
            st.error("Select at least one training day.")
        else:
            st.session_state.training_pairs.append({
                'trainee': tr_trainee,
                'trainer': tr_trainer,
                'workflow': tr_workflow,
                'days': tr_days,
            })
            st.success(f"Added: {tr_trainer} training {tr_trainee} on {tr_workflow} — {', '.join(tr_days)}")

if st.session_state.training_pairs:
    st.markdown("### Current Training Pairs")
    for i, pair in enumerate(st.session_state.training_pairs):
        col1, col2 = st.columns([4,1])
        with col1:
            st.markdown(f"**{pair['trainer']} → {pair['trainee']}** | {pair['workflow']} | {', '.join(pair['days'])}")
        with col2:
            if st.button("Remove", key=f"remove_pair_{i}"):
                st.session_state.training_pairs.pop(i)
                st.rerun()

DAYS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
IDX = {d:i for i, d in enumerate(DAYS)}

TOKEN_MAP = {
    'sunday':'Sun','sun':'Sun','su':'Sun','s':'Sun',
    'monday':'Mon','mon':'Mon','m':'Mon',
    'tuesday':'Tue','tues':'Tue','tue':'Tue','tu':'Tue','t':'Tue',
    'wednesday':'Wed','weds':'Wed','wed':'Wed','w':'Wed',
    'thursday':'Thu','thurs':'Thu','thur':'Thu','thu':'Thu','th':'Thu',
    'friday':'Fri','fri':'Fri','f':'Fri',
    'saturday':'Sat','sat':'Sat','sa':'Sat',
}
ALL_TOKENS = sorted(TOKEN_MAP.keys(), key=len, reverse=True)
DAY_ALT = r'(?:' + '|'.join(re.escape(t) for t in ALL_TOKENS) + r')'

def _token_to_day(tok: str):
    return TOKEN_MAP.get(tok.lower())

def _days_range(a: str, b: str):
    ai, bi = IDX[a], IDX[b]
    if ai <= bi:
        return DAYS[ai:bi+1]
    return DAYS[ai:] + DAYS[:bi+1]

def _apply_S_heuristic(raw: str) -> str:
    s = raw
    s = re.sub(r'(?i)\b(t|tu|tue)\s*[-_/]\s*s\b', r'\1-Sat', s)
    s = re.sub(r'(?i)\b(f|fri|friday)\s*[-_/]\s*s\b', r'\1-Sat', s)
    return s

def expand_shift_days(shift_str: str):
    raw = canon_basic(shift_str)
    raw = _apply_S_heuristic(raw)

    m = re.search(rf'\b({DAY_ALT})\b\s*-\s*\b({DAY_ALT})\b', raw, flags=re.IGNORECASE)
    if m:
        a = _token_to_day(m.group(1))
        b = _token_to_day(m.group(2))
        if a and b:
            return _days_range(a, b)

    toks = re.findall(rf'\b{DAY_ALT}\b', raw, flags=re.IGNORECASE)
    canon_toks = [_token_to_day(t) for t in toks if _token_to_day(t)]
    if len(canon_toks) >= 2:
        a, b = canon_toks[0], canon_toks[1]
        return _days_range(a, b)

    if len(canon_toks) == 1:
        a = canon_toks[0]
        start = IDX[a]
        return [DAYS[(start + k) % 7] for k in range(5)]

    return DAYS[:]

def is_on_shift(row, day):
    return day in expand_shift_days(row['Shift'])

def is_not_pto(name, day):
    return day not in st.session_state.pto_by_day.get(name, [])

def is_overtime(name, day):
    return day in st.session_state.ot_by_day.get(name, [])

def is_training_day(name, day):
    for pair in st.session_state.get('training_pairs', []):
        if name in (pair['trainee'], pair['trainer']) and day in pair['days']:
            return True
    return False

def working_pool(day) -> pd.DataFrame:
    mask = df.apply(
        lambda r: (is_on_shift(r, day) or is_overtime(r['Name'], day) or is_training_day(r['Name'], day))
                  and is_not_pto(r['Name'], day),
        axis=1
    )
    pool = df[mask].copy()
    if len(pool) == 0:
        pool = df[df['Name'].apply(lambda n: is_not_pto(n, day))].copy()
    return pool

CORE_ROLES = ['ISO','TIU','QS','FLOAT','CLS','PGD','HZN','TIH','POC','CLA','CLS_TRAINEE']
HZN_DNA_ROLE = 'HZN EXT/NORM/DIL'

# 7 ISO zones (A-G), 13 QS zones (1-13)
ISO_ZONE_LIST = [f'Zone {c}' for c in ['A','B','C','D','E','F','G','H']]
QS_ZONE_LIST = [f'Zone {i}' for i in range(1, 14)]

ISO_ZONE_ORDER = {z:i+1 for i, z in enumerate(ISO_ZONE_LIST)}
QS_ZONE_ORDER = {z:i+1 for i, z in enumerate(QS_ZONE_LIST)}

QS_STEAL_ORDER = QS_ZONE_LIST[:]

GENERAL_FLOATER_TARGET = 7
QS_BASE_FLOATER_COUNT_PER_DAY = 1

def count_yes_roles_list(row):
    return [r for r in CORE_ROLES if str(row.get(r, '')).strip().lower() == 'yes']

def skills_count(row):
    return len(count_yes_roles_list(row))

def pick(df_in, cond):
    if df_in.empty:
        return df_in.copy()
    def safe_cond(r):
        try:
            return bool(cond(r))
        except Exception:
            return False
    mask = df_in.apply(safe_cond, axis=1)
    return df_in[mask.reindex(df_in.index, fill_value=False)]

def priority_names(df_pool, already_assigned, reserve_cls=False, limit=None, prefer_more_skills=False, prefer_no_float=False):
    if df_pool is None or df_pool.empty:
        return []
    pool = df_pool[~df_pool['Name'].isin(already_assigned)].copy()
    if pool.empty:
        return []

    pool = pool.assign(
        _yes_count=pool.apply(skills_count, axis=1),
        _has_float=pool['FLOAT'].astype(str).str.strip().str.lower().eq('yes'),
        _rand=pool['Name'].apply(lambda _: random.random())
    )

    def take(frame, n=None):
        # People with no Float fallback go first when prefer_no_float is set —
        # they have nowhere else to land if they miss this pool. Float-qualified
        # people have a safe landing spot (Floater) either way, so they yield.
        sort_cols, sort_asc = [], []
        if prefer_no_float:
            sort_cols.append('_has_float'); sort_asc.append(True)
        if prefer_more_skills:
            sort_cols += ['_yes_count', '_rand']; sort_asc += [False, True]
        else:
            sort_cols += ['_rand']; sort_asc += [True]
        frame = frame.sort_values(sort_cols, ascending=sort_asc)
        names = frame['Name'].tolist()
        return names if n is None else names[:n]

    if reserve_cls:
        non_cls = pool[pool['CLS'].str.strip().str.lower() != 'yes']
        cls_only = pool[pool['CLS'].str.strip().str.lower() == 'yes']
        first = take(non_cls, None)
        if limit is None or len(first) >= limit:
            return first if limit is None else first[:limit]
        need = limit - len(first)
        second = take(cls_only, need)
        return first + second

    return take(pool, limit)

def priority_names_excluding(df_pool, already_assigned, exclude_set=None, reserve_cls=False, limit=None, prefer_more_skills=False, prefer_no_float=False):
    exclude_set = exclude_set or set()
    if df_pool is None or df_pool.empty:
        return []
    base = df_pool[~df_pool['Name'].isin(exclude_set)]
    primary = priority_names(base, already_assigned, reserve_cls, limit, prefer_more_skills, prefer_no_float)
    if limit is None or len(primary) >= limit:
        return primary
    need = limit - len(primary)
    topup = df_pool[df_pool['Name'].isin(exclude_set)]
    return primary + priority_names(topup, already_assigned.union(set(primary)), reserve_cls, need, prefer_more_skills, prefer_no_float)

def block_rank(workflow: str) -> float:
    # Display order: ISO > QS > PGD > TIH > HZN > Floater > POC
    w = str(workflow)
    if w.startswith('Tecan Maintenance'):
        return 1
    if w.startswith('ISO '):
        return 1
    if w.startswith('QS Zone '):
        return 2
    if w.startswith('QS Floater'):
        return 2.5
    if w.startswith('PGD'):
        return 3
    if w.startswith('TIH'):
        return 4
    if w.startswith('TIU'):
        return 4.5
    if w.startswith('HZN EXT/NORM/DIL'):
        return 5
    if w.startswith('Floater'):
        return 6
    if w.startswith('HZN POC Swap'):
        return 7
    if w.startswith('DNEasy/Mix-1') or w.startswith('DNEasy'):
        return 7.5
    if w.startswith('POC'):
        return 8
    return 10

def zone_rank(workflow: str) -> int:
    w = str(workflow)
    if w.startswith('ISO Zone '):
        z = w.replace('ISO ', '')
        return ISO_ZONE_ORDER.get(z, 999)
    if w.startswith('QS Zone '):
        z = w.replace('QS ', '')
        return QS_ZONE_ORDER.get(z, 999)
    if w.startswith('Tecan Maintenance/Rack Disposal Zone '):
        letter = w[-1]
        return ord(letter) - ord('A')
    if w.startswith('ISO Zone') or 'Zone A' in w or 'Zone B' in w:
        # paired monday zones
        letter = w.split('Zone ')[-1][0] if 'Zone ' in w else 'Z'
        return ord(letter) - ord('A')
    if w.startswith('Floater'):
        # Floater A-G sort before Floater 1-2
        m = re.search(r'Floater ([A-G])', w)
        if m:
            return ord(m.group(1)) - ord('A')
        m = re.search(r'Floater (\d+)', w)
        return 100 + int(m.group(1)) if m else 200
    if w.startswith('QS Floater'):
        return 1000
    return 500

def _df_row_by_name(name: str):
    return df.loc[df['Name'] == name].iloc[0]

def safe_assign(assign_map, assigned, day, name, role):
    """Assign a role to a person only if they haven't been assigned yet.
    Returns True if assignment was made, False if person was already assigned."""
    if name in assigned:
        return False
    assign_map[(day, name)].append(role)
    assigned.add(name)
    return True

def _qs_zone_assignments_for_day(assign_map, day):
    pairs = []
    for (dkey, name), roles in assign_map.items():
        if dkey != day:
            continue
        for r in roles:
            if r.startswith('QS Zone '):
                z = r.replace('QS ', '').strip()
                if z in QS_ZONE_LIST:
                    pairs.append((z, name))
    rank = {z:i for i, z in enumerate(QS_STEAL_ORDER)}
    pairs.sort(key=lambda p: rank.get(p[0], 999))
    return pairs

def _steal_from_qs(assign_map, day, want_role, predicate):
    # Steal from the HIGHEST numbered QS zone to preserve low zones (1, 2, 3...)
    pairs = _qs_zone_assignments_for_day(assign_map, day)
    for zone, name in reversed(pairs):
        row = _df_row_by_name(name)
        if predicate(row):
            roles = assign_map[(day, name)]
            roles = [r for r in roles if r != f'QS {zone}']
            assign_map[(day, name)] = roles
            assign_map[(day, name)].append(want_role)
            return True
    return False

def _steal_from_floaters(assign_map, day, want_role, predicate):
    """Steal from a floater assignment to fill a critical role."""
    for (dkey, name), roles in list(assign_map.items()):
        if dkey != day:
            continue
        if any(r.startswith('Floater') for r in roles):
            row = _df_row_by_name(name)
            if predicate(row):
                assign_map[(day, name)] = [r for r in roles if not r.startswith('Floater')]
                assign_map[(day, name)].append(want_role)
                return True
    return False

def enforce_qs_minimum(assign_map, day, pool, assigned):
    """
    Guarantee QS zones are filled starting from Zone 1 with no gaps.
    Finds the next unfilled zone in sequence and fills it.
    Steals from backups/floaters if no unassigned QS-qualified person is available.
    """
    # Find which zone numbers are already filled
    filled_qs_zones = set()
    for (dkey, _name), roles in assign_map.items():
        if dkey != day:
            continue
        for r in roles:
            if r.startswith('QS Zone '):
                filled_qs_zones.add(r.replace('QS ', '').strip())

    # Find the next zone that needs filling (first gap from Zone 1 upward)
    next_zone = None
    for z in QS_ZONE_LIST:
        if z not in filled_qs_zones:
            next_zone = z
            break

    # All zones filled or none needed
    if next_zone is None:
        return

    # Only enforce if Zone 1 is missing — that's the hard requirement
    if next_zone != 'Zone 1':
        return

    zone_label = f'QS {next_zone}'

    # Try unassigned QS-qualified workers first
    qs_unassigned = pick(pool, lambda r: r['QS'].strip().lower() == 'yes' and r['Name'] not in assigned)
    names = priority_names(qs_unassigned, assigned, reserve_cls=True, limit=1)
    if names:
        n = names[0]
        assign_map[(day, n)].append(zone_label)
        assigned.add(n)
        return

    # Steal from a floater who has QS qualification
    stolen = _steal_from_floaters(
        assign_map, day, zone_label,
        predicate=lambda row: row['QS'].strip().lower() == 'yes'
    )
    if stolen:
        return

    # Last resort: steal from a backup slot
    for (dkey, name), roles in list(assign_map.items()):
        if dkey != day:
            continue
        if any(r.endswith('Backup') or r == 'General Support' for r in roles):
            row = _df_row_by_name(name)
            if row['QS'].strip().lower() == 'yes':
                assign_map[(day, name)] = [r for r in roles if not (r.endswith('Backup') or r == 'General Support')]
                assign_map[(day, name)].append(zone_label)
                return

def is_cls_or_trainee(row):
    cls = str(row.get('CLS','')).strip().lower() == 'yes'
    trainee = str(row.get('CLS_TRAINEE','')).strip().lower() == 'yes'
    return cls or trainee

def reserve_hzn_ext(day, pool, assigned, weekly_hzn_ext_used, weekly_poc_used, limit=3, swap_floor=4):
    """Reserve HZN EXT/NORM/DIL people for the day, driven purely by roster qualifications.

    Freshness beats everything: nobody repeats HZN EXT until every eligible person
    on shift has done it this week. Within equal freshness, prefer in order:
      tier 1: HZN-qualified, not POC-qualified
      tier 2: HZN+POC qualified who already used their weekly POC slot
              (they can't take the POC swap again, so EXT costs the swap nothing)
      tier 3: HZN+POC qualified and still swap-eligible — only tapped while at
              least `swap_floor` swap-eligible people remain for HZN POC Swap
    On Sun/Mon there is no POC work, so all HZN-qualified people are fair game.
    """
    def _hzn(r): return str(r.get('HZN', '')).strip().lower() == 'yes'
    def _poc(r): return str(r.get('POC', '')).strip().lower() == 'yes'

    tier1 = pick(pool, lambda r: _hzn(r) and not _poc(r) and r['Name'] not in assigned)
    if day in ('Sun', 'Mon'):
        tier2 = pick(pool, lambda r: _hzn(r) and _poc(r) and r['Name'] not in assigned)
        tier3 = pool.iloc[0:0]
        tier3_budget = 0
    else:
        tier2 = pick(pool, lambda r: _hzn(r) and _poc(r)
                     and r['Name'] in weekly_poc_used and r['Name'] not in assigned)
        tier3 = pick(pool, lambda r: _hzn(r) and _poc(r)
                     and r['Name'] not in weekly_poc_used and r['Name'] not in assigned)
        tier3_budget = max(0, len(tier3) - swap_floor)

    picked = []
    for fresh in (True, False):
        for tier_idx, tier_pool in enumerate((tier1, tier2, tier3)):
            if len(picked) >= limit:
                return picked
            if tier_pool is None or tier_pool.empty:
                continue
            in_weekly = tier_pool['Name'].isin(weekly_hzn_ext_used)
            seg = tier_pool[~in_weekly] if fresh else tier_pool[in_weekly]
            if seg.empty:
                continue
            need = limit - len(picked)
            if tier_idx == 2:
                need = min(need, tier3_budget)
                if need <= 0:
                    continue
            got = priority_names(seg, assigned.union(picked), reserve_cls=True, limit=need)
            if tier_idx == 2:
                tier3_budget -= len(got)
            picked.extend(got)
    return picked

def enforce_tih_minimum(assign_map, day, pool, assigned, tih_reserved_names=None):
    # Commit the people already reserved upfront for TIH (see the
    # tih_reserved block before ISO/Float/QS). They were locked into
    # `assigned` at reservation time to keep ISO/Float/QS/PGD off them, so
    # this must be a raw append (matching the PGD/DNE commit pattern) —
    # safe_assign would refuse to act on a name already in `assigned` and
    # silently drop the role, leaving them reserved but never actually
    # placed on TIH. Everything below is a shortage fallback.
    for n in (tih_reserved_names or []):
        assign_map[(day, n)].append('TIH_CLS')

    # Sun-Wed: need at least 2 CLS or CLS_TRAINEE on TIH
    # Other days: need at least 1 CLS on TIH
    if day in ['Sun','Mon','Tue','Wed']:
        min_tih = 2
    else:
        min_tih = 1

    current_tih = sum(
        1 for (d, _n), roles in assign_map.items()
        if d == day and any(r.startswith('TIH') for r in roles)
    )
    needed = max(0, min_tih - current_tih)
    if needed == 0:
        return

    # Sun-Wed: CLS or CLS_TRAINEE qualifies; other days: CLS only
    if day in ['Sun','Mon','Tue','Wed']:
        tih_pool = pick(pool, lambda r: is_cls_or_trainee(r)
                        and str(r.get('TIH','')).strip().lower() == 'yes'
                        and r['Name'] not in assigned)
    else:
        tih_pool = pick(pool, lambda r: str(r.get('CLS','')).strip().lower() == 'yes'
                        and str(r.get('TIH','')).strip().lower() == 'yes'
                        and r['Name'] not in assigned)

    names = priority_names(tih_pool, assigned, reserve_cls=False, limit=needed)
    for n in names:
        assign_map[(day, n)].append('TIH_CLS')
        assigned.add(n)
        needed -= 1

    for _ in range(needed):
        _steal_from_qs(
            assign_map, day, 'TIH_CLS',
            predicate=lambda row: row.get('TIH','') == 'yes' and is_cls_or_trainee(row)
        )

def enforce_sun_mon_mins(assign_map, day, pool, assigned):
    have_tih = any(
        d == day and any(r.startswith('TIH') for r in roles)
        for (d, _n), roles in assign_map.items()
    )
    if not have_tih:
        cls_pool = pick(pool, lambda r: r['CLS'].strip().lower() == 'yes' and r['Name'] not in assigned)
        tih_cls_pool = pick(cls_pool, lambda r: r['TIH'].strip().lower() == 'yes')
        names = priority_names(tih_cls_pool, assigned, reserve_cls=False, limit=1)
        if names:
            n = names[0]
            assign_map[(day, n)].append('TIH_CLS')
            assigned.add(n)
        else:
            _steal_from_qs(
                assign_map,
                day,
                'TIH_CLS',
                predicate=lambda row: row['TIH'] == 'yes' and row['CLS'] == 'yes'
            )

    have_hzn = any(
        d == day and any(r.startswith('HZN EXT/NORM/DIL') for r in roles)
        for (d, _n), roles in assign_map.items()
    )
    if not have_hzn:
        hzn_pool = pick(pool, lambda r: r['HZN'].strip().lower() == 'yes' and r['Name'] not in assigned)
        names = priority_names(hzn_pool, assigned, reserve_cls=True, limit=1)
        if names:
            n = names[0]
            assign_map[(day, n)].append('HZN EXT/NORM/DIL')
            assigned.add(n)
        else:
            _steal_from_qs(
                assign_map,
                day,
                'HZN EXT/NORM/DIL',
                predicate=lambda row: row['HZN'] == 'yes'
            )

def count_roles_for_day(assign_map, day, prefix):
    total = 0
    for (dkey, _name), roles in assign_map.items():
        if dkey != day:
            continue
        total += sum(1 for r in roles if str(r).startswith(prefix))
    return total

def backup_label_for_row(row, day=None):
    # Everyone gets a real role — route extras to QS or Floater
    if str(row.get('QS', '')).strip().lower() == 'yes':
        return 'General Floater'
    if str(row.get('FLOAT', '')).strip().lower() == 'yes':
        return 'Floater'
    return 'General Floater'

def final_fill_no_unassigned(day, pool, assigned, assign_map):
    # Everyone gets a real assignment — no backups ever
    working_names = set(pool['Name'])
    leftovers = sorted(
        list(working_names - assigned),
        key=lambda n: int(df.loc[df['Name'] == n, '__roster_index'].iloc[0])
    )

    # Count current QS zones filled so we can continue numbering
    qs_zone_list_active = [z for z in QS_ZONE_LIST if z != 'Zone 5']
    filled_qs = set()
    for (dkey, _n), roles in assign_map.items():
        if dkey == day:
            for r in roles:
                if r.startswith('QS Zone '):
                    filled_qs.add(r.replace('QS ', '').strip())
    next_qs_idx = len(filled_qs)

    current_floaters = count_roles_for_day(assign_map, day, 'Floater')
    zone_letters = ['A','B','C','D','E','F','G','H']

    for name in leftovers:
        row = _df_row_by_name(name)
        qs_yes = str(row.get('QS', '')).strip().lower() == 'yes'
        float_yes = str(row.get('FLOAT', '')).strip().lower() == 'yes'

        if qs_yes and next_qs_idx < len(qs_zone_list_active):
            # Put them on the next available QS zone
            assign_map[(day, name)].append(f'QS {qs_zone_list_active[next_qs_idx]}')
            next_qs_idx += 1
        elif float_yes:
            current_floaters += 1
            if day in ['Sun', 'Mon']:
                assign_map[(day, name)].append(f'General Floater')
            else:
                letter_idx = current_floaters - 1
                label = f'Floater {zone_letters[letter_idx]}' if letter_idx < len(zone_letters) else f'Floater {letter_idx - len(zone_letters) + 1}'
                assign_map[(day, name)].append(label)
        elif qs_yes:
            # QS qualified but all zones full — still put them on QS support
            assign_map[(day, name)].append('General Floater')
        else:
            assign_map[(day, name)].append('General Floater')
        assigned.add(name)

if st.button("Generate Weekly Schedule"):
    rows = []
    prev_tiu = set()
    prev_dne = set()

    roster = df[['Name','__roster_index','Shift']].copy().sort_values('__roster_index')

    prev_pgd = set()
    prev_hzn_ext = set()
    prev_hzn_poc = set()
    prev_iso = set()
    prev_float = set()
    prev_qs = set()
    weekly_poc_used = set()  # tracks everyone who did POC this week
    weekly_iso_count = {}    # ISO count per person this week
    weekly_pgd_used  = set() # tracks everyone who did PGD this week
    weekly_hzn_ext_used = set()  # tracks everyone who did HZN EXT this week
    role_long = []       # (day, display_name, raw_role_token) — one row per assignment
    day_headcount = {}   # day -> number of people working that day
    for day in days:
        pool = working_pool(day)
        working_names = set(pool['Name'])
        day_headcount[day] = len(working_names)

        assigned = set()
        assign_map = defaultdict(list)

        # --- 0. Pre-reserve critical roles before ISO/Float/QS consumes everyone ---
        # Order matters: PGD first (smallest pool), then HZN EXT, then HZN POC Swap.
        # Each step excludes the previous to guarantee no double-assignments.
        # --- Training pairs: assign first, lock both people in ---
        # Trainer and trainee are both added to assigned so nothing else grabs them
        training_display = {}  # name -> display override "Trainer:Trainee"
        for pair in st.session_state.get('training_pairs', []):
            if day not in pair['days']:
                continue
            trainee  = pair['trainee']
            trainer  = pair['trainer']
            workflow = pair['workflow']
            # both must be in today's pool
            if trainee not in working_names or trainer not in working_names:
                continue
            if trainer in assigned or trainee in assigned:
                continue
            # assign both to the training workflow
            role_label = f'{workflow} Training'
            assign_map[(day, trainer)].append(role_label)
            assign_map[(day, trainee)].append(role_label)
            assigned.add(trainer)
            assigned.add(trainee)
            # track for display: only trainer row shows "Trainer:Trainee", trainee row hidden
            training_display[trainer] = f'{trainer}:{trainee}'
            training_display[trainee] = '__HIDE__'
            # count toward weekly caps if applicable
            if workflow == 'ISO' and day != 'Sun':
                weekly_iso_count[trainer] = weekly_iso_count.get(trainer, 0) + 1
                weekly_iso_count[trainee] = weekly_iso_count.get(trainee, 0) + 1
            if workflow in ('POC', 'DNEasy'):
                weekly_poc_used.add(trainer)
                weekly_poc_used.add(trainee)

        pgd_reserved          = set()
        pgd_reserved_names    = []
        dne_reserved          = set()
        dne_reserved_names    = []
        hzn_ext_reserved      = set()
        hzn_ext_reserved_names = []
        hzn_poc_reserved      = set()
        swap_reserved         = []

        # --- TIH minimum, reserved FIRST for every day (2 on Sun-Wed, 1 on
        # Thu-Sat) — same upfront-reservation pattern PGD/HZN use below. If
        # this ran late instead (the old behavior), ISO/Float/QS could
        # already have consumed the whole crew — a real problem on
        # Monday's ~20-person shift — forcing TIH to steal an
        # already-assigned QS person as its only option.
        tih_min_today = 2 if day in ['Sun', 'Mon', 'Tue', 'Wed'] else 1
        if day in ['Sun', 'Mon', 'Tue', 'Wed']:
            tih_reserve_pool = pick(pool, lambda r: is_cls_or_trainee(r)
                                     and str(r.get('TIH', '')).strip().lower() == 'yes')
        else:
            tih_reserve_pool = pick(pool, lambda r: str(r.get('CLS', '')).strip().lower() == 'yes'
                                     and str(r.get('TIH', '')).strip().lower() == 'yes')
        tih_reserved_names = priority_names(tih_reserve_pool, assigned, reserve_cls=False, limit=tih_min_today)
        tih_reserved = set(tih_reserved_names)
        assigned.update(tih_reserved)  # lock in immediately so nothing downstream can steal them

        if day not in ['Sun', 'Mon']:
            # 1. PGD (1 person) — picked FIRST, added to assigned immediately so nothing steals them
            pgd_pool_all = pool[(pool['PGD'] == 'yes') & (~pool['Name'].isin(assigned))].copy()
            pgd_candidates = [n for n in pgd_pool_all['Name'].tolist() if n not in weekly_pgd_used]
            if not pgd_candidates:
                pgd_candidates = [n for n in pgd_pool_all['Name'].tolist() if n not in prev_pgd]
            if not pgd_candidates:
                pgd_candidates = pgd_pool_all['Name'].tolist()
            random.shuffle(pgd_candidates)
            pgd_pick = pgd_candidates[0] if pgd_candidates else None
            if pgd_pick:
                pgd_reserved_names = [pgd_pick]
                pgd_reserved = {pgd_pick}
                assigned.add(pgd_pick)

            # 1b. DNEasy/Mix-1 (1 CLS+POC person) — reserved before HZN POC Swap eats them all
            dne_reserve_pool = pick(
                pool,
                lambda r: r['CLS'].strip().lower() == 'yes'
                          and r['POC'].strip().lower() == 'yes'
                          and r['Name'] not in pgd_reserved
                          and r['Name'] not in weekly_poc_used
            )
            dne_reserved_names = priority_names_excluding(
                dne_reserve_pool, assigned, exclude_set=prev_dne, reserve_cls=False, limit=1
            )
            dne_reserved = set(dne_reserved_names)
            assigned.update(dne_reserved)  # lock in immediately

            # 2. HZN EXT/NORM/DIL (2-3 people) — non-POC preferred but not required,
            # weekly rotation so nobody repeats before everyone eligible has gone
            hzn_ext_reserved_names = reserve_hzn_ext(
                day, pool, assigned, weekly_hzn_ext_used, weekly_poc_used, limit=3
            )
            hzn_ext_reserved = set(hzn_ext_reserved_names)

            # 3. HZN POC Swap (4 people, exclude PGD and HZN EXT people)
            swap_reserve_pool = pick(
                pool,
                lambda r: r['POC'].strip().lower() == 'yes'
                          and r['HZN'].strip().lower() == 'yes'
                          and r['Name'] not in pgd_reserved
                          and r['Name'] not in hzn_ext_reserved
                          and r['Name'] not in dne_reserved
                          and r['Name'] not in weekly_poc_used
            )
            swap_reserved = priority_names_excluding(
                swap_reserve_pool, assigned, exclude_set=prev_hzn_poc, reserve_cls=True, limit=4
            )
            hzn_poc_reserved = set(swap_reserved)

        else:
            # Sun/Mon: reserve 2-3 HZN EXT people. There is no POC work on these days,
            # so every HZN-qualified person is eligible — no POC filter.
            hzn_ext_reserved_names = reserve_hzn_ext(
                day, pool, assigned, weekly_hzn_ext_used, weekly_poc_used, limit=3
            )
            hzn_ext_reserved = set(hzn_ext_reserved_names)

        # --- 1. ISO + Floater paired assignment, then QS ---
        # reserved set: nobody in here can be touched by ISO/Float/QS
        reserved = hzn_poc_reserved | hzn_ext_reserved | pgd_reserved | dne_reserved | tih_reserved

        ISO_MIN = 4
        QS_MIN  = 4

        # Single-pass: assign ISO first, then Float from remaining, no double simulation
        def iso_under_cap(r):
            n = r['Name']
            count = weekly_iso_count.get(n, 0)
            quals = sum(1 for c in CORE_ROLES if str(r.get(c, '')).strip().lower() == 'yes')
            return count < (2 if quals <= 2 else 1)

        if day == 'Sun':
            # Sunday: 7 people on Tecan Maintenance/Rack Disposal — ISO+FLOAT qualified
            # Does NOT count toward weekly ISO cap
            sun_pool = pick(pool, lambda r: r['ISO'].strip().lower() == 'yes'
                            and r['FLOAT'].strip().lower() == 'yes'
                            and r['Name'] not in reserved)
            sun_names = priority_names_excluding(sun_pool, assigned, exclude_set=prev_iso, reserve_cls=True, limit=len(ISO_ZONE_LIST), prefer_more_skills=True)
            for i, name in enumerate(sun_names):
                safe_assign(assign_map, assigned, day, name, f'Tecan Maintenance/Rack Disposal {ISO_ZONE_LIST[i]}')
            iso_all = []
            float_all = []

        elif day == 'Mon':
            # Monday spec (explicit): 4 people cover all 8 ISO zones in pairs —
            # one person on A/B, one on C/D, one on E/F, one on G/H. No
            # separate Monday floaters; the pairing itself is the coverage.
            ISO_PAIRS = [('Zone A', 'Zone B'), ('Zone C', 'Zone D'),
                         ('Zone E', 'Zone F'), ('Zone G', 'Zone H')]
            mon_iso_pool = pick(pool, lambda r: r['ISO'].strip().lower() == 'yes' and r['Name'] not in reserved)
            mon_iso_names = priority_names_excluding(
                mon_iso_pool, assigned, exclude_set=prev_iso, reserve_cls=True,
                limit=len(ISO_PAIRS), prefer_more_skills=True
            )
            for (z1, z2), name in zip(ISO_PAIRS, mon_iso_names):
                safe_assign(assign_map, assigned, day, name, f'ISO {z1}/{z2}')
            iso_all, float_all = [], []

        else:
            # Tue-Sat: normal ISO + Float assignment with weekly cap
            iso_pool_all = pick(pool, lambda r: r['ISO'].strip().lower() == 'yes' and r['Name'] not in reserved)
            iso_pool_capped = pick(iso_pool_all, iso_under_cap)
            iso_pool_filtered = iso_pool_capped if len(iso_pool_capped) >= ISO_MIN else iso_pool_all
            iso_all = priority_names_excluding(iso_pool_filtered, assigned, exclude_set=prev_iso, reserve_cls=True, limit=len(ISO_ZONE_LIST), prefer_more_skills=True)

            float_pool_filtered = pick(pool, lambda r: r['FLOAT'].strip().lower() == 'yes'
                                       and r['Name'] not in reserved
                                       and r['Name'] not in set(iso_all))
            float_all = priority_names(float_pool_filtered, assigned | set(iso_all), reserve_cls=True, limit=len(ISO_ZONE_LIST))

        max_pairs = min(len(iso_all), len(float_all), len(ISO_ZONE_LIST))
        n_pairs   = min(ISO_MIN, max_pairs)

        # Expand pairs beyond floor only if QS stays above minimum
        for candidate_pairs in range(n_pairs + 1, max_pairs + 1):
            consumed = set(iso_all[:candidate_pairs]) | set(float_all[:candidate_pairs])
            remaining_qs = pick(
                pool,
                lambda r: r['QS'].strip().lower() == 'yes'
                          and r['Name'] not in consumed
                          and r['Name'] not in reserved
            ).shape[0]
            if remaining_qs >= QS_MIN:
                n_pairs = candidate_pairs
            else:
                break

        # Commit ISO/Float — Sun and Mon are both handled directly above
        # (Tecan Maint / paired zones), with no separate Floater step.
        if day not in ['Sun', 'Mon']:
            for i, name in enumerate(iso_all[:n_pairs]):
                safe_assign(assign_map, assigned, day, name, f'ISO {ISO_ZONE_LIST[i]}')

            # Floaters labeled to match ISO zone letter (A-G), then 1-2 for extras
            zone_letters = ['A','B','C','D','E','F','G','H']
            for i, name in enumerate(float_all[:n_pairs]):
                if i < len(zone_letters):
                    label = f'Floater {zone_letters[i]}'
                else:
                    label = f'Floater {i - len(zone_letters) + 1}'
                safe_assign(assign_map, assigned, day, name, label)

        # --- 2. QS (Zones 1-13, skip Zone 5) ---
        qs_zone_list_filtered = [z for z in QS_ZONE_LIST if z != 'Zone 5']
        qs_pool  = pick(pool, lambda r: r['QS'].strip().lower() == 'yes'
                        and r['Name'] not in assigned and r['Name'] not in reserved)
        # Monday spec: 4 zones stay deliberately empty (8 filled of 12).
        qs_limit_today = 8 if day == 'Mon' else len(qs_zone_list_filtered)
        qs_names = priority_names(qs_pool, assigned, reserve_cls=True, limit=qs_limit_today)
        for i, name in enumerate(qs_names):
            safe_assign(assign_map, assigned, day, name, f'QS {qs_zone_list_filtered[i]}')

        # QS Floaters removed — no longer assigned

        # --- 3. TIU ---
        # Must exclude `reserved`, not just `assigned` — HZN EXT/HZN POC swap
        # people are held in `reserved` but aren't locked into `assigned`
        # until their own commit step runs later in this loop. Without this,
        # TIU could poach one of them first and silently leave Horizon a
        # person short every day this collision hit.
        cls_pool = pick(pool, lambda r: r['CLS'].strip().lower() == 'yes'
                        and r['Name'] not in assigned and r['Name'] not in reserved)
        tiu_pool = pick(cls_pool, lambda r: r['TIU'].strip().lower() == 'yes')
        if day == 'Mon':
            # Monday: 1 CLS on TIU/Stickers
            for name in priority_names_excluding(tiu_pool, assigned, exclude_set=prev_tiu, reserve_cls=False, limit=1):
                safe_assign(assign_map, assigned, day, name, 'TIU/Stickers')
        elif day == 'Sun':
            tiu_count = 1
            for name in priority_names_excluding(tiu_pool, assigned, exclude_set=prev_tiu, reserve_cls=False, limit=tiu_count):
                safe_assign(assign_map, assigned, day, name, 'TIU')
        else:
            tiu_count = 2
            for name in priority_names_excluding(tiu_pool, assigned, exclude_set=prev_tiu, reserve_cls=False, limit=tiu_count):
                safe_assign(assign_map, assigned, day, name, 'TIU')

        # --- 4. HZN EXT/NORM/DIL (2-3 people, pre-reserved) ---
        for name in hzn_ext_reserved_names:
            safe_assign(assign_map, assigned, day, name, 'HZN EXT/NORM/DIL')

        # --- 4b. HZN POC Swap (Tue-Sat) — assign the pre-reserved people now ---
        if day not in ['Sun', 'Mon']:
            swap_selected = list(hzn_poc_reserved)
            half = min(2, (len(swap_selected) + 1) // 2)
            start_hzn = swap_selected[:half]
            start_poc = swap_selected[half:]
            for name in start_hzn:
                safe_assign(assign_map, assigned, day, name, 'HZN POC Swap (First Half HZN / Second Half POC)')
            for name in start_poc:
                safe_assign(assign_map, assigned, day, name, 'HZN POC Swap (First Half POC / Second Half HZN)')

        # Wednesday: ensure at least 2 CLS or CLS_TRAINEE on POC
        if day == 'Wed':
            current_poc = sum(
                1 for (d, _n), roles in assign_map.items()
                if d == day and any('HZN POC Swap' in r for r in roles)
                and is_cls_or_trainee(_df_row_by_name(_n))
            )
            needed_poc = max(0, 2 - current_poc)
            if needed_poc > 0:
                poc_extra_pool = pick(
                    pool,
                    lambda r: is_cls_or_trainee(r)
                              and str(r.get('POC','')).strip().lower() == 'yes'
                              and str(r.get('HZN','')).strip().lower() == 'yes'
                              and r['Name'] not in assigned
                              and r['Name'] not in weekly_poc_used
                )
                for name in priority_names(poc_extra_pool, assigned, reserve_cls=False, limit=needed_poc):
                    safe_assign(assign_map, assigned, day, name, 'HZN POC Swap (First Half POC / Second Half HZN)')

        # --- 5. DNEasy/Mix-1 (Tue-Sat, pre-reserved CLS+POC person) ---
        if day not in ['Sun','Mon']:
            for name in dne_reserved_names:
                assign_map[(day, name)].append('DNEasy/Mix-1')

        # --- 6. TIH enforcement ---
        enforce_tih_minimum(assign_map, day, pool, assigned, tih_reserved_names)

        tih_cla_pool = pick(
            pool,
            lambda r: r['TIH'].strip().lower() == 'yes' and r['CLA'].strip().lower() == 'yes' and r['Name'] not in assigned
        )
        for name in priority_names(tih_cla_pool, assigned, reserve_cls=True, limit=1):
            safe_assign(assign_map, assigned, day, name, 'TIH_CLA')

        # --- 7. PGD (Tue-Sat, 1 person, different person each day) ---
        if day not in ['Sun', 'Mon']:
            for name in pgd_reserved_names:
                assign_map[(day, name)].append('PGD')


        # --- 9. Sun/Mon minimums ---
        if day in ['Sun','Mon']:
            enforce_sun_mon_mins(assign_map, day, pool, assigned)

        # --- 10. QS Zone 1 enforcement (last safety net) ---
        enforce_qs_minimum(assign_map, day, pool, assigned)

        # --- 11. Fill any remaining unassigned workers ---
        final_fill_no_unassigned(day, pool, assigned, assign_map)

        # Capture raw (unjoined) role tokens for the week-grid view, before
        # they get string-joined into the CSV-facing 'Workflow' column below.
        for (dkey, name), roles in assign_map.items():
            if dkey != day:
                continue
            disp = training_display.get(name, name)
            if disp == '__HIDE__':
                continue
            for token in roles:
                if token.endswith('Training') and ':' in str(disp):
                    trainer_name, trainee_name = disp.split(':', 1)
                    role_long.append((day, f'{trainer_name} \u2192 {trainee_name}', token))
                elif token.endswith('Training'):
                    continue
                elif token.startswith('ISO ') and '/Zone ' in token:
                    # Monday's paired zones ("ISO Zone A/Zone B") — one person
                    # covering two zones. Show them in both zone rows.
                    z1, z2 = token[len('ISO '):].split('/')
                    role_long.append((day, name, f'ISO {z1}'))
                    role_long.append((day, name, f'ISO {z2}'))
                else:
                    role_long.append((day, name, token))

        today_rows = []
        for _, person in roster.iterrows():
            name = person['Name']
            if name not in working_names:
                continue

            roles = assign_map.get((day, name), []).copy()
            if not roles:
                row = _df_row_by_name(name)
                roles = [backup_label_for_row(row, day)]

            wf = " / ".join(roles) if len(roles) == 2 else ", ".join(roles)
            # Use Trainer:Trainee display if this person is in a training pair today
            display_name = training_display.get(name, name)
            if display_name == '__HIDE__':
                continue
            today_rows.append((day, int(person['__roster_index']), display_name, wf))

        if len(today_rows) != len(working_names):
            already = {n for _d, _idx, n, _w in today_rows}
            for n in (working_names - already):
                row = _df_row_by_name(n)
                today_rows.append((
                    day,
                    int(df.loc[df['Name'] == n, '__roster_index'].iloc[0]),
                    n,
                    backup_label_for_row(row, day)
                ))

        # Update weekly POC tracker
        weekly_poc_used.update(
            n for (dkey, n), roles in assign_map.items()
            if dkey == day and any('HZN POC Swap' in r or r == 'DNEasy/Mix-1' or r == 'DNEasy' for r in roles)
        )

        for (dkey, n), roles in assign_map.items():
            if dkey == day and any(r.startswith('ISO Zone') or r.startswith('ISO ') for r in roles):
                if day != 'Sun':  # Sunday Tecan Maintenance doesn't count toward cap
                    weekly_iso_count[n] = weekly_iso_count.get(n, 0) + 1

        prev_tiu = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'TIU' for r in roles)]
        )
        prev_dne = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'DNEasy/Mix-1' or r == 'DNEasy' for r in roles)]
        )
        prev_pgd = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'PGD' for r in roles)]
        )
        weekly_pgd_used.update(prev_pgd)
        prev_hzn_ext = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'HZN EXT/NORM/DIL' for r in roles)]
        )
        weekly_hzn_ext_used.update(prev_hzn_ext)
        prev_hzn_poc = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any('HZN POC Swap' in r for r in roles)]
        )
        prev_iso = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r.startswith('ISO ') for r in roles)]
        )
        prev_float = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r.startswith('Floater') for r in roles)]
        )
        prev_qs = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r.startswith('QS Zone') for r in roles)]
        )

        rows.extend(today_rows)

    out = pd.DataFrame(rows, columns=['Day','__roster_index','Name','Workflow'])
    st.session_state['full_schedule'] = out
    st.session_state['role_long'] = role_long
    st.session_state['day_headcount'] = day_headcount

st.subheader("Generated Weekly Schedule")
if 'full_schedule' in st.session_state:
    grid = build_grid(st.session_state.get('role_long', []))
    day_headcount = st.session_state.get('day_headcount', {})
    st.markdown(render_week_grid_html(grid, day_headcount), unsafe_allow_html=True)

    st.download_button(
        "Download Schedule CSV",
        st.session_state['full_schedule'].to_csv(index=False).encode(),
        "weekly_schedule.csv",
        "text/csv"
    )

    with st.expander("View as a flat per-day table (for spot-checking)"):
        df_out = st.session_state['full_schedule']
        selected_day = st.radio("Day:", days, horizontal=True, key="flat_day_radio")
        view = df_out[df_out['Day'] == selected_day].copy()
        view['__block'] = view['Workflow'].map(block_rank)
        view['__zone'] = view['Workflow'].map(zone_rank)
        view.sort_values(['__block','__zone','__roster_index'], inplace=True, kind='stable')
        view = view.drop(columns=['__block','__zone']).reset_index(drop=True)
        st.caption(f"Showing {len(view)} working staff for {selected_day}")
        st.dataframe(view[['Name','Workflow']], use_container_width=True)
else:
    st.info("Click 'Generate Weekly Schedule' to begin.")
