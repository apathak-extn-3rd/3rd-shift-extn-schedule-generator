"""
3rd Shift Schedule Board
------------------------
Single-file Streamlit app.

  * Sidebar: roster, PTO / sick, overtime, training pairs, generate.
  * Main:    an interactive week board (drag-and-drop, click-to-place,
             call-outs, day-of adds, coverage checks, undo, change log).
  * Edits sync back to Python, so the CSV download and the flat table always
    reflect what the board shows. The board also autosaves to disk so a page
    refresh (or another lead opening the app) picks up where you left off.

The board UI is a custom Streamlit component whose HTML/JS lives in the
BOARD_HTML string at the bottom of this file; it's written to a small folder
next to this script on startup, so this stays a one-file deploy.
"""
import base64
import html as _html
import inspect
import json
import os
import random
import re
import tempfile
import uuid
from collections import OrderedDict, defaultdict
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

# =====================================================================
# Page + look
# =====================================================================
st.set_page_config(layout="wide", page_title="3rd Shift Schedule Board", page_icon="🗓️")

BASE_DIR = Path(__file__).resolve().parent
DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']

try:  # newer Streamlit uses width='stretch'; older uses **STRETCH
    STRETCH = {'width': 'stretch'} if 'width' in inspect.signature(st.button).parameters else {'use_container_width': True}
except Exception:
    STRETCH = {}
DAY_IDX = {d: i for i, d in enumerate(DAYS)}

APP_CSS = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  html, body, [class*="css"], .stApp { font-family: "IBM Plex Sans", -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
  .stApp { background: #0f1522; color: #dce3ef; }
  .block-container { padding-top: 1.1rem; padding-bottom: 2rem; max-width: 100%; }
  section[data-testid="stSidebar"] { background: #121a29; border-right: 1px solid #24304a; }
  section[data-testid="stSidebar"] .block-container { padding-top: 1rem; }
  h1, h2, h3 { font-weight: 600; letter-spacing: -0.01em; }
  .hero { display: flex; align-items: center; gap: 18px; margin: 0 0 12px 0; }
  .hero img { height: 34px; opacity: .95; }
  .hero .t { font-size: 22px; font-weight: 600; color: #dce3ef; line-height: 1.2; }
  .hero .s { font-size: 13px; color: #98a4ba; margin-top: 2px; }
  .hero .stats { margin-left: auto; display: flex; gap: 22px; }
  .hero .stat b { display: block; font-size: 18px; font-weight: 600; color: #dce3ef; font-variant-numeric: tabular-nums; }
  .hero .stat span { font-size: 12px; color: #98a4ba; }
  .note { color: #98a4ba; font-size: 12.5px; }
  .pill { display: inline-block; font-size: 11.5px; color: #98a4ba; border: 1px solid #24304a; border-radius: 6px; padding: 1px 7px; margin-right: 4px; }
  .stButton > button, .stDownloadButton > button { border-radius: 9px; border: 1px solid #2f3d5c; background: #1b2436; color: #dce3ef; }
  .stButton > button:hover, .stDownloadButton > button:hover { border-color: #7aa2f7; color: #fff; }
  .stButton > button[kind="primary"] { background: #7aa2f7; border-color: #7aa2f7; color: #0b1020; font-weight: 600; }
  .stButton > button[kind="primary"]:hover { background: #8fb1ff; }
  div[data-testid="stExpander"] { border: 1px solid #24304a; border-radius: 10px; background: #141c2c; }
  div[data-testid="stExpander"] summary { font-weight: 500; }
  .stDataFrame { border-radius: 10px; overflow: hidden; }
  div[data-baseweb="select"] > div { background: #0f1522; border-color: #24304a; border-radius: 9px; }
  .entry { display: flex; align-items: center; gap: 8px; padding: 4px 0; border-bottom: 1px solid #1c2537; font-size: 13px; }
  .entry .nm { flex: 1; color: #dce3ef; }
  .entry .dy { color: #98a4ba; font-size: 12px; }
  footer { visibility: hidden; }
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


def get_base64_logo(img_path):
    if not os.path.exists(img_path):
        return ""
    img = Image.open(img_path)
    buffered = BytesIO()
    img.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()


# =====================================================================
# Board layout (categories / rows / colors) shared by generator + UI
# =====================================================================
ISO_ROWS = [f'Zone {c}' for c in 'ABCDEFGH']
QS_ROWS = [f'Zone {i}' for i in range(1, 14) if i != 5]
FLOATER_LETTER_ROWS = [f'Floater {c}' for c in 'ABCDEFGH']
FLOATER_NUMBERED_ROWS = [f'Floater {i}' for i in range(1, 11)]

CATEGORY_ROWS = {
    'ISO / TECAN MAINT': ISO_ROWS,
    'QS AUTOMATED EXT': QS_ROWS,
    'HORIZON': ['EXT/NORM/DIL', 'POC Swap'],
    'POC': ['DNEasy/Mix-1'],
    'PGD': ['PGD'],
    'TIH': ['TIH'],
    'TIU': ['TIU'],
    'FLOATERS': FLOATER_LETTER_ROWS + FLOATER_NUMBERED_ROWS,
}
CATEGORY_ORDER = ['ISO / TECAN MAINT', 'QS AUTOMATED EXT', 'HORIZON', 'POC',
                  'PGD', 'TIH', 'TIU', 'FLOATERS', 'TRAINING']
CATEGORY_LABELS = {
    'ISO / TECAN MAINT': 'ISO / Tecan maint',
    'QS AUTOMATED EXT': 'QS automated ext',
    'HORIZON': 'Horizon',
    'POC': 'POC',
    'PGD': 'PGD',
    'TIH': 'TIH',
    'TIU': 'TIU',
    'FLOATERS': 'Floaters',
    'TRAINING': 'Training',
}
CATEGORY_COLORS = {
    'ISO / TECAN MAINT': '#34d399',
    'QS AUTOMATED EXT':  '#f87171',
    'HORIZON':           '#fb923c',
    'POC':               '#a78bfa',
    'PGD':               '#c084fc',
    'TIH':               '#60a5fa',
    'TIU':               '#38bdf8',
    'FLOATERS':          '#fbbf24',
    'TRAINING':          '#4ade80',
}
# Roster columns a person needs marked "yes" to be considered qualified for a row.
CATEGORY_QUALS = {
    'ISO / TECAN MAINT': ['ISO'],
    'QS AUTOMATED EXT': ['QS'],
    'HORIZON': ['HZN'],
    'POC': ['CLS', 'POC'],
    'PGD': ['PGD'],
    'TIH': ['TIH'],
    'TIU': ['CLS', 'TIU'],
    'FLOATERS': ['FLOAT'],
    'TRAINING': [],
}
ROW_QUALS_OVERRIDE = {('HORIZON', 'POC Swap'): ['HZN', 'POC']}
# Numbered floater slots are overflow slots; anyone can land there without a flag.
ROW_QUALS_OVERRIDE.update({('FLOATERS', r): [] for r in FLOATER_NUMBERED_ROWS})
ROW_TAGS = {
    ('HORIZON', 'POC Swap'): ['H\u2192P', 'P\u2192H'],
    ('TIH', 'TIH'): ['', 'CLA'],
}
TRAINING_TAGS = ['Trainer', 'Trainee']


def row_id(cat, sub):
    return f'{cat}||{sub}'


def row_short(cat, sub):
    if cat == 'ISO / TECAN MAINT':
        return f'ISO {sub}'
    if cat == 'QS AUTOMATED EXT':
        return f'QS {sub}'
    if cat == 'HORIZON':
        return 'HZN EXT/NORM/DIL' if sub == 'EXT/NORM/DIL' else 'HZN POC Swap'
    return sub


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
        m2 = re.search(r'Floater (\d+)$', token)
        if m2:
            return ('FLOATERS', f'Floater {m2.group(1)}', None)
        return ('FLOATERS', 'Floater 1', None)
    if token.endswith('Training'):
        return ('TRAINING', token, None)
    return None


def board_config(df):
    """Static config for the board component: rows, colours, roster, coverage rules."""
    cats = []
    for cat in CATEGORY_ORDER:
        rows = []
        if cat == 'TRAINING':
            rows.append({'id': row_id(cat, 'Training'), 'label': 'Training', 'short': 'Training',
                         'quals': [], 'tags': TRAINING_TAGS})
        else:
            for sub in CATEGORY_ROWS[cat]:
                rows.append({
                    'id': row_id(cat, sub), 'label': sub, 'short': row_short(cat, sub),
                    'quals': ROW_QUALS_OVERRIDE.get((cat, sub), CATEGORY_QUALS[cat]),
                    'tags': ROW_TAGS.get((cat, sub), []),
                })
        cats.append({'id': cat, 'label': CATEGORY_LABELS[cat], 'color': CATEGORY_COLORS[cat],
                     'quals': CATEGORY_QUALS[cat], 'rows': rows})

    roster = []
    for _, r in df.iterrows():
        roster.append({
            'name': r['Name'],
            'idx': int(r['__roster_index']),
            'quals': count_yes_roles_list(r),
            'shift': str(r['Shift']),
            'shiftDays': expand_shift_days(r['Shift']),
        })

    all_days = {d: 1 for d in DAYS}
    tue_sat = {d: 1 for d in DAYS[2:]}
    coverage = [
        {'label': 'QS Zone 1', 'rows': [row_id('QS AUTOMATED EXT', 'Zone 1')], 'need': all_days},
        {'label': 'TIH', 'rows': [row_id('TIH', 'TIH')],
         'need': {'Sun': 2, 'Mon': 2, 'Tue': 2, 'Wed': 2, 'Thu': 1, 'Fri': 1, 'Sat': 1},
         'filter': {'Sun': 'cls_or_trainee', 'Mon': 'cls_or_trainee', 'Tue': 'cls_or_trainee',
                    'Wed': 'cls_or_trainee', 'Thu': 'cls', 'Fri': 'cls', 'Sat': 'cls'}},
        {'label': 'HZN EXT', 'rows': [row_id('HORIZON', 'EXT/NORM/DIL')],
         'need': {'Sun': 1, 'Mon': 1, 'Tue': 2, 'Wed': 2, 'Thu': 2, 'Fri': 2, 'Sat': 2}},
        {'label': 'POC swap', 'rows': [row_id('HORIZON', 'POC Swap')], 'need': {d: 4 for d in DAYS[2:]}},
        {'label': 'POC CLS', 'rows': [row_id('HORIZON', 'POC Swap')], 'need': {'Wed': 2}, 'filter': 'cls_or_trainee'},
        {'label': 'DNEasy', 'rows': [row_id('POC', 'DNEasy/Mix-1')], 'need': tue_sat},
        {'label': 'PGD', 'rows': [row_id('PGD', 'PGD')], 'need': tue_sat},
        {'label': 'TIU', 'rows': [row_id('TIU', 'TIU')],
         'need': {'Sun': 1, 'Mon': 1, 'Tue': 2, 'Wed': 2, 'Thu': 2, 'Fri': 2, 'Sat': 2}},
        {'label': 'ISO', 'cat': 'ISO / TECAN MAINT', 'need': {d: 4 for d in DAYS},
         'jump': row_id('ISO / TECAN MAINT', 'Zone A')},
    ]
    return {'categories': cats, 'roster': roster, 'coverage': coverage}


# =====================================================================
# Roster loading
# =====================================================================
STD_COLS_MAP = {
    'name': 'Name', 'employee name': 'Name',
    'shift': 'Shift',
    'iso': 'ISO', 'tiu': 'TIU', 'qs': 'QS', 'float': 'FLOAT',
    'cls': 'CLS', 'pgd': 'PGD', 'hzn': 'HZN', 'tih': 'TIH',
    'poc': 'POC', 'cla': 'CLA', 'cls trainee': 'CLS_TRAINEE',
}
REQUIRED_COLS = ['Name', 'Shift', 'ISO', 'TIU', 'QS', 'FLOAT', 'CLS', 'PGD', 'HZN', 'TIH', 'POC', 'CLA', 'CLS_TRAINEE']


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
    s = (s.replace('â€"', '-').replace('â€"', '-').replace('–', '-').replace('—', '-')
          .replace('_', '-').replace('/', '-'))
    s = re.sub(r'\bto\b|\bthru\b|\bthrough\b', '-', s, flags=re.IGNORECASE)
    s = re.sub(r'[^A-Za-z\- ]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


DATA_FILE = BASE_DIR / "today_active_workers_corrected.csv"


def _clean_roster(df: pd.DataFrame) -> pd.DataFrame:
    df.insert(0, '__roster_index', range(len(df)))
    df.columns = [str(c).strip() for c in df.columns]
    df = normalize_columns(df)
    df['Name'] = df['Name'].astype(str).str.strip()
    role_cols = [c for c in REQUIRED_COLS if c not in ['Name', 'Shift']]
    for c in role_cols:
        df[c] = (
            df[c].astype(str).str.strip().str.lower()
            .replace({'nan': '', 'none': '', 'no': '', 'false': '', '0': ''})
        )
    df['Shift'] = df['Shift'].astype(str).map(canon_basic)
    df = df[df['Name'].str.len() > 0]
    df = df[df['Name'].str.lower() != 'nan']
    df = df.sort_values('__roster_index').drop_duplicates(subset=['Name'], keep='first').reset_index(drop=True)
    return df


def _read_csv_any(src) -> pd.DataFrame:
    if hasattr(src, 'seek'):
        src.seek(0)
    try:
        return pd.read_csv(src, encoding="utf-8-sig")
    except UnicodeDecodeError:
        if hasattr(src, 'seek'):
            src.seek(0)
        return pd.read_csv(src, encoding="cp1252")


def load_data():
    if 'roster_upload' in st.session_state and st.session_state.roster_upload is not None:
        return _clean_roster(_read_csv_any(st.session_state.roster_upload))
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame()
    return _clean_roster(_read_csv_any(DATA_FILE))


# =====================================================================
# Shift parsing
# =====================================================================
IDX = DAY_IDX
TOKEN_MAP = {
    'sunday': 'Sun', 'sun': 'Sun', 'su': 'Sun', 's': 'Sun',
    'monday': 'Mon', 'mon': 'Mon', 'm': 'Mon',
    'tuesday': 'Tue', 'tues': 'Tue', 'tue': 'Tue', 'tu': 'Tue', 't': 'Tue',
    'wednesday': 'Wed', 'weds': 'Wed', 'wed': 'Wed', 'w': 'Wed',
    'thursday': 'Thu', 'thurs': 'Thu', 'thur': 'Thu', 'thu': 'Thu', 'th': 'Thu',
    'friday': 'Fri', 'fri': 'Fri', 'f': 'Fri',
    'saturday': 'Sat', 'sat': 'Sat', 'sa': 'Sat',
}
ALL_TOKENS = sorted(TOKEN_MAP.keys(), key=len, reverse=True)
DAY_ALT = r'(?:' + '|'.join(re.escape(t) for t in ALL_TOKENS) + r')'


def _token_to_day(tok: str):
    return TOKEN_MAP.get(tok.lower())


def _days_range(a: str, b: str):
    ai, bi = IDX[a], IDX[b]
    if ai <= bi:
        return DAYS[ai:bi + 1]
    return DAYS[ai:] + DAYS[:bi + 1]


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


# =====================================================================
# Session state + autosave
# =====================================================================
AUTOSAVE = BASE_DIR / "board_autosave.json"


def _init_state():
    st.session_state.setdefault('pto_by_day', {})
    st.session_state.setdefault('ot_by_day', {})
    st.session_state.setdefault('training_pairs', [])
    st.session_state.setdefault('board_state', None)
    st.session_state.setdefault('board_height', 780)


def autosave():
    try:
        payload = {
            'board': st.session_state.board_state,
            'pto_by_day': st.session_state.pto_by_day,
            'ot_by_day': st.session_state.ot_by_day,
            'training_pairs': st.session_state.training_pairs,
            'saved_at': datetime.now().isoformat(timespec='seconds'),
        }
        AUTOSAVE.write_text(json.dumps(payload), encoding='utf-8')
    except Exception:
        pass


def restore_autosave(force=False):
    if not AUTOSAVE.exists():
        return False
    try:
        data = json.loads(AUTOSAVE.read_text(encoding='utf-8'))
    except Exception:
        return False
    if force or st.session_state.board_state is None:
        st.session_state.board_state = data.get('board')
        st.session_state.pto_by_day = data.get('pto_by_day', {}) or {}
        st.session_state.ot_by_day = data.get('ot_by_day', {}) or {}
        st.session_state.training_pairs = data.get('training_pairs', []) or []
        if data.get('board'):
            st.session_state.restored_at = data.get('saved_at')
        # widget keys must be dropped so the day pickers pick up restored values
        for k in [k for k in st.session_state.keys() if str(k).startswith(('pto_days_', 'ot_days_'))]:
            del st.session_state[k]
        return True
    return False


_init_state()
if 'autosave_checked' not in st.session_state:
    st.session_state.autosave_checked = True
    restore_autosave()


# =====================================================================
# Board component (HTML written next to the script on startup)
# =====================================================================
BOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --ink:#0f1522; --slate:#161e2e; --slate2:#1b2436; --line:#24304a; --line2:#2f3d5c;
  --fog:#dce3ef; --mist:#98a4ba; --dim:#5c6a85; --ghost:#3a465e;
  --amber:#f5b841; --red:#f47b7b; --green:#4cd6a4; --blue:#7aa2f7;
  --fh:780px;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0;background:transparent;color:var(--fog);overflow:hidden;
  font-family:"IBM Plex Sans",-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;font-size:13px;line-height:1.45;
  -webkit-font-smoothing:antialiased}
button{font:inherit;color:inherit}
.wrap{background:var(--ink);border:1px solid var(--line);border-radius:14px;overflow:hidden;display:flex;flex-direction:column;height:var(--fh)}
.wrap.expanded{height:auto}
.toolbar{display:flex;align-items:center;gap:10px;padding:10px 12px;border-bottom:1px solid var(--line);background:var(--slate);flex-wrap:wrap}
.seg{display:inline-flex;background:var(--ink);border:1px solid var(--line);border-radius:9px;padding:2px;gap:1px}
.seg button{background:transparent;border:0;color:var(--mist);padding:5px 10px;border-radius:7px;cursor:pointer;font-weight:500;display:inline-flex;align-items:center}
.seg button:hover{color:var(--fog)}
.seg button.on{background:var(--line2);color:#fff}
.seg button.gap::after{content:"";display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--amber);margin-left:6px}
.search{position:relative;display:inline-block}
.search input{background:var(--ink);border:1px solid var(--line);border-radius:9px;color:var(--fog);padding:6px 10px 6px 30px;width:190px;outline:none;font:inherit}
.search input:focus{border-color:var(--blue)}
.search svg{position:absolute;left:9px;top:8px;width:14px;height:14px;fill:none;stroke:var(--dim);stroke-width:2;pointer-events:none}
.spacer{flex:1}
.tb{background:transparent;border:1px solid var(--line);border-radius:9px;color:var(--mist);padding:6px 10px;cursor:pointer;display:inline-flex;align-items:center;gap:6px}
.tb:hover:not(:disabled){color:var(--fog);border-color:var(--line2);background:var(--slate2)}
.tb:disabled{opacity:.35;cursor:default}
.tb .n{font-size:11px;background:var(--line2);color:var(--fog);border-radius:6px;padding:0 6px}
.sync{font-size:12px;color:var(--dim);display:inline-flex;align-items:center;gap:6px;padding-left:4px}
.sync i{width:7px;height:7px;border-radius:50%;background:var(--green);display:inline-block}
.sync.pending i{background:var(--amber)}
.board{overflow:auto;flex:1;position:relative;overscroll-behavior:contain}
.wrap.expanded .board{overflow:visible;flex:none}
table{border-collapse:separate;border-spacing:0;table-layout:fixed;min-width:100%}
col.rc{width:176px}
col.dc{width:214px}
.wrap.dayview col.dc{width:auto}
th,td{border-bottom:1px solid var(--line);border-left:1px solid var(--line);vertical-align:top;text-align:left;font-weight:400}
th:first-child,td:first-child{border-left:0}
thead th{position:sticky;top:0;z-index:5;background:var(--slate);padding:9px 10px 8px;border-bottom:1px solid var(--line2)}
thead th.corner{left:0;z-index:7;color:var(--mist);font-size:12px}
.dh-top{display:flex;align-items:baseline;gap:8px}
.dh-name{font-weight:600;font-size:14px;color:var(--fog)}
.dh-count{font-size:12px;color:var(--mist);font-variant-numeric:tabular-nums}
.dh-count b{color:var(--fog);font-weight:600;font-size:15px}
.dh-tools{margin-left:auto;display:inline-flex;gap:3px;align-self:center}
.icon{width:24px;height:24px;border-radius:7px;border:1px solid var(--line);background:transparent;color:var(--mist);cursor:pointer;display:inline-flex;align-items:center;justify-content:center;padding:0;line-height:1;font-size:15px}
.icon:hover{background:var(--slate2);color:var(--fog);border-color:var(--line2)}
.dh-badges{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;min-height:18px}
.bd{font-size:11px;padding:1px 7px;border-radius:6px;border:1px solid transparent;font-variant-numeric:tabular-nums;white-space:nowrap}
.bd.ok{color:var(--green);border-color:rgba(76,214,164,.25);background:rgba(76,214,164,.06)}
.bd.warn{color:var(--amber);border-color:rgba(245,184,65,.35);background:rgba(245,184,65,.09);cursor:pointer}
.bd.warn:hover{background:rgba(245,184,65,.18)}
.bd.note{color:var(--mist);border-color:var(--line);cursor:pointer}
.bd.note:hover{background:var(--slate2)}
.rolecell{position:sticky;left:0;z-index:2;background:var(--ink);padding:7px 10px;color:#c7d0e0;font-weight:500;white-space:nowrap;border-left:3px solid var(--accent,#3a465e) !important}
.rolecell .sub{display:block;font-size:11px;color:var(--dim);font-weight:400}
tr.catrow th{background:color-mix(in srgb,var(--accent) 16%,var(--ink));color:var(--accent);font-weight:600;padding:5px 10px;font-size:12px;border-top:1px solid var(--line2)}
tr.catrow td{background:color-mix(in srgb,var(--accent) 10%,var(--ink));border-top:1px solid var(--line2)}
tbody tr:not(.catrow):hover td.cell{background:rgba(255,255,255,.018)}
.cell{padding:5px 6px;min-height:38px;position:relative;transition:background .12s,box-shadow .12s}
.cell .chips{display:flex;flex-wrap:wrap;gap:4px;align-items:center;min-height:26px}
.cell .empty{color:var(--ghost);padding:2px 4px;user-select:none}
.addbtn{width:22px;height:22px;border-radius:6px;border:1px dashed var(--ghost);background:transparent;color:var(--dim);cursor:pointer;opacity:.35;padding:0;line-height:1;font-size:15px}
.cell:hover .addbtn,.addbtn:focus{opacity:1;color:var(--mist);border-color:var(--dim)}
.cell.ok{background:rgba(76,214,164,.07) !important;box-shadow:inset 0 0 0 1px rgba(76,214,164,.45)}
.cell.nope{background:rgba(244,123,123,.05) !important;box-shadow:inset 0 0 0 1px rgba(244,123,123,.28)}
.cell.over{background:rgba(122,162,247,.16) !important;box-shadow:inset 0 0 0 2px var(--blue)}
.chip{display:inline-flex;align-items:center;gap:5px;padding:3px 8px 3px 7px;border-radius:999px;background:var(--slate2);border:1px solid var(--line2);border-left:3px solid var(--accent,#3a465e);cursor:grab;user-select:none;max-width:100%;position:relative;transition:opacity .12s}
.chip:hover{border-color:#3b4b6e;border-left-color:var(--accent,#3a465e);background:#21304a}
.chip:active{cursor:grabbing}
.chip .nm{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:160px}
.wrap.dayview .chip .nm{max-width:none}
.chip .tg{font-size:10px;font-weight:600;padding:1px 5px;border-radius:4px;background:#2a3852;color:#aebbd6;white-space:nowrap;cursor:pointer}
.chip .tg.ot{background:rgba(122,162,247,.18);color:var(--blue);cursor:default}
.chip .tg.was{background:transparent;color:var(--dim);font-weight:400;padding:0;cursor:default;white-space:normal}
.chip .wn{color:var(--amber);font-weight:700;font-size:12px}
.chip.warn{border-color:rgba(245,184,65,.55)}
.chip.sel{box-shadow:0 0 0 2px var(--blue);background:#233355}
.chip.dim{opacity:.18}
.chip.dragging{opacity:.4}
.chip.flash{animation:flash 1.2s ease-out}
@keyframes flash{0%{box-shadow:0 0 0 3px rgba(122,162,247,.9)}100%{box-shadow:0 0 0 0 rgba(122,162,247,0)}}
tr.bench .rolecell{--accent:var(--amber)}
tr.bench td.cell{background:rgba(245,184,65,.035)}
tr.outrow .rolecell{--accent:var(--red)}
tr.outrow .chip{--accent:var(--red);opacity:.85}
tr.outrow td.cell{background:rgba(244,123,123,.03)}
tr.jump td,tr.jump th{animation:rowflash 1.4s ease-out}
@keyframes rowflash{0%{background-color:rgba(245,184,65,.22)}100%{}}
.statusbar{display:flex;align-items:center;gap:14px;padding:7px 12px;border-top:1px solid var(--line);background:var(--slate);font-size:12px;color:var(--mist);min-height:34px;flex-wrap:wrap}
.statusbar .hint{color:var(--blue);font-weight:500}
.statusbar .k{color:var(--dim)}
.legend{display:flex;gap:12px;flex-wrap:wrap;margin-left:auto}
.legend span{display:inline-flex;align-items:center;gap:5px;color:var(--dim)}
.legend i{width:8px;height:8px;border-radius:2px;display:inline-block}
.pop-bg{position:fixed;inset:0;z-index:49}
.pop{position:fixed;z-index:50;background:#141c2c;border:1px solid var(--line2);border-radius:11px;box-shadow:0 14px 40px rgba(0,0,0,.55);min-width:250px;max-width:330px;padding:6px;animation:pop .12s ease-out}
@keyframes pop{from{opacity:0;transform:translateY(-3px)}to{opacity:1;transform:none}}
.pop .hd{padding:8px 10px 7px;border-bottom:1px solid var(--line);margin-bottom:4px}
.pop .hd b{display:block;font-size:14px;font-weight:600}
.pop .hd .q{color:var(--mist);font-size:12px;margin-top:2px;line-height:1.5}
.pop .hd .q em{font-style:normal;color:var(--fog)}
.pop .it{display:flex;align-items:center;gap:8px;width:100%;text-align:left;background:transparent;border:0;padding:7px 10px;border-radius:7px;cursor:pointer;color:var(--fog)}
.pop .it:hover{background:var(--slate2)}
.pop .it.danger{color:var(--red)}
.pop .it small{margin-left:auto;color:var(--dim)}
.pop .sep{height:1px;background:var(--line);margin:4px 0}
.pop .cov{padding:6px 10px 8px;font-size:12px;line-height:1.7}
.pop .cov .g{color:var(--amber)} .pop .cov .o{color:var(--green)}
.modal-bg{position:fixed;inset:0;background:rgba(6,10,18,.66);z-index:60;backdrop-filter:blur(2px)}
.modal{position:absolute;left:50%;transform:translateX(-50%);width:min(620px,94vw);background:#141c2c;border:1px solid var(--line2);border-radius:14px;box-shadow:0 24px 60px rgba(0,0,0,.6);display:flex;flex-direction:column;max-height:min(580px,calc(100vh - 24px));animation:pop .14s ease-out}
.modal .mh{display:flex;align-items:center;gap:10px;padding:14px 16px 10px}
.modal .mh b{font-size:15px;font-weight:600}
.modal .mh .x{margin-left:auto}
.modal .ms{padding:0 16px 10px;display:flex;gap:8px;align-items:center}
.modal .ms input{flex:1;background:var(--ink);border:1px solid var(--line);border-radius:9px;color:var(--fog);padding:7px 10px;outline:none;font:inherit}
.modal .ms input:focus{border-color:var(--blue)}
.modal .mb{overflow:auto;padding:0 8px 10px}
.sec{padding:10px 8px 4px;color:var(--mist);font-size:12px;font-weight:600;display:flex;align-items:baseline;gap:8px}
.sec small{color:var(--dim);font-weight:400}
.pr{display:flex;align-items:center;gap:10px;padding:6px 8px;border-radius:8px}
.pr:hover{background:var(--slate2)}
.pr .who{flex:1;min-width:0}
.pr .who b{font-weight:500;display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pr .who span{font-size:11px;color:var(--dim)}
.pr .who span.unq{color:var(--amber)}
.pr .acts{display:inline-flex;gap:4px;flex-shrink:0}
.pr .acts button{background:transparent;border:1px solid var(--line);color:var(--mist);border-radius:7px;padding:4px 9px;cursor:pointer;font-size:12px;white-space:nowrap}
.pr .acts button:hover{border-color:var(--blue);color:var(--fog)}
.pr .acts button.go{background:var(--blue);border-color:var(--blue);color:#0b1020;font-weight:600}
.rowpick{display:flex;flex-direction:column;gap:2px}
.rowpick .ct{padding:8px 8px 3px;font-size:12px;font-weight:600;color:var(--accent)}
.rowpick button{display:flex;align-items:center;gap:8px;background:transparent;border:0;border-left:3px solid var(--accent);border-radius:7px;padding:6px 10px;color:var(--fog);cursor:pointer;text-align:left;width:100%}
.rowpick button:hover:not(:disabled){background:var(--slate2)}
.rowpick button:disabled{opacity:.45;cursor:default}
.rowpick button small{margin-left:auto;color:var(--dim);white-space:nowrap}
.rowpick button.unq small{color:var(--amber)}
.logl{list-style:none;margin:0;padding:0 8px 8px}
.logl li{display:grid;grid-template-columns:44px 34px 1fr;gap:8px;padding:6px 8px;border-bottom:1px solid var(--line);font-size:12px}
.logl li span:first-child{color:var(--dim);font-variant-numeric:tabular-nums}
.logl li span:nth-child(2){color:var(--mist)}
textarea.copy{width:100%;min-height:280px;background:var(--ink);color:var(--fog);border:1px solid var(--line);border-radius:9px;padding:10px;font:12px/1.5 ui-monospace,Menlo,Consolas,monospace;resize:vertical}
.toast{position:fixed;left:50%;transform:translateX(-50%);background:#1e2a42;border:1px solid var(--line2);color:var(--fog);padding:9px 14px;border-radius:10px;z-index:70;box-shadow:0 10px 30px rgba(0,0,0,.5);font-size:13px;max-width:90vw}
.toast.warn{border-color:rgba(245,184,65,.5)}
.empty-state{padding:56px 20px;text-align:center;color:var(--mist)}
.empty-state b{display:block;color:var(--fog);font-size:16px;font-weight:600;margin-bottom:6px}
:focus-visible{outline:2px solid var(--blue);outline-offset:2px}
@media (prefers-reduced-motion:reduce){*{animation:none !important;transition:none !important}}
</style>
</head>
<body>
<div id="app"></div>
<div id="layer"></div>
<script>
'use strict';
var app=document.getElementById('app'), layer=document.getElementById('layer');

/* ---------------- Streamlit bridge ---------------- */
function stSend(type,data){window.parent.postMessage(Object.assign({isStreamlitMessage:true,type:type},data||{}),'*');}
function stHeight(h){stSend('streamlit:setFrameHeight',{height:h});}
function stValue(v){stSend('streamlit:setComponentValue',{value:v,dataType:'json'});}

var DAYS=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
var BENCH='__bench', OUT='__out';
var CFG=null, S=null, ROWS={}, ROSTER={};
var hist=[], fut=[];
var view='week', query='', sel=null, expanded=false, frameH=780, drag=null, lastMoved=null, pushT=null, toastT=null, lastY=80;

window.addEventListener('message',function(ev){
  var d=ev.data; if(!d||d.type!=='streamlit:render')return;
  var a=d.args||{};
  frameH=Number(a.height)||780;
  document.documentElement.style.setProperty('--fh',frameH+'px');
  var cfg,st;
  try{cfg=JSON.parse(a.config_json);st=JSON.parse(a.state_json);}
  catch(e){app.innerHTML='<div class="empty-state"><b>Could not read the board data</b>'+esc(String(e))+'</div>';stHeight(200);return;}
  var changed=!S;
  CFG=cfg; indexConfig();
  if(!S||S.rev!==st.rev){S=st;hist=[];fut=[];sel=null;drag=null;closeLayer();changed=true;}
  else if((st.seq||0)>(S.seq||0)){S=st;changed=true;}
  ensureShape();
  if(changed)render(); else updateHeight();
});
window.addEventListener('resize',updateHeight);
stSend('streamlit:componentReady',{apiVersion:1});

/* ---------------- config / state shape ---------------- */
function indexConfig(){
  ROWS={}; ROSTER={};
  CFG.categories.forEach(function(c){c.rows.forEach(function(r){
    ROWS[r.id]=Object.assign({},r,{cat:c.id,catLabel:c.label,color:c.color,quals:r.quals||c.quals||[],tags:r.tags||[]});
  });});
  (CFG.roster||[]).forEach(function(p){ROSTER[p.name]=p;});
}
function ensureShape(){
  S.seq=S.seq||0; S.log=S.log||[];
  S.assignments=S.assignments||{}; S.bench=S.bench||{}; S.out=S.out||{}; S.added=S.added||{};
  DAYS.forEach(function(d){
    S.assignments[d]=S.assignments[d]||{}; S.bench[d]=S.bench[d]||[]; S.out[d]=S.out[d]||[]; S.added[d]=S.added[d]||[];
    Object.keys(S.assignments[d]).forEach(function(rid){if(!ROWS[rid])addDynamicRow(rid);});
  });
}
function addDynamicRow(rid){
  var parts=rid.split('||'), cat=parts[0], label=parts[1]||rid;
  var c=null; CFG.categories.forEach(function(x){if(x.id===cat)c=x;});
  if(!c){CFG.categories.forEach(function(x){if(x.id==='TRAINING')c=x;});}
  if(!c){c={id:'OTHER',label:'Other',color:'#9aa4b8',quals:[],rows:[]};CFG.categories.push(c);}
  var r={id:rid,label:label,short:label,quals:[],tags:(c.id==='TRAINING'?['Trainer','Trainee']:[])};
  c.rows.push(r);
  ROWS[rid]=Object.assign({},r,{cat:c.id,catLabel:c.label,color:c.color});
}

/* ---------------- helpers ---------------- */
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
function uid(){return Math.random().toString(36).slice(2,10)+Date.now().toString(36).slice(-3);}
function pad(n){return String(n).padStart(2,'0');}
function person(n){return ROSTER[n]||{name:n,quals:[],shiftDays:DAYS,shift:'',idx:9999};}
function onShift(n,d){var p=ROSTER[n];return p?p.shiftDays.indexOf(d)>=0:true;}
function hasQ(n,q){return person(n).quals.indexOf(q)>=0;}
function isCls(n){return hasQ(n,'CLS');}
function isClsOrTr(n){return hasQ(n,'CLS')||hasQ(n,'CLS_TRAINEE');}
function rowShort(loc){if(loc===BENCH)return 'Unassigned';if(loc===OUT)return 'Out';var r=ROWS[loc];return r?(r.short||r.label):loc;}
function rowQuals(loc){var r=ROWS[loc];return r?(r.quals||[]):[];}
function rowTags(loc){var r=ROWS[loc];return r&&r.tags?r.tags:[];}
function qualOK(n,loc){if(loc===BENCH||loc===OUT)return true;if(!ROSTER[n])return true;return rowQuals(loc).every(function(q){return hasQ(n,q);});}
function missingQ(n,loc){return rowQuals(loc).filter(function(q){return !hasQ(n,q);});}
function listAt(d,loc){if(loc===BENCH)return S.bench[d];if(loc===OUT)return S.out[d];if(!S.assignments[d][loc])S.assignments[d][loc]=[];return S.assignments[d][loc];}
function findChip(d,loc,id){var l=listAt(d,loc);for(var i=0;i<l.length;i++)if(l[i].id===id)return {list:l,idx:i,chip:l[i]};return null;}
function placesOf(d,n){var res=[];Object.keys(S.assignments[d]).forEach(function(rid){S.assignments[d][rid].forEach(function(c){if(c.name===n)res.push({loc:rid,chip:c});});});S.bench[d].forEach(function(c){if(c.name===n)res.push({loc:BENCH,chip:c});});return res;}
function isOut(d,n){return S.out[d].some(function(c){return c.name===n;});}
function workingSet(d){var s={};Object.keys(S.assignments[d]).forEach(function(rid){S.assignments[d][rid].forEach(function(c){s[c.name]=1;});});S.bench[d].forEach(function(c){s[c.name]=1;});return Object.keys(s);}
function addName(arr,n){if(arr.indexOf(n)<0)arr.push(n);}
function delName(arr,n){var i=arr.indexOf(n);if(i>=0)arr.splice(i,1);}
function removeAll(d,n){Object.keys(S.assignments[d]).forEach(function(rid){S.assignments[d][rid]=S.assignments[d][rid].filter(function(c){return c.name!==n;});});S.bench[d]=S.bench[d].filter(function(c){return c.name!==n;});}
function defaultTag(loc,cur){var t=rowTags(loc);if(!t.length)return null;if(cur&&t.indexOf(cur)>=0)return cur;return t[0]||null;}
function allRows(){var out=[];CFG.categories.forEach(function(c){c.rows.forEach(function(r){out.push(ROWS[r.id]);});});return out;}
function shownDays(){return view==='week'?DAYS:[view];}
function hasAnyOnDay(d){return workingSet(d).length>0||S.out[d].length>0;}
function cssEsc(s){return (window.CSS&&CSS.escape)?CSS.escape(s):s.replace(/([^a-zA-Z0-9_-])/g,'\\$1');}

/* ---------------- mutations (return true when something changed) ---------------- */
function doMove(fd,fl,id,td,tl,split){
  var f=findChip(fd,fl,id); if(!f)return false; var n=f.chip.name;
  if(fd===td&&fl===tl&&!split)return false;
  if(tl===OUT){ if(td!==fd){toast('Call-outs stay on the same day.',true);return false;} if(fl===OUT)return false; return doCallOut(fd,n); }
  if(fl===OUT&&td!==fd){toast('Bring '+n+' back on '+fd+' first, then move them.',true);return false;}
  if(td!==fd&&placesOf(td,n).length){toast(n+' is already on '+td+'.',true);return false;}
  var chip;
  if(split&&fl!==OUT){chip={id:uid(),name:n,tag:null};}
  else{chip=f.list.splice(f.idx,1)[0];}
  chip.tag=defaultTag(tl,chip.tag); delete chip.roles;
  if(td!==fd){
    S.out[td]=S.out[td].filter(function(c){return c.name!==n;});
    if(!placesOf(fd,n).length){delName(S.added[fd],n);if(onShift(n,fd))S.out[fd].push({id:uid(),name:n,tag:null,roles:['moved to '+td]});}
  }
  listAt(td,tl).push(chip);
  if((td!==fd||fl===OUT)&&!onShift(n,td))addName(S.added[td],n);
  lastMoved=chip.id; return true;
}
function doCallOut(d,n){
  var pl=placesOf(d,n); if(!pl.length)return false;
  var roles=pl.filter(function(p){return p.loc!==BENCH;}).map(function(p){return rowShort(p.loc);});
  removeAll(d,n);
  if(!isOut(d,n))S.out[d].push({id:uid(),name:n,tag:null,roles:roles.length?roles:['unassigned']});
  delName(S.added[d],n); return true;
}
function doAdd(d,n,tl){
  if(placesOf(d,n).length){toast(n+' is already on '+d+'.',true);return false;}
  tl=tl||BENCH;
  S.out[d]=S.out[d].filter(function(c){return c.name!==n;});
  var chip={id:uid(),name:n,tag:defaultTag(tl,null)};
  listAt(d,tl).push(chip);
  if(!onShift(n,d))addName(S.added[d],n);
  lastMoved=chip.id; return true;
}
function doRemoveFromDay(d,n){removeAll(d,n);S.out[d]=S.out[d].filter(function(c){return c.name!==n;});delName(S.added[d],n);return true;}
function doAssignExisting(d,n,tl,mode){
  var pl=placesOf(d,n); if(!pl.length)return doAdd(d,n,tl);
  if(mode==='split'){var chip={id:uid(),name:n,tag:defaultTag(tl,null)};listAt(d,tl).push(chip);lastMoved=chip.id;return true;}
  var src=null; pl.forEach(function(p){if(p.loc===BENCH&&!src)src=p;}); if(!src)src=pl[0];
  return doMove(d,src.loc,src.chip.id,d,tl,false);
}
function doCycleTag(d,loc,id){var f=findChip(d,loc,id);if(!f)return false;var t=rowTags(loc);if(!t.length)return false;var cur=f.chip.tag||'';var i=t.indexOf(cur);f.chip.tag=t[(i+1)%t.length]||null;return true;}

/* ---------------- history / sync ---------------- */
function snap(){return JSON.stringify([S.assignments,S.bench,S.out,S.added]);}
function restore(js){var o=JSON.parse(js);S.assignments=o[0];S.bench=o[1];S.out=o[2];S.added=o[3];}
function commit(msg,day,fn){
  var before=snap(), ok=false;
  try{ok=fn();}catch(e){console.error(e);ok=false;}
  if(!ok){restore(before);return false;}
  hist.push(before); if(hist.length>150)hist.shift(); fut=[];
  S.seq=(S.seq||0)+1; addLog(msg,day); sel=null; closeLayer(); render(); push(); return true;
}
function addLog(m,d){var t=new Date();S.log.unshift({t:pad(t.getHours())+':'+pad(t.getMinutes()),d:d||'',m:m});if(S.log.length>400)S.log.length=400;}
function undo(){if(!hist.length)return;fut.push(snap());restore(hist.pop());S.seq++;addLog('Undid the last change','');sel=null;closeLayer();render();push();}
function redo(){if(!fut.length)return;hist.push(snap());restore(fut.pop());S.seq++;addLog('Redid a change','');sel=null;closeLayer();render();push();}
function push(){setSync(true);clearTimeout(pushT);pushT=setTimeout(function(){stValue(S);setSync(false);},150);}
function setSync(p){var e=document.querySelector('.sync');if(!e)return;e.classList.toggle('pending',p);e.lastChild.textContent=p?'Saving':'Saved';}

/* ---------------- coverage ---------------- */
function coverage(d){
  var res=[];
  (CFG.coverage||[]).forEach(function(rule){
    var need=rule.need?rule.need[d]:null; if(!need)return;
    var rows=rule.rows||[];
    if(rule.cat)rows=allRows().filter(function(r){return r.cat===rule.cat;}).map(function(r){return r.id;});
    var filt=(rule.filter&&typeof rule.filter==='object')?rule.filter[d]:rule.filter;
    var names={};
    rows.forEach(function(rid){(S.assignments[d][rid]||[]).forEach(function(c){
      if(filt==='cls'&&!isCls(c.name))return;
      if(filt==='cls_or_trainee'&&!isClsOrTr(c.name))return;
      names[c.name]=1;
    });});
    res.push({label:rule.label,need:need,have:Object.keys(names).length,jump:rule.jump||rows[0]});
  });
  return res;
}

/* ---------------- render ---------------- */
function render(){
  if(!S||!CFG)return;
  var oldB=document.getElementById('board'); var sx=oldB?oldB.scrollLeft:0, sy=oldB?oldB.scrollTop:0;
  var days=shownDays();
  var hasAny=DAYS.some(hasAnyOnDay);
  var h='<div class="wrap'+(expanded?' expanded':'')+(view!=='week'?' dayview':'')+'">'+toolbarHTML()+'<div class="board" id="board">';
  if(!hasAny){
    h+='<div class="empty-state"><b>No schedule yet</b>Set PTO, overtime and training in the sidebar, then generate the week.</div>';
  }else{
    h+='<table><colgroup><col class="rc">'+days.map(function(){return '<col class="dc">';}).join('')+'</colgroup>';
    h+='<thead><tr><th class="corner">Role / zone</th>'+days.map(dayHeadHTML).join('')+'</tr></thead><tbody>';
    h+='<tr class="bench"><th class="rolecell">Unassigned<span class="sub">working, no role yet</span></th>'+days.map(function(d){return cellHTML(d,BENCH);}).join('')+'</tr>';
    CFG.categories.forEach(function(c){
      var rows=c.rows.map(function(r){return ROWS[r.id];});
      if(c.id==='TRAINING'){rows=rows.filter(function(r){return r.id==='TRAINING||Training'||DAYS.some(function(d){return (S.assignments[d][r.id]||[]).length>0;});});}
      if(!rows.length)return;
      h+='<tr class="catrow" style="--accent:'+c.color+'"><th class="rolecell">'+esc(c.label)+'</th><td colspan="'+days.length+'"></td></tr>';
      rows.forEach(function(r){
        h+='<tr data-row="'+esc(r.id)+'" style="--accent:'+c.color+'"><th class="rolecell">'+esc(r.label)+'</th>'+days.map(function(d){return cellHTML(d,r.id);}).join('')+'</tr>';
      });
    });
    h+='<tr class="outrow"><th class="rolecell">Out<span class="sub">PTO, sick, left early</span></th>'+days.map(function(d){return cellHTML(d,OUT);}).join('')+'</tr>';
    h+='</tbody></table>';
  }
  h+='</div>'+statusHTML()+'</div>';
  app.innerHTML=h;
  var b=document.getElementById('board'); b.scrollLeft=sx; b.scrollTop=sy;
  applyQuery();
  if(lastMoved){var el=b.querySelector('[data-chip="'+lastMoved+'"]');if(el){el.classList.add('flash');if(!expanded)el.scrollIntoView({block:'nearest',inline:'nearest'});}lastMoved=null;}
  if(sel)markTargets(sel.day,sel.name,sel.loc);
  updateHeight();
}
function toolbarHTML(){
  var segs=['week'].concat(DAYS).map(function(v){
    var on=v===view; var gap=v!=='week'&&hasAnyOnDay(v)&&coverage(v).some(function(c){return c.have<c.need;});
    return '<button data-act="view" data-view="'+v+'" class="'+(on?'on':'')+(gap?' gap':'')+'" title="'+(gap?'Coverage gaps on '+v:'')+'">'+(v==='week'?'Week':v)+'</button>';
  }).join('');
  return '<div class="toolbar"><div class="seg">'+segs+'</div>'+
    '<label class="search"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg><input id="q" type="search" placeholder="Find a name" value="'+esc(query)+'" aria-label="Find a name"></label>'+
    '<span class="spacer"></span>'+
    '<button class="tb" data-act="undo" '+(hist.length?'':'disabled')+' title="Undo (Ctrl+Z)">Undo</button>'+
    '<button class="tb" data-act="redo" '+(fut.length?'':'disabled')+' title="Redo (Ctrl+Y)">Redo</button>'+
    '<button class="tb" data-act="log" title="Every change made on this board">Changes'+(S.log.length?' <span class="n">'+S.log.length+'</span>':'')+'</button>'+
    '<button class="tb" data-act="expand">'+(expanded?'Fit to screen':'Show full board')+'</button>'+
    '<span class="sync"><i></i>Saved</span></div>';
}
function dayHeadHTML(d){
  var n=workingSet(d).length, cov=coverage(d), gaps=cov.filter(function(c){return c.have<c.need;});
  var b='';
  if(!hasAnyOnDay(d))b='<span class="bd note">no one scheduled</span>';
  else if(gaps.length)b=gaps.map(function(g){return '<span class="bd warn" data-act="jump" data-loc="'+esc(g.jump)+'" title="'+esc(g.label)+' needs '+g.need+', has '+g.have+'. Click to jump to the row.">'+esc(g.label)+' '+g.have+'/'+g.need+'</span>';}).join('');
  else b='<span class="bd ok">Covered</span>';
  if(S.bench[d].length)b+='<span class="bd note" data-act="jumpbench" title="Working, no role yet">'+S.bench[d].length+' unassigned</span>';
  if(S.out[d].length)b+='<span class="bd note" data-act="jumpout" title="PTO, sick or called out">'+S.out[d].length+' out</span>';
  return '<th class="dayhead" data-day="'+d+'"><div class="dh-top"><span class="dh-name">'+d+'</span><span class="dh-count"><b>'+n+'</b> working</span>'+
    '<span class="dh-tools"><button class="icon" data-act="addperson" data-day="'+d+'" title="Add a person to '+d+'">+</button><button class="icon" data-act="daymenu" data-day="'+d+'" title="More for '+d+'">&#8943;</button></span></div>'+
    '<div class="dh-badges">'+b+'</div></th>';
}
function cellHTML(d,loc){
  var list=listAt(d,loc);
  var inner=list.map(function(c){return chipHTML(d,loc,c);}).join('');
  if(loc!==OUT){inner+='<button class="addbtn" data-act="assign" data-day="'+d+'" data-loc="'+esc(loc)+'" title="'+(loc===BENCH?'Add someone to '+d:'Fill '+esc(rowShort(loc))+' on '+d)+'">+</button>';}
  else if(!list.length){inner='<span class="empty">&mdash;</span>';}
  return '<td class="cell" data-day="'+d+'" data-loc="'+esc(loc)+'"><div class="chips">'+inner+'</div></td>';
}
function chipHTML(d,loc,c){
  var n=c.name, warn=loc!==OUT&&!qualOK(n,loc), ot=loc!==OUT&&!onShift(n,d), p=ROSTER[n];
  var title=n+(p?'. Roles: '+(p.quals.join(', ')||'none marked')+(p.shift?'. Shift: '+p.shift:''):'');
  var t='<span class="chip'+(warn?' warn':'')+(sel&&sel.id===c.id?' sel':'')+'" draggable="true" data-chip="'+c.id+'" data-day="'+d+'" data-loc="'+esc(loc)+'" data-name="'+esc(n)+'" title="'+esc(title)+'"><span class="nm">'+esc(n)+'</span>';
  if(c.tag)t+='<span class="tg" data-act="tag" title="Click to switch">'+esc(c.tag)+'</span>';
  if(ot)t+='<span class="tg ot" title="Not on their regular shift">OT</span>';
  if(loc===OUT&&c.roles&&c.roles.length)t+='<span class="tg was">'+esc(c.roles.join(', '))+'</span>';
  if(warn)t+='<span class="wn" title="Not marked for '+esc(missingQ(n,loc).join(', '))+'">!</span>';
  return t+'</span>';
}
function statusHTML(){
  var hint=sel?'<span class="hint">Placing '+esc(sel.name)+'</span><span class="k">click a cell to move them there, or press Esc</span>'
               :'<span class="k">Drag names between cells. Click a name for options. Click + in any slot to fill it.</span>';
  var leg='<span class="legend">'+CFG.categories.filter(function(c){return c.id!=='TRAINING'&&c.id!=='OTHER';}).map(function(c){return '<span><i style="background:'+c.color+'"></i>'+esc(c.label)+'</span>';}).join('')+'</span>';
  return '<div class="statusbar">'+hint+leg+'</div>';
}
function applyQuery(){document.querySelectorAll('.chip').forEach(function(c){c.classList.toggle('dim',!!query&&c.dataset.name.toLowerCase().indexOf(query)<0);});}
function updateHeight(){requestAnimationFrame(function(){stHeight(expanded?document.body.scrollHeight+4:frameH+4);});}
function markTargets(d,name,fromLoc){
  document.querySelectorAll('.cell').forEach(function(c){
    var loc=c.dataset.loc, sameDay=c.dataset.day===d;
    if(loc===OUT){c.classList.toggle('ok',sameDay);c.classList.toggle('nope',!sameDay);return;}
    if(fromLoc===OUT&&!sameDay){c.classList.remove('ok');c.classList.add('nope');return;}
    var q=qualOK(name,loc); c.classList.toggle('ok',q); c.classList.toggle('nope',!q);
  });
}
function clearMarks(){document.querySelectorAll('.cell.ok,.cell.nope,.cell.over').forEach(function(c){c.classList.remove('ok','nope','over');});}
function jumpTo(loc){
  var tr=null;
  if(loc===BENCH)tr=document.querySelector('tr.bench'); else if(loc===OUT)tr=document.querySelector('tr.outrow');
  else tr=document.querySelector('tr[data-row="'+cssEsc(loc)+'"]');
  if(!tr)return;
  tr.scrollIntoView({block:'center'}); tr.classList.remove('jump'); void tr.offsetWidth; tr.classList.add('jump');
}
function moveMsg(n,fd,fl,td,tl){
  if(fd===td)return 'Moved '+n+': '+rowShort(fl)+' to '+rowShort(tl)+' ('+td+')';
  return 'Moved '+n+' from '+fd+' ('+rowShort(fl)+') to '+td+' ('+rowShort(tl)+')';
}
function toast(msg,warn){
  var t=document.querySelector('.toast'); if(!t){t=document.createElement('div');t.className='toast';document.body.appendChild(t);}
  t.textContent=msg; t.classList.toggle('warn',!!warn);
  t.style.top=Math.min(window.innerHeight-60,Math.max(12,lastY+28))+'px';
  clearTimeout(toastT); toastT=setTimeout(function(){t.remove();},3400);
}

/* ---------------- selection (click to move) ---------------- */
function setSel(ch){sel={id:ch.dataset.chip,day:ch.dataset.day,loc:ch.dataset.loc,name:ch.dataset.name};closeLayer();render();}
function clearSel(){sel=null;render();}
function placeSel(d,loc){
  var s=sel; if(!s)return;
  if(loc===OUT){ if(d!==s.day){toast('Call-outs stay on the same day.',true);return;} commit('Called out '+s.name+' ('+s.day+')',s.day,function(){return doCallOut(s.day,s.name);}); return; }
  var unq=!qualOK(s.name,loc);
  var ok=commit(moveMsg(s.name,s.day,s.loc,d,loc),d,function(){return doMove(s.day,s.loc,s.id,d,loc,false);});
  if(ok&&unq)toast(s.name+' is not marked for '+missingQ(s.name,loc).join(', ')+'. Placed anyway.',true);
  if(!ok&&s.day===d&&s.loc===loc)clearSel();
}

/* ---------------- events ---------------- */
app.addEventListener('click',function(ev){
  lastY=ev.clientY;
  var t=ev.target.closest('[data-act],.chip,.cell'); if(!t)return;
  var act=t.dataset.act;
  if(act==='view'){view=t.dataset.view;closeLayer();render();return;}
  if(act==='undo'){undo();return;}
  if(act==='redo'){redo();return;}
  if(act==='expand'){expanded=!expanded;closeLayer();render();return;}
  if(act==='log'){openLog();return;}
  if(act==='jump'){jumpTo(t.dataset.loc);return;}
  if(act==='jumpbench'){jumpTo(BENCH);return;}
  if(act==='jumpout'){jumpTo(OUT);return;}
  if(act==='addperson'){openPicker(t.dataset.day,BENCH,ev);return;}
  if(act==='daymenu'){openDayMenu(t.dataset.day,t);return;}
  if(act==='tag'){var ch=t.closest('.chip');if(sel){placeSel(ch.dataset.day,ch.dataset.loc);return;}commit('Switched tag for '+ch.dataset.name+' ('+rowShort(ch.dataset.loc)+', '+ch.dataset.day+')',ch.dataset.day,function(){return doCycleTag(ch.dataset.day,ch.dataset.loc,ch.dataset.chip);});return;}
  if(act==='assign'){ if(sel){placeSel(t.dataset.day,t.dataset.loc);return;} openPicker(t.dataset.day,t.dataset.loc,ev); return; }
  if(t.classList.contains('chip')){
    if(sel){
      if(sel.id===t.dataset.chip){clearSel();return;}
      if(!(sel.day===t.dataset.day&&sel.loc===t.dataset.loc)){placeSel(t.dataset.day,t.dataset.loc);return;}
    }
    openChipMenu(t); return;
  }
  if(t.classList.contains('cell')){
    if(sel){placeSel(t.dataset.day,t.dataset.loc);}
    else if(t.dataset.loc!==OUT){openPicker(t.dataset.day,t.dataset.loc,ev);}
  }
});
app.addEventListener('input',function(ev){if(ev.target&&ev.target.id==='q'){query=ev.target.value.trim().toLowerCase();applyQuery();}});
app.addEventListener('dragstart',function(ev){
  var ch=ev.target.closest?ev.target.closest('.chip'):null; if(!ch)return;
  drag={id:ch.dataset.chip,day:ch.dataset.day,loc:ch.dataset.loc,name:ch.dataset.name};
  ev.dataTransfer.effectAllowed='move'; try{ev.dataTransfer.setData('text/plain',ch.dataset.name);}catch(e){}
  ch.classList.add('dragging'); closeLayer(); sel=null; markTargets(drag.day,drag.name,drag.loc);
});
app.addEventListener('dragend',function(ev){var ch=ev.target.closest?ev.target.closest('.chip'):null;if(ch)ch.classList.remove('dragging');drag=null;clearMarks();});
app.addEventListener('dragover',function(ev){
  var c=ev.target.closest?ev.target.closest('.cell'):null; if(!c||!drag)return;
  ev.preventDefault(); ev.dataTransfer.dropEffect='move';
  if(!c.classList.contains('over')){document.querySelectorAll('.cell.over').forEach(function(x){x.classList.remove('over');});c.classList.add('over');}
});
app.addEventListener('dragleave',function(ev){var c=ev.target.closest?ev.target.closest('.cell'):null;if(c&&!c.contains(ev.relatedTarget))c.classList.remove('over');});
app.addEventListener('drop',function(ev){
  var c=ev.target.closest?ev.target.closest('.cell'):null; if(!c||!drag)return;
  ev.preventDefault(); lastY=ev.clientY; var d=drag; drag=null; clearMarks();
  var td=c.dataset.day, tl=c.dataset.loc;
  if(tl===OUT){ if(td!==d.day){toast('Call-outs stay on the same day.',true);return;} if(d.loc===OUT)return; commit('Called out '+d.name+' ('+d.day+')',d.day,function(){return doCallOut(d.day,d.name);}); return; }
  if(td===d.day&&tl===d.loc)return;
  var unq=!qualOK(d.name,tl);
  var ok=commit(moveMsg(d.name,d.day,d.loc,td,tl),td,function(){return doMove(d.day,d.loc,d.id,td,tl,false);});
  if(ok&&unq)toast(d.name+' is not marked for '+missingQ(d.name,tl).join(', ')+'. Placed anyway.',true);
});
document.addEventListener('keydown',function(ev){
  var tag=(document.activeElement&&document.activeElement.tagName)||'';
  if(ev.key==='Escape'){ if(layer.innerHTML){closeLayer();} else if(sel){clearSel();} else if(tag==='INPUT'){document.activeElement.blur();} return; }
  if(tag==='INPUT'||tag==='TEXTAREA'||tag==='SELECT')return;
  var mod=ev.ctrlKey||ev.metaKey, k=ev.key.toLowerCase();
  if(mod&&k==='z'&&!ev.shiftKey){ev.preventDefault();undo();}
  else if(mod&&(k==='y'||(k==='z'&&ev.shiftKey))){ev.preventDefault();redo();}
});

/* ---------------- popovers & modals ---------------- */
function closeLayer(){layer.innerHTML='';}
function openPop(anchorEl,html){
  var r=anchorEl.getBoundingClientRect();
  layer.innerHTML='<div class="pop-bg"></div><div class="pop" role="menu">'+html+'</div>';
  var pop=layer.querySelector('.pop'), vw=window.innerWidth, vh=window.innerHeight;
  var x=r.left, y=r.bottom+6, pw=pop.offsetWidth, ph=pop.offsetHeight;
  if(x+pw>vw-8)x=Math.max(8,vw-8-pw);
  if(y+ph>vh-8)y=Math.max(8,Math.min(r.top-6-ph,vh-8-ph));
  pop.style.left=x+'px'; pop.style.top=y+'px';
  layer.querySelector('.pop-bg').addEventListener('click',closeLayer);
  return pop;
}
function openModal(html,anchorY){
  layer.innerHTML='<div class="modal-bg"><div class="modal" role="dialog">'+html+'</div></div>';
  var bg=layer.querySelector('.modal-bg'), m=layer.querySelector('.modal');
  var vh=window.innerHeight, mh=m.offsetHeight;
  var top=(anchorY==null)?Math.max(12,Math.min(60,(vh-mh)/2)):Math.max(12,Math.min(anchorY-60,vh-mh-12));
  m.style.top=top+'px';
  bg.addEventListener('click',function(e){if(e.target===bg)closeLayer();});
  var x=m.querySelector('[data-x]'); if(x)x.addEventListener('click',closeLayer);
  var inp=m.querySelector('input[type=search]'); if(inp)inp.focus();
  return m;
}
function mBtn(m,label,cls,small){return '<button class="it '+(cls||'')+'" data-m="'+m+'">'+esc(label)+(small?'<small>'+esc(small)+'</small>':'')+'</button>';}
function closeX(){return '<span class="x"><button class="icon" data-x="1" title="Close">&times;</button></span>';}

function openChipMenu(ch){
  var d=ch.dataset.day, loc=ch.dataset.loc, id=ch.dataset.chip, n=ch.dataset.name, p=person(n);
  var places=placesOf(d,n).map(function(x){return rowShort(x.loc);});
  var h='<div class="hd"><b>'+esc(n)+'</b><div class="q">'+(p.quals.length?esc(p.quals.join(', ')):'no roles marked on the roster')+(p.shift?'<br>Shift: <em>'+esc(p.shift)+'</em>':'')+'<br>'+d+': <em>'+esc(loc===OUT?'out':places.join(' + '))+'</em></div></div>';
  if(loc===OUT){
    h+=mBtn('return','Bring back as unassigned')+mBtn('assign','Bring back into a role');
    h+='<div class="sep"></div>'+mBtn('remove','Remove from '+d,'danger');
  }else{
    h+=mBtn('pick','Move: click a cell to place')+mBtn('assign','Move to a role')+mBtn('split','Add a second role');
    if(loc!==BENCH)h+=mBtn('bench','Send to unassigned');
    var tags=rowTags(loc); if(tags.length){var cur=(findChip(d,loc,id)||{chip:{}}).chip.tag||'';var nxt=tags[(tags.indexOf(cur)+1)%tags.length];h+=mBtn('tag','Switch tag','',nxt?'to '+nxt:'clear');}
    h+='<div class="sep"></div>'+mBtn('out','Call out for '+d,'danger');
    if(S.added[d].indexOf(n)>=0)h+=mBtn('remove','Undo add to '+d,'danger');
  }
  var pop=openPop(ch,h);
  pop.addEventListener('click',function(ev){
    var b=ev.target.closest('[data-m]'); if(!b)return; var m=b.dataset.m;
    if(m==='pick'){setSel(ch);}
    else if(m==='assign'){openRowPicker(d,n,{fromLoc:loc,chipId:id,mode:'move'},ev);}
    else if(m==='split'){openRowPicker(d,n,{fromLoc:loc,chipId:id,mode:'split'},ev);}
    else if(m==='bench'){commit('Unassigned '+n+' from '+rowShort(loc)+' ('+d+')',d,function(){return doMove(d,loc,id,d,BENCH,false);});}
    else if(m==='tag'){commit('Switched tag for '+n+' ('+rowShort(loc)+', '+d+')',d,function(){return doCycleTag(d,loc,id);});}
    else if(m==='out'){commit('Called out '+n+' ('+d+')',d,function(){return doCallOut(d,n);});}
    else if(m==='return'){commit('Brought back '+n+' as unassigned ('+d+')',d,function(){return doMove(d,OUT,id,d,BENCH,false);});}
    else if(m==='remove'){commit('Removed '+n+' from '+d,d,function(){return doRemoveFromDay(d,n);});}
  });
}

function openRowPicker(d,n,opts,ev){
  var cur=placesOf(d,n).map(function(x){return x.loc;});
  var h='<div class="mh"><b>'+(opts.mode==='split'?'Add a second role for ':(opts.fromLoc===OUT?'Bring back ':'Move '))+esc(n)+' on '+d+'</b>'+closeX()+'</div>'+
        '<div class="ms"><input type="search" placeholder="Filter roles" id="rq"></div><div class="mb"><div class="rowpick" id="rp">';
  CFG.categories.forEach(function(c){
    var rows=c.rows.map(function(r){return ROWS[r.id];}); if(!rows.length)return;
    h+='<div class="ct" style="--accent:'+c.color+'">'+esc(c.label)+'</div>';
    rows.forEach(function(r){
      var cnt=(S.assignments[d][r.id]||[]).length, q=qualOK(n,r.id), here=cur.indexOf(r.id)>=0;
      var meta=here?'current':((q?'':'not marked '+missingQ(n,r.id).join('/')+', ')+(cnt?cnt+' assigned':'empty'));
      h+='<button data-row="'+esc(r.id)+'" class="'+(q?'':'unq')+'" style="--accent:'+c.color+'" '+(here?'disabled':'')+'><span>'+esc(r.label)+'</span><small>'+esc(meta)+'</small></button>';
    });
  });
  h+='</div></div>';
  var m=openModal(h,ev&&ev.clientY);
  m.querySelector('#rq').addEventListener('input',function(e){var v=e.target.value.toLowerCase();m.querySelectorAll('#rp button').forEach(function(b){b.style.display=b.textContent.toLowerCase().indexOf(v)>=0?'':'none';});});
  m.querySelector('#rp').addEventListener('click',function(e){
    var b=e.target.closest('button[data-row]'); if(!b||b.disabled)return; var tl=b.dataset.row; var unq=!qualOK(n,tl); var ok;
    if(opts.mode==='split')ok=commit('Added '+n+' to '+rowShort(tl)+' as a second role ('+d+')',d,function(){return doAssignExisting(d,n,tl,'split');});
    else if(opts.fromLoc===OUT)ok=commit('Brought back '+n+' into '+rowShort(tl)+' ('+d+')',d,function(){return doMove(d,OUT,opts.chipId,d,tl,false);});
    else ok=commit(moveMsg(n,d,opts.fromLoc,d,tl),d,function(){return doMove(d,opts.fromLoc,opts.chipId,d,tl,false);});
    if(ok&&unq)toast(n+' is not marked for '+missingQ(n,tl).join(', ')+'. Placed anyway.',true);
  });
}

function openPicker(d,loc,ev){
  var isBench=loc===BENCH, title=isBench?'Add someone to '+d:'Fill '+rowShort(loc)+' on '+d;
  var working=workingSet(d), outNames=S.out[d].map(function(c){return c.name;});
  var roster=(CFG.roster||[]).slice().sort(function(a,b){return (a.idx||0)-(b.idx||0);}).map(function(p){return p.name;});
  var secs=[];
  if(!isBench){
    var bench=S.bench[d].map(function(c){return c.name;});
    secs.push({t:'Unassigned',s:'working '+d+', no role yet',people:bench,act:'assign'});
    var others=working.filter(function(nm){return bench.indexOf(nm)<0&&!placesOf(d,nm).some(function(p){return p.loc===loc;});});
    secs.push({t:'Already on a role',s:'move them here, or give them this as a second role',people:others,act:'moveorsplit'});
  }
  secs.push({t:'Out on '+d,s:'PTO, sick or called out',people:outNames,act:'return'});
  var notWorking=roster.filter(function(nm){return working.indexOf(nm)<0&&outNames.indexOf(nm)<0;});
  secs.push({t:'Not scheduled '+d,s:'add as extra coverage',people:notWorking,act:'add'});
  var h='<div class="mh"><b>'+esc(title)+'</b>'+closeX()+'</div><div class="ms"><input type="search" id="pq" placeholder="Search names"></div><div class="mb" id="pl">';
  var any=false;
  secs.forEach(function(sec){
    if(!sec.people.length)return; any=true;
    var sorted=sec.people.slice().sort(function(a,b){return (qualOK(b,loc)?1:0)-(qualOK(a,loc)?1:0)||((person(a).idx||0)-(person(b).idx||0));});
    h+='<div class="sec">'+esc(sec.t)+'<small>'+esc(sec.s)+'</small></div>';
    sorted.forEach(function(nm){
      var p=person(nm), q=qualOK(nm,loc), meta=[];
      var roles=placesOf(d,nm).filter(function(x){return x.loc!==BENCH;}).map(function(x){return rowShort(x.loc);});
      if(!q)meta.push('not marked '+missingQ(nm,loc).join('/'));
      if(sec.act==='moveorsplit'&&roles.length)meta.push('now: '+roles.join(' + '));
      if(sec.act==='return'){var oc=null;S.out[d].forEach(function(c){if(c.name===nm)oc=c;});if(oc&&oc.roles)meta.push('was: '+oc.roles.join(', '));}
      if(sec.act==='add'&&p.shift)meta.push('shift '+p.shift);
      if(p.quals.length&&sec.act!=='moveorsplit')meta.push(p.quals.join(' '));
      var acts='';
      if(sec.act==='assign')acts='<button class="go" data-a="assign">Assign</button>';
      else if(sec.act==='moveorsplit')acts='<button class="go" data-a="move">Move here</button><button data-a="split">Second role</button>';
      else if(sec.act==='return')acts='<button class="go" data-a="return">'+(isBench?'Bring back':'Bring back here')+'</button>';
      else acts='<button class="go" data-a="add">'+(isBench?'Add to '+d:'Add and assign')+'</button>';
      h+='<div class="pr" data-name="'+esc(nm)+'"><div class="who"><b>'+esc(nm)+'</b><span class="'+(q?'':'unq')+'">'+esc(meta.join(', '))+'</span></div><div class="acts">'+acts+'</div></div>';
    });
  });
  if(!any)h+='<div class="empty-state">Everyone on the roster is already on '+d+'.</div>';
  h+='</div>';
  var m=openModal(h,ev&&ev.clientY);
  m.querySelector('#pq').addEventListener('input',function(e){
    var v=e.target.value.toLowerCase();
    m.querySelectorAll('.pr').forEach(function(r){r.style.display=r.dataset.name.toLowerCase().indexOf(v)>=0?'':'none';});
    m.querySelectorAll('.sec').forEach(function(s){var el=s.nextElementSibling,vis=false;while(el&&el.classList.contains('pr')){if(el.style.display!=='none')vis=true;el=el.nextElementSibling;}s.style.display=vis?'':'none';});
  });
  m.querySelector('#pl').addEventListener('click',function(e){
    var b=e.target.closest('[data-a]'); if(!b)return;
    var nm=b.closest('.pr').dataset.name, a=b.dataset.a, unq=!qualOK(nm,loc), ok=false, where=rowShort(loc);
    if(a==='assign'||a==='move')ok=commit('Assigned '+nm+' to '+where+' ('+d+')',d,function(){return doAssignExisting(d,nm,loc,'move');});
    else if(a==='split')ok=commit('Added '+nm+' to '+where+' as a second role ('+d+')',d,function(){return doAssignExisting(d,nm,loc,'split');});
    else if(a==='return')ok=commit('Brought back '+nm+(isBench?' as unassigned':' into '+where)+' ('+d+')',d,function(){return doAdd(d,nm,loc);});
    else if(a==='add')ok=commit('Added '+nm+' to '+d+(isBench?'':' as '+where),d,function(){return doAdd(d,nm,loc);});
    if(ok&&unq)toast(nm+' is not marked for '+missingQ(nm,loc).join(', ')+'. Placed anyway.',true);
  });
}

function openDayMenu(d,anchor){
  var cov=coverage(d);
  var h='<div class="hd"><b>'+d+'</b><div class="q">'+workingSet(d).length+' working, '+S.bench[d].length+' unassigned, '+S.out[d].length+' out</div></div>'+
        mBtn('focus',view===d?'Back to week view':'Focus on '+d)+mBtn('copy','Copy '+d+' as text')+mBtn('add','Add a person');
  if(cov.length)h+='<div class="sep"></div><div class="cov">'+cov.map(function(c){return '<span class="'+(c.have<c.need?'g':'o')+'">'+esc(c.label)+' '+c.have+'/'+c.need+'</span>';}).join('<br>')+'</div>';
  var pop=openPop(anchor,h);
  pop.addEventListener('click',function(ev){
    var b=ev.target.closest('[data-m]'); if(!b)return; var m=b.dataset.m;
    if(m==='focus'){view=(view===d?'week':d);closeLayer();render();}
    else if(m==='copy'){copyDay(d);}
    else if(m==='add'){openPicker(d,BENCH,ev);}
  });
}
function dayText(d){
  var L=[d+' - '+workingSet(d).length+' working'];
  CFG.categories.forEach(function(c){
    var rows=c.rows.map(function(r){return ROWS[r.id];}).filter(function(r){return (S.assignments[d][r.id]||[]).length;});
    if(!rows.length)return;
    L.push(''); L.push(c.label);
    rows.forEach(function(r){L.push('  '+r.label+': '+S.assignments[d][r.id].map(function(ch){return ch.name+(ch.tag?' ['+ch.tag+']':'');}).join(', '));});
  });
  if(S.bench[d].length){L.push('');L.push('Unassigned: '+S.bench[d].map(function(c){return c.name;}).join(', '));}
  if(S.out[d].length){L.push('Out: '+S.out[d].map(function(c){return c.name+(c.roles?' ('+c.roles.join(', ')+')':'');}).join(', '));}
  return L.join('\n');
}
function copyDay(d){
  var txt=dayText(d); closeLayer();
  var done=function(){toast('Copied '+d+' to the clipboard');};
  var fallback=function(){var m=openModal('<div class="mh"><b>'+d+' as text</b>'+closeX()+'</div><div class="mb" style="padding:0 16px 16px"><textarea class="copy" readonly>'+esc(txt)+'</textarea></div>',null);var ta=m.querySelector('textarea');ta.focus();ta.select();};
  if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(txt).then(done,fallback);} else fallback();
}
function openLog(){
  var h='<div class="mh"><b>Changes on this board</b>'+closeX()+'</div><div class="mb">';
  if(!S.log.length)h+='<div class="empty-state">No changes yet.</div>';
  else h+='<ul class="logl">'+S.log.map(function(e){return '<li><span>'+esc(e.t)+'</span><span>'+esc(e.d)+'</span><span>'+esc(e.m)+'</span></li>';}).join('')+'</ul>';
  openModal(h+'</div>',null);
}
</script>
</body>
</html>
"""


def _component_dir():
    candidates = [BASE_DIR / "_schedule_board", Path(tempfile.gettempdir()) / "schedule_board_component"]
    for d in candidates:
        try:
            d.mkdir(parents=True, exist_ok=True)
            f = d / "index.html"
            if not f.exists() or f.read_text(encoding='utf-8') != BOARD_HTML:
                f.write_text(BOARD_HTML, encoding='utf-8')
            return d
        except Exception:
            continue
    return None


# =====================================================================
# Roster + sidebar
# =====================================================================
df = load_data()

with st.sidebar:
    logo_b64 = get_base64_logo(BASE_DIR / "natera.png")
    if logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{logo_b64}" style="height:30px;margin-bottom:6px">', unsafe_allow_html=True)
    st.markdown("### Week setup")

    with st.expander("Roster", expanded=df.empty):
        if df.empty:
            st.warning(f"`{DATA_FILE.name}` wasn't found next to the app. Upload a roster CSV to continue.")
        else:
            st.markdown(f'<span class="note">{len(df)} people loaded from `{DATA_FILE.name}`.</span>', unsafe_allow_html=True)
        st.file_uploader("Use a different roster CSV", type=['csv'], key='roster_upload',
                         help="Same columns as today_active_workers_corrected.csv")

if df.empty:
    st.markdown('<div class="hero"><div><div class="t">3rd shift schedule board</div>'
                '<div class="s">Add a roster in the sidebar to get started.</div></div></div>', unsafe_allow_html=True)
    st.stop()

name_options = df['Name'].tolist()


def day_picker(label, current, key):
    current = [d for d in (current or []) if d in DAYS]
    default = current or None
    if hasattr(st, 'pills'):
        val = st.pills(label, DAYS, selection_mode='multi', default=default, key=key)
        return [d for d in DAYS if d in (val or [])]
    val = st.multiselect(label, DAYS, default=default, key=key)
    return [d for d in DAYS if d in (val or [])]


def _clear_days(store, emp, prefix):
    st.session_state[store][emp] = []
    st.session_state.pop(f'{prefix}{emp}', None)


def _clear_all_days(store, prefix):
    st.session_state[store] = {}
    for k in [k for k in st.session_state.keys() if str(k).startswith(prefix)]:
        del st.session_state[k]


def _entries_block(store, prefix, empty_text):
    entries = {e: d for e, d in st.session_state[store].items() if d}
    if not entries:
        st.markdown(f'<span class="note">{empty_text}</span>', unsafe_allow_html=True)
        return
    for emp, days_ in entries.items():
        c1, c2 = st.columns([6, 1])
        with c1:
            st.markdown(f'<div class="entry"><span class="nm">{_html.escape(emp)}</span><span class="dy">{", ".join(days_)}</span></div>',
                        unsafe_allow_html=True)
        with c2:
            st.button("×", key=f'rm_{prefix}{emp}', help=f"Clear {emp}", on_click=_clear_days, args=(store, emp, prefix))
    st.button("Clear all", key=f'clear_all_{prefix}', on_click=_clear_all_days, args=(store, prefix))


with st.sidebar:
    with st.expander("PTO / sick", expanded=True):
        emp = st.selectbox("Employee", name_options, key='pto_emp')
        st.session_state.pto_by_day[emp] = day_picker(
            "Days off", st.session_state.pto_by_day.get(emp, []), key=f'pto_days_{emp}')
        _entries_block('pto_by_day', 'pto_days_', "No PTO entered for this week.")

    with st.expander("Overtime"):
        ot_emp = st.selectbox("Employee", name_options, key='ot_emp')
        st.session_state.ot_by_day[ot_emp] = day_picker(
            "Extra days", st.session_state.ot_by_day.get(ot_emp, []), key=f'ot_days_{ot_emp}')
        _entries_block('ot_by_day', 'ot_days_', "No overtime entered for this week.")

    TRAINING_WORKFLOWS = ['QS', 'ISO', 'Floating', 'HZN', 'PGD', 'TIH', 'POC', 'DNEasy', 'TIU']

    def _add_training_pair():
        tr, te = st.session_state.tr_trainer, st.session_state.tr_trainee
        days_ = [d for d in DAYS if d in (st.session_state.get('tr_days') or [])]
        if tr == te:
            st.session_state.tr_msg = ('error', "Trainee and trainer must be different people.")
            return
        if not days_:
            st.session_state.tr_msg = ('error', "Pick at least one training day.")
            return
        st.session_state.training_pairs.append(
            {'trainee': te, 'trainer': tr, 'workflow': st.session_state.tr_workflow, 'days': days_})
        st.session_state.tr_msg = ('ok', f"Added {tr} training {te} on {st.session_state.tr_workflow}.")
        st.session_state.tr_days = []

    def _remove_training_pair(i):
        if 0 <= i < len(st.session_state.training_pairs):
            st.session_state.training_pairs.pop(i)

    with st.expander("Training pairs"):
        st.selectbox("Trainee", name_options, key='tr_trainee')
        st.selectbox("Trainer", name_options, key='tr_trainer')
        st.selectbox("Workflow", TRAINING_WORKFLOWS, key='tr_workflow')
        day_picker("Training days", st.session_state.get('tr_days', []), key='tr_days')
        st.button("Add training pair", on_click=_add_training_pair, key='tr_add')
        msg = st.session_state.pop('tr_msg', None)
        if msg:
            (st.error if msg[0] == 'error' else st.success)(msg[1])
        if st.session_state.training_pairs:
            for i, pair in enumerate(st.session_state.training_pairs):
                c1, c2 = st.columns([6, 1])
                with c1:
                    st.markdown(f'<div class="entry"><span class="nm">{_html.escape(pair["trainer"])} → {_html.escape(pair["trainee"])}</span>'
                                f'<span class="dy">{pair["workflow"]}, {", ".join(pair["days"])}</span></div>',
                                unsafe_allow_html=True)
                with c2:
                    st.button("×", key=f'rm_pair_{i}', on_click=_remove_training_pair, args=(i,))


# =====================================================================
# Scheduler (logic unchanged from the original generator)
# =====================================================================
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


CORE_ROLES = ['ISO', 'TIU', 'QS', 'FLOAT', 'CLS', 'PGD', 'HZN', 'TIH', 'POC', 'CLA', 'CLS_TRAINEE']
HZN_DNA_ROLE = 'HZN EXT/NORM/DIL'

ISO_ZONE_LIST = [f'Zone {c}' for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']]
QS_ZONE_LIST = [f'Zone {i}' for i in range(1, 14)]
ISO_ZONE_ORDER = {z: i + 1 for i, z in enumerate(ISO_ZONE_LIST)}
QS_ZONE_ORDER = {z: i + 1 for i, z in enumerate(QS_ZONE_LIST)}
QS_STEAL_ORDER = QS_ZONE_LIST[:]


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


def _df_row_by_name(name: str):
    return df.loc[df['Name'] == name].iloc[0]


def safe_assign(assign_map, assigned, day, name, role):
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
    rank = {z: i for i, z in enumerate(QS_STEAL_ORDER)}
    pairs.sort(key=lambda p: rank.get(p[0], 999))
    return pairs


def _steal_from_qs(assign_map, day, want_role, predicate):
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
    filled_qs_zones = set()
    for (dkey, _name), roles in assign_map.items():
        if dkey != day:
            continue
        for r in roles:
            if r.startswith('QS Zone '):
                filled_qs_zones.add(r.replace('QS ', '').strip())

    next_zone = None
    for z in QS_ZONE_LIST:
        if z not in filled_qs_zones:
            next_zone = z
            break
    if next_zone is None:
        return
    if next_zone != 'Zone 1':
        return

    zone_label = f'QS {next_zone}'
    qs_unassigned = pick(pool, lambda r: r['QS'].strip().lower() == 'yes' and r['Name'] not in assigned)
    names = priority_names(qs_unassigned, assigned, reserve_cls=True, limit=1)
    if names:
        n = names[0]
        assign_map[(day, n)].append(zone_label)
        assigned.add(n)
        return

    stolen = _steal_from_floaters(
        assign_map, day, zone_label,
        predicate=lambda row: row['QS'].strip().lower() == 'yes'
    )
    if stolen:
        return

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
    cls = str(row.get('CLS', '')).strip().lower() == 'yes'
    trainee = str(row.get('CLS_TRAINEE', '')).strip().lower() == 'yes'
    return cls or trainee


def reserve_hzn_ext(day, pool, assigned, weekly_hzn_ext_used, weekly_poc_used, limit=3, swap_floor=4):
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
    for n in (tih_reserved_names or []):
        assign_map[(day, n)].append('TIH_CLS')

    if day in ['Sun', 'Mon', 'Tue', 'Wed']:
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

    if day in ['Sun', 'Mon', 'Tue', 'Wed']:
        tih_pool = pick(pool, lambda r: is_cls_or_trainee(r)
                        and str(r.get('TIH', '')).strip().lower() == 'yes'
                        and r['Name'] not in assigned)
    else:
        tih_pool = pick(pool, lambda r: str(r.get('CLS', '')).strip().lower() == 'yes'
                        and str(r.get('TIH', '')).strip().lower() == 'yes'
                        and r['Name'] not in assigned)

    names = priority_names(tih_pool, assigned, reserve_cls=False, limit=needed)
    for n in names:
        assign_map[(day, n)].append('TIH_CLS')
        assigned.add(n)
        needed -= 1

    for _ in range(needed):
        _steal_from_qs(
            assign_map, day, 'TIH_CLS',
            predicate=lambda row: row.get('TIH', '') == 'yes' and is_cls_or_trainee(row)
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
            _steal_from_qs(assign_map, day, 'TIH_CLS',
                           predicate=lambda row: row['TIH'] == 'yes' and row['CLS'] == 'yes')

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
            _steal_from_qs(assign_map, day, 'HZN EXT/NORM/DIL',
                           predicate=lambda row: row['HZN'] == 'yes')


def final_fill_no_unassigned(day, pool, assigned, assign_map):
    working_names = set(pool['Name'])
    leftovers = sorted(
        list(working_names - assigned),
        key=lambda n: int(df.loc[df['Name'] == n, '__roster_index'].iloc[0])
    )
    qs_zone_list_active = [z for z in QS_ZONE_LIST if z != 'Zone 5']
    filled_qs = set()
    for (dkey, _n), roles in assign_map.items():
        if dkey == day:
            for r in roles:
                if r.startswith('QS Zone '):
                    filled_qs.add(r.replace('QS ', '').strip())
    next_qs_idx = len(filled_qs)
    overflow_counter = 0
    for name in leftovers:
        row = _df_row_by_name(name)
        qs_yes = str(row.get('QS', '')).strip().lower() == 'yes'
        if qs_yes and next_qs_idx < len(qs_zone_list_active):
            assign_map[(day, name)].append(f'QS {qs_zone_list_active[next_qs_idx]}')
            next_qs_idx += 1
        else:
            overflow_counter += 1
            assign_map[(day, name)].append(f'Floater {overflow_counter}')
        assigned.add(name)


def generate_week():
    """Runs the weekly generator. Returns {day: {'assign': {name: [tokens]},
    'training': {name: (Trainer|Trainee, workflow)}, 'working': set(names)}}."""
    days_data = {}
    prev_tiu = set()
    prev_dne = set()
    prev_pgd = set()
    prev_hzn_ext = set()
    prev_hzn_poc = set()
    prev_iso = set()
    weekly_poc_used = set()
    weekly_iso_count = {}
    weekly_pgd_used = set()
    weekly_hzn_ext_used = set()

    for day in DAYS:
        pool = working_pool(day)
        working_names = set(pool['Name'])
        assigned = set()
        assign_map = defaultdict(list)
        training_roles = {}

        # --- Training pairs: assign first, lock both people in ---
        for pair in st.session_state.get('training_pairs', []):
            if day not in pair['days']:
                continue
            trainee, trainer, workflow = pair['trainee'], pair['trainer'], pair['workflow']
            if trainee not in working_names or trainer not in working_names:
                continue
            if trainer in assigned or trainee in assigned:
                continue
            role_label = f'{workflow} Training'
            assign_map[(day, trainer)].append(role_label)
            assign_map[(day, trainee)].append(role_label)
            assigned.add(trainer)
            assigned.add(trainee)
            training_roles[trainer] = ('Trainer', workflow)
            training_roles[trainee] = ('Trainee', workflow)
            if workflow == 'ISO' and day != 'Sun':
                weekly_iso_count[trainer] = weekly_iso_count.get(trainer, 0) + 1
                weekly_iso_count[trainee] = weekly_iso_count.get(trainee, 0) + 1
            if workflow in ('POC', 'DNEasy'):
                weekly_poc_used.add(trainer)
                weekly_poc_used.add(trainee)

        pgd_reserved = set()
        pgd_reserved_names = []
        dne_reserved = set()
        dne_reserved_names = []
        hzn_ext_reserved = set()
        hzn_ext_reserved_names = []
        hzn_poc_reserved = set()
        swap_reserved = []

        # --- TIH minimum, reserved first ---
        tih_min_today = 2 if day in ['Sun', 'Mon', 'Tue', 'Wed'] else 1
        if day in ['Sun', 'Mon', 'Tue', 'Wed']:
            tih_reserve_pool = pick(pool, lambda r: is_cls_or_trainee(r)
                                    and str(r.get('TIH', '')).strip().lower() == 'yes')
        else:
            tih_reserve_pool = pick(pool, lambda r: str(r.get('CLS', '')).strip().lower() == 'yes'
                                    and str(r.get('TIH', '')).strip().lower() == 'yes')
        tih_reserved_names = priority_names(tih_reserve_pool, assigned, reserve_cls=False, limit=tih_min_today)
        tih_reserved = set(tih_reserved_names)
        assigned.update(tih_reserved)

        if day not in ['Sun', 'Mon']:
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
            assigned.update(dne_reserved)

            hzn_ext_reserved_names = reserve_hzn_ext(
                day, pool, assigned, weekly_hzn_ext_used, weekly_poc_used, limit=3
            )
            hzn_ext_reserved = set(hzn_ext_reserved_names)

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
            hzn_ext_reserved_names = reserve_hzn_ext(
                day, pool, assigned, weekly_hzn_ext_used, weekly_poc_used, limit=3
            )
            hzn_ext_reserved = set(hzn_ext_reserved_names)

        reserved = hzn_poc_reserved | hzn_ext_reserved | pgd_reserved | dne_reserved | tih_reserved

        ISO_MIN = 4
        QS_MIN = 4

        def iso_under_cap(r):
            n = r['Name']
            count = weekly_iso_count.get(n, 0)
            quals = sum(1 for c in CORE_ROLES if str(r.get(c, '')).strip().lower() == 'yes')
            return count < (2 if quals <= 2 else 1)

        if day == 'Sun':
            sun_pool = pick(pool, lambda r: r['ISO'].strip().lower() == 'yes'
                            and r['FLOAT'].strip().lower() == 'yes'
                            and r['Name'] not in reserved)
            sun_names = priority_names_excluding(sun_pool, assigned, exclude_set=prev_iso, reserve_cls=True,
                                                 limit=len(ISO_ZONE_LIST), prefer_more_skills=True)
            for i, name in enumerate(sun_names):
                safe_assign(assign_map, assigned, day, name, f'Tecan Maintenance/Rack Disposal {ISO_ZONE_LIST[i]}')
            iso_all = []
            float_all = []
        elif day == 'Mon':
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
            iso_pool_all = pick(pool, lambda r: r['ISO'].strip().lower() == 'yes' and r['Name'] not in reserved)
            iso_pool_capped = pick(iso_pool_all, iso_under_cap)
            iso_pool_filtered = iso_pool_capped if len(iso_pool_capped) >= ISO_MIN else iso_pool_all
            iso_all = priority_names_excluding(iso_pool_filtered, assigned, exclude_set=prev_iso, reserve_cls=True,
                                               limit=len(ISO_ZONE_LIST), prefer_more_skills=True)
            float_pool_filtered = pick(pool, lambda r: r['FLOAT'].strip().lower() == 'yes'
                                       and r['Name'] not in reserved
                                       and r['Name'] not in set(iso_all))
            float_all = priority_names(float_pool_filtered, assigned | set(iso_all), reserve_cls=True, limit=len(ISO_ZONE_LIST))

        max_pairs = min(len(iso_all), len(float_all), len(ISO_ZONE_LIST))
        n_pairs = min(ISO_MIN, max_pairs)
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

        if day not in ['Sun', 'Mon']:
            for i, name in enumerate(iso_all[:n_pairs]):
                safe_assign(assign_map, assigned, day, name, f'ISO {ISO_ZONE_LIST[i]}')
            zone_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
            for i, name in enumerate(float_all[:n_pairs]):
                if i < len(zone_letters):
                    label = f'Floater {zone_letters[i]}'
                else:
                    label = f'Floater {i - len(zone_letters) + 1}'
                safe_assign(assign_map, assigned, day, name, label)

        # --- QS ---
        qs_zone_list_filtered = [z for z in QS_ZONE_LIST if z != 'Zone 5']
        qs_pool = pick(pool, lambda r: r['QS'].strip().lower() == 'yes'
                       and r['Name'] not in assigned and r['Name'] not in reserved)
        qs_limit_today = 8 if day == 'Mon' else len(qs_zone_list_filtered)
        qs_names = priority_names(qs_pool, assigned, reserve_cls=True, limit=qs_limit_today)
        for i, name in enumerate(qs_names):
            safe_assign(assign_map, assigned, day, name, f'QS {qs_zone_list_filtered[i]}')

        # --- TIU ---
        cls_pool = pick(pool, lambda r: r['CLS'].strip().lower() == 'yes'
                        and r['Name'] not in assigned and r['Name'] not in reserved)
        tiu_pool = pick(cls_pool, lambda r: r['TIU'].strip().lower() == 'yes')
        if day == 'Mon':
            for name in priority_names_excluding(tiu_pool, assigned, exclude_set=prev_tiu, reserve_cls=False, limit=1):
                safe_assign(assign_map, assigned, day, name, 'TIU/Stickers')
        elif day == 'Sun':
            for name in priority_names_excluding(tiu_pool, assigned, exclude_set=prev_tiu, reserve_cls=False, limit=1):
                safe_assign(assign_map, assigned, day, name, 'TIU')
        else:
            for name in priority_names_excluding(tiu_pool, assigned, exclude_set=prev_tiu, reserve_cls=False, limit=2):
                safe_assign(assign_map, assigned, day, name, 'TIU')

        # --- HZN EXT ---
        for name in hzn_ext_reserved_names:
            safe_assign(assign_map, assigned, day, name, 'HZN EXT/NORM/DIL')

        # --- HZN POC Swap ---
        if day not in ['Sun', 'Mon']:
            swap_selected = list(hzn_poc_reserved)
            half = min(2, (len(swap_selected) + 1) // 2)
            for name in swap_selected[:half]:
                safe_assign(assign_map, assigned, day, name, 'HZN POC Swap (First Half HZN / Second Half POC)')
            for name in swap_selected[half:]:
                safe_assign(assign_map, assigned, day, name, 'HZN POC Swap (First Half POC / Second Half HZN)')

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
                    and str(r.get('POC', '')).strip().lower() == 'yes'
                    and str(r.get('HZN', '')).strip().lower() == 'yes'
                    and r['Name'] not in assigned
                    and r['Name'] not in weekly_poc_used
                )
                for name in priority_names(poc_extra_pool, assigned, reserve_cls=False, limit=needed_poc):
                    safe_assign(assign_map, assigned, day, name, 'HZN POC Swap (First Half POC / Second Half HZN)')

        # --- DNEasy ---
        if day not in ['Sun', 'Mon']:
            for name in dne_reserved_names:
                assign_map[(day, name)].append('DNEasy/Mix-1')

        # --- TIH ---
        enforce_tih_minimum(assign_map, day, pool, assigned, tih_reserved_names)
        tih_cla_pool = pick(
            pool,
            lambda r: r['TIH'].strip().lower() == 'yes' and r['CLA'].strip().lower() == 'yes' and r['Name'] not in assigned
        )
        for name in priority_names(tih_cla_pool, assigned, reserve_cls=True, limit=1):
            safe_assign(assign_map, assigned, day, name, 'TIH_CLA')

        # --- PGD ---
        if day not in ['Sun', 'Mon']:
            for name in pgd_reserved_names:
                assign_map[(day, name)].append('PGD')

        if day in ['Sun', 'Mon']:
            enforce_sun_mon_mins(assign_map, day, pool, assigned)

        enforce_qs_minimum(assign_map, day, pool, assigned)
        final_fill_no_unassigned(day, pool, assigned, assign_map)

        # --- weekly trackers ---
        weekly_poc_used.update(
            n for (dkey, n), roles in assign_map.items()
            if dkey == day and any('HZN POC Swap' in r or r == 'DNEasy/Mix-1' or r == 'DNEasy' for r in roles)
        )
        for (dkey, n), roles in assign_map.items():
            if dkey == day and any(r.startswith('ISO Zone') or r.startswith('ISO ') for r in roles):
                if day != 'Sun':
                    weekly_iso_count[n] = weekly_iso_count.get(n, 0) + 1
        prev_tiu = {n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'TIU' for r in roles)}
        prev_dne = {n for (dkey, n), roles in assign_map.items() if dkey == day and any(r in ('DNEasy/Mix-1', 'DNEasy') for r in roles)}
        prev_pgd = {n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'PGD' for r in roles)}
        weekly_pgd_used.update(prev_pgd)
        prev_hzn_ext = {n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'HZN EXT/NORM/DIL' for r in roles)}
        weekly_hzn_ext_used.update(prev_hzn_ext)
        prev_hzn_poc = {n for (dkey, n), roles in assign_map.items() if dkey == day and any('HZN POC Swap' in r for r in roles)}
        prev_iso = {n for (dkey, n), roles in assign_map.items() if dkey == day and any(r.startswith('ISO ') for r in roles)}

        days_data[day] = {
            'assign': {n: list(roles) for (dkey, n), roles in assign_map.items() if dkey == day},
            'training': training_roles,
            'working': working_names,
        }
    return days_data


# =====================================================================
# Board state <-> generator output <-> CSV
# =====================================================================
def _chip(name, tag=None, roles=None):
    c = {'id': uuid.uuid4().hex[:10], 'name': name, 'tag': tag}
    if roles:
        c['roles'] = roles
    return c


def board_from_days(days_data):
    roster_idx = {n: int(i) for n, i in zip(df['Name'], df['__roster_index'])}
    board = {'rev': uuid.uuid4().hex[:8], 'seq': 0, 'assignments': {}, 'bench': {}, 'out': {}, 'added': {},
             'log': [], 'generated_at': datetime.now().strftime('%a %b %d, %H:%M')}
    for day in DAYS:
        dd = days_data[day]
        A = OrderedDict()
        bench = []
        for name in sorted(dd['assign'], key=lambda n: roster_idx.get(n, 9999)):
            tokens = dd['assign'][name]
            if not tokens:
                bench.append(_chip(name))
                continue
            for token in tokens:
                if token.endswith('Training'):
                    role, wf = dd['training'].get(name, ('Trainer', token.replace(' Training', '')))
                    A.setdefault(row_id('TRAINING', f'{wf} Training'), []).append(_chip(name, role))
                elif token.startswith('ISO ') and '/Zone ' in token:
                    for z in token[len('ISO '):].split('/'):
                        A.setdefault(row_id('ISO / TECAN MAINT', z.strip()), []).append(_chip(name))
                else:
                    cls = classify_role(token)
                    if cls is None:
                        bench.append(_chip(name))
                        continue
                    cat, sub, tag = cls
                    A.setdefault(row_id(cat, sub), []).append(_chip(name, tag))
        for name in sorted(dd['working'] - set(dd['assign']), key=lambda n: roster_idx.get(n, 9999)):
            bench.append(_chip(name))
        # People who would normally work today but are on PTO show in the Out row,
        # so a lead can bring them back with one click if plans change.
        out = []
        for _, r in df.iterrows():
            n = r['Name']
            if day in st.session_state.pto_by_day.get(n, []) and day in expand_shift_days(r['Shift']):
                out.append(_chip(n, roles=['PTO']))
        board['assignments'][day] = A
        board['bench'][day] = bench
        board['out'][day] = out
        board['added'][day] = []
    return board


def _token_for(day, cat, sub, tag):
    if cat == 'ISO / TECAN MAINT':
        return f'Tecan Maintenance/Rack Disposal {sub}' if day == 'Sun' else f'ISO {sub}'
    if cat == 'QS AUTOMATED EXT':
        return f'QS {sub}'
    if cat == 'HORIZON':
        if sub == 'POC Swap':
            return ('HZN POC Swap (First Half POC / Second Half HZN)' if tag == 'P\u2192H'
                    else 'HZN POC Swap (First Half HZN / Second Half POC)')
        return 'HZN EXT/NORM/DIL'
    if cat == 'POC':
        return 'DNEasy/Mix-1'
    if cat == 'PGD':
        return 'PGD'
    if cat == 'TIH':
        return 'TIH_CLA' if tag == 'CLA' else 'TIH_CLS'
    if cat == 'TIU':
        return 'TIU/Stickers' if day == 'Mon' else 'TIU'
    if cat == 'TRAINING':
        return f'{sub} ({tag})' if tag else sub
    return sub


def _workflow_string(day, items):
    items = list(items)
    toks = []
    iso_zones = [sub for cat, sub, _ in items if cat == 'ISO / TECAN MAINT']
    if day == 'Mon' and len(iso_zones) >= 2:
        toks.append('ISO ' + '/'.join(iso_zones))
        items = [i for i in items if i[0] != 'ISO / TECAN MAINT']
    toks += [_token_for(day, *i) for i in items]
    return ' / '.join(toks) if len(toks) == 2 else ', '.join(toks)


def board_to_frame(board):
    roster_idx = {n: int(i) for n, i in zip(df['Name'], df['__roster_index'])}
    rows = []
    for day in DAYS:
        per = OrderedDict()
        for rid, chips in (board.get('assignments', {}).get(day) or {}).items():
            cat, _, sub = rid.partition('||')
            for c in chips:
                per.setdefault(c['name'], []).append((cat, sub, c.get('tag')))
        for c in board.get('bench', {}).get(day) or []:
            per.setdefault(c['name'], [])
        added = set(board.get('added', {}).get(day) or [])
        for name, items in per.items():
            wf = _workflow_string(day, items) if items else 'Unassigned'
            rows.append((day, roster_idx.get(name, 9999), name, wf, 'Added day-of' if name in added else ''))
        for c in board.get('out', {}).get(day) or []:
            rows.append((day, roster_idx.get(c['name'], 9999), c['name'], '',
                         'Out (' + ', '.join(c.get('roles') or ['called out']) + ')'))
    out = pd.DataFrame(rows, columns=['Day', '__roster_index', 'Name', 'Workflow', 'Status'])
    if not out.empty:
        out['__d'] = out['Day'].map(DAY_IDX)
        out = out.sort_values(['__d', '__roster_index'], kind='stable').drop(columns='__d').reset_index(drop=True)
    return out


def block_rank(workflow: str) -> float:
    w = str(workflow)
    if w.startswith('Tecan Maintenance') or w.startswith('ISO '):
        return 1
    if w.startswith('QS Zone '):
        return 2
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
    if w.startswith('DNEasy'):
        return 7.5
    if w.endswith('Training') or 'Training' in w:
        return 9
    if w == 'Unassigned':
        return 9.5
    if w == '':
        return 11
    return 10


def zone_rank(workflow: str) -> int:
    w = str(workflow)
    if w.startswith('ISO Zone '):
        return ISO_ZONE_ORDER.get(w.replace('ISO ', '').split('/')[0], 999)
    if w.startswith('QS Zone '):
        return QS_ZONE_ORDER.get(w.replace('QS ', ''), 999)
    if w.startswith('Tecan Maintenance/Rack Disposal Zone '):
        return ord(w[-1]) - ord('A')
    if w.startswith('Floater'):
        m = re.search(r'Floater ([A-H])', w)
        if m:
            return ord(m.group(1)) - ord('A')
        m = re.search(r'Floater (\d+)', w)
        return 100 + int(m.group(1)) if m else 200
    return 500


# =====================================================================
# Sidebar: generate / board controls
# =====================================================================
with st.sidebar:
    st.markdown("### Board")
    board = st.session_state.board_state
    edits = int(board.get('seq', 0)) if board else 0
    confirm_ok = True
    if board and edits > 0:
        confirm_ok = st.checkbox(f"Replace the current board ({edits} edit{'s' if edits != 1 else ''} will be lost)",
                                 key='confirm_regen')
    def _do_generate():
        st.session_state.board_state = board_from_days(generate_week())
        st.session_state.pop('confirm_regen', None)
        st.session_state.pop('restored_at', None)
        autosave()

    def _do_reload():
        if not restore_autosave(force=True):
            st.session_state.board_msg = "Nothing has been saved yet."

    def _do_clear():
        st.session_state.board_state = None
        st.session_state.pop('confirm_regen', None)
        autosave()

    st.button("Generate week" if not board else "Generate a new week", type="primary",
              disabled=not confirm_ok, on_click=_do_generate, **STRETCH)
    c1, c2 = st.columns(2)
    with c1:
        st.button("Reload saved", on_click=_do_reload, **STRETCH,
                  help="Pull the last autosaved board (e.g. after another lead edited it)")
    with c2:
        st.button("Clear board", on_click=_do_clear, disabled=board is None, **STRETCH)
    msg = st.session_state.pop('board_msg', None)
    if msg:
        st.info(msg)

    with st.expander("Display"):
        st.session_state.board_height = st.slider("Board height", 520, 1400, st.session_state.board_height, 20)
        st.markdown('<span class="note">Tip: use "Show full board" on the board to drop the inner scrollbar entirely.</span>',
                    unsafe_allow_html=True)


# =====================================================================
# Main
# =====================================================================
board = st.session_state.board_state
logo_b64 = get_base64_logo(BASE_DIR / "natera.png")
logo_html = f'<img src="data:image/png;base64,{logo_b64}" alt="">' if logo_b64 else ''
n_pto = sum(1 for v in st.session_state.pto_by_day.values() if v)
n_ot = sum(1 for v in st.session_state.ot_by_day.values() if v)
sub = (f"Generated {board.get('generated_at', '')}" if board else "No week generated yet")
if board and board.get('seq'):
    sub += f". {board['seq']} edit{'s' if board['seq'] != 1 else ''} since."
if st.session_state.get('restored_at'):
    sub += f" Restored from autosave at {st.session_state.restored_at[11:16]}."
st.markdown(f"""
<div class="hero">{logo_html}<div><div class="t">3rd shift schedule board</div><div class="s">{sub}</div></div>
<div class="stats">
  <div class="stat"><b>{len(df)}</b><span>on roster</span></div>
  <div class="stat"><b>{n_pto}</b><span>with PTO</span></div>
  <div class="stat"><b>{n_ot}</b><span>on overtime</span></div>
  <div class="stat"><b>{len(st.session_state.training_pairs)}</b><span>training pairs</span></div>
</div></div>
""", unsafe_allow_html=True)

if board is None:
    st.info("Set PTO, overtime and training pairs in the sidebar, then click **Generate week**. "
            "After that, everything on the board is drag-and-drop: move people between roles and days, "
            "call someone out, add someone who picked up a shift, and fill gaps from a qualified-first picker.")
    st.stop()

comp_dir = _component_dir()
if comp_dir is None:
    st.error("Couldn't write the board component files (read-only file system). The static table below still works.")
else:
    _board_component = components.declare_component("schedule_board", path=str(comp_dir))
    ret = _board_component(
        config_json=json.dumps(board_config(df)),
        state_json=json.dumps(board),
        height=int(st.session_state.board_height),
        key='board_widget',
        default=None,
    )
    if isinstance(ret, dict) and ret.get('rev') == board.get('rev') and int(ret.get('seq', 0)) > int(board.get('seq', 0)):
        st.session_state.board_state = ret
        board = ret
        autosave()

schedule_df = board_to_frame(board)

c1, c2, c3 = st.columns([1, 1, 4])
with c1:
    st.download_button("Download CSV", schedule_df.to_csv(index=False).encode('utf-8'),
                       "weekly_schedule.csv", "text/csv", **STRETCH)
with c2:
    st.download_button("Download board (JSON)", json.dumps(board, indent=1).encode('utf-8'),
                       "schedule_board.json", "application/json", **STRETCH,
                       help="A full snapshot of the board, including the change log")

with st.expander("Flat per-day table (for spot-checking)"):
    selected_day = st.radio("Day", DAYS, horizontal=True, key="flat_day_radio")
    view = schedule_df[schedule_df['Day'] == selected_day].copy()
    view['__block'] = view['Workflow'].map(block_rank)
    view['__zone'] = view['Workflow'].map(zone_rank)
    view.sort_values(['__block', '__zone', '__roster_index'], inplace=True, kind='stable')
    view = view.drop(columns=['__block', '__zone', '__roster_index']).reset_index(drop=True)
    working = int((~view['Status'].str.startswith('Out')).sum())
    st.caption(f"{working} working on {selected_day}, {len(view) - working} out")
    st.dataframe(view[['Name', 'Workflow', 'Status']], hide_index=True)

with st.expander("Change log"):
    log = board.get('log') or []
    if not log:
        st.markdown('<span class="note">No edits yet. Every move, call-out and add made on the board shows up here.</span>',
                    unsafe_allow_html=True)
    else:
        st.dataframe(pd.DataFrame(log).rename(columns={'t': 'Time', 'd': 'Day', 'm': 'Change'}),
                     hide_index=True, height=min(420, 38 + 35 * len(log)))
