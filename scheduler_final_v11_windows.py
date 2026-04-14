import streamlit as st
import pandas as pd
import random
import os
from collections import defaultdict
from PIL import Image
import base64
from io import BytesIO
from pathlib import Path
import re

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
    'poc': 'POC', 'cla': 'CLA',
}
REQUIRED_COLS = ['Name','Shift','ISO','TIU','QS','FLOAT','CLS','PGD','HZN','TIH','POC','CLA']

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

DATA_FILE = BASE_DIR / "today_active_workers_08APR2026.xlsx"

def load_data():
    if not os.path.exists(DATA_FILE):
        st.error(f"{DATA_FILE} not found.")
        return pd.DataFrame()
    try:
        if str(DATA_FILE).endswith('.xlsx'):
            df = pd.read_excel(DATA_FILE)
        else:
            try:
                df = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
            except UnicodeDecodeError:
                df = pd.read_csv(DATA_FILE, encoding="cp1252")
    except Exception as e:
        st.error(f"Could not load file: {e}")
        return pd.DataFrame()
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

def working_pool(day) -> pd.DataFrame:
    mask = df.apply(
        lambda r: (is_on_shift(r, day) or is_overtime(r['Name'], day)) and is_not_pto(r['Name'], day),
        axis=1
    )
    pool = df[mask].copy()
    if len(pool) == 0:
        pool = df[df['Name'].apply(lambda n: is_not_pto(n, day))].copy()
    return pool

CORE_ROLES = ['ISO','TIU','QS','FLOAT','CLS','PGD','HZN','TIH','POC','CLA']
HZN_DNA_ROLE = 'HZN EXT/NORM/DIL'

# 7 ISO zones (A-G), 13 QS zones (1-13)
ISO_ZONE_LIST = [f'Zone {c}' for c in ['A','B','C','D','E','F','G']]
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

def priority_names(df_pool, already_assigned, reserve_cls=False, limit=None, prefer_more_skills=False):
    if df_pool is None or df_pool.empty:
        return []
    pool = df_pool[~df_pool['Name'].isin(already_assigned)].copy()
    if pool.empty:
        return []

    rng_local = random.Random()
    pool = pool.assign(
        _yes_count=pool.apply(skills_count, axis=1),
        _rand=pool['Name'].apply(lambda _: rng_local.random())
    )

    def take(frame, n=None):
        # prefer_more_skills: versatile people go to ISO, leaving specialists free for other roles
        asc = not prefer_more_skills
        frame = frame.sort_values(['_yes_count', '_rand'], ascending=[asc, True])
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

def priority_names_excluding(df_pool, already_assigned, exclude_set=None, reserve_cls=False, limit=None, prefer_more_skills=False):
    exclude_set = exclude_set or set()
    if df_pool is None or df_pool.empty:
        return []
    base = df_pool[~df_pool['Name'].isin(exclude_set)]
    primary = priority_names(base, already_assigned, reserve_cls, limit, prefer_more_skills)
    if limit is None or len(primary) >= limit:
        return primary
    need = limit - len(primary)
    topup = df_pool[df_pool['Name'].isin(exclude_set)]
    return primary + priority_names(topup, already_assigned.union(set(primary)), reserve_cls, need, prefer_more_skills)

def block_rank(workflow: str) -> float:
    # Display order: ISO > QS > PGD > TIH > HZN > Floater > POC
    w = str(workflow)
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
    if w.startswith('QS Floater'):
        return 1000
    if w.startswith('Floater'):
        m = re.search(r'(\d+)$', w)
        return int(m.group(1)) if m else 1000
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

def enforce_tih_minimum(assign_map, day, pool, assigned):
    have_tih = any(
        d == day and any(r.startswith('TIH') for r in roles)
        for (d, _n), roles in assign_map.items()
    )
    if have_tih:
        return

    cls_pool = pick(pool, lambda r: r['CLS'].strip().lower() == 'yes' and r['Name'] not in assigned)
    tih_cls_pool = pick(cls_pool, lambda r: r['TIH'].strip().lower() == 'yes')
    names = priority_names(tih_cls_pool, assigned, reserve_cls=False, limit=1)
    if names:
        n = names[0]
        assign_map[(day, n)].append('TIH_CLS')
        assigned.add(n)
        return

    _steal_from_qs(
        assign_map,
        day,
        'TIH_CLS',
        predicate=lambda row: row['TIH'] == 'yes' and row['CLS'] == 'yes'
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

def backup_label_for_row(row):
    priority = [
        ('QS', 'QS Floater Backup'),
        ('FLOAT', 'Floater Backup'),
        ('ISO', 'ISO Backup'),
        ('HZN', 'HZN Backup'),
        ('TIH', 'TIH Backup'),
        ('PGD', 'PGD Backup'),
        ('TIU', 'TIU Backup'),
        ('POC', 'POC Backup'),
    ]
    for col, label in priority:
        if str(row.get(col, '')).strip().lower() == 'yes':
            return label
    return 'General Support'

def final_fill_no_unassigned(day, pool, assigned, assign_map):
    working_names = set(pool['Name'])
    leftovers = sorted(
        list(working_names - assigned),
        key=lambda n: int(df.loc[df['Name'] == n, '__roster_index'].iloc[0])
    )

    current_qs_floaters = count_roles_for_day(assign_map, day, 'QS Floater')
    current_floaters = count_roles_for_day(assign_map, day, 'Floater')

    qs_leftovers = []
    float_leftovers = []
    other_leftovers = []

    for name in leftovers:
        row = _df_row_by_name(name)
        qs_yes = str(row['QS']).strip().lower() == 'yes'
        float_yes = str(row['FLOAT']).strip().lower() == 'yes'

        if qs_yes:
            qs_leftovers.append(name)
        elif float_yes:
            float_leftovers.append(name)
        else:
            other_leftovers.append(name)

    for name in qs_leftovers:
        current_qs_floaters += 1
        assign_map[(day, name)].append(f'QS Floater {current_qs_floaters}')
        assigned.add(name)

    for name in float_leftovers:
        current_floaters += 1
        assign_map[(day, name)].append(f'Floater {current_floaters}')
        assigned.add(name)

    for name in other_leftovers:
        row = _df_row_by_name(name)
        assign_map[(day, name)].append(backup_label_for_row(row))
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
    weekly_iso_count = {}    # name -> how many times they've done ISO this week

    for day in days:
        pool = working_pool(day)
        working_names = set(pool['Name'])

        assigned = set()
        assign_map = defaultdict(list)

        # --- 0. Pre-reserve critical roles before ISO/Float/QS consumes everyone ---
        # Order matters: PGD first (smallest pool), then HZN EXT, then HZN POC Swap.
        # Each step excludes the previous to guarantee no double-assignments.
        pgd_reserved          = set()
        pgd_reserved_names    = []
        hzn_ext_reserved      = set()
        hzn_ext_reserved_names = []
        hzn_poc_reserved      = set()
        swap_reserved         = []

        if day not in ['Sun', 'Mon']:
            # 1. PGD (1 person) — picked FIRST, added to assigned immediately so nothing steals them
            pgd_pool_all = pool[pool['PGD'] == 'yes'].copy()
            # prefer someone not used yesterday, but fall back if needed
            pgd_candidates = [n for n in pgd_pool_all['Name'].tolist() if n not in prev_pgd]
            if not pgd_candidates:
                pgd_candidates = pgd_pool_all['Name'].tolist()
            pgd_pick = pgd_candidates[0] if pgd_candidates else None
            if pgd_pick:
                pgd_reserved_names = [pgd_pick]
                pgd_reserved = {pgd_pick}
                assigned.add(pgd_pick)  # lock in NOW

            # 2. HZN EXT/NORM/DIL (2-3 people, non-POC preferred, exclude PGD person)
            hzn_ext_reserve_pool = pick(
                pool,
                lambda r: r['HZN'].strip().lower() == 'yes'
                          and r['POC'].strip().lower() != 'yes'
                          and r['Name'] not in pgd_reserved
            )
            hzn_ext_reserved_names = priority_names_excluding(
                hzn_ext_reserve_pool, assigned, exclude_set=prev_hzn_ext, reserve_cls=True, limit=3
            )
            hzn_ext_reserved = set(hzn_ext_reserved_names)

            # 3. HZN POC Swap (4 people, exclude PGD and HZN EXT people)
            swap_reserve_pool = pick(
                pool,
                lambda r: r['POC'].strip().lower() == 'yes'
                          and r['HZN'].strip().lower() == 'yes'
                          and r['Name'] not in pgd_reserved
                          and r['Name'] not in hzn_ext_reserved
                          and r['Name'] not in weekly_poc_used
            )
            swap_reserved = priority_names_excluding(
                swap_reserve_pool, assigned, exclude_set=prev_hzn_poc, reserve_cls=True, limit=4
            )
            hzn_poc_reserved = set(swap_reserved)

        else:
            # Sun/Mon: reserve 2-3 HZN EXT people (no POC swap, no PGD)
            hzn_ext_reserve_pool = pick(
                pool,
                lambda r: r['HZN'].strip().lower() == 'yes' and r['POC'].strip().lower() != 'yes'
            )
            hzn_ext_reserved_names = priority_names_excluding(
                hzn_ext_reserve_pool, assigned, exclude_set=prev_hzn_ext, reserve_cls=True, limit=3
            )
            hzn_ext_reserved = set(hzn_ext_reserved_names)

        # --- 1. ISO + Floater paired assignment, then QS ---
        # reserved set: nobody in here can be touched by ISO/Float/QS
        reserved = hzn_poc_reserved | hzn_ext_reserved | pgd_reserved

        ISO_MIN = 4
        QS_MIN  = 4

        # Single-pass: assign ISO first, then Float from remaining, no double simulation
        # ISO frequency cap: >2 quals = max 1x/week, <=2 quals = max 2x/week
        def iso_freq_ok(r):
            n = r['Name']
            count = weekly_iso_count.get(n, 0)
            quals = sum(1 for c in CORE_ROLES if str(r.get(c, '')).strip().lower() == 'yes')
            max_times = 1 if quals > 2 else 2
            return count < max_times

        iso_pool_filtered = pick(pool, lambda r: r['ISO'].strip().lower() == 'yes'
                                  and r['Name'] not in reserved
                                  and iso_freq_ok(r))
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

        # Commit — same lists, no re-seeding, guaranteed consistent
        for i, name in enumerate(iso_all[:n_pairs]):
            safe_assign(assign_map, assigned, day, name, f'ISO {ISO_ZONE_LIST[i]}')

        for i, name in enumerate(float_all[:n_pairs]):
            safe_assign(assign_map, assigned, day, name, f'Floater {i+1}')

        # --- 2. QS (Zones 1-13, strictly sequential, no gaps) ---
        qs_pool  = pick(pool, lambda r: r['QS'].strip().lower() == 'yes'
                        and r['Name'] not in assigned and r['Name'] not in reserved)
        qs_names = priority_names(qs_pool, assigned, reserve_cls=True, limit=len(QS_ZONE_LIST))
        for i, name in enumerate(qs_names):
            safe_assign(assign_map, assigned, day, name, f'QS {QS_ZONE_LIST[i]}')

        # QS Floater
        qs_floater_pool  = pick(pool, lambda r: r['QS'].strip().lower() == 'yes'
                                 and r['Name'] not in assigned and r['Name'] not in reserved)
        qs_floater_names = priority_names(qs_floater_pool, assigned, reserve_cls=True, limit=QS_BASE_FLOATER_COUNT_PER_DAY)
        for idx, name in enumerate(qs_floater_names, start=1):
            safe_assign(assign_map, assigned, day, name, f'QS Floater {idx}')

        # --- 3. TIU ---
        tiu_count = 1 if day in ['Sun','Mon'] else 2
        cls_pool = pick(pool, lambda r: r['CLS'].strip().lower() == 'yes' and r['Name'] not in assigned)
        tiu_pool = pick(cls_pool, lambda r: r['TIU'].strip().lower() == 'yes')
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

        # --- 5. DNEasy (Tue-Sat) ---
        # Requires CLS qualification. If POC/HZN Swap is happening today,
        # must ensure the DNEasy person is CLS-qualified (they run Mix-1/DNEasy).
        if day not in ['Sun','Mon']:
            dne_pool = pick(
                pool,
                lambda r: r['CLS'].strip().lower() == 'yes'
                          and r['POC'].strip().lower() == 'yes'
                          and r['Name'] not in assigned
                          and r['Name'] not in weekly_poc_used
            )
            for name in priority_names_excluding(dne_pool, assigned, exclude_set=prev_dne, reserve_cls=False, limit=1):
                safe_assign(assign_map, assigned, day, name, 'DNEasy/Mix-1')

        # --- 6. TIH enforcement ---
        enforce_tih_minimum(assign_map, day, pool, assigned)

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

        today_rows = []
        for _, person in roster.iterrows():
            name = person['Name']
            if name not in working_names:
                continue

            roles = assign_map.get((day, name), []).copy()
            if not roles:
                row = _df_row_by_name(name)
                roles = [backup_label_for_row(row)]

            wf = " / ".join(roles) if len(roles) == 2 else ", ".join(roles)
            today_rows.append((day, int(person['__roster_index']), name, wf))

        if len(today_rows) != len(working_names):
            already = {n for _d, _idx, n, _w in today_rows}
            for n in (working_names - already):
                row = _df_row_by_name(n)
                today_rows.append((
                    day,
                    int(df.loc[df['Name'] == n, '__roster_index'].iloc[0]),
                    n,
                    backup_label_for_row(row)
                ))

        # Update weekly POC tracker
        weekly_poc_used.update(
            n for (dkey, n), roles in assign_map.items()
            if dkey == day and any('HZN POC Swap' in r or r == 'DNEasy/Mix-1' or r == 'DNEasy' for r in roles)
        )

        # Update weekly ISO count
        for (dkey, n), roles in assign_map.items():
            if dkey == day and any(r.startswith('ISO Zone') for r in roles):
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
        prev_hzn_ext = set(
            [n for (dkey, n), roles in assign_map.items() if dkey == day and any(r == 'HZN EXT/NORM/DIL' for r in roles)]
        )
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

st.subheader("Generated Weekly Schedule")
if 'full_schedule' in st.session_state:
    df_out = st.session_state['full_schedule']
    selected_day = st.radio("Select a day to view schedule:", days, horizontal=True)
    view = df_out[df_out['Day'] == selected_day].copy()

    view['__block'] = view['Workflow'].map(block_rank)
    view['__zone'] = view['Workflow'].map(zone_rank)
    view.sort_values(['__block','__zone','__roster_index'], inplace=True, kind='stable')
    view = view.drop(columns=['__block','__zone']).reset_index(drop=True)

    st.caption(f"Showing {len(view)} working staff for {selected_day}")
    st.dataframe(view[['Name','Workflow']], use_container_width=True)
    st.download_button(
        "Download Schedule CSV",
        st.session_state['full_schedule'].to_csv(index=False).encode(),
        "weekly_schedule.csv",
        "text/csv"
    )
else:
    st.info("Click 'Generate Weekly Schedule' to begin.")
