import re
from collections import OrderedDict

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
.sched-wrap { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
  background: #0e1117; color: #e6e6e6; border-radius: 10px; overflow: hidden;
  border: 1px solid #232936; }
.sched-table { width: 100%; border-collapse: collapse; font-size: 13px; table-layout: fixed; }
.sched-table col.role-col { width: 15%; }
.sched-table thead th { position: sticky; top: 0; z-index: 3; background: #161b22;
  color: #8b949e; font-weight: 600; font-size: 11px; letter-spacing: .04em;
  text-transform: uppercase; padding: 10px 8px; border-bottom: 1px solid #232936;
  text-align: center; }
.sched-table thead th.role-head { text-align: left; left: 0; z-index: 4; }
.sched-table thead .headcount { display: block; margin-top: 2px; font-size: 15px;
  font-weight: 700; color: #e6e6e6; letter-spacing: 0; text-transform: none; }
.sched-table tbody td, .sched-table tbody th { padding: 7px 8px; border-bottom: 1px solid #1c212b;
  vertical-align: middle; }
.sched-table tbody tr:hover td, .sched-table tbody tr:hover th { background: #171c26; }
.cat-row td { background: var(--accent-bg); padding: 6px 10px; font-size: 11px;
  font-weight: 700; letter-spacing: .06em; color: var(--accent); text-transform: uppercase;
  border-top: 1px solid #232936; border-bottom: 1px solid #232936; }
.role-cell { position: sticky; left: 0; background: #12161f; z-index: 2;
  border-left: 3px solid var(--accent); font-weight: 500; color: #c9d1d9;
  white-space: nowrap; }
.cell-name { color: #e6e6e6; }
.cell-empty { color: #4b5361; text-align: center; }
.tag { display: inline-block; margin-left: 4px; padding: 1px 5px; border-radius: 3px;
  font-size: 10px; font-weight: 700; background: #232936; color: #8b949e; }
.name-sep { color: #4b5361; }
.legend { display: flex; flex-wrap: wrap; gap: 14px; padding: 10px 14px; background: #161b22;
  border-top: 1px solid #232936; font-size: 11px; color: #8b949e; }
.legend .dot { display: inline-block; width: 8px; height: 8px; border-radius: 2px;
  margin-right: 5px; vertical-align: middle; }
"""


def render_week_grid_html(grid, day_headcount=None):
    """Returns a <div class='sched-wrap'>...</div> fragment (CSS included via <style>)."""
    day_headcount = day_headcount or {}
    parts = [f"<style>{CSS}</style>", "<div class='sched-wrap'><table class='sched-table'>"]
    parts.append("<colgroup><col class='role-col'>" + "<col>" * 7 + "</colgroup>")
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
            f"<tr class='cat-row' style='--accent:{accent}; --accent-bg:{accent}1a'>"
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

    parts.append("</tbody></table>")
    parts.append(
        "<div class='legend'>" +
        "".join(
            f"<span><span class='dot' style='background:{c}'></span>{cat}</span>"
            for cat, c in CATEGORY_COLORS.items() if cat != 'TRAINING'
        ) +
        "</div></div>"
    )
    return "".join(parts)
