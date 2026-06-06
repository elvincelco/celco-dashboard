#!/usr/bin/env python3
"""Offline replacement for celco_wip_monitor.html drop-zone + Publish to Team flow.

Reads the Working GP XLS, replicates parseData() in Python, extracts the
WIP_TEMPLATE JS string from celco_wip_monitor.html, substitutes the data and
meta markers, and writes wip_published.html.

KEEP IN SYNC with celco_wip_monitor.html — specifically the canonical CELCO_DATA
block (suffix-first resolution: CEILINGS, CODE_TO_DIV, EXCLUDED_CODES,
JOB_OVERRIDES, NAME_TO_CODE, CODE_TO_SUBDIV_NAME, resolveJob/resolveWip),
budgetTier, and the parseData job-dict shape. The HTML is the source of truth;
this script is a static Python copy.

2026-06 re-sync: replaced the old string-based classifyDivision/SUBDIV_CEILINGS
with the suffix-first CELCO_DATA resolver to match the rewritten template
(FIRE/FIRESTR 75%, CONSTR 45%, CST 50%, ELE folded into MAINTENANCE, etc.).
"""
import json, math, re, sys, os
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import pandas as pd

# Render timestamps in Pacific time regardless of where this runs (the scheduled
# task sandbox is UTC; local manual runs are typically Pacific).
PACIFIC = ZoneInfo('America/Los_Angeles')

XLS_PATH = sys.argv[1]
MONITOR_PATH = sys.argv[2]
OUT_PATH = sys.argv[3]

# ═════════════════════════════════════════════════════════════════════════════
# CANONICAL RESOLUTION (suffix-first) — ported verbatim from the CELCO_DATA block
# in celco_wip_monitor.html. Subdivision, parent division, and ceiling resolve
# from the JOB NUMBER suffix code first; the DASH division string is only a
# fallback. KEEP IN SYNC with the HTML template (source of truth).
# ═════════════════════════════════════════════════════════════════════════════

# Ceiling = max cost as a % of sold amount. GP floor = 1 − ceiling. Keyed by suffix.
CEILINGS = {
    # STRUCTURE
    'FIRE': 0.75, 'FIRESTR': 0.75,
    'STR': 0.55, 'US': 0.55, 'SUPP': 0.55,
    'CST': 0.50, 'CONSTR': 0.45,
    # EMERGENCY
    'WTR': 0.25, 'MLD': 0.35,
    'RT': 0.40, 'TARP': 0.40, 'BU': 0.40, 'TMP': 0.40,
    'CONWTR': 0.45, 'CLN': 0.50, 'ABAT': 0.60, 'TEST': 0.80,
    # MAINTENANCE
    'MNT': 0.55, 'MWO': 0.60, 'MIPMP': 0.70,
    'MI': 0.50, 'MIL': 0.50, 'REM': 0.70,
    'ELE': 0.55,   # Electrical — rolled into Maintenance (was its own division)
    # PLUMBING
    'PLM': 0.55, 'RPLM': 0.55,
    # IRVINE
    'IRVWTR': 0.25, 'IRVMLD': 0.25, 'IRVSTR': 0.65, 'IRVPLM': 0.65, 'IRVELE': 0.65,
    # MEDICAL
    'MEDWTR': 0.50, 'MEDSTR': 0.50,
    # LEGACY / retired-but-mapped
    'RFG': 0.55, 'HOA': 0.55, 'INSP': 0.50, 'COMP-BID': 0.50,
}

CODE_TO_DIV = {
    'FIRE': 'STRUCTURE', 'FIRESTR': 'STRUCTURE', 'STR': 'STRUCTURE', 'US': 'STRUCTURE',
    'SUPP': 'STRUCTURE', 'CST': 'STRUCTURE', 'CONSTR': 'STRUCTURE',
    'RFG': 'STRUCTURE', 'HOA': 'STRUCTURE', 'INSP': 'STRUCTURE', 'COMP-BID': 'STRUCTURE',
    'WTR': 'EMERGENCY', 'MLD': 'EMERGENCY', 'RT': 'EMERGENCY', 'TARP': 'EMERGENCY',
    'BU': 'EMERGENCY', 'TMP': 'EMERGENCY', 'CONWTR': 'EMERGENCY', 'CLN': 'EMERGENCY',
    'ABAT': 'EMERGENCY', 'TEST': 'EMERGENCY',
    'MNT': 'MAINTENANCE', 'MWO': 'MAINTENANCE', 'MIPMP': 'MAINTENANCE',
    'MI': 'MAINTENANCE', 'MIL': 'MAINTENANCE', 'REM': 'MAINTENANCE', 'ELE': 'MAINTENANCE',
    'PLM': 'PLUMBING', 'RPLM': 'PLUMBING',
    'IRVWTR': 'IRVINE', 'IRVMLD': 'IRVINE', 'IRVSTR': 'IRVINE', 'IRVPLM': 'IRVINE', 'IRVELE': 'IRVINE',
    'MEDWTR': 'MEDICAL', 'MEDSTR': 'MEDICAL',
    # NOTE: bare 'IRV' and bare 'CON' intentionally absent — handled as legacy in resolve_job.
}

EXCLUDED_CODES = {'WAR', 'EST', 'MAR', 'SHOP', 'OFFICE', 'INVT', 'WHS', 'LAB', 'VEH', 'NUVE', 'VAND'}

# Closed historical exceptions matched by stem (UPPERCASED job number startswith).
JOB_OVERRIDES = [
    {'stem': '01-09-017125IRV',  'exclude': True,  'reason': 'Irvine operational overhead (permanent)'},
    {'stem': '001-09-017125IRV', 'exclude': True,  'reason': 'Irvine operational overhead (permanent)'},
    {'stem': '01-09-011739-IRV', 'exclude': True,  'reason': 'Vehicle/operational overhead'},
    {'stem': '01-09-014143CON',  'code': 'CONSTR', 'reason': 'legacy CON -> Structure'},
    {'stem': '26-0165-CON',      'code': 'CONSTR', 'reason': 'legacy CON -> Structure'},
    {'stem': '01-09-016305CON',  'code': 'CONWTR', 'reason': 'legacy CON -> Emergency'},
    {'stem': '01-09-016666CON',  'code': 'CONWTR', 'reason': 'legacy CON -> Emergency'},
    {'stem': '01-09-016881CON',  'code': 'CONWTR', 'reason': 'legacy CON -> Emergency'},
    {'stem': '01-09-016905CON',  'code': 'CONWTR', 'reason': 'legacy CON -> Emergency'},
    {'stem': '01-09-017266CON',  'code': 'CONWTR', 'reason': 'legacy CON -> Emergency'},
    {'stem': '01-09-014699REM',  'exclude': True,  'reason': 'Remodel dead job'},
    {'stem': '01-09-016947REM',  'exclude': True,  'reason': 'Remodel dead job'},
    {'stem': '01-09-000-3218',   'exclude': True,  'reason': 'miscreated job'},
    {'stem': '01-09-000-35',     'exclude': True,  'reason': 'miscreated job'},
    {'stem': '01-09-003109',     'exclude': True,  'reason': 'miscreated job'},
]

# DASH "Division" string (case-insensitive) -> canonical suffix code (fallback only).
NAME_TO_CODE = {
    'structure': 'STR', 'upsales': 'US', 'supplement': 'SUPP', 'consulting': 'CST',
    'fire restoration': 'FIRESTR', 'fire mitigation': 'FIRE', 'fire damage': 'FIRE',
    'contents str': 'CONSTR', 'contents wtr': 'CONWTR', 'contents': 'CONSTR',
    'water mitigation': 'WTR', 'mold remediation': 'MLD', 'cleaning': 'CLN',
    'abatement': 'ABAT', 'roof tarp': 'RT', 'temp services': 'TMP',
    'maintenance': 'MNT', 'maint work order': 'MWO', 'maint insp pmp': 'MIPMP',
    'maintenance inspection leads': 'MIL', 'remodel': 'REM',
    'plumbing': 'PLM', 'residential plumbing': 'RPLM', 'electrical': 'ELE',
    'irvine structure': 'IRVSTR', 'irvine water mitigation': 'IRVWTR',
    'irvine mold remediation': 'IRVMLD', 'irvine plumbing': 'IRVPLM', 'irvine electrical': 'IRVELE',
    'medical facilities str': 'MEDSTR', 'medical facilities wtr': 'MEDWTR',
    'testing services': 'TEST',
}

NAME_INTERNAL = {'warranty', 'inventory', 'labor', 'vehicles', 'warehouse',
                 'nuve home thermostat', 'vandalism'}

CODE_TO_SUBDIV_NAME = {
    'FIRE': 'Fire / Fire Restoration', 'FIRESTR': 'Fire / Fire Restoration',
    'STR': 'Structure', 'US': 'Upsales', 'SUPP': 'Supplement', 'CST': 'Consulting',
    'CONSTR': 'Contents (Structure)',
    'RFG': 'Roofing', 'HOA': 'HOA', 'INSP': 'Inspection', 'COMP-BID': 'Comparative Bid',
    'CON': 'Contents (legacy)',
    'WTR': 'Water Mitigation', 'MLD': 'Mold Remediation',
    'RT': 'Tarp / Board-up', 'TARP': 'Tarp / Board-up', 'BU': 'Tarp / Board-up', 'TMP': 'Temp',
    'CONWTR': 'Contents (Water Mitigation)', 'CLN': 'Cleaning', 'ABAT': 'Abatement',
    'TEST': 'Testing Services',
    'MNT': 'Maintenance', 'MWO': 'Maint Work Order', 'MIPMP': 'Maint Insp PMP',
    'MI': 'MI', 'MIL': 'MIL', 'REM': 'Remodel', 'ELE': 'Electrical',
    'PLM': 'Plumbing', 'RPLM': 'Residential Plumbing',
    'IRVWTR': 'Irvine Water Mitigation', 'IRVMLD': 'Irvine Mold Remediation',
    'IRVSTR': 'Irvine Structure', 'IRVPLM': 'Irvine Plumbing', 'IRVELE': 'Irvine Electrical',
    'MEDWTR': 'Medical Water Mitigation', 'MEDSTR': 'Medical Structure',
}

_COMPOUND = ['COMP-BID']


def find_override(jobnum):
    if not jobnum:
        return None
    j = str(jobnum).upper().strip()
    for ov in JOB_OVERRIDES:
        if j.startswith(ov['stem'].upper()):
            return ov
    return None


def extract_job_code(jobnum):
    # Real formats: hyphenated (26-0165-STR, -STR-2), glued (010625ELE),
    # compound word (COMP-BID), oddball (26-0165-CON), slash miscode (.../MLD).
    if not jobnum:
        return None
    j = str(jobnum).upper().strip()
    if '/' in j:
        j = j.split('/')[0]                       # slash compound -> first part
    for comp in _COMPOUND:
        if j.endswith(comp):
            return comp
    j = re.sub(r'-\d+$', '', j)                    # strip trailing numeric revision
    m = re.search(r'(IRV[A-Z]+|MED[A-Z]+|CON[A-Z]+|[A-Z]+)$', j)
    return m.group(1) if m else None


def has_slash_compound(jobnum):
    return bool(jobnum) and '/' in str(jobnum)


def ceiling_for_code(code):
    if not code:
        return None
    return CEILINGS.get(str(code).upper().strip())


def division_for_code(code):
    if not code:
        return None
    if code in EXCLUDED_CODES:
        return '__EXCLUDED__'
    return CODE_TO_DIV.get(code, '__UNKNOWN__')


def code_for_name(name):
    if not name:
        return None
    key = str(name).lower().strip()
    if key in NAME_INTERNAL:
        return '__INTERNAL__'
    return NAME_TO_CODE.get(key)


def subdiv_name_for_code(code):
    if not code:
        return None
    return CODE_TO_SUBDIV_NAME.get(str(code).upper().strip())


def resolve_job(job_number, division_string):
    """Mirror of CELCO_DATA.resolveJob(). Returns dict with code/division/ceiling/
    source/excluded (+ mismatch flags, which don't affect baked output)."""
    out = {'code': None, 'division': None, 'ceiling': None, 'mismatch': False,
           'mismatchReason': '', 'source': 'suffix', 'excluded': False, 'excludeReason': ''}

    # 1) Job-number override (closed historical exceptions)
    ov = find_override(job_number)
    if ov:
        if ov.get('exclude'):
            out['excluded'] = True
            out['excludeReason'] = ov.get('reason', '')
            out['division'] = '__EXCLUDED__'
            out['source'] = 'override'
            return out
        out['code'] = ov['code']
        out['division'] = division_for_code(ov['code'])
        out['ceiling'] = ceiling_for_code(ov['code'])
        out['source'] = 'override'
        return out

    suf_code = extract_job_code(job_number)
    slash = has_slash_compound(job_number)
    name_code = code_for_name(division_string)

    # 2) Legacy bare IRV (predates IRVMLD)
    if suf_code == 'IRV':
        legacy = name_code if (name_code and name_code.startswith('IRV')) else 'IRVMLD'
        out['code'] = legacy
        out['division'] = 'IRVINE'
        out['ceiling'] = ceiling_for_code(legacy)
        out['source'] = 'legacy-irv'
        return out
    # Legacy bare CON not in overrides -> flag, don't guess.
    if suf_code == 'CON':
        out['code'] = 'CON'
        out['division'] = '__UNKNOWN__'
        out['ceiling'] = None
        out['source'] = 'legacy-con'
        out['mismatch'] = True
        out['mismatchReason'] = 'bare CON not in override list — needs classification'
        return out

    suf_div = division_for_code(suf_code)

    if suf_div == '__EXCLUDED__':
        out['code'] = suf_code
        out['division'] = '__EXCLUDED__'
        out['excluded'] = True
        out['excludeReason'] = 'excluded code ' + str(suf_code)
        out['source'] = 'suffix'
        return out
    if suf_code and suf_div != '__UNKNOWN__':
        out['code'] = suf_code
        out['division'] = suf_div
        out['ceiling'] = ceiling_for_code(suf_code)
        out['source'] = 'suffix'
    elif name_code and name_code != '__INTERNAL__':
        out['code'] = name_code
        out['division'] = division_for_code(name_code)
        out['ceiling'] = ceiling_for_code(name_code)
        out['source'] = 'name'
    elif name_code == '__INTERNAL__':
        out['division'] = '__EXCLUDED__'
        out['excluded'] = True
        out['excludeReason'] = 'internal DASH division'
        out['source'] = 'name'
        return out
    else:
        out['code'] = suf_code or None
        out['division'] = '__UNKNOWN__'
        out['source'] = 'none'
        return out

    # 3) Cross-check suffix vs DASH string (mismatch flags only; no effect on output)
    if slash:
        out['mismatch'] = True
        out['mismatchReason'] = 'slash-compound job number — graded by ' + str(out['code'])
    elif out['source'] == 'suffix' and name_code and name_code != '__INTERNAL__' and name_code != suf_code:
        n_div = division_for_code(name_code)
        n_ceil = ceiling_for_code(name_code)
        dash_is_bare_parent = (CODE_TO_DIV.get(name_code) == out['division']) and \
                              (str(division_string).lower().strip() == str(out['division']).lower())
        if (not dash_is_bare_parent) and (n_div != out['division'] or n_ceil != out['ceiling']):
            out['mismatch'] = True
            out['mismatchReason'] = ('suffix=' + str(suf_code) + ' vs DASH="' +
                                     str(division_string) + '"')
    return out


def resolve_wip(job_num, raw_div):
    """Mirror of resolveWip() wrapper in celco_wip_monitor.html."""
    r = resolve_job(job_num or '', raw_div or '')
    excluded = bool(r['excluded']) or r['division'] == '__EXCLUDED__'
    unknown = (not excluded) and (r['division'] == '__UNKNOWN__' or not r['division'])
    name = subdiv_name_for_code(r['code']) or (raw_div or '')
    return {
        'code': r['code'] or None,
        'division': '__EXCLUDED__' if excluded else ('__UNKNOWN__' if unknown else r['division']),
        'ceiling': r['ceiling'],
        'subdivName': name,
        'excluded': excluded,
        'unknown': unknown,
    }


# ── Replicate budgetTier() ──
def budget_tier(cost_pct, ceiling_pct):
    if ceiling_pct is None:
        return 'na'
    if cost_pct is None or not math.isfinite(cost_pct):
        return 'na'
    if cost_pct > ceiling_pct:
        return 'over'
    if cost_pct >= ceiling_pct * 0.90:
        return 'warning'
    return 'on'


# ── Helpers ──
def s(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ''
    return str(v).strip()


def num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, float) and math.isnan(v):
        return 0.0
    try:
        f = float(v)
        if math.isnan(f):
            return 0.0
        return f
    except (TypeError, ValueError):
        try:
            f = float(str(v).replace('$', '').replace(',', '').replace('%', '').strip())
            return 0.0 if math.isnan(f) else f
        except Exception:
            return 0.0


# Mirror parseExcelDate() — accept pandas Timestamp, Excel serial, or string.
def parse_excel_date(v):
    if v is None or v == '':
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        ts = v.to_pydatetime()
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=timezone.utc)
        return ts.astimezone(timezone.utc)
    if isinstance(v, datetime):
        if v.tzinfo is None:
            v = v.replace(tzinfo=timezone.utc)
        return v.astimezone(timezone.utc)
    try:
        n = float(v)
        if 40000 < n < 60000:
            from datetime import timedelta
            base = datetime(1899, 12, 30, tzinfo=timezone.utc)
            return base + timedelta(days=n)
    except (TypeError, ValueError):
        pass
    try:
        d = pd.to_datetime(str(v), errors='coerce')
        if pd.isna(d):
            return None
        if d.tzinfo is None:
            d = d.tz_localize('UTC')
        return d.to_pydatetime().astimezone(timezone.utc)
    except Exception:
        return None


# ── Parse XLS ──
df = pd.read_excel(XLS_PATH)

expected = ['Job Number', 'Job Name', 'Division', 'Foreperson', 'Coordinator',
            'Total Estimates', 'Total Job Cost', 'Working Gross Profit(%)']
missing = [c for c in expected if c not in df.columns]
if missing:
    print(f'FATAL: missing columns: {missing}', file=sys.stderr)
    print(f'Got: {list(df.columns)}', file=sys.stderr)
    sys.exit(2)

has_status_col = 'Job Status' in df.columns
has_date_started = 'Date Started' in df.columns

today = datetime.now(timezone.utc)

all_jobs = []
unknown_divs = {}
skipped_non_wip = 0
missing_date_started = 0

for _, r in df.iterrows():
    job_num = s(r.get('Job Number'))
    if not job_num:
        continue
    if has_status_col and s(r.get('Job Status')) != 'Work in Progress':
        skipped_non_wip += 1
        continue
    raw_div = s(r.get('Division'))

    _res = resolve_wip(job_num, raw_div)
    if _res['excluded']:
        continue                                   # internal/administrative — drop silently
    div_key = '__UNKNOWN__' if _res['unknown'] else _res['division']
    if _res['unknown']:
        key = raw_div or '(blank)'
        unknown_divs[key] = unknown_divs.get(key, 0) + 1
    subdiv_name = _res['subdivName'] or raw_div

    est = num(r.get('Total Estimates'))
    cost = num(r.get('Total Job Cost'))
    foreman = re.sub(r'\s+', ' ', s(r.get('Foreperson')))

    ceiling = _res['ceiling']
    cost_pct = (cost / est) if est > 0 else None
    ceiling_dollars = (est * ceiling) if (ceiling is not None and est > 0) else None
    variance_dollars = (ceiling_dollars - cost) if (ceiling_dollars is not None) else None
    tier = budget_tier(cost_pct, ceiling)

    # Age in WIP
    date_started_dt = parse_excel_date(r.get('Date Started')) if has_date_started else None
    if has_date_started and date_started_dt is None:
        missing_date_started += 1
    if date_started_dt is not None:
        delta_days = (today - date_started_dt).total_seconds() / 86400.0
        days_in_wip = max(0, int(delta_days))
        utc = date_started_dt.astimezone(timezone.utc)
        date_started_iso = utc.strftime('%Y-%m-%dT%H:%M:%S.') + f'{utc.microsecond // 1000:03d}Z'
    else:
        days_in_wip = None
        date_started_iso = None

    all_jobs.append({
        'jobNum':    job_num,
        'jobName':   s(r.get('Job Name')),
        'rawDiv':    raw_div,
        'subdivName': subdiv_name,
        'divKey':    div_key,
        'foreman':   foreman,
        'coord':     s(r.get('Coordinator')),
        'estimates': est,
        'jobCost':   cost,
        'gpDollar':  est - cost,
        'gpPct':     num(r.get('Working Gross Profit(%)')),
        'ceiling':   ceiling,
        'costPct':   cost_pct,
        'ceilingDollars':  ceiling_dollars,
        'varianceDollars': variance_dollars,
        'tier':      tier,
        'daysInWip': days_in_wip,
        'dateStartedISO': date_started_iso,
    })

print(f'Parsed {len(all_jobs)} jobs')
if not has_status_col:
    print('  WARN: Job Status column missing — WIP/Pending Sales filter skipped')
elif skipped_non_wip:
    print(f'  Skipped {skipped_non_wip} non-WIP rows (Pending Sales etc.)')
if not has_date_started:
    print('  WARN: Date Started column missing — all daysInWip will render blank')
elif missing_date_started:
    print(f'  WARN: {missing_date_started} WIP rows had no Date Started; ages render blank')
if unknown_divs:
    print(f'  Unknown divisions: {unknown_divs}')

# ── Build meta ──
now = datetime.now(PACIFIC)
h = now.hour
ampm = 'PM' if h >= 12 else 'AM'
h12 = (h % 12) or 12
meta = {
    'updated': f'{now.month}/{now.day}/{now.year} {h12}:{now.minute:02d} {ampm}',
    'jobCount': len(all_jobs),
}

# ── Extract WIP_TEMPLATE ──
with open(MONITOR_PATH, 'r', encoding='utf-8') as f:
    src = f.read()

m = re.search(r'var\s+WIP_TEMPLATE\s*=\s*("(?:\\.|[^"\\])*")\s*;', src, re.DOTALL)
if not m:
    print('FATAL: could not extract WIP_TEMPLATE', file=sys.stderr)
    sys.exit(3)

template = json.loads(m.group(1))

# ── Substitute markers ──
data_json = json.dumps(all_jobs, separators=(',', ':'), ensure_ascii=False, allow_nan=False)
meta_json = json.dumps(meta, separators=(',', ':'), ensure_ascii=False, allow_nan=False)

if 'XWIP_DATA_MARKERX' not in template:
    print('FATAL: XWIP_DATA_MARKERX not found in template', file=sys.stderr)
    sys.exit(4)
if 'XWIP_META_MARKERX' not in template:
    print('FATAL: XWIP_META_MARKERX not found in template', file=sys.stderr)
    sys.exit(5)

output = template.replace('XWIP_DATA_MARKERX', data_json, 1).replace('XWIP_META_MARKERX', meta_json, 1)

with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write(output)

print(f'Wrote {OUT_PATH} ({len(output)} bytes)')
print(f'  starts: {output[:60]!r}')
print(f'  ends:   {output[-60:]!r}')
