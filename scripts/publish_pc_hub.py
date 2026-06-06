#!/usr/bin/env python3
"""Offline replacement for celco_pc_hub.html drop-zone + Publish to Team flow.

Reads the 5 source reports (IMSR, Payment, AR, Working GP, Job Aging),
replicates the JS parsers in celco_pc_hub.html, builds the `baked` payload,
and writes celco_pc_hub_published.html with `<script>window.__BAKED_DATA__ = …</script>`
injected at the `<!-- BAKED_DATA -->` marker.

USAGE:
    publish_pc_hub.py <hub_html> <out_html> <imsr.xlsx> <payment.xlsx> \
                      <ar.xlsx> <wgp.xlsx> <aging.xls_or_html> [label]

Skip a report by passing an empty string ("") in its slot.

KEEP IN SYNC with celco_pc_hub.html — specifically parseAR, parsePayments,
parseIMSR, parseWGP, parseAging, the canonical CELCO_DATA block (suffix-first
resolution: CEILINGS, CODE_TO_DIV, EXCLUDED_CODES, JOB_OVERRIDES, NAME_TO_CODE,
CODE_TO_SUBDIV_NAME, resolveJob/resolvePc), budgetTier, excelDateToJS,
daysBetween, publishToTeam, rehydrateBakedData. The HTML is the source of truth;
this script is a static Python copy.

NOTE on division routing: only parseWGP bakes resolved fields (ceiling, subdivName,
tier). The AR/Payments/IMSR/Aging parsers carry raw DASH strings; the PC Hub's
operational routing (inDiv / MARIA_CODES, the ELE->Maintenance and CST/INSP/CLN/MED
-> Maria overrides) all runs at RENDER time inside the published template, so it is
deliberately NOT mirrored here.

2026-06 re-sync: replaced the old string-based SUBDIV_CEILINGS with the suffix-first
CELCO_DATA resolver to match the rewritten template (FIRE/FIRESTR 75%, CONSTR 45%,
CST 50%, ELE folded into MAINTENANCE, etc.), and added subdivName to the WGP rows.

This script exists because `mcp__Claude_in_Chrome__file_upload` cannot read
sandbox-created files in scheduled-task runs. Same workaround pattern as
publish_wip.py for celco_wip_monitor.html.
"""
import json, math, re, sys, os
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import pandas as pd
from bs4 import BeautifulSoup

PACIFIC = ZoneInfo('America/Los_Angeles')

# ═════════════════════════════════════════════════════════════════════════════
# CANONICAL RESOLUTION (suffix-first) — ported verbatim from the CELCO_DATA block
# in celco_pc_hub.html. Subdivision, parent division, and ceiling resolve from the
# JOB NUMBER suffix code first; the DASH division string is only a fallback.
# This block is byte-identical to the one in publish_wip.py. KEEP IN SYNC with the
# HTML template (source of truth).
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
    if not jobnum:
        return None
    j = str(jobnum).upper().strip()
    if '/' in j:
        j = j.split('/')[0]
    for comp in _COMPOUND:
        if j.endswith(comp):
            return comp
    j = re.sub(r'-\d+$', '', j)
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
    """Mirror of CELCO_DATA.resolveJob()."""
    out = {'code': None, 'division': None, 'ceiling': None, 'mismatch': False,
           'mismatchReason': '', 'source': 'suffix', 'excluded': False, 'excludeReason': ''}

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

    if suf_code == 'IRV':
        legacy = name_code if (name_code and name_code.startswith('IRV')) else 'IRVMLD'
        out['code'] = legacy
        out['division'] = 'IRVINE'
        out['ceiling'] = ceiling_for_code(legacy)
        out['source'] = 'legacy-irv'
        return out
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


def resolve_pc(job_num, raw_div):
    """Mirror of resolvePc() wrapper in celco_pc_hub.html (identical to resolveWip)."""
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


def budget_tier(cost_pct, ceiling_pct):
    if ceiling_pct is None:
        return 'na'
    if cost_pct is None or (isinstance(cost_pct, float) and not math.isfinite(cost_pct)):
        return 'na'
    if cost_pct > ceiling_pct:
        return 'over'
    if cost_pct >= ceiling_pct * 0.90:
        return 'warning'
    return 'on'


# ── Excel date helpers — mirror excelDateToJS() and daysBetween() ──
def excel_date_to_dt(v):
    """Return a UTC-aware datetime, or None. Mirrors JS excelDateToJS()."""
    if v is None or v == '' or (isinstance(v, float) and math.isnan(v)):
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


def days_between(d1, d2):
    """floor((d2 - d1) / 1 day). Returns int or None."""
    if d1 is None or d2 is None:
        return None
    delta = (d2 - d1).total_seconds() / 86400.0
    return int(math.floor(delta))


def iso_z(dt):
    """JS Date#toISOString format: ms precision, trailing Z."""
    if dt is None:
        return None
    dt = dt.astimezone(timezone.utc)
    return dt.strftime('%Y-%m-%dT%H:%M:%S.') + f'{dt.microsecond // 1000:03d}Z'


def round2(x):
    """Match JS Math.round(x*100)/100 — round-half-away-from-zero on .5."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return 0.0
    sign = -1 if x < 0 else 1
    return sign * math.floor(abs(x) * 100 + 0.5) / 100


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
        return 0.0 if math.isnan(f) else f
    except (TypeError, ValueError):
        try:
            f = float(str(v).replace('$', '').replace(',', '').replace('%', '').strip())
            return 0.0 if math.isnan(f) else f
        except Exception:
            return 0.0


def num_or_none(v):
    """parseFloat() in JS returns NaN for non-numeric → we treat as None."""
    if v is None or v == '':
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    try:
        f = float(v)
        return None if math.isnan(f) else f
    except (TypeError, ValueError):
        try:
            f = float(str(v).replace('$', '').replace(',', '').replace('%', '').strip())
            return None if math.isnan(f) else f
        except Exception:
            return None


# ── XLSX → list of rows (mirrors XLSX.utils.sheet_to_json header:1) ──
def read_xlsx_rows(path):
    """Return [headers, row1_cells, row2_cells, ...] as a list of lists."""
    df = pd.read_excel(path, header=None, engine='openpyxl')
    rows = []
    for _, r in df.iterrows():
        rows.append([('' if pd.isna(c) else c) for c in r.tolist()])
    return rows


def idx_of(headers, label):
    for i, h in enumerate(headers):
        if str(h).strip() == label:
            return i
    return -1


# ── parseAR ──
def parse_ar(rows):
    headers = [s(h) for h in rows[0]]
    jnI  = idx_of(headers, 'Job Number')
    diI  = idx_of(headers, 'Date Invoiced')
    jnmI = idx_of(headers, 'Job Name')
    tiI  = idx_of(headers, 'Total Invoiced')
    tcI  = idx_of(headers, 'Total Collected')
    boI  = idx_of(headers, 'Balance Owing')
    jsI  = idx_of(headers, 'Job Status')
    coI  = idx_of(headers, 'Coordinator')
    today_utc = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
    items = []
    for r in rows[1:]:
        if not r or jnI < 0 or jnI >= len(r) or not r[jnI]:
            continue
        bal_raw = r[boI] if boI >= 0 and boI < len(r) else None
        bal = num_or_none(bal_raw)
        if bal is None:
            continue
        inv = num(r[tiI]) if tiI >= 0 and tiI < len(r) else 0.0
        col = num(r[tcI]) if tcI >= 0 and tcI < len(r) else 0.0
        date_obj = excel_date_to_dt(r[diI]) if diI >= 0 and diI < len(r) else None
        days_aged = days_between(date_obj, today_utc) if date_obj else None
        items.append({
            'job':         s(r[jnI]),
            'name':        s(r[jnmI]) if jnmI >= 0 and jnmI < len(r) else '',
            'inv':         round2(inv),
            'collected':   round2(col),
            'bal':         round2(bal),
            'status':      s(r[jsI]) if jsI >= 0 and jsI < len(r) else '',
            'coordinator': s(r[coI]) if coI >= 0 and coI < len(r) else '',
            'dateObj':     iso_z(date_obj),
            'daysAged':    days_aged,
        })
    return items


# ── parsePayments ──
def parse_payments(rows):
    headers = [s(h) for h in rows[0]]
    dI   = idx_of(headers, 'Division')
    eI   = idx_of(headers, 'Estimator')
    jnI  = idx_of(headers, 'Job Number')
    aI   = idx_of(headers, 'Payment Amount')
    dtI  = idx_of(headers, 'Payment Date')
    coI  = idx_of(headers, 'Coordinator')
    items = []
    for r in rows[1:]:
        if not r:
            continue
        amt = num_or_none(r[aI]) if aI >= 0 and aI < len(r) else None
        if amt is None or amt <= 0:
            continue
        date_obj = excel_date_to_dt(r[dtI]) if dtI >= 0 and dtI < len(r) else None
        items.append({
            'job':         s(r[jnI]) if jnI >= 0 and jnI < len(r) else '',
            'div':         s(r[dI])  if dI  >= 0 and dI  < len(r) else '',
            'est':         re.sub(r'\s+', ' ', s(r[eI])) if eI >= 0 and eI < len(r) else '',
            'amt':         round2(amt),
            'coordinator': s(r[coI]) if coI >= 0 and coI < len(r) else '',
            'dateObj':     iso_z(date_obj),
        })
    return items


# ── parseIMSR ──
def parse_imsr(rows):
    headers = [s(h) for h in rows[0]]
    dI   = idx_of(headers, 'Division')
    eI   = idx_of(headers, 'Estimator')
    cI   = idx_of(headers, 'Loss Category')
    amtI = idx_of(headers, 'Invoiced Subtotal')
    if amtI < 0:
        amtI = 16  # fallback to known position
    coI  = idx_of(headers, 'Coordinator')
    items = []
    for r in rows[1:]:
        if not r:
            continue
        amt = num_or_none(r[amtI]) if amtI >= 0 and amtI < len(r) else None
        if amt is None or amt == 0:
            continue
        if dI >= 0 and dI < len(r):
            if not r[dI]:
                continue
        date_obj = excel_date_to_dt(r[5]) if len(r) > 5 else None
        items.append({
            'job':         s(r[4]) if len(r) > 4 else '',
            'div':         s(r[dI]) if dI >= 0 and dI < len(r) else '',
            'est':         re.sub(r'\s+', ' ', s(r[eI])) if eI >= 0 and eI < len(r) else '',
            'cat':         s(r[cI]) if cI >= 0 and cI < len(r) else '',
            'amt':         round2(amt),
            'dateObj':     iso_z(date_obj),
            'coordinator': s(r[coI]) if coI >= 0 and coI < len(r) else '',
        })
    return items


# ── parseWGP ──
# Suffix-first resolution (ceiling, subdivName, tier) baked here to match the
# template's parseWGP, which resolves via resolvePc() at parse time.
def parse_wgp(rows):
    headers = [s(h) for h in rows[0]]
    jnI  = idx_of(headers, 'Job Number')
    jnmI = idx_of(headers, 'Job Name')
    dI   = idx_of(headers, 'Division')
    fI   = idx_of(headers, 'Foreperson')
    coI  = idx_of(headers, 'Coordinator')
    teI  = idx_of(headers, 'Total Estimates')
    tcI  = idx_of(headers, 'Total Job Cost')
    gpI  = idx_of(headers, 'Working Gross Profit(%)')
    jsI  = idx_of(headers, 'Job Status')
    dsI  = idx_of(headers, 'Date Started')
    today_utc = datetime.now(timezone.utc)
    items = []
    for r in rows[1:]:
        if not r or jnI < 0 or jnI >= len(r) or not r[jnI]:
            continue
        job_num = s(r[jnI])
        est  = num(r[teI]) if teI >= 0 and teI < len(r) else 0.0
        cost = num(r[tcI]) if tcI >= 0 and tcI < len(r) else 0.0
        raw_div = s(r[dI]) if dI >= 0 and dI < len(r) else ''
        _res = resolve_pc(job_num, raw_div)
        ceiling = _res['ceiling']
        subdiv_name = _res['subdivName'] or raw_div
        cost_pct = (cost / est) if est > 0 else None
        tier = budget_tier(cost_pct, ceiling)
        date_started = excel_date_to_dt(r[dsI]) if dsI >= 0 and dsI < len(r) else None
        days_in_wip = max(0, days_between(date_started, today_utc)) if date_started else None
        gp_pct = num(r[gpI]) if gpI >= 0 and gpI < len(r) else 0.0
        items.append({
            'job':             job_num,
            'name':            s(r[jnmI]) if jnmI >= 0 and jnmI < len(r) else '',
            'rawDiv':          raw_div,
            'subdivName':      subdiv_name,
            'foreman':         re.sub(r'\s+', ' ', s(r[fI])) if fI >= 0 and fI < len(r) else '',
            'coordinator':     s(r[coI]) if coI >= 0 and coI < len(r) else '',
            'est':             est,
            'cost':            cost,
            'gpPct':           gp_pct,
            'status':          s(r[jsI]) if jsI >= 0 and jsI < len(r) else '',
            'ceiling':         ceiling,
            'costPct':         cost_pct,
            'ceilingDollars':  (est * ceiling) if (ceiling is not None and est > 0) else None,
            'varianceDollars': (est * ceiling - cost) if (ceiling is not None and est > 0) else None,
            'tier':            tier,
            'daysInWip':       days_in_wip,
        })
    return items


# ── parseAging ──
def parse_aging(html_text):
    soup = BeautifulSoup(html_text, 'html.parser')
    table = soup.find('table')
    if not table:
        raise SystemExit('FATAL: Job Aging HTML has no <table>')
    trs = table.find_all('tr')
    if not trs:
        raise SystemExit('FATAL: Job Aging table is empty')
    headers = [c.get_text(strip=True) for c in trs[0].find_all(['th', 'td'])]
    status_cols = [
        'Pending Sales', 'Pre Production', 'Work in Progress',
        'Completed Without Paperwork', 'Invoice Pending',
        'Accounts Receivable', 'Waiting Final Closure',
    ]
    status_idx = {sc: headers.index(sc) for sc in status_cols if sc in headers}
    closed_idx = headers.index('Closed') if 'Closed' in headers else -1
    items = []
    skipped_closed = 0
    skipped_ar_ended = 0
    nbsp = '\xa0'
    re_days = re.compile(r'^(-?\d+)\s*Days?')
    for tr in trs[1:]:
        cells = [c.get_text(strip=True) for c in tr.find_all(['th', 'td'])]
        if len(cells) < 11:
            continue
        job_raw = cells[0]
        if not job_raw:
            continue
        comma = job_raw.find(',')
        jobnum = (job_raw[:comma] if comma >= 0 else job_raw).strip()
        job_name = job_raw[comma + 1:].strip() if comma >= 0 else ''
        if not jobnum:
            continue
        if closed_idx >= 0 and closed_idx < len(cells):
            cv = cells[closed_idx]
            if cv and cv.strip() and cv != nbsp:
                skipped_closed += 1
                continue
        ar_idx = status_idx.get('Accounts Receivable', -1)
        wfc_idx = status_idx.get('Waiting Final Closure', -1)
        ar_val = cells[ar_idx] if 0 <= ar_idx < len(cells) else ''
        wfc_val = cells[wfc_idx] if 0 <= wfc_idx < len(cells) else ''
        ar_ended = bool(ar_val and ar_val.strip() and ar_val != nbsp and 'to Present' not in ar_val)
        wfc_ended = bool(wfc_val and wfc_val.strip() and wfc_val != nbsp and 'to Present' not in wfc_val)
        if ar_ended or wfc_ended:
            skipped_ar_ended += 1
            continue
        current_status = None
        days_in_status = None
        for sc in reversed(status_cols):
            i = status_idx.get(sc, -1)
            if i < 0 or i >= len(cells):
                continue
            v = cells[i]
            if v and 'to Present' in v:
                current_status = sc
                m = re_days.match(v)
                if m:
                    days_in_status = int(m.group(1))
                break
        if not current_status:
            continue
        items.append({
            'job':           jobnum,
            'name':          job_name,
            'currentStatus': current_status,
            'daysInStatus':  days_in_status,
        })
    print(f'  Job Aging: parsed {len(items)} open jobs; skipped {skipped_closed} closed, {skipped_ar_ended} AR-ended')
    return items


# ── Aging input may be .xls (HTML-disguised) or .html ──
def read_aging_html(path):
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in ('utf-8', 'latin-1'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


# ── Build baked, inject, write ──
def main():
    if len(sys.argv) < 8:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    hub_path  = sys.argv[1]
    out_path  = sys.argv[2]
    imsr_path = sys.argv[3]
    pay_path  = sys.argv[4]
    ar_path   = sys.argv[5]
    wgp_path  = sys.argv[6]
    aging_path = sys.argv[7]
    label     = sys.argv[8] if len(sys.argv) > 8 else 'Refreshed by Cowork'

    loaded = {'ar': False, 'payments': False, 'imsr': False, 'aging': False, 'wgp': False}
    data   = {'ar': [], 'payments': [], 'imsr': [], 'aging': [], 'wgp': []}
    file_names = {}

    def display_label(orig_path):
        bn = os.path.basename(orig_path)
        bn = re.sub(r'^\d+_(IMSR|PAYMENT|AR|WORKING_GP|JOB_AGING)_', '', bn)
        return f'{label} — {bn}'

    if imsr_path:
        print(f'Reading IMSR: {imsr_path}')
        data['imsr'] = parse_imsr(read_xlsx_rows(imsr_path))
        loaded['imsr'] = True
        file_names['imsr'] = display_label(imsr_path)
        print(f'  IMSR: {len(data["imsr"])} invoices')

    if pay_path:
        print(f'Reading Payments: {pay_path}')
        data['payments'] = parse_payments(read_xlsx_rows(pay_path))
        loaded['payments'] = True
        file_names['payments'] = display_label(pay_path)
        print(f'  Payments: {len(data["payments"])} rows')

    if ar_path:
        print(f'Reading AR: {ar_path}')
        data['ar'] = parse_ar(read_xlsx_rows(ar_path))
        loaded['ar'] = True
        file_names['ar'] = display_label(ar_path)
        print(f'  AR: {len(data["ar"])} rows')

    if wgp_path:
        print(f'Reading Working GP: {wgp_path}')
        data['wgp'] = parse_wgp(read_xlsx_rows(wgp_path))
        loaded['wgp'] = True
        file_names['wgp'] = display_label(wgp_path)
        print(f'  WGP: {len(data["wgp"])} jobs')

    if aging_path:
        print(f'Reading Job Aging: {aging_path}')
        data['aging'] = parse_aging(read_aging_html(aging_path))
        loaded['aging'] = True
        file_names['aging'] = display_label(aging_path)

    baked = {
        'publishedAt': iso_z(datetime.now(timezone.utc)),
        'publishedBy': 'Cowork',
        'LOADED':      loaded,
        'DATA':        data,
        'fileNames':   file_names,
    }
    baked_json = json.dumps(baked, separators=(',', ':'), ensure_ascii=False, allow_nan=False)

    with open(hub_path, 'r', encoding='utf-8') as f:
        html = f.read()
    marker = '<!-- BAKED_DATA -->'
    if marker not in html:
        raise SystemExit(f'FATAL: marker "{marker}" not found in {hub_path}')
    close_tag_re = re.compile(r'<(/script)', re.IGNORECASE)
    safe_json = close_tag_re.sub(r'<\\\1', baked_json)
    inject = '<script>window.__BAKED_DATA__ = ' + safe_json + ';</script>'
    out_html = html.replace(marker, inject, 1)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(out_html)
    print(f'Wrote {out_path} ({len(out_html)} bytes)')
    print(f'  baked JSON: {len(baked_json)} chars')
    print(f'  loaded: {loaded}')


if __name__ == '__main__':
    main()
