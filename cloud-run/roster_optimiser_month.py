#!/usr/bin/env python3
"""
Apunga roster optimiser — month-horizon build (CP-SAT).  v4
New vs v3:
  - SITES.csv is now an input (Region / LGA / Ticks joined per visit).
  - Tick avoidance (HARD): staff with 'No ticks' never assigned to tick sites.
  - Preferred sites/regions (SOFT): tokens in 'Preferred Sites' and
    'Preferred regions' resolve against Region, LGA (with aliases) and
    site-name word-match; assignment to a preferred site earns R_PREF.
  - Preferred pairing (SOFT): 'Preferred Staff / Pairing' pairs earn R_PAIR
    when rostered onto the same visit (incl. when one is the supervisor).
  - Carpooling: passenger->driver map (optional 'Carpool with' Staff column,
    plus CARPOOL_SEED below). When passenger + driver share a visit (driver
    as crew or supervisor), the passenger's travel cost — in the objective
    AND the fairness budget — collapses to CARPOOL_MIN instead of their own
    matrix time. Incentive, not obligation: solo assignment stays possible.
"""
import pandas as pd, re, sys, math, os
from ortools.sat.python import cp_model

STAFF_CSV  = 'Staff.csv'
MATRIX_CSV = 'Travelmatrix.csv'
DEMAND_CSV = 'Demand.csv'
SITES_CSV  = 'Sites.csv'
OUT_CSV    = 'MasterRoster.csv'

W_TRANSIT_MULT, P_LONE_PROTECTED, P_UNQUAL_OVER, P_NO_ATM = 1.5, 10000, 10000, 30
P_UNDERSTAFF = 2000   # per missing crew member (below safety rules, above travel)
P_MINDAYS    = 1500   # per day short of a staff member's weekly MinDays
P_NO_UTE     = 5000   # 'Force Work Ute' visit with no work ute (supervisor or crew)
FAIRNESS_WEIGHT = 1
DAY_BUDGET_MIN  = 50           # travel budget granted per working day
FORTNIGHT_DAYS  = 14
NO_TRAVEL_LIMIT = 10**6
DAYS = ['Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat', 'Sun']
WEEKDAYS = ['Mon', 'Tue', 'Wed', 'Thur', 'Fri']
# Days with no availability column in the Staff tab (currently Sat/Sun) are
# treated as opt-in: nobody is auto-assigned, only staff PINNED on the Demand
# row can work them. Add 'Sat'/'Sun' columns to the Staff tab and they become
# ordinary rostered days automatically.

# --- new soft weights (placeholders — calibrate against real rosters) ----------
R_PREF      = 10   # reward (travel-min-equivalent) for a preferred site/region
R_PAIR      = 8    # reward for a preferred pairing on the same visit
R_SITE_ATM  = 300  # reward for the site's Preferred ATM being on that visit (grooming
                   # them to take the site over — deliberately heavy, beats travel)
P_NOT_PREF_ATM = 50  # penalty when the preferred ATM is present but not given the role
CARPOOL_MIN = 15   # effective one-way minutes for a passenger riding with driver

# passenger -> driver (StaffIDs). Overridden/extended by a 'Carpool with'
# column on the Staff tab if present.
CARPOOL_SEED = {'Mitch': 'Cam', 'Callum': 'Rosie'}

# fuzzy tokens in the preference columns -> canonical Region/LGA values
PREF_ALIASES = {
    'iwc': ['Inner West'], 'city': ['City of Sydney'], 'waverly': ['Waverley'],
    'eastern suburbs': ['East'], 'upper north shore': ['Hornsby', 'Ku-ring-gai'],
    'cumberland': ['Cumberland Council'], 'blacktown': ['Blacktown Council'],
}

def toks(v): return [] if pd.isna(v) else [t.strip() for t in re.split(r'[,\n/]', str(v)) if t.strip()]
def tru(v): return str(v).strip().lower() in ('true', 'yes', '1', 'x')

# ---- input plumbing --------------------------------------------------------------
# solve() takes a payload dict of {name: list-of-rows} (first row = header, exactly
# as Google Sheets returns it) OR {name: DataFrame}. Nothing here touches the disk,
# the network, or Google, so the same code runs locally and inside a container.

def _frame(payload, name, header='infer'):
    src = payload[name]
    if isinstance(src, pd.DataFrame): return src.copy()
    rows = list(src)
    if not rows: raise ValueError(f"'{name}' is empty")
    if header is None: return pd.DataFrame(rows)
    return pd.DataFrame(rows[1:], columns=[str(c).strip() for c in rows[0]])

def _fail(msg, warns=()):
    return {'ok': False, 'error': msg,
            'run_log': {'columns': ['Severity', 'Area', 'What happened', 'What to do'],
                        'rows': [['ACTION NEEDED', 'Run failed', msg,
                                  'Fix the problem above and run again.']]
                                + [['CHECK', 'Warning', w, ''] for w in warns]},
            'master_roster': {'columns': [], 'rows': []},
            'stats': {}, 'warnings': list(warns)}

def solve(payload, verbose=True):
    """payload -> {ok, master_roster, run_log, stats, warnings, console}.

    Never raises for ordinary data problems — returns ok=False with a readable
    message instead, so a caller (CLI or web service) can always show something.
    verbose=False captures the console output into result['console'] rather than
    printing it."""
    if verbose:
        return _solve(payload)
    import io, contextlib
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            res = _solve(payload)
    except Exception as e:
        return dict(_fail(f"Unexpected error: {type(e).__name__}: {e}"), console=buf.getvalue())
    res['console'] = buf.getvalue()
    return res

_ALT_RE = re.compile(r'alt', re.I)
def _is_alt(v):
    return v is not None and not (not isinstance(v, str) and pd.isna(v)) and bool(_ALT_RE.search(str(v)))

_ISO_RE = re.compile(r'^\s*(\d{4})-(\d{1,2})-(\d{1,2})')
def _to_date(v, dayfirst):
    """Sheet date -> Timestamp. ISO 'yyyy-mm-dd' (what the sheet button sends) is
    never ambiguous, so it bypasses the day/month guess; pandas would otherwise read
    '2026-11-02' as 11 Feb when dayfirst=True."""
    m = _ISO_RE.match(str(v))
    if m:
        try: return pd.Timestamp(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError: return pd.NaT
    return pd.to_datetime(v, dayfirst=dayfirst, errors='coerce')

def _solve(payload):
    # ---- staff -------------------------------------------------------------------
    raw = _frame(payload, 'staff', header=None)
    hr = next(i for i in range(15) if 'StaffID' in [str(x).strip() for x in raw.iloc[i]]
              and 'Transport' in [str(x).strip() for x in raw.iloc[i]])
    S = raw.iloc[hr+1:].copy(); S.columns = [str(x).strip() for x in raw.iloc[hr]]
    S = S[S['Staff'].notna() & (S['Staff'].astype(str).str.strip() != '')]
    S = S[~S['Staff'].astype(str).str.startswith(('Wk', 'TOTAL'))]
    DAY_COLS = [d for d in DAYS if d in S.columns]          # which days the sheet knows about
    OPTIN_DAYS = [d for d in DAYS if d not in DAY_COLS]     # e.g. Sat/Sun -> pinned staff only
    def cell(r, c): return r.get(c) if c in S.columns else None
    def celltxt(r, c):
        v = cell(r, c)
        return '' if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip()
    staff = {}
    for _, r in S.iterrows():
        sid = str(r['StaffID']).strip()
        if not sid: continue
        g = str(r.get('Gender', '')).strip().upper()
        cat, status = str(r.get('Category', '')).strip(), str(r.get('Cat. Status', '')).strip()
        mt = pd.to_numeric(r.get('Max travel time'), errors='coerce')
        quals = {q for q, c in [('whitecard', 'Whitecard'), ('irata', 'IRATA'), ('wwc', 'WWC'), ('chainsaw', 'Chainsaw')] if str(r.get(c, '')).strip()}
        staff[sid] = {
            'protected': (g in ('F', '')), 'male': (g == 'M'), 'atm': (cat == 'ATM'),
            'unq': (cat == 'Field Staff' and status in ('Unqualified', 'New')),
            'transit': 'public' in str(r.get('Transport', '')).lower(),
            'mt': (NO_TRAVEL_LIMIT if pd.isna(mt) else int(mt)),
            # "7.5" = available; "Alt" (or legacy "7.5Alt") = available on alternate
            # weeks only; anything else ("Office", "Nursery", blank) = not available
            'avail': {d: (d in DAY_COLS and (_is_alt(r.get(d)) or pd.notna(pd.to_numeric(
                          _ALT_RE.sub('', str(r.get(d, ''))).strip(), errors='coerce')))) for d in DAYS},
            'alt': {d for d in DAYS if d in DAY_COLS and _is_alt(r.get(d))},
            'ute': str(cell(r, 'Work Ute') or '').strip().lower() in ('yes', 'y', 'true', '1', 'x'),
            'excl': toks(r.get('Excluded sites')), 'nw': set(toks(r.get('Not work with Staff'))), 'quals': quals,
            'excl_reg_toks': toks(cell(r, 'Excluded regions')),
            'no_ticks': tru(cell(r, 'No ticks')),
        'adhoc': 'adhoc' in celltxt(r, 'Avail. Pattern').lower().replace(' ', ''),
        'away_raw': celltxt(r, 'Away'),
        'back_raw': celltxt(r, 'Returning'),
        'spray': celltxt(r, 'Ok to spray?').lower(),
            'pref_toks': toks(cell(r, 'Preferred Sites')) + toks(cell(r, 'Preferred regions')),
            'pair_toks': toks(cell(r, 'Preferred Staff / Pairing')),
            'carpool': (str(cell(r, 'Carpool with')).strip()
                        if 'Carpool with' in S.columns and pd.notna(cell(r, 'Carpool with')) else ''),
            'mindays': int(md) if pd.notna(md := pd.to_numeric(cell(r, 'MinDays'), errors='coerce')) else 0,
        }
    _nw_unknown = []
    for a in list(staff):
        for b in list(staff[a]['nw']):
            if b in staff: staff[b]['nw'].add(a)
            else: _nw_unknown.append((a, b))      # reported once warns exists

    warns = []
    for a, b in _nw_unknown:
        warns.append(f"not-work-with name unmatched for {a}: '{b}' — NOT applied")

    # ---- Away / Returning ------------------------------------------------------------
    # These are typed by hand, so infer the convention from the column as a whole
    # rather than assuming: if any first component is > 12 it must be day-first.
    _dparts = []
    for st in staff.values():
        for v in (st['away_raw'], st['back_raw']):
            mm = re.match(r'\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})', v)
            if mm: _dparts.append((int(mm.group(1)), int(mm.group(2))))
    if any(a > 12 for a, b in _dparts):   AWAY_DAYFIRST, _away_sure = True, True
    elif any(b > 12 for a, b in _dparts): AWAY_DAYFIRST, _away_sure = False, True
    else:                                 AWAY_DAYFIRST, _away_sure = True, not _dparts

    def _pdate(v):
        if not v: return None
        d = _to_date(v, AWAY_DAYFIRST)
        return None if pd.isna(d) else d.normalize()

    for sid, st in staff.items():
        st['away'] = _pdate(st['away_raw'])
        st['back'] = _pdate(st['back_raw'])
        if st['away_raw'] and st['away'] is None:
            warns.append(f"away date unreadable for {sid}: '{st['away_raw']}' — IGNORED, "
                         "they may be rostered while away")
        if st['back_raw'] and st['back'] is None:
            warns.append(f"returning date unreadable for {sid}: '{st['back_raw']}' — treated "
                         "as still away")
        if st['away'] and st['back'] and st['back'] < st['away']:
            warns.append(f"{sid}: Returning ({st['back_raw']}) is before Away ({st['away_raw']}) "
                         "— check the dates")

    def is_away(s, dt):
        st = staff[s]
        if not st['away'] or dt < st['away']: return False
        return st['back'] is None or dt <= st['back']

    # ---- sites (Region / LGA / Ticks) ----------------------------------------------
    SI = _frame(payload, 'sites')
    SI = SI[SI['Roster site name (Official name)'].notna()]
    site_info = {}
    for _, r in SI.iterrows():
        nm = str(r['Roster site name (Official name)']).strip()
        if not nm or nm == 'nan': continue
        site_info[nm] = {'pref_atm': str(r.get('Preferred ATM', '')).strip() if pd.notna(r.get('Preferred ATM')) else '',
                         'region': str(r.get('Region', '')).strip() if pd.notna(r.get('Region')) else '',
                         'lga':    str(r.get('LGA', '')).strip()    if pd.notna(r.get('LGA'))    else '',
                         'ticks':  tru(r.get('Ticks'))}
    TICK_SITES = {n for n, i in site_info.items() if i['ticks']}

    def resolve_pref_token(tok):
        """token -> set of site names, matching Region, LGA (aliased) and site-name word."""
        out, tl = set(), tok.strip().lower()
        targets = PREF_ALIASES.get(tl, [tok.strip()])
        for tgt in targets:
            for n, i in site_info.items():
                if i['region'].lower() == tgt.lower() or i['lga'].lower() == tgt.lower():
                    out.add(n)
        pat = re.compile(r'\b' + re.escape(tok.strip()) + r'\b', re.I)
        out |= {n for n in site_info if pat.search(n)}
        return out

    def resolve_excl_token(tok):
        """token -> (set of site names, how it matched).

        Exclusions resolve against Region and LGA EXACTLY (with the same aliases as
        preferences). Only if a token matches neither do we fall back to matching the
        site name, so 'North' (the region) can never accidentally exclude 'North Head'.
        """
        t = tok.strip()
        targets = PREF_ALIASES.get(t.lower(), [t])
        hits = {n for tgt in targets for n, i in site_info.items()
                if i['region'].lower() == tgt.lower() or i['lga'].lower() == tgt.lower()}
        if hits: return hits, 'region/council'
        pat = re.compile(r'\b' + re.escape(t) + r'\b', re.I)
        hits = {n for n in site_info if pat.search(n)}
        if hits: return hits, 'site name'
        return set(), 'no match'

    def resolve_site_token(tok):
        """'Manly SWW' -> 'Manly SWW (Manly Scenic Walkway)'. Returns (sites, how)."""
        t = tok.strip().lower()
        if not t: return set(), 'none'
        exact = {nm for nm in site_info if nm.lower() == t}
        if exact: return exact, 'exact'
        short = lambda nm: nm.split(' (')[0].split('/')[0].strip().lower()
        sh = {nm for nm in site_info if short(nm) == t}
        if len(sh) == 1: return sh, 'short name'
        if len(sh) > 1: return sh, 'ambiguous'
        pre = {nm for nm in site_info if nm.lower().startswith(t)}
        if len(pre) == 1: return pre, 'prefix'
        if len(pre) > 1: return pre, 'ambiguous'
        return set(), 'none'

    excl_sites = {}            # staff -> set of FULL site names from Excluded sites
    excl_site_notes = []       # (staff, token, resolved, how)
    for sid, st in staff.items():
        blocked = set()
        for tok in st['excl']:
            hit, how = resolve_site_token(tok)
            if how == 'none':
                warns.append(f"excluded site unmatched for {sid}: '{tok}' — NOT applied")
            elif how == 'ambiguous':
                warns.append(f"excluded site ambiguous for {sid}: '{tok}' could be "
                             + ' | '.join(sorted(hit)) + " — NOT applied")
            else:
                blocked |= hit
                excl_site_notes.append((sid, tok, sorted(hit)[0], how))
        excl_sites[sid] = blocked

    excl_regions = {}          # staff -> set of site names blocked by Excluded regions
    excl_notes = []            # (staff, token, n_sites, how) for the run log
    for sid, st in staff.items():
        blocked = set()
        for tok in st['excl_reg_toks']:
            hit, how = resolve_excl_token(tok)
            if not hit:
                warns.append(f"excluded region unmatched for {sid}: '{tok}' — NOT applied")
            else:
                blocked |= hit
            excl_notes.append((sid, tok, len(hit), how))
        if blocked: excl_regions[sid] = blocked

    pref_sites = {}
    for sid, st in staff.items():
        ps = set()
        for tok in st['pref_toks']:
            hit = resolve_pref_token(tok)
            if not hit: warns.append(f"preference token unmatched for {sid}: '{tok}'")
            ps |= hit
        if ps: pref_sites[sid] = ps

    # ---- pairing + carpool ---------------------------------------------------------
    carpool = dict(CARPOOL_SEED)
    for sid, st in staff.items():
        if st['carpool']:
            if st['carpool'] in staff: carpool[sid] = st['carpool']
            else: warns.append(f"carpool driver '{st['carpool']}' for {sid} not in staff — ignored")
    carpool = {p: d for p, d in carpool.items() if p in staff and d in staff}

    PAIRS = set()
    for sid, st in staff.items():
        for tok in st['pair_toks']:
            t = tok.strip()
            if t in staff and t != sid:
                PAIRS.add(tuple(sorted((sid, t))))
            elif sid in carpool and (t.lower().startswith('lift') or carpool[sid].lower() in t.lower()):
                pass  # carpool note handled by the carpool map — not a pairing
            else:
                warns.append(f"pairing token unmatched for {sid}: '{t}'")

    # ---- travel ------------------------------------------------------------------
    M = _frame(payload, 'matrix')
    travel = {(str(r['staff_id']).strip(), str(r['site_id']).strip()): float(r['minutes']) for _, r in M.iterrows()}
    ALL = set(M['site_id'].astype(str).str.strip())
    tmin = lambda s, si: travel.get((s, si))
    _orphans = sorted(ALL - set(site_info))      # reported once, as a count, below

    # ---- demand ------------------------------------------------------------------
    D = _frame(payload, 'demand'); cols = {c.strip().lower(): c for c in D.columns}; col = lambda n: cols.get(n.lower())
    # infer the date convention from the WHOLE column: if any first part >12 it's d/m,
    # if any second part >12 it's m/d; ambiguous-only columns default to m/d (Sheet format)
    _parts = [re.match(r'\s*(\d{1,2})/(\d{1,2})/(\d{2,4})', str(v)) for v in D[col('Date')].dropna()]
    _parts = [(int(m.group(1)), int(m.group(2))) for m in _parts if m]
    DAYFIRST = any(a > 12 for a, b in _parts) and not any(b > 12 for a, b in _parts)
    def crew_of(r):
        cs = r.get(col('Crew size')) if col('Crew size') else None
        if cs is not None and pd.notna(cs) and str(cs).strip() != '': return int(round(float(cs)))
        sh = r.get(col('Site Hours')) if col('Site Hours') else None
        if sh is not None and pd.notna(sh) and str(sh).strip() != '': return max(2, math.ceil((float(sh) - 8) / 7.5 + 1))
        return None

    DEM, PIN, OFF, LOCK, seen_sup = [], {}, {}, set(), set()
    pin_overrides = []
    pin_away, pin_spray = [], []
    for i, r in D.iterrows():
        site = str(r[col('Site')]).strip() if col('Site') else ''
        if not site or site == 'nan': continue
        vid = f"V{i+1}"
        if site not in ALL:
            warns.append(f"skipped (site not in matrix): {site}"); continue
        dt = _to_date(str(r[col('Date')]).strip(), DAYFIRST)
        if pd.isna(dt): warns.append(f"skipped (bad date): {site}"); continue
        sup = str(r[col('Supervisor')]).strip() if pd.notna(r[col('Supervisor')]) else ''
        if sup and sup not in staff:
            warns.append(f"{site}: supervisor '{sup}' not in staff — treating as no-supervisor visit"); sup = ''
        if sup:
            k = (dt, sup)
            if k in seen_sup:
                warns.append(f"{str(r[col('Date')]).strip()} {site}: {sup} already supervises elsewhere this day — left WITHOUT a supervisor, assign one"); sup = ''
            else:
                seen_sup.add(k)
        cs = crew_of(r)
        if not cs: warns.append(f"skipped (no crew/hours): {site}"); continue
        sh_raw = r.get(col('Site Hours')) if col('Site Hours') else None
        site_hours = ('%g' % float(sh_raw)) if (sh_raw is not None and pd.notna(sh_raw) and str(sh_raw).strip() != '') else ''
        note = str(r[col('Notes')]).strip() if col('Notes') and pd.notna(r.get(col('Notes'))) else ''
        DEM.append({'v': vid, 'date': dt, 'day': str(r[col('Day')]).strip(), 'site': site, 'sup': sup, 'crew': cs,
                    'hours': site_hours, 'note': note,
                'spray': tru(r.get(col('Spraying?'))) if col('Spraying?') else False,
                'ute': tru(r.get(col('Force Work Ute'))) if col('Force Work Ute') else False,
                    'rq': [q.lower() for q in toks(r.get(col('Required quals')))], 'atm': tru(r.get(col('Force ATM')))})
        PIN[vid] = toks(r.get(col('Pin staff'))); OFF[vid] = set(toks(r.get(col('Keep off'))))
        if tru(r.get(col('Lock crew'))): LOCK.add(vid)

    if not DEM: return _fail("No usable demand rows — check the Demand tab.", warns)

    # ---- demanded sites that SITES doesn't know about --------------------------
    # Tick status, region, council, exclusions and Preferred ATM all come from
    # SITES. A demanded site missing from SITES (usually a rename in SITES that
    # wasn't carried into Demand) would otherwise be treated as tick-free and
    # region-less — silently switching off safety rules. FAIL SAFE: assume it may
    # have ticks until SITES says otherwise.
    UNKNOWN_SITES = sorted({v['site'] for v in DEM if v['site'] not in site_info})
    for _u in UNKNOWN_SITES:
        _n = sum(1 for v in DEM if v['site'] == _u)
        warns.append(f"DEMANDED SITE NOT IN SITES: {_u} ({_n} visit(s))")
    TICK_SITES = TICK_SITES | set(UNKNOWN_SITES)
    min_date = min(v['date'] for v in DEM)
    for v in DEM: v['fn'] = (v['date'] - min_date).days // FORTNIGHT_DAYS
    dates = sorted({v['date'] for v in DEM})

    # ---- model -------------------------------------------------------------------
    m = cp_model.CpModel(); pool = list(staff); x = {}
    def block_reason(s, v):
        """Why s cannot work visit v — None if they can. A pin overrides the
        availability grid (managers know about ad-hoc and one-off arrangements)
        but never the safety/consent rules."""
        if s == v['sup']: return 'is already the supervisor'
        st = staff[s]
        pinned = s in PIN[v['v']]
        if s in OFF[v['v']]: return 'is listed in Keep off for this visit'
        if v['site'] in excl_sites.get(s, ()): return 'has this site in their Excluded sites'
        if v['site'] in excl_regions.get(s, ()): return 'has this region/council excluded'
        if st['no_ticks'] and v['site'] in TICK_SITES: return 'cannot work tick sites'
        missing = [q for q in v['rq'] if q not in st['quals']]
        if missing: return 'is missing the required qualification(s): ' + ', '.join(missing)
        t = tmin(s, v['site'])
        if t is None: return 'has no travel time to this site (needs geocoding / matrix rebuild)'
        if t > st['mt']: return f'is {t:.0f} min away, over their {st["mt"]} min limit'
        if v.get('spray'):
            sp = st['spray']
            if sp.startswith('n'):          # an explicit No is consent — pins cannot override
                return 'is marked "No" for spraying and this is a spraying visit'
            if not sp.startswith('y') and not pinned:
                return ('is not marked "Yes" for spraying'
                        + (' (marked Unsure)' if sp else ' (blank in the Staff tab)'))
        if not pinned:
            if is_away(s, v['date']):
                return f'is away from {st["away_raw"]}' + (f' to {st["back_raw"]}' if st['back_raw'] else '')
            if v['day'] in OPTIN_DAYS: return 'is not pinned (weekend work is pin-only)'
            if not st['avail'].get(v['day'], False): return f'is not available on {v["day"]}'
        return None

    def elig(s, v):
        return block_reason(s, v) is None
    for v in DEM:
        for s in pool:
            if elig(s, v): x[(s, v['v'])] = m.NewBoolVar(f"x_{s}_{v['v']}")

    SHORT = {}
    for v in DEM:
        mem = [s for s in pool if (s, v['v']) in x]
        need = max(0, v['crew'] - (1 if v['sup'] else 0))
        # Managers pinning more people than the crew size means they want that many:
        # raise the cap rather than making the whole model infeasible.
        npins = sum(1 for s in PIN[v['v']] if (s, v['v']) in x)
        cap = max(need, npins)
        if cap > need:
            warns.append(f"{v['site']} ({v['day']}): {npins} staff pinned but crew size is "
                         f"{v['crew']} — crew size raised to fit the pins")
        m.Add(sum(x[(s, v['v'])] for s in mem) <= cap)
        sh = m.NewIntVar(0, need, f"short_{v['v']}")
        m.Add(sh >= need - sum(x[(s, v['v'])] for s in mem))
        SHORT[v['v']] = sh
        for s in PIN[v['v']]:
            if (s, v['v']) in x:
                m.Add(x[(s, v['v'])] == 1)
                if s in staff and is_away(s, v['date']):
                    pin_away.append((s, v['site'], v['date']))
                elif s in staff and not staff[s]['avail'].get(v['day'], False) \
                        and v['day'] not in OPTIN_DAYS:
                    pin_overrides.append((s, v['site'], v['day'], v['date'], staff[s]['adhoc']))
                if s in staff and v.get('spray') and not staff[s]['spray'].startswith('y'):
                    pin_spray.append((s, v['site'], v['date'], staff[s]['spray']))
            elif s not in staff:
                warns.append(f"PIN NOT HONOURED: {s} -> {v['site']} ({v['day']}) :: "
                             "that name is not a StaffID in the Staff tab")
            else:
                warns.append(f"PIN NOT HONOURED: {s} -> {v['site']} ({v['day']}) :: "
                             + block_reason(s, v))
        if v['v'] in LOCK:
            for s in mem:
                if s not in PIN[v['v']]: m.Add(x[(s, v['v'])] == 0)

    # double-booking per person per DATE (crew + supervisor duty)
    for s in pool:
        for d in dates:
            cv = [x[(s, v['v'])] for v in DEM if v['date'] == d and (s, v['v']) in x]
            sc = sum(1 for v in DEM if v['date'] == d and v['sup'] == s)
            if cv or sc: m.Add(sum(cv) + sc <= 1)

    # ---- "Alt" days: one week on, one week off ------------------------------------
    # Nobody knows from the sheet which week is "on", so the solver may pick, but it
    # can never use the same weekday in two consecutive weeks. Fixed duties
    # (supervising, pins) are counted first; if they already break the pattern the
    # rule is skipped for that pair and reported, rather than making the run fail.
    def _monday(dt): return (dt - pd.Timedelta(days=dt.weekday())).normalize()
    alt_clash = []                                   # (staff, day, week1, week2)
    for s in pool:
        for day in sorted(staff[s]['alt']):
            by_wk = {}
            for v in DEM:
                if v['day'] == day: by_wk.setdefault(_monday(v['date']), []).append(v)
            for wk in sorted(by_wk):
                nxt = wk + pd.Timedelta(days=7)
                if nxt not in by_wk: continue
                vs = by_wk[wk] + by_wk[nxt]
                fixed = sum(1 for v in vs if v['sup'] == s)
                pinned = {v['date'] for v in vs if s in PIN[v['v']] and (s, v['v']) in x}
                if fixed + len(pinned) >= 2:
                    alt_clash.append((s, day, wk, nxt)); continue
                m.Add(sum(x[(s, v['v'])] for v in vs if (s, v['v']) in x) + fixed <= 1)

    # ---- together vars (shared by carpool + pairing) -------------------------------
    _tog = {}
    def together(a, b, v):
        """BoolVar true iff a and b are both crew on visit v."""
        key = (min(a, b), max(a, b), v['v'])
        if key in _tog: return _tog[key]
        if (a, v['v']) not in x or (b, v['v']) not in x: return None
        t = m.NewBoolVar(f"tog_{key[0]}_{key[1]}_{v['v']}")
        m.Add(t <= x[(a, v['v'])]); m.Add(t <= x[(b, v['v'])])
        m.Add(t >= x[(a, v['v'])] + x[(b, v['v'])] - 1)
        _tog[key] = t
        return t

    pen = []
    UTE_VAR = {}                       # vid -> BoolVar (ute present) | None (nobody eligible)
    for v in DEM:
        vid, sup = v['v'], v['sup']; mem = [s for s in pool if (s, vid) in x]
        sup_prot = 1 if (sup and staff[sup]['protected']) else 0
        sup_male = 1 if (sup and staff[sup]['male']) else 0
        sup_unq  = 1 if (sup and staff[sup]['unq']) else 0
        sup_atm  = sup and staff[sup]['atm']
        for i, a in enumerate(mem):
            if sup and sup in staff[a]['nw']: m.Add(x[(a, vid)] == 0)
            for b in mem[i+1:]:
                if b in staff[a]['nw']: m.Add(x[(a, vid)] + x[(b, vid)] <= 1)
        pr = sum(x[(s, vid)] for s in mem if staff[s]['protected']) + sup_prot
        ml = sum(x[(s, vid)] for s in mem if staff[s]['male']) + sup_male
        e1 = m.NewBoolVar(f"e{vid}"); m.Add(pr == 1).OnlyEnforceIf(e1); m.Add(pr != 1).OnlyEnforceIf(e1.Not())
        ma = m.NewBoolVar(f"a{vid}"); m.Add(ml >= 1).OnlyEnforceIf(ma); m.Add(ml == 0).OnlyEnforceIf(ma.Not())
        lo = m.NewBoolVar(f"l{vid}"); m.AddBoolAnd([e1, ma]).OnlyEnforceIf(lo); m.AddBoolOr([e1.Not(), ma.Not()]).OnlyEnforceIf(lo.Not())
        pen.append(P_LONE_PROTECTED * lo)
        uq = sum(x[(s, vid)] for s in mem if staff[s]['unq']) + sup_unq
        ov = m.NewIntVar(0, v['crew'], f"o{vid}"); m.Add(ov >= uq - 2); pen.append(P_UNQUAL_OVER * ov)
        at = [x[(s, vid)] for s in mem if staff[s]['atm']] + ([1] if sup_atm else [])
        if at:
            h = m.NewBoolVar(f"h{vid}"); m.Add(sum(at) >= 1).OnlyEnforceIf(h); m.Add(sum(at) == 0).OnlyEnforceIf(h.Not())
            if v['atm']: m.Add(h == 1)
            else: pen.append(P_NO_ATM * h.Not())
        elif v['atm']: warns.append(f"{vid}: ATM required, none eligible")
        pen.append(P_UNDERSTAFF * SHORT[vid])
        # 'Force Work Ute': the supervisor's own ute, else someone on the crew with one
        if v.get('ute') and not (sup and staff[sup]['ute']):
            utes = [x[(s, vid)] for s in mem if staff[s]['ute']]
            if utes:
                u = m.NewBoolVar(f"ute{vid}"); m.Add(sum(utes) >= 1).OnlyEnforceIf(u)
                pen.append(P_NO_UTE * u.Not()); UTE_VAR[vid] = u
            else:
                UTE_VAR[vid] = None

    # ---- carpool adjustments -------------------------------------------------------
    # ride[(p, vid)] -> var true iff passenger p rides with their driver on v.
    # When riding, p's effective travel becomes CARPOOL_MIN in objective AND fairness.
    ride, obj_adj = {}, []          # obj_adj: extra (negative) objective terms
    fair_adj = {}                   # (s, fn) -> [(coef, var)] fairness corrections
    for p, drv in carpool.items():
        for v in DEM:
            if (p, v['v']) not in x: continue
            t = tmin(p, v['site'])
            base_obj  = int(round(t * (W_TRANSIT_MULT if staff[p]['transit'] else 1)))
            base_fair = int(round(t))
            save_obj, save_fair = max(0, base_obj - CARPOOL_MIN), max(0, base_fair - CARPOOL_MIN)
            if v['sup'] == drv:      # driver is the supervisor: together whenever p is assigned
                rv = x[(p, v['v'])]
            else:
                rv = together(p, drv, v)
                if rv is None: continue
            ride[(p, v['v'])] = rv
            obj_adj.append(-save_obj * rv)
            fair_adj.setdefault((p, v['fn']), []).append((-save_fair, rv))

    # ---- pairing rewards -----------------------------------------------------------
    pair_terms = []                 # (a, b, vid, var) satisfied-pairing indicators
    for a, b in sorted(PAIRS):
        for v in DEM:
            if v['sup'] == b and (a, v['v']) in x:
                pair_terms.append((a, b, v['v'], x[(a, v['v'])]))
            elif v['sup'] == a and (b, v['v']) in x:
                pair_terms.append((a, b, v['v'], x[(b, v['v'])]))
            else:
                tg = together(a, b, v)
                if tg is not None: pair_terms.append((a, b, v['v'], tg))
    obj_adj += [-R_PAIR * t for *_, t in pair_terms]

    # ---- site's Preferred ATM: pull them onto that site's visits ---------------------
    pref_atm_hits, pref_atm_bad = [], []
    for v in DEM:
        pa = site_info.get(v['site'], {}).get('pref_atm', '')
        if not pa: continue
        v['pref_atm'] = pa
        if pa not in staff:
            if pa not in [b[0] for b in pref_atm_bad]:
                pref_atm_bad.append((pa, v['site']))
            continue
        if (pa, v['v']) in x:
            pref_atm_hits.append((pa, v['v']))
            obj_adj.append(-R_SITE_ATM * x[(pa, v['v'])])
    for pa, site in pref_atm_bad:
        warns.append(f"preferred ATM '{pa}' for {site} is not a StaffID — ignored")

    # ---- preferred-site rewards ----------------------------------------------------
    pref_hits = [(s, v['v']) for v in DEM for s in pool
                 if (s, v['v']) in x and v['site'] in pref_sites.get(s, ())]
    obj_adj += [-R_PREF * x[k] for k in pref_hits]

    # ---- fairness: per-person per-fortnight excess over (50 * days worked) --------
    fns = sorted({v['fn'] for v in DEM})
    excess_vars = []
    for s in pool:
        for fn in fns:
            vfn = [v for v in DEM if v['fn'] == fn]
            crew_terms = [int(round(tmin(s, v['site']))) * x[(s, v['v'])] for v in vfn if (s, v['v']) in x]
            sup_const_travel = sum(int(round(tmin(s, v['site']) or 0)) for v in vfn if v['sup'] == s)
            sup_days = sum(1 for v in vfn if v['sup'] == s)
            if not crew_terms and not sup_days: continue
            days_worked = sum(x[(s, v['v'])] for v in vfn if (s, v['v']) in x) + sup_days  # <=1 per date
            travel_expr = sum(crew_terms) + sup_const_travel \
                          + sum(c * var for c, var in fair_adj.get((s, fn), []))
            ex = m.NewIntVar(0, 100000, f"ex_{s}_{fn}")
            m.Add(ex >= travel_expr - DAY_BUDGET_MIN * days_worked)
            excess_vars.append(ex)

    # ---- MinDays: weekly minimum days for flagged staff (soft, heavy penalty) -------
    def monday_of(dt): return (dt - pd.Timedelta(days=dt.weekday())).normalize()
    week_dates = {}
    for v in DEM: week_dates.setdefault(monday_of(v['date']), set()).add((v['date'], v['day']))
    mindays_deficits = []                      # (staff, week, req, deficit_var)
    for s in pool:
        md = staff[s]['mindays']
        if not md: continue
        for wk, dds in sorted(week_dates.items()):
            if not any(day in WEEKDAYS for _, day in dds): continue   # weekend-only week fragment
            # a week the Demand only touches for 1-2 weekdays (month starts/ends
            # mid-week) can't fairly carry a weekly minimum: skip it
            if len({d for d, day in dds if day in WEEKDAYS}) < 3: continue
            # days this week the person could actually work: available AND (eligible
            # for >=1 visit that date, or supervising that date)
            # (Alt days don't count towards the weekly minimum: half the weeks are off)
            possible = [d for d, day in dds if staff[s]['avail'].get(day, False)
                        and day not in staff[s]['alt']
                        and (any((s, v['v']) in x for v in DEM if v['date'] == d)
                             or any(v['sup'] == s for v in DEM if v['date'] == d))]
            req = min(md, len(possible))
            # nobody to warn about if they have no regular (non-Alt) day at all
            if not any(staff[s]['avail'].get(d2) and d2 not in staff[s]['alt'] for d2 in WEEKDAYS): continue
            # away for every regular day they'd have had this week: leave, not a problem
            _reg = [d for d, day in dds if staff[s]['avail'].get(day, False) and day not in staff[s]['alt']]
            if req < md and _reg and all(is_away(s, d) for d in _reg): continue
            if req < md:
                warns.append(f"MinDays: {s} wants {md}/wk but only {len(possible)} possible "
                             f"in week of {wk.strftime('%d %b')} (availability/eligibility)")
            if req == 0: continue
            days_expr = sum(x[(s, v['v'])] for v in DEM if v['date'].normalize() >= wk
                            and (v['date'] - wk).days < 7 and (s, v['v']) in x)                     + sum(1 for v in DEM if v['sup'] == s and v['date'].normalize() >= wk
                              and (v['date'] - wk).days < 7)
            dv = m.NewIntVar(0, md, f"mindef_{s}_{wk.strftime('%m%d')}")
            m.Add(dv >= req - days_expr)
            mindays_deficits.append((s, wk, req, dv))
    pen += [P_MINDAYS * dv for *_, dv in mindays_deficits]

    obj = []
    for v in DEM:
        for s in pool:
            if (s, v['v']) in x:
                t = tmin(s, v['site']); obj.append(int(round(t * (W_TRANSIT_MULT if staff[s]['transit'] else 1))) * x[(s, v['v'])])
    m.Minimize(sum(obj) + sum(pen) + FAIRNESS_WEIGHT * sum(excess_vars) + sum(obj_adj))

    sv = cp_model.CpSolver()
    sv.parameters.max_time_in_seconds = float(os.environ.get('SOLVE_SECONDS', 60))
    sv.parameters.num_search_workers = int(os.environ.get('SOLVE_WORKERS', 8))
    st = sv.Solve(m)
    print("Status:", sv.StatusName(st))
    if warns:
        print("\n*** WARNINGS ***");  [print("  !", w) for w in warns]
    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return _fail("No possible roster with these constraints (INFEASIBLE). "
                     "Usually an impossible pin or a supervisor clash.", warns)

    def riding(s, vid): return (s, vid) in ride and sv.Value(ride[(s, vid)])
    def eff(s, v):      return CARPOOL_MIN if riding(s, v['v']) else tmin(s, v['site'])

    cur = None; total = 0
    for v in sorted(DEM, key=lambda v: (v['date'], v['site'])):
        ds = v['date'].strftime('%a %d %b')
        if ds != cur: cur = ds; print(f"\n=== {ds} ===")
        crew = [s for s in pool if (s, v['v']) in x and sv.Value(x[(s, v['v'])])]
        ct = sum(eff(s, v) for s in crew); total += ct
        suptxt = v['sup'] if v['sup'] else '(no sup)'
        short = sv.Value(SHORT[v['v']])
        smark = f"  ** SHORT {short}" if short else ""
        def tag(s):
            mk = ('cp' if riding(s, v['v']) else '') + ('*' if v['site'] in pref_sites.get(s, ()) else '')
            return f"{s}({eff(s, v):.0f}{mk})"
        print(f"  {v['site'][:34]:34s} sup {suptxt:8s} -> " +
              ", ".join(tag(s) for s in sorted(crew, key=lambda s: eff(s, v))) + f"  [{ct:.0f}]{smark}")
    print(f"\nTOTAL CREW TRAVEL: {total:.0f} min over {len(dates)} days, {len(DEM)} visits  (carpool-adjusted; cp=riding, *=preferred site)")

    # fairness report
    print("\n--- fairness (per fortnight: travel vs 50-min/day budget, carpool-adjusted) ---")
    for fn in fns:
        vfn = [v for v in DEM if v['fn'] == fn]
        rows = []
        for s in pool:
            days = sum(1 for v in vfn if (s, v['v']) in x and sv.Value(x[(s, v['v'])])) + sum(1 for v in vfn if v['sup'] == s)
            if not days: continue
            trav = sum(eff(s, v) for v in vfn if (s, v['v']) in x and sv.Value(x[(s, v['v'])]))
            trav += sum((tmin(s, v['site']) or 0) for v in vfn if v['sup'] == s)
            rows.append((s, days, trav, trav - 50*days))
        over = [r for r in rows if r[3] > 0]
        rows.sort(key=lambda r: -r[3])
        print(f"Fortnight {fn}: {len(rows)} people worked, {len(over)} over budget. "
              f"avg/day high-to-low:")
        for s, days, trav, exc in rows[:6]:
            print(f"    {s:12s} {days}d  {trav:4.0f}min  ({trav/days:4.1f}/day)  {'OVER +'+str(int(exc)) if exc>0 else 'ok'}")

    # carpool report
    print("\n--- carpool ---")
    for p, drv in sorted(carpool.items()):
        tot = sum(1 for v in DEM if (p, v['v']) in x and sv.Value(x[(p, v['v'])]))
        rd  = sum(1 for v in DEM if riding(p, v['v']))
        print(f"    {p:8s} rides with {drv:8s}: {rd}/{tot} working days shared")

    # pairing report
    bycount = {}
    for a, b, vid, t in pair_terms:
        if sv.Value(t): bycount[(a, b)] = bycount.get((a, b), 0) + 1
    print("\n--- preferred pairings honoured (visits together) ---")
    for (a, b) in sorted(PAIRS):
        print(f"    {a} + {b}: {bycount.get((a, b), 0)}")

    # preferred-site report
    ph, tot_by = {}, {}
    for s, vid in pref_hits:
        if sv.Value(x[(s, vid)]): ph[s] = ph.get(s, 0) + 1
    for v in DEM:
        for s in pool:
            if s in pref_sites and (s, v['v']) in x and sv.Value(x[(s, v['v'])]): tot_by[s] = tot_by.get(s, 0) + 1
    print("\n--- preferred sites/regions: days on preferred / days worked ---")
    for s in sorted(pref_sites):
        if tot_by.get(s): print(f"    {s:12s} {ph.get(s, 0)}/{tot_by[s]}")

    # ---- MinDays compliance ----------------------------------------------------------
    if mindays_deficits:
        print("\n--- MinDays compliance (per week: worked / required) ---")
        for s in sorted({s for s, *_ in mindays_deficits}):
            parts = []
            for s2, wk, req, dv in mindays_deficits:
                if s2 != s: continue
                got = req - sv.Value(dv)
                parts.append(f"wk {wk.strftime('%d %b')}: {got}/{req}" + (" SHORT" if sv.Value(dv) else ""))
            print(f"    {s:10s} (min {staff[s]['mindays']}/wk)  " + " | ".join(parts))

    # ---- excess staff report: available but unused / under-used --------------------
    demand_dates = sorted({(v['date'], v['day']) for v in DEM})
    med_travel = {}
    dem_sites = {v['site'] for v in DEM}
    for s in pool:
        ts = sorted(t for site in dem_sites if (t := tmin(s, site)) is not None)
        med_travel[s] = ts[len(ts)//2] if ts else None
    util = []
    for s in pool:
        avail = sum(1 for d, day in demand_dates if staff[s]['avail'].get(day, False))
        if not avail: continue
        worked = len({v['date'] for v in DEM if ((s, v['v']) in x and sv.Value(x[(s, v['v'])])) or v['sup'] == s})
        elig = sum(1 for v in DEM if (s, v['v']) in x)
        util.append((s, worked, avail, elig, med_travel[s]))
    unused = [u for u in util if u[1] == 0]
    low = sorted((u for u in util if 0 < u[1] < 0.5 * u[2]), key=lambda u: u[1] / u[2])
    print(f"\n--- excess staff: {len(unused)} available-but-never-rostered, {len(low)} under 50% utilised ---")
    print("    (days worked / days available | eligible visits | median travel to this month's sites)")
    for s, w, a, e, mt in sorted(unused, key=lambda u: -(u[4] or 0)) + low[:10]:
        why = ('no travel matrix data (supervisor-only?)' if mt is None
               else 'not eligible for any visit' if e == 0
               else f'far from demand (median {mt:.0f} min)' if mt >= 60 else '')
        print(f"    {s:12s} {w}/{a} days | {e:3d} eligible | {('%.0f min' % mt) if mt is not None else ' n/a '}  {why}")

    understaffed = [(v['date'].strftime('%a %d %b'), v['site'], sv.Value(SHORT[v['v']])) for v in DEM if sv.Value(SHORT[v['v']]) > 0]
    print(f"\n--- understaffing: {sum(s for _,_,s in understaffed)} slots unfilled across {len(understaffed)} visits ---")
    for d, site, s in understaffed: print(f"    {d}  {site[:36]}  short {s}")

    # constraint check: safety rules + tick avoidance
    bad = sum(1 for v in DEM if (lambda full: (sum(staff[s]['protected'] for s in full)==1 and sum(staff[s]['male'] for s in full)>=1) or sum(staff[s]['unq'] for s in full)>2)([s for s in pool if (s,v['v']) in x and sv.Value(x[(s,v['v'])])] + ([v['sup']] if v['sup'] else [])))
    tickbad = sum(1 for v in DEM if v['site'] in TICK_SITES
                  for s in pool if (s, v['v']) in x and sv.Value(x[(s, v['v'])]) and staff[s]['no_ticks'])
    print("\nconstraint check:", "ALL OK" if bad + tickbad == 0 else f"{bad} safety + {tickbad} tick VIOLATIONS")

    # ---- stage 2: ATM role designation (post-roster, cannot change assignments) -----
    # One paid ATM per visit, chosen among ATM-qualified people PRESENT on that visit.
    # Crew ATMs preferred; supervisor only as fallback if no ATM crew. Balance:
    # minimise sum of squared per-person shift counts over the whole horizon.
    assigned = {v['v']: [s for s in pool if (s, v['v']) in x and sv.Value(x[(s, v['v'])])] for v in DEM}
    atm_cands = {}
    for v in DEM:
        crew_atms = [s for s in assigned[v['v']] if staff[s]['atm']]
        if crew_atms: atm_cands[v['v']] = crew_atms
        # NB: a supervisor is never the paid ATM — sup and ATM rates don't stack.
        # A visit with only a supervisor (or no ATM crew) simply has no paid ATM.

    designated = {}
    if atm_cands:
        m2 = cp_model.CpModel(); y = {}
        for vid, cands in atm_cands.items():
            for s in cands: y[(s, vid)] = m2.NewBoolVar(f"y_{s}_{vid}")
            m2.AddExactlyOne(y[(s, vid)] for s in cands)
        atm_people = sorted({s for cands in atm_cands.values() for s in cands})
        sq = []
        for s in atm_people:
            n = m2.NewIntVar(0, len(atm_cands), f"n_{s}")
            m2.Add(n == sum(y[(s, vid)] for vid in atm_cands if (s, vid) in y))
            q = m2.NewIntVar(0, len(atm_cands)**2, f"q_{s}")
            m2.AddMultiplicationEquality(q, [n, n]); sq.append(q)
        # If the site has a Preferred ATM and they are on the visit, give THEM the paid
        # role — otherwise the reward above just seats them without the grooming.
        pref_pen = []
        for vid, cands in atm_cands.items():
            v = next(z for z in DEM if z['v'] == vid)
            pa = site_info.get(v['site'], {}).get('pref_atm', '')
            if pa and pa in cands:
                pref_pen.append(y[(pa, vid)].Not())
        m2.Minimize(sum(sq) + P_NOT_PREF_ATM * sum(pref_pen))
        sv2 = cp_model.CpSolver(); sv2.parameters.max_time_in_seconds = 20
        st2 = sv2.Solve(m2)
        if st2 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            designated = {vid: next(s for s in cands if sv2.Value(y[(s, vid)]))
                          for vid, cands in atm_cands.items()}
        else:
            warns.append("ATM designation solve failed — no ATM roles assigned")
        counts = {}
        for s in designated.values(): counts[s] = counts.get(s, 0) + 1
        print(f"\n--- ATM role designation: {len(designated)}/{len(DEM)} visits have a paid ATM ---")
        for s in sorted(counts, key=lambda s: -counts[s]):
            elig_days = sum(1 for vid, c in atm_cands.items() if s in c)
            print(f"    {s:12s} {counts[s]} ATM shifts (eligible on {elig_days} visits)")
        novisits = [v for v in DEM if v['v'] not in atm_cands]
        if novisits: print(f"    ({len(novisits)} visits with no ATM-qualified person present)")


    # ---- MasterRoster.csv: one editable row per visit --------------------------------
    # Columns: Date | Day | Notes | Site | Site Hours | Supervisor | ATM | Staff1..N | Short
    # This is the draft source of truth. Managers edit names in place (swap staff,
    # clear a row for a cancellation), then re-render the calendar + staff view.
    visit_rows = []
    for v in sorted(DEM, key=lambda v: (v['date'], v['site'])):
        atm = designated.get(v['v'], '')
        crew = [s for s in pool if (s, v['v']) in x and sv.Value(x[(s, v['v'])]) and s != atm]
        crew.sort(key=lambda s: eff(s, v))
        visit_rows.append({'Date': f"{v['date'].month}/{v['date'].day}/{v['date'].year}",
                           'Day': v['day'], 'Notes': v['note'], 'Site': v['site'].replace('\n', ' / '),
                           'Site Hours': v['hours'], 'Supervisor': v['sup'], 'ATM': atm,
                           'crew': crew, 'Short': sv.Value(SHORT[v['v']]) or ''})
    maxcrew = max((len(r['crew']) for r in visit_rows), default=0)
    staff_cols = [f'Staff{i+1}' for i in range(maxcrew)]
    for r in visit_rows:
        for i2, c in enumerate(staff_cols): r[c] = r['crew'][i2] if i2 < len(r['crew']) else ''
        del r['crew']
    cols_out = ['Date', 'Day', 'Notes', 'Site', 'Site Hours', 'Supervisor', 'ATM'] + staff_cols + ['Short']
    master_roster = {'columns': cols_out,
                     'rows': [[('' if r.get(c) is None else r.get(c)) for c in cols_out] for r in visit_rows]}

    # ---- RunLog.csv: plain-English report for managers -------------------------------
    # Every warning the console prints also lands here, classified and paired with a
    # recommended action, so nothing is lost once this runs behind a button.
    import re as _re
    from datetime import datetime as _dt

    LOG = []   # (severity, area, what happened, what to do)
    def log(sev, area, what, todo=''): LOG.append((sev, area, what, todo))

    # --- headline numbers -------------------------------------------------------------
    _n_days = len(dates); _n_vis = len(DEM)
    _n_people = len({s for v in DEM for s in pool if (s, v['v']) in x and sv.Value(x[(s, v['v'])])}
                    | {v['sup'] for v in DEM if v['sup']})
    _short_slots = sum(sv.Value(SHORT[v['v']]) for v in DEM)
    _short_visits = sum(1 for v in DEM if sv.Value(SHORT[v['v']]))
    _solve_ok = sv.StatusName(st)

    log('SUMMARY', 'Roster', f"{_n_vis} visits across {_n_days} days, {_n_people} staff rostered", '')
    log('SUMMARY', 'Travel', f"{total:.0f} minutes total crew travel "
        f"({total/max(1,_n_vis):.0f} min per visit average)", '')
    log('SUMMARY', 'Solver', f"Solution quality: {_solve_ok}"
        + (' (best possible)' if _solve_ok == 'OPTIMAL' else ''), '')
    log('SUMMARY', 'ATM roles', f"{len(designated)} of {_n_vis} visits have a paid ATM", '')

    # --- crew shortfalls ---------------------------------------------------------------
    if _short_slots:
        log('ACTION NEEDED', 'Crew shortfall',
            f"{_short_slots} crew slots unfilled across {_short_visits} visits",
            'See the rows flagged in the Short column of Master Roster. Either accept a '
            'smaller crew, free up staff, or move the visit to another day.')
        for v in sorted(DEM, key=lambda v: (v['date'], v['site'])):
            sh = sv.Value(SHORT[v['v']])
            if sh:
                log('ACTION NEEDED', 'Crew shortfall',
                    f"{v['date'].strftime('%a %d %b')} — {v['site']}: {sh} short "
                    f"(wanted {v['crew']})",
                    'No eligible staff left that day. Check availability, quals or exclusions.')
    else:
        log('OK', 'Crew shortfall', 'Every visit fully crewed', '')

    # --- visits with no ATM ------------------------------------------------------------
    _no_atm = [v for v in DEM if v['v'] not in designated]
    if _no_atm:
        log('CHECK', 'ATM roles',
            f"{len(_no_atm)} visits have no ATM-qualified person on site",
            'Normal when the crew has no ATM. If one of these needs an ATM, tick Force ATM '
            'on that Demand row and re-run.')
        for v in _no_atm[:10]:
            log('CHECK', 'ATM roles', f"{v['date'].strftime('%a %d %b')} — {v['site']}: no ATM", '')

    # --- classify the raw warnings ------------------------------------------------------
    def classify(w):
        """raw warning -> (severity, area, plain english, recommended action)"""
        if w.startswith('excluded site unmatched') or w.startswith('excluded site ambiguous'):
            return ('ACTION NEEDED', 'Excluded sites', w,
                    'This exclusion is NOT being applied, so they may be rostered there. '
                    'Use the site name exactly as it appears in SITES.')
        if w.startswith('not-work-with name unmatched'):
            return ('ACTION NEEDED', 'Not work with', w,
                    'This pairing rule is NOT being applied, so they may be rostered together. '
                    'Check the spelling matches a StaffID.')
        if w.startswith('DEMANDED SITE NOT IN SITES'):
            site = w.split(': ', 1)[1]
            return ('ACTION NEEDED', 'Unknown site',
                    f"{site} is on the Demand tab but not in SITES",
                    'Usually a site renamed in SITES but not in Demand. Until fixed, tick rules, '
                    'exclusions, region and Preferred ATM cannot be checked for it, so it is '
                    'treated as a TICK site (tick-sensitive staff kept off). Make the Demand '
                    'name match SITES exactly, then run buildTravelMatrix.')
        if w.startswith('PIN NOT HONOURED'):
            m = _re.match(r"PIN NOT HONOURED: (\S+) -> (.+?) \((\w+)\) :: (.+)", w)
            if m:
                who, site, day, why = m.group(1), m.group(2), m.group(3), m.group(4)
                return ('ACTION NEEDED', 'Pinned staff',
                        f"{who} could not be pinned to {site} ({day}) — {who} {why}",
                        'Availability alone no longer blocks a pin, so this is a harder rule: '
                        'fix the Staff tab, pick someone else, or remove the pin.')
            return ('ACTION NEEDED', 'Pinned staff', w.replace('PIN NOT HONOURED: ', ''),
                    'Fix the Staff tab or remove the pin.')
        if w.startswith('skipped (site not in matrix)'):
            return ('ACTION NEEDED', 'Unknown site', w.split(': ', 1)[-1] + ' — not rostered at all',
                    'This site is missing from SITES or has no travel times. Add it, run '
                    'geocodeStaff / resolvePins, then buildTravelMatrix.')
        if w.startswith('skipped (bad date)'):
            return ('ACTION NEEDED', 'Demand tab', w.split(': ', 1)[-1] + ' — date unreadable, visit skipped',
                    'Fix the Date cell on that Demand row (M/D/YYYY).')
        if w.startswith('skipped (no crew/hours)'):
            return ('ACTION NEEDED', 'Demand tab', w.split(': ', 1)[-1] + ' — no crew size or hours, visit skipped',
                    'Fill in Site Hours or Crew size on that Demand row.')
        if 'already supervises elsewhere' in w:
            return ('ACTION NEEDED', 'Supervisor clash', w.split(' — ')[0].strip(),
                    'One supervisor cannot cover two sites the same day. Assign a different '
                    'supervisor on the Demand row, or move the visit.')
        if "not in staff" in w and 'supervisor' in w:
            return ('ACTION NEEDED', 'Supervisor unknown', w.split(' — ')[0].strip(),
                    'That name is not a StaffID in the Staff tab. Check spelling.')
        if 'pinned but crew size is' in w:
            return ('CHECK', 'Pinned staff', w.split(' — ')[0].strip(),
                    'More staff pinned than the crew size. They were all kept and the crew '
                    'size raised. Adjust Crew size on the Demand row if that was not intended.')
        if w.startswith('away date unreadable') or w.startswith('returning date unreadable'):
            return ('ACTION NEEDED', 'Away dates', w,
                    'Use a real date (e.g. 24/07/2026). Until then they can be rostered '
                    'while away.')
        if 'Returning' in w and 'is before Away' in w:
            return ('ACTION NEEDED', 'Away dates', w, 'Swap the two dates.')
        if w.startswith("preferred ATM '"):
            return ('ACTION NEEDED', 'Preferred ATM', w,
                    'The Preferred ATM column on SITES must contain a StaffID from the Staff tab.')
        if w.startswith('MinDays'):
            return ('CHECK', 'Minimum days', w.replace('MinDays: ', ''),
                    'Their guaranteed days cannot all be met this week — usually availability. '
                    'Lower MinDays or open up more days.')
        if w.startswith('ATM required, none eligible') or 'ATM required' in w:
            return ('CHECK', 'ATM roles', w,
                    'Force ATM is ticked but no ATM-qualified staff are free. Untick it or free someone up.')
        if 'not in matrix' in w or 'missing from Sites.csv' in w:
            return ('CHECK', 'Site data', w,
                    'Add this site to SITES (with Region and Ticks) so colours and tick rules work.')
        if w.startswith('excluded region unmatched'):
            return ('ACTION NEEDED', 'Excluded regions',
                    w.replace('excluded region unmatched for ', 'Unrecognised excluded region for '),
                    'This exclusion was NOT applied — they could be rostered somewhere they '
                    'refuse to go. Use a Region (North/South/East/West/Central) or a council name.')
        if w.startswith('preference token unmatched'):
            return ('INFO', 'Preferences', w.replace('preference token unmatched for ', 'Unrecognised preference for '),
                    'Spelling in the Preferred Sites / Preferred regions column. Ignored for now.')
        if w.startswith('pairing token unmatched'):
            return ('INFO', 'Preferences', w.replace('pairing token unmatched for ', 'Unrecognised pairing for '),
                    'That name is not a StaffID. Fix the Preferred Staff / Pairing cell, or ignore.')
        if w.startswith('carpool driver'):
            return ('INFO', 'Carpool', w, 'Check the Carpool with column uses a valid StaffID.')
        return ('CHECK', 'Other', w, '')

    for w in warns:
        log(*classify(w))

    # --- staff who never got work -------------------------------------------------------
    for s, wk, av, el, mt in sorted(unused, key=lambda u: -(u[4] or 0)):
        if el == 0 and mt is None:
            log('ACTION NEEDED', 'Missing travel data',
                f"{s} was available {av} days but has no travel times, so could not be rostered",
                'Add their suburb to Location in the Staff tab, run geocodeStaff, then '
                'buildTravelMatrix. (Ignore if they are a fixed-site supervisor.)')
        else:
            log('CHECK', 'Not rostered',
                f"{s} was available {av} days but got no work"
                + (f" (median travel {mt:.0f} min)" if mt else ''),
                'Usually travel distance. Give them a MinDays value in the Staff tab if they '
                'must be rostered.')

    # --- minimum days not met ------------------------------------------------------------
    if mindays_deficits:
        for s, wk, req, dv in mindays_deficits:
            if sv.Value(dv):
                log('ACTION NEEDED', 'Minimum days',
                    f"{s} got {req - sv.Value(dv)} of {req} guaranteed days, week of {wk.strftime('%d %b')}",
                    'Not enough eligible visits that week. Check their availability and site exclusions.')

    # --- safety check --------------------------------------------------------------------
    if bad + tickbad:
        log('ACTION NEEDED', 'Safety rules',
            f"{bad} crew-composition and {tickbad} tick-site problems",
            'Review before publishing this roster.')
    elif UNKNOWN_SITES:
        log('CHECK', 'Safety rules',
            f"Crew safety rules satisfied, but {len(UNKNOWN_SITES)} site(s) could not be checked "
            "against SITES — see Unknown site above",
            'Those sites were treated as tick sites to stay safe. Fix the names to verify fully.')
    else:
        log('OK', 'Safety rules', 'All crew safety rules satisfied', '')

    if _orphans:
        log('INFO', 'Site data',
            f"{len(_orphans)} site name(s) in TravelMatrix are no longer in SITES (old names "
            "after a rename). Harmless — they are simply ignored.", '')


    # --- spare capacity: who was free each day but not rostered -------------------------
    _assigned_on = {}
    for v in DEM:
        d = v['date']
        _assigned_on.setdefault(d, set())
        if v['sup']: _assigned_on[d].add(v['sup'])
        for s in pool:
            if (s, v['v']) in x and sv.Value(x[(s, v['v'])]): _assigned_on[d].add(s)

    # an 'Alt' day is off if the same weekday was worked the week before or after
    _worked = {s: {d for d, ss in _assigned_on.items() if s in ss} for s in pool}
    def _alt_blocked(s, d):
        day = WEEKDAYS[d.weekday()] if d.weekday() < 5 else None
        if day not in staff[s]['alt']: return False
        return any((d + pd.Timedelta(days=k)) in _worked[s] for k in (-7, 7))
    def _could_work(s, d):
        """available, not away, not on an Alt off-week, eligible for >=1 visit that day"""
        day = next(v['day'] for v in DEM if v['date'] == d)
        return (staff[s]['avail'].get(day, False) and not is_away(s, d) and not _alt_blocked(s, d)
                and any((s, v['v']) in x for v in DEM if v['date'] == d))

    _spare_total = 0
    for d in dates:
        day = next(v['day'] for v in DEM if v['date'] == d)
        if day in OPTIN_DAYS: continue                     # weekend: availability unknown
        free = sorted(s for s in pool if s not in _assigned_on.get(d, set()) and _could_work(s, d))
        _spare_total += len(free)
        if free:
            shown = ', '.join(free[:14]) + (f" (+{len(free)-14} more)" if len(free) > 14 else '')
            log('INFO', 'Spare staff',
                f"{d.strftime('%a %d %b')}: {len(free)} available and unrostered — {shown}",
                'These people can fill a gap on this day without breaking availability.')
        else:
            log('CHECK', 'Spare staff', f"{d.strftime('%a %d %b')}: nobody spare — everyone available is rostered",
                'No slack this day. Adding a visit means moving someone or leaving a site short.')
    log('SUMMARY', 'Spare staff',
        f"{_spare_total} spare staff-days across the month (average "
        f"{_spare_total/max(1,len([d for d in dates if next(v['day'] for v in DEM if v['date']==d) not in OPTIN_DAYS])):.1f} per weekday)",
        'More people are available than the crews need, so the optimiser drops the '
        'longest trips first. See "Days given" below for who that leaves short.')

    # --- days given vs days available, per person -------------------------------------
    # The crews are sized from Demand, so when supply beats demand someone sits out.
    # Travel is minimised, so it is usually the people who live furthest away (or use
    # public transport, costed x1.5). Make that visible person by person.
    _given = []
    for s in pool:
        could = [d for d in dates if next(v['day'] for v in DEM if v['date'] == d) not in OPTIN_DAYS
                 and (_could_work(s, d) or d in _worked[s])]
        if not could: continue
        got = sorted(_worked[s]); idle = [d for d in could if d not in _worked[s]]
        if not idle: continue
        near = [min([t for v in DEM if v['date'] == d and (s, v['v']) in x
                     for t in [tmin(s, v['site'])] if t is not None] or [0]) for d in idle]
        _given.append((len(got) / len(could), s, got, could, idle, sum(near) / len(near)))
    _given.sort()
    _low = [g for g in _given if g[0] < 0.5]
    if _low:
        log('SUMMARY', 'Days given', f"{len(_low)} staff got under half the days they were available",
            'Usually distance: they are the furthest from that day\'s sites. To give them more, '
            'pin them on a visit, raise a Crew size, or add their days to MinDays.')
    for ratio, s, got, could, idle, near in _given:
        sev = 'CHECK' if ratio < 0.5 else 'INFO'
        why = f"nearest site on their free days averaged {near:.0f} min away"
        if staff[s]['transit']: why += ', public transport (costed x1.5)'
        log(sev, 'Days given',
            f"{s}: {len(got)} of {len(could)} available days. Free on "
            + ', '.join(d.strftime('%a %d') for d in idle) + f" ({why})", '')

    # --- pinned staff working outside their availability -----------------------------
    if pin_overrides:
        planned = [p for p in pin_overrides if p[4]]        # ad-hoc: expected, no fuss
        unusual = [p for p in pin_overrides if not p[4]]
        if unusual:
            log('CHECK', 'Pinned staff',
                f"{len(unusual)} pinned assignment(s) ignore the availability grid",
                'The pin was honoured because you asked for it. Confirm each person '
                'actually can work that day.')
            for s, site, day, dt, _ in unusual:
                log('CHECK', 'Pinned staff',
                    f"{s} pinned to {site} on {dt.strftime('%a %d %b')} but the Staff tab "
                    f"shows them unavailable on {day}",
                    f"Either confirm with {s}, or update their {day} availability.")
        for s, site, day, dt, _ in planned:
            log('INFO', 'Pinned staff',
                f"{s} (ad-hoc) pinned to {site} on {dt.strftime('%a %d %b')}",
                'Ad-hoc staff are only ever rostered by pinning — working as intended.')

    # --- pinned staff with a very long trip -------------------------------------------
    PIN_TRAVEL_WARN = 120      # minutes one way
    _far = []
    for v in DEM:
        for s in PIN[v['v']]:
            if (s, v['v']) not in x: continue
            t = tmin(s, v['site'])
            if t is not None and t > PIN_TRAVEL_WARN:
                _far.append((s, v['site'], v['date'], t))
    for s, site, dt, t in sorted(_far, key=lambda z: -z[3]):
        log('CHECK', 'Pinned staff',
            f"{s} is pinned to {site} on {dt.strftime('%a %d %b')} — {t:.0f} minutes travel each way",
            'A pin overrides the travel balancing, so this trip is locked in. '
            'Check it is really what you want.')

    # --- Force Work Ute ---------------------------------------------------------------
    _ute_vis = [v for v in DEM if v.get('ute')]
    if _ute_vis:
        log('SUMMARY', 'Work ute', f"{len(_ute_vis)} visits need a work ute", '')
        for v in _ute_vis:
            vid, when = v['v'], f"{v['date'].strftime('%a %d %b')} {v['site']}"
            if v['sup'] and staff[v['sup']]['ute']:
                log('OK', 'Work ute', f"{when}: supervisor {v['sup']} has one", ''); continue
            who = [s for s in pool if (s, vid) in x and sv.Value(x[(s, vid)]) and staff[s]['ute']]
            if who:
                log('INFO', 'Work ute', f"{when}: {', '.join(who)} brings a work ute"
                    + (f" ({v['sup']} has none)" if v['sup'] else ''), '')
            else:
                log('ACTION NEEDED', 'Work ute',
                    f"{when}: no work ute on this visit"
                    + (' — nobody with a ute was free and eligible' if UTE_VAR.get(vid) is None else ''),
                    'Pin someone with a work ute (Work Ute = Yes in Staff), or change the supervisor.')

    # --- alternate-week ("Alt") days -----------------------------------------------------
    _alt_rows = []
    for s in sorted(pool):
        for day in sorted(staff[s]['alt']):
            worked = sorted({v['date'] for v in DEM if v['day'] == day and (
                v['sup'] == s or ((s, v['v']) in x and sv.Value(x[(s, v['v'])])))})
            if worked: _alt_rows.append((s, day, worked))
    if _alt_rows:
        log('SUMMARY', 'Alternate weeks',
            f"{len(_alt_rows)} staff rostered on an 'Alt' day (one week on, one week off)",
            'The sheet does not say which week is "on", so the optimiser picked. '
            'Check these against what each person expects.')
        for s, day, worked in _alt_rows:
            log('INFO', 'Alternate weeks',
                f"{s} ({day}, alternate weeks): rostered " + ', '.join(d.strftime('%d %b') for d in worked), '')
    for s, day, w1, w2 in alt_clash:
        log('CHECK', 'Alternate weeks',
            f"{s} is 'Alt' on {day} but supervises or is pinned on {day} in both weeks of "
            f"{w1.strftime('%d %b')} and {w2.strftime('%d %b')}",
            'Fixed duties win, so they are rostered both weeks. Change the supervisor or pin '
            'if that breaks their pattern.')

    # --- away dates ---------------------------------------------------------------------
    _away_people = [s for s in pool if staff[s]['away']]
    if _away_people:
        log('SUMMARY', 'Away dates', f"{len(_away_people)} staff have away dates recorded", '')
        for s in sorted(_away_people):
            _p = staff[s]        # NB: do not use 'st' here — that is the solver status
            covered = [v for v in DEM if is_away(s, v['date'])]
            if covered:
                log('INFO', 'Away dates',
                    f"{s} away {_p['away_raw']}" + (f" to {_p['back_raw']}" if _p['back_raw'] else ' (no return date)')
                    + f" — kept off {len(covered)} visits in this period", '')
    if not _away_sure and _dparts:
        log('CHECK', 'Away dates',
            'Away/Returning dates are ambiguous (no value above the 12th), so they were '
            f"read as {'day/month' if AWAY_DAYFIRST else 'month/day'}",
            'Check they mean what you intended, or type a value like 24/07/2026 somewhere '
            'in the column to make the format unambiguous.')
    for s, site, dt in pin_away:
        log('CHECK', 'Away dates',
            f"{s} is marked AWAY but is pinned to {site} on {dt.strftime('%a %d %b')}",
            'The pin was honoured. Confirm they are actually back, or remove the pin.')

    # --- spraying -------------------------------------------------------------------------
    _spray_visits = [v for v in DEM if v.get('spray')]
    if _spray_visits:
        _yes = [s for s in pool if staff[s]['spray'].startswith('y')]
        log('SUMMARY', 'Spraying', f"{len(_spray_visits)} spraying visits, "
            f"{len(_yes)} staff marked OK to spray", '')
        if not _yes:
            log('ACTION NEEDED', 'Spraying',
                'No staff are marked "Yes" in the Ok to spray? column',
                'Spraying visits can only be crewed by pinning. Fill in the Staff tab.')
        for s, site, dt, mark in pin_spray:
            log('CHECK', 'Spraying',
                f"{s} pinned to spraying visit {site} on {dt.strftime('%a %d %b')} but is marked "
                + (f"'{mark}'" if mark else 'blank') + ' for spraying',
                'The pin was honoured. Confirm they are happy to spray.')

    # --- preferred ATM ----------------------------------------------------------------------
    _pa_sites = {v['site']: site_info[v['site']]['pref_atm'] for v in DEM
                 if site_info.get(v['site'], {}).get('pref_atm')}
    if _pa_sites:
        got = 0
        for v in DEM:
            pa = site_info.get(v['site'], {}).get('pref_atm', '')
            if not pa: continue
            on = (pa, v['v']) in x and sv.Value(x[(pa, v['v'])])
            role = designated.get(v['v']) == pa
            if on: got += 1
            log('INFO' if on else 'CHECK', 'Preferred ATM',
                f"{v['site']} on {v['date'].strftime('%a %d %b')}: preferred ATM {pa} "
                + ('is on site and holds the ATM role' if role else
                   'is on site but another ATM holds the role' if on else 'could NOT be rostered'),
                '' if on else 'Usually availability, or they are already booked elsewhere that day.')
        log('SUMMARY', 'Preferred ATM',
            f"{len(_pa_sites)} sites have a preferred ATM; they were rostered on {got} visits", '')

    # --- excluded sites matched by a short name ---------------------------------------
    for sid, tok, full, how in excl_site_notes:
        if how != 'exact':
            log('INFO', 'Excluded sites',
                f"{sid}: '{tok}' read as {full} (matched on {how})",
                'Check this is the site they meant.')

    # --- excluded regions ----------------------------------------------------------------
    if excl_notes:
        log('SUMMARY', 'Excluded regions',
            f"{len({s for s, _, n, _ in excl_notes if n})} staff have region/council exclusions applied", '')
        for sid, tok, n, how in excl_notes:
            if n:
                log('INFO', 'Excluded regions',
                    f"{sid}: '{tok}' blocks {n} sites (matched on {how})",
                    'Matched by site name rather than a region/council — check it is what you meant.'
                    if how == 'site name' else '')

    # --- site data completeness (exclusions and colours depend on it) --------------------
    _dem_sites = {v['site'] for v in DEM}
    _no_lga = sorted(n for n in _dem_sites if not site_info.get(n, {}).get('lga'))
    _no_reg = sorted(n for n in _dem_sites if not site_info.get(n, {}).get('region'))
    if _no_lga or _no_reg:
        log('CHECK', 'Site data',
            f"{len(_no_lga)} rostered sites have no council and {len(_no_reg)} have no region",
            'Region/council exclusions cannot block these sites, and they render grey on the '
            'calendar. Fill the Region and LGA columns in SITES.')
        for n in (_no_lga or _no_reg)[:10]:
            log('CHECK', 'Site data', f"{n}: "
                + ', '.join(filter(None, ['no council' if n in _no_lga else '',
                                          'no region' if n in _no_reg else ''])), '')

    # --- carpool ------------------------------------------------------------------------
    if carpool:
        for p, drv in sorted(carpool.items()):
            worked = [v for v in DEM if (p, v['v']) in x and sv.Value(x[(p, v['v'])])]
            rode = [v for v in worked if riding(p, v['v'])]
            solo = [v for v in worked if not riding(p, v['v'])]
            if not worked:
                log('INFO', 'Carpool', f"{p} (rides with {drv}) was not rostered this month", '')
            elif not solo:
                log('OK', 'Carpool', f"{p} rode with {drv} on all {len(rode)} working days", '')
            else:
                log('CHECK', 'Carpool',
                    f"{p} rode with {drv} on {len(rode)} of {len(worked)} days — travelling "
                    f"alone on {', '.join(v['date'].strftime('%a %d %b') for v in solo[:4])}"
                    + ('…' if len(solo) > 4 else ''),
                    f"Check {p} has a way to get to those sites, or move them to a site {drv} is on.")
    else:
        log('INFO', 'Carpool', 'No carpool pairs configured',
            "Add a 'Carpool with' column to the Staff tab (passenger's row, driver's StaffID).")

    # --- weekend visits -------------------------------------------------------------------
    _wend = [v for v in DEM if v['day'] in OPTIN_DAYS]
    if _wend:
        log('CHECK', 'Weekend work',
            f"{len(_wend)} weekend visits — crew comes only from Pin staff",
            "The Staff tab has no Sat/Sun availability columns, so nobody is auto-assigned. "
            "Pin the staff who agreed to work, or add Sat/Sun columns to roster them normally.")
        for v in _wend:
            n_crew = sum(1 for s in pool if (s, v['v']) in x and sv.Value(x[(s, v['v'])]))
            log('CHECK' if n_crew == 0 else 'INFO', 'Weekend work',
                f"{v['date'].strftime('%a %d %b')} — {v['site']}: "
                f"{n_crew} crew pinned" + (' (supervisor only)' if n_crew == 0 else ''),
                'Add names to Pin staff on that Demand row.' if n_crew == 0 else '')

    # --- write ----------------------------------------------------------------------------
    _ORDER = {'SUMMARY': 0, 'ACTION NEEDED': 1, 'CHECK': 2, 'INFO': 3, 'OK': 4}
    LOG.sort(key=lambda r: _ORDER.get(r[0], 5))
    _counts = {k: sum(1 for r in LOG if r[0] == k) for k in _ORDER}
    _hdr = [('SUMMARY', 'Run', f"Generated {_dt.now().strftime('%a %d %b %Y, %I:%M %p')}", ''),
            ('SUMMARY', 'Run', f"{_counts['ACTION NEEDED']} things need action, "
                               f"{_counts['CHECK']} worth checking", '')]
    run_log = {'columns': ['Severity', 'Area', 'What happened', 'What to do'],
               'rows': [list(r) for r in (_hdr + LOG)]}
    return {'ok': True,
            'master_roster': master_roster,
            'run_log': run_log,
            'stats': {'visits': len(DEM), 'days': len(dates), 'staff_rostered': _n_people,
                      'total_travel_min': round(total), 'status': sv.StatusName(st),
                      'unfilled_slots': int(_short_slots),
                      'actions': _counts['ACTION NEEDED'], 'checks': _counts['CHECK']},
            'warnings': list(warns)}


# ---- CLI: read the CSVs, write the CSVs (unchanged behaviour) ----------------------
def _main():
    payload = {'staff':  pd.read_csv(STAFF_CSV, header=None),
               'sites':  pd.read_csv(SITES_CSV),
               'matrix': pd.read_csv(MATRIX_CSV),
               'demand': pd.read_csv(DEMAND_CSV)}
    res = solve(payload)
    pd.DataFrame(res['run_log']['rows'], columns=res['run_log']['columns']) \
      .to_csv('RunLog.csv', index=False)
    if not res['ok']:
        print('FAILED:', res['error']); print('wrote RunLog.csv'); return 1
    mr = res['master_roster']
    pd.DataFrame(mr['rows'], columns=mr['columns']).to_csv(OUT_CSV, index=False)
    s = res['stats']
    print(f"\nwrote {OUT_CSV}: {len(mr['rows'])} visit rows -> paste into the 'Master Roster' tab, "
          "then Apunga Roster menu > Render both")
    print(f"wrote RunLog.csv: {s['actions']} action, {s['checks']} check")
    return 0

if __name__ == '__main__':
    sys.exit(_main())
