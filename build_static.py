#!/usr/bin/env python3
"""
Employee_Dashboard_Static.html — the read-only, no-JavaScript build.

Why it exists: iOS previews HTML through Quick Look, which renders it with
JavaScript disabled, so the interactive dashboards show their layout and none of
their figures on an iPhone. Every number here is computed in Python and written
out as plain HTML and CSS. Nothing to install, nothing to host, no script tag in
the file at all.

How it pages without JavaScript: each screen is a <section class="page"> with an
id, and the CSS rule `.page:target{display:block}` shows whichever one the URL
fragment names. Ordinary <a href="#..."> links move between them. `#home` is the
default and is hidden by `.page:target ~ #home` once any page is selected, which
is why #home is emitted last. Paging is why this scales — at 900 people a single
scrolling document ran to roughly 294 phone screens.

Structure:
    home --+-- practice          scale, grades, mix
           +-- capacity --+-- stretched / low utilized / burnout   names needing a decision
           |              +-- util-<comp>                   the full month-by-month grid
           +-- directory ---- dir-<comp> ---- dir-<comp>-<grade>   people, with bios
           +-- capability      skills and who holds them
           +-- value           billable value by grade, competency, territory

Rules, identical to the interactive builds:
  - Managing Directors are HIDDEN, not excluded: their hours count in every
    total, rate and load, but they get no utilization row, pill or exception
    listing of their own.
  - Population is "all contributors" - everyone who worked in the period,
    leavers included - matching what the desktop build opens on.
  - The window is FY Jul-Jun.

Called from generate_report.py after the interactive builds; it does not touch
them. Delete this file and the other three still build.
"""

import html
import re

# ---- grade handling, mirroring the templates ----------------------------
ROLE_GROUPS = [
    ('Managing Directors', r'managing\s*director|md\b'),
    ('Directors',          r'^director$|^d\b'),
    ('Senior Managers',    r'senior\s*manager|sr\.?\s*manager|^sm\b'),
    ('Managers',           r'^manager$|^m\b'),
    ('Senior Associates',  r'senior\s*associate|^sa\d*$|^sa\b'),
    ('Associates',         r'^associate(\s*\d+)?$|^a\d*$|^a\b'),
]
GRADE_CARD_ROLE = {
    'Managing Directors': 'Managing Director', 'Directors': 'Director',
    'Senior Managers': 'Senior Manager', 'Managers': 'Manager',
    'Senior Associates': 'Senior Associate', 'Associates': 'Associate',
}
GRADE_ORDER = ['Managing Directors', 'Directors', 'Senior Managers',
               'Managers', 'Senior Associates', 'Associates', 'Other']
_MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

# Photos are base64-embedded. At 900 people that can run to hundreds of
# megabytes, which nobody can email, so there is a budget: past it the build
# falls back to initials, which look fine and cost nothing.
PHOTO_BUDGET_BYTES = 6 * 1024 * 1024

ICON = {
    'practice': '<path d="M3 21V8l9-5 9 5v13"/><path d="M9 21v-6h6v6"/>',
    'capacity': '<path d="M3 3v18h18"/><path d="M7 15l4-5 3 3 5-7"/>',
    'directory': '<circle cx="9" cy="8" r="3.2"/><path d="M3.5 20a5.5 5.5 0 0 1 11 0"/>'
                 '<path d="M17 11.5a2.7 2.7 0 1 0-2-4.6"/><path d="M16.5 20a5 5 0 0 0-1.7-3.6"/>',
    'capability': '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.4"/>',
    'value': '<path d="M12 1v22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
    'grid': '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>',
    'alert': '<path d="M12 3 2 20h20L12 3z"/><path d="M12 10v4"/><path d="M12 17h.01"/>',
}


def svg(kind, cls='ic'):
    return ('<svg class="%s" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" '
            'stroke-linecap="round" stroke-linejoin="round">%s</svg>' % (cls, ICON.get(kind, '')))


def group_for_role(role):
    r = str(role or '').strip()
    for title, pat in ROLE_GROUPS:
        if re.search(pat, r, re.I):
            return title
    return 'Other'


def is_md(role):
    return group_for_role(role) == 'Managing Directors'


def e(v):
    return html.escape('' if v is None else str(v), quote=True)


def slug(v):
    s = re.sub(r'[^a-z0-9]+', '-', str(v or '').lower()).strip('-')
    return s or 'none'


def num(v):
    if v is None or v == '' or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '').rstrip('%').strip()
    try:
        return float(s)
    except ValueError:
        return None


def month_key(my):
    try:
        mo, yr = str(my).split('-')
        return (int(yr), _MON.index(mo[:3].title()) + 1)
    except Exception:
        return (0, 0)


def fmt_month(my):
    try:
        mo, yr = str(my).split('-')
        return "%s '%s" % (mo[:3].title(), yr[2:])
    except Exception:
        return str(my)


def util_colour(v):
    if v is None:
        return ('#F4F4F4', '#9A9A9A')
    if v < 40:
        return ('#FFE8D6', '#8A4B12')
    if v < 60:
        return ('#FCC288', '#6B3A0E')
    if v < 100:
        return ('#EB8C00', '#FFFFFF')
    return ('#E0301E', '#FFFFFF')


def n0(x):
    return '{:,}'.format(int(round(x)))


def money(v):
    if v >= 1e6:
        return '$%.1fM' % (v / 1e6)
    if v >= 1e3:
        return '$%dk' % round(v / 1000)
    return '$%d' % round(v)


def pct(v):
    return '&mdash;' if v is None else '%d%%' % round(v)


def rate_lookup(rates):
    table = {}
    for r in rates:
        v = num(r.get('Hourly Rate (USD)') or r.get('Hourly Rate') or r.get('Rate'))
        if v is not None:
            table[(str(r.get('Role', '')).strip().lower(),
                   str(r.get('Territory', '')).strip().lower())] = v

    def rate(role, terr):
        raw = str(role or '').strip()
        names = [raw]
        if re.fullmatch(r'a2', raw, re.I) or re.search(r'associate\s*2', raw, re.I):
            names.append('Associate 2')
        g = GRADE_CARD_ROLE.get(group_for_role(raw))
        if g and g != raw:
            names.append(g)
        t = str(terr or '').strip().lower()
        for nm in names:
            nl = str(nm).strip().lower()
            if (nl, t) in table:
                return table[(nl, t)]
            if (nl, 'standard') in table:
                return table[(nl, 'standard')]
        return 0.0
    return rate


CSS = """
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:#EFEFEF;color:#1F1F1F;
     font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;
     font-size:15px;line-height:1.5}
a{color:inherit;text-decoration:none;-webkit-tap-highlight-color:rgba(208,74,2,.12)}

/* ---- paging: one <section class="page"> per screen, switched by :target ---- */
.page{display:none}
/* :target scrolls the section to y=0, which would be behind the sticky header. */
.page,#home{scroll-margin-top:70px}
.page:target{display:block}
#home{display:block}
.page:target ~ #home{display:none}

.bar{height:4px;background:linear-gradient(90deg,#E0301E,#D04A02 38%,#EB8C00 72%,#FFB600)}
.top{background:#1F1F1F;color:#fff;padding:11px 15px;display:flex;align-items:center;gap:11px;
     position:sticky;top:0;z-index:5}
.top img{height:23px;background:#fff;border-radius:3px;padding:3px 5px}
.top .t{font-size:14.5px;font-weight:700;letter-spacing:.2px;line-height:1.25}
.top .d{font-size:10.5px;color:#B4B4B4;margin-top:1px}
.top .hm{margin-left:auto;border:1px solid #4A4A4A;border-radius:7px;padding:6px 11px;
         font-size:11.5px;font-weight:700;color:#EDEDED;white-space:nowrap}
.wrap{max-width:900px;margin:0 auto;padding:14px 13px 60px}

.crumb{display:flex;flex-wrap:wrap;align-items:center;gap:5px;font-size:11.5px;font-weight:600;
       color:#7A7A7A;margin:0 2px 12px}
.crumb a{color:#D04A02}
.crumb i{font-style:normal;color:#C4C4C4}
h1.pg{margin:0 0 3px;font-size:23px;font-weight:800;letter-spacing:-.5px;line-height:1.2}
p.lede{margin:0 0 16px;font-size:13.5px;color:#6B6B6B;line-height:1.55}
h2.sec{display:flex;align-items:center;gap:8px;margin:24px 0 10px;font-size:11px;font-weight:800;
       letter-spacing:1.2px;text-transform:uppercase;color:#D04A02}
h2.sec:after{content:'';flex:1;height:1px;background:#DEDEDE}

.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:9px}
.kpi{background:#fff;border-radius:11px;box-shadow:0 1px 2px rgba(0,0,0,.07);padding:13px 14px;
     border-left:4px solid #D04A02}
.kpi.y{border-left-color:#FFB600}.kpi.o{border-left-color:#EB8C00}
.kpi.r{border-left-color:#E0301E}.kpi.g{border-left-color:#1F8A44}
.kpi .v{font-size:25px;font-weight:800;letter-spacing:-.8px;line-height:1.05}
.kpi .l{font-size:10px;font-weight:700;letter-spacing:.55px;text-transform:uppercase;color:#7A7A7A;
        margin-top:5px;line-height:1.35}
.kpi .s{font-size:11.5px;color:#8A8A8A;margin-top:3px}

.card{background:#fff;border-radius:11px;box-shadow:0 1px 2px rgba(0,0,0,.07);padding:15px 16px;margin-bottom:11px}
.note{font-size:12.5px;color:#7A7A7A;line-height:1.55;margin:10px 2px 0}

.navlist{display:grid;gap:9px}
.nav{display:flex;align-items:center;gap:13px;background:#fff;border-radius:11px;padding:14px 15px;
     box-shadow:0 1px 2px rgba(0,0,0,.07);border-left:4px solid transparent}
.nav:active{background:#FAFAFA}
.nav .ic{width:21px;height:21px;flex:none;color:#D04A02}
.nav .tx{flex:1;min-width:0}
.nav b{display:block;font-size:15px;font-weight:700}
.nav em{display:block;font-style:normal;font-size:12px;color:#8A8A8A;margin-top:2px;line-height:1.4}
.nav .ch{color:#C4C4C4;font-size:19px;line-height:1;flex:none}
.nav.a{border-left-color:#D04A02}.nav.b{border-left-color:#EB8C00}
.nav.c{border-left-color:#FFB600}.nav.d{border-left-color:#E0301E}.nav.e{border-left-color:#A32020}

.bl{display:flex;align-items:center;gap:9px;margin:7px 0}
.bl .k{flex:0 0 40%;font-size:12.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bl .t{display:block;flex:1;height:9px;background:#F0F0F0;border-radius:5px;overflow:hidden}
.bl .f{display:block;height:100%;background:#D04A02;border-radius:5px}
.bl .v{flex:0 0 auto;font-size:12px;font-weight:700;font-variant-numeric:tabular-nums;min-width:46px;text-align:right}

.grp{display:flex;align-items:baseline;gap:7px;font-size:10.5px;font-weight:800;letter-spacing:.8px;
     text-transform:uppercase;color:#D04A02;margin:16px 0 5px;padding-left:8px;border-left:3px solid #D04A02}
details.per{background:#fff;border-radius:10px;box-shadow:0 1px 2px rgba(0,0,0,.07);margin-bottom:7px;overflow:hidden}
details.per>summary{cursor:pointer;list-style:none;display:flex;gap:11px;align-items:center;padding:11px 13px}
details.per>summary::-webkit-details-marker{display:none}
.av{width:38px;height:38px;border-radius:50%;flex:none;object-fit:cover;background:#D04A02;color:#fff;
    display:flex;align-items:center;justify-content:center;font-size:12.5px;font-weight:700;letter-spacing:.3px}
.pn{font-weight:700;font-size:14.5px;line-height:1.25}
.pm{font-size:11.5px;color:#8A8A8A;margin-top:2px;line-height:1.35}
.pill{font-size:11px;font-weight:800;border-radius:6px;padding:3px 7px;flex:none}
.bio{padding:2px 14px 14px 62px;font-size:13px;color:#4A4A4A;line-height:1.6;border-top:1px solid #F2F2F2}
.bio dl{display:grid;grid-template-columns:auto 1fr;gap:3px 12px;margin:9px 0 0;font-size:12px}
.bio dt{color:#9A9A9A;font-weight:600}
.bio dd{margin:0}
.tags{margin-top:9px;display:flex;flex-wrap:wrap;gap:5px}
.tag{background:#F4F4F4;border-radius:5px;padding:3px 7px;font-size:11px;color:#5A5A5A}

.scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0 -16px;padding:0 16px}
table{border-collapse:separate;border-spacing:0;font-size:12px;white-space:nowrap}
thead th{background:#1F1F1F;color:#fff;font-size:9.5px;letter-spacing:.5px;text-transform:uppercase;
         font-weight:700;padding:7px 8px;text-align:left;position:sticky;top:0}
td{padding:5px 8px;border-bottom:1px solid #F4F4F4}
td.nm{font-weight:600;padding-right:14px}
td.u{text-align:center;font-weight:700}
td.num{text-align:right;font-variant-numeric:tabular-nums}
.legend{display:flex;flex-wrap:wrap;gap:10px;font-size:10.5px;color:#7A7A7A;margin:10px 2px 0}
.legend i{width:12px;height:12px;border-radius:3px;display:inline-block;vertical-align:-2px;margin-right:4px}
.left{font-size:9px;font-weight:800;color:#8A8A8A;border:1px solid #DEDEDE;border-radius:4px;
      padding:1px 4px;margin-left:6px;vertical-align:1px;letter-spacing:.4px}

.ex{display:flex;align-items:center;gap:11px;background:#fff;border-radius:10px;padding:10px 13px;
    margin-bottom:6px;box-shadow:0 1px 2px rgba(0,0,0,.07)}
.ex .tx{flex:1;min-width:0}
.ex b{display:block;font-size:14px}
.ex em{display:block;font-style:normal;font-size:11.5px;color:#8A8A8A;margin-top:2px}
.empty{background:#fff;border-radius:11px;padding:22px 16px;text-align:center;color:#8A8A8A;font-size:13.5px}
footer{font-size:11px;color:#9A9A9A;text-align:center;padding:26px 16px 0;line-height:1.7}
@media print{.page{display:block!important}.top{position:static}details.per{break-inside:avoid}}
"""


def build(data, enriched, leavers, photos, logo_uri, generated_str, out_file, photo_for):
    rates = data.get('rates') or []
    rate = rate_lookup(rates)
    util = data.get('util_jj') or data.get('utilization') or []

    months = sorted({str(r.get('Month Year')).strip() for r in util
                     if str(r.get('Month Year') or '').strip()}, key=month_key)
    last = months[-1] if months else ''
    cell, meta = {}, {}
    for r in util:
        w = str(r.get('WorkdayID') or '').strip()
        m = str(r.get('Month Year') or '').strip()
        if not w or not m:
            continue
        c = cell.setdefault((w, m), [0.0, 0.0, 0.0])
        c[0] += num(r.get('Chargeable Hours')) or 0
        c[1] += num(r.get('Standard Hours')) or 0
        c[2] += num(r.get('Training Hours')) or 0
        meta[(w, m)] = r

    roster = {str(x.get('WorkdayID') or '').strip(): x for x in enriched}
    leaver_by_id = {str(x.get('WorkdayID') or '').strip(): x for x in leavers}
    md_ids = {w for w, x in roster.items() if is_md(x.get('Role'))}
    all_ids = set(roster) | {w for (w, _m) in cell if w not in roster}
    shown_ids = all_ids - md_ids

    def mu(w, m):
        c = cell.get((w, m))
        return None if not c or not c[1] else c[0] / c[1] * 100

    def agg(ids, ms):
        ch = std = tr = 0.0
        for w in ids:
            for m in ms:
                c = cell.get((w, m))
                if c:
                    ch += c[0]; std += c[1]; tr += c[2]
        return ch, std, tr

    def avg(w):
        ch, std, _ = agg({w}, months)
        return None if not std else ch / std * 100

    def group_avg(ids):
        ch, std, _ = agg(ids, months)
        return None if not std else ch / std * 100

    CH, STD, TR = agg(all_ids, months)
    fy_avg = None if not STD else CH / STD * 100
    # The Low utilized cut-off: the latest month's average across everyone,
    # Managing Directors included, so it matches the rate the page reports.
    _lch, _lstd, _ = agg(all_ids, [last])
    last_avg = None if not _lstd else _lch / _lstd * 100

    # ---- exceptions -----------------------------------------------------
    L3 = months[-3:]
    stretched, low_util, burnout = [], [], []
    for w in shown_ids:
        p = roster.get(w) or leaver_by_id.get(w)
        if not p:
            continue
        v = mu(w, last)
        if v is not None and v > 100:
            stretched.append((v, p, w))
        if v is not None and last_avg is not None and v < last_avg:
            low_util.append((v, p, w))
        vs = [mu(w, m) for m in L3]
        vs = [x for x in vs if x is not None]
        if len(vs) >= min(2, len(L3)) and sum(vs) / len(vs) > 100:
            burnout.append((sum(vs) / len(vs), p, w))
    stretched.sort(key=lambda t: -t[0])
    low_util.sort(key=lambda t: t[0])
    burnout.sort(key=lambda t: -t[0])

    # ---- value ----------------------------------------------------------
    value = hours = unbilled = 0.0
    by_grade_v, by_comp_v, by_terr_v = {}, {}, {}
    for (w, m), r in meta.items():
        h = num(r.get('Chargeable Hours')) or 0
        rt = rate(r.get('Role'), r.get('Territory') or r.get('Territory Filter'))
        v = h * rt
        value += v
        hours += h
        unbilled += max(0.0, (num(r.get('Standard Hours')) or 0) - h) * rt
        g = group_for_role(r.get('Role'))
        by_grade_v[g] = by_grade_v.get(g, 0) + v
        ck = str(r.get('Competency Group') or r.get('Competency') or '').strip() or '(not set)'
        by_comp_v[ck] = by_comp_v.get(ck, 0) + v
        tk = str(r.get('Territory') or r.get('Territory Filter') or '').strip() or '(not set)'
        by_terr_v[tk] = by_terr_v.get(tk, 0) + v

    # ---- skills -----------------------------------------------------------
    holders, cats, skills_by_emp = {}, {}, {}
    for row in data.get('skills') or []:
        w = str(row.get('WorkdayID') or '').strip()
        cat = str(row.get('Skill Category') or '').strip() or 'Uncategorised'
        for s in re.split(r'[,;]', str(row.get('Skills') or '')):
            s = s.strip()
            if not s:
                continue
            skills_by_emp.setdefault(w, set()).add(s)
            if w in roster:
                k = s.lower()
                holders.setdefault(k, {'name': s, 'ids': set()})['ids'].add(w)
                cats.setdefault(k, set()).add(cat)
    thin = sum(1 for k in holders if len(holders[k]['ids']) <= 3)

    # ---- people organised competency -> grade -----------------------------
    tree = {}
    for x in enriched:
        c = str(x.get('Competency Group') or x.get('Competency') or '').strip() or '(not set)'
        tree.setdefault(c, {}).setdefault(group_for_role(x.get('Role')), []).append(x)
    comp_names = sorted(tree, key=lambda c: -sum(len(v) for v in tree[c].values()))

    # The utilization pages cover a different population from the directory:
    # leavers belong here (they worked in the period) and Managing Directors do
    # not (their hours count in the totals, but they get no row). Built
    # separately for exactly that reason.
    def comp_of(x):
        return str(x.get('Competency Group') or x.get('Competency') or '').strip() or '(not set)'
    util_by_comp = {}
    for x in enriched:
        util_by_comp.setdefault(comp_of(x), set()).add(str(x.get('WorkdayID') or '').strip())
    for x in leavers:
        util_by_comp.setdefault(comp_of(x), set()).add(str(x.get('WorkdayID') or '').strip())
    for c in list(util_by_comp):
        util_by_comp[c] = (util_by_comp[c] & all_ids) - md_ids
        if not util_by_comp[c]:
            del util_by_comp[c]
    util_comps = sorted(util_by_comp, key=lambda c: -len(util_by_comp[c]))
    for c in tree:
        for g in tree[c]:
            tree[c][g].sort(key=lambda x: str(x.get('Name') or ''))

    # ---- photo budget ------------------------------------------------------
    matched = [x for x in enriched if x.get('_photo')]
    photo_bytes = sum(len(x['_photo']) for x in matched)
    use_photos = photo_bytes <= PHOTO_BUDGET_BYTES
    if not use_photos:
        print('[note] Static build: %d photos would add %.0f MB, over the %.0f MB budget - '
              'using initials instead.' % (len(matched), photo_bytes / 1048576,
                                           PHOTO_BUDGET_BYTES / 1048576))

    locs = sorted({str(x.get('Location') or '').strip() for x in enriched if str(x.get('Location') or '').strip()})
    terrs = sorted({str(x.get('Territory') or '').strip() for x in enriched if str(x.get('Territory') or '').strip()})
    exps = [num(x.get('Experience')) for x in enriched]
    exps = [v for v in exps if v is not None]
    genders = [str(x.get('Gender') or '').strip().lower() for x in enriched]
    males, females = genders.count('male'), genders.count('female')
    seniors = [x for x in enriched if group_for_role(x.get('Role')) in
               ('Managing Directors', 'Directors', 'Senior Managers', 'Managers')]
    grades = {}
    for x in enriched:
        grades.setdefault(group_for_role(x.get('Role')), []).append(x)

    O = []
    a = O.append

    def kpi(v, l, cls='', sub=''):
        a('<div class="kpi %s"><div class="v">%s</div><div class="l">%s</div>%s</div>'
          % (cls, v, l, ('<div class="s">%s</div>' % sub) if sub else ''))

    def bars(pairs, fmt=str, order=None):
        d = dict(pairs)
        keys = [k for k in order if k in d] if order else [k for k, _ in pairs]
        mx = max([d[k] for k in keys] or [1]) or 1
        for k in keys:
            a('<div class="bl"><span class="k">%s</span><span class="t"><span class="f" '
              'style="width:%d%%"></span></span><span class="v">%s</span></div>'
              % (e(k), max(4, round(d[k] / mx * 100)), fmt(d[k])))

    def crumb(*parts):
        bits = []
        for i, (label, href) in enumerate(parts):
            if i:
                bits.append('<i>&rsaquo;</i>')
            bits.append(('<a href="#%s">%s</a>' % (href, e(label))) if href else e(label))
        a('<div class="crumb">%s</div>' % ''.join(bits))

    def person_block(x):
        w = str(x.get('WorkdayID') or '').strip()
        ph = x.get('_photo') if use_photos else ''
        av = ('<img class="av" src="%s" alt="">' % ph) if ph else \
             ('<span class="av">%s</span>' % e(x.get('_initials') or '?'))
        u = None if is_md(x.get('Role')) else avg(w)
        bg, fg = util_colour(u)
        pill = ('<span class="pill" style="background:%s;color:%s">%s</span>' % (bg, fg, pct(u))) \
               if u is not None else ''
        line = ' · '.join([t for t in [str(x.get('Role') or ''), str(x.get('Location') or ''),
                                            str(x.get('Territory') or '')] if t])
        a('<details class="per"><summary>%s<span class="tx"><span class="pn">%s</span>'
          '<div class="pm">%s</div></span>%s</summary><div class="bio">'
          % (av, e(x.get('Name')), e(line), pill))
        if x.get('Employee Description'):
            a('<div>%s</div>' % e(x.get('Employee Description')))
        a('<dl>')
        for lbl, key in [('Workday ID', 'WorkdayID'), ('Competency', 'Competency'),
                         ('Email', 'emailid'), ('RL Manager', 'RL Manager'),
                         ('Join date', 'Join Date'), ('Experience', 'Experience'),
                         ('Qualification', 'Qualification'), ('Work mode', 'Work Mode'),
                         ('Status', 'Status')]:
            v = x.get(key)
            if v not in (None, ''):
                a('<dt>%s</dt><dd>%s</dd>' % (e(lbl), e('%s yrs' % v if key == 'Experience' else v)))
        a('</dl>')
        sk = sorted(skills_by_emp.get(w, []))
        if sk:
            a('<div class="tags">%s</div>' % ''.join('<span class="tag">%s</span>' % e(s) for s in sk))
        a('</div></details>')

    def ex_list(rows, unit, cap=60):
        if not rows:
            a('<div class="empty">No one in this band &mdash; nothing to action.</div>')
            return
        for v, p, _w in rows[:cap]:
            bg, fg = util_colour(v)
            line = ' · '.join([t for t in [str(p.get('Role') or ''),
                                                str(p.get('Competency') or p.get('Competency Group') or ''),
                                                str(p.get('Location') or '')] if t])
            a('<div class="ex"><span class="pill" style="background:%s;color:%s">%s</span>'
              '<span class="tx"><b>%s</b><em>%s</em></span></div>'
              % (bg, fg, pct(v), e(p.get('Name')), e(line)))
        if len(rows) > cap:
            a('<p class="note">Showing the %d most extreme of <b>%s</b>. The rest are in the '
              'month-by-month grid.</p>' % (cap, n0(len(rows))))
        a('<p class="note">Measured on %s. %s</p>' % (e(fmt_month(last)), unit))

    def util_table(ids):
        rows = [(w, roster.get(w) or leaver_by_id.get(w), avg(w))
                for w in ids if w in all_ids and w not in md_ids]
        rows = [r for r in rows if r[1]]
        rows.sort(key=lambda t: -(t[2] if t[2] is not None else -1))
        if not rows:
            a('<div class="empty">No utilization reported for this group.</div>')
            return
        a('<div class="card"><div class="scroll"><table><thead><tr><th>Employee</th><th>FY</th>')
        for m in months:
            a('<th>%s</th>' % e(fmt_month(m)))
        a('</tr></thead><tbody>')
        for w, p, av in rows:
            gone = '' if w in roster else '<span class="left">LEFT</span>'
            a('<tr><td class="nm">%s%s</td>' % (e(p.get('Name')), gone))
            bg, fg = util_colour(av)
            a('<td class="u" style="background:%s;color:%s">%s</td>'
              % (bg, fg, pct(av) if av is not None else ''))
            for m in months:
                v = mu(w, m)
                bg, fg = util_colour(v)
                a('<td class="u" style="background:%s;color:%s">%s</td>'
                  % (bg, fg, pct(v) if v is not None else ''))
            a('</tr>')
        a('</tbody></table></div><div class="legend">')
        for lbl, c in [('&lt;40%', '#FFE8D6'), ('40&ndash;60%', '#FCC288'),
                       ('60&ndash;100%', '#EB8C00'), ('100%+', '#E0301E')]:
            a('<span><i style="background:%s"></i>%s</span>' % (c, lbl))
        a('</div></div>')

    # ================= document ==========================================
    a('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">')
    a('<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">')
    a('<title>Deals Practice &mdash; Read-only</title><style>%s</style></head><body>' % CSS)
    a('<div class="bar"></div><div class="top">')
    if logo_uri:
        a('<img src="%s" alt="PwC">' % logo_uri)
    a('<div><div class="t">Deals &mdash; People &amp; Capability</div>'
      '<div class="d">As of %s &middot; read-only</div></div>'
      '<a class="hm" href="#home">Menu</a></div>' % e(generated_str))

    # ---------------- practice --------------------------------------------
    a('<section class="page" id="practice"><div class="wrap">')
    crumb(('Menu', 'home'), ('The practice', None))
    a('<h1 class="pg">The practice</h1><p class="lede">Who we are, at %s.</p>' % e(generated_str))
    a('<div class="kpis">')
    kpi(n0(len(enriched)), 'Head Count')
    kpi(n0(len(comp_names)), 'Competencies', 'y')
    kpi(n0(len(locs)), 'Locations', 'o')
    kpi(('%.1f' % (sum(exps) / len(exps))) if exps else '&mdash;', 'Avg years experience', 'o')
    kpi('%s : %s' % (n0(males), n0(females)), 'Male : female', 'r')
    kpi(n0(len(terrs)), 'Territories served', 'y')
    a('</div>')
    a('<h2 class="sec">By grade</h2><div class="card">')
    bars([(g, len(v)) for g, v in grades.items()], n0, GRADE_ORDER)
    jun = len(enriched) - len(seniors)
    a('<p class="note">Leverage &mdash; <b>%s</b> below Manager for <b>%s</b> at Manager and above (%.1f&times;).</p>'
      % (n0(jun), n0(len(seniors)), (jun / len(seniors)) if seniors else 0))
    a('</div>')
    a('<h2 class="sec">By competency</h2><div class="card">')
    bars([(c, sum(len(v) for v in tree[c].values())) for c in comp_names], n0)
    a('</div>')
    a('<h2 class="sec">By location</h2><div class="card">')
    lc = {}
    for x in enriched:
        k = str(x.get('Location') or '').strip() or '(not set)'
        lc[k] = lc.get(k, 0) + 1
    bars(sorted(lc.items(), key=lambda t: -t[1])[:12], n0)
    a('</div></div></section>')

    # ---------------- capacity ---------------------------------------------
    a('<section class="page" id="capacity"><div class="wrap">')
    crumb(('Menu', 'home'), ('Capacity', None))
    a('<h1 class="pg">Capacity</h1><p class="lede">FY Jul&ndash;Jun, covering everyone who worked in '
      'the period. Latest month with data is %s.</p>' % e(fmt_month(last)))
    a('<div class="kpis">')
    kpi(pct(fy_avg), 'Average utilization', '', 'FY Jul&ndash;Jun')
    kpi(n0(CH), 'Chargeable hours', 'y')
    kpi(n0(STD), 'Standard hours', 'o')
    kpi(n0(TR), 'Training hours', 'o')
    a('</div>')
    a('<h2 class="sec">Needs a decision</h2><div class="navlist">')
    for href, cls, title, n, sub in [
            ('stretched', 'd', 'Stretched', len(stretched), 'Above 100% in ' + fmt_month(last)),
            ('lowutil', 'c', 'Low utilized', len(low_util),
             ('Below the %s average in ' % pct(last_avg)) + fmt_month(last)),
            ('burnout', 'e', 'Burnout watch', len(burnout), 'Three-month average above 100%')]:
        a('<a class="nav %s" href="#%s">%s<span class="tx"><b>%s &middot; %s</b><em>%s</em></span>'
          '<span class="ch">&rsaquo;</span></a>' % (cls, href, svg('alert'), title, n0(n), sub))
    a('</div>')
    a('<h2 class="sec">Utilization by grade</h2><div class="card">')
    gu = {}
    for g in GRADE_ORDER:
        ids = {str(x.get('WorkdayID') or '').strip() for x in grades.get(g, [])} - md_ids
        v = group_avg(ids) if ids else None
        if v is not None:
            gu[g] = v
    bars(list(gu.items()), pct, GRADE_ORDER)
    a('</div>')
    a('<h2 class="sec">Utilization by competency</h2><div class="card">')
    cu = {}
    for c in comp_names:
        ids = {str(x.get('WorkdayID') or '').strip() for v in tree[c].values() for x in v} - md_ids
        v = group_avg(ids) if ids else None
        if v is not None:
            cu[c] = v
    bars(sorted(cu.items(), key=lambda t: -t[1]), pct)
    a('</div>')
    a('<h2 class="sec">Month by month</h2><div class="navlist">')
    for c in util_comps:
        a('<a class="nav b" href="#util-%s">%s<span class="tx"><b>%s</b>'
          '<em>%s people &middot; the full grid</em></span><span class="ch">&rsaquo;</span></a>'
          % (slug(c), svg('grid'), e(c), n0(len(util_by_comp[c]))))
    a('</div>')
    nmd, nlv = len(md_ids), len(all_ids - set(roster))
    bits = []
    if nmd:
        bits.append('Includes <b>%s</b> Managing Director%s &mdash; counted in the figures above, but never '
                    'listed individually.' % (n0(nmd), '' if nmd == 1 else 's'))
    if nlv:
        bits.append('Covers everyone who worked in the period, including <b>%s</b> who ha%s since left.'
                    % (n0(nlv), 's' if nlv == 1 else 've'))
    if bits:
        a('<p class="note">%s</p>' % ' '.join(bits))
    a('</div></section>')

    for pid, title, rows, unit in [
            ('stretched', 'Stretched', stretched, 'Above 100% of standard hours.'),
            ('lowutil', 'Low utilized', low_util,
             'Below the team average for the latest month (%s).' % pct(last_avg)),
            ('burnout', 'Burnout watch', burnout,
             'Three-month average above 100%, needing at least two months reported.')]:
        a('<section class="page" id="%s"><div class="wrap">' % pid)
        crumb(('Menu', 'home'), ('Capacity', 'capacity'), (title, None))
        a('<h1 class="pg">%s</h1><p class="lede">%s people.</p>' % (e(title), n0(len(rows))))
        ex_list(rows, unit)
        a('</div></section>')

    for c in util_comps:
        ids = util_by_comp[c]
        a('<section class="page" id="util-%s"><div class="wrap">' % slug(c))
        crumb(('Menu', 'home'), ('Capacity', 'capacity'), (c, None))
        a('<h1 class="pg">%s</h1><p class="lede">Month-by-month utilization, %s&ndash;%s. '
          'Scroll the table sideways.</p>'
          % (e(c), e(fmt_month(months[0]) if months else ''), e(fmt_month(last))))
        util_table(ids)
        a('</div></section>')

    # ---------------- directory ---------------------------------------------
    a('<section class="page" id="directory"><div class="wrap">')
    crumb(('Menu', 'home'), ('Directory', None))
    a('<h1 class="pg">Directory</h1><p class="lede">%s people. Choose a competency, then a grade.</p>'
      % n0(len(enriched)))
    a('<div class="navlist">')
    for i, c in enumerate(comp_names):
        n = sum(len(v) for v in tree[c].values())
        gl = ', '.join('%s %s' % (n0(len(tree[c][g])), g) for g in GRADE_ORDER if g in tree[c])
        a('<a class="nav %s" href="#dir-%s">%s<span class="tx"><b>%s</b><em>%s people<br>%s</em></span>'
          '<span class="ch">&rsaquo;</span></a>'
          % ('abcde'[i % 5], slug(c), svg('directory'), e(c), n0(n), e(gl)))
    a('</div></div></section>')

    for c in comp_names:
        a('<section class="page" id="dir-%s"><div class="wrap">' % slug(c))
        crumb(('Menu', 'home'), ('Directory', 'directory'), (c, None))
        n = sum(len(v) for v in tree[c].values())
        a('<h1 class="pg">%s</h1><p class="lede">%s people. Choose a grade.</p>' % (e(c), n0(n)))
        a('<div class="navlist">')
        for i, g in enumerate([g for g in GRADE_ORDER if g in tree[c]]):
            ids = {str(x.get('WorkdayID') or '').strip() for x in tree[c][g]} - md_ids
            v = group_avg(ids) if ids else None
            sub = '%s people' % n0(len(tree[c][g]))
            if v is not None:
                sub += ' &middot; %s average utilization' % pct(v)
            a('<a class="nav %s" href="#dir-%s-%s">%s<span class="tx"><b>%s</b><em>%s</em></span>'
              '<span class="ch">&rsaquo;</span></a>'
              % ('abcde'[i % 5], slug(c), slug(g), svg('directory'), e(g), sub))
        a('</div></div></section>')
        for g in [g for g in GRADE_ORDER if g in tree[c]]:
            a('<section class="page" id="dir-%s-%s"><div class="wrap">' % (slug(c), slug(g)))
            crumb(('Menu', 'home'), ('Directory', 'directory'),
                  (c, 'dir-%s' % slug(c)), (g, None))
            a('<h1 class="pg">%s</h1><p class="lede">%s &middot; %s people. Tap a name for the full bio.</p>'
              % (e(g), e(c), n0(len(tree[c][g]))))
            for x in tree[c][g]:
                person_block(x)
            a('</div></section>')

    # ---------------- capability ---------------------------------------------
    a('<section class="page" id="capability"><div class="wrap">')
    crumb(('Menu', 'home'), ('Capability', None))
    a('<h1 class="pg">Capability</h1><p class="lede">Every recorded skill and how many people hold it.</p>')
    a('<div class="kpis">')
    kpi(n0(len(holders)), 'Distinct skills')
    kpi(n0(len({x for v in cats.values() for x in v})), 'Skill categories', 'y')
    kpi(n0(thin), 'Thinly covered', 'r', '3 or fewer people')
    kpi(n0(len(holders) - thin), 'Deeply covered', 'g', 'more than 3')
    a('</div>')
    a('<h2 class="sec">Every skill, thinnest first</h2><div class="card"><div class="scroll">'
      '<table><thead><tr><th>Skill</th><th>Category</th><th>People</th></tr></thead><tbody>')
    for k in sorted(holders, key=lambda k: (len(holders[k]['ids']), holders[k]['name'].lower())):
        n = len(holders[k]['ids'])
        col = '#E0301E' if n <= 3 else '#1F1F1F'
        a('<tr><td class="nm">%s</td><td style="color:#8A8A8A">%s</td>'
          '<td class="num" style="color:%s;font-weight:700">%s</td></tr>'
          % (e(holders[k]['name']), e(' · '.join(sorted(cats.get(k, [])))), col, n0(n)))
    a('</tbody></table></div><p class="note">Skills held by three or fewer people are in red &mdash; '
      'those are the ones a holiday or a departure is felt on.</p></div></div></section>')

    # ---------------- value ---------------------------------------------------
    if rates and hours:
        a('<section class="page" id="value"><div class="wrap">')
        crumb(('Menu', 'home'), ('Value', None))
        a('<h1 class="pg">Value</h1><p class="lede">Chargeable hours priced through the rate card, '
          'matched on grade and territory.</p>')
        a('<div class="kpis">')
        kpi(money(value), 'Billable value', '', 'FY Jul&ndash;Jun')
        kpi('$%d' % round(value / hours), 'Blended rate / hr', 'y')
        kpi(money(unbilled), 'Unbilled at card rate', 'o')
        kpi(money(value / max(1, len(all_ids))), 'Value per person', 'o')
        a('</div>')
        for title, dct, order in [('By grade', by_grade_v, GRADE_ORDER),
                                  ('By competency', by_comp_v, None),
                                  ('By territory', by_terr_v, None)]:
            items = list(dct.items()) if order else sorted(dct.items(), key=lambda t: -t[1])
            if not items:
                continue
            a('<h2 class="sec">%s</h2><div class="card">' % e(title))
            bars(items, money, order)
            a('</div>')
        a('<p class="note">Sample rates for demonstration only &mdash; not PwC&rsquo;s rate card. '
          'The method is real; the amounts are illustrative.</p>')
        a('</div></section>')

    # ---------------- home (emitted last so `.page:target ~ #home` can hide it)
    a('<section id="home"><div class="wrap">')
    a('<h1 class="pg">Deals practice</h1>'
      '<p class="lede">A read-only snapshot of people, capacity and capability. No JavaScript, so it '
      'opens anywhere &mdash; including the preview on an iPhone. Tap a section.</p>')
    a('<div class="kpis">')
    kpi(n0(len(enriched)), 'Head Count')
    kpi(pct(fy_avg), 'Avg utilization', 'y', 'FY Jul&ndash;Jun')
    kpi(n0(len(holders)), 'Distinct skills', 'o')
    if rates and hours:
        kpi(money(value), 'Billable value', 'g', 'indicative')
    a('</div>')
    a('<h2 class="sec">Sections</h2><div class="navlist">')
    navs = [('practice', 'a', 'practice', 'The practice',
             '%s people &middot; %s competencies &middot; %s locations'
             % (n0(len(enriched)), n0(len(comp_names)), n0(len(locs)))),
            ('capacity', 'b', 'capacity', 'Capacity',
             '%s average utilization &middot; %s stretched &middot; %s low utilized'
             % (pct(fy_avg), n0(len(stretched)), n0(len(low_util)))),
            ('directory', 'c', 'directory', 'Directory',
             'By competency, then grade &middot; %s people' % n0(len(enriched))),
            ('capability', 'd', 'capability', 'Capability',
             '%s skills &middot; %s thinly covered' % (n0(len(holders)), n0(thin)))]
    if rates and hours:
        navs.append(('value', 'e', 'value', 'Value',
                     '%s billable &middot; $%d blended rate' % (money(value), round(value / hours))))
    for href, cls, icon, title, sub in navs:
        a('<a class="nav %s" href="#%s">%s<span class="tx"><b>%s</b><em>%s</em></span>'
          '<span class="ch">&rsaquo;</span></a>' % (cls, href, svg(icon), title, sub))
    a('</div>')
    a('<footer>Generated %s from Employee Details.xlsx.<br>Read-only companion to the interactive '
      'dashboard &mdash; every figure matches it.%s</footer>'
      % (e(generated_str), '' if use_photos else '<br>Photos omitted to keep the file small enough to send.'))
    a('</div></section>')

    a('</body></html>')
    out_file.write_text('\n'.join(O), encoding='utf-8')
    return out_file
