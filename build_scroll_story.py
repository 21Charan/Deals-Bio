#!/usr/bin/env python3
"""
Deals Scroll Story — a single-page, scroll-driven showcase of the practice.

Reads the same workbook as the main dashboard and writes ONE self-contained
HTML file (logo + photos embedded, no network, no libraries):

    reads   ../03_Output files/Employee Details.xlsx
              Employee Details          roster, competencies, experience, join dates
              Utilization Full_Jul_Jun  hours for the FY Jul-Jun delivery figures
            ../01_Source/Images/<Employee ID or WorkdayID>.png
            ../01_Source/PwC Logo.jpg
            ../01_Source/Skill Images/<Skill name>.svg|png   (floating skill orbs)
    writes  ../03_Output files/Deals_Scroll_Story.html   (beside Employee_Dashboard.html)

That one HTML file is the only thing written.

The page is deliberately overall-level: headline KPIs, competencies, and the
leadership structure (Managing Director -> Directors -> competencies). It never
lists individual staff, so it holds up at 800 people.

Figures use the same definitions as the dashboard Overview, so the two agree:
  * Employees        every row of Employee Details
  * Offices          distinct Location
  * Territories      distinct Territory, ignoring "Other..."
  * Avg experience   mean of Experience
  * Avg yrs at PwC   mean time since Join Date (as at the build date)
  * Avg utilization  total chargeable / total standard hours, FY, everyone who
                     worked in the period (leavers included), MDs included
  * Chargeable/training hours   FY totals, same population
  * Spare capacity   latest month, roster only: sum of max(0, standard - chargeable)
                     / one full month (upper quartile of person-month standard hours)
Competency figures split the same totals by the competency each hours row
carries, so the competencies plus "Other" add back to the headline.

One deliberate difference: competencies are the "Competency Filter" column
(FDD, Valuations, Data & Analytics, BRS). The dashboard counts the raw
"Competency Group" column, which spells BRS two ways and counts Others.

Managing Directors count in every figure. On this page the MD appears by name
and photo at the centre of the leadership diagram, with no personal metrics.

All content is rendered here in Python, so the page is fully readable with
JavaScript off (iOS Quick Look). JavaScript only adds motion and click-through.

Usage:  python build_scroll_story.py
"""
import base64
import html
import json
import math
import re
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

from build_dark_dashboard import dark_html

# ---- Configuration: the same folder layout generate_report.py uses -------------------------
SCRIPT_DIR    = Path(__file__).resolve().parent          # 02_Scripts & ETL
ROOT          = SCRIPT_DIR.parent                         # Deals Skills and Bio
SOURCE_DIR    = ROOT / "01_Source"                        # Images, logo, Skill Images
OUTPUT_DIR    = ROOT / "03_Output files"                  # workbook in, both dashboards out
WORKBOOK      = OUTPUT_DIR / "Employee Details.xlsx"
IMAGES_DIR    = SOURCE_DIR / "Images"
LOGO          = SOURCE_DIR / "PwC Logo.jpg"
SKILL_IMG_DIR = SOURCE_DIR / "Skill Images"
OUT           = OUTPUT_DIR / "Deals_Scroll_Story.html"     # beside Employee_Dashboard.html
UTIL_SHEET = "Utilization Full_Jul_Jun"
RATES_SHEET = "Hourly Rates"
# the dashboard's skills sheet, under any of the names generate_report.py accepts for it
SKILLS_SHEETS = ("Employee Skills", "Skill Mapping", "Employee Skills Hierarchy")
RATE_ROLE = {"MD": "Managing Director", "Director": "Director", "SM": "Senior Manager", "M": "Manager",
             "SA": "Senior Associate", "A2": "Associate 2", "A1": "Associate"}
# Role as written in the workbook -> the grade key used throughout this script. The workbook may
# carry full names ("Senior Manager") or grade codes ("SM"); both land on the same key, grouped
# exactly as generate_report.py's ROLE_RENAME groups them for the dashboard (every Senior
# Associate level is one grade; A1 is "Associate", A2 "Associate 2"). Anything else is kept as
# written and sorts after the known grades.
ROLE_KEY = {
    "md": "MD", "managing director": "MD",
    "d": "Director", "director": "Director",
    "sm": "SM", "sr manager": "SM", "senior manager": "SM",
    "m": "M", "manager": "M",
    "sa": "SA", "sa1": "SA", "sa2": "SA", "sa3": "SA", "senior associate": "SA",
    "a2": "A2", "associate 2": "A2",
    "a1": "A1", "a": "A1", "a3": "A1", "associate": "A1",
}


def grade_key(v):
    s = str(v or "").strip()
    return ROLE_KEY.get(s.lower(), s)

# The dashboard (Directory, Skill Atlas, Capability Risk, Pulse, Team Analytics, Rate Analysis)
# is built by build_dark_dashboard.dark_html() and embedded in this same file — see dash_panel().
LIGHT_DASHBOARD_HREF = "../../03_Output files/Employee_Dashboard.html"   # the original light build
NO_STORE_META = ('<meta http-equiv="Cache-Control" content="no-store, no-cache, must-revalidate">'
                 '<meta http-equiv="Pragma" content="no-cache">')
TEAM_NAME = "Deals"
OTHER = "Other"          # the bucket for a blank or "Other..." competency or territory
TERR_SHORT = {"United States of America": "USA"}   # grid headers only; the full name shows on hover
# The hero mark: the practice stacked, one slab per grade, bottom tier first. Each slab is as
# wide as the people standing on it, so the widest tier is wherever the bench actually is — the
# shape is read, not labelled. Group grades together by listing them in one tier; empty ones drop.
# The Managing Director is not a slab: a small star floats above the top one (STAR_GRADE).
HERO_TIERS = [
    (("A1",), "#FFD26A"),
    (("A2",), "#FFB600"),
    (("SA",), "#EB8C00"),
    (("M",), "#FD5108"),
    (("SM",), "#E0301E"),
    (("Director",), "#E669A2"),
]
STAR_GRADE = "MD"
STACK_S = 46.0        # units per slab half-width
STACK_H = 13.0        # slab thickness
STACK_GAP = 3.0       # air between slabs, so each grade reads as its own layer
STACK_RY = -32.0      # the turn it rests at (degrees), and the one the SVG below is drawn at
STACK_RX = 27.0       # the tilt it rests at: how far you look down on it; a drag returns here
UTIL_HEADROOM = 1.6     # the Capacity chart's utilization scale tops out at least this far above the busiest month
SPARK_LABEL_GAP = 20.0   # a month's figure is written above its dot only when months are this far apart (SVG units)
STAR_R = 8.0          # the star's outer radius
STAR_LIFT = 20.0      # how high the star floats above the top slab, so it clears it at the resting tilt
LIGHT = (-0.35, 0.8, 0.5)   # where the light comes from, in view space: above, left, towards you
ORB_COUNT = 6          # skill orbs on the opening screen; they swap every 2–5 s (see BODY_JS)

COMPETENCY_NAMES = {   # Competency Filter value -> display name. Anything else is "Other".
    "FDD": "Financial Due Diligence",
    "Valuation": "Valuations",
    "Data & Analytics": "Data & Analytics",
    "BRS": "Business Restructuring Services",
}
COMPETENCY_CODE = {"Financial Due Diligence": "FDD", "Valuations": "VAL",
                   "Data & Analytics": "D&amp;A", "Business Restructuring Services": "BRS"}
GRADES = [  # most senior first
    ("MD", "Managing Director"), ("Director", "Director"), ("SM", "Senior Manager"),
    ("M", "Manager"), ("SA", "Senior Associate"), ("A2", "Associate 2"), ("A1", "Associate"),
]
GRADE_RANK = {g: i for i, (g, _) in enumerate(GRADES)}
GRADE_LABEL = dict(GRADES)
ACCENTS = ["#FD5108", "#EB8C00", "#E669A2", "#FFB600", "#E0301E"]
PHOTO_EXTS = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}
IMG_MIME = {**PHOTO_EXTS, ".svg": "image/svg+xml"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

e = html.escape


# ---------------------------------------------------------------- data
def norm_id(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") and s[:-2].isdigit() else s


def num(v):
    if v is None or v == "" or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "").rstrip("%"))
    except ValueError:
        return None


def comp_of(filter_value):
    """Competency Filter -> display name. Values not in COMPETENCY_NAMES keep their own name,
    so a new competency in the workbook appears without a code change. Blank and Other* are "Other"."""
    v = str(filter_value or "").strip()
    if not v or v.lower().startswith("other") or v.lower() == "none":
        return "Other"
    return COMPETENCY_NAMES.get(v, v)


def comp_code(name):
    if name in COMPETENCY_CODE:
        return COMPETENCY_CODE[name]
    words = [w for w in re.split(r"[\s&/-]+", name) if w and w.lower() not in ("and", "of", "the")]
    return e("".join(w[0] for w in words[:4]).upper() if len(words) > 1 else name[:3].upper())


def sheet_rows(wb, name):
    if name not in wb.sheetnames:
        return []
    rows = wb[name].iter_rows(values_only=True)
    header, seen = [], Counter()
    for h in next(rows):            # repeated names made unique, the way generate_report.py does it
        h = str(h).strip() if h is not None else ""
        seen[h] += 1
        header.append(h if seen[h] == 1 or not h else ("PwC Experience" if h == "Experience" else f"{h} ({seen[h]})"))
    return [dict(zip(header, r)) for r in rows if any(c not in (None, "") for c in r)]


def pick(row, *names):
    idx = {re.sub(r"[^a-z0-9]", "", k.lower()): k for k in row}
    for n in names:
        k = idx.get(re.sub(r"[^a-z0-9]", "", n.lower()))
        if k is not None and row.get(k) not in (None, ""):
            return row[k]
    return None


def month_key(v):
    if isinstance(v, (datetime, date)):
        return f"{v.year:04d}-{v.month:02d}"
    s = str(v or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.match(r"([A-Za-z]{3,})[\s\-/]+(\d{4})", s)
    if m and m.group(1)[:3].title() in MONTHS:
        return f"{m.group(2)}-{MONTHS.index(m.group(1)[:3].title()) + 1:02d}"
    return s[:7]


def month_label(key):
    y, m = key.split("-")
    return f"{MONTHS[int(m) - 1]} '{y[2:]}"


DATE_FORMATS = ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d.%m.%Y",
                "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y")


def to_date(v):
    """A date cell as a date: a real date, an Excel serial number, or text in a common format
    (day-first before month-first, as dates are written here). None when it cannot be read."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        serial = float(v)
    except (TypeError, ValueError):
        serial = None
    if serial is not None and 20000 < serial < 80000:        # Excel's day count from 1899-12-30
        return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
    text = str(v or "").strip()
    for t in (text, text.split(" ")[0], text.split("T")[0]):
        for fmt in DATE_FORMATS:
            try:
                return datetime.strptime(t, fmt).date()
            except ValueError:
                pass
    return None


def read_workbook():
    if not WORKBOOK.exists():
        raise SystemExit(f"Could not find the workbook at {WORKBOOK}. Check the paths at the top of this script.")
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True, data_only=True)
    today = date.today()
    people = []
    for rec in sheet_rows(wb, "Employee Details"):
        if not str(rec.get("Name") or "").strip():
            continue
        rec["Name"] = str(rec["Name"]).strip()
        rec["Role"] = grade_key(rec.get("Role"))
        rec["_id"] = norm_id(rec.get("WorkdayID"))
        rec["_comp"] = comp_of(rec.get("Competency Filter"))
        rec["_exp"] = num(rec.get("Experience"))
        jd = to_date(rec.get("Join Date"))
        # time at PwC from the join date; without a readable one, the "PwC Experience" column
        rec["_pwc"] = (today - jd).days / 365.25 if jd else num(rec.get("PwC Experience"))
        people.append(rec)

    util = []   # one entry per hours row
    for r in sheet_rows(wb, UTIL_SHEET):
        wid = norm_id(pick(r, "Workday ID", "WorkdayID"))
        mk = month_key(pick(r, "EoM", "Month Year"))
        if not wid or not mk:
            continue
        util.append({"id": wid, "month": mk, "std": num(pick(r, "Standard Hours")),
                     "ch": num(pick(r, "Chargeable Hours")) or 0.0, "tr": num(pick(r, "Training Hours")) or 0.0,
                     "pct": num(pick(r, "Utilisation%", "Utilization")), "comp": comp_of(pick(r, "Competency Filter")),
                     "terr": str(pick(r, "Territory") or "").strip(),
                     "role": grade_key(pick(r, "EMP Designation", "Role"))})
    # Skills the way the dashboard reads them: the skills sheet, one or more rows per person, each
    # cell split on "," and ";", matched case-insensitively and shown as first written. Only when
    # the workbook has no skills sheet does the Skills column on Employee Details stand in.
    sheet = next((n for n in SKILLS_SHEETS if n in wb.sheetnames), None)
    skill_rows = sheet_rows(wb, sheet) if sheet else []
    split = lambda v: [s.strip() for s in re.split(r"[,;]", str(v or "")) if s.strip()]
    held = defaultdict(list)
    for r in skill_rows:
        held[norm_id(pick(r, "WorkdayID", "Workday ID"))].extend(split(r.get("Skills")))
    shown = {}
    for rec in people:
        mine = []
        for s in (held.get(rec["_id"], []) if sheet else split(rec.get("Skills"))):
            d = shown.setdefault(s.lower(), s)
            if d not in mine:
                mine.append(d)
        rec["_skills"] = mine
    return people, util, sheet_rows(wb, RATES_SHEET), skill_rows


class Delivery:
    """Utilization figures over the hours sheet, defined exactly as the dashboard Overview."""

    def __init__(self, util, roster_ids):
        self.util, self.roster = util, roster_ids
        self.months = sorted({u["month"] for u in util})
        self.latest = self.months[-1] if self.months else ""
        pm = defaultdict(float)
        for u in util:
            if u["std"]:
                pm[(u["id"], u["month"])] += u["std"]
        stds = sorted(v for v in pm.values() if v > 0)
        self.full_month = stds[min(len(stds) - 1, int(len(stds) * 0.75))] if stds else 0

    def __bool__(self):
        return bool(self.util)

    def figures(self, keys=None, field="comp", ids=None):
        """Figures over the hours rows whose `field` ("comp" or "terr") is in `keys`; everything when keys is None.
        With `ids`, only those people's rows (a Director's reporting line)."""
        rows = [u for u in self.util if (keys is None or u[field] in keys) and (ids is None or u["id"] in ids)]
        std = sum(u["std"] or 0 for u in rows)
        ch = sum(u["ch"] for u in rows)
        pcts = [u["pct"] for u in rows if u["pct"] is not None]
        avg = ch / std * 100 if std else (sum(pcts) / len(pcts) if pcts else None)
        spare = 0.0
        latest = defaultdict(lambda: [0.0, 0.0, None])       # person-month in the latest month
        for u in rows:
            if u["month"] == self.latest and u["id"] in self.roster:
                cur = latest[u["id"]]
                cur[0] += u["std"] or 0
                cur[1] += u["ch"]
                cur[2] = u["pct"]
        for s, c, p in latest.values():
            if s > 0 and self.full_month > 0:
                spare += max(0.0, s - c) / self.full_month
            elif p is not None:
                spare += max(0.0, 100 - p) / 100
        series = []
        for m in self.months:
            ms = sum(u["std"] or 0 for u in rows if u["month"] == m)
            mc = sum(u["ch"] for u in rows if u["month"] == m)
            series.append(mc / ms * 100 if ms else None)
        return {"util": avg, "ch": ch, "tr": sum(u["tr"] for u in rows), "spare": spare, "series": series}

    def per_person(self, keys, field="comp"):
        """id -> [standard hrs, chargeable hrs, spare FTE in the latest month] over the same rows as
        figures(keys), so a slice of people can be summed in the page when the filters narrow it."""
        out = defaultdict(lambda: [0.0, 0.0, 0.0])
        latest = defaultdict(lambda: [0.0, 0.0, None])
        for u in self.util:
            if u[field] not in keys:
                continue
            out[u["id"]][0] += u["std"] or 0
            out[u["id"]][1] += u["ch"]
            if u["month"] == self.latest and u["id"] in self.roster:
                cur = latest[u["id"]]
                cur[0] += u["std"] or 0
                cur[1] += u["ch"]
                cur[2] = u["pct"]
        for i, (s, c, p) in latest.items():
            if s > 0 and self.full_month > 0:
                out[i][2] = max(0.0, s - c) / self.full_month
            elif p is not None:
                out[i][2] = max(0.0, 100 - p) / 100
        return out

    def person_months(self):
        """(id, month) -> [standard, chargeable, reported %]"""
        pm = defaultdict(lambda: [0.0, 0.0, None])
        for u in self.util:
            cur = pm[(u["id"], u["month"])]
            cur[0] += u["std"] or 0
            cur[1] += u["ch"]
            if u["pct"] is not None:
                cur[2] = u["pct"]
        return pm

    @staticmethod
    def pct_of(cell):
        if cell is None:
            return None
        std, ch, pct = cell
        return ch / std * 100 if std > 0 else pct

    def bands(self, ids):
        """Latest-month utilization bands and a 3-month burnout watch, over `ids` (MDs excluded by the caller)."""
        pm = self.person_months()
        last3 = self.months[-3:]
        vals = [v for v in (self.pct_of(pm.get((i, self.latest))) for i in ids) if v is not None]
        sustained = 0
        for i in ids:
            got = [self.pct_of(pm.get((i, m))) for m in last3]
            got = [g for g in got if g is not None]
            if len(got) == len(last3) and last3 and sum(got) / len(got) > 100:
                sustained += 1
        return {"n": len(vals), "bench": sum(1 for v in vals if v < 40), "healthy": sum(1 for v in vals if 80 <= v <= 100),
                "hot": sum(1 for v in vals if v > 100), "burn": sustained,
                "latest": sum(v for v in vals) / len(vals) if vals else None}

    def spare_by_person(self):
        pm = self.person_months()
        out = {}
        for (i, m), (std, ch, pct) in pm.items():
            if m != self.latest or i not in self.roster:
                continue
            if std > 0 and self.full_month > 0:
                out[i] = max(0.0, std - ch) / self.full_month
            elif pct is not None:
                out[i] = max(0.0, 100 - pct) / 100
        return out

    def by_month(self):
        rows = []
        for m in self.months:
            std = sum(u["std"] or 0 for u in self.util if u["month"] == m)
            ch = sum(u["ch"] for u in self.util if u["month"] == m)
            tr = sum(u["tr"] for u in self.util if u["month"] == m)
            rows.append({"m": m, "std": std, "ch": ch, "tr": tr, "util": ch / std * 100 if std else None})
        return rows

    def value(self, rate_of):
        """Chargeable hours priced through the rate card; also what the unused standard hours would be worth."""
        billed = unbilled = hours = 0.0
        by_grade = defaultdict(float)
        by_month = defaultdict(float)
        for u in self.util:
            rate = rate_of(u["role"], u["terr"])
            if not rate:
                continue
            billed += u["ch"] * rate
            hours += u["ch"]
            by_grade[u["role"]] += u["ch"] * rate
            by_month[u["month"]] += u["ch"] * rate
            if u["std"]:
                unbilled += max(0.0, u["std"] - u["ch"]) * rate
        return {"billed": billed, "unbilled": unbilled, "hours": hours, "people": len({u["id"] for u in self.util}),
                "by_grade": by_grade, "by_month": [by_month.get(m, 0.0) for m in self.months]}


def rate_lookup(rates):
    """rate_of(grade code, territory) from the Hourly Rates card, falling back to the Standard territory."""
    table = {}
    for r in rates:
        table[(str(r.get("Role") or "").strip(), str(r.get("Territory") or "").strip())] = num(r.get("Hourly Rate (USD)"))

    def rate_of(role, terr):
        name = RATE_ROLE.get(role, role)
        return table.get((name, terr)) or table.get((name, "Standard")) or 0
    return rate_of


def data_uri(path):
    return f"data:{IMG_MIME[path.suffix.lower()]};base64," + base64.b64encode(path.read_bytes()).decode("ascii")


def photo_for(rec):
    for key in ("Employee ID", "WorkdayID"):
        nid = norm_id(rec.get(key))
        for ext in PHOTO_EXTS:
            p = IMAGES_DIR / f"{nid}{ext}"
            if nid and p.exists():
                return data_uri(p)
    return ""


def ranked(counter, n=None):
    """`most_common`, with ties broken by name.

    Counts are built from sets (`set(p["_skills"])`), and set iteration order changes between
    runs, so plain `most_common()` shuffles equally-common items and two rebuilds of the same
    workbook produce different files. Sorting the ties makes a rebuild byte-for-byte comparable.
    """
    items = sorted(counter.items(), key=lambda kv: (-kv[1], str(kv[0])))
    return items[:n] if n else items


def initials(name):
    parts = name.split()
    return (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()


def team_size(name, children):
    seen, stack = set(), list(children.get(name, []))
    while stack:
        n = stack.pop()
        if n not in seen:
            seen.add(n)
            stack.extend(children.get(n, []))
    return len(seen)


def count_word(n):
    words = ["No", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"]
    return words[n] if 0 <= n < len(words) else str(n)


def fmt(n):
    return f"{int(round(n)):,}"


def mean(vals):
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else None


def rank(rec):
    return GRADE_RANK.get(rec["Role"], len(GRADES))


def is_other_territory(t):
    return str(t).strip().lower().startswith("other")


# ---------------------------------------------------------------- sections
def nav(logo_uri):
    logo = f'<img src="{logo_uri}" alt="PwC">' if logo_uri else "<b>PwC</b>"
    return f"""
<div class="progress" aria-hidden="true"><i></i></div>
<header class="top" id="top">
  <div class="top-in">
    <a class="brand" href="#hero"><span class="logo">{logo}</span><span>{TEAM_NAME} Skills &amp; Bio</span></a>
    <nav class="links" aria-label="Sections">
      <a href="#competencies">Competencies</a><a href="#leadership">Leadership</a><a href="#capacity">Capacity</a><a href="#shape">Team shape</a><a href="#value">Value</a>
    </nav>
    <button type="button" class="btn btn-sm" data-dash>Open dashboard</button>
  </div>
</header>"""


MARQUEE_MAX_CHARS = 15   # skills this long or longer stay out of the scrolling strip (the tabs list them all)
MARQUEE_SEC_PER_CHIP = 3.2  # the strip's pace; its duration grows with the number of chips so speed stays even


def marquee(skill_counts):
    items = [(s, c) for s, c in ranked(skill_counts) if len(s) < MARQUEE_MAX_CHARS]
    half = (len(items) + 1) // 2
    rows = []
    for i, chunk in enumerate((items[:half], items[half:])):
        chips = "".join(f'<span class="chip">{e(s)}<sup>{c}</sup></span>' for s, c in chunk)
        dur = max(40, len(chunk) * MARQUEE_SEC_PER_CHIP) * (1.15 if i else 1)
        rows.append(f'<div class="mq-row{" rev" if i else ""}"><div class="mq-track" style="animation-duration:{dur:.0f}s">'
                    f'{chips}<span class="dup" aria-hidden="true">{chips}</span></div></div>')
    return f"""
<section class="marquee" id="skills" aria-label="Skills across the team">
  <p class="mq-label">Skills on the team &middot; number of people who hold each</p>
  {''.join(rows)}
</section>"""


def statement(n_people, n_comp, avg_exp, n_terr):
    exp = f"{avg_exp:.1f}" if avg_exp is not None else "\u2014"
    text = (f"We are [{n_people}] {TEAM_NAME} professionals across [{n_comp}] competencies, "
            f"averaging [{exp}] years of experience each, serving clients in [{n_terr}] territories.")
    words = []
    for w in text.split():
        if w.startswith("[") and w.endswith("]"):
            words.append(f'<span class="w num">{e(w[1:-1])}</span>')
        else:
            words.append(f'<span class="w">{e(w)}</span>')
    return f"""
<section class="statement" id="story">
  <div class="wrap"><p class="big">{' '.join(words)}</p></div>
</section>"""


def kpi(value, label, note="", dec=0, suffix="", count=True):
    """One headline figure. data-count lets the page count up to it; the text is already final."""
    shown = f"{value:,.{dec}f}{suffix}" if isinstance(value, (int, float)) else e(str(value))
    attr = f' data-count="{value:.{dec}f}" data-dec="{dec}" data-suffix="{e(suffix)}"' if count and isinstance(value, (int, float)) else ""
    return (f'<div class="kp rv"><b{attr}>{shown}</b><span>{e(label)}</span>'
            + (f'<small>{e(note)}</small>' if note else "") + "</div>")


def stagger(tiles, step=110):
    """Each tile arrives a beat after the one before it, left to right."""
    return "".join(t.replace('class="kp rv"', f'class="kp rv" style="--d:{i * step}ms"', 1) for i, t in enumerate(tiles))


def kpis(row1, row2):
    return f"""
<section class="kpis-sec" id="numbers">
  <div class="wrap">
    <div class="kp-row"><p class="kp-h rv">Our people</p><div class="kp-grid">{stagger(row1)}</div></div>
    <div class="kp-row"><p class="kp-h rv">How we deliver</p><div class="kp-grid">{stagger(row2)}</div></div>
  </div>
</section>"""


def sparkline(series, months, col, w=320, h=70, caption=True, labels=True):
    """Monthly utilization as a line with a dot on every month, and each month's figure above its
    dot when the months sit far enough apart to be read (SPARK_LABEL_GAP)."""
    pts = [(i, v) for i, v in enumerate(series) if v is not None]
    if len(pts) < 2:
        return ""
    vals = [v for _, v in pts]
    if caption:                      # panel: 0-based scale with a 100% reference
        lo, hi = 0, max(120, math.ceil(max(vals) / 20) * 20)
    else:                            # card: fitted to the data so the shape reads in a small space
        pad = max(4, (max(vals) - min(vals)) * 0.15)
        lo, hi = min(vals) - pad, max(vals) + pad
    step = (w - 8) / max(len(series) - 1, 1)
    labels = labels and step >= SPARK_LABEL_GAP
    head = 13 if labels else 4       # room above the line for the figures
    X = lambda i: 4 + i * step
    Y = lambda v: h - 4 - (v - lo) / (hi - lo) * (h - 4 - head)
    line = " ".join(f"{X(i):.1f},{Y(v):.1f}" for i, v in pts)
    last_i, last_v = pts[-1]
    dots = "".join(f'<circle cx="{X(i):.1f}" cy="{Y(v):.1f}" r="{3.5 if i == last_i else 2.2}" fill="{col}"/>' for i, v in pts)
    text = ""
    if labels:
        for i, v in pts:
            anchor = "start" if i == 0 else "end" if i == len(series) - 1 else "middle"
            text += (f'<text x="{X(i) - (2 if anchor == "end" else 0):.1f}" y="{Y(v) - 6:.1f}" text-anchor="{anchor}" '
                     f'class="sp-lb{" last" if i == last_i else ""}">{v:.0f}%</text>')
    svg = (f'<svg class="spark" viewBox="0 0 {w} {h}" role="img" '
            f'aria-label="Monthly utilization {month_label(months[0])} to {month_label(months[-1])}">'
            + (f'<line x1="0" x2="{w}" y1="{Y(100):.1f}" y2="{Y(100):.1f}" class="sp-ref"/>' if lo <= 100 <= hi else "") +
            f'<polyline points="{line}" fill="none" stroke="{col}" stroke-width="2.2" stroke-linejoin="round" stroke-linecap="round"/>'
            f'{dots}{text}</svg>')
    if not caption:
        return svg
    return (svg + f'<div class="sp-t"><span>{month_label(months[0])}</span><span>dashed line = 100%</span>'
            f'<span>{month_label(months[-1])} &middot; <b>{last_v:.0f}%</b></span></div>')


def delivery_strip(f):
    if not f:
        return ""
    util = f"{f['util']:.0f}%" if f["util"] is not None else "—"
    return (f'<div class="dstrip">'
            f'<div><b>{util}</b><span>utilization</span></div>'
            f'<div><b>{fmt(f["ch"])}</b><span>chargeable hrs</span></div>'
            f'<div><b>{f["spare"]:.1f}</b><span>spare FTE</span></div>'
            f'<div><b>{fmt(f["tr"])}</b><span>training hrs</span></div></div>')


LEVELS = [  # grade rows, most senior first. Group grades together by listing them in one row.
    ("Dir", "MD & Directors", ("MD", "Director")),
    ("SM", "Senior Managers", ("SM",)), ("M", "Managers", ("M",)),
    ("SA", "Senior Associates", ("SA",)), ("A", "Associates", ("A1", "A2")),
]

# Team shape reads at a finer grain than the competency cards: same groups, with the two
# Associate grades on their own rows. Both lists must cover every grade in GRADES.
SHAPE_LEVELS = [
    ("Dir", "MD & Directors", ("MD", "Director")),
    ("SM", "Senior Managers", ("SM",)), ("M", "Managers", ("M",)),
    ("SA", "Senior Associates", ("SA",)),
    ("A2", "Associate 2", ("A2",)), ("A1", "Associate", ("A1",)),
]


def pyramid(members):
    """One bar per row of LEVELS, most junior at the bottom, centred so the team's shape reads at a glance."""
    counts = [(short, full, sum(1 for m in members if m["Role"] in roles)) for short, full, roles in LEVELS]
    known = {r for _, _, roles in LEVELS for r in roles}
    other = sum(1 for m in members if m["Role"] not in known)
    if other:
        counts.append(("Oth", "Other grades", other))
    counts = [c for c in counts if c[2]]      # a grade nobody holds is a row of nothing to read
    top = max((c for _, _, c in counts), default=0) or 1
    rows = "".join(
        f'<div class="py2" title="{e(full)}: {c}"><span>{short}</span><span class="py2-b"><i style="width:{c / top * 100:.1f}%"></i></span><b>{c}</b></div>'
        for short, full, c in counts)
    return f'<div class="pyr" role="img" aria-label="Headcount by grade">{rows}</div>'


def group_card(i, anchor, name, members, f, deliv, lead_html, meta):
    """One card in the Competencies section — used for both the competency and the territory view."""
    col = ACCENTS[i % len(ACCENTS)]
    sk = ranked(Counter(s for m in members for s in set(m["_skills"])), 5)
    chips = "".join(f'<span class="chip">{e(s)}<sup>{c}</sup></span>' for s, c in sk)
    trend = ""
    if f:
        spark = sparkline(f["series"], deliv.months, col, w=300, h=52, caption=False)
        # the header figure is the whole period's average for this competency or territory (total
        # chargeable over total standard hours, the same figure as the utilization tile), not the
        # last month: that one is already the last point on the line
        if spark and f["util"] is not None:
            trend = (f'<div class="trend"><div class="tr-h"><span class="k">Utilization trend</span>'
                     f'<span>{month_label(deliv.months[0])} &rarr; {month_label(deliv.latest)} &middot; avg <b>{f["util"]:.0f}%</b></span></div>{spark}</div>')
    meta_html = "".join(f"<div><dt>{e(k)}</dt><dd>{v}</dd></div>" for k, v in meta)
    return f"""
      <article class="comp rv" id="{anchor}" style="--c:{col};--d:{(i % 2) * 110}ms">
        <div class="comp-head">
          <div class="ch-l"><h3>{e(name)}</h3><p class="lead">{lead_html}</p></div>
          <div class="ch-r"><p class="hc"><b data-count="{len(members)}">{len(members)}</b> people</p>{pyramid(members)}</div>
        </div>
        {delivery_strip(f)}
        {trend}
        <p class="k">Deepest skills</p>
        <div class="chips">{chips}</div>
        <dl class="meta">{meta_html}</dl>
      </article>"""


def competencies(comps, deliv, comp_leads, people):
    """Competency cards, with a toggle to the same cards cut by client territory."""
    fy = f" &middot; FY {month_label(deliv.months[0])} &ndash; {month_label(deliv.latest)}" if deliv else ""
    places = lambda ms: e(", ".join(sorted({str(m.get("Location") or "").strip() for m in ms} - {""})))

    comp_cards = []
    for i, (name, members) in enumerate(comps):
        leads = comp_leads.get(name) or []
        lead = ("Led by " + ", ".join(f"<b>{e(d['Name'])}</b>" for d in leads) + ", Director" + ("s" if len(leads) > 1 else "")
                if leads else "Guided directly by the Managing Director")
        terrs = e(", ".join(sorted({str(m.get("Territory") or "").strip() for m in members} - {""})))
        comp_cards.append(group_card(i, f"c{i}", name, members, deliv.figures({name}) if deliv else None, deliv, lead,
                                     [("Offices", places(members)), ("Client territories", terrs)]))

    by_terr = defaultdict(list)
    for p in people:
        t = str(p.get("Territory") or "").strip()
        if t and not is_other_territory(t):
            by_terr[t].append(p)
    terrs = sorted(by_terr.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    terr_cards = []
    for i, (name, members) in enumerate(terrs):
        mix = ranked(Counter(m["_comp"] for m in members if m["_comp"] != "Other"))
        lead = f"Largest: <b>{e(mix[0][0])}</b>" if mix else "No mapped competency"
        codes = " &middot; ".join(f"{comp_code(c)} {n}" for c, n in mix) or "—"
        terr_cards.append(group_card(i, f"t{i}", name, members, deliv.figures({name}, "terr") if deliv else None, deliv, lead,
                                     [("Offices", places(members)), ("Competencies", codes)]))

    n_c, n_t = len(comps), len(terrs)
    return f"""
<section class="competencies" id="competencies">
  <div class="wrap">
    <div class="comp-intro">
      <div>
        <p class="eyebrow rv">Competencies &amp; territories</p>
        <h2 class="rv"><span class="cv-h" data-v="comp">{count_word(n_c)} competenc{"ies" if n_c != 1 else "y"}. One bench of expertise.</span><span class="cv-h" data-v="terr">{count_word(n_t)} client territor{"ies" if n_t != 1 else "y"}. One practice.</span></h2>
        <p class="sub rv"><span class="cv-h" data-v="comp">How each competency is led, staffed and deployed{fy}.</span><span class="cv-h" data-v="terr">How each client territory is staffed and deployed{fy}.</span></p>
      </div>
      <div class="tg rv" role="group" aria-label="Group cards by">
        <button type="button" data-v="comp" aria-pressed="true">Competency</button><button type="button" data-v="terr" aria-pressed="false">Territory</button>
      </div>
    </div>
    <div class="cv" data-v="comp"><div class="comp-grid">{''.join(comp_cards)}</div></div>
    <div class="cv" data-v="terr"><p class="k cv-nojs">By client territory</p><div class="comp-grid">{''.join(terr_cards)}</div></div>
  </div>
</section>"""


def money(v):
    if v >= 1_000_000:
        return f"${v / 1_000_000:.1f}M"
    if v >= 1_000:
        return f"${v / 1_000:.0f}k"
    return f"${v:.0f}"


def stat_tiles(items, cols=4):
    return f'<div class="tiles" style="--tc:{cols}">' + "".join(
        f'<div class="tile rv" style="--d:{i*60}ms"><b>{v}</b><span>{e(label)}</span>'
        + (f"<small>{e(note)}</small>" if note else "") + "</div>"
        for i, (v, label, note) in enumerate(items)) + "</div>"


def fits(text, column, font=12):
    """Whether a label is narrow enough to sit over its bar without touching the next one."""
    return len(text) * font * 0.62 <= column * 0.9


def capacity(deliv, people):
    """Practice-level Pulse: the monthly picture and where people sit in the latest month."""
    if not deliv:
        return ""
    rows = deliv.by_month()
    band = deliv.bands([p["_id"] for p in people if p["Role"] != "MD"])   # counts name individuals, so MDs are out
    t = deliv.figures()
    # Each bar is the month's standard hours: chargeable at the bottom, non-chargeable (everything
    # else, training included) on top. Utilization is the chargeable share of that bar, so it is
    # written once, above the bar, rather than as a second line on its own scale.
    W, H, L, R, T, B = 1000, 300, 62, 62, 30, 40
    for r in rows:
        r["nc"] = max(0.0, (r["std"] or 0) - r["ch"])
    top_v = max((r["ch"] + r["nc"] for r in rows), default=1) or 1
    step = next(s_ for s_ in (250, 500, 1000, 2000, 2500, 5000, 10000, 20000, 25000, 50000, 100000)
                if top_v / s_ <= 4) if top_v <= 400000 else 10 ** len(str(int(top_v // 4)))
    hmax = math.ceil(top_v / step) * step
    ticks = [step * k for k in range(int(hmax / step) + 1)]
    tick = lambda v: f"{v / 1000:g}k" if v >= 1000 else f"{v:g}"
    n = len(rows)
    col = (W - L - R) / max(n, 1)
    bw = col * 0.56
    X = lambda i: L + (i + 0.5) * col
    Yh = lambda v: H - B - v / hmax * (H - T - B)
    # Utilization trend on its own light scale at the right. Its ticks sit on the same gridlines as
    # the hours, and the scale is roomy enough (UTIL_HEADROOM) that the line runs through the orange
    # part of the bars, clear of the hour figures, rather than over the grey segment's numbers.
    k = max(len(ticks) - 1, 1)
    u_top = max((r["util"] for r in rows if r["util"] is not None), default=100)
    u_step = next((s_ for s_ in (20, 25, 30, 40, 50, 60, 75, 100) if s_ * k >= u_top * UTIL_HEADROOM), 100)
    umax = u_step * k
    Yu = lambda v: H - B - v / umax * (H - T - B)
    grid = "".join(
        f'<line class="cx" x1="{L}" y1="{Yh(v):.1f}" x2="{W - R}" y2="{Yh(v):.1f}"/>'
        f'<text class="cy" x="{L - 8}" y="{Yh(v) + 4:.1f}" text-anchor="end">{tick(v)}</text>'
        f'<text class="cy cy-u" x="{W - R + 8}" y="{Yh(v) + 4:.1f}">{u_step * j:g}%</text>' for j, v in enumerate(ticks))
    bars, dl = "", ""
    for i, r in enumerate(rows):
        x0, y_ch, y_top = X(i) - bw / 2, Yh(r["ch"]), Yh(r["ch"] + r["nc"])
        bars += (f'<rect class="bch" x="{x0:.1f}" y="{y_ch:.1f}" width="{bw:.1f}" height="{max(H - B - y_ch, 0):.1f}" rx="3"/>'
                 f'<rect class="bnc" x="{x0:.1f}" y="{y_top:.1f}" width="{bw:.1f}" height="{max(y_ch - y_top - 2, 0):.1f}" rx="3"/>')
        # each segment's hours inside it when the segment is tall and wide enough; utilization on top
        for v, y0, y1, cls in ((r["ch"], y_ch, H - B, "dl-ch"), (r["nc"], y_top, y_ch, "dl-nc")):
            txt = f"{v:,.0f}"
            if v and y1 - y0 >= 20 and fits(txt, col, 11):
                dl += f'<text class="dl {cls}" x="{X(i):.1f}" y="{(y0 + y1) / 2 + 4:.1f}" text-anchor="middle">{txt}</text>'
        if r["util"] is not None and fits(f"{r['util']:.0f}%", col, 13):
            dl += f'<text class="dl dl-u" x="{X(i):.1f}" y="{y_top - 8:.1f}" text-anchor="middle">{r["util"]:.0f}%</text>'
    labels = "".join(
        f'<text class="cy{" alt" if i % 2 else ""}" x="{X(i):.1f}" y="{H - 12}" text-anchor="middle">{month_label(r["m"])}</text>'
        for i, r in enumerate(rows) if n <= 12 or i % 2 == 0)
    pts = [(X(i), Yu(r["util"])) for i, r in enumerate(rows) if r["util"] is not None]
    trend = (f'<g class="ut"><polyline class="ut-l" points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in pts)}" fill="none"/>'
             + "".join(f'<circle class="ut-d" cx="{x:.1f}" cy="{y:.1f}" r="3.6"/>' for x, y in pts) + "</g>") if len(pts) > 1 else ""
    chart = (f'<svg class="chart rv" viewBox="0 0 {W} {H}" role="img" aria-label="Chargeable and non-chargeable hours by month, with utilization">'
             f'{grid}{bars}{trend}{dl}{labels}</svg>')
    lm = month_label(deliv.latest)
    tiles = [(f"{t['util']:.0f}%" if t["util"] is not None else "—", "Average utilization", f"FY {month_label(deliv.months[0])} – {lm}"),
             (f"{band['latest']:.0f}%" if band["latest"] is not None else "—", "Latest month", lm),
             (band["bench"], "On the bench", f"under 40% · {lm}"),
             (band["healthy"], "Healthy band", f"80–100% · {lm}"),
             (band["hot"], "Stretched", f"over 100% · {lm}"),
             (band["burn"], "Burnout watch", "3 months over 100%"),
             (f"{t['spare']:.1f}", "Spare capacity (FTE)", lm),
             (fmt(t["tr"]), "Training hours", "FY")]
    return f"""
<section class="capacity" id="capacity">
  <div class="wrap">
    <div class="sec-head">
      <p class="eyebrow rv">Capacity</p>
      <h2 class="rv">How busy the practice is.</h2>
      <p class="sub rv">Each month's standard hours, split into chargeable and non-chargeable (training included), with utilization on top. Band counts are people in {e(lm)}, Managing Directors excluded.</p>
    </div>
    <div class="sh-card">
      <div class="lgd"><span class="lg lg-ch">Chargeable hours</span><span class="lg lg-nc">Non-chargeable hours</span><span class="lg lg-ut">Utilization trend (right scale)</span><span class="lg lg-pct"><b>%</b> utilization, the chargeable share</span></div>
      {chart}
    </div>
    {stat_tiles(tiles)}
  </div>
</section>"""


def value_section(deliv, rate_of, people):
    """Rate Analysis at practice level — indicative value, not booked revenue."""
    if not deliv:
        return ""
    v = deliv.value(rate_of)
    if not v["billed"]:
        return ""
    blended = v["billed"] / v["hours"] if v["hours"] else 0
    rows = [(full, sum(v["by_grade"].get(g, 0) for g in roles)) for _, full, roles in LEVELS]
    known = {g for _, _, roles in LEVELS for g in roles}
    rest = sum(x for g, x in v["by_grade"].items() if g not in known)
    if rest:
        rows.append(("Other grades", rest))
    rows = [(g, x) for g, x in rows if x]
    mx = max((x for _, x in rows), default=1) or 1
    bars_html = "".join(
        f'<div class="off rv" style="--d:{i*50}ms"><span>{e(g)}</span><span class="off-b"><i style="width:{x / mx * 100:.0f}%"></i></span>'
        f'<b>{money(x)}</b><span class="pr-s">{x / v["billed"] * 100:.0f}%</span></div>' for i, (g, x) in enumerate(rows))
    best = max(zip(deliv.months, v["by_month"]), key=lambda kv: kv[1], default=("", 0))
    senior = sum(v["by_grade"].get(g, 0) for g in ("MD", "Director", "SM", "M"))
    top_grade = max(rows, key=lambda kv: kv[1], default=("—", 0))
    tiles = [(money(v["billed"]), "Billable value", "chargeable hours at card rate"),
             (f"${blended:.0f}", "Blended rate / hour", "across every grade"),
             (money(v["unbilled"]), "Unbilled at card rate", "the spare capacity, priced"),
             (money(v["billed"] / max(v["people"], 1)), "Value per person", f"{v['people']} people who worked in the period")]
    months = deliv.months
    W, H, B = 600, 200, 26
    vmax = max(v["by_month"] + [1])
    bw = (W - 8) / max(len(months), 1) * 0.58
    X = lambda i: 4 + (i + 0.5) / max(len(months), 1) * (W - 8)
    Y = lambda x: H - B - x / vmax * (H - B - 24)
    mbars = "".join(
        f'<rect class="bch" x="{X(i) - bw / 2:.1f}" y="{Y(x):.1f}" width="{bw:.1f}" height="{max(H - B - Y(x), 0):.1f}" rx="3"/>'
        for i, x in enumerate(v["by_month"]))
    mlabels = "".join(
        f'<text class="cy{" alt" if i % 2 else ""}" x="{X(i):.1f}" y="{H - 8}" text-anchor="middle">{month_label(m)}</text>'
        for i, m in enumerate(months))
    col = (W - 8) / max(len(months), 1)
    mlabels += "".join(                      # each month's value over its bar, where it fits
        f'<text class="dl dl-v" x="{X(i):.1f}" y="{Y(x) - 6:.1f}" text-anchor="middle">{money(x)}</text>'
        for i, x in enumerate(v["by_month"]) if x and fits(money(x), col, 11))
    chart = (f'<svg class="chart rv" viewBox="0 0 {W} {H}" role="img" aria-label="Billable value by month">'
             f'<line class="cx" x1="0" x2="{W}" y1="{H - B:.1f}" y2="{H - B:.1f}"/>{mbars}{mlabels}</svg>')
    return f"""
<section class="value-sec" id="value">
  <div class="wrap">
    <div class="sec-head">
      <p class="eyebrow rv">Value</p>
      <h2 class="rv">What the hours are worth.</h2>
      <p class="sub rv">Chargeable hours priced through the Hourly Rates card by grade and territory. Indicative value, not booked revenue.</p>
    </div>
    {stat_tiles(tiles)}
    <div class="shape-grid">
      <div class="sh-card"><p class="k">Billable value by grade</p><div class="val-rows">{bars_html}</div>
        <p class="lev"><b>{e(top_grade[0])}</b> generate the most, at {top_grade[1] / v["billed"] * 100:.0f}% of the value.
        Manager grade and above: {senior / v["billed"] * 100:.0f}%.</p></div>
      <div class="sh-card"><p class="k">Billable value by month</p><div class="val-rows">{chart}</div>
        <p class="lev">Best month <b>{e(month_label(best[0]))}</b> at {money(best[1])}. Average {money(sum(v["by_month"]) / max(len(v["by_month"]), 1))} a month.</p></div>
    </div>
  </div>
</section>"""


def polar(r, deg):
    return 50 + r * math.cos(math.radians(deg)), 50 + r * math.sin(math.radians(deg))


def bars(counter, order=None):
    items = [(k, counter[k]) for k in order if counter.get(k)] if order else ranked(counter)
    if not items:
        return ""
    m = max(v for _, v in items)
    return '<ul class="mb">' + "".join(
        f'<li><span>{e(k)}</span><span class="bar"><i style="width:{v / m * 100:.1f}%"></i></span><b>{v}</b></li>'
        for k, v in items) + "</ul>"


SENIOR = ("MD", "Director", "SM")   # "Senior Manager and above", for the leadership orbit


def line_of(name, children):
    """Everyone under a person through the reporting lines, direct and indirect, nearest first."""
    seen, out, queue = {name}, [], list(children.get(name, []))
    while queue:
        n = queue.pop(0)
        if n not in seen:
            seen.add(n)
            out.append(n)
            queue.extend(children.get(n, []))
    return out


def leadership(mds, directors, comps, comp_color, comp_leads, children, photos, deliv, skill_cls, mark, people):
    """MD at the centre, Directors on one ring with their competencies under their names, and on an
    outer layer the competencies no Director leads, with their Senior Managers circling them.
    The panel on the left is about each Director's own team: everyone in their reporting line."""
    if not directors and not mds:
        return ""
    members = dict(comps)
    by_name = {p["Name"]: p for p in people}
    n_dir = len(directors)
    d_ang = {d["Name"]: -90 + i * 360 / max(n_dir, 1) for i, d in enumerate(directors)}

    def team_of(d):
        """(everyone in their line, delivery figures over that line)"""
        names = line_of(d["Name"], children)
        team = [by_name[n] for n in names if n in by_name]
        f = deliv.figures(ids={p["_id"] for p in team}) if deliv and team else None
        return team, f

    # Inner ring: Directors only, each with the competencies they lead written under their name.
    lines, nodes = [], []
    for i, d in enumerate(directors):
        a = d_ang[d["Name"]]
        x, y = polar(34, a)
        col = comp_color.get(d["_comp"], ACCENTS[0])
        led = [n for n, _ in comps if d in (comp_leads.get(n) or [])]
        hc = sum(len(members[n]) for n in led)
        full = "".join(f"<i>{e(n)}</i>" for n in led) or "<i>No competency mapped</i>"
        short = " &middot; ".join(comp_code(n) for n in led)
        lines.append(f'<line class="ln dash" data-k="d{i}" x1="500" y1="500" x2="{x*10:.1f}" y2="{y*10:.1f}"/>')
        nodes.append(
            f'<a class="dn" href="#lead-d{i}" data-go="d{i}" data-k="d{i}" style="--x:{x:.2f}%;--y:{y:.2f}%;--c:{col};--d:{i*110}ms" '
            f'aria-label="Open {e(d["Name"])}">{avatar(d, photos)}'
            f'<span class="pl"><b>{e(d["Name"])}</b>'
            f'<span class="cl-full">{full}</span><span class="cl-code">{short}</span>'
            f'<em>{hc} people</em></span></a>')

    # Outer layer: competencies no Director leads, shown only when they have a Senior Manager or
    # above, with those Senior Managers circling the competency. Teams with no one that senior
    # (support functions and the like) are left off the leadership picture.
    rest = []
    for j, (n, m) in enumerate(comps):
        sms = sorted((p for p in m if p["Role"] in SENIOR), key=lambda p: (rank(p), p["Name"]))
        if not comp_leads.get(n) and sms:
            rest.append((j, n, m, sms))
    k = len(rest)
    for q, (j, name, mem, sms) in enumerate(rest):
        # in the gaps between Directors while they fit, otherwise evenly round the ring
        if n_dir and k <= n_dir:
            slot = round(q * n_dir / k)
            a = -90 + 180 / n_dir + slot * 360 / n_dir
        else:
            a = -90 + 180 / k + q * 360 / k
        x, y = polar(47, a)
        lines.append(f'<line class="ln dot" data-k="md" x1="500" y1="500" x2="{x*10:.1f}" y2="{y*10:.1f}"/>')
        sats = "".join(f'<span class="cn-sat" style="--a:{s * 360 / len(sms):.1f}deg" title="{e(p["Name"])} &middot; '
                       f'{e(GRADE_LABEL.get(p["Role"], p["Role"]))}">{avatar(p, photos)}</span>'
                       for s, p in enumerate(sms))
        n_sm = sum(1 for p in sms if p["Role"] == "SM")
        who = f'{n_sm} Senior Manager{"s" if n_sm != 1 else ""}' if n_sm == len(sms) else f'{len(sms)} senior leaders'
        nodes.append(
            f'<a class="cn" href="#c{j}" data-k="md" title="{e(name)} &middot; {e(", ".join(p["Name"] for p in sms))}" '
            f'style="--x:{x:.2f}%;--y:{y:.2f}%;--c:{comp_color[name]};--d:{400 + q*90}ms" aria-label="{e(name)}">'
            f'<span class="cn-orb" style="--spin:{22 + q * 4}s">{sats}</span>'
            f'<span class="cn-b">{comp_code(name)}</span>'
            f'<span class="cn-n"><b>{who}</b>{len(mem)} people</span></a>')
    legend = ""
    if rest:
        legend = ('<div class="orbit-legend"><p class="k">Outer ring &middot; led by Senior Managers, guided by the Managing Director</p><div class="chips">'
                  + "".join(f'<a class="chip" href="#c{j}"><i class="dot" style="background:{comp_color[n]}"></i><b>{comp_code(n)}</b>&nbsp;{e(n)}<sup>{len(m)}</sup></a>'
                            for j, n, m, _ in rest) + "</div></div>")

    hub_av = "".join(avatar(m, photos) for m in mds[:3])
    md_names = " &amp; ".join(e(m["Name"]) for m in mds[:3])
    hub = (f'<a class="hub md" href="#lead-all" data-go="all" aria-label="Leadership overview">'
           f'<span class="md-avs">{hub_av}</span>'
           f'<span class="pl"><b class="md-name">{md_names}</b>Managing Director{"s" if len(mds) > 1 else ""}</span></a>') if mds else \
          f'<div class="hub">{mark}</div>'

    # ---- panels: each Director's own team
    teams = [team_of(d) for d in directors]
    picks = []
    for i, d in enumerate(directors):
        team, f = teams[i]
        led = [n for n, _ in comps if d in (comp_leads.get(n) or [])]
        util = f"{f['util']:.0f}%" if f and f["util"] is not None else "—"
        picks.append(
            f'<a class="pick" href="#lead-d{i}" data-go="d{i}" style="--c:{comp_color.get(d["_comp"], ACCENTS[0])}">{avatar(d, photos)}'
            f'<span class="pn"><b>{e(d["Name"])}</b><span class="s">{e(", ".join(led) or "No competency mapped")}</span></span>'
            f'<span class="ps"><b>{len(team)}</b><small>in team</small><i>{f"{util} util." if team else "no reports"}</i></span></a>')
    md_line = (f"{md_names} guide{'s' if len(mds) == 1 else ''} the practice" if mds else "The practice")
    views = [f"""
      <div class="tm-view on" id="lead-all" data-view="all">
        <p class="eyebrow">Leadership</p>
        <h2>{md_line}.<br>{count_word(n_dir)} Director{"s" if n_dir != 1 else ""} run{"" if n_dir != 1 else "s"} the competencies.</h2>
        <p class="sub">The Managing Director sits at the centre, with each Director's competencies written under their name.
        The outer ring holds the competencies led by Senior Managers, who circle them. Below, each Director's own team:
        everyone in their reporting line, directly or through their managers, and that team's utilization. Click one to open it.</p>
        <div class="picks">{''.join(picks)}</div>
      </div>"""]

    for i, d in enumerate(directors):
        col = comp_color.get(d["_comp"], ACCENTS[0])
        led = [n for n, _ in comps if d in (comp_leads.get(n) or [])]
        team, f = teams[i]
        desc = str(d.get("Employee Description") or "").strip()
        sk = ranked(Counter(s for p in team for s in set(p["_skills"])), 10)

        def chip(s, c=None):
            ic = f'<i class="ci ic-{skill_cls[s.lower()]}"></i>' if s.lower() in skill_cls else ""
            sup = f"<sup>{c}</sup>" if c else ""
            return f'<span class="chip">{ic}{e(s)}{sup}</span>'

        if team:
            util = f"{f['util']:.0f}%" if f and f["util"] is not None else "—"
            exp_avg, pwc_avg = mean(p["_exp"] for p in team), mean(p["_pwc"] for p in team)
            tiles = [(len(team), "In their team"), (util, "Team utilization"),
                     (fmt(f["ch"]) if f else "—", "Chargeable hrs"),
                     (f"{f['spare']:.1f}" if f else "—", f"Spare FTE · {month_label(deliv.latest)}" if f else "Spare FTE"),
                     (f"{exp_avg:.1f}" if exp_avg is not None else "—", "Avg experience"),
                     (f"{pwc_avg:.1f}" if pwc_avg is not None else "—", "Avg yrs at PwC"),
                     (fmt(f["tr"]) if f else "—", "Training hrs")]
            body = f"""
        <p class="k">Their team &middot; everyone who reports to {e(d["Name"].split()[0])}, directly or through their managers</p>
        <div class="kpis">{''.join(f'<div><b>{v}</b><span>{l}</span></div>' for v, l in tiles)}</div>
        {f'<p class="k">Monthly utilization &middot; their team</p>{sparkline(f["series"], deliv.months, col, h=84)}' if f else ""}
        <div class="an">
          <div><p class="k">Grade mix</p>{bars(Counter(p["Role"] for p in team), [g for g, _ in GRADES] + sorted({p["Role"] for p in team} - set(GRADE_RANK)))}</div>
          <div><p class="k">Offices</p>{bars(Counter(str(p.get("Location") or "Unassigned") for p in team))}</div>
          <div><p class="k">Client territories</p>{bars(Counter(str(p.get("Territory") or "Unassigned") for p in team))}</div>
          <div><p class="k">Top skills</p><div class="chips">{''.join(chip(s, c) for s, c in sk)}</div></div>
        </div>"""
        else:
            body = f'<p class="sub">No one reports to {e(d["Name"])} in the workbook, so there is no team to summarise here.</p>'
        views.append(f"""
      <div class="tm-view" id="lead-d{i}" data-view="d{i}" style="--c:{col}">
        <a class="back" href="#lead-all" data-go="all">&larr; Leadership overview</a>
        <div class="tm-head">{avatar(d, photos)}<div><h3>{e(d["Name"])}</h3>
          <p>Director &middot; {e(str(d.get("Location") or ""))} &middot; leads <b>{e(", ".join(led) or "—")}</b></p></div></div>
        {f"<blockquote>{e(desc)}</blockquote>" if desc else ""}
        {body}
      </div>""")

    return f"""
<section class="teams" id="leadership">
  <div class="wrap tm-grid">
    <div class="tm-panel">{''.join(views)}</div>
    <div class="tm-orbit-col"><div class="orbit rv">
      <svg class="wires" viewBox="0 0 1000 1000" aria-hidden="true"><circle class="ring" cx="500" cy="500" r="340"/><circle class="ring" cx="500" cy="500" r="470"/>{''.join(lines)}</svg>
      {hub}{''.join(nodes)}
    </div>{legend}</div>
  </div>
</section>"""


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-") or "x"


def load_skill_images(skill_counts):
    """Every image in Skill Images/ becomes an orb. File stem = skill name."""
    if not SKILL_IMG_DIR.exists():
        return []
    lower = {k.lower(): v for k, v in skill_counts.items()}
    items = []
    for p in sorted(SKILL_IMG_DIR.iterdir()):
        if p.is_file() and p.suffix.lower() in IMG_MIME:
            name = p.stem.strip()
            placeholder = p.suffix.lower() == ".svg" and "PLACEHOLDER" in p.read_text(errors="ignore")
            items.append({"name": name, "cls": slug(name), "uri": data_uri(p), "count": lower.get(name.lower(), 0),
                          "placeholder": placeholder})
    items.sort(key=lambda s: (-s["count"], s["name"].lower()))
    return items


def stack_geom(people):
    """Everything the hero mark needs, worked out once: the slabs, the star, and the box they sit in.
    grade_stack() draws the resting frame from this and STACK_JS re-draws it as it turns. Heights
    are centred on the middle of the stack, so it turns about its own centre."""
    counts = Counter(str(p.get("Role") or "").strip() for p in people)
    slabs = []
    for grades, colour in HERO_TIERS:
        n = sum(counts.get(g, 0) for g in grades)
        if n:
            slabs.append({"n": n, "c": colour})
    if not slabs:
        return None
    top = max(s["n"] for s in slabs)
    for i, s in enumerate(slabs):
        s["w"] = round((0.30 + s["n"] / top * 0.62) * STACK_S, 2)
        s["y0"] = i * (STACK_H + STACK_GAP)
        s["y1"] = s["y0"] + STACK_H
    height = slabs[-1]["y1"]
    star = None
    if counts.get(STAR_GRADE):
        star = {"y": height + STAR_LIFT + STAR_R, "r": STAR_R, "n": counts[STAR_GRADE]}
        height = star["y"] + STAR_R
    mid = height / 2
    for s in slabs:
        s["y0"] = round(s["y0"] - mid, 2)
        s["y1"] = round(s["y1"] - mid, 2)
    if star:
        star["y"] = round(star["y"] - mid, 2)
    g = {"ry": STACK_RY, "rx": STACK_RX, "light": LIGHT, "t": slabs, "star": star}
    # the box is the resting frame's own outline; the SVG overflows it while it is being turned
    xs, ys = [], []
    for poly, _ in stack_polys(g, STACK_RY, STACK_RX):
        xs += [p[0] for p in poly]
        ys += [p[1] for p in poly]
    pad = 4
    g["box"] = [round(min(xs) - pad, 1), round(min(ys) - pad, 1),
                round(max(xs) - min(xs) + 2 * pad, 1), round(max(ys) - min(ys) + 2 * pad, 1)]
    return g


def shade(hex_colour, k):
    r, g_, b = (int(hex_colour[i:i + 2], 16) for i in (1, 3, 5))
    return "#%02x%02x%02x" % tuple(max(0, min(255, round(c * k))) for c in (r, g_, b))


def stack_polys(g, ry, rx):
    """The visible faces at one turn (ry) and tilt (rx), in the order they are painted, each as
    (screen points, fill). STACK_JS runs the same maths in the page."""
    th, ph = math.radians(ry), math.radians(rx)
    ct, st, cp, sp = math.cos(th), math.sin(th), math.cos(ph), math.sin(ph)
    lx, ly, lz = g["light"]
    ll = math.sqrt(lx * lx + ly * ly + lz * lz)

    def P(x, y, z):
        x1, z1 = x * ct + z * st, -x * st + z * ct
        return (x1, -(y * cp - z1 * sp))

    def lit(n, colour, top):
        nx, ny, nz = n
        x1, z1 = nx * ct + nz * st, -nx * st + nz * ct
        v = (x1, ny * cp - z1 * sp, ny * sp + z1 * cp)          # the normal, in view space
        if v[2] <= 1e-6:
            return None                                           # facing away: not drawn
        d = max(0.0, (v[0] * lx + v[1] * ly + v[2] * lz) / ll)
        return shade(colour, (0.62 if top else 0.5) + 0.5 * d)

    out = []
    layers = list(range(len(g["t"]) + (1 if g["star"] else 0)))
    if sp < 0:
        layers.reverse()                  # seen from below: the top of the stack is nearest
    for i in layers:
        if i == len(g["t"]):
            cx, cy = P(0, g["star"]["y"], 0)
            r = g["star"]["r"]
            pts = [(cx + (r if k % 2 == 0 else r * 0.42) * math.sin(k * math.pi / 5),
                    cy - (r if k % 2 == 0 else r * 0.42) * math.cos(k * math.pi / 5)) for k in range(10)]
            out.append((pts, "star"))
            continue
        s = g["t"][i]
        w, y0, y1 = s["w"], s["y0"], s["y1"]
        q = [(-w, -w), (w, -w), (w, w), (-w, w)]
        normals = [(0, 0, -1), (1, 0, 0), (0, 0, 1), (-1, 0, 0)]
        for f in range(4):
            fill = lit(normals[f], s["c"], False)
            if fill:
                a, b = q[f], q[(f + 1) % 4]
                out.append(([P(a[0], y0, a[1]), P(b[0], y0, b[1]), P(b[0], y1, b[1]), P(a[0], y1, a[1])], fill))
        for n, y in (((0, 1, 0), y1), ((0, -1, 0), y0)):
            fill = lit(n, s["c"], True)
            if fill:
                out.append(([P(x, y, z) for x, z in q], fill))
    return out


def grade_stack(g, extra=""):
    """The hero mark as SVG, drawn at its resting angle so it still reads with JavaScript off."""
    if not g:
        return ""
    paths = "".join(
        (f'<path class="stack-star" d="M{"L".join(f"{x:.1f},{y:.1f}" for x, y in pts)}Z">'
         f'<title>Managing Director{"s" if g["star"]["n"] > 1 else ""}</title></path>') if fill == "star" else
        f'<path d="M{"L".join(f"{x:.1f},{y:.1f}" for x, y in pts)}Z" fill="{fill}"/>'
        for pts, fill in stack_polys(g, g["ry"], g["rx"]))
    return (f'<div class="stack-stage {extra}" title="Drag to turn">'
            f'<svg viewBox="{" ".join(str(v) for v in g["box"])}" role="img" '
            f'aria-label="The practice by grade, one slab per grade, the Managing Director as a star on top">'
            f'<g class="stack-g">{paths}</g></svg>'
            f'<span class="stack-glow"></span></div>')


def avatar(rec, photos, cls=""):
    pid = photos.get(rec["Name"])
    if pid:
        return f'<span class="av {cls} ph-{pid}" role="img" aria-label="{e(rec["Name"])}"></span>'
    return f'<span class="av {cls}" aria-hidden="true">{e(initials(rec["Name"]))}</span>'


def hero(n_people, n_comp, n_skills, skill_imgs, mark):
    pos = [(19, 17), (7, 52), (17, 86), (82, 22), (93, 55), (80, 87)]
    orbs = []
    for i in range(min(ORB_COUNT, len(skill_imgs), len(pos))):
        s = skill_imgs[i]
        x, y = pos[i]
        label = e(s["name"]) + (f' &middot; {s["count"]} {"person" if s["count"] == 1 else "people"}' if s["count"] else "")
        orbs.append(
            f'<div class="orb" data-i="{i}" style="left:{x}%;top:{y}%;--fd:{6 + i * 1.3:.1f}s;--dl:-{i * 1.7:.1f}s">'
            f'<div class="orb-in"><span class="face"><i class="ic ic-{s["cls"]}"></i></span>'
            f'<span class="face back"><i class="ic"></i></span></div>'
            f'<span class="orb-l">{label}</span></div>')
    return f"""
<section class="hero" id="hero">
  <div class="hero-bg" aria-hidden="true"><span class="g g1"></span><span class="g g2"></span><span class="grid"></span></div>
  <div class="orbs">{''.join(orbs)}</div>
  <div class="wrap hero-in">
    <div class="rv"><div class="hero-mark">{mark}</div></div>
    <p class="eyebrow rv" style="--d:60ms">PwC {TEAM_NAME} &middot; Skills &amp; Bio</p>
    <h1 class="rv" style="--d:120ms">The people behind<br><em>every deal</em></h1>
    <p class="lede rv" style="--d:200ms">{n_people} professionals. {n_comp} competencies. {n_skills} skills.<br>One team, mapped end to end.</p>
    <div class="cta-row rv" style="--d:280ms">
      <a class="btn" href="#competencies">Explore competencies</a>
      <button type="button" class="btn btn-ghost" data-dash>Open the full dashboard &rarr;</button>
    </div>
  </div>
  <a class="cue" href="#skills" aria-label="Scroll down"><span></span></a>
</section>"""


def shape(people, children):
    """Team shape, filterable by grade, competency, client territory, or one cell of the grid.

    The practice-level figures are rendered here, so the section reads with JavaScript off. With
    JavaScript, SHAPE_JS recounts the same blocks from a compact per-person table (see __SHAPE
    below) — cheaper than rendering a panel for every grade x competency x territory combination,
    and it keeps one set of definitions rather than one per filter.
    """
    total_all = len(people) or 1

    # ---- the control: one bar per grade group, always practice-level
    groups = [(short, full, [p for p in people if p["Role"] in roles]) for short, full, roles in SHAPE_LEVELS]
    known = {r for _, _, roles in SHAPE_LEVELS for r in roles}
    rest = [p for p in people if p["Role"] not in known]
    if rest:
        groups.append(("Oth", "Other grades", rest))
    groups = [(k, label, mem) for k, label, mem in groups if mem]
    mx = max((len(m) for _, _, m in groups), default=1) or 1
    rows = "".join(
        f'<button type="button" class="pr rv" data-ts="g:{k}" style="--d:{i*70}ms" aria-pressed="false">'
        f'<span class="pr-l">{e(label)}</span>'
        f'<span class="pr-b"><i style="width:{max(len(m) / mx * 100, 1.5):.1f}%"></i></span>'
        f'<b>{len(m)}</b><span class="pr-s">{len(m) / total_all * 100:.0f}%</span></button>'
        for i, (k, label, m) in enumerate(groups))

    junior = sum(1 for p in people if p["Role"] in ("A1", "A2", "SA"))
    senior = sum(1 for p in people if p["Role"] in ("M", "SM", "Director", "MD"))
    lev = (f'<p class="lev">Leverage <b>{junior / senior:.1f}&times;</b> &mdash; {junior} people below Manager '
           f'for every {senior} at Manager and above.</p>' if senior else "")

    # The grid carries an "Other" row and column so it covers everyone: people whose competency is
    # blank or "Other", whose territory is blank or "Other...", and any territory past the top six.
    # Without them the grid's total sat below the headcount and the two figures looked at odds.
    def terr_of(p):
        t = str(p.get("Territory") or "").strip()
        return t if t and not is_other_territory(t) else OTHER
    named = [t for t, _ in ranked(Counter(terr_of(p) for p in people if terr_of(p) != OTHER), 6)]
    spill = any(terr_of(p) == OTHER or terr_of(p) not in named for p in people)
    terrs = named + ([OTHER] if spill else [])
    comps = [c for c, _ in ranked(Counter(p["_comp"] for p in people if p["_comp"] != OTHER))]
    if any(p["_comp"] == OTHER for p in people):
        comps.append(OTHER)

    def cell_of(p):
        t = terr_of(p)
        return p["_comp"], (t if t in named else OTHER)
    offices = [k for k, _ in ranked(Counter(str(p.get("Location") or "Unassigned").strip() for p in people))]
    BANDS = [("0–2 yrs", 0, 3), ("3–5 yrs", 3, 6), ("6–10 yrs", 6, 11), ("11+ yrs", 11, 999)]

    def band_rows(field, stagger=60):
        counts = [(lab, sum(1 for p in people if p[field] is not None and lo <= p[field] < hi)) for lab, lo, hi in BANDS]
        bmx = max((c for _, c in counts), default=1) or 1
        return "".join(
            f'<div class="off rv" style="--d:{i*stagger}ms"><span>{e(lab)}</span>'
            f'<span class="off-b"><i style="width:{c / bmx * 100:.0f}%"></i></span>'
            f'<b>{c}</b><span class="pr-s">{c / total_all * 100:.0f}%</span></div>'
            for i, (lab, c) in enumerate(counts))

    # competency x client territory — every header and every cell is a filter
    cell = Counter(cell_of(p) for p in people)
    top = max((cell[(c, t)] for c in comps for t in terrs), default=0) or 1
    head = "".join(f'<th><button type="button" class="hm-h" data-ts="t:{i}" title="{e(t)}">{e(TERR_SHORT.get(t, t))}</button></th>'
                   for i, t in enumerate(terrs))
    body = ""
    for ci, c in enumerate(comps):
        tds = ""
        for ti, t in enumerate(terrs):
            n = cell[(c, t)]
            tds += (f'<td><button type="button" class="hm" data-ts="x:{ci}:{ti}" style="--a:{0.12 + 0.88 * (n / top):.2f}">{n}</button></td>'
                    if n else '<td><span class="hm zero">&middot;</span></td>')
        tds += f"<td class='tot'>{sum(cell[(c, t)] for t in terrs)}</td>"
        body += (f'<tr><th scope="row"><button type="button" class="hm-h" data-ts="c:{ci}">{e(c)}</button></th>{tds}</tr>')
    foot = "".join(f"<td class='tot'>{sum(cell[(c, t)] for c in comps)}</td>" for t in terrs)
    cols = (f'<colgroup><col class="hm-c0"><col span="{len(terrs)}" class="hm-cx"><col class="hm-ct"></colgroup>'
            if terrs else "")
    grid = f"""<div class="hm-wrap rv" id="ts-grid"><table class="hm-t" style="--cols:{max(len(terrs), 1)}">
        {cols}
        <thead><tr><th></th>{head}<th class="tot">All</th></tr></thead>
        <tbody>{body}</tbody>
        <tfoot><tr><th scope="row">All competencies</th>{foot}<td class="tot">{sum(cell[(c, t)] for c in comps for t in terrs)}</td></tr></tfoot>
      </table></div>"""

    omx = max((sum(1 for p in people if str(p.get("Location") or "Unassigned").strip() == k) for k in offices), default=1) or 1
    otiles = "".join(
        f'<div class="off rv" style="--d:{i*50}ms"><span>{e(k)}</span><span class="off-b">'
        f'<i style="width:{v / omx * 100:.0f}%"></i></span>'
        f'<b>{v}</b><span class="pr-s">{v / total_all * 100:.0f}%</span></div>'
        for i, (k, v) in enumerate((k, sum(1 for p in people if str(p.get("Location") or "Unassigned").strip() == k))
                                   for k in offices))

    joiners = sum(1 for p in people if p["_pwc"] is not None and p["_pwc"] < 1)
    g = Counter(str(p.get("Gender") or "").strip().lower() for p in people)
    mf = g.get("male", 0) + g.get("female", 0)
    ratio = f"{round(g.get('male', 0) / mf * 100)}% : {100 - round(g.get('male', 0) / mf * 100)}%" if mf else "—"
    exps = [p["_exp"] for p in people if p["_exp"] is not None]
    leads = sum(1 for p in people if children.get(p["Name"]))
    mix = (f'<div class="mix-rows">'
           f'<div class="mx"><b>{joiners}</b><span>joined in the last 12 months</span></div>'
           f'<div class="mx"><b>{ratio}</b><span>male : female</span></div>'
           f'<div class="mx"><b>{sum(exps) / len(exps):.1f}</b><span>average years of experience</span></div>'
           f'<div class="mx"><b>{leads}</b><span>people with direct reports</span></div></div>'
           if exps else
           f'<div class="mix-rows">'
           f'<div class="mx"><b>{joiners}</b><span>joined in the last 12 months</span></div>'
           f'<div class="mx"><b>{ratio}</b><span>male : female</span></div>'
           f'<div class="mx"><b>{leads}</b><span>people with direct reports</span></div></div>')

    # ---- the per-person table the browser filters on: [grade, competency, territory, office,
    #      experience, years at PwC, gender, has direct reports]. -1 means "not one of the listed".
    gkey = {}
    for k, _, mem in groups:
        for p in mem:
            gkey[p["Name"]] = k
    ci = {c: i for i, c in enumerate(comps)}
    ti = {t: i for i, t in enumerate(terrs)}
    oi = {o: i for i, o in enumerate(offices)}
    recs = [[gkey.get(p["Name"], "Oth"), ci.get(cell_of(p)[0], -1),
             ti.get(cell_of(p)[1], -1),
             oi.get(str(p.get("Location") or "Unassigned").strip(), -1),
             None if p["_exp"] is None else round(p["_exp"], 2),
             None if p["_pwc"] is None else round(p["_pwc"], 2),
             str(p.get("Gender") or "").strip().lower()[:1],
             1 if children.get(p["Name"]) else 0] for p in people]
    data = json.dumps({"g": [[k, label] for k, label, _ in groups], "c": comps, "t": terrs,
                       "o": offices, "b": [[lab, lo, hi] for lab, lo, hi in BANDS], "p": recs},
                      separators=(",", ":"))

    return f"""
<section class="shape" id="shape">
  <div class="wrap">
    <div class="sec-head">
      <p class="eyebrow rv">Team shape</p>
      <h2 class="rv">Grade by grade. Office by office.</h2>
      <p class="sub rv">How the practice is built: the grade pyramid, which competency serves which client territory, where everyone sits, and how long they have been here. <span class="ts-hint">Click a grade, a competency, a territory or one cell of the grid to filter everything below.</span></p>
      <div class="ts-notes rv"><p class="ts-note" id="ts-note">Every figure below covers the whole practice. Pick a grade, a competency, a territory or a cell to see it on its own.</p></div>
    </div>
    <div class="shape-grid">
      <div class="sh-card"><p class="k">Grade pyramid</p>{rows}{lev}
        <p class="k">Experience</p><div class="exp-rows" id="ts-exp">{band_rows("_exp")}</div></div>
      <div class="sh-card"><p class="k" id="ts-grid-k">Competency &times; client territory &middot; people</p>{grid}</div>
    </div>
    <div class="shape-grid offices">
      <div class="sh-card"><p class="k">Offices</p><div class="off-grid one" id="ts-off">{otiles}</div></div>
      <div class="sh-card"><p class="k">Time at PwC</p><div id="ts-ten">{band_rows("_pwc")}</div>
        <p class="k">Mix</p><div id="ts-mix">{mix}</div></div>
    </div>
  </div>
</section>
<script>window.__SHAPE={data};</script>"""


def thumb_uri(rec, px=208):
    """A small square photo for the Competency Teams view.

    The full-size portraits already cost about a megabyte on the landing page; the dashboard is a
    second document, so embedding them again would double that for a circle drawn at 104px. This
    returns a 208px JPEG instead (a few KB). Without Pillow, or without a photo, the view falls
    back to initials.
    """
    src = photo_for(rec)
    if not src:
        return ""
    try:
        from io import BytesIO
        from PIL import Image
        raw = base64.b64decode(src.split(",", 1)[1])
        im = Image.open(BytesIO(raw)).convert("RGB")
        side = min(im.size)
        im = im.crop(((im.width - side) // 2, (im.height - side) // 2,
                      (im.width + side) // 2, (im.height + side) // 2)).resize((px, px), Image.LANCZOS)
        buf = BytesIO()
        im.save(buf, "JPEG", quality=82, optimize=True)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception as err:                       # Pillow missing, or an image it cannot read
        print(f"[warn] could not shrink the photo for {rec.get('Name')}: {err} — using initials")
        return ""


def competency_teams(people, comps, comp_color, comp_leads, deliv):
    """The "Competency Teams" view added to the dashboard's Team Analytics tab.

    Each competency starts from its top-level people: everyone whose RL manager sits outside it
    (its Directors, and anyone who reports into another team), shown side by side. Picking one
    opens their team as a top-down org chart, one branch at a time; a person with no one below
    them opens their bio. Built from the roster, so it needs no hooks into the dashboard's code.
    """
    if not comps:
        return "", "", "{}"

    css, out = [], []
    s_ = lambda v: str(v or "").strip()
    order = lambda q: (rank(q), q["Name"])

    def node(p):
        return {"n": p["Name"], "r": GRADE_LABEL.get(p["Role"], p["Role"]), "o": s_(p.get("Location")),
                "i": initials(p["Name"]), "w": p["_id"]}

    for idx, (name, members) in enumerate(comps):
        leads = comp_leads.get(name) or []
        names = {p["Name"] for p in members}

        reports = defaultdict(list)                  # only inside this competency
        for p in members:
            mgr = s_(p.get("RL Manager"))
            if mgr in names and mgr != p["Name"]:
                reports[mgr].append(p)

        seen = set()

        def subtree(person):
            """The person and everyone under them, most senior first at each level."""
            seen.add(person["Name"])
            kids = [k for k in sorted(reports.get(person["Name"], []), key=order) if k["Name"] not in seen]
            for k in kids:
                seen.add(k["Name"])
            out_node = node(person)
            out_node["k"] = [subtree(k) for k in kids]
            return out_node

        # Top level: whoever reports outside this competency, most senior first. Then anyone a
        # reporting loop kept out of every line, so no member is ever left off the chart.
        tops = [p for p in sorted(members, key=order)
                if s_(p.get("RL Manager")) not in names or s_(p.get("RL Manager")) == p["Name"]]
        heads = []
        loose = []                     # no team, and their manager sits in another competency
        for p in tops + sorted(members, key=order):
            if p["Name"] in seen:
                continue
            h = subtree(p)
            h["mg"] = s_(p.get("RL Manager"))
            if not h["k"] and p["Role"] not in ("MD", "Director"):
                loose.append(h)        # shown under their manager, which here means in the card below
                continue
            photo = thumb_uri(p)
            if photo:
                cls = f"ct-ph{idx}-{len(heads)}"
                css.append(f".{cls}{{background-image:url({photo})}}")
                h["ph"] = cls
            heads.append(h)
        # People with no one reporting to them are not top level: they appear once their manager is
        # opened. Those whose manager is in another competency have no manager here to open, so they
        # share one card at the end of the row instead of each taking a place in it.
        if loose:
            heads.append({"n": f"Outside {COMPETENCY_CODE.get(name, name).replace('&amp;', '&')}",
                          "r": "Report to other competencies", "o": "", "i": f"+{len(loose)}", "w": "",
                          "grp": 1, "k": loose})

        f = deliv.figures({name}) if deliv else None
        pp = deliv.per_person({name}) if deliv else {}
        exps = [p["_exp"] for p in members if p["_exp"] is not None]
        grades = Counter(GRADE_LABEL.get(p["Role"], p["Role"]) for p in members)

        out.append({
            "n": name, "col": comp_color.get(name, "#FD5108"), "lead": leads[0]["Name"] if leads else "",
            "h": heads,
            "st": {"people": len(members),
                   "leads": sum(1 for p in members if reports.get(p["Name"])),
                   "util": (round(f["util"]) if f and f["util"] is not None else None),
                   "ch": (round(f["ch"]) if f else None),
                   "fte": (round(f["spare"], 1) if f else None),
                   "exp": (round(sum(exps) / len(exps), 1) if exps else None),
                   "off": len({s_(p.get("Location")) for p in members if s_(p.get("Location"))}),
                   "terr": len({s_(p.get("Territory")) for p in members
                                if s_(p.get("Territory")) and not is_other_territory(s_(p.get("Territory")))}),
                   "g": [[g, c] for g, c in sorted(grades.items(), key=lambda kv: GRADE_RANK.get(
                       next((k for k, v in GRADE_LABEL.items() if v == kv[0]), ""), 99))]},
            # one row per member, so the tiles can be re-added in the page for a slice or a line:
            # name, grade, experience, office, territory (blank if "other"), has a team,
            # standard hrs, chargeable hrs, spare FTE. The filters themselves are matched against
            # the dashboard's own EMPLOYEES records, which carry the root build's renamed values.
            "m": [[p["Name"], GRADE_LABEL.get(p["Role"], p["Role"]), p["_exp"], s_(p.get("Location")),
                   "" if is_other_territory(s_(p.get("Territory"))) else s_(p.get("Territory")),
                   1 if reports.get(p["Name"]) else 0]
                  + [round(v, 2) for v in pp.get(p["_id"], [0.0, 0.0, 0.0])] for p in members],
        })

    chips = "".join(
        f'<button type="button" class="ct-chip{" is-on" if i == 0 else ""}" data-ct="{i}">'
        f'<span class="ct-dot" style="background:{c["col"]}"></span>'
        f'<b>{e(c["n"])}</b><span class="ct-n">{c["st"]["people"]} people</span></button>'
        for i, c in enumerate(out))

    # the same card as the tab's other views, so switching views keeps the content where it was
    html = f"""<div class="ct-wrap ana-card full" id="ct-wrap" hidden>
  <h3 class="ana-h">Competency Teams</h3>
  <p class="ana-sub ct-intro">Each competency starts from the people who lead a team in it; everyone else appears once their manager is opened, and anyone whose manager sits in another competency is gathered in one card at the end. Pick a leader to open their team as an org chart, then click anyone with a team to go a level down, one branch at a time; large teams wrap into rows. Click someone with no one below them to open their bio. The filters above narrow who is counted; anyone outside the slice stays in the chart, dimmed, so the reporting line still reads.</p>
  <div class="ct-picker" id="ct-picker">{chips}</div>
  <div class="ct-sum" id="ct-sum"></div>
  <nav class="ct-crumbs" id="ct-crumbs" aria-label="Reporting line"></nav>
  <div class="ct-heads" id="ct-heads"></div>
  <div class="ct-tree" id="ct-tree"></div>
</div>"""
    lm = month_label(deliv.latest) if deliv and deliv.latest else ""
    return html, "\n".join(css), json.dumps({"c": out, "lm": lm}, separators=(",", ":"))


CT_CSS = r"""
/* ===== Competency Teams — a Team Analytics view added by build_scroll_story.py ===== */
.ct-wrap[hidden]{display:none}
.ct-intro{max-width:860px}
.ct-picker{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:18px}
.ct-chip{display:inline-flex;align-items:center;gap:9px;background:rgba(255,255,255,.04);border:1px solid var(--ds-line);
  border-radius:999px;padding:9px 16px;color:var(--ds-mu);font:inherit;font-size:13px;cursor:pointer;transition:all .18s}
.ct-chip:hover{color:var(--ds-tx);background:rgba(255,255,255,.08)}
.ct-chip.is-on{color:var(--ds-tx);border-color:rgba(253,81,8,.6);background:rgba(253,81,8,.12)}
.ct-chip[hidden]{display:none}
.ct-chip b{font-weight:600}
.ct-chip .ct-n{color:var(--ds-mu2)}
.ct-dot{width:9px;height:9px;border-radius:50%;flex:none}

/* summary strip: who is in focus, then their tiles */
.ct-sum{display:grid;grid-template-columns:minmax(210px,280px) minmax(0,1fr);gap:20px;align-items:center;
  border:1px solid var(--ds-line);border-radius:14px;padding:16px 18px;margin:0 0 16px;background:rgba(255,255,255,.02)}
.ct-sum h4{margin:0 0 3px;font-family:var(--serif);font-size:19px;font-weight:400;color:var(--ds-tx)}
.ct-sum .ct-meta{margin:0 0 9px;color:var(--ds-mu);font-size:12.5px}
.ct-gm{display:flex;flex-wrap:wrap;gap:5px}
.ct-gm span{font-size:11px;color:var(--ds-mu);border:1px solid var(--ds-line);border-radius:999px;padding:2px 8px}
.ct-gm b{color:var(--ds-tx);font-weight:600}
.ct-tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(108px,1fr));gap:10px}
.ct-tile{border:1px solid var(--ds-line);border-radius:12px;padding:10px 12px}
.ct-tile b{display:block;font-family:var(--serif);font-size:20px;color:var(--ds-tx);line-height:1.1}
.ct-tile span{font-size:11px;color:var(--ds-mu2)}

/* the reporting line you have opened */
.ct-crumbs{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin:0 0 14px;font-size:13px;min-height:22px}
.ct-crumb{background:none;border:0;padding:2px 4px;font:inherit;color:var(--pwc-orange);cursor:pointer;border-radius:6px}
.ct-crumb:hover{background:rgba(255,255,255,.07)}
.ct-crumb.is-on{color:var(--ds-tx);cursor:default;font-weight:600}
.ct-sep{color:var(--ds-mu2)}

/* avatars */
.ct-av{flex:none;width:38px;height:38px;border-radius:50%;background:#24242b center/cover no-repeat;border:1.5px solid var(--ds-line2);
  display:grid;place-items:center;overflow:hidden;font-weight:600;font-size:12.5px;color:var(--ds-tx)}
.ct-av img{width:100%;height:100%;object-fit:cover}

/* top-level people, side by side */
.ct-heads{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px;margin:0 0 8px}
.ct-head{display:flex;align-items:center;gap:13px;padding:14px 16px;background:#1b1b21;border:1px solid var(--ds-line2);
  border-radius:16px;color:inherit;font:inherit;text-align:left;cursor:pointer;transition:border-color .2s,transform .2s,box-shadow .2s}
.ct-head:hover{border-color:rgba(253,81,8,.55);transform:translateY(-2px)}
.ct-head.is-on{border-color:var(--pwc-orange);box-shadow:0 0 0 3px rgba(253,81,8,.18)}
.ct-head .ct-av{width:54px;height:54px;font-size:16px;font-family:var(--serif)}
.ct-head .ct-txt b{font-size:14px}
.ct-head .ct-cnt{display:block;margin-top:5px;font-size:11.5px;color:#FFB08A}
.ct-head .ct-out{display:block;font-size:11px;color:var(--ds-mu2)}
.ct-heads.compact{grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:8px;margin-bottom:22px}
.ct-heads.compact .ct-head{padding:8px 12px;border-radius:12px}
.ct-heads.compact .ct-head .ct-av{width:34px;height:34px;font-size:12px}
.ct-heads.compact .ct-cnt,.ct-heads.compact .ct-out{display:none}
.ct-hint{color:var(--ds-mu);font-size:13px;margin:14px 0 4px}

/* person card, used in the chart */
.ct-card{display:flex;align-items:center;gap:10px;width:228px;padding:9px 12px 9px 9px;background:#1b1b21;
  border:1px solid var(--ds-line2);border-radius:14px;color:inherit;font:inherit;text-align:left;cursor:pointer;
  transition:border-color .2s,transform .2s,box-shadow .2s;position:relative;z-index:1}
.ct-card:hover{border-color:rgba(253,81,8,.55);transform:translateY(-2px)}
.ct-card.is-open{border-color:var(--pwc-orange);box-shadow:0 0 0 3px rgba(253,81,8,.18)}
.ct-card.is-root{width:250px;padding:12px 16px 12px 12px;cursor:default;
  background:linear-gradient(135deg,rgba(253,81,8,.16),rgba(253,81,8,.03) 60%),#1b1b21;border-color:rgba(253,81,8,.5)}
.ct-card.is-root:hover{transform:none}
.ct-card.is-root .ct-av{width:50px;height:50px;font-size:15px;font-family:var(--serif)}
.ct-txt{display:flex;flex-direction:column;min-width:0;flex:1}
.ct-txt b{font-size:13px;font-weight:600;color:var(--ds-tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ct-txt span{font-size:11.5px;color:var(--ds-mu);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ct-badge{flex:none;font-size:11px;font-weight:600;padding:3px 8px;border-radius:999px;background:rgba(253,81,8,.14);color:#FFB08A;white-space:nowrap}
.ct-card.is-open .ct-badge{background:var(--pwc-orange);color:#fff}
.ct-txt .ct-off{color:var(--ds-mu2);font-size:11px}
.ct-bio{flex:none;font-size:11px;color:var(--ds-mu2);white-space:nowrap}
.ct-card:hover .ct-bio{color:var(--pwc-orange)}
.ct-wrap .dim{opacity:.35}

/* the org chart: straight lines with rounded corners, drawn from each <li>'s own borders */
.ct-tree{--ct-ln:#3d3d45;--ct-gap:30px;overflow-x:auto;padding:4px 4px 18px;position:relative}
.ct-org,.ct-org ul{display:flex;justify-content:center;margin:0;padding:0;list-style:none}
.ct-org{min-width:max-content;margin:0 auto}
.ct-org ul{position:relative;padding-top:var(--ct-gap)}
.ct-org li{position:relative;display:flex;flex-direction:column;align-items:center;padding:var(--ct-gap) 8px 0}
.ct-org > li{padding-top:0}
.ct-org ul::before{content:'';position:absolute;top:0;left:50%;height:var(--ct-gap);border-left:1.5px solid var(--ct-ln)}
.ct-org li::before,.ct-org li::after{content:'';position:absolute;top:0;width:50%;height:var(--ct-gap);border-top:1.5px solid var(--ct-ln)}
.ct-org li::before{right:50%}
.ct-org li::after{left:50%;border-left:1.5px solid var(--ct-ln)}
.ct-org li:first-child::before,.ct-org li:last-child::after{border:0}
.ct-org li:last-child::before{border-right:1.5px solid var(--ct-ln);border-radius:0 14px 0 0}
.ct-org li:first-child::after{border-radius:14px 0 0 0}
.ct-org li:only-child::before{display:none}
.ct-org li:only-child::after{border-top:0;border-left:1.5px solid var(--ct-ln);border-radius:0}
.ct-org > li::before,.ct-org > li::after{display:none}
/* the line down to the branch you opened is drawn in orange */
.ct-org li.on-path > ul::before{border-left-color:var(--pwc-orange)}
.ct-org ul.ct-new{animation:ct-in .35s ease both}
/* a wide level: rows of up to four cards in one panel, one line down to it from the manager */
.ct-org .ct-rows{display:grid;grid-template-columns:repeat(var(--cols),228px);gap:12px;padding:14px;
  border:1px solid var(--ds-line);border-radius:18px;background:rgba(255,255,255,.02);position:relative;z-index:1}
.ct-org .ct-cap{margin:18px 0 0;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:var(--pwc-orange);position:relative;z-index:1;
  background:var(--ds-card);padding:0 8px}
.ct-org .ct-cap+ul{padding-top:14px}
.ct-org .ct-cap+ul::before{height:14px}
@media (max-width:1100px){.ct-org .ct-rows{grid-template-columns:repeat(2,228px)}}
@keyframes ct-in{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:none}}
@media (prefers-reduced-motion:reduce){.ct-org ul.ct-new{animation:none}.ct-card,.ct-head{transition:none}}
.ct-empty{color:var(--ds-mu);font-size:14px;padding:30px 0}
@media (max-width:900px){.ct-sum{grid-template-columns:1fr}}
"""

CT_JS = r"""
(function(){
  var D = window.__CT, wrap = document.getElementById('ct-wrap');
  if (!D || !D.c || !D.c.length || !wrap) return;
  var toggle = document.getElementById('ana-view-toggle'),
      body = document.getElementById('ana-body'),
      filters = document.getElementById('ana-filters'),
      fy = document.getElementById('ana-fy-toggle'),
      picker = document.getElementById('ct-picker'),
      sum = document.getElementById('ct-sum'),
      crumbs = document.getElementById('ct-crumbs'),
      headsEl = document.getElementById('ct-heads'),
      tree = document.getElementById('ct-tree');
  if (!toggle || !body || !tree) return;

  var ci = 0, hi = -1, path = [];      /* competency, top-level person picked, branch opened below them */

  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  function num(n){ return n == null ? '—' : n.toLocaleString(); }

  /* --- the tab's own filters (the same controls Workforce Mix uses) decide who is in the slice.
     The selection is read off each control's chips, and matched against the dashboard's own
     EMPLOYEES records so role and competency names are the ones the filters list. --- */
  var MS = ['ana-mgr-ms', 'ana-role-ms', 'ana-loc-ms', 'ana-comp-ms', 'ana-terr-ms'],
      FIELD = ['RL Manager', 'Role', 'Location', 'Competency Filter', 'Territory Filter'],
      EMP = {}, inSlice = null;           /* null = no filters: everyone counts */
  try { EMPLOYEES.forEach(function(e){ var k = String(e.Name || '').trim(); (EMP[k] = EMP[k] || []).push(e); }); } catch (err) {}
  function readFilters(){
    var s = MS.map(function(id){
      return [].map.call(document.querySelectorAll('#' + id + ' .ms-chip-x'), function(x){ return x.getAttribute('data-v'); });
    });
    if (!Object.keys(EMP).length || !s.some(function(v){ return v.length; })){ inSlice = null; return; }
    inSlice = function(name){
      return (EMP[name] || []).some(function(e){
        function v(f){ return String(e[f] == null ? '' : e[f]).trim(); }
        if (s[0].length && s[0].indexOf(v('RL Manager')) < 0 && s[0].indexOf(v('Name')) < 0) return false;
        for (var i = 1; i < FIELD.length; i++) if (s[i].length && s[i].indexOf(v(FIELD[i])) < 0) return false;
        return true;
      });
    };
  }
  function on(name){ return !inSlice || inSlice(name); }

  /* --- the tree --- */
  function line(n, acc){                 /* the person and everyone under them */
    acc = acc || [];
    acc.push(n.n);
    (n.k || []).forEach(function(k){ line(k, acc); });
    return acc;
  }
  function anyOn(n){ return line(n).some(on); }
  function opened(){                     /* the head, then each person opened below them */
    var c = D.c[ci], out = [];
    if (hi < 0) return out;
    var n = c.h[hi];
    out.push(n);
    for (var d = 0; d < path.length && n; d++){ n = (n.k || [])[path[d]]; if (n) out.push(n); }
    return out;
  }

  /* tiles over a set of people (a whole competency when names is null), inside the slice */
  function stats(c, names){
    var m = c.m.filter(function(r){ return (!names || names.indexOf(r[0]) >= 0) && on(r[0]); });
    if (!names && m.length === c.m.length) return c.st;
    var s = {people: m.length, leads: 0, off: 0, terr: 0, g: []}, std = 0, ch = 0, fte = 0, ex = [],
        off = {}, ter = {}, g = {};
    m.forEach(function(r){
      g[r[1]] = (g[r[1]] || 0) + 1;
      if (r[2] != null) ex.push(r[2]);
      if (r[3]) off[r[3]] = 1;
      if (r[4]) ter[r[4]] = 1;
      s.leads += r[5]; std += r[6]; ch += r[7]; fte += r[8];
    });
    var hasU = c.st.ch != null;
    s.util = hasU && std ? Math.round(ch / std * 100) : null;
    s.ch = hasU ? Math.round(ch) : null;
    s.fte = hasU ? Math.round(fte * 10) / 10 : null;
    s.exp = ex.length ? Math.round(ex.reduce(function(a, b){ return a + b; }, 0) / ex.length * 10) / 10 : null;
    s.off = Object.keys(off).length; s.terr = Object.keys(ter).length;
    var order = c.st.g.map(function(x){ return x[0]; });
    s.g = Object.keys(g).sort(function(a, b){
      var i = order.indexOf(a), j = order.indexOf(b);
      return (i < 0 ? 99 : i) - (j < 0 ? 99 : j);
    }).map(function(k){ return [k, g[k]]; });
    return s;
  }

  /* --- pieces --- */
  function avatar(n){
    if (n.ph) return '<span class="ct-av ' + n.ph + '"></span>';
    var e = (EMP[n.n] || [])[0], src = e && e._photo;
    return '<span class="ct-av">' + (src ? '<img loading="lazy" alt="" src="' + esc(src) + '">' : esc(n.i)) + '</span>';
  }
  function reportsTxt(n){
    if (n.grp) return n.k.length + (n.k.length === 1 ? ' person' : ' people');
    var k = (n.k || []).length, all = line(n).length - 1;
    return k ? k + (k === 1 ? ' report' : ' reports') + (all > k ? ' · ' + all + ' in line' : '') : 'no reports';
  }

  function drawSum(){
    var c = D.c[ci], o = opened(), f = o[o.length - 1];
    var s = f ? stats(c, line(f)) : stats(c, null), cut = !f && s !== c.st;
    var title = f ? esc(f.n) : esc(c.n),
        meta = f ? (f.grp ? reportsTxt(f) + ' in ' + esc(c.n) + ' whose managers sit in other competencies'
                          : esc(f.r) + (f.o ? ' · ' + esc(f.o) : '') + ' · ' + reportsTxt(f))
                 : (c.lead ? 'Led by ' + esc(c.lead) : 'Guided by the Managing Director') + ' · ' +
                   c.h.length + (c.h.length === 1 ? ' top-level person' : ' top-level people') +
                   (cut ? ' · ' + s.people + ' of ' + c.st.people + ' in the current filters' : '');
    var tiles = [[num(s.people), f ? 'people in line' : 'people'], [s.util == null ? '—' : s.util + '%', 'utilization'],
                 [s.fte == null ? '—' : s.fte.toFixed(1), 'available FTE' + (D.lm ? ' · ' + esc(D.lm) : '')],
                 [num(s.ch), 'chargeable hours'], [s.exp == null ? '—' : s.exp, 'avg years exp.'],
                 [s.off + ' / ' + s.terr, 'offices / territories']];
    sum.innerHTML = '<div><h4>' + title + '</h4><p class="ct-meta">' + meta + '</p><div class="ct-gm">' +
      s.g.map(function(g){ return '<span>' + esc(g[0]) + ' <b>' + g[1] + '</b></span>'; }).join('') + '</div></div>' +
      '<div class="ct-tiles">' + tiles.map(function(t){
        return '<div class="ct-tile"><b>' + t[0] + '</b><span>' + t[1] + '</span></div>'; }).join('') + '</div>';
  }

  function drawCrumbs(){
    var o = opened();
    crumbs.innerHTML = (o.length ? '<button type="button" class="ct-crumb" data-go="-2">' + esc(D.c[ci].n) + '</button>'
                                 : '<span class="ct-crumb is-on">' + esc(D.c[ci].n) + '</span>') +
      o.map(function(n, i){
        return '<span class="ct-sep">›</span>' + (i === o.length - 1
          ? '<span class="ct-crumb is-on">' + esc(n.n) + '</span>'
          : '<button type="button" class="ct-crumb" data-go="' + (i - 1) + '">' + esc(n.n) + '</button>');
      }).join('');
  }

  function drawHeads(){
    var c = D.c[ci];
    headsEl.classList.toggle('compact', hi >= 0);
    headsEl.innerHTML = c.h.map(function(h, i){
      var outside = h.grp ? '<span class="ct-out">their managers sit in other competencies</span>'
                  : h.mg && h.r !== 'Director' ? '<span class="ct-out">reports to ' + esc(h.mg) + '</span>' : '';
      return '<button type="button" class="ct-head' + (i === hi ? ' is-on' : '') + (anyOn(h) ? '' : ' dim') +
        '" data-h="' + i + '">' + avatar(h) + '<span class="ct-txt"><b>' + esc(h.n) + '</b><span>' + esc(h.r) +
        (h.o ? ' · ' + esc(h.o) : '') + '</span>' + outside + '<span class="ct-cnt">' + reportsTxt(h) + '</span></span></button>';
    }).join('');
  }

  /* one card; a person with a team opens it, a person without one opens their bio */
  function card(n, d, i, root){
    var k = (n.k || []).length, open = !root && path[d] === i && k;
    var cls = 'ct-card' + (root ? ' is-root' : '') + (open ? ' is-open' : '') + (n.grp || on(n.n) ? '' : ' dim');
    var tail = root ? '<span class="ct-badge">' + k + '</span>'
             : k ? '<span class="ct-badge">' + k + (open ? ' ▴' : ' ▾') + '</span>'
             : '<span class="ct-bio">Bio ›</span>';
    var at = root ? ' aria-label="' + esc(n.n) + '"' : ' data-d="' + d + '" data-i="' + i + '"' + (k ? ' aria-expanded="' + !!open + '"' : '');
    return '<button type="button" class="' + cls + '"' + at + ' title="' + esc(n.n + ' · ' + n.r + (n.o ? ' · ' + n.o : '')) +
      '">' + avatar(n) + '<span class="ct-txt"><b>' + esc(n.n) + '</b><span>' + esc(n.r) + '</span>' +
      (n.o ? '<span class="ct-off">' + esc(n.o) + '</span>' : '') + '</span>' + tail + '</button>';
  }
  var ROW_MAX = 4;                 /* more people than this on a level: rows of up to four, not one long line */
  function level(n, d){
    if (!(n.k || []).length) return '';
    var fresh = d === path.length ? ' ct-new' : '';
    if (n.k.length <= ROW_MAX){
      return '<ul class="' + fresh + '">' + n.k.map(function(q, i){
        var deeper = path[d] === i && (q.k || []).length;
        return '<li' + (deeper ? ' class="on-path"' : '') + '>' + card(q, d, i, false) + (deeper ? level(q, d + 1) : '') + '</li>';
      }).join('') + '</ul>';
    }
    /* A wide level wraps into rows inside one panel, hung from its manager by a single line. The
       team of whoever is opened in it follows underneath, captioned, rather than under their card. */
    var open = n.k[path[d]], deeper = open && (open.k || []).length ? open : null;
    return '<ul class="ct-wide' + fresh + '"><li' + (deeper ? ' class="on-path"' : '') + '>' +
      '<div class="ct-rows" style="--cols:' + Math.min(ROW_MAX, n.k.length) + '">' +
      n.k.map(function(q, i){ return card(q, d, i, false); }).join('') + '</div>' +
      (deeper ? '<p class="ct-cap">' + esc(deeper.n) + '&rsquo;s team</p>' + level(deeper, d + 1) : '') + '</li></ul>';
  }
  function drawTree(){
    var c = D.c[ci];
    if (hi < 0){
      tree.innerHTML = '<p class="ct-hint">Pick a top-level person above to see their team.</p>';
      return;
    }
    var h = c.h[hi];
    tree.innerHTML = (h.k || []).length
      ? '<ul class="ct-org"><li class="on-path">' + card(h, -1, -1, true) + level(h, 0) + '</li></ul>'
      : '<ul class="ct-org"><li>' + card(h, -1, -1, true) + '</li></ul><p class="ct-hint">No one in this competency reports to ' +
        esc(h.n) + '. <button type="button" class="ct-crumb" data-bio="' + esc(h.w) + '">Open their bio ›</button></p>';
    /* bring the level just opened into view: it grows downwards, so the page scrolls, not the chart */
    var fresh = tree.querySelector('.ct-new');
    if (fresh && path.length){
      var r = fresh.getBoundingClientRect();
      if (r.bottom > window.innerHeight) fresh.scrollIntoView({block: 'nearest', behavior: 'smooth'});
    }
  }

  function render(){
    /* a competency with no one in the slice drops out of the picker, as it would from Workforce Mix */
    var shown = D.c.map(function(c){ return c.m.filter(function(r){ return on(r[0]); }).length; });
    if (!shown[ci]){
      var first = shown.findIndex(function(k){ return k > 0; });
      if (first >= 0){ ci = first; hi = -1; path = []; }
    }
    [].forEach.call(picker.querySelectorAll('.ct-chip'), function(b, i){
      b.classList.toggle('is-on', i === ci);
      b.hidden = !shown[i];
      var n = b.querySelector('.ct-n');
      if (n) n.textContent = (inSlice && shown[i] !== D.c[i].m.length ? shown[i] + ' of ' : '') + D.c[i].st.people + ' people';
    });
    if (!shown[ci]){
      sum.innerHTML = ''; crumbs.innerHTML = ''; headsEl.innerHTML = '';
      tree.innerHTML = '<p class="ct-empty">No competency team has anyone in the current filters.</p>';
      return;
    }
    drawSum(); drawCrumbs(); drawHeads(); drawTree();
  }

  function bio(w, list){
    if (typeof showBio === 'function' && w) showBio(w, false, list && list.length > 1 ? list : null, !(list && list.length > 1));
  }

  picker.addEventListener('click', function(ev){
    var b = ev.target.closest('.ct-chip'); if (!b) return;
    ci = +b.getAttribute('data-ct'); hi = -1; path = []; render();
  });
  headsEl.addEventListener('click', function(ev){
    var b = ev.target.closest('[data-h]'); if (!b) return;
    var i = +b.getAttribute('data-h');
    hi = (hi === i ? -1 : i); path = []; render();
  });
  crumbs.addEventListener('click', function(ev){
    var b = ev.target.closest('[data-go]'); if (!b) return;
    var go = +b.getAttribute('data-go');
    if (go === -2){ hi = -1; path = []; } else path = path.slice(0, go + 1);
    render();
  });
  tree.addEventListener('click', function(ev){
    var b = ev.target.closest('[data-bio]');
    if (b){ bio(b.getAttribute('data-bio')); return; }
    b = ev.target.closest('.ct-card[data-d]'); if (!b) return;
    var d = +b.getAttribute('data-d'), i = +b.getAttribute('data-i');
    var parent = opened()[d], q = parent && (parent.k || [])[i];
    if (!q) return;
    if ((q.k || []).length){                     /* has a team: open it here, closing its siblings' */
      path = path[d] === i ? path.slice(0, d) : path.slice(0, d).concat(i);
      render();
    } else {                                     /* the end of the line: their bio, stepping through the siblings */
      bio(q.w, parent.k.map(function(x){ return x.w; }));
    }
  });

  /* the filters and the "showing N people" line stay, as on Workforce Mix; only the fiscal-year
     toggle goes, since nothing here is a monthly series */
  function chrome(show){
    body.style.display = show ? 'none' : '';
    wrap.hidden = !show;
    if (fy) fy.style.display = show ? 'none' : '';
  }
  toggle.addEventListener('click', function(ev){
    var b = ev.target.closest ? ev.target.closest('button[data-view]') : null;
    if (!b) return;
    if (b.getAttribute('data-view') === 'cteam'){
      ev.stopPropagation(); ev.preventDefault();
      [].forEach.call(toggle.querySelectorAll('button'), function(x){ x.classList.toggle('active', x === b); });
      chrome(true); readFilters(); render();
    } else {
      /* The dashboard's own handler ignores a click on the view it thinks is still showing (it
         never saw Competency Teams), so it would leave the highlight here. Move it ourselves. */
      [].forEach.call(toggle.querySelectorAll('button'), function(x){ x.classList.toggle('active', x === b); });
      chrome(false);
    }
  }, true);

  /* any change to the filters — a tick, a chip removed, Clear, Reset, or another tab's synced
     filter — repaints this view while it is showing */
  var queued = false;
  if (filters && window.MutationObserver) new MutationObserver(function(){
    if (wrap.hidden || queued) return;
    queued = true;
    requestAnimationFrame(function(){ queued = false; readFilters(); render(); });
  }).observe(filters, {childList: true, subtree: true});
})();
"""


def closing(generated):
    return f"""
<section class="closing">
  <div class="hero-bg" aria-hidden="true"><span class="g g1"></span><span class="g g2"></span></div>
  <div class="wrap">
    <h2 class="rv">Find the right people<br>for your next deal.</h2>
    <p class="lede rv" style="--d:100ms">Search every profile, skill and reporting line in the full dashboard.</p>
    <div class="cta-row rv" style="--d:180ms"><button type="button" class="btn" data-dash>Open the full dashboard &rarr;</button><a class="btn btn-ghost" href="#hero">Back to top</a></div>
  </div>
</section>
<footer class="foot"><div class="wrap">PwC {TEAM_NAME} &middot; Skills &amp; Bio &middot; Generated {e(generated)} from Employee Details.xlsx &middot; Internal use only</div></footer>"""


# ---------------------------------------------------------------- page
CSS = r"""
:root{--bg:#0B0B0D;--bg2:#111114;--card:#151519;--line:rgba(255,255,255,.08);--line2:rgba(255,255,255,.16);
--tx:#F5F5F4;--mu:#A5A5AC;--mu2:#72727A;--or:#FD5108;--tg:#EB8C00;--yl:#FFB600;--pk:#E669A2;
--serif:Georgia,'Times New Roman',serif;--sans:'Helvetica Neue',Arial,sans-serif;--ease:cubic-bezier(.2,.7,.1,1)}
*{box-sizing:border-box}
html{scroll-behavior:smooth;-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--tx);font-family:var(--sans);font-size:16px;line-height:1.55;overflow-x:hidden}
a{color:inherit;text-decoration:none}
h1,h2,h3{font-family:var(--serif);font-weight:400;letter-spacing:-.02em;margin:0}
.wrap{max-width:1200px;margin:0 auto;padding:0 32px}
.eyebrow{font-size:12px;letter-spacing:.16em;text-transform:uppercase;color:var(--tg);margin:0 0 18px;font-weight:600}
.k{font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:var(--mu2);margin:26px 0 10px;font-weight:600}
.btn{display:inline-flex;align-items:center;gap:8px;background:var(--or);color:#fff;border:0;cursor:pointer;font-family:inherit;border-radius:999px;padding:14px 24px;font-weight:600;font-size:15px;transition:transform .2s var(--ease),background .2s}
.btn:hover{transform:translateY(-2px);background:#ff6a2b}
.btn-ghost{background:transparent;border:1px solid var(--line2);color:var(--tx)}
.btn-ghost:hover{background:rgba(255,255,255,.06)}
.btn-sm{padding:9px 16px;font-size:13px}
.chips{display:flex;flex-wrap:wrap;gap:8px}
.chip{display:inline-flex;align-items:baseline;gap:6px;border:1px solid var(--line2);border-radius:999px;padding:6px 13px;font-size:13px;color:#dcdce0;white-space:nowrap;background:rgba(255,255,255,.02)}
.chip sup{font-size:11px;color:var(--tg);top:0;font-weight:700}

/* nav */
.progress{position:fixed;top:0;left:0;right:0;height:2px;z-index:60}
.progress i{display:block;height:100%;width:100%;background:linear-gradient(90deg,var(--or),var(--yl));transform-origin:0 50%;transform:scaleX(0)}
.top{position:fixed;top:0;left:0;right:0;z-index:50;transition:background .3s,border-color .3s;border-bottom:1px solid transparent}
.top.scrolled{background:rgba(11,11,13,.78);-webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);border-color:var(--line)}
/* the bar runs the full width of the window rather than the 1200px column the sections use,
   so the brand sits in the left corner and the dashboard button in the right one */
.top-in{display:flex;align-items:center;gap:28px;height:68px;padding:0 clamp(20px,3vw,52px)}
.brand{display:flex;align-items:center;gap:12px;font-weight:600;font-size:14px}
.logo{background:#fff;border-radius:8px;padding:4px 8px;display:flex}
.logo img{height:22px;display:block}
.links{display:flex;gap:22px;margin-left:auto;font-size:14px;color:var(--mu)}
.links a:hover{color:var(--tx)}

/* hero */
.hero{position:relative;min-height:100vh;min-height:100svh;display:flex;align-items:center;text-align:center;overflow:hidden;padding:120px 0 90px}
.hero-bg{position:absolute;inset:0;pointer-events:none;overflow:hidden}
.g{position:absolute;border-radius:50%;filter:blur(90px);opacity:.42}
.g1{width:52vw;height:52vw;left:-10vw;top:-18vw;background:radial-gradient(circle,var(--or),transparent 65%);animation:drift 18s ease-in-out infinite alternate}
.g2{width:46vw;height:46vw;right:-14vw;top:10vh;background:radial-gradient(circle,var(--pk),transparent 65%);opacity:.26;animation:drift 22s ease-in-out infinite alternate-reverse}
.g3{width:40vw;height:40vw;left:30vw;bottom:-26vw;background:radial-gradient(circle,var(--yl),transparent 65%);opacity:.2;animation:drift 26s ease-in-out infinite alternate}
.grid{position:absolute;inset:0;background-image:linear-gradient(var(--line) 1px,transparent 1px),linear-gradient(90deg,var(--line) 1px,transparent 1px);background-size:64px 64px;-webkit-mask-image:radial-gradient(ellipse at 50% 45%,#000 10%,transparent 70%);mask-image:radial-gradient(ellipse at 50% 45%,#000 10%,transparent 70%)}
@keyframes drift{to{transform:translate(6vw,4vw) scale(1.12)}}
.hero-in{position:relative;width:100%}
.lede{font-size:clamp(15px,1.25vw,18px);color:var(--mu);margin:22px auto 0;max-width:580px}
.cta-row{display:flex;gap:14px;justify-content:center;flex-wrap:wrap;margin-top:38px}
.cue{position:absolute;bottom:28px;left:50%;transform:translateX(-50%);width:26px;height:42px;border:1px solid var(--line2);border-radius:14px}
.cue span{position:absolute;left:50%;top:9px;width:3px;height:8px;margin-left:-1.5px;border-radius:2px;background:var(--tx);animation:cue 1.8s ease-in-out infinite}
@keyframes cue{50%{transform:translateY(12px);opacity:.2}}

/* marquee */
.marquee{border-top:1px solid var(--line);border-bottom:1px solid var(--line);padding:30px 0;background:var(--bg2);overflow:hidden;-webkit-mask-image:linear-gradient(90deg,transparent,#000 12%,#000 88%,transparent);mask-image:linear-gradient(90deg,transparent,#000 12%,#000 88%,transparent)}
.mq-label{text-align:center;font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:var(--mu2);margin:0 0 18px}
.mq-row{overflow:hidden;padding:5px 0}
.mq-track{display:flex;gap:10px;width:max-content;animation:mq 60s linear infinite}
.mq-track .dup{display:flex;gap:10px;padding-right:10px}
.mq-row.rev .mq-track{animation-direction:reverse;animation-duration:70s}
.marquee:hover .mq-track{animation-play-state:paused}
@keyframes mq{to{transform:translateX(-50%)}}

/* statement */
.statement{padding:22vh 0}
.big{font-family:var(--serif);font-size:clamp(30px,4.6vw,62px);line-height:1.18;letter-spacing:-.015em;margin:0;max-width:1050px}
.w{transition:opacity .35s,color .35s}
.num{color:var(--tg);font-style:italic}

/* headline KPIs — two rows of five */
.kpis-sec{border-top:1px solid var(--line);border-bottom:1px solid var(--line);padding:76px 0 70px}
.kp-row+.kp-row{margin-top:50px}
.kp-h{font-size:12px;letter-spacing:.16em;text-transform:uppercase;color:var(--tg);font-weight:600;margin:0 0 4px}
.kp-grid{display:grid;grid-template-columns:repeat(5,1fr)}
.kp{padding:24px 22px 6px;border-left:1px solid var(--line)}
.kp:first-child{border-left:0;padding-left:0}
.kp b{display:block;font-family:var(--serif);font-weight:400;font-size:clamp(40px,4.4vw,66px);line-height:1;letter-spacing:-.03em;white-space:nowrap;font-variant-numeric:tabular-nums lining-nums}
.kp span{display:block;margin-top:12px;color:var(--mu);font-size:13px;text-transform:uppercase;letter-spacing:.1em}
.kp small{display:block;margin-top:4px;color:var(--mu2);font-size:12px}

/* competencies — two-column grid */
.competencies{padding:130px 0 90px}
.comp-intro{display:flex;justify-content:space-between;align-items:flex-end;gap:32px;margin-bottom:44px}
.comp-intro h2,.sec-head h2{font-size:clamp(34px,4vw,54px);line-height:1.06}
.comp-intro .sub{margin:14px 0 0;max-width:none}
.tg{display:none;flex:none;padding:4px;border:1px solid var(--line2);border-radius:999px;background:var(--card)}
.js .tg{display:inline-flex}
.tg button{font:inherit;font-size:14px;font-weight:600;color:var(--mu);background:transparent;border:0;border-radius:999px;padding:9px 18px;cursor:pointer;transition:background .25s,color .25s}
.tg button[aria-pressed="true"]{background:var(--or);color:#fff}
.tg button:focus-visible{outline:2px solid var(--tg);outline-offset:2px}
.comp-intro h2,.comp-intro .sub{display:grid}
.cv-h{grid-area:1/1}
.cv-h[data-v="terr"]{visibility:hidden}
.js .cv[data-v="terr"]{display:none}
.js .cv-nojs{display:none}
.cv-nojs{margin:48px 0 18px}
.js .competencies[data-view="terr"] .cv[data-v="comp"]{display:none}
.js .competencies[data-view="terr"] .cv-h[data-v="comp"]{visibility:hidden}
.js .competencies[data-view="terr"] .cv[data-v="terr"]{display:block}
.js .competencies[data-view="terr"] .cv-h[data-v="terr"]{visibility:visible}
.motion .cv.swap .comp{animation:fadeUp .55s var(--ease) both;animation-delay:var(--d,0ms)}
.sub{color:var(--mu);margin:22px 0 30px;max-width:460px}
.note{color:var(--mu2);font-size:13px;margin-top:12px}
.comp-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:24px}
.comp{position:relative;display:flex;flex-direction:column;border:1px solid var(--line);border-radius:22px;padding:30px;background:radial-gradient(120% 70% at 0 0,color-mix(in srgb,var(--c) 15%,transparent),transparent 55%),var(--card);overflow:hidden}
.comp::before{content:"";position:absolute;left:0;top:0;height:3px;width:100%;background:var(--c)}
.comp-head{display:flex;justify-content:space-between;gap:20px}
.ch-l{min-width:0}
.comp h3{font-size:clamp(26px,2.4vw,34px);line-height:1.1}
.lead{color:var(--mu);margin:8px 0 0;font-size:14px}
.lead b{color:var(--tx);font-weight:600}
.ch-r{flex:none;width:150px;text-align:right}
.hc{margin:0 0 10px;color:var(--mu);font-size:14px}
.hc b{font-family:var(--serif);font-weight:400;font-size:44px;line-height:1;color:var(--tx);margin-right:4px}
.pyr{display:grid;gap:3px}
.py2{display:grid;grid-template-columns:24px 1fr 22px;gap:6px;align-items:center;font-size:10px;color:var(--mu2)}
.py2 span:first-child{text-align:left;letter-spacing:.04em}
.py2 b{color:var(--tx);font-weight:600;font-size:11px;text-align:right}
.py2-b{display:flex;justify-content:center;height:9px;background:rgba(255,255,255,.04);border-radius:3px}
.py2-b i{display:block;height:100%;background:var(--c);border-radius:3px;transition:transform .9s var(--ease)}
.trend{margin-top:20px}
.tr-h{display:flex;justify-content:space-between;align-items:baseline;gap:10px;font-size:12px;color:var(--mu2)}
.tr-h .k{margin:0}
.tr-h b{color:var(--tx)}
.trend .spark{margin-top:6px}
.comp .k{margin-top:20px}
.meta{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin:auto 0 0;padding-top:20px;border-top:1px solid var(--line)}
.comp .chips{margin-bottom:22px}
/* Cards side by side share their rows, so each part (header, figures, trend, skills, offices)
   starts at the same height in both cards however tall the other card's header or chips run.
   Each part is pinned to its row, so a card without a part (no hours sheet) keeps the rest in line.
   Browsers without subgrid keep the plain column layout above. */
@supports (grid-template-rows:subgrid){
  .comp{display:grid;grid-row:span 6;grid-template-rows:subgrid;row-gap:0;align-content:start}
  .comp>.comp-head{grid-row:1}
  .comp>.dstrip{grid-row:2;align-self:start}
  .comp>.trend{grid-row:3;align-self:start}
  .comp>.k{grid-row:4;align-self:end}
  .comp>.chips{grid-row:5;align-self:start}
  .comp>.meta{grid-row:6;align-self:end}
}
.meta dt{font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--mu2)}
.meta dd{margin:6px 0 0;font-size:13px}

/* hero — grade stack, skill orbs */
.hero h1{font-family:var(--sans);text-transform:uppercase;font-weight:400;font-size:clamp(34px,5.1vw,78px);line-height:1.02;letter-spacing:-.02em}
.hero h1 em{font-style:normal;color:var(--or)}
.hero-in{z-index:2}
.hero-mark{margin:0 auto clamp(20px,3vh,34px);width:max-content;transform:translateY(calc(var(--sy,0) * -0.3px)) rotate(calc(var(--sy,0) * 0.05deg))}
.stack-stage{--sh:clamp(180px,calc(30vh - 20px),310px);position:relative;cursor:grab;touch-action:pan-y;-webkit-user-select:none;user-select:none}
.stack-stage.grabbing{cursor:grabbing}
/* the slabs are drawn at their resting angle by the builder, so this reads without JavaScript;
   STACK_JS redraws them frame by frame once it loads, so a drag can take the turn over */
.stack-stage svg{display:block;height:var(--sh);width:auto;overflow:visible;filter:drop-shadow(0 18px 34px rgba(0,0,0,.55))}
.stack-g path{stroke:rgba(0,0,0,.18);stroke-width:.5;stroke-linejoin:round}
.stack-g .stack-star{fill:#FFF6DC;stroke:none;filter:drop-shadow(0 0 3px rgba(255,182,0,.95)) drop-shadow(0 0 9px rgba(253,81,8,.55))}
.stack-glow{position:absolute;left:50%;bottom:-8%;width:160%;height:36%;transform:translateX(-50%);background:radial-gradient(ellipse,rgba(253,81,8,.38),transparent 70%);filter:blur(9px);pointer-events:none;z-index:-1}
.orbs{position:absolute;inset:0;z-index:1;pointer-events:none}
.orb{position:absolute;width:112px;height:112px;margin:-56px 0 0 -56px;animation:bob var(--fd,7s) ease-in-out var(--dl,0s) infinite alternate}
.orb-in{position:relative;width:100%;height:100%;transform-style:preserve-3d;transition:transform .9s var(--ease)}
.orb-in.flip{transform:rotateY(180deg)}
.face{position:absolute;inset:0;border-radius:50%;border:1px solid rgba(255,255,255,.2);background:radial-gradient(circle at 35% 28%,#26262b,#0e0e11 72%);box-shadow:0 14px 40px rgba(0,0,0,.55),inset 0 1px 0 rgba(255,255,255,.08);-webkit-backface-visibility:hidden;backface-visibility:hidden}
.face.back{transform:rotateY(180deg)}
.face .ic{position:absolute;inset:25%;background-size:contain;background-repeat:no-repeat;background-position:center}
.orb-l{position:absolute;top:100%;left:50%;transform:translateX(-50%);margin-top:10px;white-space:nowrap;font-size:12px;color:var(--mu);letter-spacing:.03em}
@keyframes bob{from{transform:translateY(-10px)}to{transform:translateY(12px)}}

/* leadership — MD at the centre, Directors, competencies */
.teams{padding:120px 0;background:var(--bg2);border-top:1px solid var(--line)}
.tm-grid{display:grid;grid-template-columns:minmax(0,5fr) minmax(0,7fr);gap:56px;align-items:start}
.tm-orbit-col{position:sticky;top:96px}
.orbit{position:relative;width:100%;max-width:660px;aspect-ratio:1/1;margin:0 auto 40px}
.wires{position:absolute;inset:0;width:100%;height:100%;overflow:visible}
.ring{fill:none;stroke:var(--line);stroke-width:2}
.ln{stroke:rgba(222,212,192,.55);stroke-width:3;fill:none;transition:opacity .45s}
.ln.dash{stroke-dasharray:18 14;animation:flow 1.8s linear infinite}
.ln.dot{stroke-dasharray:0 12;stroke-linecap:round;stroke-width:6}
@keyframes flow{to{stroke-dashoffset:-32}}
.hub{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);z-index:2}
.hub-stack{--sh:128px}
.hub.md{display:flex;flex-direction:column;align-items:center}
.md-avs{display:flex}
.md-avs .av{width:124px;height:124px;font-size:40px;border:3px solid var(--or);box-shadow:0 0 70px rgba(253,81,8,.45)}
.md-avs .av+.av{margin-left:-40px}
.hub.md:hover .av{transform:scale(1.04)}
.hub .pl{position:static;transform:none;margin-top:10px}
.av{display:inline-flex;align-items:center;justify-content:center;flex:none;border-radius:50%;overflow:hidden;background:radial-gradient(circle at 30% 25%,#35353c,#17171b 72%);background-size:cover;background-position:center 20%;color:var(--tx);font-family:var(--serif);transition:transform .3s var(--ease)}
.dn,.cn{position:absolute;left:var(--x);top:var(--y);transform:translate(-50%,-50%);transition:opacity .45s}
.dn{width:88px;height:88px;border-radius:50%}
.dn .av{width:100%;height:100%;font-size:28px;border:2px solid var(--c);animation:ping 2.8s ease-out infinite}
.dn:hover .av,.dn:focus-visible .av,.dn.sel .av{transform:scale(1.08)}
@keyframes ping{0%{box-shadow:0 0 0 0 color-mix(in srgb,var(--c) 50%,transparent)}70%,100%{box-shadow:0 0 0 20px transparent}}
.cn{width:60px;height:60px}
.cn-n{position:absolute;top:100%;left:50%;transform:translateX(-50%);margin-top:34px;font-size:11px;color:var(--mu);white-space:nowrap;text-align:center}
.cn-n b{display:block;color:var(--tx);font-weight:600;font-size:11.5px}
/* a competency's Senior Managers, circling it and staying upright as they go */
.cn-orb{position:absolute;inset:0;animation:spin var(--spin,24s) linear infinite;pointer-events:none}
.cn-sat{position:absolute;left:50%;top:50%;width:30px;height:30px;margin:-15px 0 0 -15px;pointer-events:auto;
  transform:rotate(var(--a)) translateX(50px) rotate(calc(-1 * var(--a)))}
.cn-sat .av{width:100%;height:100%;font-size:11px;font-family:var(--sans);font-weight:600;border:1.5px solid var(--c);
  animation:spin var(--spin,24s) linear infinite reverse}
.cn:hover .cn-orb,.cn:hover .cn-sat .av{animation-play-state:paused}
@keyframes spin{to{transform:rotate(360deg)}}
.hub.md .pl .md-name{color:var(--or);font-size:18px;letter-spacing:.01em;text-shadow:0 0 18px rgba(253,81,8,.45)}
.pl i,.pl em{display:block;font-style:normal;font-size:11px;line-height:1.35}
.pl i{color:#cfcfd4}
.pl em{color:var(--mu2);margin-top:2px}
.cl-code{display:none}
.orbit-legend{max-width:660px;margin:22px auto 0}
.orbit-legend .k{margin-top:0}
.orbit-legend .chip{align-items:center}
.orbit-legend .chip b{font-weight:700}
.dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:4px}
.cn-b{display:flex;align-items:center;justify-content:center;width:100%;height:100%;border-radius:50%;border:2px solid var(--c);background:radial-gradient(circle at 35% 28%,color-mix(in srgb,var(--c) 40%,#1b1b20),#121215 76%);font-weight:700;font-size:13px;letter-spacing:.04em;color:#fff;transition:transform .3s var(--ease)}
.cn:hover .cn-b,.cn:focus-visible .cn-b{transform:scale(1.08)}
.pl{position:absolute;top:100%;left:50%;transform:translateX(-50%);margin-top:8px;text-align:center;white-space:nowrap;font-size:12px;color:var(--mu);line-height:1.3;background:rgba(17,17,20,.82);padding:2px 7px;border-radius:7px;z-index:1}
.pl b{display:block;color:var(--tx);font-size:14px;font-weight:600}
.orbit .dim{opacity:.16!important}
.motion .orbit .dn,.motion .orbit .cn{left:50%;top:50%;opacity:0;transition:left .9s var(--ease) var(--d,0ms),top .9s var(--ease) var(--d,0ms),opacity .5s ease var(--d,0ms)}
.motion .orbit.in .dn,.motion .orbit.in .cn{left:var(--x);top:var(--y);opacity:1}
.motion .orbit .ln{opacity:0;transition:opacity .8s ease .6s}
.motion .orbit.in .ln{opacity:1}
.motion .orbit.settled .dn,.motion .orbit.settled .cn,.motion .orbit.settled .ln{transition:opacity .45s}

.tm-view h2{font-size:clamp(34px,4vw,54px);line-height:1.06}
.picks{display:grid;gap:10px}
.pick{display:flex;align-items:center;gap:14px;padding:12px 16px;border:1px solid var(--line);border-radius:16px;background:var(--card);transition:border-color .2s,transform .25s var(--ease)}
.pick:hover,.pick:focus-visible{border-color:var(--c);transform:translateX(4px)}
.pick .av{width:48px;height:48px;font-size:17px;border:2px solid var(--c)}
.pick b{display:block;font-weight:600}
.pick .s{display:block;font-size:13px;color:var(--mu)}
.pick .pn{min-width:0;flex:1}
/* three labelled figures, each in its own column, so a count never runs into a percentage */
/* the team, large, with its utilization small underneath: one figure per line, nothing side by side */
.pick .ps{margin-left:auto;text-align:right;white-space:nowrap}
.pick .ps b{display:inline;font-family:var(--serif);font-weight:400;font-size:30px;line-height:1}
.pick .ps small{margin-left:5px;font-size:11px;color:var(--mu2);letter-spacing:.06em;text-transform:uppercase}
.pick .ps i{display:block;font-style:normal;margin-top:5px;font-size:12px;color:var(--mu)}
.back{display:inline-flex;align-items:center;gap:8px;font-size:14px;color:var(--mu);border:1px solid var(--line2);border-radius:999px;padding:8px 15px;margin-bottom:26px;transition:color .2s,border-color .2s}
.back:hover{color:var(--tx);border-color:var(--c)}
.tm-head{display:flex;gap:18px;align-items:center}
.tm-head .av{width:78px;height:78px;font-size:26px;border:2px solid var(--c)}
.tm-head h3{font-size:clamp(32px,3.4vw,46px);line-height:1}
.tm-head p{margin:8px 0 0;color:var(--mu)}
.tm-head p b{color:var(--tx);font-weight:600}
blockquote{margin:24px 0 0;padding-left:20px;border-left:2px solid var(--c);font-family:var(--serif);font-style:italic;font-size:18px;line-height:1.5;color:#d6d6da}
.chip .ci{width:14px;height:14px;align-self:center;background-size:contain;background-repeat:no-repeat;background-position:center}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--line);border:1px solid var(--line);border-radius:16px;overflow:hidden;margin:26px 0 6px}
.kpis div{background:var(--card);padding:14px}
.kpis div:last-child:nth-child(4n+3){grid-column:span 2}
.kpis b{display:block;font-family:var(--serif);font-weight:400;font-size:28px;line-height:1}
.kpis span{display:block;margin-top:8px;font-size:12px;color:var(--mu2)}
.spark{display:block;width:100%;height:auto;overflow:visible}
.spark .sp-lb{font-size:7px;fill:var(--mu);font-family:var(--sans)}
.spark .sp-lb.last{fill:var(--tx);font-weight:700}
.sp-ref{stroke:var(--line2);stroke-dasharray:4 5}
.sp-t{display:flex;justify-content:space-between;font-size:11px;color:var(--mu2);margin-top:6px}
.sp-t b{color:var(--tx)}
.an{display:grid;grid-template-columns:1fr 1fr;gap:0 28px}
.mb{list-style:none;margin:0;padding:0}
.mb li{display:grid;grid-template-columns:minmax(0,110px) 1fr 24px;gap:10px;align-items:center;font-size:13px;padding:4px 0;color:var(--mu)}
.mb li span:first-child{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mb li b{color:var(--tx);text-align:right;font-weight:600}
.mb .bar{height:6px;background:rgba(255,255,255,.06);border-radius:4px;overflow:hidden}
.mb .bar i{display:block;height:100%;background:var(--c);border-radius:4px;transform-origin:0 50%}
.dstrip{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--line);border:1px solid var(--line);border-radius:14px;overflow:hidden;margin-top:22px}
.dstrip div{background:rgba(21,21,25,.94);padding:12px 14px}
.dstrip b{display:block;font-family:var(--serif);font-weight:400;font-size:24px;line-height:1}
.dstrip span{display:block;margin-top:6px;font-size:12px;color:var(--mu2)}
.js .tm-view{display:none}
.js .tm-view.on{display:block}
.motion .tm-view.on{animation:fadeUp .6s var(--ease) both}
.motion .tm-view.on .mb .bar i{animation:grow 1s var(--ease) .15s both}
@keyframes fadeUp{from{opacity:0;transform:translateY(18px)}}
@keyframes grow{from{transform:scaleX(0)}}

/* team shape */
.shape{padding:120px 0}
.sec-head{margin-bottom:40px}
.sec-head .sub{margin-top:14px;max-width:620px}
.shape-grid{display:grid;grid-template-columns:minmax(0,5fr) minmax(0,7fr);gap:24px;align-items:stretch}
.sh-card{min-width:0;display:flex;flex-direction:column;border:1px solid var(--line);border-radius:22px;padding:26px 28px 28px;background:var(--card)}
.sh-card .k{margin-top:22px}
.sh-card .k:first-child{margin-top:0}
.pr{display:grid;grid-template-columns:130px 1fr 28px 38px;align-items:center;gap:14px;padding:7px 0}
.pr-l{color:var(--mu);font-size:13px}
.pr-b{display:flex;justify-content:center;height:26px;background:rgba(255,255,255,.035);border-radius:6px}
.pr-b i{display:block;height:100%;border-radius:6px;background:linear-gradient(90deg,var(--or),var(--tg));transition:transform 1.1s var(--ease)}
.pr b{text-align:right;font-family:var(--serif);font-weight:400;font-size:20px}
.pr-s{color:var(--mu2);font-size:12px;text-align:right}
.lev{margin:auto 0 0;padding-top:16px;border-top:1px solid var(--line);color:var(--mu);font-size:14px}
/* Value: two cards of the same width as the tiles above, with the grade list and the month chart
   filling the same band, so both takeaway lines start on one baseline. */
/* same 12px gutter as the tiles above, so the split between the cards falls exactly on the
   boundary between the second and third tile */
.value-sec .shape-grid{grid-template-columns:1fr 1fr;gap:12px;margin-top:12px}
.value-sec .val-rows{flex:1;display:flex;flex-direction:column;justify-content:space-evenly;padding:4px 0}
.value-sec .val-rows .chart{margin:auto 0}
/* heading / content / takeaway share their row heights across both cards, so the two sides line up
   whatever the text length. Without subgrid the flex fallback above still bottoms the takeaway out. */
@supports (grid-template-rows:subgrid){
  .value-sec .shape-grid{grid-template-rows:auto 1fr auto}
  .value-sec .sh-card{display:grid;grid-row:span 3;grid-template-rows:subgrid}
  .value-sec .lev{margin:0}
}

/* Team shape — the pyramid rows are the filter */
.pr[data-ts]{width:100%;text-align:left;background:none;border:0;font:inherit;color:inherit;cursor:pointer;
  border-radius:10px;margin:0 -10px;padding:7px 10px;width:calc(100% + 20px);transition:background .18s}
.pr[data-ts]:hover{background:rgba(255,255,255,.045)}
.pr[data-ts]:focus-visible{outline:2px solid var(--tg);outline-offset:2px}
.pr[data-ts][aria-pressed="true"]{background:rgba(253,81,8,.12)}
.pr[data-ts][aria-pressed="true"] .pr-l{color:var(--tx);font-weight:600}
.ts-hint{color:var(--mu2)}
.ts-notes{margin-top:18px;min-height:1.6em}
.ts-note{margin:0;color:var(--mu);font-size:14px}
.ts-note b{color:var(--tx)}
.ts-clear{background:none;border:0;padding:0 0 1px;margin-left:6px;font:inherit;color:var(--tg);cursor:pointer;border-bottom:1px solid rgba(235,140,0,.45)}
.ts-clear:hover{color:var(--yl)}
html:not(.js) .ts-hint{display:none}
/* every header and every filled cell of the grid is a filter too */
.hm-h{background:none;border:0;padding:2px 4px;margin:-2px -4px;font:inherit;color:inherit;text-align:inherit;
  cursor:pointer;border-radius:6px;transition:color .15s,background .15s}
.hm-h:hover{color:var(--tx);background:rgba(255,255,255,.06)}
.hm-h:focus-visible,button.hm:focus-visible{outline:2px solid var(--tg);outline-offset:2px}
.hm-h.is-on{color:var(--tx);font-weight:700;background:rgba(253,81,8,.14)}
button.hm{width:100%;border:0;font:inherit;font-weight:600;color:#fff;cursor:pointer}
button.hm:hover{box-shadow:0 0 0 2px rgba(255,255,255,.35) inset}
button.hm.zero{background:rgba(255,255,255,.04);color:var(--mu2)}
button.hm.zero:hover{background:rgba(255,255,255,.09);color:var(--mu)}
button.hm.is-on{box-shadow:0 0 0 2px var(--yl) inset}
.shape.ts-filtered .hm.zero{opacity:.5}
.ts-sub{color:var(--mu2)}
.exp-rows{display:grid;gap:6px}
.lev b{color:var(--tx);font-family:var(--serif);font-size:19px;font-weight:400}
.hm-wrap{min-width:0;overflow-x:auto;-webkit-overflow-scrolling:touch}
.hm-t{border-collapse:collapse;width:100%;min-width:420px;table-layout:fixed;font-size:13px}
.hm-c0{width:34%}
.hm-ct{width:9%}
.hm-t th,.hm-t td{padding:5px 6px;text-align:center;font-weight:400}
.hm-t thead th{font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--mu2);padding-bottom:10px}
.hm-t th[scope=row]{text-align:left;color:var(--mu);padding-right:14px;line-height:1.3}
.hm{display:block;border-radius:7px;padding:9px 0;color:#fff;font-weight:600;background:color-mix(in srgb,var(--or) calc(var(--a) * 100%),transparent);transition:transform .5s var(--ease)}
.hm.zero{background:rgba(255,255,255,.03);color:var(--mu2);font-weight:400}
.hm-t .tot{color:var(--mu);font-weight:600}
.hm-t tfoot th,.hm-t tfoot td{border-top:1px solid var(--line);padding-top:12px;color:var(--mu)}
.offices{margin-top:24px}
.off-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px 28px}
.off b{min-width:22px}
.off{display:grid;grid-template-columns:minmax(0,1.4fr) minmax(50px,2fr) auto 38px;align-items:center;gap:10px;padding:3px 0;font-size:13px;color:var(--mu)}
.off span:first-child{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.off b{color:var(--tx);text-align:right;font-weight:600}
.off-b{height:6px;background:rgba(255,255,255,.05);border-radius:4px;overflow:hidden}
.off-b i{display:block;height:100%;background:var(--tg);border-radius:4px;transform-origin:0 50%;transition:transform 1s var(--ease)}

/* practice-level sections: capacity, skills, value, directory */
.tiles{display:grid;grid-template-columns:repeat(var(--tc,4),minmax(0,1fr));gap:12px;margin-top:22px}
.tile{border:1px solid var(--line);border-radius:16px;padding:18px 20px;background:var(--card)}
.tile b{display:block;font-family:var(--serif);font-weight:400;font-size:34px;line-height:1}
.tile span{display:block;margin-top:8px;font-size:13px;color:var(--mu)}
.tile small{display:block;margin-top:4px;font-size:11px;color:var(--mu2)}
.chart{width:100%;height:auto;display:block;margin-top:10px}
.cx{stroke:var(--line)}
.cy{fill:var(--mu2);font-size:12px}
.bch{fill:var(--or);opacity:.9}
.dl{font-size:12px;font-weight:600;font-family:var(--sans)}
.dl-ch{fill:rgba(255,255,255,.92);font-size:11px}
.dl-nc{fill:var(--mu);font-size:11px}
.dl-u{fill:#fff;font-size:13px;font-weight:700}
.dl-v{fill:var(--mu);font-size:11px}
.lgd{display:flex;gap:20px;flex-wrap:wrap;font-size:12px;color:var(--mu2)}
.lg{display:inline-flex;align-items:center;gap:7px;color:var(--mu2)}
.lg::before{content:"";width:10px;height:10px;border-radius:3px;background:var(--dot,var(--or))}
.lg-ch{--dot:var(--or)}
.lg-nc{--dot:#4a4a52}
.lg-pct::before{display:none}
.lg-ut::before{height:2px!important;width:16px!important;border-radius:1px!important;background:rgba(255,255,255,.85)!important}
.cy-u{fill:var(--mu2);opacity:.75}
.ut-l{stroke:rgba(255,255,255,.85);stroke-width:2;stroke-linejoin:round;stroke-linecap:round}
.ut-d{fill:#16161A;stroke:#fff;stroke-width:1.8}
.motion .chart .ut{opacity:0;transition:opacity .8s var(--ease) .7s}
.motion .chart.in .ut{opacity:1}
.lg-pct b{color:#fff;font-weight:700}
.bnc{fill:#3a3a42}
.motion .chart .bch,.motion .chart .bnc{transform:scaleY(0);transform-origin:50% 100%}
.motion .chart.in .bch,.motion .chart.in .bnc{transform:none;transition:transform .9s var(--ease)}

.mix-rows{display:grid;grid-template-columns:repeat(2,1fr);gap:14px 18px;margin-top:4px}
.mx b{display:block;font-family:var(--serif);font-weight:400;font-size:26px;line-height:1}
.mx span{display:block;margin-top:5px;font-size:12px;color:var(--mu2)}
.off-grid.one{grid-template-columns:1fr}

/* closing */
.closing{position:relative;text-align:center;padding:170px 0;overflow:hidden;border-top:1px solid var(--line)}
.closing .wrap{position:relative}
.closing h2{font-size:clamp(40px,6vw,84px);line-height:1.02}
.foot{border-top:1px solid var(--line);padding:30px 0;color:var(--mu2);font-size:13px}

/* animation states — only when JS has confirmed it can run them (html.motion) */
.motion .rv{opacity:0;transform:translateY(34px);transition:opacity .9s var(--ease),transform .9s var(--ease);transition-delay:var(--d,0ms)}
.motion .rv.in{opacity:1;transform:none}
.motion .w{opacity:.16}
.motion .w.on{opacity:1}
.motion .py2-b i,.motion .off-b i,.motion .pr-b i{transform:scaleX(0)}
.motion .in .py2-b i,.motion .rv.in.off .off-b i,.motion .rv.in.pr .pr-b i{transform:scaleX(1)}
.motion .hm-wrap .hm{transform:scale(.8);opacity:0}
.motion .hm-wrap.in .hm{transform:none;opacity:1;transition:transform .5s var(--ease),opacity .5s}

/* responsive */
@media (max-width:900px){
  .wrap{padding:0 16px}
  .links{display:none}
  .top-in{justify-content:space-between;padding:0 16px}
  .brand span+span{display:none}
  .shape-grid,.comp-grid,.value-sec .shape-grid{grid-template-columns:1fr;gap:24px}
  /* one card per row, so there is nothing to line the two of them up against */
  .value-sec .shape-grid{grid-template-rows:none}
  .value-sec .sh-card{display:flex;grid-row:auto}
  .value-sec .lev{margin:auto 0 0}
  .comp-intro{flex-direction:column;align-items:flex-start;gap:20px;margin-bottom:32px}
  .comp{padding:22px}
  .ch-r{width:120px}
  .hc b{font-size:34px}
  .kp-grid{grid-template-columns:repeat(2,1fr)}
  .kp,.kp:first-child{padding:20px 0 6px;border-left:0}
  .kp:nth-child(even){padding-left:16px;border-left:1px solid var(--line)}
  .kp b{font-size:40px}
  .kp span{font-size:11px}
  .meta{grid-template-columns:1fr}
  .pr{grid-template-columns:96px 1fr 24px 32px;gap:8px}
  .pr-l{font-size:12px}
  .sh-card{padding:20px}
  .off-grid,.off-grid.one{grid-template-columns:1fr}
  .tiles{grid-template-columns:repeat(2,minmax(0,1fr))}
  .tile b{font-size:28px}
  .cy{font-size:26px}
  .chart .alt,.chart .dl{display:none}
  .mix-rows{grid-template-columns:1fr}
  .hm-t{font-size:12px}
  .statement{padding:14vh 0}
  .kpis,.dstrip{grid-template-columns:repeat(2,1fr)}
  .kpis b{font-size:26px}
  .dstrip b{font-size:24px}
  .teams{overflow:hidden}
  .tm-grid{grid-template-columns:1fr;gap:28px}
  .tm-orbit-col{position:static;order:-1}
  .orbit{max-width:420px}
  .dn{width:54px;height:54px}.dn .av{font-size:18px}
  .cn{width:40px;height:40px}.cn-b{font-size:9px}.cn-n{font-size:9px;margin-top:3px}
  .md-avs .av{width:66px;height:66px;font-size:22px}
  .pl{font-size:9px;margin-top:4px;padding:1px 5px}.pl b{font-size:11px}
  .dn .pl{padding:1px 5px}.dn .pl b{font-size:10px}
  .cl-full,.pl em{display:none}.cl-code{display:block;font-size:9px;color:var(--mu)}
  .stack-stage{--sh:152px}
  .hub-stack{--sh:112px}
  .an{grid-template-columns:1fr}
  .orb{width:62px;height:62px;margin:-31px 0 0 -31px}
  .orb-l{display:none}
  .orb:nth-child(1){left:14%!important;top:13%!important}
  .orb:nth-child(2){left:86%!important;top:17%!important}
  .orb:nth-child(3){left:13%!important;top:90%!important}
  .orb:nth-child(4){left:87%!important;top:88%!important}
  .orb:nth-child(n+5){display:none}
  .hero-mark{margin-bottom:22px}
}
@media (prefers-reduced-motion:reduce){
  html{scroll-behavior:auto}
  .g,.mq-track,.cue span,.orb,.ln.dash,.dn .av,.cn-orb,.cn-sat .av{animation:none}
}

/* the full dashboard, embedded in this same file — see dash_panel() */
.dpanel{position:fixed;inset:0;z-index:300;background:var(--bg);display:none}
.dpanel.on{display:block}
/* the frame is positioned so it paints — and takes clicks — above the waiting message below it */
.dpanel iframe{display:block;position:relative;z-index:1;width:100%;height:100%;border:0;background:var(--bg)}
.dpanel-wait{position:absolute;inset:0;display:grid;place-items:center;color:var(--mu);font-size:14px;letter-spacing:.02em}
html.dpanel-on{overflow:hidden}
html:not(.js) [data-dash]{display:none}
"""

def dash_panel(dash_src):
    """The full dashboard, carried in this file and mounted on first use.

    It keeps its own stylesheet and script, so it goes in an iframe rather than in the page:
    the two builds share class names (.card, .panel, .btn) and would otherwise repaint each
    other. The source sits in a text/plain block — inert until something asks for it — so the
    landing page still opens instantly. "</script" is escaped so it cannot close that block.
    """
    return ('<div class="dpanel" id="dpanel" aria-hidden="true">'
            '<p class="dpanel-wait">Opening the dashboard&hellip;</p></div>\n'
            '<script type="text/plain" id="dash-src">'
            + dash_src.replace("</script", "<\\/script") + "</script>")


STACK_JS = r"""
(function(){
  var D = window.__STACK, stages = [].slice.call(document.querySelectorAll('.stack-stage'));
  if (!D || !D.t || !stages.length) return;
  var motion = /\bmotion\b/.test(document.documentElement.className);
  var L = D.light, LL = Math.sqrt(L[0]*L[0] + L[1]*L[1] + L[2]*L[2]);
  var NORM = [[0,0,-1],[1,0,0],[0,0,1],[-1,0,0]];

  function shade(hex, k){
    return '#' + [1,3,5].map(function(i){
      var c = Math.max(0, Math.min(255, Math.round(parseInt(hex.substr(i, 2), 16) * k)));
      return (c < 16 ? '0' : '') + c.toString(16);
    }).join('');
  }
  function path(pts){ return 'M' + pts.map(function(p){ return p[0].toFixed(1) + ',' + p[1].toFixed(1); }).join('L') + 'Z'; }

  /* the same maths as stack_polys() in the builder: every visible face, in painting order */
  function polys(ry, rx){
    var th = ry * Math.PI / 180, ph = rx * Math.PI / 180,
        ct = Math.cos(th), st = Math.sin(th), cp = Math.cos(ph), sp = Math.sin(ph), out = [];
    function P(x, y, z){ var z1 = -x * st + z * ct; return [x * ct + z * st, -(y * cp - z1 * sp)]; }
    function lit(n, c, top){
      var x1 = n[0] * ct + n[2] * st, z1 = -n[0] * st + n[2] * ct,
          vx = x1, vy = n[1] * cp - z1 * sp, vz = n[1] * sp + z1 * cp;
      if (vz <= 1e-6) return null;
      var d = Math.max(0, (vx * L[0] + vy * L[1] + vz * L[2]) / LL);
      return shade(c, (top ? 0.62 : 0.5) + 0.5 * d);
    }
    var layers = [], i, n = D.t.length + (D.star ? 1 : 0);
    for (i = 0; i < n; i++) layers.push(i);
    if (sp < 0) layers.reverse();
    layers.forEach(function(i){
      if (i === D.t.length){
        var c = P(0, D.star.y, 0), r = D.star.r, pts = [];
        for (var k = 0; k < 10; k++){
          var rr = k % 2 ? r * 0.42 : r;
          pts.push([c[0] + rr * Math.sin(k * Math.PI / 5), c[1] - rr * Math.cos(k * Math.PI / 5)]);
        }
        out.push([pts, 'star']);
        return;
      }
      var s = D.t[i], w = s.w, q = [[-w,-w],[w,-w],[w,w],[-w,w]], f, fill;
      for (f = 0; f < 4; f++){
        fill = lit(NORM[f], s.c, false);
        if (fill){
          var a = q[f], b = q[(f + 1) % 4];
          out.push([[P(a[0], s.y0, a[1]), P(b[0], s.y0, b[1]), P(b[0], s.y1, b[1]), P(a[0], s.y1, a[1])], fill]);
        }
      }
      [[[0,1,0], s.y1], [[0,-1,0], s.y0]].forEach(function(t){
        fill = lit(t[0], s.c, true);
        if (fill) out.push([q.map(function(c){ return P(c[0], t[1], c[1]); }), fill]);
      });
    });
    return out;
  }

  function norm(a){ return ((a % 360) + 540) % 360 - 180; }   /* -180..180 */

  stages.forEach(function(stage){
    var g = stage.querySelector('.stack-g');
    if (!g) return;
    var title = (g.querySelector('.stack-star title') || {}).textContent || '';
    /* the angles the SVG was drawn at, so nothing jumps when this takes over */
    var ry = D.ry, rx = D.rx, spin = 20 / 1000, vy = 0, drag = null, last = 0;

    function apply(){
      g.innerHTML = polys(ry, rx).map(function(p){
        return p[1] === 'star'
          ? '<path class="stack-star" d="' + path(p[0]) + '"><title>' + title + '</title></path>'
          : '<path d="' + path(p[0]) + '" fill="' + p[1] + '"/>';
      }).join('');
    }

    function frame(t){
      var dt = last ? Math.min(t - last, 64) : 0, moved = false;
      last = t;
      if (!drag){
        if (Math.abs(vy) > 0.001){            /* the throw you gave it, easing back to the idle turn */
          ry += vy * dt;
          vy *= Math.pow(0.9, dt / 16);
          moved = true;
        } else if (motion){
          ry += spin * dt;
          moved = true;
        }
        var off = norm(D.rx - rx);            /* tipped over: settle back to the resting tilt, the short way */
        if (Math.abs(off) > 0.05){
          rx += off * (1 - Math.pow(0.9, dt / 16));
          moved = true;
        } else if (off){ rx = D.rx; moved = true; }
      }
      if (moved) apply();
      requestAnimationFrame(frame);
    }
    requestAnimationFrame(frame);

    stage.addEventListener('pointerdown', function(ev){
      /* a mouse or pen turns it both ways; a finger only turns it, so a vertical swipe still scrolls */
      drag = {x: ev.clientX, y: ev.clientY, t: ev.timeStamp, dx: 0, tilt: ev.pointerType !== 'touch'};
      vy = 0;
      stage.classList.add('grabbing');
      if (stage.setPointerCapture) { try { stage.setPointerCapture(ev.pointerId); } catch (e) {} }
    });
    stage.addEventListener('pointermove', function(ev){
      if (!drag) return;
      var dx = ev.clientX - drag.x, dy = ev.clientY - drag.y, dt = Math.max(ev.timeStamp - drag.t, 1);
      ry += dx * 0.45;
      if (drag.tilt) rx += dy * 0.45;
      drag.dx = dx / dt * 0.45;                 /* degrees per ms, for the release */
      drag.x = ev.clientX; drag.y = ev.clientY; drag.t = ev.timeStamp;
      apply();
      if (ev.cancelable) ev.preventDefault();   /* pan-y still lets the page scroll vertically */
    });
    function release(){
      if (!drag) return;
      vy = Math.max(-1.2, Math.min(1.2, drag.dx));
      drag = null;
      stage.classList.remove('grabbing');
    }
    stage.addEventListener('pointerup', release);
    stage.addEventListener('pointercancel', release);
    stage.addEventListener('lostpointercapture', release);
  });
})();
"""

PACK_JS = r"""
(function(){
  /* Skill chips keep their rank order, but a chip too wide for the rest of a line no longer leaves
     that space empty: the next chip that does fit moves up into it (first fit, row by row). */
  var boxes = [].slice.call(document.querySelectorAll('.comp .chips, .tm-view .chips'));
  if (!boxes.length || !window.ResizeObserver) return;
  function pack(box){
    var W = box.clientWidth;
    if (!W) return;                                   /* hidden (the other toggle view): packed when shown */
    var gap = parseFloat(getComputedStyle(box).columnGap) || 0, n = 0;
    var left = [].slice.call(box.children).map(function(el, i){
      if (el._rank == null) el._rank = i;
      return {el: el, w: el.getBoundingClientRect().width};
    }).sort(function(a, b){ return a.el._rank - b.el._rank; });
    while (left.length){
      var used = 0;
      left = left.filter(function(c){
        var need = used ? used + gap + c.w : c.w;
        if (used && need > W + 0.5) return true;          /* does not fit this line: try it on the next */
        used = need; c.el.style.order = n++;
        return false;
      });
    }
  }
  var ro = new ResizeObserver(function(es){ es.forEach(function(en){ pack(en.target); }); });
  boxes.forEach(function(b){ ro.observe(b); });
})();
"""

SHAPE_JS = r"""
(function(){
  var sec = document.getElementById('shape'), D = window.__SHAPE;
  if (!sec || !D || !D.p) return;
  var $ = function(id){ return document.getElementById(id); };
  var exp = $('ts-exp'), grid = $('ts-grid'), off = $('ts-off'), ten = $('ts-ten'),
      mix = $('ts-mix'), note = $('ts-note');
  if (!exp || !grid || !off || !ten || !mix || !note) return;

  var N = D.p.length, active = null;          /* null = the whole practice */
  var G = 0, C = 1, T = 2, O = 3, E = 4, PW = 5, SEX = 6, REPS = 7;

  function esc(s){ return String(s).replace(/[&<>"]/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  function pct(n, of){ return of ? Math.round(n / of * 100) : 0; }

  /* who the current filter covers */
  function match(r){
    if (!active) return true;
    if (active.k === 'g') return r[G] === active.v;
    if (active.k === 'c') return r[C] === active.v;
    if (active.k === 't') return r[T] === active.v;
    return r[C] === active.c && r[T] === active.t;     /* one cell */
  }
  function subset(){ return D.p.filter(match); }

  function label(){
    if (!active) return '';
    if (active.k === 'g'){
      for (var i = 0; i < D.g.length; i++) if (D.g[i][0] === active.v) return D.g[i][1];
      return active.v;
    }
    if (active.k === 'c') return D.c[active.v];
    if (active.k === 't') return D.t[active.v];
    return D.c[active.c] + ' × ' + D.t[active.t];
  }

  /* --- the blocks, same markup as the Python build so the styling is shared --- */
  function rowsHtml(pairs, n){
    var mx = 1, i;
    for (i = 0; i < pairs.length; i++) mx = Math.max(mx, pairs[i][1]);
    return pairs.map(function(pr, i){
      return '<div class="off rv in" style="--d:' + (i * 50) + 'ms"><span>' + esc(pr[0]) + '</span>' +
             '<span class="off-b"><i style="width:' + Math.round(pr[1] / mx * 100) + '%"></i></span>' +
             '<b>' + pr[1] + '</b><span class="pr-s">' + pct(pr[1], n) + '%</span></div>';
    }).join('');
  }
  function bands(rows, field){
    return D.b.map(function(b){
      return [b[0], rows.filter(function(r){ return r[field] !== null && r[field] >= b[1] && r[field] < b[2]; }).length];
    });
  }
  function offices(rows){
    var out = [];
    D.o.forEach(function(name, i){
      var n = rows.filter(function(r){ return r[O] === i; }).length;
      if (n) out.push([name, n]);
    });
    return out;
  }
  function mixHtml(rows){
    var n = rows.length,
        joined = rows.filter(function(r){ return r[PW] !== null && r[PW] < 1; }).length,
        m = rows.filter(function(r){ return r[SEX] === 'm'; }).length,
        f = rows.filter(function(r){ return r[SEX] === 'f'; }).length,
        exps = rows.filter(function(r){ return r[E] !== null; }),
        leads = rows.filter(function(r){ return r[REPS]; }).length,
        ratio = (m + f) ? Math.round(m / (m + f) * 100) + '% : ' + (100 - Math.round(m / (m + f) * 100)) + '%' : '—',
        avg = exps.length ? (exps.reduce(function(a, r){ return a + r[E]; }, 0) / exps.length).toFixed(1) : null,
        tile = function(v, l){ return '<div class="mx"><b>' + v + '</b><span>' + l + '</span></div>'; };
    return '<div class="mix-rows">' + tile(joined, 'joined in the last 12 months') + tile(ratio, 'male : female') +
      (avg === null ? '' : tile(avg, 'average years of experience')) +
      tile(leads, 'people with direct reports') + '</div>';
  }

  /* every pairing that has anyone in the practice stays clickable whatever the current filter,
     so you can move straight from one cell to another instead of clearing first */
  var base = {};
  D.p.forEach(function(r){ if (r[C] >= 0 && r[T] >= 0) base[r[C] + ':' + r[T]] = (base[r[C] + ':' + r[T]] || 0) + 1; });

  /* the pyramid is the control, so it holds still while a grade is picked; filter by a competency,
     a territory or a cell and it becomes a readout too — the grade split of those people */
  var bars = [].slice.call(sec.querySelectorAll('.pr[data-ts]'));
  function paintPyramid(rows){
    var byGrade = {}, live = active && active.k !== 'g', mx = 1;
    (live ? rows : D.p).forEach(function(r){ byGrade[r[G]] = (byGrade[r[G]] || 0) + 1; });
    bars.forEach(function(b){ mx = Math.max(mx, byGrade[b.getAttribute('data-ts').slice(2)] || 0); });
    var of = live ? rows.length : N;
    bars.forEach(function(b){
      var n = byGrade[b.getAttribute('data-ts').slice(2)] || 0;
      b.querySelector('b').textContent = n;
      b.querySelector('.pr-s').textContent = pct(n, of) + '%';
      b.querySelector('.pr-b i').style.width = Math.max(n / mx * 100, n ? 1.5 : 0).toFixed(1) + '%';
    });
  }

  /* the grid is both a readout and the control: the counts follow the filter */
  function gridHtml(rows){
    var cnt = {}, top = 1, ci, ti, n;
    rows.forEach(function(r){ if (r[C] >= 0 && r[T] >= 0) cnt[r[C] + ':' + r[T]] = (cnt[r[C] + ':' + r[T]] || 0) + 1; });
    for (ci = 0; ci < D.c.length; ci++) for (ti = 0; ti < D.t.length; ti++) top = Math.max(top, cnt[ci + ':' + ti] || 0);
    var on = function(k, v){ return active && active.k === k && active.v === v ? ' is-on' : ''; };
    var head = D.t.map(function(t, i){
      return '<th><button type="button" class="hm-h' + on('t', i) + '" data-ts="t:' + i + '">' + esc(t) + '</button></th>';
    }).join('');
    var body = D.c.map(function(c, ci){
      var tds = D.t.map(function(t, ti){
        var n = cnt[ci + ':' + ti] || 0,
            sel = active && active.k === 'x' && active.c === ci && active.t === ti ? ' is-on' : '',
            key = ' data-ts="x:' + ci + ':' + ti + '"';
        if (n) return '<td><button type="button" class="hm' + sel + '"' + key +
                      ' style="--a:' + (0.12 + 0.88 * (n / top)).toFixed(2) + '">' + n + '</button></td>';
        /* nobody here under this filter, but the pairing exists — keep it reachable */
        return base[ci + ':' + ti]
          ? '<td><button type="button" class="hm zero' + sel + '"' + key + '>·</button></td>'
          : '<td><span class="hm zero">·</span></td>';
      }).join('');
      var tot = D.t.reduce(function(a, t, ti){ return a + (cnt[ci + ':' + ti] || 0); }, 0);
      return '<tr><th scope="row"><button type="button" class="hm-h' + on('c', ci) + '" data-ts="c:' + ci + '">' +
             esc(c) + '</button></th>' + tds + '<td class="tot">' + tot + '</td></tr>';
    }).join('');
    var foot = D.t.map(function(t, ti){
      return '<td class="tot">' + D.c.reduce(function(a, c, ci){ return a + (cnt[ci + ':' + ti] || 0); }, 0) + '</td>';
    }).join('');
    var all = Object.keys(cnt).reduce(function(a, k){ return a + cnt[k]; }, 0);
    var cols = D.t.length ? '<colgroup><col class="hm-c0"><col span="' + D.t.length + '" class="hm-cx"><col class="hm-ct"></colgroup>' : '';
    return '<table class="hm-t" style="--cols:' + Math.max(D.t.length, 1) + '">' + cols +
           '<thead><tr><th></th>' + head + '<th class="tot">All</th></tr></thead><tbody>' + body + '</tbody>' +
           '<tfoot><tr><th scope="row">All competencies</th>' + foot + '<td class="tot">' + all + '</td></tr></tfoot></table>';
  }

  function render(){
    var rows = subset(), n = rows.length,
        /* the grid only counts the competencies and territories it lists, so its total can be
           lower than the headcount. Say so rather than leave two figures that look at odds. */
        inGrid = rows.filter(function(r){ return r[C] >= 0 && r[T] >= 0; }).length;
    exp.innerHTML = rowsHtml(bands(rows, E), n);
    ten.innerHTML = rowsHtml(bands(rows, PW), n);
    off.innerHTML = rowsHtml(offices(rows), n);
    mix.innerHTML = mixHtml(rows);
    grid.innerHTML = gridHtml(rows);
    grid.classList.add('in');
    paintPyramid(rows);
    note.innerHTML = active
      ? 'Showing <b>' + esc(label()) + '</b> · ' + n + (n === 1 ? ' person' : ' people') +
        ' (' + pct(n, N) + '% of the practice). ' +
        (inGrid < n ? '<span class="ts-sub">The grid counts ' + inGrid +
          ' — the rest sit outside the competencies and territories it lists.</span> ' : '') +
        '<button type="button" class="ts-clear" data-ts-clear>Show the whole practice</button>'
      : 'Every figure below covers the whole practice. Pick a grade, a competency, a territory or a cell to see it on its own.';
    [].forEach.call(sec.querySelectorAll('.pr[data-ts]'), function(b){
      var k = b.getAttribute('data-ts');
      b.setAttribute('aria-pressed', active && active.k === 'g' && 'g:' + active.v === k ? 'true' : 'false');
    });
    sec.classList.toggle('ts-filtered', !!active);
  }

  function parse(key){
    var bits = key.split(':');
    if (bits[0] === 'x') return {k: 'x', c: +bits[1], t: +bits[2]};
    if (bits[0] === 'g') return {k: 'g', v: bits[1]};
    return {k: bits[0], v: +bits[1]};
  }
  function same(a, b){
    if (!a || !b || a.k !== b.k) return false;
    return a.k === 'x' ? (a.c === b.c && a.t === b.t) : a.v === b.v;
  }

  sec.addEventListener('click', function(ev){
    var t = ev.target;
    if (t.closest && t.closest('[data-ts-clear]')){ active = null; render(); return; }
    var btn = t.closest ? t.closest('[data-ts]') : null;
    if (!btn) return;
    var want = parse(btn.getAttribute('data-ts'));
    active = same(active, want) ? null : want;      /* clicking the active one clears it */
    render();
  });

  render();
})();
"""

DASH_JS = r"""
(function(){
  var root = document.documentElement,
      panel = document.getElementById('dpanel'),
      src = document.getElementById('dash-src'),
      frame = null, y = 0;
  if (!panel || !src) return;

  function open(){
    y = window.pageYOffset;
    if (!frame){
      frame = document.createElement('iframe');
      frame.title = 'Full dashboard';
      frame.addEventListener('load', function(){
        var wait = panel.querySelector('.dpanel-wait');
        if (wait) wait.parentNode.removeChild(wait);
      });
      panel.appendChild(frame);
      frame.srcdoc = src.textContent.replace(/<\\\/script/g, '</script');
    }
    panel.classList.add('on');
    panel.setAttribute('aria-hidden', 'false');
    root.classList.add('dpanel-on');
  }
  function close(){
    panel.classList.remove('on');
    panel.setAttribute('aria-hidden', 'true');
    root.classList.remove('dpanel-on');
    window.scrollTo(0, y);          /* back to where the dashboard was opened from */
  }

  [].forEach.call(document.querySelectorAll('[data-dash]'), function(b){
    b.addEventListener('click', open);
  });
  /* the Overview button lives inside the dashboard, so it asks us to close */
  window.addEventListener('message', function(ev){ if (ev.data === 'ds-close') close(); });
  document.addEventListener('keydown', function(ev){
    if (ev.key === 'Escape' && panel.classList.contains('on')) close();
  });
})();
"""

HEAD_JS = r"""(function(){var d=document.documentElement;d.className+=' js';try{var m=window.matchMedia&&window.matchMedia('(prefers-reduced-motion: reduce)').matches;
if('IntersectionObserver' in window&&window.requestAnimationFrame&&!m){d.className+=' motion';}}catch(_){}})();"""

BODY_JS = r"""
(function(){
  var root = document.documentElement;
  var motion = /\bmotion\b/.test(root.className);
  function each(sel, fn, scope){ [].forEach.call((scope || document).querySelectorAll(sel), fn); }
  function noMotion(){ motion = false; root.className = root.className.replace(/\bmotion\b/g, ''); }

  // ---- 1 · scroll animation (reveal, counters, words, progress, cube)
  if (motion) try {
    var io = new IntersectionObserver(function(es){
      es.forEach(function(en){ if(en.isIntersecting){ en.target.classList.add('in'); countIn(en.target); io.unobserve(en.target); } });
    }, {rootMargin:'0px 0px -12% 0px', threshold:0.08});
    each('.rv', function(el){ io.observe(el); });

    // Coming back from the dashboard, the browser restores the scroll position. Anything
    // already scrolled past never intersects again, so reveal it at once (no animation,
    // figures left at their final value) — otherwise the page above reads as blank.
    function settleAbove(){
      each('.rv:not(.in)', function(el){
        if (el.getBoundingClientRect().bottom < 4){
          each('[data-count]', function(n){ n._done = true; }, el);
          if (el.matches('[data-count]')) el._done = true;
          el.style.transition = 'none';
          el.classList.add('in');
          void el.offsetWidth;
          el.style.transition = '';
        }
      });
    }
    if (window.pageYOffset > 4) settleAbove();
    window.addEventListener('pageshow', function(){ if (window.pageYOffset > 4) settleAbove(); });
    window.addEventListener('load', function(){ if (window.pageYOffset > 4) settleAbove(); });

    var COUNT_MIN = 50;
    function countIn(scope){
      var els = scope.matches && scope.matches('[data-count]') ? [scope] : [];
      els = els.concat([].slice.call(scope.querySelectorAll('[data-count]')));
      els.forEach(function(el){
        if (el._done) return; el._done = true;
        var to = parseFloat(el.getAttribute('data-count')), t0 = null, dur = 1300;
        var dec = +(el.getAttribute('data-dec') || 0), suf = el.getAttribute('data-suffix') || '';
        /* Only whole figures big enough to be worth watching climb (hours, utilization). Small
           counts and averages (4 competencies, 8.4 years) just arrive with their tile's fade:
           ticking 1-2-3-4 or 0.0-8.4 is what made the row look busy. */
        if (!isFinite(to) || dec || Math.abs(to) < COUNT_MIN) return;
        var tile = el.closest('.rv'), wait = tile ? (parseFloat(getComputedStyle(tile).transitionDelay) || 0) * 1000 : 0;
        var show = function(x){ el.textContent = Math.round(x).toLocaleString('en-US') + suf; };
        function step(t){ if(t0===null) t0=t; var p=Math.min(1,(t-t0)/dur), v=1-Math.pow(2,-10*p);
          show(p < 1 ? to*v : to); if(p<1) requestAnimationFrame(step); }
        show(0); setTimeout(function(){ requestAnimationFrame(step); }, wait);
      });
    }

    var bar = document.querySelector('.progress i'), top = document.querySelector('.top');
    var words = [].slice.call(document.querySelectorAll('.statement .w')), stmt = document.querySelector('.statement');
    var hcube = document.querySelector('.hero-mark');
    var ticking = false;
    function onScroll(){
      ticking = false;
      var y = window.pageYOffset, vh = window.innerHeight, docH = document.documentElement.scrollHeight - vh;
      if (bar) bar.style.transform = 'scaleX(' + (docH > 0 ? y/docH : 0) + ')';
      if (top) top.classList.toggle('scrolled', y > 24);
      if (hcube && y < vh * 1.2) hcube.style.setProperty('--sy', y);
      if (stmt && words.length){
        var r = stmt.getBoundingClientRect();
        var p = (vh*0.85 - r.top) / (r.height*0.75);
        var n = Math.round(Math.max(0, Math.min(1, p)) * words.length);
        for (var i=0;i<words.length;i++) words[i].classList.toggle('on', i < n);
      }
    }
    window.addEventListener('scroll', function(){ if(!ticking){ ticking = true; requestAnimationFrame(onScroll); } }, {passive:true});
    window.addEventListener('resize', onScroll);
    onScroll();

  } catch (err) {
    noMotion();
    if (window.console) console.warn('Scroll animation disabled:', err);
  }

  // ---- 2 · skill orbs swap every 2–5 s
  if (motion) try {
    var SK = window.__SKILLS || [], orbs = [].slice.call(document.querySelectorAll('.orb'));
    if (orbs.length && SK.length > orbs.length){
      var shown = orbs.map(function(o){ return +o.getAttribute('data-i'); }), ptr = orbs.length;
      var pick = function(){
        for (var t = 0; t < SK.length; t++){ var k = (ptr + t) % SK.length;
          if (shown.indexOf(k) < 0){ ptr = (k + 1) % SK.length; return k; } }
        return -1;
      };
      orbs.forEach(function(o, idx){
        var inner = o.querySelector('.orb-in'), faces = o.querySelectorAll('.face .ic'), lab = o.querySelector('.orb-l'), flipped = false;
        function swap(){
          if (!document.hidden && window.pageYOffset < window.innerHeight && o.offsetParent !== null){
            var k = pick();
            if (k >= 0){
              (flipped ? faces[0] : faces[1]).className = 'ic ic-' + SK[k].c;
              flipped = !flipped; inner.classList.toggle('flip', flipped); shown[idx] = k;
              setTimeout(function(){ lab.textContent = SK[k].n + (SK[k].p ? ' · ' + SK[k].p + (SK[k].p === 1 ? ' person' : ' people') : ''); }, 350);
            }
          }
          setTimeout(swap, 2000 + Math.random() * 3000);
        }
        setTimeout(swap, 1800 + idx * 650 + Math.random() * 1500);
      });
    }
  } catch (err) { if (window.console) console.warn('Skill orbs disabled:', err); }

  // ---- 2b · Competency / territory toggle
  try {
    var cs = document.querySelector('.competencies');
    if (cs){
      cs.addEventListener('click', function(ev){
        var b = ev.target.closest ? ev.target.closest('.tg button') : null;
        if (!b) return;
        var v = b.getAttribute('data-v');
        if (cs.getAttribute('data-view') === v || (!cs.getAttribute('data-view') && v === 'comp')) return;
        cs.setAttribute('data-view', v);
        each('.tg button', function(x){ x.setAttribute('aria-pressed', x === b ? 'true' : 'false'); }, cs);
        each('.cv', function(g){
          g.classList.remove('swap');
          if (g.getAttribute('data-v') === v){
            each('.rv', function(c){ c.classList.add('in'); }, g);   // shown at once, animated by .swap
            void g.offsetWidth; g.classList.add('swap');
          }
        }, cs);
      });
    }
  } catch (err) { if (window.console) console.warn('Competency toggle disabled:', err); }

  // ---- 3 · Leadership: click a Director to focus their competencies
  try {
    var tm = document.querySelector('.teams');
    if (tm){
      var orbit = tm.querySelector('.orbit');
      var setView = function(v, scroll){
        each('.tm-view', function(x){ x.classList.toggle('on', x.getAttribute('data-view') === v); }, tm);
        each('[data-k]', function(n){
          var ks = (n.getAttribute('data-k') || '').split(' ');
          n.classList.toggle('dim', v !== 'all' && ks.indexOf(v) < 0);
          if (n.classList.contains('dn')) n.classList.toggle('sel', ks.indexOf(v) >= 0);
        }, orbit);
        if (scroll){
          var r = tm.getBoundingClientRect();
          if (Math.abs(r.top - 70) > 40)
            window.scrollTo({top: window.pageYOffset + r.top - 70, behavior: motion ? 'smooth' : 'auto'});
        }
      };
      tm.addEventListener('click', function(ev){
        var a = ev.target.closest ? ev.target.closest('[data-go]') : null;
        if (!a) return;
        ev.preventDefault();
        setView(a.getAttribute('data-go'), true);
      });
      if (motion){
        var io3 = new IntersectionObserver(function(es){
          if (es[0].isIntersecting){ setTimeout(function(){ orbit.classList.add('settled'); }, 1800); io3.disconnect(); }
        }, {threshold:0.2});
        io3.observe(orbit);
      }
    }
  } catch (err) {
    root.className = root.className.replace(/\bjs\b/g, '');
    if (window.console) console.warn('Leadership diagram disabled:', err);
  }
})();
"""


def build():
    people, util, rates, skills_rows = read_workbook()
    if not people:
        raise SystemExit("No rows found in 'Employee Details'.")
    deliv = Delivery(util, {p["_id"] for p in people})

    children = defaultdict(list)
    for p in people:
        mgr = str(p.get("RL Manager") or "").strip()
        if mgr:
            children[mgr].append(p["Name"])

    by_comp = defaultdict(list)
    for p in people:
        if p["_comp"] != "Other":
            by_comp[p["_comp"]].append(p)
    comps = sorted(by_comp.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    comp_color = {name: ACCENTS[i % len(ACCENTS)] for i, (name, _) in enumerate(comps)}
    others = sum(1 for p in people if p["_comp"] == "Other")

    by_exp = lambda p: (-(p["_exp"] or 0), p["Name"])
    mds = sorted((p for p in people if p["Role"] == "MD"), key=by_exp)
    # the top grade group reads most senior first, "MD & Directors", or "MDs & Directors" when there are several
    for levels in (LEVELS, SHAPE_LEVELS):
        levels[0] = (levels[0][0], ("MDs" if len(mds) > 1 else "MD") + " & Directors", levels[0][2])
    directors = sorted((p for p in people if p["Role"] == "Director"), key=by_exp)
    comp_leads = {name: [d for d in directors if d["_comp"] == name] for name, _ in comps}

    skill_counts = Counter(s for p in people for s in set(p["_skills"]))
    territories = {str(p.get("Territory") or "").strip() for p in people} - {""}
    territories = {t for t in territories if not is_other_territory(t)}
    cities = {str(p.get("Location") or "").strip() for p in people} - {""}

    # photos only for the people the page shows by name (MDs + Directors), each embedded once
    photos, img_css = {}, []
    for p in [q for q in people if q["Role"] == "SM"]:
        uri = thumb_uri(p, 96)
        if uri:
            pid = p["_id"] or slug(p["Name"])
            photos[p["Name"]] = pid
            img_css.append(f'.ph-{pid}{{background-image:url("{uri}")}}')
    for p in mds + directors:
        uri = photo_for(p)
        if uri:
            pid = p["_id"] or slug(p["Name"])
            photos[p["Name"]] = pid
            img_css.append(f'.ph-{pid}{{background-image:url("{uri}")}}')
    skill_imgs = load_skill_images(skill_counts)
    img_css += [f'.ic-{s["cls"]}{{background-image:url("{s["uri"]}")}}' for s in skill_imgs]
    skill_cls = {s["name"].lower(): s["cls"] for s in skill_imgs if not s["placeholder"]}  # chips skip placeholders
    skills_js = json.dumps([{"n": s["name"], "c": s["cls"], "p": s["count"]} for s in skill_imgs]).replace("</", "<\\/")

    logo_uri = data_uri(LOGO) if LOGO.exists() else ""
    generated = datetime.now().strftime("%d %b %Y")

    # the Competency Teams view added to the dashboard's Team Analytics tab
    ct_html, ct_css, ct_data = competency_teams(people, comps, comp_color, comp_leads, deliv)
    ct_view = ({"html": ct_html, "css": CT_CSS + "\n" + ct_css, "js": CT_JS, "data": ct_data,
                "view": "cteam", "label": "Competency Teams"} if ct_html else None)

    exp_avg, pwc_avg = mean(p["_exp"] for p in people), mean(p["_pwc"] for p in people)
    row1 = [kpi(len(people), "Employees"), kpi(len(comps), "Competencies", f"+ {others} in leadership & other teams" if others else ""),
            kpi(len(cities), "Offices"), kpi(round(exp_avg, 1) if exp_avg is not None else "—", "Avg years experience", dec=1),
            kpi(round(pwc_avg, 1) if pwc_avg is not None else "—", "Avg years at PwC", f"as at {generated}", dec=1)]
    if deliv:
        t = deliv.figures()
        fy = f"FY {month_label(deliv.months[0])} – {month_label(deliv.latest)}"
        row2 = [kpi(round(t["util"]) if t["util"] is not None else "—", "Avg utilization", fy, suffix="%"),
                kpi(round(t["ch"]), "Chargeable hours", fy), kpi(round(t["spare"], 1), "Spare capacity (FTE)", month_label(deliv.latest), dec=1),
                kpi(round(t["tr"]), "Training hours", fy), kpi(len(territories), "Client territories")]
    else:
        row2 = [kpi("—", l, "no utilization sheet") for l in ("Avg utilization", "Chargeable hours", "Spare capacity (FTE)", "Training hours")] \
               + [kpi(len(territories), "Client territories")]

    stack = stack_geom(people)
    # the same geometry grade_stack() drew the resting frame from, so STACK_JS carries on from it
    stack_js = json.dumps({k: stack[k] for k in ("ry", "rx", "light", "t", "star")}) if stack else "null"
    body = "".join([
        nav(logo_uri),
        "<main>",
        hero(len(people), len(comps), len(skill_counts), skill_imgs, grade_stack(stack)),
        marquee(skill_counts),
        statement(len(people), len(comps), exp_avg, len(territories)),
        kpis(row1, row2),
        competencies(comps, deliv, comp_leads, people),
        leadership(mds, directors, comps, comp_color, comp_leads, children, photos, deliv, skill_cls,
                   grade_stack(stack, "hub-stack"), people),
        capacity(deliv, people),
        # Skills & depth and the Directory are the dark dashboard's own tabs (Skill Atlas,
        # Capability Risk, Directory), reached from the button in the nav.
        shape(people, children),
        value_section(deliv, rate_lookup(rates), people),
        closing(generated),
        "</main>",
    ])

    page = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
{NO_STORE_META}
<title>{TEAM_NAME} Skills &amp; Bio — Our Practice</title>
<meta name="description" content="A scroll-through showcase of the PwC {TEAM_NAME} practice: headline figures, competencies and leadership.">
<script>{HEAD_JS}</script>
<style>{CSS}</style>
<style>{''.join(img_css)}</style>
</head>
<body>
{body}
{dash_panel(dark_html(ct_view))}
<script>window.__SKILLS={skills_js};</script>
<script>window.__STACK={stack_js};</script>
<script>{DASH_JS}</script>
<script>{STACK_JS}</script>
<script>{PACK_JS}</script>
<script>{SHAPE_JS}</script>
<script>{BODY_JS}</script>
</body>
</html>
"""
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(page, encoding="utf-8")
    print(f"[ok] {len(people)} people, {len(comps)} competencies, {len(mds)} MD, {len(directors)} directors, "
          f"{len(util)} hours rows, {len(rates)} rate rows, {len(skill_imgs)} skill images -> {OUT.name} ({OUT.stat().st_size/1024:.0f} KB)")


if __name__ == "__main__":
    build()
