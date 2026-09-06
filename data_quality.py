"""
Data quality report for the Employee Bio Dashboard.

Reads the same workbook the generator reads, plus the Images folder, and
writes an Excel file of everything that is missing or inconsistent.  It
changes nothing.

    Summary          one row per person, with their RL manager's email and a
                     Notes column naming every column they are missing, so one
                     row is one message to send
    Completeness     every column in the extract, how complete it is
    Missing - <col>  one tab per column that has gaps, with contact details
    Utilization Gaps people with no utilization, holes mid-series, duplicates
    Value Checks     values that are present but disagree with each other
    Not On Roster    IDs in the utilization sheets with no roster record

The column list is not hard-coded: whatever columns the Employee Details
sheet carries are the columns that get checked, so adding one to the extract
next month needs no change here.  Only the severity of a few well-known
columns, and the checks that are not about a column at all (photos, skills,
utilization), are named in the code.

Run it on its own:

    python data_quality.py

or let generate_report.py call build() at the end of a build.
"""

import datetime
import re
import sys
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:                                          # pragma: no cover
    sys.exit("openpyxl is not installed.  pip install openpyxl")

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
WORKBOOK_CANDIDATES = [
    ROOT / "03_Output files" / "Employee Details.xlsx",
    HERE / "Employee Details.xlsx",
    ROOT / "01_Source" / "Employee Details.xlsx",
]
DEFAULT_IMAGES = ROOT / "01_Source" / "Images"
DEFAULT_OUT = ROOT / "03_Output files" / "Data_Quality_Report.xlsx"

PHOTO_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
MAX_ISSUE_TABS = 20          # beyond this the long tail goes to "Other Gaps"

ORANGE = "D04A02"
WHITE = "FFFFFF"

# Severity is about what breaks if the value is absent:
#   High   - the person is mis-filed, mis-counted or unreachable
#   Medium - a tab or a filter loses them
#   Low    - cosmetic
# Anything not named here is Medium: a column someone added to the extract is
# presumed to be there for a reason.
SEVERITY = {
    "name": "High", "workdayid": "High", "role": "High", "designation": "High",
    "rlmanager": "High", "rlmanageremail": "High", "emailid": "High", "email": "High",
    "location": "High", "competency": "High",
    "employeedescription": "Low", "qualification": "Low", "gender": "Low",
    "workmode": "Low", "coach": "Low", "talentconsultant": "Low",
}
SEV_ORDER = {"High": 0, "Medium": 1, "Low": 2}

# One spelling per state, so the vocabulary cannot drift extract by extract.
KNOWN_STATUS = {"active", "not active", "left", "on leave"}

ROLE_GROUPS = [
    ("Managing Directors", r"managing\s*director|^md\b"),
    ("Directors",          r"^director$|^d$"),
    ("Senior Managers",    r"senior\s*manager|sr\.?\s*manager|^sm$"),
    ("Managers",           r"^manager$|^m$"),
    ("Senior Associates",  r"senior\s*associate|^sa\d*$"),
    ("Associates",         r"^associate(\s*\d+)?$|^a\d*$"),
]
GRADE_NAME = {
    "Managing Directors": "Managing Director", "Directors": "Director",
    "Senior Managers": "Senior Manager", "Managers": "Manager",
    "Senior Associates": "Senior Associate", "Associates": "Associate",
}


# ---------------------------------------------------------------------------
# Helpers — the same normalisation the dashboard applies, so a value this
# report calls present is one the dashboard can actually use.
# ---------------------------------------------------------------------------
def flat(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def find_col(headers, *names):
    idx = {flat(h): h for h in headers if h is not None}
    for n in names:
        if flat(n) in idx:
            return idx[flat(n)]
    return None


def blank(v):
    return v is None or str(v).strip() == "" or str(v).strip().lower() in ("none", "nan")


def norm_id(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def is_email(v):
    return bool(EMAIL_RE.match(str(v).strip()))


def is_number(v):
    try:
        float(str(v).strip())
        return True
    except (TypeError, ValueError):
        return False


def is_date(v):
    if hasattr(v, "year"):
        return True
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y", "%Y/%m/%d"):
        try:
            datetime.datetime.strptime(s[:10], fmt)
            return True
        except ValueError:
            pass
    return False


def validator_for(column):
    """Pick a format check from the column's own name, so a new 'Backup Email'
    column is validated the same way 'RL Manager Email' is."""
    f = flat(column)
    if "email" in f:
        return is_email, "email address"
    if "date" in f:
        return is_date, "date"
    if f in ("experience", "totalexperience", "pwcexperience") or "hours" in f or "rate" in f:
        return is_number, "number"
    return None, ""


def month_key(raw):
    if raw is None or raw == "":
        return ""
    if hasattr(raw, "year"):
        return "%04d-%02d" % (raw.year, raw.month)
    s = str(raw).strip()
    m = re.match(r"^(\d{4})[-/](\d{1,2})", s)
    if m:
        return "%04d-%02d" % (int(m.group(1)), int(m.group(2)))
    m = re.match(r"^([A-Za-z]{3,})[-\s/](\d{4})$", s)
    if m:
        mon = ["jan", "feb", "mar", "apr", "may", "jun",
               "jul", "aug", "sep", "oct", "nov", "dec"].index(m.group(1)[:3].lower()) + 1
        return "%04d-%02d" % (int(m.group(2)), mon)
    return ""


def month_span(first, last):
    y, m = int(first[:4]), int(first[5:7])
    ey, em = int(last[:4]), int(last[5:7])
    out = []
    while (y, m) <= (ey, em):
        out.append("%04d-%02d" % (y, m))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


def rows_of(ws):
    rows = list(ws.values)
    if not rows:
        return [], []
    headers = list(rows[0])
    return headers, [dict(zip(headers, r)) for r in rows[1:]
                     if any(c is not None and str(c).strip() != "" for c in r)]


def group_for_role(role):
    r = str(role or "").strip()
    for title, pat in ROLE_GROUPS:
        if re.search(pat, r, re.I):
            return title
    return "Other"


# ---------------------------------------------------------------------------
# The checks
# ---------------------------------------------------------------------------
def check_people(roster, headers, photos, skill_ids):
    """One record per person per gap.

    Every column in the sheet is checked for blanks; columns whose name says
    what they hold (email, date, hours) are checked for shape as well.  Photo
    and Skills are added on top because they live outside the sheet.
    """
    out = []
    c_name = find_col(headers, "Name") or "Name"
    c_id = find_col(headers, "WorkdayID", "Workday ID")
    c_emp = find_col(headers, "Employee ID")
    c_role = find_col(headers, "Role", "Designation")
    c_mgr = find_col(headers, "RL Manager")
    c_mgr_mail = find_col(headers, "RL Manager Email")
    c_mail = find_col(headers, "emailid", "Email", "Email ID")
    c_loc = find_col(headers, "Location")
    columns = [h for h in headers if h is not None and str(h).strip() != ""]

    for p in roster:
        base = dict(
            Name=p.get(c_name, ""),
            WorkdayID=norm_id(p.get(c_id, "")),
            EmployeeID=norm_id(p.get(c_emp, "")),
            Role=p.get(c_role, ""),
            Location=p.get(c_loc, ""),
            Email=p.get(c_mail, "") if c_mail else "",
            Manager=p.get(c_mgr, "") if c_mgr else "",
            ManagerEmail=p.get(c_mgr_mail, "") if c_mgr_mail else "",
        )
        for col in columns:
            sev = SEVERITY.get(flat(col), "Medium")
            v = p.get(col)
            note = ""
            if col in (c_mgr, c_mgr_mail) and group_for_role(base["Role"]) == "Managing Directors":
                note = "May be the top of the reporting tree"
            if blank(v):
                out.append(dict(base, Column=str(col), Issue="Blank",
                                Detail="", Severity=sev, Note=note))
                continue
            valid, what = validator_for(col)
            if valid and not valid(v):
                out.append(dict(base, Column=str(col), Issue="Not a valid " + what,
                                Detail=str(v)[:90], Severity=sev, Note=note))

        # Photo — matched on the image filename stem, the same three keys the
        # generator tries, in the same order.
        if photos is not None:
            keys = [k for k in (base["EmployeeID"], norm_id(p.get("PhotoID", "")),
                                base["WorkdayID"]) if k]
            if not any(k in photos for k in keys):
                out.append(dict(base, Column="Photo", Issue="Blank",
                                Detail="no image file matches this person",
                                Severity="High",
                                Note=("Add %s to the Images folder"
                                      % " or ".join(k + ".png" for k in dict.fromkeys(keys)))
                                if keys else "No ID on this record to match an image on"))
        # Skills — the roster column or the skills sheet will do; neither is a gap.
        if skill_ids is not None and base["WorkdayID"] not in skill_ids:
            col_sk = find_col(headers, "Skills")
            if not col_sk or blank(p.get(col_sk)):
                out.append(dict(base, Column="Skills", Issue="Blank",
                                Detail="no skills on the roster or in the skills sheet",
                                Severity="Medium",
                                Note="Drops out of Skill Atlas and staffing search"))
    return out


def util_sheets(wb):
    """The utilization sheets the dashboard actually reads.

    'Utilization Full_X' supersedes 'Employee Utilization_X' in the generator,
    so checking both would report every person as missing from a sheet the
    dashboard never looks at.
    """
    named = {}
    for ws in wb.worksheets:
        t = ws.title.strip()
        if "utilization" not in t.lower() and "utilisation" not in t.lower():
            continue
        window = t.split("_", 1)[1].lower() if "_" in t else t.lower()
        full = t.lower().startswith(("utilization full", "utilisation full"))
        prev = named.get(window)
        if prev is None or (full and not prev[0]):
            named[window] = (full, ws)
    return [ws for _, (full, ws) in sorted(named.items())]


def check_utilization(wb, roster_ids, id_of_name):
    rows, unmatched = [], []
    for ws in util_sheets(wb):
        t = ws.title.strip()
        headers, data = rows_of(ws)
        c_id = find_col(headers, "Workday ID", "WorkdayID")
        c_mo = find_col(headers, "EoM", "Month Year", "Month")
        if not c_id or not c_mo:
            continue
        c_std = find_col(headers, "Standard Hours")
        c_nm = find_col(headers, "EMP Name", "Name")

        by_person, seen, no_std = {}, {}, {}
        for r in data:
            wid, mo = norm_id(r.get(c_id)), month_key(r.get(c_mo))
            if not wid or not mo:
                continue
            by_person.setdefault(wid, set()).add(mo)
            seen[(wid, mo)] = seen.get((wid, mo), 0) + 1
            if c_std and (r.get(c_std) in (None, "", 0) or not is_number(r.get(c_std))):
                no_std[wid] = no_std.get(wid, 0) + 1
            if wid not in roster_ids:
                unmatched.append(dict(Sheet=t, WorkdayID=wid,
                                      Name=r.get(c_nm, "") if c_nm else "",
                                      Detail="in the utilization sheet but not on the roster",
                                      Note="Reads as a leaver — correct if they have left"))
        for wid in sorted(roster_ids):
            months = sorted(by_person.get(wid, ()))
            name = id_of_name.get(wid, "")
            if not months:
                rows.append(dict(Sheet=t, WorkdayID=wid, Name=name,
                                 Issue="No utilization at all", Months="",
                                 Detail="no rows in this sheet", Severity="High"))
                continue
            holes = [m for m in month_span(months[0], months[-1]) if m not in set(months)]
            if holes:
                rows.append(dict(Sheet=t, WorkdayID=wid, Name=name,
                                 Issue="Months missing mid-series",
                                 Months="%s .. %s" % (months[0], months[-1]),
                                 Detail=", ".join(holes), Severity="Medium"))
            if no_std.get(wid):
                rows.append(dict(Sheet=t, WorkdayID=wid, Name=name,
                                 Issue="Rows with no Standard Hours",
                                 Months="%s .. %s" % (months[0], months[-1]),
                                 Detail="%d row(s) — these are left out of the average"
                                        % no_std[wid], Severity="Medium"))
        dupes = sorted((k, n) for k, n in seen.items() if n > 1 and k[0] in roster_ids)
        for (wid, mo), n in dupes:
            rows.append(dict(Sheet=t, WorkdayID=wid, Name=id_of_name.get(wid, ""),
                             Issue="Duplicate person-month", Months=mo,
                             Detail="%d rows for this month — the hours are summed" % n,
                             Severity="High"))
    seen_u, out_u = set(), []
    for u in unmatched:
        k = (u["Sheet"], u["WorkdayID"])
        if k not in seen_u:
            seen_u.add(k)
            out_u.append(u)
    return rows, out_u


def check_values(wb, roster, headers, photos):
    """Values that are present but disagree with each other."""
    out = []
    c_role = find_col(headers, "Role", "Designation")
    c_terr = find_col(headers, "Territory")
    c_tf = find_col(headers, "Territory Filter")
    c_comp = find_col(headers, "Competency")
    c_cg = find_col(headers, "Competency Group")
    c_cf = find_col(headers, "Competency Filter")
    c_status = find_col(headers, "Status")

    def mapping(src, dst):
        if not src or not dst:
            return
        m = {}
        for p in roster:
            a, b = str(p.get(src) or "").strip(), str(p.get(dst) or "").strip()
            if a and b:
                m.setdefault(a, set()).add(b)
        for a, bs in sorted(m.items()):
            if len(bs) > 1:
                out.append(dict(Check="Inconsistent mapping", Value=a,
                                Detail="%s '%s' maps to %d different %s values: %s"
                                       % (src, a, len(bs), dst, ", ".join(sorted(bs))),
                                Severity="High"))

    mapping(c_terr, c_tf)
    mapping(c_comp, c_cg)
    mapping(c_comp, c_cf)

    # Values that differ only by case, spacing or punctuation — checked on every
    # column that behaves like a category (few distinct values, many rows).
    for col in [h for h in headers if h is not None]:
        vals = [str(p.get(col) or "").strip() for p in roster]
        vals = [v for v in vals if v]
        distinct = set(vals)
        if not vals or len(distinct) > max(3, len(vals) * 0.6):
            continue                      # free text, not a category
        groups = {}
        for v in distinct:
            groups.setdefault(flat(v), set()).add(v)
        for _, variants in sorted(groups.items()):
            if len(variants) > 1:
                out.append(dict(Check="Near-duplicate value", Value=str(col),
                                Detail="written %d ways: %s"
                                       % (len(variants), ", ".join(sorted(variants))),
                                Severity="Medium"))

    if c_status:
        counts = {}
        for p in roster:
            v = str(p.get(c_status) or "").strip()
            counts[v or "(blank)"] = counts.get(v or "(blank)", 0) + 1
        for v, n in sorted(counts.items()):
            known = v.strip().lower() in KNOWN_STATUS
            out.append(dict(Check="Status value", Value=v,
                            Detail="%d %s%s" % (n, "person" if n == 1 else "people",
                                                "" if known else " — not a recognised status"),
                            Severity="Low" if known else "Medium"))

    rates_ws = next((w for w in wb.worksheets if flat(w.title) == flat("Hourly Rates")), None)
    if rates_ws:
        rh, rd = rows_of(rates_ws)
        c_rr, c_rt = find_col(rh, "Role", "Grade"), find_col(rh, "Territory")
        c_rv = find_col(rh, "Hourly Rate (USD)", "Hourly Rate", "Rate")
        table = {(flat(r.get(c_rr)), flat(r.get(c_rt))) for r in rd
                 if c_rv is not None and not blank(r.get(c_rv))}
        pairs = {}
        for p in roster:
            role = str(p.get(c_role) or "").strip()
            terr = str(p.get(c_terr) or "").strip() or str(p.get(c_tf) or "").strip()
            if role:
                pairs[(role, terr)] = pairs.get((role, terr), 0) + 1
        for (role, terr), n in sorted(pairs.items()):
            grade = GRADE_NAME.get(group_for_role(role), role)
            if not any((flat(name), flat(t)) in table
                       for name in (role, grade) for t in (terr, "standard")):
                out.append(dict(Check="No rate card entry", Value="%s / %s" % (role, terr),
                                Detail="%d %s priced at nothing on the Rate Analysis tab"
                                       % (n, "person" if n == 1 else "people"),
                                Severity="Medium"))

    if photos:
        keys = set()
        for p in roster:
            for c in ("Employee ID", "PhotoID", "WorkdayID", "Workday ID"):
                col = find_col(headers, c) or c
                keys.add(norm_id(p.get(col, "")))
        for stem in sorted(photos):
            if stem not in keys:
                out.append(dict(Check="Image matches nobody", Value=stem,
                                Detail="%s is in the Images folder but nobody on the roster "
                                       "carries that ID" % photos[stem], Severity="Low"))
    return out


# ---------------------------------------------------------------------------
# Writing the workbook
# ---------------------------------------------------------------------------
HDR_FONT = Font(name="Arial", size=10, bold=True, color=WHITE)
HDR_FILL = PatternFill("solid", fgColor=ORANGE)
BODY = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=14, bold=True)
NOTE = Font(name="Arial", size=9, italic=True, color="595959")

CONTACT_HDR = ["Name", "WorkdayID", "Employee ID", "Role", "Location",
               "Email", "RL Manager", "RL Manager Email"]
CONTACT_W = [22, 12, 12, 16, 14, 26, 20, 28]


def contact_row(p):
    return [p["Name"], p["WorkdayID"], p["EmployeeID"], p["Role"], p["Location"],
            p["Email"], p["Manager"], p["ManagerEmail"]]


def safe_title(name, taken):
    t = re.sub(r"[\[\]:*?/\\]", "-", str(name)).strip()[:31] or "Sheet"
    base, n = t, 2
    while t.lower() in taken:
        suffix = " (%d)" % n
        t = base[:31 - len(suffix)] + suffix
        n += 1
    taken.add(t.lower())
    return t


def sheet(wb, title, headers, rows, widths, taken, first_row=1):
    ws = wb.create_sheet(safe_title(title, taken))
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=first_row, column=i, value=h)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(vertical="center")
    # Written by coordinate rather than append(), so a sheet whose header does
    # not start at row 1 still lays its rows out directly beneath it.
    for ri, r in enumerate(rows, start=first_row + 1):
        for ci, v in enumerate(r, start=1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = BODY
            c.alignment = Alignment(vertical="top")
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=first_row + 1, column=1).coordinate
    if rows:
        ws.auto_filter.ref = "A%d:%s%d" % (first_row, get_column_letter(len(headers)),
                                           first_row + len(rows))
    ws.sheet_view.showGridLines = False
    return ws


def build(workbook_path=None, images_dir=None, out_path=None, quiet=False):
    workbook_path = Path(workbook_path) if workbook_path else next(
        (p for p in WORKBOOK_CANDIDATES if p.exists()), None)
    if not workbook_path or not workbook_path.exists():
        raise SystemExit("Workbook not found. Looked in: %s"
                         % ", ".join(str(p) for p in WORKBOOK_CANDIDATES))
    images_dir = Path(images_dir) if images_dir else DEFAULT_IMAGES
    out_path = Path(out_path) if out_path else DEFAULT_OUT

    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    details = next((w for w in wb.worksheets if flat(w.title).startswith("employeedetail")),
                   wb.worksheets[0])
    headers, roster = rows_of(details)

    photos = {}
    if images_dir.exists():
        for f in sorted(images_dir.iterdir()):
            if f.suffix.lower() in PHOTO_EXTS:
                photos[norm_id(f.stem)] = f.name

    skills_ws = next((w for w in wb.worksheets
                      if flat(w.title) in ("employeeskills", "skillmapping",
                                           "employeeskillshierarchy")), None)
    skill_ids = None
    if skills_ws:
        sh, sd = rows_of(skills_ws)
        c_sid = find_col(sh, "WorkdayID", "Workday ID")
        skill_ids = {norm_id(r.get(c_sid)) for r in sd if c_sid and not blank(r.get(c_sid))}

    c_id = find_col(headers, "WorkdayID", "Workday ID")
    c_name = find_col(headers, "Name")
    roster_ids = {norm_id(p.get(c_id)) for p in roster if not blank(p.get(c_id))}
    id_of_name = {norm_id(p.get(c_id)): p.get(c_name, "") for p in roster}

    gaps = check_people(roster, headers, photos, skill_ids)
    util, unmatched = check_utilization(wb, roster_ids, id_of_name)
    values = check_values(wb, roster, headers, photos)

    out = openpyxl.Workbook()
    out.remove(out.active)
    taken = set()

    # --- Summary: one row per person, everything they are missing -----------
    by_person = {}
    for g in gaps:
        by_person.setdefault(g["WorkdayID"] or g["Name"], []).append(g)
    people_rows = []
    for _, items in by_person.items():
        p = items[0]
        worst = min(SEV_ORDER[i["Severity"]] for i in items)
        cols = sorted({i["Column"] for i in items},
                      key=lambda c: (SEV_ORDER[next(i["Severity"] for i in items
                                                    if i["Column"] == c)], c))
        people_rows.append((worst, -len(cols), str(p["Name"]),
                            contact_row(p) + [len(cols),
                                              ["High", "Medium", "Low"][worst],
                                              ", ".join(cols)]))
    people_rows.sort(key=lambda t: (t[0], t[1], t[2]))

    ws = sheet(out, "Summary",
               CONTACT_HDR + ["Gaps", "Worst", "Missing columns (what to ask them for)"],
               [r[3] for r in people_rows], CONTACT_W + [7, 9, 80], taken, first_row=5)
    ws["A1"] = "Data quality — Employee Bio Dashboard"
    ws["A1"].font = TITLE
    ws["A2"] = ("Workbook: %s · %d people on the roster · %d with something missing · "
                "generated %s" % (workbook_path.name, len(roster), len(people_rows),
                                  datetime.datetime.now().strftime("%d %b %Y %H:%M")))
    ws["A2"].font = NOTE
    ws["A3"] = ("One row per person. Filter RL Manager to give a manager their whole team's "
                "list in one message. Severity: High = mis-filed, mis-counted or unreachable · "
                "Medium = a tab or filter loses them · Low = cosmetic.")
    ws["A3"].font = NOTE

    # --- Completeness: every column, how complete ---------------------------
    counts = {}
    for g in gaps:
        counts[g["Column"]] = counts.get(g["Column"], 0) + 1
    all_columns = [str(h) for h in headers if h is not None and str(h).strip() != ""]
    all_columns += ["Photo"] + (["Skills"] if skills_ws else [])
    seen_c, ordered = set(), []
    for c in all_columns:                                    # keep sheet order, no repeats
        if c.lower() not in seen_c:
            seen_c.add(c.lower())
            ordered.append(c)
    comp = sheet(out, "Completeness",
                 ["Column", "Severity", "People missing", "Complete"],
                 [[c, SEVERITY.get(flat(c), "High" if c == "Photo" else "Medium"), None, None]
                  for c in ordered], [34, 10, 16, 12], taken, first_row=3)
    comp["A1"] = "Every column in the extract. Counts are live formulas over the Summary sheet."
    comp["A1"].font = NOTE
    comp["F1"], comp["G1"] = "People on roster", len(roster)
    comp["F1"].font, comp["G1"].font = NOTE, BOLD
    for i, c in enumerate(ordered, start=4):
        # Counted off the Summary sheet's "Missing columns" list, so the two can
        # never disagree; the wildcards stop 'Competency' matching 'Competency Group'.
        comp.cell(row=i, column=3,
                  value=('=COUNTIF(Summary!$K$6:$K$100000,"*"&$A%d&",*")'
                         '+COUNTIF(Summary!$K$6:$K$100000,"*, "&$A%d)'
                         '+COUNTIF(Summary!$K$6:$K$100000,$A%d)' % (i, i, i))).font = BODY
        pc = comp.cell(row=i, column=4,
                       value='=IF($G$1=0,"",($G$1-C%d)/$G$1)' % i)
        pc.font, pc.number_format = BODY, "0%"

    # --- One tab per column that has gaps -----------------------------------
    ranked = sorted(counts.items(),
                    key=lambda kv: (SEV_ORDER[next(g["Severity"] for g in gaps
                                                   if g["Column"] == kv[0])], -kv[1], kv[0]))
    overflow = []
    for col, _n in ranked[:MAX_ISSUE_TABS]:
        items = [g for g in gaps if g["Column"] == col]
        sheet(out, "Missing - " + col, CONTACT_HDR + ["Issue", "Detail", "What to do"],
              [contact_row(g) + [g["Issue"], g["Detail"], g["Note"]] for g in
               sorted(items, key=lambda g: str(g["Name"]))],
              CONTACT_W + [18, 34, 46], taken)
    for col, _n in ranked[MAX_ISSUE_TABS:]:
        overflow += [g for g in gaps if g["Column"] == col]
    if overflow:
        sheet(out, "Other Gaps", CONTACT_HDR + ["Column", "Issue", "Detail"],
              [contact_row(g) + [g["Column"], g["Issue"], g["Detail"]] for g in overflow],
              CONTACT_W + [24, 18, 34], taken)

    # --- The checks that are not about a person's record --------------------
    sheet(out, "Utilization Gaps",
          ["Sheet", "WorkdayID", "Name", "Issue", "Months", "Detail", "Severity"],
          [[u["Sheet"], u["WorkdayID"], u["Name"], u["Issue"], u["Months"],
            u["Detail"], u["Severity"]] for u in util], [26, 12, 22, 28, 20, 46, 10], taken)
    sheet(out, "Value Checks", ["Check", "Column or value", "Detail", "Severity"],
          [[v["Check"], v["Value"], v["Detail"], v["Severity"]] for v in values],
          [24, 26, 78, 10], taken)
    sheet(out, "Not On Roster", ["Sheet", "WorkdayID", "Name", "Detail", "Note"],
          [[u["Sheet"], u["WorkdayID"], u["Name"], u["Detail"], u["Note"]] for u in unmatched],
          [26, 12, 22, 46, 40], taken)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        out.save(out_path)
    except PermissionError:
        # Excel locks the file it has open. Losing this month's report to a
        # window someone left open is worse than a second filename.
        alt = out_path.with_name(out_path.stem + " (new)" + out_path.suffix)
        out.save(alt)
        if not quiet:
            print("[note] %s is open in Excel — wrote %s instead."
                  % (out_path.name, alt.name))
        out_path = alt
    if not quiet:
        print("[ok] Wrote %s — %d people to chase, %d gaps across %d columns, "
              "%d utilization gaps, %d value checks"
              % (out_path.name, len(people_rows), len(gaps), len(counts),
                 len(util), len(values)))
    return out_path


if __name__ == "__main__":
    build()
