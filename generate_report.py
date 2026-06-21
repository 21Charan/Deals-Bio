#!/usr/bin/env python3
"""
Employee Bio Dashboard - PwC-branded, single-file HTML.

Reads:
  Employee Details.xlsx   (sheets: Employee Details / Employee Skills /
                           Employee Monthly Utilization)
  Images/<PhotoID>.png    (photos referenced by the PhotoID column)
  template.html           (page template, lives next to this script)

Writes:
  Employee_Dashboard.html (self-contained, images base64-embedded,
                           shareable with clients - works offline)

Reusable: keep the schema the same, replace the .xlsx and/or add photos,
then rerun:
    python generate_report.py
"""

import argparse
import base64
import json
import datetime
import struct
import zlib
import zipfile
from io import BytesIO
from pathlib import Path

import openpyxl

# Optional: Pillow for photo downscaling. Massive HTML-size win at scale.
# pip install Pillow
try:
    from PIL import Image, ImageOps
    HAVE_PIL = True
except ImportError:
    HAVE_PIL = False

# Maximum dimension (px) for embedded photos. 240 covers both the directory
# card and the bio modal at high-DPI quality while keeping file size tiny.
PHOTO_MAX_DIM = 320
PHOTO_JPEG_QUALITY = 72

# ---- Configuration -------------------------------------------------------
SCRIPT_DIR   = Path(__file__).resolve().parent          # 02_Scripts & ETL
ROOT         = SCRIPT_DIR.parent                         # Deals Skills and Bio
SOURCE_DIR   = ROOT / "01_source"                        # raw input files + Images + logo
OUTPUT_DIR   = ROOT / "03_Output files"                  # generated HTML
OUTPUT_DIR.mkdir(exist_ok=True)
INPUT_FILE   = OUTPUT_DIR / "Employee Details.xlsx"      # master file (built by your ETL) lives with the outputs
IMAGES_DIR   = SOURCE_DIR / "Images"
TEMPLATE     = SCRIPT_DIR / "template.html"
OUTPUT_FILE  = OUTPUT_DIR / "Employee_Dashboard.html"

# PwC logo lives in the working folder. Looked up by stem in this order.
LOGO_CANDIDATES = ["PwC Logo", "pwc logo", "PwC_Logo", "logo"]

PHOTO_EXTS   = [".png", ".jpg", ".jpeg", ".webp", ".gif"]
MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".webp": "image/webp", ".gif": "image/gif"}

# Two columns in the source share the name "Experience". Rename the second
# one (years at PwC) for clarity.
EXP_TOTAL = "Experience"
EXP_PWC   = "PwC Experience"

# Role grade codes -> full display names. Applied at load; a no-op if the
# workbook already uses full names. Keeps display/filters/grouping consistent.
ROLE_RENAME = {
    "md": "Managing Director", "managing director": "Managing Director",
    "d": "Director", "director": "Director",
    "sm": "Senior Manager", "sr manager": "Senior Manager", "senior manager": "Senior Manager",
    "m": "Manager", "manager": "Manager",
    "sa": "Senior Associate", "sa1": "Senior Associate", "sa2": "Senior Associate",
    "sa3": "Senior Associate", "senior associate": "Senior Associate",
    "a2": "Associate 2", "associate 2": "Associate 2",
    "a1": "Associate", "a": "Associate", "a3": "Associate", "associate": "Associate",
}


def rename_role(v):
    return ROLE_RENAME.get(str(v or "").strip().lower(), v)


def normalise_headers(raw):
    """De-duplicate header names. Second 'Experience' becomes 'PwC Experience'."""
    seen, out = {}, []
    for h in raw:
        h = "" if h is None else str(h).strip()
        if h in seen:
            seen[h] += 1
            out.append(EXP_PWC if h == "Experience" else f"{h} ({seen[h]})")
        else:
            seen[h] = 1
            out.append(h)
    return out


def read_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = normalise_headers(rows[0])
    records = []
    for raw in rows[1:]:
        if all(c is None or (isinstance(c, str) and c.strip() == "") for c in raw):
            continue
        rec = {}
        for i, header in enumerate(headers):
            v = raw[i] if i < len(raw) else None
            if isinstance(v, datetime.datetime):
                v = v.strftime("%Y-%m-%d")
            elif isinstance(v, datetime.date):
                v = v.isoformat()
            rec[header] = "" if v is None else v
        records.append(rec)
    return records


def _recover_xlsx_bytes(raw):
    """Rebuild a clean .xlsx when its zip central directory / end-of-archive
    record is missing. This can happen when the file is read mid-save or the
    sync layer truncates the tail. We scan the intact local file headers at the
    front of the archive and re-zip every member we can fully decompress."""
    out = BytesIO()
    i, sig = 0, b"PK\x03\x04"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        while True:
            j = raw.find(sig, i)
            if j < 0:
                break
            try:
                (_ver, _flag, method, _mt, _md, _crc, csize, usize,
                 fnlen, eflen) = struct.unpack("<HHHHHIIIHH", raw[j + 4:j + 30])
            except struct.error:
                break
            name = raw[j + 30:j + 30 + fnlen].decode("utf-8", "replace")
            start = j + 30 + fnlen + eflen
            comp = raw[start:start + csize] if csize else b""
            i = j + 4
            try:
                data = zlib.decompress(comp, -15) if method == 8 else comp
            except Exception:
                continue
            if usize and len(data) != usize:   # this member was itself truncated
                continue
            z.writestr(name, data)
    out.seek(0)
    return out


def load_workbook_data(path):
    if not path.exists():
        raise FileNotFoundError(f"Could not find input file: {path}")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        print(f"[warn] Workbook archive looks truncated ({e}); recovering from local headers...")
        wb = openpyxl.load_workbook(_recover_xlsx_bytes(path.read_bytes()), data_only=True)
    keymap = {
        "Employee Details": "details",
        "Employee Skills": "skills",
        "Skill Mapping": "skills",          # real-file sheet name
        "Employee Monthly Utilization": "utilization",
        "Employee Utilization_Jul_Jun": "util_jj",   # FY Jul-Jun view (Pulse toggle)
        "Employee Utilization_Apr_Mar": "util_am",   # FY Apr-Mar view (Pulse toggle)
    }
    data = {"details": [], "skills": [], "utilization": [], "util_jj": [], "util_am": []}
    for ws in wb.worksheets:
        key = keymap.get(ws.title.strip())
        if key is None and not data["details"]:
            key = "details"  # fall back: first sheet is details
        if key:
            data[key] = read_sheet(ws)
    # Fallbacks so the dashboard works whether or not the FY-specific sheets are
    # present. If only the legacy monthly sheet exists, use it for both FY views;
    # if only FY sheets exist, use Jul-Jun as the global utilization source.
    if not data["utilization"]:
        data["utilization"] = data["util_jj"] or data["util_am"]
    if not data["util_jj"]:
        data["util_jj"] = data["utilization"]
    if not data["util_am"]:
        data["util_am"] = data["utilization"]
    return data


def _encode_one(path):
    """Embed the image exactly as it is in the folder, with no re-encoding,
    so quality is preserved. Sizing/compression is handled upstream by the
    user's own image script, so we must not compress it again here."""
    ext = path.suffix.lower()
    mime = MIME.get(ext, "image/png")
    raw = path.read_bytes()
    return (f"data:{mime};base64," + base64.b64encode(raw).decode("ascii"), len(raw))


def _norm_id(v):
    """Normalise an id (number or text) to a plain string for matching.
    Handles int/float (534034.0 -> '534034') and stray whitespace."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def load_photos(images_dir):
    photos = {}
    if not images_dir.exists():
        return photos
    total_bytes = 0
    for f in images_dir.iterdir():
        if f.is_file() and f.suffix.lower() in PHOTO_EXTS:
            uri, size = _encode_one(f)
            total_bytes += size
            stem = f.stem.strip()
            photos[_norm_id(stem)] = uri   # normalised (matches numeric Employee ID)
            photos[stem] = uri             # also keep the raw filename stem
    if photos:
        avg = total_bytes / len(set(photos.values())) if photos else 0
        mode = "embedded as-is (original quality)"
        print(f"[ok] Photos {mode}: {len(set(photos.values()))} files, "
              f"total {total_bytes/1024:.0f} KB, avg {avg/1024:.0f} KB/photo")
    return photos


def photo_for(rec, photos):
    # Photos are matched by the image filename stem. Real data names images by
    # Employee ID, so try that first; fall back to PhotoID / WorkdayID.
    for key in ("Employee ID", "PhotoID", "WorkdayID"):
        nid = _norm_id(rec.get(key, ""))
        if not nid:
            continue
        uri = photos.get(nid, "")
        if uri:
            return uri
    return ""


def load_logo():
    """Find the PwC logo in the working folder and return a base64 data URI."""
    for stem in LOGO_CANDIDATES:
        for ext in PHOTO_EXTS:
            p = SOURCE_DIR / f"{stem}{ext}"
            if p.exists():
                b64 = base64.b64encode(p.read_bytes()).decode("ascii")
                return f"data:{MIME[ext.lower()]};base64,{b64}"
    return ""


def initials(name):
    parts = [p for p in str(name or "").split() if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def main():
    parser = argparse.ArgumentParser(description="Build the employee dashboard HTML.")
    parser.add_argument("--pbi-url", default="",
                        help="Optional Power BI report URL linked from bios in the restricted build")
    args = parser.parse_args()

    data = load_workbook_data(INPUT_FILE)
    for _r in data["details"]:                     # normalise Role grade codes
        _r["Role"] = rename_role(_r.get("Role"))
    photos = load_photos(IMAGES_DIR)
    logo_uri = load_logo()

    enriched = []
    for r in data["details"]:
        c = dict(r)
        c["_photo"] = photo_for(r, photos)
        c["_initials"] = initials(r.get("Name"))
        enriched.append(c)

    # Photo-match diagnostic (helps spot Employee ID vs image-name mismatches)
    if photos:
        matched = sum(1 for e in enriched if e.get("_photo"))
        print(f"[ok] Photos matched to employees: {matched} / {len(enriched)} (by Employee ID)")
        if matched < len(enriched):
            unmatched = [_norm_id(e.get("Employee ID")) for e in enriched if not e.get("_photo")]
            avail = sorted({_norm_id(k) for k in photos})
            print(f"[warn] Unmatched Employee IDs (sample): {unmatched[:8]}")
            print(f"       Image file names available (sample): {avail[:8]}")

    # Stats
    exps = [e.get(EXP_TOTAL) for e in enriched if isinstance(e.get(EXP_TOTAL), (int, float))]
    avg_exp = round(sum(exps) / len(exps), 1) if exps else 0
    ous   = sorted({str(e.get("OU", "")).strip()       for e in enriched if str(e.get("OU", "")).strip()})
    locs  = sorted({str(e.get("Location", "")).strip() for e in enriched if str(e.get("Location", "")).strip()})
    roles = sorted({str(e.get("Role", "")).strip()     for e in enriched if str(e.get("Role", "")).strip()})

    def opts(values):
        return "\n".join(f'<option value="{v}">{v}</option>' for v in values)

    # Day with no leading zero ("31 May 2026" not "31 May 2026, 17:42")
    today = datetime.date.today()
    generated_str = f"{today.day} {today.strftime('%B %Y')}"

    template = TEMPLATE.read_text(encoding="utf-8")

    def render(share):
        """share=True -> restricted build: utilization stripped + full Address redacted."""
        emp = enriched
        util = data["utilization"]
        util_jj = data["util_jj"]
        util_am = data["util_am"]
        if share:
            util = util_jj = util_am = []                  # no utilization anywhere
            emp = [{k: v for k, v in e.items() if k != "Address"} for e in enriched]  # redact full address
        return (template
            .replace("__EMPLOYEES_JSON__", json.dumps(emp, default=str))
            .replace("__SKILLS_JSON__",    json.dumps(data["skills"], default=str))
            .replace("__UTIL_JSON__",      json.dumps(util, default=str))
            .replace("__UTIL_JJ_JSON__",   json.dumps(util_jj, default=str))
            .replace("__UTIL_AM_JSON__",   json.dumps(util_am, default=str))
            .replace("__TOTAL__",          str(len(enriched)))
            .replace("__AVG_EXP__",        str(avg_exp))
            .replace("__NUM_OUS__",        str(len(ous)))
            .replace("__NUM_LOCS__",       str(len(locs)))
            .replace("__OU_OPTIONS__",     opts(ous))
            .replace("__ROLE_OPTIONS__",   opts(roles))
            .replace("__LOC_OPTIONS__",    opts(locs))
            .replace("__GENERATED__",      generated_str)
            .replace("__LOGO_URI__",       logo_uri)
            .replace("__EXP_TOTAL__",      EXP_TOTAL)
            .replace("__EXP_PWC__",        EXP_PWC)
            .replace("__SHARE_MODE__",     "true" if share else "false")
            .replace("__PBI_URL__",        args.pbi_url))

    builds = [
        (False, OUTPUT_DIR / "Employee_Dashboard.html",        "MD / full (all data)"),
        (True,  OUTPUT_DIR / "Employee_Dashboard_Shared.html", "Restricted (no utilization, Address redacted)"),
    ]

    print(f"[ok] Employees: {len(enriched)} | Skill rows: {len(data['skills'])} | Util rows: {len(data['utilization'])}")
    ids_with_skills = {str(r.get("WorkdayID")) for r in data["skills"] if r.get("WorkdayID") not in (None, "")}
    no_skill_emps = [e for e in enriched if str(e.get("WorkdayID")) not in ids_with_skills]
    if no_skill_emps:
        names = ", ".join(str(e.get("Name", "?")) for e in no_skill_emps[:5])
        more = f" + {len(no_skill_emps) - 5} more" if len(no_skill_emps) > 5 else ""
        print(f"[warn] {len(no_skill_emps)} employee(s) have NO rows in the skills sheet: {names}{more}")
    print(f"[ok] Logo found: {'yes' if logo_uri else 'no'}")
    for share, out_file, label in builds:
        out_file.write_text(render(share), encoding="utf-8")
        size_kb = out_file.stat().st_size / 1024
        print(f"[ok] {label:46s} -> {out_file.name}  ({size_kb:,.0f} KB)")


if __name__ == "__main__":
    main()
