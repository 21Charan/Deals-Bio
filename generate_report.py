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

import base64
import json
import datetime
from io import BytesIO
from pathlib import Path

import openpyxl

# Optional: Pillow for photo downscaling. Massive HTML-size win at scale.
# pip install Pillow
try:
    from PIL import Image
    HAVE_PIL = True
except ImportError:
    HAVE_PIL = False

# Maximum dimension (px) for embedded photos. 240 covers both the directory
# card and the bio modal at high-DPI quality while keeping file size tiny.
PHOTO_MAX_DIM = 240
PHOTO_JPEG_QUALITY = 72

# ---- Configuration -------------------------------------------------------
SCRIPT_DIR   = Path(__file__).resolve().parent
INPUT_FILE   = SCRIPT_DIR / "Employee Details.xlsx"
IMAGES_DIR   = SCRIPT_DIR / "Images"
TEMPLATE     = SCRIPT_DIR / "template.html"
OUTPUT_FILE  = SCRIPT_DIR / "Employee_Dashboard.html"

# PwC logo lives in the working folder. Looked up by stem in this order.
LOGO_CANDIDATES = ["PwC Logo", "pwc logo", "PwC_Logo", "logo"]

PHOTO_EXTS   = [".png", ".jpg", ".jpeg", ".webp", ".gif"]
MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".webp": "image/webp", ".gif": "image/gif"}

# Two columns in the source share the name "Experience". Rename the second
# one (years at PwC) for clarity.
EXP_TOTAL = "Experience"
EXP_PWC   = "PwC Experience"


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


def load_workbook_data(path):
    if not path.exists():
        raise FileNotFoundError(f"Could not find input file: {path}")
    wb = openpyxl.load_workbook(path, data_only=True)
    keymap = {
        "Employee Details": "details",
        "Employee Skills": "skills",
        "Employee Monthly Utilization": "utilization",
    }
    data = {"details": [], "skills": [], "utilization": []}
    for ws in wb.worksheets:
        key = keymap.get(ws.title.strip())
        if key is None and not data["details"]:
            key = "details"  # fall back: first sheet is details
        if key:
            data[key] = read_sheet(ws)
    return data


def _encode_one(path):
    """Return (data_uri, bytes_size) for a single image file.

    With Pillow: re-encode to a JPEG thumbnail (max 240px) for huge savings.
    Without Pillow: embed the original bytes (with a warning printed by caller).
    """
    if HAVE_PIL:
        try:
            with Image.open(path) as img:
                # Flatten alpha onto white so JPEG (no alpha) doesn't go black.
                if img.mode in ("RGBA", "LA", "P"):
                    if img.mode == "P":
                        img = img.convert("RGBA")
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
                    img = bg
                elif img.mode != "RGB":
                    img = img.convert("RGB")
                # Downscale (in-place, keeps aspect ratio). LANCZOS = high quality.
                img.thumbnail((PHOTO_MAX_DIM, PHOTO_MAX_DIM), Image.LANCZOS)
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=PHOTO_JPEG_QUALITY, optimize=True)
                data = buf.getvalue()
                return ("data:image/jpeg;base64," + base64.b64encode(data).decode("ascii"), len(data))
        except Exception as e:
            print(f"  [warn] Could not process {path.name}: {e}. Embedding original.")
    # Fallback path
    ext = path.suffix.lower()
    mime = MIME.get(ext, "image/png")
    raw = path.read_bytes()
    return (f"data:{mime};base64," + base64.b64encode(raw).decode("ascii"), len(raw))


def load_photos(images_dir):
    photos = {}
    if not images_dir.exists():
        return photos
    total_bytes = 0
    for f in images_dir.iterdir():
        if f.is_file() and f.suffix.lower() in PHOTO_EXTS:
            uri, size = _encode_one(f)
            total_bytes += size
            pid = f.stem.strip()
            photos[pid] = uri
            if pid.isdigit():
                photos[str(int(pid))] = uri
    if photos:
        avg = total_bytes / len(set(photos.values())) if photos else 0
        mode = "downscaled" if HAVE_PIL else "ORIGINAL (install Pillow to downscale)"
        print(f"[ok] Photos {mode}: {len(set(photos.values()))} files, "
              f"total {total_bytes/1024:.0f} KB, avg {avg/1024:.0f} KB/photo")
    return photos


def photo_for(rec, photos):
    pid = rec.get("PhotoID", "")
    if pid in (None, ""):
        return ""
    if isinstance(pid, float) and pid.is_integer():
        pid = int(pid)
    return photos.get(str(pid).strip(), "")


def load_logo():
    """Find the PwC logo in the working folder and return a base64 data URI."""
    for stem in LOGO_CANDIDATES:
        for ext in PHOTO_EXTS:
            p = SCRIPT_DIR / f"{stem}{ext}"
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
    data = load_workbook_data(INPUT_FILE)
    photos = load_photos(IMAGES_DIR)
    logo_uri = load_logo()

    enriched = []
    for r in data["details"]:
        c = dict(r)
        c["_photo"] = photo_for(r, photos)
        c["_initials"] = initials(r.get("Name"))
        enriched.append(c)

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
    html = (template
        .replace("__EMPLOYEES_JSON__", json.dumps(enriched, default=str))
        .replace("__SKILLS_JSON__",    json.dumps(data["skills"], default=str))
        .replace("__UTIL_JSON__",      json.dumps(data["utilization"], default=str))
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
    )
    OUTPUT_FILE.write_text(html, encoding="utf-8")

    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"[ok] Employees:    {len(enriched)}")
    print(f"[ok] Skills rows:  {len(data['skills'])}")
    print(f"[ok] Util rows:    {len(data['utilization'])}")
    # Sanity check: anyone with no rows in Employee Skills?
    ids_with_skills = {str(r.get("WorkdayID")) for r in data["skills"] if r.get("WorkdayID") not in (None, "")}
    no_skill_emps = [e for e in enriched if str(e.get("WorkdayID")) not in ids_with_skills]
    if no_skill_emps:
        names = ", ".join(str(e.get("Name", "?")) for e in no_skill_emps[:5])
        more = f" + {len(no_skill_emps) - 5} more" if len(no_skill_emps) > 5 else ""
        print(f"[warn] {len(no_skill_emps)} employee(s) have NO rows in Employee Skills sheet: {names}{more}")
        print(f"       They will not appear in the Skill Atlas / Staffing Match.")
    print(f"[ok] Logo found:   {'yes' if logo_uri else 'no'}")
    print(f"[ok] Output:       {OUTPUT_FILE.name}  ({size_kb:,.0f} KB)")


if __name__ == "__main__":
    main()
