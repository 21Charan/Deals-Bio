# Employee Bio Dashboard

A PwC-branded, single-file HTML dashboard generated from an Excel workbook.
Designed to be shared with clients as one self-contained file — photos and
data are embedded inline, so the page works offline with no external
dependencies.

## Features

- **Directory** — searchable, sortable card grid grouped by role hierarchy
  (Managing Directors → Directors → Senior Managers → Managers → Senior
  Associates → Associates). Cards sized for 4–5 per row so 700+ employees
  remain easy to scan.
- **Filters** — OU, Role, Location, Territory, Competency, and a
  checkbox-based Skills multi-select (searchable, OR-match across both the
  Skills column and the Employee Skills sheet), plus live full-text search
  and a Reset button to clear everything at once.
- **KPI cards** — Total Employees, Operating Units, Locations,
  Competencies, Managers &amp; above (all computed live from the data).
- **Bio popup** — click any card to open a modal with the full profile.
  Every column in the Excel that has a value is shown, so adding new columns
  in the sheet automatically surfaces them in the modal. Includes a smooth
  area+line chart for monthly utilization. Arrow keys or on-screen prev/next
  walk through the currently visible employees. ESC or click-outside closes.
- **Skill Atlas tab** — designed for executive (MD) consumption:
  - **Executive Insights banner**: three auto-generated narrative callouts
    — your deepest capability, concentration risk (skills held by 1 person),
    and thinnest area.
  - **Competency Cards**: one card per Skill Category with a strength
    badge (Deep / Solid / Thin), people + skills count, and the top 6
    skills as labelled horizontal bars. Click any bar to drill in. Use the
    "+ N more skills" button to expand a card to its full skill list.
  - **Skill detail panel**: distribution by Competency / Territory /
    Location / Role, top co-occurring skills ("often paired with…"), and
    the list of people with the skill. Each person opens their bio modal.
  - **Staffing Match builder**: chip-input for *must-have* and
    *nice-to-have* skills with type-ahead. Strict must-have logic
    (candidate must have ALL musts) plus weighted scoring across
    nice-to-haves. Copy emails of the top 10 in one click.
- **Pulse tab** — workforce capacity view:
  - KPI strip (this-month avg, on-bench, healthy, stretched, burnout watch).
  - Filterable monthly utilization heatmap (PwC orange/rose ramp). Click any
    cell or name to open the bio. Paginated at 60 employees per page so it
    stays fast at scale.
  - Three insight cards: On the Bench (lowest util this month), Watch List
    (sustained 3-month avg over 100%), and New Joiners (last 12 months).
- **Boardroom tab** — strategic view for Managing Directors:
  - Filter strip (Competency, Experience band, Role) with Reset.
  - KPI strip: Team Size, **Leverage Ratio** (junior:senior), Median
    Experience, **Capability Coverage** (% of Competency × senior-tier
    pairs that have at least one person).
  - **Talent Funnel** — centred bar pyramid of role tiers (MD &rarr; A).
    Width shows scale. Empty tiers render as faded thin bars so leverage
    gaps are visible.
  - **Capability Matrix** — Competency &times; Role heatmap. Zero cells
    are flagged as capability gaps; dark cells are concentrations.
  - **Drill-down panel** — click a tier or matrix cell to populate a
    "Selected segment" grid with the people in that slice. Click anyone
    to open their bio.
- **Photos & logo** are base64-embedded — the output HTML is one file you
  can email to a client. No CDN, no internet required at viewing time.

## Quick start

```bash
pip install openpyxl
python generate_report.py
```

The script reads:

- `Employee Details.xlsx` — three sheets:
  - `Employee Details` (one row per employee)
  - `Employee Skills` (one row per skill-category entry)
  - `Employee Monthly Utilization` (monthly utilization %)
- `Images/` — photos named after the `PhotoID` column (e.g. `1.png`, `2.png`).
  Missing photos fall back to a PwC-orange initials avatar.
- `PwC Logo.jpg` — shown in the header. Optional.
- `template.html` — the HTML/CSS/JS template the script populates.

…and writes `Employee_Dashboard.html`.

## Schema

`Employee Details` sheet — first row must be the header. Expected columns:

| Column                 | Notes                                              |
|------------------------|----------------------------------------------------|
| Name                   |                                                    |
| WorkdayID              | Used as the primary key across sheets              |
| Role                   | e.g. "Managing Director", "Director", "SM", "SA3"  |
| Experience             | Total years (first occurrence)                     |
| Skills                 | Comma-separated, used if no per-category rows      |
| emailid                |                                                    |
| PhotoID                | Filename stem in `Images/` (e.g. `1` → `1.png`)    |
| Employee Description   | Free-text bio                                      |
| OU                     | Operating Unit                                     |
| Location               |                                                    |
| Experience             | Years at PwC (second occurrence — auto-renamed)    |
| Join Date              |                                                    |
| Qualification          |                                                    |
| Team                   | Comma-separated names; matching names link to bios |
| Team Lead              | Lead's Name or WorkdayID; used by the Org Tree     |
| Territory              | Adds an extra Directory filter                     |
| Competency             | Adds an extra Directory filter + drives the KPI    |

`Employee Skills` sheet (the *authoritative* source of skills): `Skill
Category`, `Skills`, `WorkdayID`. One skill per row is the cleanest format,
but the script also handles legacy rows with multiple skills separated by
`,` or `;`. The free-text `Skills` column on `Employee Details` is
intentionally ignored by the filter, search, atlas, and staffing builder
to avoid delimiter conflicts and duplicates.

`Employee Monthly Utilization` sheet: `WorkdayID`, `Month Year`, `Utilization`, `FY`.

## Reusability

The schema is fixed; the data is not. To regenerate with new data:

1. Replace `Employee Details.xlsx` (keep the same headers).
2. Drop new photos into `Images/` named after each `PhotoID`.
3. Run `python generate_report.py`.

A fresh `Employee_Dashboard.html` appears in the folder.

## Files in this repository

- `generate_report.py` — the generator script.
- `template.html` — the HTML/CSS/JS template with `__PLACEHOLDER__` tokens
  the script replaces.
- `.gitignore` — excludes employee data, photos, the PwC logo, and the
  generated output from version control.
- `README.md` — this file.

## Why no data files in the repo

The real `Employee Details.xlsx`, photos in `Images/`, and the generated
HTML all contain PII (names, employee IDs, emails, faces) and are excluded
by `.gitignore`. They live only on the operator's machine. Clone this repo
and supply your own data file matching the schema above.
"# Deals-Bio" 
