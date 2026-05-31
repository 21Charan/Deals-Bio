# Employee Bio Dashboard

A PwC-branded, single-file HTML dashboard generated from an Excel workbook.
Designed to be shared with clients as one self-contained file — photos and
data are embedded inline, so the page works offline with no external
dependencies.

## Features

- **Directory** — searchable, sortable card grid grouped by role hierarchy
  (Managing Directors → Directors → Senior Managers → Managers → Senior
  Associates → Associates).
- **Filters** — OU, Role, Location, plus live full-text search.
- **Bio popup** — click any card to open a modal with the full profile
  (photo, contacts, qualification, join date, team, skills by category,
  monthly utilization chart). Arrow keys or on-screen prev/next walk through
  the currently visible employees. ESC or click-outside closes.
- **Analytics tab** — donut charts (OU / Role / Location) and a ranked
  average-utilization bar chart.
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

`Employee Skills` sheet: `Skill Category`, `Skills`, `WorkdayID`.

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
