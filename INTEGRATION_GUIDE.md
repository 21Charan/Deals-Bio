# Integration Guide — porting Skill Atlas, Capability Risk & Pulse into your Directory template

This explains how to lift the three new tabs out of the generated dashboard
(`template.html` / `Employee_Dashboard.html`) and drop them into your own
office **Directory** template. It is written so an AI assistant can follow it
mechanically: hand it this file plus both templates and ask it to merge.

---

## 1. How the dashboard is built (context)

`generate_report.py` reads `Employee Details.xlsx` (three sheets) and a layout
file `template.html`, then writes a single self-contained `Employee_Dashboard.html`
with the data injected inline (photos base64-embedded, works offline).

The generator fills these placeholder tokens in `template.html`:

| Token | Filled with |
|-------|-------------|
| `__EMPLOYEES_JSON__` | JSON array, one object per row of the **Employee Details** sheet (plus `_photo`, `_initials`) |
| `__SKILLS_JSON__` | JSON array of the **Employee Skills** sheet rows (`Skill Category`, `Skills`, `WorkdayID`) |
| `__UTIL_JSON__` | JSON array of the **Employee Monthly Utilization** sheet rows |
| `__EXP_TOTAL__`, `__EXP_PWC__` | the two experience column names |
| `__SHARE_MODE__` | `true`/`false` (team build strips utilization) |
| `__PBI_URL__` | optional Power BI link |

**Your Directory template already populated by the same generator/schema**, so
the data and most shared scaffolding already exist — the three modules plug in
with minimal wiring.

---

## 2. Where each module lives in `template.html`

Search markers have been added so blocks are easy to find:

- **HTML** — each tab's button + `<section>` is wrapped in
  `<!-- ===== PORTABLE MODULE START: <name> (HTML) ===== -->` …
  `<!-- ===== PORTABLE MODULE END: <name> (HTML) ===== -->`.
  (Tab buttons live together in `<nav class="tabs">`; copy the matching
  `<button data-tab="...">` too.)
- **JS** — banner comments mark `// ===== PORTABLE MODULE: <name> — JS =====`
  and `// ===== SHARED FOUNDATION: ... =====`. The Capability Risk JS is a
  self-contained IIFE under `//  Capability Risk tab`.
- **CSS** — block comments: `/* ---- Skill Atlas ---- */`,
  `/* ---- Capability Risk ---- */`, `/* ---- Shared layout ... ---- */`,
  `/* ---- Staffing Builder ---- */`, `/* ---- Pulse ... ---- */`.

---

## 3. SHARED FOUNDATION — must exist in the host template

Every module depends on this scaffolding. If your Directory template was built
by the same generator it already has most of it; confirm each item exists.

**Data (from the generator):** `EMPLOYEES`, `SKILLS`, `UTIL` consts injected
from the three placeholders.

**Helper functions / globals (top of `<script>`):**
- `esc(str)` — HTML-escape.
- `uniqueValues(key)` and `fillSelect(id, values)` — populate filter dropdowns.
- `SKILLS_BY_EMP`, `UTIL_BY_EMP` — indexes keyed by `WorkdayID`.
- Colour consts `ORANGE, TANGERINE, YELLOW, ROSE, PINK, BURGUNDY, BLACK, GREY, GREY_D`
  and the `PALETTE` array. **All visuals use PwC palette only.**
- CSS variables `--pwc-*` in the `<style>` `:root` (orange, tangerine, greys, etc.).

**Skills index (shared by Skill Atlas + Capability Risk)** — block
`// ===== SHARED FOUNDATION: skills index =====`:
- builds `e._skills` (a `Set` of lowercase skill keys) on every employee,
- `SKILL_DISPLAY` (key → display name), `SKILL_CATEGORY` (key → category),
- `ALL_SKILLS`, `SKILL_COUNT`, and `ORDERED_CATS`.

**Bio modal (all modules open bios):** the `#bio-modal` markup near the end of
`<body>` and the `showBio(workdayId)` function. Skill Atlas, Capability Risk
and Pulse all call `showBio(...)`.

**Tab switching:** the generic `nav.tabs button` click handler that toggles
`.tab-panel.active`. A new `<button data-tab="x">` + `<section id="tab-x">`
"just works" with it.

**Utilization foundation (needed by Capability Risk + Pulse)** — in the Pulse
JS region:
- `ALL_MONTHS`, `UTIL_LOOKUP`, `LATEST_MONTH`,
- per-employee `e._latestUtil`, `e._avgUtil`, `e._sustained3`,
- `utilColor(value)` → `{bg, dark}` PwC ramp,
- `isManagerRole(role)` (defined near the Directory KPIs) — used to flag seniors.

---

## 4. Per-module dependency summary

| Module | HTML IDs (key) | CSS blocks | JS needs |
|--------|----------------|-----------|----------|
| **Skill Atlas** | `tab-skills`, `atlas-cards`, `atlas-cloud`, `atlas-detail`, `atlas-insights`, staffing builder ids | `/* Skill Atlas */`, `/* Staffing Builder */`, shared `.adt-*` detail styles | Skills index, `showBio`, `esc`, `PALETTE`, `ORDERED_CATS` |
| **Capability Risk** | `tab-caprisk`, `cr-territory/competency/cat/role/loc`, `cr-view-toggle`, `cr-scatter`, `cr-kp-table`, `cr-scarce-detail` | `/* Capability Risk */` + **shared layout** (`.smap-controls/.smap-bar/.smap-meta/.smap-grid/.smap-stage/.smap-caption/.smap-net-svgwrap/.smap-people/.smap-person/.sp-*/.smap-detail/.smap-tagrow/.venn-empty/.adt-emoji`) | Skills index, **utilization foundation**, `isManagerRole`, `utilColor`, `showBio`, `esc`, `uniqueValues`, `ORDERED_CATS` |
| **Pulse** | `tab-pulse`, `pl-avg/bench/healthy/stretched/burnout`, heatmap container | `/* Pulse */` | **utilization foundation**, `showBio`, `esc`, `uniqueValues` |

> **Important shared-CSS note:** the layout classes prefixed `.smap-*` were
> introduced for an earlier "Skill Map" tab (now removed) but are **reused by
> Capability Risk** (filter bar, stage, side panel) and by the detail-panel
> people list. They live in the `/* ---- Shared layout ... ---- */` CSS block —
> copy that block whenever you copy Capability Risk. Likewise `.adt-*` (atlas
> detail) and `.venn-empty` are shared.

---

## 5. Step-by-step merge (for each module)

1. **Tab button** — copy the module's `<button data-tab="...">` into your
   template's `<nav class="tabs">`.
2. **Section** — copy everything between that module's
   `PORTABLE MODULE START/END (HTML)` markers into your `<main>`.
3. **CSS** — copy the module's CSS block(s) into your `<style>`. For Capability
   Risk also copy the **Shared layout** block (if not already present).
4. **JS** — copy the module's JS (between its banner and the next banner) into
   your `<script>`, **after** the shared foundation.
5. **Confirm shared foundation** (section 3) exists once in the host. If your
   template lacks the skills index or utilization foundation, copy those shared
   blocks too — only once.
6. **Regenerate** with `python generate_report.py` and open the output.

---

## 6. Gotchas

- **PwC colours only.** Competency/category colours come from `PALETTE`; the
  utilization heatmap from `utilColor`. Don't introduce non-brand colours.
- **Share/team build (`--share`)** strips `UTIL`. Capability Risk's *Scarce vs
  Stretched* view and the whole Pulse tab degrade gracefully (Pulse button is
  removed; Scarce view shows a "no utilization data" message). Key-Person Risk
  still works without utilization.
- **Scale.** Designed for your real data (~900 employees, 14 competencies,
  ~200 skills). These modules are single-pass over people plus a ~200-point
  scatter/table — no heavy computation, so they stay fast.
- **Data quality.** Skill matching is case-insensitive on the lowercase key;
  the free-text `Skills` column on Employee Details is intentionally ignored in
  favour of the `Employee Skills` sheet. Watch for location/spelling dupes
  (e.g. "Hyd" vs "Hyderabad") in source data.
- **Never hand-edit `Employee_Dashboard.html`** — it's regenerated. Edit
  `template.html` and re-run the generator.
